# -*- coding: utf-8 -*-
"""
관측 소요시간 테이블 생성기 (예약 시작 -> 완료처리)
출력: data/manual/service_duration_v2.csv

원천은 reservations 의 done_dt - use_dt 다.
원본 분석 문서 규약 6-1 은 done_dt 를 **시간대 분석**(몇 시에 붐비는가)에 쓰지 말라고 한다.
60.5%가 익일 오전 9~10시 배치이기 때문이다. 그 금지는 그대로 유효하다.
다만 **당일 처리된 건에 한해서는** done_dt - use_dt 가 시술 소요시간의 유효한 추정치다.
근거 4개:
  (1) 대상 945건에 초 단위 동일 타임스탬프 그룹이 0개 -> 마감 일괄정산 서명이 없다
  (2) 시작 시각대별 중앙 종료시각이 16.7시 -> 19.9시로 퍼진다 -> 마감 고정이 아니다
  (3) 병행이 물리적으로 불가능한 표본(디자이너-일 완료 1건)에서 같은 값이 재현된다
  (4) 사장님 인터뷰(뿌리탈색 4~5h / 전체탈색 8h)와 정합. 가정 A(3.5h)는 불일치

한계 — 반드시 같이 읽을 것:
  * 완료 2,732건 중 관측 가능한 것은 34.6% 뿐이다
  * 절단 편향: 시작 시각별 당일처리율이 10~13시 37~40% -> 19시 12.8% 로 떨어진다.
    늦게 시작해 늦게 끝난 건이 빠지므로 **관측값은 하한에 가깝다**
  * 연도 쏠림: 2025년 28.7% vs 2026년 42.1%
  * 측정 대상은 정확히는 "예약 시작 -> 완료처리 클릭" 이며 시술 종료 그 자체가 아니다

⚠️ 이 값을 라인별로 합산해 가동률 분모에 넣지 말 것.
   탈색 방치 시간에 다른 손님을 병행하므로 겹치는 구간이 중복 계산된다.
   실제로 합산하면 DESIGNER_A 가동률이 141.6% 가 된다. 캐파 지표는 구간 합집합을 쓸 것.

집계 단위는 **예약 1건 = 값 1개** 다. 라인 합산을 하지 않는다.
  - 메뉴 레벨   : 유료 메뉴(is_free==0)가 정확히 1종인 예약만, 그 메뉴명으로 묶음
  - 카테고리 레벨: 같은 표본 전체를 01번 rep_cat 규칙으로 묶음
    (rep_cat 은 무료 아쿠아클리닉 라인도 포함해 계산한다. 01번 §0-4 와 동일)

실행: python make_service_duration.py
"""
import pandas as pd, numpy as np, os, sys
from pathlib import Path

OUT    = Path(os.environ.get("SALON_OUT", "./data/processed"))
MANUAL = Path(os.environ.get("SALON_MANUAL", "./data/manual"))

# 01번 §0-4 와 동일. 여기서 새로 정의하지 말고 순서를 그대로 맞춘다.
PRIORITY = ["탈색", "펌", "염색", "컷", "클리닉", "스타일링", "기타", "확인필요"]
EXCLUDE  = ["비시술"]                      # 선불권 — 시술이 아니므로 분석 제외

N_SUFFICIENT = 30                          # 신뢰도 '충분'
N_REFERENCE  = 10                          # 신뢰도 '참고'

COLS = ["level", "category", "menu", "n", "중앙값_h", "p25_h", "p75_h", "평균_h",
        "최소_h", "최대_h", "무료클리닉동반율", "신뢰도", "겹침없음_n", "겹침없음_중앙값_h"]


def confidence(n):
    if n >= N_SUFFICIENT: return f"충분(n>={N_SUFFICIENT})"
    if n >= N_REFERENCE:  return f"참고(n{N_REFERENCE}~{N_SUFFICIENT - 1})"
    return f"소표본(n<{N_REFERENCE})"


def load():
    """reservations / menu_lines / menu_category 를 읽고 01번과 같은 파생을 만든다"""
    R = pd.read_parquet(OUT / "reservations.parquet")
    L = pd.read_parquet(OUT / "menu_lines.parquet")
    C = pd.read_csv(OUT / "menu_category.csv", encoding="utf-8-sig")

    cat_of, free_of = C.set_index("menu")["category"], C.set_index("menu")["is_free"]
    LX = L.copy()
    LX["category"] = LX.menu.map(cat_of)
    LX["is_free"]  = LX.menu.map(free_of).astype(int)
    assert LX.category.notna().all(), "카테고리 매핑 실패 라인 존재 — make_menu_category.py 를 먼저 돌릴 것"
    LX = LX[~LX.category.isin(EXCLUDE)].copy()

    # 대표 시술 (01번 §0-4 와 글자 그대로 동일한 규칙)
    _p = {c: i for i, c in enumerate(PRIORITY)}
    rep = (LX.assign(_o=LX.category.map(_p))
             .sort_values(["예약번호", "_o"])
             .drop_duplicates("예약번호")
             .set_index("예약번호")["category"])
    R["rep_cat"] = R.예약번호.map(rep)
    return R, LX, C


def build_sample(R, LX):
    """관측 대상 표본을 만든다. 필터 4단은 이 함수 밖에서 바뀌지 않는다."""
    D = R[R.status == "이용완료"].copy()
    n0 = len(D)

    D = D[D.use_dt.notna() & D.done_dt.notna()]
    n1 = len(D)

    same_day = D.done_dt.dt.normalize() == D.use_dt.dt.normalize()
    D = D[same_day]
    n2 = len(D)

    D["gap_h"] = (D.done_dt - D.use_dt).dt.total_seconds() / 3600
    D = D[D.gap_h > 0].copy()
    n3 = len(D)

    print(f"필터 1) status == 이용완료               {n0:>6,}건")
    print(f"필터 2) use_dt / done_dt 둘 다 존재       {n1:>6,}건  ({n1-n0:+,})")
    print(f"필터 3) done_dt.date == use_dt.date      {n2:>6,}건  ({n2-n1:+,})")
    print(f"필터 4) gap_h > 0                        {n3:>6,}건  ({n3-n2:+,})")
    print(f"  -> 관측 표본 {n3:,}건 = 완료 {n0:,}건의 {n3/n0*100:.1f}%")

    # 유료 메뉴가 정확히 1종인 예약 (메뉴 레벨의 모집단)
    paid = LX[LX.is_free == 0]
    only1 = paid.groupby("예약번호").menu.agg(["nunique", "first"])
    only1 = only1[only1["nunique"] == 1]
    D["pure_menu"] = D.예약번호.map(only1["first"])

    # 무료 클리닉 동반 여부
    has_free = LX[LX.is_free == 1].예약번호.unique()
    D["has_free_clinic"] = D.예약번호.isin(has_free)

    # 겹침없음 표본 = 그날 그 디자이너의 **완료 건이 1건뿐**인 날
    #   병행 시술이 물리적으로 불가능하므로 오염 검증본이 된다.
    #   판정 모집단은 표본이 아니라 완료 전체다(표본 안에서만 세면 과대 계상된다).
    ALLD = R[R.status == "이용완료"].copy()
    ALLD["date"] = ALLD.use_dt.dt.normalize()
    day_n = ALLD.groupby(["designer", "date"]).예약번호.size().rename("day_n")
    D["date"]  = D.use_dt.dt.normalize()
    D["day_n"] = pd.MultiIndex.from_frame(D[["designer", "date"]]).map(day_n)
    D["solo"]  = D.day_n == 1
    print(f"  -> 겹침없음(디자이너-일 완료 1건) 표본 {int(D.solo.sum()):,}건")
    return D


def agg(g, level, category, menu):
    v, s = g.gap_h, g[g.solo].gap_h
    return {
        "level": level, "category": category, "menu": menu,
        "n": len(v),
        "중앙값_h": round(float(v.median()), 2),
        "p25_h":   round(float(v.quantile(.25)), 2),
        "p75_h":   round(float(v.quantile(.75)), 2),
        "평균_h":   round(float(v.mean()), 2),
        "최소_h":   round(float(v.min()), 2),
        "최대_h":   round(float(v.max()), 2),
        "무료클리닉동반율": round(float(g.has_free_clinic.mean() * 100), 1),
        "신뢰도": confidence(len(v)),
        "겹침없음_n": len(s),
        "겹침없음_중앙값_h": round(float(s.median()), 2) if len(s) else np.nan,
    }


def build_table(D, C):
    cat_of = C.set_index("menu")["category"].to_dict()
    rows = []

    # --- 메뉴 레벨: 유료 메뉴 1종 예약만 ---
    M = D[D.pure_menu.notna()]
    for menu, g in M.groupby("pure_menu"):
        rows.append(agg(g, "메뉴", cat_of[menu], menu))

    # --- 카테고리 레벨: 표본 전체를 rep_cat 으로 ---
    for cat, g in D[D.rep_cat.notna()].groupby("rep_cat"):
        rows.append(agg(g, "카테고리", cat, np.nan))

    df = pd.DataFrame(rows, columns=COLS)
    df["_lv"] = (df.level == "카테고리").astype(int)          # 메뉴 먼저
    df = (df.sort_values(["_lv", "category", "n"], ascending=[True, True, False])
            .drop(columns="_lv").reset_index(drop=True))
    return df


def compare_with_v1(df, path):
    """기존 service_duration_measured.csv 와 메뉴별 중앙값을 대조한다"""
    if not path.exists():
        print(f"\n기존 CSV 를 찾지 못해 대조를 건너뛴다: {path}")
        return
    v1 = pd.read_csv(path, encoding="utf-8-sig")
    a = v1[v1.level == "메뉴"].set_index("menu")[["n", "중앙값_h"]]
    b = df[df.level == "메뉴"].set_index("menu")[["n", "중앙값_h", "겹침없음_n", "겹침없음_중앙값_h"]]
    j = a.join(b, how="outer", lsuffix="_v1", rsuffix="_v2")
    j["n차"]      = j.n_v2 - j.n_v1
    j["중앙값차"] = (j.중앙값_h_v2 - j.중앙값_h_v1).round(2)
    j = j.sort_values("n_v2", ascending=False)

    print("\n" + "=" * 96)
    print("기존 service_duration_measured.csv 대조 (메뉴 레벨)")
    print("=" * 96)
    print(j.to_string())

    only_v1 = sorted(set(a.index) - set(b.index))
    only_v2 = sorted(set(b.index) - set(a.index))
    if only_v1: print(f"\nv1 에만 있는 메뉴 {len(only_v1)}종: {only_v1}")
    if only_v2: print(f"v2 에만 있는 메뉴 {len(only_v2)}종: {only_v2}")

    big = j[j.중앙값차.abs() >= 0.5].sort_values("중앙값차", key=abs, ascending=False)
    print(f"\n중앙값 차이 0.5h 이상 — {len(big)}종")
    if len(big):
        print(big[["n_v1", "n_v2", "중앙값_h_v1", "중앙값_h_v2", "중앙값차"]].to_string())
    print(f"\nn 이 늘어난 메뉴 {int((j.n차 > 0).sum())}종 / 줄어든 메뉴 {int((j.n차 < 0).sum())}종")
    print(f"v1 총 n {int(a.n.sum()):,} -> v2 총 n {int(b.n.sum()):,}")
    print("  v1 은 재현 불가한 추가 필터가 한 단 더 들어가 있다(특히 탈색). v2 는 레시피가 문서화돼 있다.")


def main():
    R, LX, C = load()
    print("=" * 96)
    print("관측 소요시간 테이블 생성 (예약 시작 -> 완료처리)")
    print("=" * 96)
    D  = build_sample(R, LX)
    df = build_table(D, C)

    MANUAL.mkdir(parents=True, exist_ok=True)
    p = MANUAL / "service_duration_v2.csv"
    df.to_csv(p, index=False, encoding="utf-8-sig")          # 엑셀에서 바로 열리도록 BOM
    print(f"\n저장: {p}  ({len(df)}행 = 메뉴 {int((df.level=='메뉴').sum())} / 카테고리 {int((df.level=='카테고리').sum())})")

    print("\n--- 카테고리 레벨 ---")
    print(df[df.level == "카테고리"][["category", "n", "중앙값_h", "p25_h", "p75_h",
                                      "겹침없음_n", "겹침없음_중앙값_h", "신뢰도"]].to_string(index=False))
    print("\n--- 메뉴 레벨 (n>=10) ---")
    print(df[(df.level == "메뉴") & (df.n >= N_REFERENCE)]
            [["category", "menu", "n", "중앙값_h", "p25_h", "p75_h",
              "겹침없음_n", "겹침없음_중앙값_h", "신뢰도"]].to_string(index=False))

    print(f"\n메뉴 레벨 n 합 {int(df[df.level=='메뉴'].n.sum()):,} / 카테고리 레벨 n 합 {int(df[df.level=='카테고리'].n.sum()):,}")
    print("  두 레벨은 같은 표본을 다르게 자른 것이다. 메뉴 레벨이 '유료 메뉴 1종' 부분집합이라 더 작다.")

    compare_with_v1(df, MANUAL / "service_duration_measured.csv")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
