# -*- coding: utf-8 -*-
"""
메뉴 -> 시술 카테고리 단일 매핑 테이블 생성기
출력: data/processed/menu_category.csv  (menu / category / is_free / is_addon / menu_ko)

원본 분석 문서 규약 6-7에 따라 이 CSV가 카테고리 분류의 **단일 기준**이다.
스크립트마다 정규식을 새로 짜지 말고 이 파일을 읽어 쓸 것.

카테고리: 탈색 / 염색 / 컷 / 펌 / 클리닉 / 스타일링 / 기타 / 비시술 / 확인필요
  - 매직·셋팅·다운펌은 모두 '펌'에 포함 (기존 리포트 수치와 호환 유지)
  - is_free=1 : 청구 0원 부가서비스 (아쿠아 계열)
  - is_addon=1: 단독 예약이 사실상 없고 다른 시술에 얹혀 나가는 메뉴
  - '확인필요' : 무엇을 하는 시술인지 확정되지 않아 사장님 확인이 필요한 메뉴
  - menu_ko   : 화면 표시용 한국어 메뉴명. 외국어 표기(네이버 자동번역)는 대응 한국어 메뉴, 한국어 메뉴는 자기 이름.
                분류·집계에는 쓰지 않는다 — 계산은 menu 로 한다

실행: python make_menu_category.py
"""
import pandas as pd, os, re, sys
from pathlib import Path

OUT = Path(os.environ.get("SALON_OUT", "./data/processed"))

CATEGORY = {
    # ---------- 탈색 (36종) ----------
    "뿌리탈색이벤트": "탈색", "무한탈색이벤트": "탈색", "수험생탈색이벤트": "탈색",
    "블랙빼기": "탈색", "수험생탈색": "탈색", "뿌리탈색": "탈색", "탈색이벤트": "탈색",
    "무한탈색+원하는컬러": "탈색", "가을맞이 개강 무한탈색이벤트": "탈색",
    "개강 무한탈색이벤트": "탈색", "전체탈색": "탈색", "블랙빼기무제한탈색": "탈색",
    "블랙빼기무제한": "탈색", "디자인탈색(옴브레,솜브레,시크릿투톤)": "탈색",
    "페이퍼뿌리탈색": "탈색", "탈색": "탈색", "프리미엄탈색": "탈색",
    "개강이벤트(무한탈색)": "탈색", "수험탈색이벤트": "탈색", "솜브레": "탈색", "옴브레": "탈색",
    # 다국어 표기 (네이버 자동번역)
    "无限脱色活动": "탈색", "根元ブリーチイベント": "탈색", "Root Bleaching Event": "탈색",
    "Indefinite Bleaching Event": "탈색", "考生脱色": "탈색", "发根脱色事件": "탈색",
    "迎秋 开讲 无限脱色活动": "탈색", "黑色头巾": "탈색",          # 블랙빼기 오역, 단가 200,000 일치
    "Bleach the examinees": "탈색", "無限ブリーチイベント": "탈색",
    "無限ブリーチ+ご希望のカラー": "탈색", "开讲 无限脱色活动": "탈색", "整体脱色": "탈색",
    "无限脱色+想要的颜色": "탈색", "Infinite bleaching + desired color": "탈색",

    # ---------- 염색 (13종) ----------
    "알파,베타뿌리염색": "염색", "베타뿌리염색": "염색", "감마,델타염색": "염색",
    "감마염색": "염색", "베타염색": "염색", "감마,델타뿌리염색": "염색",
    "알파,베타염색": "염색", "엡실론뿌리염색": "염색", "감마뿌리염색": "염색",
    "엡실론 전체염색": "염색",
    "페이크블랙": "염색",                                       # 탈색모에 블랙 입히는 염색. 최빈 단가 150,000 (단가 2종, 미해결)
    "伽马、德尔塔染色": "염색", "Alpha and Beta dyed their hair": "염색",

    # ---------- 컷 (10종) ----------
    "여성컷": "컷", "남자컷": "컷", "앞머리컷": "컷", "디자인컷": "컷", "스케일링 디자인컷": "컷",
    "女士发型": "컷", "design cut": "컷", "女性カット": "컷", "a woman's cut": "컷",
    "设计图": "컷",                                             # 디자인컷 오역, listed 가산 0

    # ---------- 펌 (매직·셋팅·다운펌 포함, 16종) ----------
    "일반펌": "펌", "매직,볼륨매직": "펌", "다운펌+컷": "펌", "열펌": "펌",
    "매직셋팅": "펌", "복구매직,볼륨매직": "펌", "복구펌": "펌", "크리닉뿌리매직": "펌",
    "뿌리매직": "펌", "다운펌": "펌", "단백질 펌": "펌", "프리미엄 매직셋팅펌": "펌",
    "복구 매직셋팅펌": "펌", "뿌리펌(애드펌,아이롱펌등등)": "펌", "뿌리펌(애드펌)": "펌",
    "회복매직셋팅+뿌리펌+클리닉": "펌",

    # ---------- 클리닉 (12종) ----------
    "아쿠아클리닉": "클리닉", "水族诊所": "클리닉",
    "Aqua Clinic": "클리닉", "アクアトリートメント": "클리닉",
    "아쿠아": "클리닉",
    "히노끼 두피 크리닉": "클리닉", "신데렐라클리닉": "클리닉", "두피클리닉(전문)": "클리닉",
    "무실리콘 천연 아르보크리닉": "클리닉", "두피 크리닉": "클리닉",
    "항노화클리닉": "클리닉", "복구크리닉": "클리닉",

    # ---------- 스타일링 (3종) ----------
    "연예인드라이": "스타일링", "샴푸+연예인드라이": "스타일링", "포니테일": "스타일링",

    # ---------- 기타 (1종) ----------
    "붙임머리 리터치(자샵, 타샵)": "기타",

    # ---------- 비시술 (1종) ----------
    "선불권": "비시술",

    # ---------- 확인필요 (2종) ----------
    "제타": "확인필요",                 # 단독예약 32건 · 단가 39,000원. 트리트먼트로 추정되나 미확정
    "솔루션": "확인필요",               # 2행 · 단가 87,000원. 클리닉인지 염색 옵션인지 불명
}

# 청구 0원 부가서비스: 아쿠아 계열 전부
IS_FREE = {"아쿠아클리닉", "水族诊所", "Aqua Clinic",
           "アクアトリートメント", "아쿠아"}

# 단독 예약이 사실상 없고 다른 시술에 얹혀 나가는 메뉴 (단독률 < 25%)
IS_ADDON = IS_FREE | {"앞머리컷", "두피클리닉(전문)", "무실리콘 천연 아르보크리닉"}

# 외국어 표기 -> 원래 한국어 메뉴 (menu_ko). 네이버가 한국어 메뉴를 자동번역해 보여준 이름이다.
# 화면 표시용이다 — 분류·집계는 menu 로 하고, 노트북은 표·차트·print 에 내기 직전에만 이 값을 쓴다.
# 25종 모두 근거 넷 중 셋 이상이 맞는다:
#   (a) 뜻  (b) 단가(price_catalog)  (c) 판매 시기(한국어 메뉴의 판매 기간 안)
#   (d) 같은 예약의 핸즈 메모 '예약시술메뉴' — 외국어로 잡힌 예약도 핸즈에는 한국어 메뉴명으로 남아 있다
# 근거가 엇갈리거나 모자란 표기는 넣지 않는다 (menu_ko 가 원문 그대로 남는다).
MENU_KO = {
    # 탈색
    "无限脱色活动": "무한탈색이벤트", "無限ブリーチイベント": "무한탈색이벤트",
    "Indefinite Bleaching Event": "무한탈색이벤트",
    "根元ブリーチイベント": "뿌리탈색이벤트", "Root Bleaching Event": "뿌리탈색이벤트",
    "发根脱色事件": "뿌리탈색이벤트",                        # '이벤트'를 事件(사건)으로 옮겼다
    "考生脱色": "수험생탈색", "Bleach the examinees": "수험생탈색",   # '이벤트'가 없다. 핸즈 기록도 수험생탈색
    "迎秋 开讲 无限脱色活动": "가을맞이 개강 무한탈색이벤트",
    "开讲 无限脱色活动": "개강 무한탈색이벤트",               # 개강이벤트(무한탈색)과 단가가 같다. 핸즈 기록이 이쪽
    "整体脱色": "전체탈색",
    "无限脱色+想要的颜色": "무한탈색+원하는컬러", "無限ブリーチ+ご希望のカラー": "무한탈색+원하는컬러",
    "Infinite bleaching + desired color": "무한탈색+원하는컬러",
    "黑色头巾": "블랙빼기",                                  # 오역('검은 두건'). 단가·핸즈 기록이 블랙빼기
    # 염색
    "伽马、德尔塔染色": "감마,델타염색", "Alpha and Beta dyed their hair": "알파,베타염색",
    # 컷
    "女士发型": "여성컷", "女性カット": "여성컷", "a woman's cut": "여성컷",
    "design cut": "디자인컷",
    "设计图": "디자인컷",                                    # 오역('설계도'). 단가·핸즈 기록이 디자인컷
    # 클리닉 (청구 0원 부가서비스)
    "水族诊所": "아쿠아클리닉", "Aqua Clinic": "아쿠아클리닉", "アクアトリートメント": "아쿠아클리닉",
}
HANGUL = re.compile(r"[가-힣]")


def main():
    menu_items = pd.read_parquet(OUT / "menu_lines.parquet")   # FIX-3c 개명
    seen = set(menu_items["menu"].unique())

    missing = sorted(seen - set(CATEGORY))
    extra   = sorted(set(CATEGORY) - seen)
    if missing:
        raise RuntimeError(
            f"매핑 누락 {len(missing)}종 — CATEGORY 딕셔너리에 추가하세요:\n  " +
            "\n  ".join(missing))
    if extra:
        print(f"경고: 데이터에 없는 메뉴가 매핑에 {len(extra)}종 있습니다: {extra}")

    df = pd.DataFrame({"menu": sorted(seen)})
    df["category"]  = df["menu"].map(CATEGORY)
    df["is_free"]   = df["menu"].isin(IS_FREE).astype(int)
    df["is_addon"]  = df["menu"].isin(IS_ADDON).astype(int)

    # menu_ko — 대응표의 키는 외국어 메뉴, 값은 한국어 메뉴여야 하고, 둘의 카테고리·무료 여부가 같아야 한다
    for f, k in MENU_KO.items():
        assert f in CATEGORY and not HANGUL.search(f), f"MENU_KO 키가 외국어 메뉴가 아니다: {f}"
        assert k in CATEGORY and HANGUL.search(k), f"MENU_KO 값이 한국어 메뉴가 아니다: {k}"
        assert CATEGORY[f] == CATEGORY[k] and (f in IS_FREE) == (k in IS_FREE), \
            f"외국어 표기와 한국어 메뉴의 카테고리·무료 여부가 다르다: {f} -> {k}"
    df["menu_ko"]   = df["menu"].map(MENU_KO).fillna(df["menu"])

    p = OUT / "menu_category.csv"
    df.to_csv(p, index=False, encoding="utf-8-sig")   # 엑셀에서 바로 열리도록 BOM

    n = menu_items["menu"].value_counts()
    df["행수"] = df["menu"].map(n)
    print(f"저장: {p}  ({len(df)}종)")
    print("\n카테고리별 메뉴수 / 메뉴행수")
    g = df.groupby("category").agg(메뉴종수=("menu", "size"), 메뉴행수=("행수", "sum"))
    g["행비중%"] = (g["메뉴행수"] / g["메뉴행수"].sum() * 100).round(1)
    print(g.sort_values("메뉴행수", ascending=False).to_string())
    print(f"\nis_free=1 : {df['is_free'].sum()}종 / {df.loc[df['is_free']==1,'행수'].sum():,}행")
    print(f"is_addon=1: {df['is_addon'].sum()}종 / {df.loc[df['is_addon']==1,'행수'].sum():,}행")
    chk = df[df["category"] == "확인필요"]
    print(f"\n확인필요 {len(chk)}종 / {chk['행수'].sum()}행")
    print(chk[["menu", "행수"]].to_string(index=False))
    fx = df[~df["menu"].map(lambda m: bool(HANGUL.search(m)))]
    left = fx[fx["menu_ko"] == fx["menu"]]
    print(f"\nmenu_ko: 외국어 표기 {len(fx)}종 중 {len(fx) - len(left)}종 -> 한국어 메뉴 "
          f"{fx.loc[fx['menu_ko'] != fx['menu'], 'menu_ko'].nunique()}종"
          + (f" · 대응 없음(원문 유지) {len(left)}종: {left['menu'].tolist()}" if len(left) else ""))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
