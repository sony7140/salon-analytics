"""
예약 데이터 통합 파이프라인
원본 7개 파일 -> 정제 테이블 6개

실행:  python build.py
필요:  pip install -r requirements-build.txt (공개하지 않은 원본 입력 필요)
"""
import pandas as pd, numpy as np, re, io, os, sys
from datetime import datetime
from pathlib import Path

RAW = Path(os.environ.get("SALON_RAW", "./data/raw"))
OUT = Path(os.environ.get("SALON_OUT", "./data/processed"))
PASSWORD = os.environ.get("SALON_XLSX_PW", "")      # 네이버 export 복호화 암호. 평문으로 두지 않습니다.

# 파일명을 정확히 적지 않고 패턴으로 자동 탐색합니다.
# 띄어쓰기/언더바/날짜 차이가 있어도 알아서 찾습니다.
PATTERNS = {
    "naver":  ["예약자관리"],
    "hands":  ["예약상세목록"],
    "report": ["리포트"],
}

def find(kind):
    """RAW 폴더에서 키워드가 들어간 엑셀 파일을 찾아 정렬해 반환"""
    keys = PATTERNS[kind]
    hits = []
    for p in sorted(RAW.iterdir()):
        if p.name.startswith("~$") or p.is_dir():      # 엑셀 임시파일 제외
            continue
        if p.suffix.lower() not in (".xlsx", ".xls"):
            continue
        norm = p.name.replace(" ", "").replace("_", "")
        if any(k.replace(" ", "").replace("_", "") in norm for k in keys):
            hits.append(p)
    if not hits:
        raise FileNotFoundError(
            f"[{kind}] '{keys[0]}' 가 들어간 엑셀 파일을 {RAW.resolve()} 에서 찾지 못했습니다.\n"
            f"현재 폴더에 있는 파일: {[q.name for q in RAW.iterdir()]}")
    return hits

# ---------- 공통 파서 ----------
_RE_USE = re.compile(r"\s*(\d{2})\.\s*(\d{1,2})\.\s*(\d{1,2})\.\((.)\)\s*(오전|오후)\s*(\d{1,2}):(\d{2})")
_RE_LOG = re.compile(r"\s*(\d{4})-(\d{2})-(\d{2})\s*\(.\)\s*(오전|오후)\s*(\d{1,2}):(\d{2})(?::(\d{2}))?")
_RE_HDT = re.compile(r"(\d{2})-(\d{2})-(\d{2})\s*\n?\s*.\s*\n?\s*(\d{1,2}):(\d{2})")

def _ampm(h, ap):
    """네이버 export는 '오후 17:00' 처럼 24시간제와 오전/오후를 섞어 씀"""
    if h <= 12:
        if ap == "오후" and h != 12: h += 12
        if ap == "오전" and h == 12: h = 0
    return h

def parse_use_dt(s):
    if not isinstance(s, str): return pd.NaT
    m = _RE_USE.match(s)
    if not m: return pd.NaT
    yy, mm, dd, _, ap, hh, mi = m.groups()
    try: return datetime(2000+int(yy), int(mm), int(dd), _ampm(int(hh), ap), int(mi))
    except ValueError: return pd.NaT

def parse_log_dt(s):
    if not isinstance(s, str): return pd.NaT
    m = _RE_LOG.match(s)
    if not m: return pd.NaT
    y, mo, d, ap, hh, mi, ss = m.groups()
    try: return datetime(int(y), int(mo), int(d), _ampm(int(hh), ap), int(mi), int(ss or 0))
    except ValueError: return pd.NaT

def parse_hands_dt(s):
    s = str(s).replace("[네이버]", "")
    m = _RE_HDT.search(s)
    if not m: return pd.NaT
    y, mo, d, h, mi = map(int, m.groups())
    try: return datetime(2000+y, mo, d, h, mi)
    except ValueError: return pd.NaT

# ---------- 1. 네이버 예약자관리 (암호화 xlsx) ----------
# 파일명 끝에 박힌 export 시각. 예: <상호>_..._예약자관리_20260823_1235.xlsx -> "202608231235"
_RE_EXPORT = re.compile(r"_(\d{8})_(\d{4})$")

def _export_key(path):
    """
    네이버 export 시각을 정렬 가능한 문자열로 반환합니다. 큰 값 = 나중에 뽑은 파일.

    ⚠️ 파일명 전체를 사전순 정렬해서 쓰면 안 됩니다. 지금은 접두사가 같아서 우연히
       맞을 뿐이고, 접두사가 바뀌면 조용히 뒤집힙니다. 시각을 명시적으로 뽑아 씁니다.
       (월별 리포트에서 _1/_2/_3 파일명 정렬에 의존하다 같은 문제를 겪었습니다.)
    """
    m = _RE_EXPORT.search(path.stem)
    if not m:
        raise ValueError(
            f"네이버 파일명에서 export 시각을 읽지 못했습니다: {path.name}\n"
            f"'..._YYYYMMDD_HHMM.xlsx' 형식이어야 예약번호가 두 파일에 겹칠 때 "
            f"최신 export를 고를 수 있습니다.")
    return m.group(1) + m.group(2)


def load_naver():
    import msoffcrypto
    frames = []
    for path in find("naver"):
        with open(path, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=PASSWORD)
            buf = io.BytesIO(); off.decrypt(buf); buf.seek(0)
        df = pd.read_excel(buf, sheet_name="Report", header=2)   # 안내문 2줄 skip
        df = df[df["예약번호"].notna()].copy()
        df["src_file"] = path.name
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True)

    for col, fn in [("이용일시","use"),("예약신청일시","log"),("예약확정일시","log"),
                    ("이용완료일시","log"),("예약취소일시","log")]:
        if col in raw.columns:
            raw[col+"_dt"] = raw[col].apply(parse_use_dt if fn=="use" else parse_log_dt)

    # ── FIX-3a (2026-08) 1단계: 파일 간 중복을 라인 단계에서 먼저 해소합니다 ──
    # 두 export의 커버 구간이 2026-01-01 하루 겹칩니다. 그날 예약 4건이 양쪽 파일에
    # 모두 들어 있고, 라인행으로는 12행입니다.
    #   raw  9,888행  <- 전량 보존 (원본 레이어)
    #   mart 9,876행  <- 예약번호마다 최신 export 파일 하나만
    # ⚠️ 순서가 중요합니다. 예전에는 결제금액을 max로 접어서 이 중복이 드러나지
    #    않았지만, 아래에서 sum으로 바꾸는 순간 겹치는 4건이 그대로 2배가 됩니다.
    #    반드시 "파일 선택 -> 그 다음 sum" 순서를 지키세요.
    exp = {p.name: _export_key(p) for p in find("naver")}
    raw["_export"] = raw["src_file"].map(exp)
    assert raw["_export"].notna().all(), "src_file에 export 시각을 못 붙인 행이 있습니다"
    mart = raw[raw["_export"] == raw.groupby("예약번호")["_export"].transform("max")].copy()

    _dup = raw.groupby("예약번호")["src_file"].nunique()
    assert (mart.groupby("예약번호")["src_file"].nunique() == 1).all(), \
        "mart에 한 예약이 두 파일에서 온 행이 남아 있습니다 - 파일 선택이 안 먹었습니다"
    assert int((_dup > 1).sum()) == 4, f"파일 간 겹치는 예약 {int((_dup>1).sum())}건 != 4"
    assert len(raw)  == 9_888, f"raw 라인행 {len(raw):,} != 9,888 (원본이 바뀌었는지 확인)"
    assert len(mart) == 9_876, f"mart 라인행 {len(mart):,} != 9,876"

    # ── FIX-3c (2026-08): 옵션행을 롱테이블로 분리하면서 라인 단가를 살립니다 ──
    # 결제금액은 예약 전체 금액이 아니라 **그 옵션 라인 하나의 단가**입니다.
    # 아래에서 예약 단위로 sum 해버리면 "이 시술이 얼마에 팔렸는지"가 사라지고,
    # 옵션이 2개 이상인 예약에서는 합계에서 역산할 방법도 없습니다.
    # 그래서 접기 전에 라인 그대로 뽑아 둡니다 (menu_items -> menu_lines 개명).
    menu = (mart[["예약번호","가격분류 및 옵션","결제금액"]]
            .dropna(subset=["가격분류 및 옵션"])
            .rename(columns={"가격분류 및 옵션":"menu","결제금액":"line_price"})
            .assign(menu=lambda d: d["menu"].str.strip())
            .reset_index(drop=True))

    # 예전 코드는 여기서 .drop_duplicates()로 끝냈습니다. 지금 데이터에서는 한 행도
    # 지우지 않지만(7,165행 = 고유 (예약,메뉴) 7,165쌍), 단가가 붙은 뒤로는 조용히
    # 접히는 순간 금액이 사라집니다. 접지 말고 소리 내서 실패시킵니다.
    assert len(menu) == 7_165, f"메뉴 라인 {len(menu):,} != 7,165"
    assert not menu.duplicated(["예약번호","menu"]).any(), \
        "한 예약에 같은 메뉴가 두 줄 있습니다 - 단가를 어느 줄에서 읽을지 먼저 정하세요"
    assert menu["menu"].nunique() == 94, f'메뉴 {menu["menu"].nunique()}종 != 94'
    assert int(menu["line_price"].sum()) == 464_899_000, \
        f'라인단가 합 {int(menu["line_price"].sum()):,} != 464,899,000'
    # 옵션이 안 붙은 라인 2,711행은 결제금액이 전부 0입니다. 그래서 라인 합계가
    # 예약 결제금액 합계와 한 원도 안 틀립니다(아래 예약 단위 assert로 고정).
    # 이게 깨지면 menu_lines가 매출을 흘리고 있는 것입니다.
    assert int(mart.loc[mart["가격분류 및 옵션"].isna(), "결제금액"].sum()) == 0, \
        "옵션 없는 라인에 결제금액이 붙어 있습니다 - menu_lines 합계가 예약 합계와 어긋납니다"

    # ── FIX-3a 2단계: 예약 단위로 축약 ──────────────────────────────────────
    # 위에서 예약번호마다 파일을 하나로 좁혔으므로 여기 "first"들은 전부 그 파일의
    # 값입니다. 예전 코드는 sort_values("src_file") + "first" 였고 src_file만 "last"라,
    # 값은 구(舊) export에서 오는데 라벨만 신(新) export를 가리켰습니다. 주석("마지막
    # export 우선")과 코드가 정반대였습니다.
    #
    # 결제금액만 "sum"입니다. 한 예약의 옵션 라인마다 단가가 흩어져 있어서 max는
    # 가장 비싼 라인 하나만 셉니다. 값이 있는 라인이 2개 이상인 예약이 226건이고,
    # 완료 기준 7,345,000원이 통째로 누락되고 있었습니다.
    # 실결제금액은 반대로 라인마다 예약 전체 금액이 반복 기재됩니다(아래 assert로 고정).
    agg = {"유입경로":"first","예약자":"first","전화번호":"first","방문자":"first","상태":"first",
           "스타일리스트":"first","결제금액":"sum","실결제금액":"max","쿠폰적용금액":"max",
           "매장결제금액(B)":"max","매장결제수단":"first","환불금액":"max",
           "블랙리스트등록여부":"first","요청사항":"first","취소사유":"first","직원메모":"first",
           "이용일시_dt":"first","예약신청일시_dt":"first","예약확정일시_dt":"first",
           "이용완료일시_dt":"first","예약취소일시_dt":"first","src_file":"first"}
    agg = {k:v for k,v in agg.items() if k in mart.columns}
    assert (mart.groupby("예약번호")["실결제금액"].nunique(dropna=False) == 1).all(), \
        "실결제금액이 한 예약 안에서 갈립니다 - max로 접으면 안 됩니다"
    assert mart["결제금액"].notna().all(), \
        "결제금액에 결측이 있습니다 - sum은 NaN을 0으로 접으므로 먼저 확인하세요"

    # 겹치는 예약 4건은 두 파일의 값이 완전히 같습니다. 그래서 지금은 어느 파일을
    # 고르든 결과가 한 글자도 안 바뀌고, "최신 export 우선" 규칙 자체는 데이터로
    # 검증되지 않습니다(transform을 max에서 min으로 바꿔도 출력이 동일합니다).
    # 재export로 값이 갈리기 시작하면 그때부터 선택이 실제로 결과를 바꾸므로,
    # 조용히 넘어가지 않게 여기서 고정합니다.
    # 규약 6: astype(str) 뒤에 fillna("")가 필요합니다. pandas 2.x는 NaN을 "nan"으로
    # 바꾸지만 3.x는 NaN(float)을 그대로 남겨 join이 TypeError로 터집니다. 어느 쪽이든
    # 두 파일을 똑같이 처리해 비교만 하므로 판정 자체는 버전과 무관합니다.
    _dupids = _dup[_dup > 1].index
    _chk = [c for c in agg if c != "src_file"] + ["가격분류 및 옵션"]
    _sig = (raw[raw["예약번호"].isin(_dupids)]
              .assign(_row=lambda d: d[_chk].astype(str).fillna("").agg("|".join, axis=1))
              .groupby(["예약번호", "src_file"])["_row"].apply(lambda s: "//".join(sorted(s))))
    assert _sig.groupby("예약번호").nunique().eq(1).all(), \
        "겹치는 예약의 값이 파일 간에 갈립니다 - 어느 export를 고르는지가 결과를 바꿉니다"

    resv = mart.groupby("예약번호", as_index=False).agg(agg)

    # 옛 기준(max) 값을 대조용으로 함께 보존합니다. 기존 분석과 원본 분석 문서 수치가
    # 전부 max 기준이라, 이 컬럼이 없으면 어느 수치가 왜 달라졌는지 못 따집니다.
    resv = resv.merge(mart.groupby("예약번호", as_index=False)["결제금액"].max()
                          .rename(columns={"결제금액":"listed_max"}), on="예약번호", how="left")

    resv = resv.rename(columns={
        "이용일시_dt":"use_dt","예약신청일시_dt":"req_dt","예약확정일시_dt":"conf_dt",
        "이용완료일시_dt":"done_dt","예약취소일시_dt":"cancel_dt","스타일리스트":"designer",
        "실결제금액":"paid","결제금액":"listed","상태":"status","유입경로":"channel"})

    # 파생
    resv["ym"]      = resv["use_dt"].dt.to_period("M")
    resv["dow"]     = resv["use_dt"].dt.dayofweek
    resv["hour"]    = resv["use_dt"].dt.hour
    resv["lead_h"]  = (resv["use_dt"] - resv["req_dt"]).dt.total_seconds()/3600
    resv["req_h"]   = resv["req_dt"].dt.hour
    # 고객키: 2025파일은 이름 마스킹, 2026파일은 대부분 비마스킹 -> 성+전화뒤4자리
    resv["cust_key"] = (resv["예약자"].astype(str).str[0] + "|" +
                        resv["전화번호"].astype(str).str[-4:])
    resv["has_request"] = resv["요청사항"].notna()

    # ── FIX-3a (2026-08) 수용 기준 ───────────────────────────────────────────
    # ⚠️ 이 블록을 지우지 마세요. 원본 파일이 교체되면 값을 갱신해야 합니다.
    #    괄호 안은 옛 max 기준 값입니다. sum -> max로 회귀하면 그 숫자가 나옵니다.
    assert resv["예약번호"].nunique() == 4_281, f'예약 {resv["예약번호"].nunique():,} != 4,281'
    _d = resv[resv["status"] == "이용완료"]
    assert len(_d) == 2_732, f"이용완료 {len(_d):,} != 2,732"
    assert _d["listed"].sum()     == 285_217_000, \
        f'완료 listed {_d["listed"].sum():,} != 285,217,000 (max 회귀 시 277,872,000)'
    assert _d["listed_max"].sum() == 277_872_000, \
        f'완료 listed_max {_d["listed_max"].sum():,} != 277,872,000'
    assert _d["paid"].sum()       == 308_975_650, \
        f'완료 paid {_d["paid"].sum():,} != 308,975,650 (sum 전환과 무관하게 불변이어야 함)'
    assert (_d["paid"] - _d["listed"]).sum() == 23_758_650, \
        f'완료 paid-listed {(_d["paid"]-_d["listed"]).sum():,} != 23,758,650 (max 기준 31,103,650)'
    # sum과 max가 갈리는 예약 = 값이 있는 옵션 라인이 2개 이상인 예약
    assert (_d["listed"] != _d["listed_max"]).sum() == 150, \
        f'완료 sum!=max {(_d["listed"]!=_d["listed_max"]).sum():,} != 150'
    assert (_d["listed"] - _d["listed_max"]).sum() == 7_345_000, \
        f'완료 sum-max {(_d["listed"]-_d["listed_max"]).sum():,} != 7,345,000'
    assert (resv["listed"] != resv["listed_max"]).sum() == 226, \
        f'전체 sum!=max {(resv["listed"]!=resv["listed_max"]).sum():,} != 226'
    assert (resv["listed"] - resv["listed_max"]).sum() == 11_780_000, \
        f'전체 sum-max {(resv["listed"]-resv["listed_max"]).sum():,} != 11,780,000'
    # (b) paid != listed 는 방향이 섞여 있습니다. 업셀(증가)만 세면 328건입니다.
    assert (_d["paid"] != _d["listed"]).sum() == 366, "완료 paid!=listed != 366건"
    assert (_d["paid"] >  _d["listed"]).sum() == 328, "완료 paid>listed != 328건"
    assert (_d["paid"] <  _d["listed"]).sum() ==  38, "완료 paid<listed != 38건"
    assert (_d["paid"] >  _d["listed_max"]).sum() == 461, "완료 paid>listed_max != 461건"

    # ── FIX-3c 수용 기준: menu_lines 와 reservations 를 금액으로 묶어 둡니다 ──
    # 예약별 라인단가 합 == listed. menu_lines에서 행이 빠지거나 단가가 어긋나면
    # 여기서 즉시 터집니다. (listed가 sum이 아니라 max로 회귀해도 터집니다.)
    _lsum = menu.groupby("예약번호")["line_price"].sum()
    _diff = resv.set_index("예약번호")["listed"].sub(_lsum, fill_value=0)
    assert (_diff == 0).all(), \
        f"예약별 라인단가 합 != listed 인 예약 {int((_diff != 0).sum()):,}건 " \
        f"(menu_lines가 매출을 흘리고 있거나 listed 집계가 max로 회귀했습니다)"
    return resv, menu

# ---------- 1-b. 가격 카탈로그 ----------
# 동일시술 의심: 이름만 갈리고 가격도 시기도 그대로 이어지는 메뉴 쌍입니다.
#   베타뿌리염색      39,000 x 277  2025-01 ~ 2025-09
#   알파,베타뿌리염색  39,000 x 313  2025-08 ~ 2026-08
# 2025-08에 메뉴명이 바뀌고 옛 이름이 사라진 형태라, 둘을 따로 세면 같은 시술의
# 재방문·매출이 두 갈래로 쪼개집니다.
# ⚠️ 자동 판정하지 않습니다. "같은 가격"만으로는 판별이 안 됩니다 - 150,000원
#    하나에만 서로 다른 메뉴가 23종 붙어 있습니다. 확인된 쌍만 손으로 적습니다.
DUP_SERVICE = {
    "베타뿌리염색":        "베타뿌리염색",
    "알파,베타뿌리염색":   "베타뿌리염색",
}


def _price_values(s):
    """단가 분포를 '88,000x56 / 58,000x19 / 78,000x4' 형태로. 건수 내림차순."""
    vc = s.value_counts().sort_index(ascending=False)          # 동수일 때 단가 내림차순
    vc = vc.sort_values(ascending=False, kind="stable")
    return " / ".join(f"{int(p):,}x{c}" for p, c in vc.items())


def build_price_catalog(menu, resv):
    """
    menu_lines 7,165행을 메뉴 94종 x 1행으로 접습니다.

    ⚠️ 여기서 나오는 price는 "정가"가 아니라 **20개월간 관측된 청구액**입니다.
       90종은 값이 하나뿐이라 정가로 봐도 되지만 4종은 갈립니다. 갈리는 4종은
       값을 하나로 고르지 않고 price=NA 로 두고 넘깁니다. 평균이나 최빈값으로
       메우면 "정가 94종 확보"처럼 보이면서 실제로는 틀린 표가 됩니다.

    price_latest 는 "현재가"가 아니라 **가장 최근 라인 한 줄의 값**입니다.
    앞머리컷은 2025-09부터 10개월간 0원인데 2026-07-13 한 건만 5,000원이라
    price_latest 가 5,000원으로 나옵니다. 현재가로 쓰려면 price_values 의
    시기 분포를 같이 보세요.
    """
    m = menu.merge(resv[["예약번호","use_dt"]], on="예약번호", how="left")
    assert m["use_dt"].notna().all(), \
        f'menu_lines {int(m["use_dt"].isna().sum()):,}행에 이용일시가 안 붙었습니다'
    m = m.sort_values("use_dt", kind="stable")                 # first/latest 를 시간순으로

    g = m.groupby("menu", sort=True)["line_price"]
    cat = pd.DataFrame({
        "n_lines":       g.size(),
        "price_nunique": g.nunique(),
        "price_min":     g.min(),
        "price_max":     g.max(),
        "price_first":   g.first(),
        "price_latest":  g.last(),
        "price_values":  g.apply(_price_values),
    })
    _ym = m.groupby("menu")["use_dt"].agg(["min","max"])
    cat["first_ym"] = _ym["min"].dt.to_period("M")
    cat["last_ym"]  = _ym["max"].dt.to_period("M")
    cat = cat.reset_index()

    cat["price_resolved"] = cat["price_nunique"] == 1
    # 미해결 4종은 Int64 의 NA. float NaN으로 두면 조용히 0이나 평균으로 채워집니다.
    cat["price"] = cat["price_min"].where(cat["price_resolved"]).astype("Int64")

    cat["dup_service_group"]   = cat["menu"].map(DUP_SERVICE).fillna("")
    cat["dup_service_suspect"] = cat["dup_service_group"] != ""

    cat = cat[["menu","n_lines","price","price_resolved","price_nunique",
               "price_values","price_first","price_latest","price_min","price_max",
               "first_ym","last_ym","dup_service_suspect","dup_service_group"]]

    # ── FIX-3c 수용 기준 ────────────────────────────────────────────────────
    # ⚠️ 이 블록을 지우지 마세요. 원본이 교체되면 값을 갱신해야 합니다.
    assert len(cat) == 94, f"메뉴 {len(cat)}종 != 94"
    assert int(cat["n_lines"].sum()) == 7_165, f'라인 합 {int(cat["n_lines"].sum()):,} != 7,165'
    assert int(cat["price_resolved"].sum()) == 90, \
        f'정가확정 {int(cat["price_resolved"].sum())}종 != 90'
    assert cat.loc[cat["price_resolved"], "price"].notna().all(), "확정인데 price가 NA입니다"
    assert cat.loc[~cat["price_resolved"], "price"].isna().all(), "미해결인데 price가 채워졌습니다"

    # 미해결 4종은 분포까지 고정합니다. 새 단가가 하나 들어와도 여기서 걸립니다.
    _unres = dict(zip(cat.loc[~cat["price_resolved"], "menu"],
                      cat.loc[~cat["price_resolved"], "price_values"]))
    assert _unres == {
        "뿌리탈색":           "88,000x56 / 58,000x19 / 78,000x4",
        "무한탈색+원하는컬러":  "150,000x57 / 200,000x6",
        "앞머리컷":           "0x30 / 5,000x12",
        "페이크블랙":         "150,000x14 / 100,000x3",
    }, f"미해결 메뉴의 단가 분포가 달라졌습니다: {_unres}"

    # 원본 분석 문서 7절 "확정 가격"과의 대조. 여기가 틀리면 문서와 데이터가 갈린 것입니다.
    _p = cat.set_index("menu")["price"]
    for _menu, _price in {"뿌리탈색이벤트": 100_000, "무한탈색이벤트": 150_000,
                          "수험생탈색": 150_000, "블랙빼기": 200_000,
                          "블랙빼기무제한": 250_000, "알파,베타뿌리염색": 39_000,
                          "베타뿌리염색": 39_000, "감마염색": 87_000,
                          "여성컷": 29_000, "남자컷": 29_000,
                          "아쿠아클리닉": 0}.items():
        assert _p[_menu] == _price, \
            f"{_menu} 확정가 {_p[_menu]} != {_price:,} (원본 분석 문서 7절)"
    # 7절의 "뿌리탈색 88,000"은 유일가가 아니라 **최신가**입니다. 58,000 -> 78,000 ->
    # 88,000 으로 두 번 올랐고 88,000은 2025-06부터입니다.
    assert _p["뿌리탈색"] is pd.NA, "뿌리탈색이 단일가로 바뀌었습니다 - 7절 문구를 고치세요"
    assert cat.loc[cat["menu"]=="뿌리탈색", "price_latest"].item() == 88_000, \
        f"뿌리탈색 최신가 != 88,000 (원본 분석 문서 7절)"

    assert set(DUP_SERVICE) <= set(cat["menu"]), \
        f"DUP_SERVICE에 데이터에 없는 메뉴가 있습니다: {set(DUP_SERVICE) - set(cat['menu'])}"
    assert int(cat["dup_service_suspect"].sum()) == 2, \
        f'동일시술 의심 {int(cat["dup_service_suspect"].sum())}종 != 2'
    _n = cat.set_index("menu")["n_lines"]
    assert (_n["알파,베타뿌리염색"], _n["베타뿌리염색"]) == (313, 277), \
        f'동일시술 쌍 라인수 {_n["알파,베타뿌리염색"]}/{_n["베타뿌리염색"]} != 313/277'
    return cat


# ---------- 2. 핸즈 (HTML을 .xls로 위장) ----------
HANDS_COLS = ["예약시간","분류","문자","상태","매출입력","고객","담당","메뉴","상세메뉴",
              "입력일","네이버결제","x1","x2","매출입력액","메모"]

def _extract_rows(html, fname):
    """
    핸즈 HTML에서 표의 <tr>을 전부 뽑아냅니다.

    주의 - 파서가 조용히 잘라먹는 사고가 두 번 있었습니다:
      · pandas.read_html  -> 100~170행에서 중단
      · lxml 6.1.x        -> 약 430행에서 중단 (6.0.x는 정상)
    둘 다 예외를 던지지 않고 "다 읽은 척" 합니다.
    그래서 원문의 <tr> 개수를 정규식으로 직접 세서 반드시 대조합니다.
    """
    from bs4 import BeautifulSoup
    expected = len(re.findall(r"<tr[\s>]", html, re.IGNORECASE))

    best, best_parser = [], None
    for parser in ("html.parser", "lxml", "html5lib"):  # html.parser는 표준 라이브러리라 항상 존재
        try:
            soup = BeautifulSoup(html, parser)
        except Exception:
            continue                                     # 미설치 파서는 건너뜀
        tables = soup.find_all("table")
        if not tables:
            continue
        rows = max(tables, key=lambda t: len(t.find_all("tr"))).find_all("tr")
        if len(rows) > len(best):
            best, best_parser = rows, parser
        if len(rows) >= expected - 1:                    # 헤더 1행 오차 허용
            break

    if len(best) < expected - 5:
        raise RuntimeError(
            f"[{fname}] 파싱 누락: 원문 <tr> {expected:,}개 중 {len(best):,}개만 읽혔습니다 "
            f"(가장 잘 읽은 파서: {best_parser}).\n"
            f"  설치된 파서가 모두 이 파일을 처리하지 못했습니다.\n"
            f"  `pip install html5lib` 후 다시 실행해 보세요."
        )
    return best, best_parser, expected


def load_hands():
    rows = []
    for path in find("hands"):
        html = path.read_bytes().decode("cp949")     # euc-kr, utf-8 아님
        trs, parser, expected = _extract_rows(html, path.name)
        print(f"      {path.name}: {len(trs):,}행 (원문 {expected:,}) · 파서 {parser}")
        for tr in trs[1:]:
            tds = tr.find_all("td")
            if len(tds) < 5: continue
            vals = [td.get_text("\n", strip=True) for td in tds]
            color = ""
            if len(tds) > 5:
                m = re.search(r"color:\s*(#[0-9a-fA-F]{6})", tds[5].get("style","") or "")
                if m: color = m.group(1).lower()
            vals = (vals + [""]*15)[:15]
            rows.append(vals + [color, 1 if tr.find("strike") else 0, path.name])
    H = pd.DataFrame(rows, columns=HANDS_COLS+["cust_color","struck","src_file"])
    H = H.drop_duplicates()
    H["dt"] = H["예약시간"].apply(parse_hands_dt)
    H = H[H["dt"].notna()].copy()
    H["is_naver"] = H["예약시간"].str.contains("네이버")
    H["channel"]  = np.where(H["is_naver"], "네이버예약", "매장직접")
    H["naver_dz"] = H["메모"].str.extract(r"\[네이버담당자:([^\]]+)\]")[0]
    H["memo_amt"] = pd.to_numeric(H["메모"].str.extract(r"미결제금액\s*:\s*(\d+)")[0], errors="coerce")
    H["ym"] = H["dt"].dt.to_period("M")
    H["cust_grade"] = H["cust_color"].map({
        "#ff0000":"빨강","#1ddb16":"초록","#ffbb00":"주황",
        "#0100ff":"파랑","#000000":"검정"}).fillna("미등록")
    return H

# ---------- 3. 월별 리포트 ----------
# openpyxl로 원본을 직독해 확인한 3단 병합 헤더 구조 (인덱스 0-base, 엑셀 열문자 병기):
#   A(0) 날짜 · B(1) 상품구분
#   예약수    C~K (2~10) : 신청 확정 취소소계 예약자취소 사업자취소 미확정자동 완료 변경 노쇼
#   수량/인원 L~T (11~19): 신청 확정 취소소계 예약자취소 사업자취소 미확정자동 완료 변경 노쇼
#   매출     U~AE(20~30): 신청 확정 취소소계 예약자취소 사업자취소 미확정자동
#                         AA(26) 이용일기준 전체매출(취소포함) · AB(27) 완료매출 · AC(28) 순매출
#                         AD(29) 변경 · AE(30) 노쇼
#   ⚠️ 매출 블록만 "완료"가 AA/AB/AC 셋으로 쪼개져 있습니다. 한 칸 밀림이 나기 쉬운 자리입니다.

# 월별 그레인에서 유효한 컬럼만.
# E/N/W(4/13/22) 취소소계가 월별로 유효한 유일한 취소 지표입니다. 예전에는 이걸 아예
# 뽑지 않고, 대신 월별 값이 아닌 F/G(5/6)를 월별 컬럼으로 실어 날랐습니다.
REP_COLS = {0:"ym_raw", 1:"designer", 2:"신청", 4:"취소소계", 8:"완료",
            13:"수량_취소소계", 22:"매출_취소소계",
            26:"전체매출_취소포함", 27:"완료매출", 28:"순매출"}

# ⚠️ 아래 9개는 월별 값이 아닙니다. (파일, 상품구분)별 15개월 윈도우 총계가
#    임의의 한 행에 통째로 덤프돼 있습니다. 근거: 리포트_1의 2025-06 DESIGNER_A 행은
#    취소소계(E)=36인데 예약자취소(F)=638입니다. 하위항목이 소계보다 클 수 없습니다.
#    월별 테이블에 두면 안 되므로 report_period_totals 로 분리합니다.
DUMP_COLS = {5:"예약수_예약자취소",  6:"예약수_사업자취소",  7:"예약수_미확정자동",
             14:"수량_예약자취소",   15:"수량_사업자취소",   16:"수량_미확정자동",
             23:"매출_예약자취소",   24:"매출_사업자취소",   25:"매출_미확정자동"}

# 취소 소계 <-> 세부 3항목 대응. 파일 단위로만 합이 맞습니다(행 단위로는 안 맞습니다).
CANCEL_GRP = {"예약수": (4, [5, 6, 7]), "수량": (13, [14, 15, 16]), "매출": (22, [23, 24, 25])}

# 월별로 유효한 전 컬럼. 덤프 9개를 뺀 나머지 전부이며 파일 간 완전일치해야 합니다.
MONTH_VALID_COLS = [2, 3, 4, 8, 9, 10, 11, 12, 13, 17, 18, 19, 20, 21, 22, 26, 27, 28, 29, 30]

# ⚠️ 원본 파일이 교체/추가되면 이 기대값 블록을 갱신해야 합니다.
EXPECT_CANCEL_TOTAL = {
    "월별_리포트_1.xlsx": {"예약수": 1_220, "수량": 1_220, "매출": 145_205_000},
    "월별_리포트_2.xlsx": {"예약수": 1_220, "수량": 1_220, "매출": 143_900_000},
    "월별_리포트_3.xlsx": {"예약수": 1_188, "수량": 1_188, "매출": 138_811_000},
}


def load_report():
    """
    월별 리포트 3개를 두 테이블로 나눠 반환합니다.

      rep           : (ym, designer) 그레인. 월별로 유효한 컬럼만 남깁니다.
      period_totals : (src_file, designer) 그레인. 취소 세부 3항목의 윈도우 전 기간 총계.
                      월 컬럼을 붙이지 않습니다 - 애초에 월 정보가 없는 값입니다.
                      기간은 window_start / window_end 로만 표현합니다.
    """
    frames = []
    for path in find("report"):
        df = pd.read_excel(path, header=None, skiprows=3)   # 3단 헤더
        df = df[df[0].notna()].copy()                       # 날짜 없는 행 제거
        df["src_file"] = path.name
        frames.append(df)
    raw_report = pd.concat(frames, ignore_index=True)
    raw_report["ym"] = pd.PeriodIndex(
        pd.to_datetime(raw_report[0].astype(str), format="%Y.%m.%d"), freq="M")

    dump_idx = list(DUMP_COLS)
    # 결측이 있으면 아래 합계 검증이 조용히 통과합니다 (pandas sum은 NaN을 건너뜁니다).
    assert raw_report[dump_idx + [4, 13, 22]].notna().to_numpy().all(), "취소 컬럼에 결측이 있습니다"

    # (c) 취소소계 == 세부 3항목 합. 파일 단위로만 성립합니다.
    for grp, (tot, parts) in CANCEL_GRP.items():
        for src, d in raw_report.groupby("src_file"):
            a, b = d[tot].sum(), d[parts].sum().sum()
            assert a == b, f"{src} {grp}: 취소소계 {a:,.0f} != 세부합 {b:,.0f}"
            exp = EXPECT_CANCEL_TOTAL.get(src, {}).get(grp)
            assert exp is None or a == exp, f"{src} {grp}: {a:,.0f} != 기대값 {exp:,}"

    # (d) 덤프행 표시. 삭제하지 않고 마스킹만 합니다.
    #     is_dump_row : 덤프 9컬럼 중 하나라도 0이 아닌 행. 이게 완전한 마스크입니다(파일당 12행).
    #     dump_f_gt_e : 검증 리포트가 쓴 (F > E) 정의. 파일당 정확히 7행.
    #     ⚠️ 둘은 다릅니다. F 덤프와 G 덤프가 서로 다른 행에 떨어지기 때문에
    #        (F > E)는 G에만 덤프된 행 15개를 놓치고, 마스킹 후에도 37,241,000원이 남습니다.
    #        분석에서 덤프를 걸러낼 때는 반드시 is_dump_row 를 쓰세요.
    raw_report["is_dump_row"] = (raw_report[dump_idx] != 0).any(axis=1)
    raw_report["dump_f_gt_e"] = raw_report[5] > raw_report[4]
    _fge = raw_report.groupby("src_file")["dump_f_gt_e"].sum()
    _dmp = raw_report.groupby("src_file")["is_dump_row"].sum()
    assert (_fge == 7).all(),  f"(F>E) 덤프행이 파일당 7행이 아님: {_fge.to_dict()}"
    assert (_dmp == 12).all(), f"덤프행이 파일당 12행이 아님: {_dmp.to_dict()}"
    assert raw_report.loc[~raw_report["is_dump_row"], dump_idx].to_numpy().sum() == 0, \
        "덤프 마스크 밖에 취소 세부값이 남아 있습니다 - is_dump_row 정의를 다시 보세요"

    # (b) 각 덤프 컬럼은 (파일, 상품구분)당 0이 아닌 행이 최대 1개입니다.
    #     따라서 sum()이 곧 그 값이고, 어느 행에 떨어졌는지 몰라도 안전하게 뽑힙니다.
    for c in dump_idx:
        n = raw_report[raw_report[c] != 0].groupby(["src_file", 1]).size()
        assert n.empty or n.max() == 1, f"col{c}: (파일,상품구분)당 덤프행 2개 이상 - 전제 붕괴"
    window = (raw_report.groupby("src_file")["ym"]
                .agg(window_start="min", window_end="max").reset_index())
    period_totals = (raw_report.groupby(["src_file", 1])[dump_idx].sum()
                       .rename(columns=DUMP_COLS).reset_index()
                       .rename(columns={1: "designer"})
                       .merge(window, on="src_file"))

    # (g) 덤프 9컬럼을 뺀 나머지는 파일 간 완전일치해야 합니다. 가드가 없던 자리입니다.
    raw_report["_key"] = raw_report["ym"].astype(str) + "|" + raw_report[1].astype(str)
    multi = raw_report.groupby("_key")["src_file"].nunique()
    keys2, keys3 = multi[multi > 1].index, multi[multi == 3].index
    ov = raw_report[raw_report["_key"].isin(keys2)]
    assert len(keys2) == 120, f"파일 간 겹치는 (ym,상품구분) {len(keys2)} != 120 (검사가 비어버림)"
    for c in MONTH_VALID_COLS:
        piv = ov.pivot_table(index="_key", columns="src_file", values=c, aggfunc="first")
        n_bad = int((piv.nunique(axis=1) > 1).sum())
        assert n_bad == 0, f"col{c}: 월별 유효 컬럼이 파일 간 불일치 {n_bad}조합"

    # 덤프 컬럼은 반대로 파일마다 다릅니다(윈도우가 다르니 당연). 건수만 로그로 남깁니다.
    _log = []
    for c in (5, 6):
        p2 = ov.pivot_table(index="_key", columns="src_file", values=c, aggfunc="first")
        p3 = (raw_report[raw_report["_key"].isin(keys3)]
                .pivot_table(index="_key", columns="src_file", values=c, aggfunc="first"))
        _log.append(f"{DUMP_COLS[c]} {int((p2.nunique(axis=1) > 1).sum())}"
                    f"/{int((p3.nunique(axis=1) > 1).sum())}")
    print(f"      덤프컬럼 파일간 충돌 (겹침 {len(keys2)}조합 / 3파일공통 {len(keys3)}조합 기준): "
          + " · ".join(_log))

    # (f) dedup 우선순위. 예전에는 파일명 정렬 + keep="first" 였는데, _1 < _2 < _3 이
    #     우연히 최신순이라 맞았을 뿐입니다. 월별_리포트_4.xlsx가 들어오면 정렬
    #     마지막이 되어 우선순위가 뒤집힙니다. 파일명이 아니라 각 파일의 윈도우 최대
    #     ym(=최신 export)을 기준으로 고릅니다. 현재 데이터에서는 결과가 동일합니다.
    rep = (raw_report[list(REP_COLS) + ["ym", "src_file", "is_dump_row", "dump_f_gt_e"]]
             .rename(columns=REP_COLS).drop(columns=["ym_raw"]))
    rep["_prio"] = rep["src_file"].map(raw_report.groupby("src_file")["ym"].max())
    rep = (rep.sort_values("_prio", ascending=False, kind="stable")
              .drop_duplicates(subset=["ym", "designer"], keep="first")
              .drop(columns=["_prio"])
              .sort_values(["ym", "designer"]).reset_index(drop=True))

    # ── FIX-1 (2026-08): "각 파일 첫 월 제거"를 삭제했습니다 ─────────────────
    # 예전 코드:  firsts = rep.groupby("src_file")["ym_raw"].first().to_dict()
    #             rep = rep[[r.ym_raw != firsts[r.src_file] for r in rep.itertuples()]]
    # "각 파일 첫 달에 직전 기간 취소분이 통째로 몰린다"는 전제였는데 셋 다 틀렸습니다.
    #   (1) 첫 월은 오염돼 있지 않다. 첫 월의 취소소계·완료·완료매출·순매출은
    #       다른 파일의 같은 월 값과 완전일치한다 (위 (g) assert가 이걸 고정합니다).
    #   (2) 잡으려던 덤프행을 못 잡는다. 덤프행은 "그 상품구분이 파일에서 처음
    #       등장하는 행"에 박히지 첫 월에 박히지 않는다. 리포트_1의 덤프행 7개 중
    #       4개(2025-08 첫방문50%·디자이너 1명, 2026-02 디자이너 2명)가 그대로 살아남았다.
    #   (3) 2025-04를 영구 손실시킨다. 2025-04는 리포트_3에만 있는데 하필 그
    #       파일의 첫 월이라 통째로 사라졌다.
    # 삭제 효과: 16개월 128행 -> 17개월 136행, 완료 +105건, 완료매출 +9,300,000원,
    #            순매출 +3,261,500원.
    # ⚠️ 아래 assert를 지우지 마세요. 깨졌다면 첫 월 제거가 되살아났거나 원본이 바뀐 것입니다.
    #    (Period는 문자열과 == 비교가 항상 False라 str()로 감쌉니다.)
    n_ym = rep["ym"].nunique()
    assert n_ym == 17, f"월 수 {n_ym} != 17 (첫 월 제거가 되살아났는지 확인)"
    assert str(rep["ym"].min()) == "2025-04", f'최소월 {rep["ym"].min()} != 2025-04'
    assert len(rep) == 136, f"리포트 행 수 {len(rep)} != 136 (17개월 x 8상품구분)"
    assert rep["완료"].sum()     == 2_339,       f'완료 {rep["완료"].sum():,.0f} != 2,339 (첫월제거 시 2,234)'
    assert rep["완료매출"].sum() == 271_120_150, f'완료매출 {rep["완료매출"].sum():,.0f} != 271,120,150 (첫월제거 시 261,820,150)'
    assert rep["순매출"].sum()   ==  64_381_650, f'순매출 {rep["순매출"].sum():,.0f} != 64,381,650 (첫월제거 시 61,120,150)'

    # ── FIX-2 (2026-08): 취소소계 추가 + 덤프 9컬럼 분리 ─────────────────────
    assert rep["신청"].sum()          ==       3_668, f'신청 {rep["신청"].sum():,.0f} != 3,668'
    assert rep["취소소계"].sum()      ==       1_319, f'취소소계 {rep["취소소계"].sum():,.0f} != 1,319'
    assert rep["수량_취소소계"].sum() ==       1_319, f'수량_취소소계 {rep["수량_취소소계"].sum():,.0f} != 1,319'
    assert rep["매출_취소소계"].sum() == 155_155_000, f'매출_취소소계 {rep["매출_취소소계"].sum():,.0f} != 155,155,000'
    assert len(period_totals) == 24, f"period_totals {len(period_totals)}행 != 24 (3파일 x 8상품구분)"
    for src, d in period_totals.groupby("src_file"):
        for grp, (_tot, parts) in CANCEL_GRP.items():
            exp = EXPECT_CANCEL_TOTAL.get(src, {}).get(grp)
            got = d[[DUMP_COLS[p] for p in parts]].to_numpy().sum()
            assert exp is None or got == exp, f"period_totals {src} {grp}: {got:,.0f} != {exp:,}"
    return rep, period_totals

# ---------- 4. 핸즈 <-> 네이버 조인 ----------
# 이름 마스킹 대조용: 네이버 '홍*동' vs 핸즈 '홍길동(283)' -> 같은 사람일 수 있는가
# (주석 예시에도 실제 고객 값을 쓰지 말 것 — 원본 분석 문서 §10)
_RE_PAREN = re.compile(r"\(.*?\)")
_RE_KEEP   = re.compile(r"[^가-힣*]")

def _clean_name(x):
    return _RE_KEEP.sub("", _RE_PAREN.sub("", str(x)))

def _name_compat(naver_name, hands_name):
    """네이버 예약자/방문자명과 핸즈 실명이 양립 가능한가. '*'는 와일드카드."""
    a, b = _clean_name(naver_name), _clean_name(hands_name)
    if len(a) < 2 or len(b) < 2 or len(a) != len(b):
        return False
    return all(x == "*" or x == y for x, y in zip(a, b))


def _assign_1to1(cand_ids, cand_names, hands_names):
    """
    한 키(동일 이용일시+담당[+금액])에 여러 예약과 여러 핸즈행이 걸린 경우의 배정.

    ⚠️ 예전 구현은 `drop_duplicates("k2")`로 **항상 첫 예약에만** 붙였습니다.
       그 결과 (1) 핸즈행 98개가 같은 예약번호에 중복 배정돼 조용히 버려지고
             (2) 실제로는 옆 슬롯 손님인 핸즈행이 엉뚱한 예약에 붙었습니다.
       성 불일치율이 k2 유일 1.1% vs k2 동률 16.4% 로 15배 차이났습니다.

    지금은 (a) 실명이 양립하는 짝을 먼저 확정하고
           (b) 남은 것만 순서대로 1:1 배정하며
           (c) 이름 근거 없이 배정된 건 ambiguous=True 로 표시합니다.

    반환: [(hands 위치, 예약번호, ambiguous)] 리스트
    """
    n_cand = len(cand_ids)
    free = list(range(n_cand))
    out, leftover = [], []

    for hi, hname in enumerate(hands_names):
        hits = [j for j in free
                if any(_name_compat(nm, hname) for nm in cand_names[j])]
        if len(hits) == 1:
            j = hits[0]
            free.remove(j)
            out.append((hi, cand_ids[j], False))
        else:
            leftover.append(hi)

    ambiguous = n_cand > 1          # 후보가 애초에 여럿이면 이름 근거 없는 배정은 못 믿는다
    for hi in leftover:
        if not free:
            break
        j = free.pop(0)
        out.append((hi, cand_ids[j], ambiguous))
    return out


def link(resv, H):
    """
    공통 PK가 없으므로 (이용일시 + 담당자) -> 금액으로 tie-break 하는 사다리 조인.
    같은 키에 예약이 여러 건이면 실명으로 다시 tie-break 하고, 그래도 못 가르면
    배정은 하되 match_ambiguous=True 로 표시합니다. (분석 시 제외할 수 있게)
    """
    r = resv.copy()
    r["k1"] = r["use_dt"].astype(str) + "|" + r["designer"].astype(str)
    r["k2"] = r["k1"] + "|" + r["listed"].fillna(0).astype(float).round(0).astype(int).astype(str)

    hn = H[H["is_naver"]].copy().reset_index(drop=True)
    hn["k1"] = hn["dt"].astype(str) + "|" + hn["naver_dz"].astype(str)
    hn["k2"] = hn["k1"] + "|" + hn["memo_amt"].fillna(0).round(0).astype(int).astype(str)

    rid = np.full(len(hn), np.nan, dtype=object)
    amb = np.zeros(len(hn), dtype=bool)
    lvl = np.array(["미매칭"] * len(hn), dtype=object)

    name_cols = [c for c in ("예약자", "방문자") if c in r.columns]
    hands_name = hn["고객"].tolist() if "고객" in hn.columns else [""] * len(hn)

    # ⚠️ 이 루프는 "핸즈행이 아직 안 붙었는지"(pd.isna(rid))만 확인하고,
    #    "예약이 1단에서 이미 찼는지"는 확인하지 않습니다. 소진 관리가 한쪽뿐입니다.
    #
    #    k1이 유일한 예약은 k2도 유일합니다(k2 = k1 + 금액). 따라서
    #      1단: 예약 R <- 핸즈행 A (금액 일치)
    #      2단: 같은 dt+담당이지만 memo_amt가 달라 1단에 실패한 핸즈행 B가
    #           len(cand)==1 조건을 통과해 같은 예약 R에 또 배정됨
    #    이 되면 한 예약번호에 핸즈행 2개가 붙습니다. main()의 add에서
    #    drop_duplicates("예약번호")로 하나가 조용히 사라지는, 98행을 날렸던 그 형태입니다.
    #    게다가 cand가 1개라 _assign_1to1은 ambiguous=False로 표시합니다(경고도 안 뜸).
    #
    #    현재 데이터에선 실제 발생 0건이고, main()의 "조인 중복배정" 줄이 이를 감시합니다.
    #    그 줄이 0이 아니게 되면 여기를 고치세요 - 배정 완료된 예약번호 집합 used를 들고
    #    cand에서 빼면 됩니다.
    for key, level in (("k2", "일시+담당+금액"), ("k1", "일시+담당")):
        r_groups = r.groupby(key)
        h_groups = hn[pd.isna(rid)].groupby(key)          # 아직 못 붙은 것만
        for k, hg in h_groups:
            if k not in r_groups.groups:
                continue
            cand = r_groups.get_group(k)
            if key == "k1" and len(cand) > 1:
                continue                                   # 2단은 유일할 때만 (기존 규칙 유지)
            cand_ids   = cand["예약번호"].tolist()
            cand_names = [[c[n] for n in name_cols] for _, c in cand.iterrows()]
            pos = hg.index.tolist()
            for hi, r_id, is_amb in _assign_1to1(
                    cand_ids, cand_names, [hands_name[p] for p in pos]):
                p = pos[hi]
                rid[p], amb[p], lvl[p] = r_id, is_amb, (
                    level if key == "k2" else "일시+담당(유일)")

    hn["예약번호"] = rid
    hn["match_level"] = lvl
    hn["match_ambiguous"] = amb
    return hn

# ---------- 실행 ----------
def main():
    OUT.mkdir(parents=True, exist_ok=True)
    print(f"원본 폴더: {RAW.resolve()}")
    for k in PATTERNS:
        print(f"  {k}: {[p.name for p in find(k)]}")
    print()
    print("[1/4] 네이버 예약자관리 복호화 및 로드")
    resv, menu = load_naver()
    print(f"      예약 {len(resv):,}건 · 메뉴행 {len(menu):,}건 · "
          f"{resv['use_dt'].min():%Y-%m-%d} ~ {resv['use_dt'].max():%Y-%m-%d}")
    pricecat = build_price_catalog(menu, resv)
    print(f"      가격 카탈로그 {len(pricecat)}종 · 정가확정 "
          f"{int(pricecat['price_resolved'].sum())}종 / 미해결 "
          f"{int((~pricecat['price_resolved']).sum())}종 · 동일시술 의심 "
          f"{int(pricecat['dup_service_suspect'].sum())}종")

    print("[2/4] 핸즈 HTML 파싱")
    H = load_hands()
    print(f"      {len(H):,}행 (네이버유래 {H['is_naver'].sum():,} / 매장직접 {(~H['is_naver']).sum():,})")

    print("[3/4] 월별 리포트")
    rep, period_totals = load_report()
    print(f"      월별 {len(rep):,}행 · {rep['ym'].min()} ~ {rep['ym'].max()}"
          f" · 기간총계 {len(period_totals):,}행 (src_file x 상품구분)")

    print("[4/4] 핸즈 <-> 네이버 조인")
    hn = link(resv, H)
    vc = hn["match_level"].value_counts()
    for k, v in vc.items(): print(f"      {k}: {v:,}행 ({v/len(hn)*100:.1f}%)")

    # ── FIX-3a 부수효과 (2026-08) ────────────────────────────────────────────
    # listed를 max -> sum으로 바꾸자 k2(일시+담당+금액) 키가 핸즈 메모의 '미결제금액'과
    # 맞아떨어지기 시작했습니다. 핸즈 메모 금액은 예약 전체 청구액이므로, 옵션 라인
    # 하나의 단가(max)가 아니라 합계(sum)가 처음부터 맞는 값이었습니다.
    #   1단 3,047 -> 3,203 · 2단 135 -> 0 · 미매칭 25 -> 4 · 잃은 행 0 · 중복배정 0 유지
    #   신규 21행 중 핸즈 실명이 있는 9행은 전부 이름 일치, 불일치 0.
    # 2단(k1) 사다리는 이제 한 행도 쓰지 않습니다. 아래 값이 흔들리면 load_naver의
    # 결제금액 집계가 sum에서 max로 되돌아갔는지 먼저 의심하세요.
    assert vc.get("일시+담당+금액", 0) == 3_203, \
        f"1단 {vc.get('일시+담당+금액', 0):,} != 3,203 (max 회귀 시 3,047)"
    assert vc.get("일시+담당(유일)", 0) == 0, \
        f"2단 {vc.get('일시+담당(유일)', 0):,} != 0 (max 회귀 시 135)"
    assert vc.get("미매칭", 0) == 4, f"미매칭 {vc.get('미매칭', 0):,} != 4 (max 회귀 시 25)"

    # 조인 결과를 예약 테이블에 병합 (핸즈 고유 필드만)
    add = (hn[hn["예약번호"].notna()]
             .drop_duplicates("예약번호")
             .set_index("예약번호")[["고객","cust_grade","상태","매출입력",
                                     "match_level","match_ambiguous"]]
             .rename(columns={"고객":"hands_name","상태":"hands_status",
                              "매출입력":"hands_sales_entered",
                              "match_level":"hands_match_level",
                              "match_ambiguous":"hands_match_ambiguous"}))
    resv = resv.merge(add, left_on="예약번호", right_index=True, how="left")

    for name, df in [("reservations", resv), ("menu_lines", menu),
                     ("price_catalog", pricecat),
                     ("hands_all", H), ("monthly_report", rep),
                     ("report_period_totals", period_totals)]:
        p = OUT/f"{name}.parquet"
        df.to_parquet(p, index=False)
        print(f"저장: {p}  ({len(df):,}행 x {df.shape[1]}컬럼)")

    # FIX-3c 에서 menu_items.parquet -> menu_lines.parquet 로 개명했습니다.
    # 옛 파일을 남겨 두면 분석 스크립트가 단가 없는 구버전을 계속 읽습니다.
    _old = OUT/"menu_items.parquet"
    if _old.exists():
        _old.unlink()
        print(f"삭제: {_old}  (menu_lines.parquet 으로 대체됨)")

    # 검증
    print("\n--- 검증 ---")
    done = resv[resv["status"]=="이용완료"]
    calc = done.groupby([done["ym"], "designer"])["paid"].sum()
    ref  = rep.set_index(["ym","designer"])["완료매출"]
    cmp  = pd.concat([calc.rename("calc"), ref.rename("ref")], axis=1).dropna()
    ok = (cmp["calc"]-cmp["ref"]).abs() < 1000
    print(f"월x디자이너 매출 대조: {len(cmp)}셀 중 일치 {ok.sum()} ({ok.mean()*100:.0f}%)")
    print(f"핸즈에만 있는 매장직접 예약: {(~H['is_naver']).sum()}건 (네이버 데이터에 부재)")
    print(f"핸즈 조인 성공: {resv['hands_name'].notna().sum():,}건")
    # .fillna("") 필수. pandas 2.x는 astype(str)이 NaN을 문자열 "nan"(길이 3)으로 바꿔
    # 미조인 1,099건까지 실명으로 세어 1,549 -> 2,648이 됩니다. (3.x는 NaN을 보존)
    nm = resv["hands_name"].fillna("").astype(str).str.len() >= 2
    print(f"  실제 실명 확보: {nm.sum():,}건  (핸즈 고객명 입력률 약 50%)")
    ok = hn[hn["예약번호"].notna()]
    print(f"조인 중복배정: {len(ok)-ok['예약번호'].nunique()}행 (0이어야 함)")
    print(f"조인 모호(match_ambiguous): {int(ok['match_ambiguous'].sum()):,}행 "
          f"({ok['match_ambiguous'].mean()*100:.1f}%) - 고객 단위 분석에서 제외 권장")
    nmd = resv[nm]
    bad = (nmd["hands_name"].astype(str).str[0] != nmd["예약자"].astype(str).str[0]) &           (nmd["hands_name"].astype(str).str[0] != nmd["방문자"].astype(str).str[0])
    print(f"핸즈 실명 vs 네이버 성 불일치: {int(bad.sum())}행 / {len(nmd):,}행 "
          f"({bad.mean()*100:.2f}%) - 대부분 대리예약")

    # ── FIX-3c (2026-08): 단가가 갈리는 메뉴 ────────────────────────────────
    # 4종 전부 "가격 개정"으로 읽힙니다. 시기가 거의 겹치지 않습니다.
    _u = pricecat[~pricecat["price_resolved"]]
    print(f"단가 미해결 {len(_u)}종 (나머지 {int(pricecat['price_resolved'].sum())}종은 단일가):")
    for r in _u.itertuples():
        print(f"  {r.menu:<20s} {r.price_values:<38s} "
              f"{r.first_ym}~{r.last_ym}  최신 {r.price_latest:,}원")

    # ── FIX-3b (2026-08): paid - listed 는 방향이 섞여 있습니다 ──────────────
    # 옛 리포트는 493건을 통째로 "현장 업셀"이라 불렀는데, 그 안에 반대 방향
    # (할인·환불)이 섞여 있었습니다. 순차액만 보면 두 방향이 상쇄돼 규모를 놓칩니다.
    up = done[done["paid"] >  done["listed"]]
    dn = done[done["paid"] <  done["listed"]]
    eq = int((done["paid"] == done["listed"]).sum())
    print(f"결제 일치(paid==listed): {eq:,}건 / {len(done):,}건 "
          f"({eq/len(done)*100:.1f}%)  [옛 max 기준 82.0%]")
    print(f"  현장 추가 추정(paid>listed): {len(up):,}건  +{(up['paid']-up['listed']).sum():,}원")
    print(f"  반대방향(paid<listed):      {len(dn):,}건  -{(dn['listed']-dn['paid']).sum():,}원"
          f"  <- 할인·환불")
    print(f"  순차액 {(done['paid']-done['listed']).sum():,}원 "
          f"(옛 max 기준 {(done['paid']-done['listed_max']).sum():,}원, "
          f"paid>listed_max {int((done['paid']>done['listed_max']).sum()):,}건)")

if __name__ == "__main__":
    main()
