"""
핸즈 파일 진단 — 파서가 잘라먹은 것인지, 파일 자체가 짧은 것인지 구분합니다.
build.py 와 같은 폴더에 두고 실행하세요.

실행:  python check_hands.py
"""
import os, re, sys
from pathlib import Path
from bs4 import BeautifulSoup

RAW = Path(os.environ.get("SALON_RAW", "./data/raw"))

def find_hands():
    hits = []
    for p in sorted(RAW.iterdir()):
        if p.is_dir() or p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in (".xls", ".xlsx"):
            continue
        if "예약상세목록" in p.name.replace(" ", "").replace("_", ""):
            hits.append(p)
    return hits

files = find_hands()
if not files:
    print(f"*** {RAW.resolve()} 에서 예약상세목록 파일을 못 찾았습니다 ***")
    sys.exit(1)

print("=" * 70)
print("핸즈 파일 진단")
print("=" * 70)
print(f"경로: {RAW.resolve()}\n")

import bs4, lxml.etree
print(f"beautifulsoup4 {bs4.__version__} / lxml {lxml.etree.__version__}\n")

grand = {}
for p in files:
    raw = p.read_bytes()
    print("-" * 70)
    print(f"파일: {p.name}")
    print(f"크기: {len(raw)/1024:,.0f} KB")

    try:
        html = raw.decode("cp949")
    except UnicodeDecodeError as e:
        print(f"*** cp949 디코딩 실패: {e} ***")
        continue

    print(f"디코딩 후 글자수: {len(html):,}")

    # 1) 원문에서 직접 <tr> 세기 = 진짜 행 수
    raw_tr = len(re.findall(r"<tr[\s>]", html, re.IGNORECASE))
    print(f"\n  [기준] HTML 원문의 <tr> 개수 : {raw_tr:,}")

    # 2) 각 파서가 실제로 찾아낸 행 수
    for parser in ["lxml", "html.parser", "html5lib"]:
        try:
            soup = BeautifulSoup(html, parser)
            tables = soup.find_all("table")
            if not tables:
                print(f"  {parser:12s} : 테이블 없음")
                continue
            best = max(tables, key=lambda t: len(t.find_all("tr")))
            n = len(best.find_all("tr"))
            total = sum(len(t.find_all("tr")) for t in tables)
            gap = raw_tr - total
            mark = "  <-- 잘림!" if gap > 5 else ""
            print(f"  {parser:12s} : 최대테이블 {n:>6,}행 · 전체 {total:>6,}행 · "
                  f"누락 {gap:>6,}{mark}")
            grand.setdefault(parser, 0)
            grand[parser] += n
        except Exception as e:
            print(f"  {parser:12s} : 사용 불가 ({type(e).__name__}: {e})")

    # 3) 파일이 담고 있는 실제 날짜 범위
    dates = re.findall(r"(\d{2})-(\d{2})-(\d{2})\s*\n?\s*.\s*\n?\s*\d{1,2}:\d{2}", html)
    if dates:
        ds = sorted(f"20{y}-{m}-{d}" for y, m, d in dates)
        print(f"\n  파일이 담은 기간 : {ds[0]} ~ {ds[-1]}  (날짜 표기 {len(ds):,}개)")
    print()

print("=" * 70)
print("파서별 합계 (최대 테이블 기준)")
for k, v in grand.items():
    print(f"  {k:12s} {v:,}행")
print()
print("판정 방법")
print("  · 원문 <tr> 개수 ≈ 파싱 행 수  ->  파일 자체가 짧음 (사장님 내려받기 범위 문제)")
print("  · 원문 <tr> 개수 >> 파싱 행 수  ->  파서가 잘라먹음 (코드 수정 필요)")
print("  · 참고: 받아야 할 값은 2025 + 2026 합계 3,643행 + 헤더")
print("=" * 70)
