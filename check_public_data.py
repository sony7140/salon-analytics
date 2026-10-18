"""공개 데이터로 build.py의 수용 기준 32개와 저장된 단가표를 확인한다.

build.py를 import하거나 원본 입력을 읽지 않는다. AST에서 공개 테이블로
평가할 수 있는 문장만 꺼내 실행하며 기대값은 build.py에서 읽는다.
원본 파일 병합, 월별 리포트, POS 원문 조인에 관한 33개 기준은 검증할 수
없으므로 목록과 이유를 출력한다. 개인정보 익명화 자체를 검사하지는 않는다.

python check_public_data.py
python check_public_data.py --selftest  # 완료 예약 paid 한 칸 +1원, 메모리에서만

AST로 선택한 코드를 실행하므로 신뢰하는 저장소 사본에서 사용한다.
"""
import argparse
import ast
import re
import sys
from pathlib import Path

import pandas as pd


EXPECTED_IDS = {f"BUILD-{n:03}" for n in [*range(6, 10), *range(14, 42)]}
BUILTINS = {"int", "len", "set", "sum", "dict", "zip", "list", "str", "pd"}


def execute(node, namespace):
    """최적화 옵션으로 assert가 제거되지 않도록 평가한다."""
    exec(compile(ast.Module([node], []), "<build.py>", "exec", optimize=0), namespace)


def evaluate(node, namespace):
    return eval(compile(ast.Expression(node), "<build.py>", "eval"), namespace)


def loaded_names(node):
    return {n.id for n in ast.walk(node) if isinstance(n, ast.Name)
            and isinstance(n.ctx, ast.Load)}


def format_value(value):
    if isinstance(value, tuple):
        return repr(tuple(v.item() if hasattr(v, "item") else v for v in value))
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, int) and not isinstance(value, bool):
        return f"{value:,}"
    return repr(value)


def load_contract(root, tree):
    contract = {}
    pattern = r"\| (BUILD-\d{3}) \| (\d+) \| (\d+) \| (\w+) \| (.*?) \| (.*?) \|"
    for line in (root / "pipeline_contract.md").read_text(encoding="utf-8").splitlines():
        match = re.match(pattern, line)
        if match:
            bid, _, lineno, function, description, expected = match.groups()
            if int(lineno) in contract:
                raise ValueError("계약 표에 중복된 now 줄 번호")
            contract[int(lineno)] = (bid, function, description, expected)
    assert_lines = {n.lineno for n in ast.walk(tree) if isinstance(n, ast.Assert)}
    ids = {row[0] for row in contract.values()}
    if set(contract) != assert_lines or ids != {f"BUILD-{n:03}" for n in range(1, 66)}:
        raise ValueError("계약 표의 65개 번호·now 줄과 build.py의 assert가 다름")
    return contract


def verify(root, reservations=None):
    """기존 수용 기준의 AST 선택·실행 방식으로 검사 결과를 반환한다."""
    source = (root / "build.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    contract = load_contract(root, tree)
    functions = {n.name: n for n in tree.body if isinstance(n, ast.FunctionDef)}
    data = root / "data" / "processed"
    resv = pd.read_parquet(data / "reservations.parquet") if reservations is None else reservations.copy(deep=True)
    menu = pd.read_parquet(data / "menu_lines.parquet")
    saved = pd.read_parquet(data / "price_catalog.parquet")
    rows = []

    def check(node, namespace, function):
        bid, declared_function, description, expected = contract[node.lineno]
        if declared_function != function:
            raise ValueError(f"{bid}: 계약 표와 실제 함수가 다름")
        number = expected.replace(",", "")
        segment = re.sub(r"\s+", " ", ast.get_source_segment(source, node))
        if re.fullmatch(r"\d+", number) and number not in segment.replace("_", ""):
            raise ValueError(f"{bid}: 계약 기대값과 assert 상수가 다름")
        test = node.test
        rhs, value = "(참이어야 함)", "—"
        if isinstance(test, ast.Compare) and len(test.ops) == 1 and isinstance(test.ops[0], (ast.Eq, ast.Is)):
            rhs = ast.get_source_segment(source, test.comparators[0])
            try:
                value = format_value(evaluate(test.left, namespace))
            except Exception as exc:
                value = f"(계산 오류 {type(exc).__name__})"
        try:
            execute(node, namespace)
            ok = True
        except Exception:
            ok = False
        rows.append((bid, node.lineno, description, rhs, value, ok))

    namespace = {"pd": pd, "menu": menu, "resv": resv}
    available = {"menu", "resv", "_d", "_lsum", "_diff"} | BUILTINS
    for node in functions["load_naver"].body:
        if isinstance(node, ast.Assert) and loaded_names(node) <= available:
            check(node, namespace, "load_naver")
        elif (isinstance(node, ast.Assign) and len(node.targets) == 1
              and isinstance(node.targets[0], ast.Name)
              and node.targets[0].id in ("_d", "_lsum", "_diff")
              and loaded_names(node) <= ({"menu", "resv", "_d", "_lsum"} | BUILTINS)):
            execute(node, namespace)

    namespace = {"pd": pd, "menu": menu, "resv": resv}
    for node in tree.body:
        if ((isinstance(node, ast.Assign) and any(getattr(t, "id", "") == "DUP_SERVICE" for t in node.targets))
                or (isinstance(node, ast.FunctionDef) and node.name == "_price_values")):
            execute(node, namespace)
    catalog = None
    for node in functions["build_price_catalog"].body:
        if isinstance(node, ast.Assert):
            check(node, namespace, "build_price_catalog")
        elif isinstance(node, ast.For) and any(isinstance(n, ast.Assert) for n in node.body):
            assertion = next(n for n in node.body if isinstance(n, ast.Assert))
            bid, function, description, _ = contract[assertion.lineno]
            if function != "build_price_catalog":
                raise ValueError(f"{bid}: 계약 표의 함수가 다름")
            pairs = list(evaluate(node.iter, namespace))
            values = [namespace["_p"].get(key, "(없음)") for key, _ in pairs]
            try:
                execute(node, namespace)
                ok = True
            except Exception:
                ok = False
            rows.append((bid, assertion.lineno, description,
                         " · ".join(f"{v:,}" for _, v in pairs),
                         " · ".join(format_value(v) if v is not pd.NA else "<NA>" for v in values), ok))
        elif isinstance(node, ast.Return):
            catalog = evaluate(node.value, namespace)
        else:
            execute(node, namespace)
    actual_ids = {r[0] for r in rows}
    if actual_ids != EXPECTED_IDS or len(rows) != 32:
        raise ValueError("공개 데이터로 검사하는 assert 집합이 예상한 32개와 다름")
    same_catalog = (catalog is not None and list(catalog.columns) == list(saved.columns)
                    and catalog.sort_values("menu").reset_index(drop=True).equals(
                        saved.sort_values("menu").reset_index(drop=True)))
    missing = {}
    reasons = {"load_naver": "원본 파일 구조(raw · mart) 필요",
               "load_report": "월별 리포트 필요",
               "main": "핸즈 원문 조인 필요 (공개 hands_all에 조인 키 칸 없음)"}
    for bid, function, _, _ in contract.values():
        if bid not in actual_ids:
            missing.setdefault(reasons.get(function, function), []).append(bid)
    return sorted(rows), same_catalog, missing


def print_results(rows, same_catalog, missing):
    print("| BUILD | 줄 | 내용 | build.py 기대값 | 공개 데이터 값 | 결과 |")
    print("|---|---|---|---|---|---|")
    for bid, line, description, expected, value, ok in rows:
        print(f"| {bid} | {line} | {description} | {expected} | {value} | {'같음' if ok else '★다름'} |")
    print(f"\n재계산 단가표 = 저장본 (메뉴로 맞춘 값·dtype·컬럼): {same_catalog}")
    print("단가표 행 순서는 비교 대상에서 제외한다: 메뉴명 변경으로 정렬 순서가 달라질 수 있다.")
    for reason, ids in missing.items():
        print(f"대조 불가 — {reason}: {', '.join(sorted(ids))}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--selftest", action="store_true", help="메모리 사본의 paid +1원을 감지하는지 확인")
    args = parser.parse_args()
    root = Path(__file__).resolve().parent
    try:
        rows, same_catalog, missing = verify(root)
        print_results(rows, same_catalog, missing)
        ok = all(row[-1] for row in rows) and same_catalog
        if args.selftest and ok:
            resv = pd.read_parquet(root / "data/processed/reservations.parquet")
            # paid == listed 행을 고르면 금액 합 외에 두 건수 기준도 함께 깨진다.
            candidates = resv.index[(resv.status == "이용완료") & (resv.paid == resv.listed)]
            if candidates.empty:
                raise ValueError("자기검사용 완료 예약을 찾을 수 없음")
            changed = resv.copy(deep=True)
            changed.loc[candidates[0], "paid"] += 1
            mutated, mutated_catalog, _ = verify(root, changed)
            failed = {row[0] for row in mutated if not row[-1]}
            expected = {"BUILD-018", "BUILD-019", "BUILD-024", "BUILD-025"}
            selftest_ok = failed == expected and mutated_catalog
            print(f"\nSELFTEST: 완료 paid 한 칸 +1원 → 실패 {', '.join(sorted(failed))}")
            print(f"SELFTEST: 예상 실패 집합과 일치 = {selftest_ok} (파일 변경 없음)")
            ok = ok and selftest_ok
        print(f"판정: {'통과' if ok else '실패'} — 같음 {sum(row[-1] for row in rows)}/32 · 대조 불가 {sum(map(len, missing.values()))}개")
        return 0 if ok else 1
    except Exception as exc:
        # 데이터 값이나 경로가 예외 메시지에 섞여 나오지 않게 종류만 출력한다.
        print(f"판정: 실패 — 실행/계약 오류 ({type(exc).__name__})")
        return 1


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
