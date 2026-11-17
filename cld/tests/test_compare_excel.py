"""Tests for compare_excel.py.

Run with:  pytest cld/tests        or        python cld/tests/test_compare_excel.py
"""
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import compare_excel as ce  # noqa: E402

# positions inside one row of Result.data
ROW, KEY, COLUMN, VAL_A, VAL_B, KIND, TYPE_FLAG, VALUE_FLAG, DELTA, FLAGS = 1, 2, 3, 6, 7, 8, 9, 10, 11, slice(12, 16)


def book(path: Path, header: list, rows: list, title: str = "Sheet") -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(header)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


def tmp() -> Path:
    return Path(tempfile.mkdtemp(prefix="xlcmp_"))


def identical(res: ce.Result) -> bool:
    return not any(res.count[key] for _, key, _, _ in ce.SUMMARY_CHECKS)


def samples(script: str, a: str, b: str) -> tuple[Path, Path]:
    out = tmp()
    subprocess.run([sys.executable, str(ROOT / "samples" / script), str(out)], check=True, capture_output=True)
    return out / a, out / b


def by_key(res: ce.Result) -> dict:
    return {row[KEY]: row for row in res.data}


# ----------------------------------------------------------------------------- tests
def test_identical_files_by_position_and_by_key():
    a, _ = samples("make_employee_samples.py", "employees_A.xlsx", "employees_B.xlsx")
    assert identical(ce.compare_workbooks(a, a))
    assert identical(ce.compare_workbooks(a, a, key_a=["EMP ID"]))


def test_employee_samples_every_planted_difference_is_found():
    a, b = samples("make_employee_samples.py", "employees_A.xlsx", "employees_B.xlsx")
    res = ce.compare_workbooks(a, b)
    expected = {"workbook": 1, "header_count": 1, "header_names": 1, "header_case": 1, "header_order": 0,
                "header_font": 1, "header_align": 1, "row_count": 1, "rows_only": 1, "values": 5, "types": 1,
                "formulas": 0, "fmt_font": 2, "fmt_align": 1, "fmt_other": 9, "layout": 2, "extras": 0}
    assert {k: res.count[k] for k in expected} == expected
    kinds = [row[KIND] for row in res.data]
    assert kinds == ["Number changed", "Case differs only", "Spacing differs only",
                     "Data type differs (number vs text)", "Date changed", "Blank in File B", "Row only in File B"]


def test_number_samples_flags_and_thresholds():
    a, b = samples("make_number_samples.py", "numbers_A.xlsx", "numbers_B.xlsx")
    res = ce.compare_workbooks(a, b, key_a=["txn id"])
    assert res.count["values"] == 19 and res.count["types"] == 2
    rows = by_key(res)
    expect = {  # key: (type flag, value flag, number diff, thresholds > 0, 0.01, 0.1, 1)
        "T001": ("Number", "True", 0.005, ["Yes", "No", "No", "No"]),
        "T002": ("Number", "True", 0.01, ["Yes", "No", "No", "No"]),        # exactly 0.01 is not > 0.01
        "T003": ("Number", "True", 0.05, ["Yes", "Yes", "No", "No"]),
        "T004": ("Number", "True", 0.1, ["Yes", "Yes", "No", "No"]),
        "T006": ("Number", "True", 1, ["Yes", "Yes", "Yes", "No"]),
        "T008": ("Number", "True", -2.5, ["Yes", "Yes", "Yes", "Yes"]),
        "T010": ("Text", "True", None, [None] * 4),                         # case only
        "T013": ("Number", "False", 0, ["No", "No", "No", "No"]),           # 750 vs "750"
        "T014": ("Number", "True", 5, ["Yes", "Yes", "Yes", "Yes"]),        # 900 vs "905"
        "T017": ("Number", "False", 0, ["No", "No", "No", "No"]),           # "1.50" vs "1.5"
        "T018": ("Text", "True", None, [None] * 4),                         # "1,200" is not read as a number
        "T019": ("Number", "True", 0.5, ["Yes", "Yes", "Yes", "No"]),       # " 45.5 "
        "T020": ("Text", "False", None, [None] * 4),                        # date vs the same date as text
        "T021": ("Number", "True", None, [None] * 4),                       # number vs blank
    }
    for key, (type_flag, value_flag, delta, flags) in expect.items():
        row = rows[key]
        assert (row[TYPE_FLAG], row[VALUE_FLAG], row[DELTA], row[FLAGS]) == (type_flag, value_flag, delta, flags), key


def test_single_key_handles_reordered_inserted_and_deleted_rows():
    d = tmp()
    rows = [[101, "Asha", 50], [102, "Vikram", 60], [103, "Meera", 70], [104, "Rohan", 80], [105, "Sana", 90]]
    a = book(d / "a.xlsx", ["Emp ID", "Name", "Score"], rows)
    b = book(d / "b.xlsx", ["EMP id", "Name", "Score"],
             [[105, "Sana", 90], [101, "Asha", 50], [106, "Neha", 10], [104, "Rohan", 85], [102, "Vikram", 60]])
    res = ce.compare_workbooks(a, b, key_a=["emp id"])
    assert (res.count["values"], res.count["rows_only"], res.count["row_order"]) == (1, 2, 2)
    assert {row[KEY]: row[KIND] for row in res.data} == {
        "104": "Number changed", "103": "Row only in File A", "106": "Row only in File B"}
    # without a key the same files show many more cell differences
    assert ce.compare_workbooks(a, b).count["values"] > 10


def test_composite_key_with_different_names_in_each_file():
    d = tmp()
    a = book(d / "a.xlsx", ["a2", "a3", "Amount"], [["IN", 1, 10], ["IN", 2, 20], ["US", 1, 30]])
    b = book(d / "b.xlsx", ["B2", "B3", "Amount"], [["US", 1, 30], ["IN", 1, 10], ["IN", 2, 25]])
    res = ce.compare_workbooks(a, b, key_a=["A2", "A3"], key_b=["b2", "b3"])
    assert res.keyed and res.count["header_names"] == 2 and res.count["values"] == 1
    assert res.data[0][KEY] == "IN | 2" and res.data[0][DELTA] == 5


def test_blank_and_repeated_keys_are_warnings_not_differences():
    d = tmp()
    rows = [[1, "x"], [1, "y"], [None, "z"], [2, "w"]]
    a, b = book(d / "a.xlsx", ["id", "val"], rows), book(d / "b.xlsx", ["id", "val"], rows)
    res = ce.compare_workbooks(a, b, key_a=["ID"])
    assert identical(res) and len(res.notes) == 2


def test_key_errors():
    d = tmp()
    a = book(d / "a.xlsx", ["id", "val"], [[1, "x"]])
    for kwargs in ({"key_a": ["nope"]}, {"key_a": ["id", "val"], "key_b": ["id"]}):
        try:
            ce.compare_workbooks(a, a, **kwargs)
        except ValueError:
            continue
        raise AssertionError(f"expected ValueError for {kwargs}")


def test_moved_columns_are_matched_by_name():
    d = tmp()
    a = book(d / "a.xlsx", ["ID", "Name", "City"], [[1, "A", "Pune"], [2, "B", "Delhi"]])
    b = book(d / "b.xlsx", ["ID", "city", "Name"], [[1, "Pune", "A"], [2, "Goa", "B"]])
    res = ce.compare_workbooks(a, b)
    assert (res.count["header_order"], res.count["header_case"], res.count["values"]) == (2, 1, 1)
    assert res.data[0][4:6] == ["C3", "B3"]


def test_tolerance_and_header_row():
    d = tmp()
    a = book(d / "a.xlsx", ["v"], [[1.0]])
    b = book(d / "b.xlsx", ["v"], [[1.004]])
    assert not identical(ce.compare_workbooks(a, b))
    assert identical(ce.compare_workbooks(a, b, tol=0.01))
    a = book(d / "c.xlsx", ["Title"], [[], ["ID", "Amt"], [1, 5]])
    b = book(d / "d.xlsx", ["Title"], [[], ["ID", "Amt"], [1, 6]])
    res = ce.compare_workbooks(a, b, header_row=3)
    assert res.count["values"] == 1 and res.data[0][COLUMN] == "Amt"


def test_multi_sheet_workbooks_and_formulas():
    d = tmp()

    def make(path, formula, extra_sheet):
        wb = Workbook()
        ws = wb.active
        ws.title = "Orders"
        ws.append(["Qty", "Price", "Total"])
        ws.append([2, 5, formula])
        wb.create_sheet(extra_sheet).append(["x"])
        wb.save(path)
        return path

    res = ce.compare_workbooks(make(d / "a.xlsx", "=A2*B2", "OnlyA"), make(d / "b.xlsx", "=A2*B2*1", "OnlyB"))
    assert res.count["formulas"] == 1 and res.count["workbook"] == 2
    assert res.data[0][KIND] == "Formula differs"


def test_report_file_and_exit_codes():
    a, b = samples("make_number_samples.py", "numbers_A.xlsx", "numbers_B.xlsx")
    out = a.parent / "report.xlsx"
    assert ce.main([str(a), str(b), str(out), "--key", "txn id"]) == 1 and out.is_file()
    assert ce.main([str(a), str(a), str(out)]) == 0
    assert ce.main([str(a), str(a.parent / "missing.xlsx"), str(out)]) == 2


if __name__ == "__main__":
    failed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"PASS  {name}")
            except Exception as exc:  # noqa: BLE001
                failed += 1
                print(f"FAIL  {name}: {exc!r}")
    sys.exit(1 if failed else 0)
