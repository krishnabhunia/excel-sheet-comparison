"""Self-contained checks for the Excel comparison command."""

from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from excel_compare import compare_workbooks, main, write_report


def make_sheet(path: Path, headers: list[str], records: list[list[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Employees"
    ws.append(headers)
    for row in records:
        ws.append(row)
    wb.save(path)
    wb.close()


class ExcelComparisonTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory(prefix="excel_compare_tests_")
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.a = self.root / "a.xlsx"
        self.b = self.root / "b.xlsx"

    def test_identical_files_have_no_findings(self) -> None:
        make_sheet(self.a, ["ID", "Name", "Amount"], [[1, "Alice", 12], [2, "Bob", 13]])
        shutil.copyfile(self.a, self.b)
        result = compare_workbooks(self.a, self.b, batch_size=1)
        self.assertTrue(result.identical)
        self.assertEqual(0, result.findings.total)

    def test_column_order_and_case_do_not_block_value_comparison(self) -> None:
        make_sheet(self.a, ["ID", "Name", "Amount"], [[1, "Alice", 12], [2, "Bob", 13]])
        make_sheet(self.b, ["id", "Amount", "Name"], [[1, 12, "alice"], [2, 14, "Bob"]])
        result = compare_workbooks(self.a, self.b, batch_size=1)
        self.assertEqual(1, result.case_changes)
        self.assertEqual(2, result.order_changes)
        self.assertEqual(2, result.findings.counts["Value"])
        self.assertFalse(result.data_identical)

    def test_font_alignment_row_height_and_column_width(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"]])
        make_sheet(self.b, ["ID", "Name"], [[1, "Alice"]])
        wb = load_workbook(self.b)
        ws = wb.active
        ws["B2"].font = Font(name="Times New Roman", size=11)
        ws["B2"].alignment = Alignment(horizontal="right")
        ws.row_dimensions[2].height = 28
        ws.column_dimensions["B"].width = 25
        wb.save(self.b)
        wb.close()
        result = compare_workbooks(self.a, self.b)
        issues = {item.issue for item in result.findings.rows}
        self.assertIn("Font name", issues)
        self.assertIn("Horizontal alignment", issues)
        self.assertIn("Row height", issues)
        self.assertIn("Column width", issues)
        self.assertTrue(result.data_identical)
        self.assertFalse(result.formatting_identical)

    def test_additional_row_and_duplicate_header(self) -> None:
        make_sheet(self.a, ["Name", "NAME"], [["a", "b"]])
        make_sheet(self.b, ["Name", "NAME"], [["a", "b"], ["c", "d"]])
        result = compare_workbooks(self.a, self.b)
        self.assertEqual(1, result.findings.counts["Row"])
        self.assertGreater(result.findings.issue_counts["Duplicate header"], 0)

    def test_detail_cap_does_not_change_summary_counts(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.b, ["ID", "Name"], [[1, "alice"], [2, "Robert"], [3, "Extra"]])
        result = compare_workbooks(self.a, self.b, max_details=1)
        self.assertEqual(3, result.findings.total)
        self.assertTrue(result.findings.truncated)
        out = self.root / "report.xlsx"
        write_report(result, out)
        wb = load_workbook(out, read_only=True)
        self.assertEqual(3, wb["Summary"]["E24"].value)
        self.assertEqual(1, wb["Data Differences"].max_row - 4)
        self.assertIn("Showing 1/3 data findings", wb["Summary"]["A32"].value)
        wb.close()

    def test_cli_does_not_overwrite_existing_report(self) -> None:
        make_sheet(self.a, ["ID"], [[1]])
        shutil.copyfile(self.a, self.b)
        out = self.root / "report.xlsx"
        out.write_bytes(b"original report")
        rc = main([str(self.a), str(self.b), "--output", str(out)])
        self.assertEqual(2, rc)
        self.assertEqual(b"original report", out.read_bytes())

    def test_single_primary_key_aligns_reordered_rows_and_differently_named_keys(self) -> None:
        make_sheet(self.a, ["a1", "Name", "Amount"], [[1, "Alice", 10], [2, "Bob", 20]])
        make_sheet(self.b, ["b1", "Amount", "Name"], [[2, 20, "Bob"], [1, 10, "Alice"]])
        result = compare_workbooks(self.a, self.b, key_a="A1", key_b="B1", batch_size=1)
        self.assertEqual(("a1",), result.key_columns_a)
        self.assertEqual(("b1",), result.key_columns_b)
        self.assertEqual(1, result.findings.issue_counts["Mapped key header"])
        self.assertEqual(0, result.findings.counts["Value"])
        self.assertEqual(0, result.findings.counts["Row"])
        self.assertEqual(0, result.findings.counts["Formatting"])
        self.assertFalse(result.data_identical)  # Differently named headers are still reported.

    def test_reordered_rows_with_identical_headers_are_identical_when_keyed(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.b, ["ID", "Name"], [[2, "Bob"], [1, "Alice"]])
        keyed = compare_workbooks(self.a, self.b, key_a="ID", key_b="id")
        positional = compare_workbooks(self.a, self.b)
        self.assertTrue(keyed.identical)
        self.assertEqual(0, keyed.findings.total)
        self.assertFalse(positional.identical)

    def test_composite_key_aligns_by_all_parts_and_locates_actual_excel_cells(self) -> None:
        make_sheet(
            self.a, ["a2", "a3", "Value"],
            [["x", 1, "first"], ["x", 2, "second"], ["y", 1, "third"]],
        )
        make_sheet(
            self.b, ["b2", "b3", "VALUE"],
            [["x", 2, "second"], ["y", 1, "changed"], ["x", 1, "first"]],
        )
        result = compare_workbooks(
            self.a, self.b, key_a=["A2", "a3"], key_b=["B2", "b3"], batch_size=2,
        )
        self.assertEqual(1, result.findings.counts["Value"])
        self.assertEqual(0, result.findings.counts["Row"])
        value_difference = next(item for item in result.findings.rows if item.category == "Value")
        self.assertEqual("C4", value_difference.a_cell)
        self.assertEqual("C3", value_difference.b_cell)

    def test_keyed_rows_compare_row_height_by_record_instead_of_position(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.b, ["id", "Name"], [[2, "Bob"], [1, "Alice"]])
        for path, height_first, height_second in (
            (self.a, 18, 30), (self.b, 30, 18),
        ):
            wb = load_workbook(path)
            wb.active.row_dimensions[2].height = height_first
            wb.active.row_dimensions[3].height = height_second
            wb.save(path)
            wb.close()
        result = compare_workbooks(self.a, self.b, key_a="id", key_b="ID")
        self.assertEqual(0, result.findings.counts["Formatting"])
        self.assertEqual(0, result.findings.counts["Row"])
        self.assertEqual(0, result.findings.counts["Value"])

    def test_keyed_comparison_reports_missing_and_additional_keys(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.b, ["ID", "Name"], [[2, "Bob"], [3, "Chris"]])
        result = compare_workbooks(self.a, self.b, key_a="ID", key_b="id")
        self.assertEqual(1, result.findings.issue_counts["Missing data row"])
        self.assertEqual(1, result.findings.issue_counts["Additional data row"])
        self.assertEqual(0, result.findings.counts["Value"])

    def test_duplicate_and_blank_primary_keys_fail_validation(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.b, ["ID", "Name"], [[1, "Alice"], [1, "Bob"]])
        with self.assertRaisesRegex(ValueError, "File B has duplicate primary key"):
            compare_workbooks(self.a, self.b, key_a="ID", key_b="ID")
        make_sheet(self.b, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [1, "Bob"]])
        with self.assertRaisesRegex(ValueError, "File A has duplicate primary key"):
            compare_workbooks(self.a, self.b, key_a="ID", key_b="ID")
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [None, "Bob"]])
        with self.assertRaisesRegex(ValueError, "File A primary key contains a blank"):
            compare_workbooks(self.a, self.b, key_a="ID", key_b="ID")

    def test_key_definition_requires_both_sides_and_unambiguous_columns(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"]])
        make_sheet(self.b, ["id", "Name"], [[1, "Alice"]])
        with self.assertRaisesRegex(ValueError, "Supply --key-a and --key-b"):
            compare_workbooks(self.a, self.b, key_a="ID")
        with self.assertRaisesRegex(ValueError, "same number of columns"):
            compare_workbooks(self.a, self.b, key_a="ID,Name", key_b="id")
        make_sheet(self.b, ["ID", "id"], [[1, 1]])
        with self.assertRaisesRegex(ValueError, "must match exactly one header"):
            compare_workbooks(self.a, self.b, key_a="ID", key_b="id")

    def test_cli_composite_key_adds_matching_method_to_report(self) -> None:
        make_sheet(self.a, ["a2", "a3", "Name"], [[1, "x", "Alice"]])
        make_sheet(self.b, ["b2", "b3", "Name"], [[1, "x", "Alice"]])
        output = self.root / "keyed_report.xlsx"
        rc = main([
            str(self.a), str(self.b), "--output", str(output),
            "--key-a", "a2,a3", "--key-b", "b2,b3", "--batch-size", "1",
            "--temp-dir", str(self.root),
        ])
        self.assertEqual(0, rc)
        wb = load_workbook(output, read_only=True)
        self.assertIn("unique primary key", wb["Summary"]["A28"].value)
        self.assertIn("Key A: a2, a3; Key B: b2, b3", wb["Summary"]["A29"].value)
        self.assertEqual("Unmatched / mapped headers", wb["Summary"]["A16"].value)
        self.assertEqual(2, wb["Summary"]["E16"].value)
        self.assertEqual(["Summary", "Data Differences"], wb.sheetnames)
        wb.close()

    def test_data_differences_are_separate_from_other_findings(self) -> None:
        make_sheet(self.a, ["ID", "Name", "Amount"], [[1, "Alice", 100]])
        make_sheet(self.b, ["id", "Name", "Amount"], [[1, "ALICE", "100"], [2, "Bob", 200]])
        workbook = load_workbook(self.b)
        workbook.active["B2"].font = Font(name="Times New Roman", size=11)
        workbook.save(self.b)
        workbook.close()
        result = compare_workbooks(self.a, self.b)
        report = self.root / "split.xlsx"
        write_report(result, report)
        wb = load_workbook(report, read_only=True)
        self.assertEqual(["Summary", "Data Differences"], wb.sheetnames)
        data_categories = [
            row[1] for row in wb["Data Differences"].iter_rows(min_row=5, values_only=True)
            if row[1] is not None
        ]
        other_categories = [
            row[0] for row in wb["Summary"].iter_rows(min_row=37, values_only=True)
            if row[0] in {"Header", "Workbook", "Formatting"}
        ]
        self.assertEqual({"Value", "Data type", "Row"}, set(data_categories))
        self.assertEqual({"Header", "Formatting"}, set(other_categories))
        self.assertEqual(result.findings.total, len(data_categories) + len(other_categories))
        wb.close()

    def test_data_and_other_details_have_independent_caps(self) -> None:
        make_sheet(self.a, ["ID", "Name"], [[1, "Alice"], [2, "Bob"]])
        make_sheet(self.b, ["id", "Name"], [[1, "ALICE"], [2, "BOB"]])
        result = compare_workbooks(self.a, self.b, max_details=1)
        self.assertTrue(result.findings.truncated)
        self.assertEqual(1, len(result.findings.other_rows))
        self.assertEqual(1, len(result.findings.data_rows))
        self.assertEqual(3, result.findings.total)
        report = self.root / "capped_split.xlsx"
        write_report(result, report)
        wb = load_workbook(report, read_only=True)
        self.assertEqual("Header", wb["Summary"]["A37"].value)
        self.assertEqual("Value", wb["Data Differences"]["B5"].value)
        self.assertEqual(3, wb["Summary"]["E24"].value)
        wb.close()

    def test_report_handles_data_only_and_other_only_findings(self) -> None:
        make_sheet(self.a, ["Name"], [["Alice"]])
        make_sheet(self.b, ["Name"], [["Bob"]])
        result = compare_workbooks(self.a, self.b)
        report = self.root / "data_only.xlsx"
        write_report(result, report)
        wb = load_workbook(report, read_only=True)
        self.assertEqual("No other differences found", wb["Summary"]["A37"].value)
        self.assertEqual("Value", wb["Data Differences"]["B5"].value)
        wb.close()

        make_sheet(self.b, ["name"], [["Alice"]])
        result = compare_workbooks(self.a, self.b)
        report = self.root / "other_only.xlsx"
        write_report(result, report)
        wb = load_workbook(report, read_only=True)
        self.assertEqual("Header", wb["Summary"]["A37"].value)
        self.assertEqual("No data differences found", wb["Data Differences"]["A5"].value)
        wb.close()


if __name__ == "__main__":
    unittest.main()
