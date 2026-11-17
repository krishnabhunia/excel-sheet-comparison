"""Generate two small, intentionally different workbooks for key-matching demos."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font


def create_examples(folder: Path) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    examples = [
        (
            "Keyed_Sample_A.xlsx",
            ["A2", "A3", "Customer", "Amount"],
            [
                ["C001", "IN", "Asha", 500],
                ["C001", "UK", "Ravi", 750],
                ["C002", "IN", "Neha", 900],
                ["C003", "IN", "Isha", 1100],
            ],
        ),
        (
            "Keyed_Sample_B.xlsx",
            ["b2", "b3", "Customer", "Amount"],
            [
                ["C003", "IN", "Isha", 1100],
                ["C001", "UK", "Ravi", 800],
                ["C001", "IN", "Asha", 500],
                ["C004", "IN", "Omar", 1200],
            ],
        ),
    ]
    for filename, headers, rows in examples:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Customers"
        worksheet.append(headers)
        for row in rows:
            worksheet.append(row)
        if filename.endswith("B.xlsx"):
            worksheet["C3"].font = Font(name="Times New Roman", size=11)
        workbook.save(folder / filename)
        workbook.close()


if __name__ == "__main__":
    create_examples(Path(__file__).resolve().parent)
