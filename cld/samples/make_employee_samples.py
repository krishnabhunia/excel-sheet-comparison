"""Build two sample workbooks with a known set of planted differences."""
import datetime as dt
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)

HEADERS = ["Emp ID", "Employee Name", "Department", "City", "Joining Date", "Salary"]
ROWS = [
    [1001, "Asha Rao", "Engineering", "Pune", dt.date(2019, 6, 3), 58000],
    [1002, "Vikram Shah", "Engineering", "Mumbai", dt.date(2020, 1, 20), 62000],
    [1003, "Meera Iyer", "Analytics", "Pune", dt.date(2018, 11, 5), 71000],
    [1004, "Rohan Das", "Finance", "Kolkata", dt.date(2022, 4, 11), 49000],
    [1005, "Sana Khan", "Analytics", "Delhi", dt.date(2021, 9, 27), 66000],
    [1006, "Arjun Nair", "Sales", "Kochi", dt.date(2021, 3, 15), 45000],
    [1007, "Divya Menon", "Sales", "Chennai", dt.date(2023, 2, 6), 43000],
    [1008, "Kabir Sethi", "Finance", "Delhi", dt.date(2017, 8, 14), 83000],
]
WIDTHS = {"A": 10, "B": 22, "C": 16, "D": 14, "E": 15, "F": 12}


def build(variant: str) -> Workbook:
    b = variant == "B"
    wb = Workbook()
    ws = wb.active
    ws.title = "employees" if b else "Employees"          # sheet-name case

    headers = list(HEADERS)
    if b:
        headers[1] = "EMPLOYEE NAME"                       # header case
        headers.append("Remarks")                          # extra header
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
    if b:
        ws["C1"].font = Font(name="Arial", size=11, bold=True)                  # header font
        ws["D1"].alignment = Alignment(horizontal="left", vertical="center")    # header alignment

    rows = [list(r) for r in ROWS]
    if b:
        rows[1][5] = 64000                    # F3  number changed
        rows[2][3] = "PUNE"                   # D4  case only
        rows[3][2] = "Finance "               # C5  trailing space
        rows[4][0] = "1005"                   # A6  number -> text
        rows[5][4] = dt.date(2021, 3, 18)     # E7  date changed
        rows[6][3] = None                     # D8  blank in B
        rows.append([1009, "Neha Joshi", "Engineering", "Pune", dt.date(2024, 7, 1), 52000])  # extra row

    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="Calibri", size=11)
        ws.cell(row=r, column=5).number_format = "dd-mmm-yyyy"
        ws.cell(row=r, column=6).number_format = "#,##0.00" if b else "#,##0"   # number format, whole column
        if b:
            ws.cell(row=r, column=7, value="OK" if r % 2 else None)

    if b:
        ws["B5"].font = Font(name="Times New Roman", size=11)        # cell font
        ws["A2"].font = Font(name="Calibri", size=11, bold=True)     # bold
        ws["D7"].fill = PatternFill("solid", fgColor="FFFF00")       # fill
        ws["C3"].alignment = Alignment(horizontal="center")          # cell alignment

    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    if b:
        ws.column_dimensions["B"].width = 30                         # column width
    else:
        ws.freeze_panes = "A2"                                       # freeze panes only in A
    return wb


build("A").save(out / "employees_A.xlsx")
build("B").save(out / "employees_B.xlsx")
print("written employees_A.xlsx and employees_B.xlsx")
