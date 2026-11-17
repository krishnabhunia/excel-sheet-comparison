"""Sample pair for the number-threshold columns."""
import datetime as dt, sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

out = Path(sys.argv[1])
H = ["Txn ID", "Account", "Currency", "Amount", "Rate", "Status", "Value Date"]
D = dt.date(2026, 11, 2)
A = [["T001", "ACC-01", "INR", 100.00, 1.25, "Settled", D],
     ["T002", "ACC-02", "INR", 1.12, 1.25, "Settled", D],
     ["T003", "ACC-03", "USD", 250.10, 1.25, "Settled", D],
     ["T004", "ACC-04", "USD", 400.00, 5.5, "Settled", D],
     ["T005", "ACC-05", "EUR", 80, 1.25, "Settled", D],
     ["T006", "ACC-06", "EUR", 500, 1.25, "Settled", D],
     ["T007", "ACC-07", "INR", 62000, 1.25, "Settled", D],
     ["T008", "ACC-08", "INR", 300, 1.25, "Settled", D],
     ["T009", "ACC-09", "INR", 10, 0.3, "Settled", D],
     ["T010", "ACC-10", "GBP", 20, 1.25, "Settled", D],
     ["T011", "ACC-11", "GBP", 30, 1.25, "Settled", D],
     ["T012", "ACC-12", "GBP", 40, 1.25, "Pending", D],
     ["T013", "ACC-13", "INR", 750, 1.25, "Settled", D],
     ["T014", "ACC-14", "INR", 900, 1.25, "Settled", D],
     ["T015", "ACC-15", "INR", 60, 1.25, "Settled", D],
     ["T016", "ACC-16", "INR", 70, 1.25, "Settled", D],
     ["T017", "ACC-17", "INR", "1.50", 1.25, "Settled", D],
     ["T018", "ACC-18", "INR", 1200, 1.25, "Settled", D],
     ["T019", "ACC-19", "INR", 45, 1.25, "Settled", D],
     ["T020", "ACC-20", "INR", 55, 1.25, "Settled", D],
     ["T021", "ACC-21", "INR", 65, 1.25, "Settled", D]]
B = [list(r) for r in A]
B[0][3] = 100.005          # diff 0.005
B[1][3] = 1.13             # diff exactly 0.01
B[2][3] = 250.15           # diff 0.05
B[3][4] = 5.6              # diff exactly 0.1
B[4][3] = 80.5             # diff 0.5
B[5][3] = 501              # diff exactly 1
B[6][3] = 64000            # diff 2000
B[7][3] = 297.5            # diff -2.5
B[8][4] = 0.3000001        # tiny difference, 0.0000001
B[9][5] = "SETTLED"        # text: case
B[10][1] = "ACC-11 "       # text: trailing space
B[11][5] = "Failed"        # text: changed
B[12][3] = "750"           # number stored as text, same value
B[13][3] = "905"           # number vs text, different value
B[14][6] = dt.date(2026, 11, 5)   # date
B[15][2] = None            # blank text
B[16][3] = "1.5"           # text vs text, same number
B[17][3] = "1,200"         # text with a comma: not readable as a number
B[18][3] = " 45.5 "        # text with spaces that reads as a number, different value
B[19][6] = "2026-11-02"    # date stored as text, same date
B[20][3] = None            # blank number

for name, rows in (("numbers_A.xlsx", A), ("numbers_B.xlsx", B)):
    wb = Workbook(); ws = wb.active; ws.title = "Settlements"
    ws.append(H)
    for c in ws[1]:
        c.font = Font(name="Calibri", size=11, bold=True); c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center")
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        row[6].number_format = "dd-mmm-yyyy"
    for col, w in zip("ABCDEFG", (10, 12, 11, 12, 9, 12, 14)):
        ws.column_dimensions[col].width = w
    wb.save(out / name)
print("written numbers_A.xlsx and numbers_B.xlsx")
