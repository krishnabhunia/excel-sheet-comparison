# Excel Comparison Tool

Compares two Excel worksheets using **Python 3.12.4**, **Polars**, and **openpyxl**.
Creates one Excel report with two tabs:

| Worksheet | Contents |
| --- | --- |
| **Summary** | Overall comparison checks, followed by header, workbook and formatting differences |
| **Data Differences** | Changed values, different cell data types, and missing or additional data records |

The Summary keeps the original comparison totals, with the other differences
listed beneath the overview. No data-value or row-level issues are mixed into
the other-differences table.

## Set up in PyCharm or Windows PowerShell

Open a terminal in the folder containing these files:

~~~powershell
poetry env use 3.12
poetry install
~~~

If Poetry is not installed, install the two dependencies in your existing Python 3.12 environment:

~~~powershell
python -m pip install "polars>=2,<3" "openpyxl>=3.1,<4"
~~~

## Compare two Excel files

~~~powershell
poetry run python excel_compare.py "C:\Data\Sheet_A.xlsx" "C:\Data\Sheet_B.xlsx" --output "C:\Data\Comparison_Report.xlsx"
~~~

The script reads the first worksheet in each file by default. To select sheets:

~~~powershell
poetry run python excel_compare.py "C:\Data\A.xlsx" "C:\Data\B.xlsx" --sheet-a "Data" --sheet-b "Data" --output "C:\Data\Report.xlsx"
~~~

## Optional primary-key matching

**No key supplied:** The program compares row 2 of File A with row 2 of File B,
row 3 with row 3, and so on (the original behaviour).

**Single-column key:** If File A uses `a1` and File B uses `b1`:

~~~powershell
poetry run python excel_compare.py "C:\Data\A.xlsx" "C:\Data\B.xlsx" --key-a "a1" --key-b "b1" --output "C:\Data\Report.xlsx"
~~~

**Composite key:** If File A uses `a2 + a3` and File B uses `b2 + b3`:

~~~powershell
poetry run python excel_compare.py "C:\Data\A.xlsx" "C:\Data\B.xlsx" --key-a "a2,a3" --key-b "b2,b3" --output "C:\Data\Report.xlsx"
~~~

- Supply **both** `--key-a` and `--key-b`, with the same number of columns in the same corresponding order.
- Key **column names** ignore case; key **values** are exact and case-sensitive.
- The complete single/composite key must be **nonblank and unique in each file**. Duplicate, blank, missing, ambiguous or formula-based keys stop the comparison with an error.
- When keys are provided, row order does not matter. Rows absent in one file are reported with their key and original Excel row number.
- The mapped key columns are compared as equivalent even when named differently. Their header-name difference is still reported because the workbooks' headers are not identical.
- Large keyed comparisons use a temporary, disk-backed SQLite index for File B rather than loading its entire worksheet into Random Access Memory (RAM). To choose a drive with space, add `--temp-dir "D:\ExcelTemp"`; the folder must already exist. The temporary index is removed when the comparison finishes.

The package also contains `Keyed_Sample_A.xlsx`, `Keyed_Sample_B.xlsx`,
and `Keyed_Comparison_Report.xlsx` to demonstrate composite-key matching with
reordered rows, renamed key columns, changed values, and missing/extra records.

Other options:

| Option | Default | Purpose |
| --- | --- | --- |
| --header-row | 1 | Row containing the headers |
| --batch-size | 256 | Rows compared per Polars batch |
| --max-details | 100000 | Maximum detail rows **per section**; summary counts all findings even when limited |
| --key-a, --key-b | Not supplied | Key columns in the respective workbooks, comma-separated for composite keys |
| --temp-dir | System temporary folder | Existing folder for the temporary SQLite index in keyed mode |
| --overwrite | Off | Explicitly replace an existing report |

The inputs are opened read-only. The output path must differ from both input paths.

## What is checked?

- Worksheet count and selected worksheet name
- Header count, missing/extra/duplicate headers, case changes and column order
- Data values by Excel row number or, when supplied, by validated primary key
- Data types, formulas as text, leading/trailing spaces and value case
- Font name, size, bold, italic, underline, strike-through and colour
- Background fill, number format, border, horizontal/vertical alignment, text wrapping, rotation and cell protection
- Row heights, column widths, hidden rows/columns, merged ranges and freeze panes
- Additional/missing rows

The **Data Differences** worksheet retains separate original values and Excel
cell coordinates. Other issues in **Summary** show the relevant cell reference,
original value, changed value and issue details.

The report distinguishes **Data identical**, **Formatting identical**, and **Overall identical in the checked scope**.

## Matching and limitations

- Without keys, data rows are compared by **Excel row position**. With keys, matching uses their values and ignores record order.
- Columns are paired by names ignoring case. Case differences are reported; columns missing or duplicated case-insensitively are not guessed.
- Values themselves remain case-sensitive. Whitespace differences are reported.
- Formulas are compared as saved formula text; this program does not recalculate formulas or verify cached results.
- Charts, images, macros, comments, conditional formatting rules/effects, external connections, print settings and some workbook metadata are outside the checked scope. A positive result does **not** prove byte-for-byte identity.
- Only .xlsx and .xlsm inputs are supported; .xls is not.
- For very large/wide sheets, reduce --batch-size. Comparison streams input rows instead of loading all worksheet values into Polars simultaneously.
- Each report section contains at most the configured number of detailed findings.
  This prevents numerous formatting issues from hiding data differences.
  **Summary totals still include every detected finding.**

## Run the tests

~~~powershell
poetry run python -m unittest discover -p "test_excel_compare.py" -v
~~~

## Use from another Python script

~~~python
from pathlib import Path
from excel_compare import compare_workbooks, write_report

result = compare_workbooks(Path("A.xlsx"), Path("B.xlsx"))
write_report(result, Path("Comparison_Report.xlsx"))
print(result.identical, result.findings.total)
~~~
