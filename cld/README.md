# Excel sheet comparison (`cld`)

`compare_excel.py` checks whether two Excel workbooks are identical and writes an Excel report of every
difference it finds: headers, cell values, data types, formulas, fonts, alignment, number formats, fills,
borders and sheet layout.

openpyxl reads the cells and their formatting, Polars does the comparison, xlsxwriter writes the report.

## Install

```bash
pip install -r cld/requirements.txt
```

Tested with Python 3.13, Polars 2.1.0, openpyxl 3.1.5 and xlsxwriter 3.2.9. Input files must be `.xlsx` or `.xlsm`.

## Usage

```bash
python cld/compare_excel.py FILE_A.xlsx FILE_B.xlsx [REPORT.xlsx] [options]
```

| Argument | Meaning | Default |
|---|---|---|
| `FILE_A`, `FILE_B` | The two workbooks to compare | required |
| `REPORT` | Report file to create | `comparison_report.xlsx` |
| `--key "a1"` or `--key "a2,a3"` | Primary key column(s), same names in both files | rows matched by position |
| `--key-a`, `--key-b` | Primary key column(s) when the names differ between the files | same as `--key` |
| `--header-row N` | Row that holds the column headers | `1` |
| `--tolerance X` | Allowed absolute difference between two numbers | `0` |
| `--sheet-a`, `--sheet-b` | Compare only one sheet from each file | all sheets, paired by name |

Examples:

```bash
# no primary key: row N of File A is compared with row N of File B
python cld/compare_excel.py a.xlsx b.xlsx report.xlsx

# single key, named a1 in File A and b1 in File B
python cld/compare_excel.py a.xlsx b.xlsx report.xlsx --key-a "a1" --key-b "b1"

# composite key: a2 pairs with b2, a3 pairs with b3
python cld/compare_excel.py a.xlsx b.xlsx report.xlsx --key-a "a2,a3" --key-b "b2,b3"
```

Exit code: `0` identical, `1` differences found, `2` error (for example a missing file or key column).

## Matching rules

- Columns are matched by header name, ignoring case and leading/trailing spaces.
- Without a key, rows are matched by position.
- With a key, rows are matched by key value, so inserted, deleted and reordered rows are reported as such.
  Key column names are matched ignoring case. Key values are matched as exact text.
- Blank or repeated key values are reported as a warning and matched in order of appearance.
- Sheets are paired by name, ignoring case.

## Report

| Sheet | Contents |
|---|---|
| Summary | Overall result (IDENTICAL / NOT IDENTICAL) and 18 checks, each MATCH or DIFF with a count |
| Headers | One row per header: name, case, position, font and alignment in each file |
| Data Differences | One row per differing cell, plus rows that exist in only one file |
| Format Differences | Font, alignment, number format, fill and border differences |
| Layout & Other | Sheet names, column widths, row heights, merged cells, freeze panes, hidden rows and columns, comments, hyperlinks, data validation, conditional formatting |

Columns of the Data Differences sheet that need explaining:

| Column | Meaning |
|---|---|
| Key | Key value of the row (only when a primary key is used) |
| Difference Type | Number changed, Text changed, Case differs only, Spacing differs only, Data type differs, Date changed, Blank in File A/B, Formula differs, Row only in File A/B |
| Difference Type Flag | `Number` when both values are numbers or text that reads as a number (`"750"`, `" 45.5 "`); otherwise `Text` |
| Difference Type Value Flag | `True` when the values still differ once both are on the same platform; `False` when only the data type or notation differs (number `750` vs text `"750"`) |
| Number Diff (B − A) | Filled for `Number` rows |
| Abs Diff > 0, > 0.01, > 0.1, > 1 | Yes/No for `Number` rows. Strictly greater than, so a difference of exactly 0.01 is "No" under "> 0.01". Change `DIFF_THRESHOLDS` in the script for other values |

Text such as `1,200`, `₹500` or `10%` is not read as a number. Dates and TRUE/FALSE cells are compared as text.

## Samples

`samples/` holds two sample pairs, the scripts that create them and the report produced for each.

| Pair | Report | Command |
|---|---|---|
| `employees_A.xlsx`, `employees_B.xlsx` (19 planted differences) | `report_employees.xlsx` | `python compare_excel.py samples/employees_A.xlsx samples/employees_B.xlsx samples/report_employees.xlsx` |
| `numbers_A.xlsx`, `numbers_B.xlsx` (21 planted differences) | `report_numbers.xlsx` | `python compare_excel.py samples/numbers_A.xlsx samples/numbers_B.xlsx samples/report_numbers.xlsx --key "txn id"` |

Run the commands from inside `cld/`.

## Tests

```bash
pytest cld/tests
```

## Limits

- Every cell's formatting is read into memory, so this suits ordinary workbooks, not multi-gigabyte data.
- Formulas that were never calculated in Excel have no saved result; their values compare as blank and the Summary sheet says so.
- Merged cells, data validation and conditional formatting are compared by cell address, not by key.
