#!/usr/bin/env python3
"""
compare_excel.py - check whether two Excel workbooks are identical and write a difference report.

    python compare_excel.py FILE_A.xlsx FILE_B.xlsx [REPORT.xlsx]
                            [--key COLS] [--key-a COLS] [--key-b COLS]
                            [--header-row N] [--tolerance X]
                            [--sheet-a NAME] [--sheet-b NAME]

Primary key (optional)
    --key "emp id"                      same key column name in both files
    --key "a2,a3"                       composite key, same names in both files
    --key-a "a1"    --key-b "b1"        key column is named differently in each file
    --key-a "a2,a3" --key-b "b2,b3"     composite key; a2 pairs with b2, a3 pairs with b3
    Key column names are matched ignoring case and leading/trailing spaces.
    With a key, rows are matched by key value, so inserted, deleted and reordered rows are
    reported as such. Without a key, rows are matched by position (row N against row N).

Same-platform comparison (Data Differences sheet)
    Difference Type Flag        Number - both values are numbers, or text that reads as a number ("750", " 45.5 ")
                                Text   - anything else (compared as text)
    Difference Type Value Flag  True   - the values still differ once both are on the same platform
                                False  - only the data type differs (number 750 vs text "750")
    For "Number" rows the sheet also shows B - A and one Yes/No column per value in DIFF_THRESHOLDS
    (0, 0.01, 0.1, 1): "is the absolute difference greater than this?".

How it works
    openpyxl reads every cell (value, formula, font, alignment, fill, number format, borders)
    and the sheet layout. All comparisons are then done in Polars (joins / filters / group-bys).
    xlsxwriter writes the report.

Exit code: 0 = identical, 1 = differences found, 2 = error.
Requires:  pip install polars openpyxl xlsxwriter
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from collections import Counter
from pathlib import Path

import polars as pl
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

MAX_REPORT_ROWS = 200_000          # per report sheet; extra rows are counted but not listed
BLANK = "(blank)"
NULL_KEY = "\x00null"              # stands in for a blank key value so that blanks can be joined
PAIR_SCHEMA = {"row": pl.Int32, "row_b": pl.Int32, "key_text": pl.String}
DIFF_THRESHOLDS = (0, 0.01, 0.1, 1)   # one Yes/No column per threshold in the Data Differences sheet
NO_FLAGS = [None] * len(DIFF_THRESHOLDS)
NUMERIC_TEXT = r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$"   # text that can be read as a number: 750, -3.25, 1e3


def _clean(delta):
    """Number difference without floating-point noise (0.009999999999999787 -> 0.01)."""
    if delta is None:
        return None
    return round(delta, 9) or delta          # keep the raw value when it is smaller than 0.000000001


def _threshold_flags(delta) -> list:
    """Yes/No per threshold for a number difference; blank for anything that is not number vs number."""
    if delta is None:
        return list(NO_FLAGS)
    size = abs(delta)
    # round away floating-point noise so that a difference of exactly 0.01 is not reported as "> 0.01"
    return ["Yes" if (size > 0 if t == 0 else round(size, 9) > t) else "No" for t in DIFF_THRESHOLDS]

# (dataframe column, label in report, category)
PROPS = [
    ("font_name", "Font name", "font"),
    ("font_size", "Font size", "font"),
    ("bold", "Bold", "font"),
    ("italic", "Italic", "font"),
    ("underline", "Underline", "font"),
    ("font_color", "Font colour", "font"),
    ("h_align", "Horizontal alignment", "align"),
    ("v_align", "Vertical alignment", "align"),
    ("wrap", "Wrap text", "align"),
    ("number_format", "Number format", "other"),
    ("fill", "Fill colour", "other"),
    ("border", "Borders", "other"),
]
PROP_COLS = [p[0] for p in PROPS]
FONT_COLS = [p[0] for p in PROPS if p[2] == "font"]
ALIGN_COLS = [p[0] for p in PROPS if p[2] == "align"]


# --------------------------------------------------------------------------- reading
def _yn(flag) -> str:
    return "Yes" if flag else "No"


def _color(c) -> str:
    if c is None:
        return "none"
    if c.type == "rgb":
        return f"#{c.rgb[-6:]}" if isinstance(c.rgb, str) else "none"
    if c.type == "theme":
        return f"theme {c.theme}" + (f" tint {c.tint:+.2f}" if c.tint else "")
    if c.type == "indexed":
        return f"indexed {c.indexed}"
    return str(c.type)


def _fill(f) -> str:
    kind = getattr(f, "fill_type", None)
    if not kind:
        return "none"
    colour = _color(getattr(f, "fgColor", None))
    return colour if kind == "solid" else f"{kind} {colour}"


def _border(b) -> str:
    parts = []
    for side, tag in (("left", "L"), ("right", "R"), ("top", "T"), ("bottom", "B")):
        s = getattr(b, side, None)
        if s is not None and s.style:
            parts.append(f"{tag}:{s.style}")
    return " ".join(parts) or "none"


def _style(cell) -> tuple:
    f, a = cell.font, cell.alignment
    return (
        f.name or "", f"{f.sz:g}" if f.sz else "", _yn(f.b), _yn(f.i), f.u or "none", _color(f.color),
        a.horizontal or "general", a.vertical or "bottom", _yn(a.wrap_text),
        cell.number_format or "General", _fill(cell.fill), _border(cell.border),
    )


def _norm(v):
    """Python cell value -> (text, data type, float or None)."""
    if v is None:
        return None, "blank", None
    if isinstance(v, bool):
        return ("TRUE" if v else "FALSE"), "boolean", None
    if isinstance(v, (int, float)):
        f = float(v)
        text = str(int(f)) if f.is_integer() and abs(f) < 1e15 else format(f, ".15g")
        return text, "number", f
    if isinstance(v, dt.datetime):
        text = v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ")
        return text, "date", None
    if isinstance(v, dt.date):
        return v.isoformat(), "date", None
    if isinstance(v, dt.time):
        return v.isoformat(), "time", None
    if isinstance(v, str):
        return v, "text", None
    return str(v), type(v).__name__, None


def open_book(path: Path) -> dict:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(f"{path.name}: only .xlsx / .xlsm are supported (save .xls or .csv as .xlsx first)")
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")
    return {"path": path, "wb": load_workbook(path), "cached": None, "styles": {}}


def _cached_sheet(book: dict, name: str):
    """Workbook loaded a second time with data_only=True - only needed when formulas exist."""
    if book["cached"] is None:
        book["cached"] = load_workbook(book["path"], data_only=True)
    return book["cached"][name]


def read_sheet(book: dict, name: str) -> dict:
    ws = book["wb"][name]
    styles = book["styles"]
    R, C, V, T, N, F, S = [], [], [], [], [], [], []
    comments, links, formula_idx = {}, {}, []

    for r, row in enumerate(ws.iter_rows(), 1):
        for c, cell in enumerate(row, 1):
            value, formula = cell.value, None
            if cell.data_type == "f":
                formula = value if isinstance(value, str) else (getattr(value, "text", None) or str(value))
                value = None
                formula_idx.append(len(R))
            text, dtype, num = _norm(value)
            sid = cell.style_id
            if sid not in styles:
                styles[sid] = _style(cell)
            R.append(r); C.append(c); V.append(text); T.append(dtype); N.append(num); F.append(formula); S.append(sid)
            if getattr(cell, "comment", None) is not None:
                comments[cell.coordinate] = (cell.comment.text or "").strip()
            link = getattr(cell, "hyperlink", None)
            if link is not None:
                links[cell.coordinate] = link.target or link.location or ""

    uncached = 0
    if formula_idx:                       # fetch the calculated results of formula cells
        wsv = _cached_sheet(book, name)
        for i in formula_idx:
            V[i], T[i], N[i] = _norm(wsv.cell(row=R[i], column=C[i]).value)
            uncached += V[i] is None

    cells = pl.DataFrame(
        {"row": R, "col": C, "value": V, "dtype": T, "num": N, "formula": F, "sid": S},
        schema={"row": pl.Int32, "col": pl.Int32, "value": pl.String, "dtype": pl.String,
                "num": pl.Float64, "formula": pl.String, "sid": pl.Int32},
    )
    style_df = pl.DataFrame(
        [(sid, *props) for sid, props in styles.items()],
        schema={"sid": pl.Int32, **{c: pl.String for c in PROP_COLS}}, orient="row",
    )
    cells = cells.join(style_df, on="sid", how="left").drop("sid")

    used = cells.filter(pl.col("value").is_not_null() | pl.col("formula").is_not_null())
    last_row = used["row"].max() or 0
    last_col = used["col"].max() or 0
    cells = cells.filter((pl.col("row") <= last_row) & (pl.col("col") <= last_col))

    widths, hidden_cols = {}, set()
    for d in ws.column_dimensions.values():
        lo = d.min or 1
        for i in range(lo, min(d.max or lo, last_col) + 1):
            if d.width is not None:
                widths[i] = round(d.width, 2)
            if d.hidden:
                hidden_cols.add(get_column_letter(i))

    return {
        "name": name, "cells": cells, "last_row": last_row, "last_col": last_col,
        "widths": widths, "hidden_cols": hidden_cols,
        "heights": {r: round(d.height, 2) for r, d in ws.row_dimensions.items() if d.height is not None},
        "hidden_rows": {r for r, d in ws.row_dimensions.items() if d.hidden},
        "merged": {str(m) for m in ws.merged_cells.ranges},
        "freeze": ws.freeze_panes, "filter": ws.auto_filter.ref, "state": ws.sheet_state,
        "comments": comments, "links": links,
        "validations": sorted(f"{dv.sqref}: {dv.type} {dv.formula1 or ''}".strip()
                              for dv in ws.data_validations.dataValidation),
        "cond_formats": sorted(f"{cf.sqref}: {rule.type}" for cf in ws.conditional_formatting for rule in cf.rules),
        "uncached": uncached,
    }


# --------------------------------------------------------------------------- comparing
def header_list(sheet: dict, header_row: int, forced: dict | None = None) -> list[dict]:
    """One entry per column: position, header text as written, and a case-insensitive match key.
    `forced` maps a column number to a fixed match key (used to pair key columns whose names differ)."""
    forced = forced or {}
    hdr = {r["col"]: r for r in sheet["cells"].filter(pl.col("row") == header_row).iter_rows(named=True)}
    seen, out = Counter(), []
    for col in range(1, sheet["last_col"] + 1):
        cell = hdr.get(col, {})
        raw = cell.get("value")
        if col in forced:
            key = forced[col]
        else:
            base = raw.strip().casefold() if raw is not None else f"\x00blank{col}"
            seen[base] += 1
            key = base if seen[base] == 1 else f"{base}\x00{seen[base]}"      # duplicate header names
        out.append({"col": col, "raw": raw, "key": key, "cell": cell})
    return out


def resolve_key(headers: list[dict], wanted: list[str]) -> list[int] | None:
    """Key column names -> column numbers (case-insensitive). None if any name is missing."""
    index = {}
    for h in headers:
        if h["raw"] is not None:
            index.setdefault(h["raw"].strip().casefold(), h["col"])
    cols = [index.get(w.strip().casefold()) for w in wanted]
    return None if any(c is None for c in cols) else cols


def key_frame(sheet: dict, cols: list[int], header_row: int) -> pl.DataFrame:
    """One line per non-blank data row: row, k0..kn (key parts), occ (nth repeat of that key), flags."""
    cells = sheet["cells"].filter(pl.col("row") > header_row)
    kf = cells.filter(pl.col("value").is_not_null()).select("row").unique()
    kn = [f"k{i}" for i in range(len(cols))]
    for name, c in zip(kn, cols):
        kf = kf.join(cells.filter(pl.col("col") == c).select("row", pl.col("value").alias(name)), on="row", how="left")
    return (kf.sort("row")
              .with_columns(null_key=pl.any_horizontal([pl.col(k).is_null() for k in kn]),
                            key_text=pl.concat_str([pl.col(k).fill_null(BLANK) for k in kn], separator=" | "))
              .with_columns([pl.col(k).fill_null(NULL_KEY) for k in kn])
              .with_columns(occ=pl.int_range(pl.len()).over(kn), dup=pl.len().over(kn) > 1))


def _font_text(c: dict) -> str:
    if not c:
        return ""
    text = f"{c['font_name']} {c['font_size']}pt"
    text += " Bold" if c["bold"] == "Yes" else ""
    text += " Italic" if c["italic"] == "Yes" else ""
    text += " Underline" if c["underline"] != "none" else ""
    if c["font_color"] not in ("none", "theme 1"):
        text += f", colour {c['font_color']}"
    return text


def _align_text(c: dict) -> str:
    if not c:
        return ""
    return f"{c['h_align']} / {c['v_align']}" + (" / wrap" if c["wrap"] == "Yes" else "")


def _name_check(a: str | None, b: str | None) -> str:
    if a == b:
        return "Exact match"
    if a is None or b is None:
        return "Blank in one file"
    if a.strip().casefold() != b.strip().casefold():
        return "Name differs"                      # only possible for a key column pair
    same_case = a.strip() == b.strip()
    same_space = a.casefold() == b.casefold()
    if same_case:
        return "Spacing differs"
    return "Case differs" if same_space else "Case and spacing differ"


def _show(v, limit: int = 300) -> str:
    if v is None or v == "" or v == set() or v == [] or v == {}:
        return "none"
    if isinstance(v, dict):
        v = ", ".join(f"{k}={val}" for k, val in v.items())
    elif isinstance(v, (set, list, tuple)):
        items = sorted(v, key=lambda x: (len(str(x)), str(x))) if isinstance(v, set) else v
        v = ", ".join(str(x) for x in items)
    v = str(v)
    return v if len(v) <= limit else v[: limit - 1] + "…"


def _quote(v):
    return f'"{v}"' if v is not None else BLANK


class Result:
    def __init__(self):
        self.header_rows, self.data, self.formats, self.layout = [], [], [], []
        self.count = Counter()
        self.notes = []
        self.key_missing = []
        self.keyed = False
        self.row_matching = ""

    def check(self, scope: str, name: str, a, b, cat: str):
        same = a == b
        self.layout.append([scope, name, _show(a), _show(b), "MATCH" if same else "DIFF"])
        self.count[cat] += not same

    def check_map(self, scope: str, name: str, a: dict, b: dict, cat: str, default: str):
        keys = [k for k in dict.fromkeys([*a, *b]) if a.get(k) != b.get(k)]
        if not keys:
            self.layout.append([scope, name, "same in both files", "same in both files", "MATCH"])
            return
        self.layout.append([scope, name, _show({k: a.get(k, default) for k in keys}),
                            _show({k: b.get(k, default) for k in keys}), "DIFF"])
        self.count[cat] += 1


def _report_only_rows(S: dict, hdrs: list[dict], rows: pl.DataFrame, side: str, label: str, res: Result) -> None:
    """List rows that exist in only one file (rows = DataFrame with row, key_text)."""
    if rows.height == 0:
        return
    hmap = pl.DataFrame({"col": [h["col"] for h in hdrs],
                         "h": [h["raw"] if h["raw"] is not None else BLANK for h in hdrs]},
                        schema={"col": pl.Int32, "h": pl.String})
    extra = (S["cells"].filter(pl.col("value").is_not_null())
             .join(rows, on="row").join(hmap, on="col").sort("row", "col")
             .group_by("row", maintain_order=True)
             .agg(pl.format("{}={}", "h", "value").str.join(" | ").alias("preview"), pl.col("key_text").first()))
    for r in extra.iter_rows(named=True):
        row = [label, r["row"], r["key_text"], "(entire row)", "", "", "", "", f"Row only in File {side}", None, "True", None, *NO_FLAGS]
        row[4 if side == "A" else 5] = f"row {r['row']}"
        row[6 if side == "A" else 7] = _show(r["preview"])
        res.data.append(row)
        res.count["rows_only"] += 1


def compare_sheets(A: dict, B: dict, header_row: int, tol: float, res: Result,
                   key_a: list[str] | None = None, key_b: list[str] | None = None) -> bool:
    """Compare one sheet pair. Returns True when rows were matched by primary key."""
    label = A["name"] if A["name"] == B["name"] else f"{A['name']} ↔ {B['name']}"
    ha, hb = header_list(A, header_row), header_list(B, header_row)

    keyed, cols_a, cols_b = False, None, None
    if key_a:
        cols_a, cols_b = resolve_key(ha, key_a), resolve_key(hb, key_b)
        if cols_a and cols_b:
            keyed = True          # pair the key columns with each other even when their names differ
            ha = header_list(A, header_row, {c: f"\x00key{i}" for i, c in enumerate(cols_a)})
            hb = header_list(B, header_row, {c: f"\x00key{i}" for i, c in enumerate(cols_b)})
        else:
            absent = lambda hdrs, names: [n for n in names if resolve_key(hdrs, [n]) is None]
            missing = [f"File {side} has no column named {', '.join(repr(n) for n in gone)}"
                       for side, gone in (("A", absent(ha, key_a)), ("B", absent(hb, key_b))) if gone]
            res.key_missing.append(f"sheet '{label}', header row {header_row}: {'; '.join(missing)}")
            res.notes.append(f"Key not found in {res.key_missing[-1]}. This sheet was matched by row position instead.")
    by_a, by_b = {h["key"]: h for h in ha}, {h["key"]: h for h in hb}

    # ---- 1. headers ---------------------------------------------------------
    res.count["header_count"] += len(ha) != len(hb)
    for key in dict.fromkeys([*by_a, *by_b]):
        a, b = by_a.get(key), by_b.get(key)
        la = get_column_letter(a["col"]) if a else ""
        lb = get_column_letter(b["col"]) if b else ""
        if a and b:
            name = _name_check(a["raw"], b["raw"])
            pos = "Same" if a["col"] == b["col"] else f"Moved ({la} → {lb})"
            font = "Same" if all(a["cell"].get(c) == b["cell"].get(c) for c in FONT_COLS) else "Differs"
            align = "Same" if all(a["cell"].get(c) == b["cell"].get(c) for c in ALIGN_COLS) else "Differs"
            if name == "Name differs":
                res.count["header_names"] += 1
            else:
                res.count["header_case"] += name != "Exact match"
            res.count["header_order"] += pos != "Same"
            res.count["header_font"] += font != "Same"
            res.count["header_align"] += align != "Same"
        else:
            name = "Only in File A" if a else "Only in File B"
            pos = font = align = "—"
            res.count["header_names"] += 1
        spaced = "pacing" in name
        show = (lambda h: "" if h is None else (BLANK if h["raw"] is None else (_quote(h["raw"]) if spaced else h["raw"])))
        ok = name == "Exact match" and pos == "Same" and font == "Same" and align == "Same"
        res.header_rows.append([
            label, la, show(a), lb, show(b), name, pos,
            _font_text(a["cell"]) if a else "", _font_text(b["cell"]) if b else "", font,
            _align_text(a["cell"]) if a else "", _align_text(b["cell"]) if b else "", align,
            "MATCH" if ok else "DIFF",
        ])

    # ---- pair the rows: by key, or by position --------------------------------
    common_rows = min(A["last_row"], B["last_row"])
    top = min(header_row, common_rows)                       # header and title rows always pair by position
    fixed = pl.DataFrame({"row": list(range(1, top + 1)), "row_b": list(range(1, top + 1)),
                          "key_text": [None] * top}, schema=PAIR_SCHEMA)
    if keyed:
        ka, kb = key_frame(A, cols_a, header_row), key_frame(B, cols_b, header_row)
        kn = [f"k{i}" for i in range(len(cols_a))]
        full = ka.join(kb.select(*kn, "occ", pl.col("row").alias("row_b"), pl.col("key_text").alias("key_text_b")),
                       on=[*kn, "occ"], how="full", coalesce=True)
        matched = full.filter(pl.col("row").is_not_null() & pl.col("row_b").is_not_null()).select("row", "row_b", "key_text")
        only_a = full.filter(pl.col("row_b").is_null()).select("row", "key_text").sort("row")
        only_b = (full.filter(pl.col("row").is_null())
                  .select(pl.col("row_b").alias("row"), pl.col("key_text_b").alias("key_text")).sort("row"))
    else:
        rows = list(range(top + 1, common_rows + 1))
        matched = pl.DataFrame({"row": rows, "row_b": rows, "key_text": [None] * len(rows)}, schema=PAIR_SCHEMA)
        extra = lambda S: pl.DataFrame({"row": list(range(common_rows + 1, S["last_row"] + 1))},
                                       schema={"row": pl.Int32}).with_columns(key_text=pl.lit(None, dtype=pl.String))
        only_a, only_b = extra(A), extra(B)
    pairs = pl.concat([fixed, matched.cast(PAIR_SCHEMA)])

    # ---- join the two sheets cell by cell on (paired row, column name) --------
    common = [k for k in by_a if k in by_b]
    names = {k: (by_a[k]["raw"] if by_a[k]["raw"] is not None else BLANK) for k in by_a}
    map_a = pl.DataFrame({"col": [by_a[k]["col"] for k in common], "key": common,
                          "column": [names[k] for k in common]},
                         schema={"col": pl.Int32, "key": pl.String, "column": pl.String})
    map_b = pl.DataFrame({"col": [by_b[k]["col"] for k in common], "key": common},
                         schema={"col": pl.Int32, "key": pl.String})
    letters = {i: get_column_letter(i) for i in range(1, max(A["last_col"], B["last_col"], 1) + 1)}
    letter = lambda c: pl.col(c).replace_strict(letters, return_dtype=pl.String)

    a = A["cells"].join(map_a, on="col").join(pairs, on="row")
    b = B["cells"].join(map_b, on="col").select(
        pl.col("row").alias("row_b"), "key", pl.all().exclude("row", "key").name.suffix("_b"))
    j = a.join(b, on=["row_b", "key"], how="inner").with_columns(la=letter("col"), lb=letter("col_b"))

    # ---- 2. values, row by row ----------------------------------------------
    d = j.filter(pl.col("row") != header_row)
    v, vb = pl.col("value"), pl.col("value_b")
    both_num = (pl.col("dtype") == "number") & (pl.col("dtype_b") == "number")
    same = (pl.when(both_num).then((pl.col("num") - pl.col("num_b")).abs() <= tol)
              .otherwise(v.eq_missing(vb) & (pl.col("dtype") == pl.col("dtype_b"))))
    kind = (
        pl.when(pl.col("dtype") == "blank").then(pl.lit("Blank in File A"))
        .when(pl.col("dtype_b") == "blank").then(pl.lit("Blank in File B"))
        .when(both_num).then(pl.lit("Number changed"))
        .when(pl.col("dtype") != pl.col("dtype_b")).then(
            pl.when(v == vb).then(pl.format("Data type differs ({} vs {})", "dtype", "dtype_b"))
            .otherwise(pl.format("Value and data type differ ({} vs {})", "dtype", "dtype_b")))
        .when(pl.col("dtype") == "text").then(
            pl.when(v.str.strip_chars() == vb.str.strip_chars()).then(pl.lit("Spacing differs only"))
            .when(v.str.to_lowercase() == vb.str.to_lowercase()).then(pl.lit("Case differs only"))
            .when(v.str.strip_chars().str.to_lowercase() == vb.str.strip_chars().str.to_lowercase())
            .then(pl.lit("Case and spacing differ"))
            .otherwise(pl.lit("Text changed")))
        .when(pl.col("dtype") == "date").then(pl.lit("Date changed"))
        .otherwise(pl.lit("Value changed"))
    )
    # Bring both values to the same platform: a number, or text that reads as a number, becomes a number;
    # everything else is compared as text.
    def as_number(val: str, dtype: str, num: str) -> pl.Expr:
        text = pl.col(val).str.strip_chars()
        return (pl.when(pl.col(dtype) == "number").then(pl.col(num))
                  .when((pl.col(dtype) == "text") & text.str.contains(NUMERIC_TEXT))
                  .then(text.str.replace(r"^\+", "").cast(pl.Float64, strict=False)))

    vals = (d.filter(~same)
             .with_columns(kind=kind, na=as_number("value", "dtype", "num"), nb=as_number("value_b", "dtype_b", "num_b"))
             .with_columns(both=pl.col("na").is_not_null() & pl.col("nb").is_not_null())
             .with_columns(
                 delta=pl.when(pl.col("both")).then(pl.col("nb") - pl.col("na")),
                 type_flag=pl.when(pl.col("both")
                                   | ((pl.col("dtype") == "blank") & pl.col("nb").is_not_null())
                                   | ((pl.col("dtype_b") == "blank") & pl.col("na").is_not_null()))
                           .then(pl.lit("Number")).otherwise(pl.lit("Text")),
                 value_flag=pl.when(pl.col("both")).then((pl.col("nb") - pl.col("na")).abs() > tol)
                            .otherwise(v.ne_missing(vb)))
             .sort("row", "col"))

    def shown(value, dtype, kind_):
        if value is None:
            return BLANK
        padded = dtype == "text" and value != value.strip()
        text = f'"{value}"' if padded or "pacing" in kind_ else value
        return f"{text}  ({dtype})" if "ata type" in kind_ else text

    for r in vals.iter_rows(named=True):
        va, vb_ = shown(r["value"], r["dtype"], r["kind"]), shown(r["value_b"], r["dtype_b"], r["kind"])
        if r["kind"] == "Number changed" and va == vb_:        # differ only beyond the 15th digit
            va, vb_ = repr(r["num"]), repr(r["num_b"])
        res.data.append([label, r["row"], r["key_text"], r["column"], f"{r['la']}{r['row']}",
                         f"{r['lb']}{r['row_b']}", va, vb_, r["kind"], r["type_flag"], str(r["value_flag"]),
                         _clean(r["delta"]), *_threshold_flags(r["delta"])])
        res.count["types" if r["kind"].startswith("Data type differs") else "values"] += 1

    # formulas (the calculated results were already compared above)
    f, fb = pl.col("formula"), pl.col("formula_b")
    forms = (d.filter(f.ne_missing(fb))
              .with_columns(kind=pl.when(f.is_null()).then(pl.lit("Formula only in File B"))
                            .when(fb.is_null()).then(pl.lit("Formula only in File A"))
                            .otherwise(pl.lit("Formula differs")))
              .sort("row", "col"))
    for r in forms.iter_rows(named=True):
        res.data.append([label, r["row"], r["key_text"], r["column"], f"{r['la']}{r['row']}",
                         f"{r['lb']}{r['row_b']}", r["formula"] or "(no formula)",
                         r["formula_b"] or "(no formula)", r["kind"], "Text", "True", None, *NO_FLAGS])
        res.count["formulas"] += 1

    # rows that exist in only one file
    res.count["row_count"] += A["last_row"] != B["last_row"]
    _report_only_rows(A, ha, only_a, "A", label, res)
    _report_only_rows(B, hb, only_b, "B", label, res)

    # ---- 3. cell formatting --------------------------------------------------
    parts = [
        j.filter(pl.col(c) != pl.col(f"{c}_b")).select(
            "row", "row_b", "col", "column", "la", "lb", pl.lit(lab).alias("prop"), pl.lit(cat).alias("cat"),
            pl.col(c).alias("a"), pl.col(f"{c}_b").alias("b"))
        for c, lab, cat in PROPS
    ]
    fm = pl.concat(parts).with_columns(
        area=pl.when(pl.col("row") == header_row).then(pl.lit("Header")).otherwise(pl.lit("Data")),
        off=pl.col("row_b") - pl.col("row"))
    grp = ["col", "column", "la", "lb", "prop", "cat", "area", "a", "b", "off"]
    fm = (fm.sort(*grp, "row")                                   # collapse consecutive rows into one range
            .with_columns(block=((pl.col("row") - pl.col("row").shift(1)).fill_null(0) != 1).cum_sum().over(grp))
            .group_by(*grp, "block")
            .agg(first=pl.col("row").min(), last=pl.col("row").max(), cells=pl.len())
            .sort("first", "col", "prop"))
    rng = lambda L, lo, hi: f"{L}{lo}" if lo == hi else f"{L}{lo}:{L}{hi}"
    for r in fm.iter_rows(named=True):
        res.formats.append([label, r["area"], r["column"], rng(r["la"], r["first"], r["last"]),
                            rng(r["lb"], r["first"] + r["off"], r["last"] + r["off"]),
                            r["cells"], r["prop"], r["a"], r["b"]])
        if r["area"] == "Data":
            res.count[f"fmt_{r['cat']}"] += r["cells"]
        elif r["cat"] == "other":
            res.count["fmt_other"] += r["cells"]

    # ---- 4. sheet layout and everything else ---------------------------------
    s = label
    a2b = dict(zip(pairs["row"].to_list(), pairs["row_b"].to_list()))
    b2a = {rb: ra for ra, rb in a2b.items()}
    col_b2a = {by_b[k]["col"]: by_a[k]["col"] for k in common}

    def to_a(coord: str) -> str:
        """Cell address in File B -> address of the matching cell in File A."""
        col_letters, r = coordinate_from_string(coord)
        c = column_index_from_string(col_letters)
        if c in col_b2a and r in b2a:
            return f"{get_column_letter(col_b2a[c])}{b2a[r]}"
        return f"{coord} (File B cell)"

    if keyed:
        res.layout.append([s, "Row matching", "By key: " + " + ".join(key_a), "By key: " + " + ".join(key_b), "—"])
        for title, expr in (("Primary key: rows with a blank key value", pl.col("null_key")),
                            ("Primary key: rows with a repeated key value", pl.col("dup") & ~pl.col("null_key"))):
            na_, nb_ = ka.filter(expr).height, kb.filter(expr).height
            res.layout.append([s, title, na_, nb_, "OK" if na_ == nb_ == 0 else "WARN"])
            if na_ or nb_:
                res.notes.append(f"Sheet '{label}' - {title.lower()}: {na_} in File A, {nb_} in File B. "
                                 "These rows were matched in order of appearance.")
        ordered = matched.sort("row")
        moved = ordered.filter(pl.col("row_b") < pl.col("row_b").cum_max().shift(1)).height
        res.layout.append([s, "Row order (rows present in both files)", "order as in File A",
                           "same order" if not moved else f"{moved:,} row(s) out of sequence",
                           "MATCH" if not moved else "DIFF"])
        res.count["row_order"] += moved
    else:
        res.layout.append([s, "Row matching", "By row position", "By row position", "—"])

    res.check(s, "Number of headers", len(ha), len(hb), "x")
    res.check(s, "Number of data rows", max(A["last_row"] - header_row, 0), max(B["last_row"] - header_row, 0), "x")
    order_a = [h for h in ha if h["key"] in by_b]
    order_b = [h for h in hb if h["key"] in by_a]
    same_order = [h["key"] for h in order_a] == [h["key"] for h in order_b]
    res.layout.append([s, "Column order (columns present in both files)",
                       _show([h["raw"] or BLANK for h in order_a]), _show([h["raw"] or BLANK for h in order_b]),
                       "MATCH" if same_order else "DIFF"])
    res.check_map(s, "Column widths", {names[k]: A["widths"].get(by_a[k]["col"], "default") for k in common},
                  {names[k]: B["widths"].get(by_b[k]["col"], "default") for k in common}, "layout", "default")
    res.check_map(s, "Row heights", {f"row {r}": h for r, h in A["heights"].items() if r in a2b},
                  {f"row {b2a[r]}": h for r, h in B["heights"].items() if r in b2a}, "layout", "default")
    res.check(s, "Merged cells", A["merged"], B["merged"], "layout")
    res.check(s, "Freeze panes", A["freeze"], B["freeze"], "layout")
    res.check(s, "Auto-filter range", A["filter"], B["filter"], "layout")
    res.check(s, "Hidden columns", A["hidden_cols"], B["hidden_cols"], "layout")
    res.check(s, "Hidden rows", {str(r) for r in A["hidden_rows"]},
              {str(b2a[r]) if r in b2a else f"{r} (File B row)" for r in B["hidden_rows"]}, "layout")
    res.check(s, "Sheet visibility", A["state"], B["state"], "layout")
    res.check_map(s, "Cell comments", A["comments"], {to_a(k): t for k, t in B["comments"].items()}, "extras", "none")
    res.check_map(s, "Hyperlinks", A["links"], {to_a(k): t for k, t in B["links"].items()}, "extras", "none")
    res.check(s, "Data validation rules", A["validations"], B["validations"], "extras")
    res.check(s, "Conditional formatting rules", A["cond_formats"], B["cond_formats"], "extras")
    for side, S in (("A", A), ("B", B)):
        if S["uncached"]:
            res.notes.append(f"File {side}, sheet '{S['name']}': {S['uncached']} formula cell(s) have no saved result "
                             "(file was never calculated in Excel), so their values compare as blank.")
    return keyed


def compare_workbooks(path_a: Path, path_b: Path, header_row: int = 1, tol: float = 0.0,
                      sheet_a: str | None = None, sheet_b: str | None = None,
                      key_a: list[str] | None = None, key_b: list[str] | None = None) -> Result:
    key_a, key_b = key_a or key_b, key_b or key_a          # one given -> same names in both files
    if key_a and len(key_a) != len(key_b):
        raise ValueError(f"The key has {len(key_a)} column(s) for File A but {len(key_b)} for File B")

    A, B = open_book(path_a), open_book(path_b)
    na, nb = A["wb"].sheetnames, B["wb"].sheetnames
    res = Result()

    if sheet_a or sheet_b:
        pairs = [(sheet_a or na[0], sheet_b or nb[0])]
        res.notes.append("Only the sheets named on the command line were compared.")
    elif len(na) == 1 and len(nb) == 1:
        pairs = [(na[0], nb[0])]
    else:
        lookup = {n.casefold(): n for n in nb}
        pairs = [(n, lookup[n.casefold()]) for n in na if n.casefold() in lookup]
    for name in (p[0] for p in pairs):
        if name not in na:
            raise ValueError(f"Sheet '{name}' not found in File A")
    for name in (p[1] for p in pairs):
        if name not in nb:
            raise ValueError(f"Sheet '{name}' not found in File B")

    if not (sheet_a or sheet_b):
        res.check("Workbook", "Number of sheets", len(na), len(nb), "workbook")
        res.check("Workbook", "Sheet names (in order)", na, nb, "workbook")
        res.check("Workbook", "Sheets without a partner in the other file",
                  [n for n in na if n not in [p[0] for p in pairs]],
                  [n for n in nb if n not in [p[1] for p in pairs]], "workbook")

    used_key = [compare_sheets(read_sheet(A, a_name), read_sheet(B, b_name), header_row, tol, res, key_a, key_b)
                for a_name, b_name in pairs]
    if key_a and not any(used_key):
        raise ValueError("Key column(s) not found - " + " | ".join(res.key_missing))
    res.keyed = any(used_key)
    if res.keyed:
        res.row_matching = (f"By primary key - File A: {' + '.join(key_a)}  ↔  File B: {' + '.join(key_b)} "
                            "(key column names matched ignoring case)")
    else:
        res.row_matching = f"By position (header in row {header_row}; row N of File A against row N of File B)"
    return res


# --------------------------------------------------------------------------- report
SUMMARY_CHECKS = [   # (label, counter key, unit, report sheet)
    ("Sheet count and sheet names", "workbook", "check(s)", "Layout & Other"),
    ("Number of headers", "header_count", "sheet(s)", "Headers"),
    ("Header names (missing, extra or renamed columns)", "header_names", "header(s)", "Headers"),
    ("Header case and spacing", "header_case", "header(s)", "Headers"),
    ("Header order", "header_order", "header(s)", "Headers"),
    ("Header font", "header_font", "header(s)", "Headers"),
    ("Header alignment", "header_align", "header(s)", "Headers"),
    ("Number of data rows", "row_count", "sheet(s)", "Data Differences"),
    ("Rows present in only one file", "rows_only", "row(s)", "Data Differences"),
    ("Row order (when matched by key)", "row_order", "row(s)", "Layout & Other"),
    ("Cell values (row by row)", "values", "cell(s)", "Data Differences"),
    ("Data types (same value, different type)", "types", "cell(s)", "Data Differences"),
    ("Formulas", "formulas", "cell(s)", "Data Differences"),
    ("Cell font (data rows)", "fmt_font", "cell(s)", "Format Differences"),
    ("Cell alignment (data rows)", "fmt_align", "cell(s)", "Format Differences"),
    ("Number format, fill colour and borders", "fmt_other", "cell(s)", "Format Differences"),
    ("Sheet layout (widths, heights, merged cells, freeze panes, hidden rows/columns, filter)",
     "layout", "check(s)", "Layout & Other"),
    ("Comments, hyperlinks, data validation, conditional formatting", "extras", "check(s)", "Layout & Other"),
]
OK_WORDS = {"MATCH", "Exact match", "Same", "IDENTICAL", "OK", "False"}


def write_report(res: Result, path_a: Path, path_b: Path, out: Path, header_row: int, tol: float) -> bool:
    identical = not any(res.count[key] for _, key, _, _ in SUMMARY_CHECKS)
    wb = xlsxwriter.Workbook(str(out), {"strings_to_formulas": False, "strings_to_numbers": False,
                                        "strings_to_urls": False})
    base = {"font_name": "Calibri", "font_size": 11, "valign": "top", "text_wrap": True, "border": 1}
    fmt = {
        "head": wb.add_format({**base, "bold": True, "font_color": "#1F3864", "bg_color": "#DDEBF7", "valign": "vcenter"}),
        "cell": wb.add_format(base),
        "alt": wb.add_format({**base, "bg_color": "#F2F2F2"}),
        "num": wb.add_format({**base, "align": "center"}),
        "num_alt": wb.add_format({**base, "align": "center", "bg_color": "#F2F2F2"}),
        "ok": wb.add_format({**base, "bold": True, "font_color": "#1A3E8C"}),
        "ok_alt": wb.add_format({**base, "bold": True, "font_color": "#1A3E8C", "bg_color": "#F2F2F2"}),
        "bad": wb.add_format({**base, "bold": True, "font_color": "#C00000"}),
        "bad_alt": wb.add_format({**base, "bold": True, "font_color": "#C00000", "bg_color": "#F2F2F2"}),
        "title": wb.add_format({"font_name": "Calibri", "font_size": 16, "bold": True, "font_color": "#1F3864"}),
        "label": wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True}),
        "text": wb.add_format({"font_name": "Calibri", "font_size": 11}),
        "verdict_ok": wb.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "font_color": "#1A3E8C"}),
        "verdict_bad": wb.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "font_color": "#C00000"}),
    }

    def table(ws, top, columns, rows, status_cols=(), empty="No differences found"):
        """columns = [(title, width)]; a SrNo. column is added in front. Returns the next free row."""
        ws.set_column(0, 0, 6)
        ws.write(top, 0, "SrNo.", fmt["head"])
        for c, (title, width) in enumerate(columns, 1):
            ws.set_column(c, c, width)
            ws.write(top, c, title, fmt["head"])
        shown = rows[:MAX_REPORT_ROWS]
        zebra = len(shown) > 6
        for i, row in enumerate(shown):
            alt = "_alt" if zebra and i % 2 else ""
            plain = fmt["alt"] if alt else fmt["cell"]
            ws.write_number(top + 1 + i, 0, i + 1, fmt["num" + alt])
            for c, val in enumerate(row, 1):
                if val is None:
                    ws.write_blank(top + 1 + i, c, None, plain)
                elif c in status_cols and val != "—":
                    ws.write_string(top + 1 + i, c, str(val), fmt[("ok" if val in OK_WORDS else "bad") + alt])
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    ws.write_number(top + 1 + i, c, val, fmt["num" + alt])
                else:
                    ws.write_string(top + 1 + i, c, str(val), plain)
        if not shown:
            ws.write_string(top + 1, 1, empty, fmt["text"])
        if len(rows) > MAX_REPORT_ROWS:
            ws.write_string(top + 1 + len(shown), 1,
                            f"{len(rows) - MAX_REPORT_ROWS:,} more row(s) not listed (limit {MAX_REPORT_ROWS:,}).", fmt["label"])
        if shown:
            ws.autofilter(top, 0, top + len(shown), len(columns))
        ws.freeze_panes(top + 1, 0)
        ws.set_landscape(); ws.fit_to_pages(1, 0); ws.repeat_rows(top)
        return top + 2 + len(shown)

    # ---- Summary
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.write(0, 0, "Excel Comparison Report", fmt["title"])
    info = [("File A", str(path_a)), ("File B", str(path_b)),
            ("Generated", dt.datetime.now().strftime("%d-%b-%Y %H:%M:%S")),
            ("Column matching", "By header name, ignoring case and leading/trailing spaces"),
            ("Row matching", res.row_matching),
            ("Number tolerance", f"{tol:g}")]
    for i, (k, val) in enumerate(info, 2):
        ws.write(i, 0, k, fmt["label"]); ws.write(i, 2, val, fmt["text"])
    failed = sum(1 for _, key, _, _ in SUMMARY_CHECKS if res.count[key])
    ws.write(9, 0, "RESULT", fmt["label"])
    ws.write(9, 2, "IDENTICAL - no differences found" if identical else
             f"NOT IDENTICAL - {failed} of {len(SUMMARY_CHECKS)} checks found differences",
             fmt["verdict_ok" if identical else "verdict_bad"])
    rows = [[label, "DIFF" if res.count[key] else "MATCH",
             f"{res.count[key]:,} {unit}" if res.count[key] else "0", sheet]
            for label, key, unit, sheet in SUMMARY_CHECKS]
    end = table(ws, 11, [("Check", 62), ("Result", 10), ("Differences", 16), ("Details in sheet", 20)], rows, status_cols={2})
    ws.freeze_panes(0, 0)
    for i, note in enumerate(res.notes):
        ws.write(end + i, 1, f"Note: {note}", fmt["text"])

    # ---- Headers
    ws = wb.add_worksheet("Headers")
    table(ws, 0, [("Sheet", 22), ("Col (A)", 8), ("Header in File A", 22), ("Col (B)", 8), ("Header in File B", 22),
                  ("Name Check", 18), ("Position", 14), ("Font in File A", 24), ("Font in File B", 24),
                  ("Font Check", 11), ("Alignment in File A", 18), ("Alignment in File B", 18),
                  ("Alignment Check", 11), ("Status", 9)],
          res.header_rows, status_cols={6, 7, 10, 13, 14}, empty="No headers found")

    # ---- Data Differences (the Key column appears only when a primary key was used)
    ws = wb.add_worksheet("Data Differences")
    cols = [("Sheet", 22), ("Row", 7), ("Key", 22), ("Column", 20), ("Cell (A)", 9), ("Cell (B)", 9),
            ("Value in File A", 30), ("Value in File B", 30), ("Difference Type", 34), ("Difference Type Flag", 12), ("Difference Type Value Flag", 12),
            ("Number Diff (B − A)", 13),
            *[(f"Abs Diff > {t:g}", 11) for t in DIFF_THRESHOLDS]]
    data = res.data
    if not res.keyed:
        cols = cols[:2] + cols[3:]
        data = [r[:2] + r[3:] for r in data]
    flag_col = [title for title, _ in cols].index("Difference Type Value Flag") + 1
    table(ws, 0, cols, data, status_cols={flag_col})

    # ---- Format Differences
    ws = wb.add_worksheet("Format Differences")
    table(ws, 0, [("Sheet", 22), ("Area", 9), ("Column", 20), ("Cells (A)", 11), ("Cells (B)", 11),
                  ("Cell Count", 8), ("Property", 22), ("File A", 24), ("File B", 24)], res.formats)

    # ---- Layout & Other
    ws = wb.add_worksheet("Layout & Other")
    table(ws, 0, [("Scope", 22), ("Check", 40), ("File A", 44), ("File B", 44), ("Status", 9)],
          res.layout, status_cols={5})

    wb.close()
    return identical


# --------------------------------------------------------------------------- CLI
def _cols(text: str | None) -> list[str] | None:
    if not text:
        return None
    names = [t.strip() for t in text.split(",") if t.strip()]
    return names or None


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Compare two Excel workbooks and write a difference report.")
    p.add_argument("file_a", type=Path)
    p.add_argument("file_b", type=Path)
    p.add_argument("report", type=Path, nargs="?", default=Path("comparison_report.xlsx"))
    p.add_argument("--key", help='primary key column(s), same names in both files, e.g. "emp id" or "a2,a3"')
    p.add_argument("--key-a", help='primary key column(s) of File A, e.g. "a1" or "a2,a3"')
    p.add_argument("--key-b", help='primary key column(s) of File B, e.g. "b1" or "b2,b3"')
    p.add_argument("--header-row", type=int, default=1, help="row number that holds the column headers (default 1)")
    p.add_argument("--tolerance", type=float, default=0.0, help="allowed absolute difference between numbers (default 0)")
    p.add_argument("--sheet-a", help="compare only this sheet of File A")
    p.add_argument("--sheet-b", help="compare only this sheet of File B")
    args = p.parse_args(argv)
    try:
        res = compare_workbooks(args.file_a, args.file_b, args.header_row, args.tolerance, args.sheet_a, args.sheet_b,
                                _cols(args.key_a) or _cols(args.key), _cols(args.key_b) or _cols(args.key))
        identical = write_report(res, args.file_a, args.file_b, args.report, args.header_row, args.tolerance)
    except (FileNotFoundError, ValueError, KeyError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print("IDENTICAL" if identical else "NOT IDENTICAL")
    print(f"  Row matching: {res.row_matching}")
    for label, key, unit, _ in SUMMARY_CHECKS:
        if res.count[key]:
            print(f"  DIFF  {label}: {res.count[key]:,} {unit}")
    for note in res.notes:
        print(f"  NOTE  {note}")
    print(f"Report written to {args.report}")
    return 0 if identical else 1


if __name__ == "__main__":
    sys.exit(main())
