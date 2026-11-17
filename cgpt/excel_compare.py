#!/usr/bin/env python3
"""Compare Excel worksheets by row position or optional validated primary keys.

Python 3.12; dependencies: polars and openpyxl.
The source workbooks are opened read-only. A third .xlsx file is generated.
"""

from __future__ import annotations

import argparse
import json
import math
import pickle
import posixpath
import sqlite3
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time
from itertools import islice, zip_longest
from pathlib import Path
from typing import Any, Iterator, Sequence
from zipfile import ZipFile

import polars as pl
from openpyxl import Workbook, load_workbook
from openpyxl.cell.read_only import EmptyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
MAX_EXCEL_ROWS = 1_048_576
DEFAULT_MAX_DETAILS = 100_000
DATA_CATEGORIES = frozenset({"Value", "Data type", "Row"})
COLORS = {
    "navy": "263D52",
    "pale_red": "FCEAEA",
    "pale_amber": "FFF2DB",
    "pale_blue": "E9F1F8",
    "gray": "EEF0F2",
    "muted": "607080",
}


@dataclass(slots=True)
class Finding:
    category: str
    issue: str
    column: str
    a_cell: str
    a_value: Any
    b_cell: str
    b_value: Any
    note: str


@dataclass(slots=True)
class Findings:
    max_details: int
    data_rows: list[Finding] = field(default_factory=list)
    other_rows: list[Finding] = field(default_factory=list)
    counts: Counter[str] = field(default_factory=Counter)
    issue_counts: Counter[str] = field(default_factory=Counter)
    total: int = 0

    def add(
        self,
        category: str,
        issue: str,
        column: str,
        a_cell: str,
        a_value: Any,
        b_cell: str,
        b_value: Any,
        note: str,
    ) -> None:
        self.total += 1
        self.counts[category] += 1
        self.issue_counts[issue] += 1
        destination = self.data_rows if category in DATA_CATEGORIES else self.other_rows
        if len(destination) < self.max_details:
            destination.append(
                Finding(category, issue, column, a_cell, a_value, b_cell, b_value, note)
            )

    @property
    def rows(self) -> list[Finding]:
        """Retain the original programmatic interface across both report sections."""
        return [*self.other_rows, *self.data_rows]

    @property
    def data_total(self) -> int:
        return sum(self.counts[category] for category in DATA_CATEGORIES)

    @property
    def other_total(self) -> int:
        return self.total - self.data_total

    @property
    def truncated(self) -> bool:
        return self.total > len(self.data_rows) + len(self.other_rows)


@dataclass(slots=True)
class SheetLayout:
    default_row_height: float = 15.0
    default_column_width: float = 13.0
    row_heights: dict[int, float] = field(default_factory=dict)
    column_widths: dict[int, float] = field(default_factory=dict)
    hidden_rows: set[int] = field(default_factory=set)
    hidden_columns: set[int] = field(default_factory=set)
    merged_ranges: set[str] = field(default_factory=set)
    freeze_pane: str = ""

    def row_height(self, row_number: int) -> float:
        return self.row_heights.get(row_number, self.default_row_height)

    def column_width(self, column_number: int) -> float:
        return self.column_widths.get(column_number, self.default_column_width)


@dataclass(slots=True)
class StoredCell:
    """Small read-only proxy for a File B cell persisted in the key index."""

    value: Any
    data_type: str
    _style_id: int


@dataclass(slots=True)
class Comparison:
    file_a: Path
    file_b: Path
    sheet_a: str
    sheet_b: str
    sheet_count_a: int
    sheet_count_b: int
    header_count_a: int
    header_count_b: int
    row_count_a: int
    row_count_b: int
    missing_headers: list[str]
    additional_headers: list[str]
    case_changes: int
    order_changes: int
    findings: Findings
    order_notes_a: str = ""
    order_notes_b: str = ""
    key_columns_a: tuple[str, ...] = ()
    key_columns_b: tuple[str, ...] = ()

    @property
    def data_identical(self) -> bool:
        return not any(
            self.findings.counts[key] for key in ("Workbook", "Header", "Value", "Data type", "Row")
        )

    @property
    def formatting_identical(self) -> bool:
        return self.findings.counts["Formatting"] == 0

    @property
    def identical(self) -> bool:
        return self.data_identical and self.formatting_identical


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml_sheet_path(archive: ZipFile, sheet_name: str) -> str:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall(f"{{{NS_REL_PACKAGE}}}Relationship")
    }
    for sheet in workbook_root.findall(f".//{{{NS_MAIN}}}sheet"):
        if sheet.attrib.get("name") != sheet_name:
            continue
        relationship_id = sheet.attrib.get(f"{{{NS_REL_DOC}}}id")
        if relationship_id not in rels:
            raise ValueError(f"Could not resolve worksheet XML for {sheet_name!r}.")
        target = rels[relationship_id]
        return (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )
    raise ValueError(f"Worksheet {sheet_name!r} was not found in the workbook manifest.")


def read_sheet_layout(file_path: Path, sheet_name: str) -> SheetLayout:
    """Read row/column dimensions and merged cells without loading the worksheet."""
    layout = SheetLayout()
    with ZipFile(file_path) as archive:
        xml_path = _xml_sheet_path(archive, sheet_name)
        with archive.open(xml_path) as xml_stream:
            parser = ET.iterparse(xml_stream, events=("start", "end"))
            _, root = next(parser)
            sheet_data_node = None
            for event, elem in parser:
                if event == "start" and _localname(elem.tag) == "sheetData":
                    sheet_data_node = elem
                if event != "end":
                    continue
                tag = _localname(elem.tag)
                attrs = elem.attrib
                if tag == "sheetFormatPr":
                    if "defaultRowHeight" in attrs:
                        layout.default_row_height = float(attrs["defaultRowHeight"])
                    if "defaultColWidth" in attrs:
                        layout.default_column_width = float(attrs["defaultColWidth"])
                elif tag == "col":
                    start = int(attrs.get("min", "1"))
                    end = int(attrs.get("max", str(start)))
                    for col in range(start, end + 1):
                        if "width" in attrs:
                            layout.column_widths[col] = float(attrs["width"])
                        if attrs.get("hidden") == "1":
                            layout.hidden_columns.add(col)
                elif tag == "row":
                    index = int(attrs["r"])
                    if "ht" in attrs:
                        layout.row_heights[index] = float(attrs["ht"])
                    if attrs.get("hidden") == "1":
                        layout.hidden_rows.add(index)
                    # The row container needs clearing too; merely clearing
                    # each cell leaves millions of empty XML nodes attached.
                    if sheet_data_node is not None:
                        sheet_data_node.clear()
                elif tag == "mergeCell" and "ref" in attrs:
                    layout.merged_ranges.add(attrs["ref"])
                elif tag == "pane":
                    layout.freeze_pane = attrs.get("topLeftCell", "")
                if tag in {"col", "mergeCell", "sheetFormatPr", "pane"}:
                    elem.clear()
            root.clear()
    return layout


def _value_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if math.isnan(value):
            return "NaN"
        if math.isinf(value):
            return "Infinity" if value > 0 else "-Infinity"
        return format(value, ".15g")
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


def _display_value(value: Any) -> Any:
    if value is None:
        return "(blank)"
    if value == "":
        return "(empty text)"
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, (int, float, str, bool)):
        return value
    return str(value)


def _cell_type(cell: Any) -> str:
    if cell is None or cell.value is None:
        return "Blank"
    if cell.data_type == "f":
        return "Formula"
    if cell.data_type == "e":
        return "Error"
    if cell.data_type == "b":
        return "Boolean"
    if isinstance(cell.value, (datetime, date, time)):
        return "Date/time"
    if cell.data_type == "n":
        return "Number"
    return "Text"


def _color(color: Any) -> str:
    if color is None:
        return "Default"
    kind = color.type
    if kind == "rgb" and isinstance(color.rgb, str):
        return "#" + color.rgb[-6:].upper()
    if kind == "theme":
        return f"Theme {color.theme} (tint {color.tint:g})"
    if kind == "indexed":
        return f"Indexed {color.indexed}"
    if kind == "auto":
        return "Automatic"
    return str(kind or "Default")


def _fill_name(fill: Any) -> str:
    if not fill or not fill.patternType:
        return "#FFFFFF"
    if fill.patternType == "solid":
        return _color(fill.fgColor)
    return f"{fill.patternType}: {_color(fill.fgColor)}"


def _border_name(border: Any) -> str:
    result: list[str] = []
    for side_name in ("left", "right", "top", "bottom", "diagonal"):
        side = getattr(border, side_name, None)
        if side is not None and side.style:
            result.append(f"{side_name}={side.style}/{_color(side.color)}")
    return ", ".join(result) if result else "None"


def _style_id(cell: Any) -> int:
    # ReadOnlyCell exposes the stored style index as _style_id, whereas normal
    # worksheet cells expose style_id. The cache belongs to one workbook only.
    raw_style_id = getattr(cell, "_style_id", None)
    if raw_style_id is None:
        raw_style_id = getattr(cell, "style_id", 0)
    return int(raw_style_id or 0)


def _style_values(cell: Any, cache: dict[int, dict[str, Any]]) -> dict[str, Any]:
    style_id = _style_id(cell)
    if style_id in cache:
        return cache[style_id]
    font = cell.font
    alignment = cell.alignment
    values = {
        "Font name": font.name or "(default)",
        "Font size": font.sz if font.sz is not None else "(default)",
        "Bold": bool(font.bold),
        "Italic": bool(font.italic),
        "Underline": font.underline or "None",
        "Strikethrough": bool(font.strike),
        "Font colour": _color(font.color),
        "Background fill": _fill_name(cell.fill),
        "Horizontal alignment": (alignment.horizontal or "General").capitalize(),
        "Vertical alignment": (alignment.vertical or "Bottom").capitalize(),
        "Wrap text": bool(alignment.wrap_text),
        "Text rotation": alignment.textRotation or 0,
        "Number format": cell.number_format or "General",
        "Border": _border_name(cell.border),
        "Cell locked": bool(cell.protection.locked),
        "Cell hidden": bool(cell.protection.hidden),
    }
    cache[style_id] = values
    return values


def _cell_at(row: tuple[Any, ...] | None, index: int) -> Any:
    if row is None or index >= len(row):
        return EmptyCell()
    return row[index]


def _column_names(row: tuple[Any, ...]) -> list[str]:
    return [str(cell.value) if cell.value is not None else "" for cell in row]


def _header_map(headers: list[str]) -> dict[str, list[int]]:
    mapping: dict[str, list[int]] = defaultdict(list)
    for index, header in enumerate(headers):
        if header:
            mapping[header.casefold()].append(index)
    return mapping


def _resolve_key_columns(
    headers_a: list[str],
    headers_b: list[str],
    key_a: Sequence[str] | str | None,
    key_b: Sequence[str] | str | None,
) -> list[tuple[int, int]]:
    """Resolve both key definitions, preserving the order of composite parts."""
    if isinstance(key_a, str):
        key_a = key_a.split(",")
    if isinstance(key_b, str):
        key_b = key_b.split(",")
    if key_a is None and key_b is None:
        return []
    if not key_a or not key_b or len(key_a) != len(key_b):
        raise ValueError("Supply --key-a and --key-b together, with the same number of columns.")
    indices: list[list[int]] = []
    for headers, names, source in ((headers_a, key_a, "A"), (headers_b, key_b, "B")):
        mapping = _header_map(headers)
        resolved: list[int] = []
        seen: set[str] = set()
        for name in names:
            normalized = name.strip().casefold()
            if not normalized or normalized in seen:
                raise ValueError(f"File {source} primary-key column names must be nonblank and distinct.")
            seen.add(normalized)
            positions = mapping.get(normalized, [])
            if len(positions) != 1:
                raise ValueError(
                    f"File {source} primary-key column {name!r} must match exactly one header "
                    f"(ignoring case); found {len(positions)}."
                )
            resolved.append(positions[0])
        indices.append(resolved)
    return list(zip(indices[0], indices[1], strict=True))


def _key_string(row: tuple[Any, ...], positions: Sequence[int], source: str, row_no: int) -> str:
    """Use typed, exact values so 001, 1, and TRUE never collide."""
    parts: list[tuple[str, str]] = []
    for position in positions:
        cell = _cell_at(row, position)
        value = cell.value
        address = f"{get_column_letter(position + 1)}{row_no}"
        if value is None or isinstance(value, str) and not value.strip():
            raise ValueError(f"File {source} primary key contains a blank at {address}.")
        if cell.data_type == "f":
            raise ValueError(f"File {source} primary key at {address} is a formula; use saved key values.")
        if isinstance(value, bool):
            parts.append(("boolean", str(value)))
        elif isinstance(value, (int, float)):
            if not math.isfinite(float(value)):
                raise ValueError(f"File {source} primary key at {address} is not a finite number.")
            parts.append(("number", _value_text(value) or ""))
        elif isinstance(value, (datetime, date, time)):
            parts.append((type(value).__name__, value.isoformat()))
        else:
            parts.append(("text", str(value)))
    return json.dumps(parts, ensure_ascii=False, separators=(",", ":"))


def _key_label(row: tuple[Any, ...], positions: Sequence[int], headers: list[str]) -> str:
    pairs = [f"{headers[position]}={_value_text(_cell_at(row, position).value)}" for position in positions]
    label = ", ".join(pairs)
    return label if len(label) <= 200 else label[:197] + "..."


def _is_different_value(a: Any, b: Any) -> str:
    if isinstance(a, str) and isinstance(b, str):
        if a.casefold() == b.casefold() and a != b:
            return "Value case changed"
        if a.strip() == b.strip() and a != b:
            return "Trailing/leading space"
    return "Value changed"


def _compare_cell_styles(
    left: Any,
    right: Any,
    cache_a: dict[int, dict[str, Any]],
    cache_b: dict[int, dict[str, Any]],
    changed_style_cache: dict[tuple[int, int], list[tuple[str, Any, Any]]],
    column: str,
    cell_a: str,
    cell_b: str,
    note: str,
    findings: Findings,
) -> None:
    pair = (_style_id(left), _style_id(right))
    if pair not in changed_style_cache:
        a_styles = _style_values(left, cache_a)
        b_styles = _style_values(right, cache_b)
        changed_style_cache[pair] = [
            (property_name, a_value, b_styles[property_name])
            for property_name, a_value in a_styles.items()
            if a_value != b_styles[property_name]
        ]
    for property_name, a_value, b_value in changed_style_cache[pair]:
        findings.add("Formatting", property_name, column, cell_a, a_value, cell_b, b_value, note)


def _compare_headers(
    header_cells_a: tuple[Any, ...],
    header_cells_b: tuple[Any, ...],
    header_row: int,
    cache_a: dict[int, dict[str, Any]],
    cache_b: dict[int, dict[str, Any]],
    changed_style_cache: dict[tuple[int, int], list[tuple[str, Any, Any]]],
    findings: Findings,
    key_positions: list[tuple[int, int]] | None = None,
) -> tuple[list[tuple[int, int, str]], list[str], list[str], int, int, str, str]:
    headers_a = _column_names(header_cells_a)
    headers_b = _column_names(header_cells_b)
    map_a = _header_map(headers_a)
    map_b = _header_map(headers_b)
    unique_pairs: list[tuple[int, int, str]] = []
    missing: list[str] = []
    extra: list[str] = []
    case_changes = 0
    moves: list[tuple[str, str, str]] = []

    for source, headers, source_name in ((map_a, headers_a, "A"), (map_b, headers_b, "B")):
        for key, positions in source.items():
            if len(positions) > 1:
                for index in positions:
                    cell = get_column_letter(index + 1) + str(header_row)
                    findings.add(
                        "Header", "Duplicate header", headers[index],
                        cell if source_name == "A" else "—",
                        headers[index] if source_name == "A" else "Not present",
                        cell if source_name == "B" else "—",
                        headers[index] if source_name == "B" else "Not present",
                        "Ambiguous case-insensitive column; data comparison skipped",
                    )

    for index, header in enumerate(headers_a):
        if not header:
            findings.add("Header", "Blank header", f"Column {get_column_letter(index + 1)}",
                         get_column_letter(index + 1) + str(header_row), "(blank)",
                         "—", "—", "File A has an unnamed header")
    for index, header in enumerate(headers_b):
        if not header:
            findings.add("Header", "Blank header", f"Column {get_column_letter(index + 1)}",
                         "—", "—", get_column_letter(index + 1) + str(header_row),
                         "(blank)", "File B has an unnamed header")

    # Reserve explicitly mapped key columns first. A differently named key in
    # File B is paired with File A, rather than being treated as two unrelated
    # columns. Other columns still match case-insensitively by their own names.
    paired: dict[int, int] = dict(key_positions or [])
    used_b = set(paired.values())
    for key, a_positions in map_a.items():
        b_positions = map_b.get(key, [])
        if len(a_positions) == len(b_positions) == 1:
            a_idx, b_idx = a_positions[0], b_positions[0]
            if a_idx not in paired and b_idx not in used_b:
                paired[a_idx] = b_idx
                used_b.add(b_idx)

    for a_index, b_index in sorted(paired.items()):
        name_a, name_b = headers_a[a_index], headers_b[b_index]
        a_ref = get_column_letter(a_index + 1) + str(header_row)
        b_ref = get_column_letter(b_index + 1) + str(header_row)
        unique_pairs.append((a_index, b_index, name_a))
        if name_a.casefold() != name_b.casefold():
            findings.add(
                "Header", "Mapped key header", name_a, a_ref, name_a, b_ref, name_b,
                "Different names explicitly mapped as primary-key columns",
            )
        elif name_a != name_b:
            case_changes += 1
            findings.add("Header", "Case change", name_a, a_ref, name_a, b_ref, name_b,
                         "Matched ignoring header case")
        if a_index != b_index:
            moves.append((name_a, get_column_letter(a_index + 1), get_column_letter(b_index + 1)))
            findings.add(
                "Header", "Column moved", name_a, a_ref, f"Position {a_index + 1}",
                b_ref, f"Position {b_index + 1}", "Matched despite new position",
            )
        _compare_cell_styles(
            _cell_at(header_cells_a, a_index),
            _cell_at(header_cells_b, b_index),
            cache_a, cache_b, changed_style_cache,
            name_a, a_ref, b_ref, "Header formatting", findings,
        )

    for key, positions in map_a.items():
        for i in positions:
            if i not in paired and (key not in map_b or any(j in used_b for j in map_b[key])):
                missing.append(headers_a[i])
                findings.add(
                    "Header", "Missing header", headers_a[i],
                    get_column_letter(i + 1) + str(header_row), headers_a[i],
                    "—", "Not matched", "File A header has no available File B counterpart",
                )
    for key, positions in map_b.items():
        for i in positions:
            if i not in used_b and (key not in map_a or any(j in paired for j in map_a[key])):
                extra.append(headers_b[i])
                findings.add(
                    "Header", "Additional header", headers_b[i], "—", "Not matched",
                    get_column_letter(i + 1) + str(header_row), headers_b[i],
                    "File B header has no available File A counterpart",
                )

    def ordered_note(index: int) -> str:
        return " / ".join(f"{name} {a if index == 1 else b}" for name, a, b in moves[:3])

    return unique_pairs, missing, extra, case_changes, len(moves), ordered_note(1), ordered_note(2)


def _compare_layout(
    layout_a: SheetLayout,
    layout_b: SheetLayout,
    header_row: int,
    common_row_count: int,
    pairs: list[tuple[int, int, str]],
    findings: Findings,
) -> None:
    for row_no in range(header_row, header_row + common_row_count + 1):
        a_height = layout_a.row_height(row_no)
        b_height = layout_b.row_height(row_no)
        if not math.isclose(a_height, b_height):
            findings.add(
                "Formatting", "Row height", f"Row {row_no}", f"Row {row_no}", f"{a_height:g} pt",
                f"Row {row_no}", f"{b_height:g} pt", "Worksheet layout",
            )
        hidden_a = row_no in layout_a.hidden_rows
        hidden_b = row_no in layout_b.hidden_rows
        if hidden_a != hidden_b:
            findings.add("Formatting", "Hidden row", f"Row {row_no}",
                         f"Row {row_no}", str(hidden_a), f"Row {row_no}", str(hidden_b),
                         "Worksheet layout")

    for a_idx, b_idx, column in pairs:
        col_a, col_b = a_idx + 1, b_idx + 1
        width_a = layout_a.column_width(col_a)
        width_b = layout_b.column_width(col_b)
        if not math.isclose(width_a, width_b):
            findings.add(
                "Formatting", "Column width", column,
                f"Column {get_column_letter(col_a)}", f"{width_a:g}",
                f"Column {get_column_letter(col_b)}", f"{width_b:g}",
                "Worksheet layout",
            )
        hidden_a = col_a in layout_a.hidden_columns
        hidden_b = col_b in layout_b.hidden_columns
        if hidden_a != hidden_b:
            findings.add(
                "Formatting", "Hidden column", column,
                f"Column {get_column_letter(col_a)}", str(hidden_a),
                f"Column {get_column_letter(col_b)}", str(hidden_b),
                "Worksheet layout",
            )
    for name in sorted(layout_a.merged_ranges - layout_b.merged_ranges):
        findings.add("Formatting", "Merged range removed", name, name, "Present",
                     "—", "Missing", "Merged-cell layout")
    for name in sorted(layout_b.merged_ranges - layout_a.merged_ranges):
        findings.add("Formatting", "Merged range added", name, "—", "Missing",
                     name, "Present", "Merged-cell layout")
    if layout_a.freeze_pane != layout_b.freeze_pane:
        findings.add("Formatting", "Freeze pane", "Worksheet",
                     "Pane", layout_a.freeze_pane or "(none)",
                     "Pane", layout_b.freeze_pane or "(none)", "Worksheet layout")


def _batches(iterator: Iterator[tuple[Any, Any]], size: int) -> Iterator[list[tuple[Any, Any]]]:
    while batch := list(islice(iterator, size)):
        yield batch


def _compare_matched_rows(
    matched: list[tuple[tuple[Any, ...], tuple[Any, ...], int, int, str]],
    pairs: list[tuple[int, int, str]],
    cache_a: dict[int, dict[str, Any]],
    cache_b: dict[int, dict[str, Any]],
    changed_style_cache: dict[tuple[int, int], list[tuple[str, Any, Any]]],
    findings: Findings,
) -> None:
    for a_idx, b_idx, col_name in pairs:
        if not matched:
            break
        left_cells = [_cell_at(a_row, a_idx) for a_row, _, _, _, _ in matched]
        right_cells = [_cell_at(b_row, b_idx) for _, b_row, _, _, _ in matched]

        # Polars compares each batch with null-safe, vectorised expressions.
        a_series = pl.Series("a", [_value_text(cell.value) for cell in left_cells], dtype=pl.String)
        b_series = pl.Series("b", [_value_text(cell.value) for cell in right_cells], dtype=pl.String)
        changed = a_series.ne_missing(b_series).to_list()

        for position, (_, _, row_a, row_b, note) in enumerate(matched):
            left, right = left_cells[position], right_cells[position]
            a_ref = get_column_letter(a_idx + 1) + str(row_a)
            b_ref = get_column_letter(b_idx + 1) + str(row_b)
            if changed[position]:
                findings.add(
                    "Value", _is_different_value(left.value, right.value), col_name,
                    a_ref, _display_value(left.value), b_ref, _display_value(right.value), note,
                )
            left_type, right_type = _cell_type(left), _cell_type(right)
            if left_type != right_type:
                findings.add(
                    "Data type", f"{left_type} vs {right_type}", col_name,
                    a_ref, f"{left_type}: {_display_value(left.value)}",
                    b_ref, f"{right_type}: {_display_value(right.value)}",
                    "Displayed values may match while cell types differ",
                )
            _compare_cell_styles(
                left, right, cache_a, cache_b, changed_style_cache,
                col_name, a_ref, b_ref, "Data cell formatting", findings,
            )


def _compare_data_batch(
    batch: list[tuple[tuple[Any, ...] | None, tuple[Any, ...] | None]],
    pairs: list[tuple[int, int, str]],
    header_row: int,
    offset: int,
    cache_a: dict[int, dict[str, Any]],
    cache_b: dict[int, dict[str, Any]],
    changed_style_cache: dict[tuple[int, int], list[tuple[str, Any, Any]]],
    findings: Findings,
    first_header: str,
) -> tuple[int, int]:
    counts_a = sum(a_row is not None for a_row, _ in batch)
    counts_b = sum(b_row is not None for _, b_row in batch)
    matched = [
        (a_row, b_row, header_row + 1 + offset + i, header_row + 1 + offset + i,
         f"Excel row {header_row + 1 + offset + i}")
        for i, (a_row, b_row) in enumerate(batch) if a_row and b_row
    ]

    for index, (a_row, b_row) in enumerate(batch):
        row_no = header_row + 1 + offset + index
        if a_row is None and b_row is not None:
            head = _cell_at(b_row, 0).value
            findings.add("Row", "Additional data row", first_header, "—", "No row",
                         f"A{row_no}", _display_value(head),
                         f"File B row {row_no} has no counterpart")
        elif b_row is None and a_row is not None:
            head = _cell_at(a_row, 0).value
            findings.add("Row", "Missing data row", first_header,
                         f"A{row_no}", _display_value(head), "—", "No row",
                         f"File A row {row_no} has no counterpart in File B")

    _compare_matched_rows(
        matched, pairs, cache_a, cache_b, changed_style_cache, findings,
    )
    return counts_a, counts_b


def _compare_keyed_row_layout(
    layout_a: SheetLayout,
    layout_b: SheetLayout,
    row_a: int,
    row_b: int,
    key_label: str,
    findings: Findings,
) -> None:
    height_a, height_b = layout_a.row_height(row_a), layout_b.row_height(row_b)
    if not math.isclose(height_a, height_b):
        findings.add(
            "Formatting", "Row height", key_label,
            f"Row {row_a}", f"{height_a:g} pt",
            f"Row {row_b}", f"{height_b:g} pt", "Rows aligned by primary key",
        )
    hidden_a, hidden_b = row_a in layout_a.hidden_rows, row_b in layout_b.hidden_rows
    if hidden_a != hidden_b:
        findings.add(
            "Formatting", "Hidden row", key_label,
            f"Row {row_a}", str(hidden_a), f"Row {row_b}", str(hidden_b),
            "Rows aligned by primary key",
        )


def _compare_keyed_data(
    ws_a: Any,
    ws_b: Any,
    header_row: int,
    key_positions: list[tuple[int, int]],
    headers_a: list[str],
    headers_b: list[str],
    pairs: list[tuple[int, int, str]],
    cache_a: dict[int, dict[str, Any]],
    cache_b: dict[int, dict[str, Any]],
    changed_style_cache: dict[tuple[int, int], list[tuple[str, Any, Any]]],
    layout_a: SheetLayout,
    layout_b: SheetLayout,
    findings: Findings,
    batch_size: int,
    temp_dir: Path | None,
) -> tuple[int, int]:
    """Use a disk-backed SQLite key index to avoid retaining File B in memory."""
    positions_a = [a for a, _ in key_positions]
    positions_b = [b for _, b in key_positions]
    count_a = count_b = 0
    with tempfile.TemporaryDirectory(
        prefix="excel_compare_keys_", dir=str(temp_dir) if temp_dir else None,
    ) as workspace:
        connection = sqlite3.connect(str(Path(workspace) / "key_index.sqlite3"))
        try:
            connection.execute("PRAGMA journal_mode=OFF")
            connection.execute("PRAGMA synchronous=OFF")
            connection.execute("PRAGMA cache_size=-16384")
            connection.execute(
                "CREATE TABLE b_rows (k TEXT PRIMARY KEY, row_no INTEGER NOT NULL, "
                "payload BLOB NOT NULL, matched INTEGER NOT NULL DEFAULT 0)"
            )
            connection.execute(
                "CREATE TABLE a_seen (k TEXT PRIMARY KEY, row_no INTEGER NOT NULL)"
            )
            connection.execute("BEGIN")
            for row_no, row in enumerate(
                ws_b.iter_rows(min_row=header_row + 1, values_only=False),
                start=header_row + 1,
            ):
                key = _key_string(row, positions_b, "B", row_no)
                for cell in row:
                    # Style dictionaries are cached once per workbook style ID.
                    if _style_id(cell) not in cache_b:
                        _style_values(cell, cache_b)
                stored = tuple(
                    StoredCell(cell.value, cell.data_type, _style_id(cell)) for cell in row
                )
                try:
                    connection.execute(
                        "INSERT INTO b_rows (k, row_no, payload) VALUES (?, ?, ?)",
                        (key, row_no, sqlite3.Binary(pickle.dumps(stored, protocol=5))),
                    )
                except sqlite3.IntegrityError as exc:
                    previous = connection.execute(
                        "SELECT row_no FROM b_rows WHERE k=?", (key,),
                    ).fetchone()[0]
                    raise ValueError(
                        f"File B has duplicate primary key {_key_label(row, positions_b, headers_b)!r} "
                        f"at rows {previous} and {row_no}."
                    ) from exc
                count_b += 1
            connection.commit()

            matched_batch: list[tuple[tuple[Any, ...], tuple[Any, ...], int, int, str]] = []
            connection.execute("BEGIN")
            for row_no, row in enumerate(
                ws_a.iter_rows(min_row=header_row + 1, values_only=False),
                start=header_row + 1,
            ):
                key = _key_string(row, positions_a, "A", row_no)
                key_label = _key_label(row, positions_a, headers_a)
                try:
                    connection.execute(
                        "INSERT INTO a_seen (k, row_no) VALUES (?, ?)", (key, row_no),
                    )
                except sqlite3.IntegrityError as exc:
                    previous = connection.execute(
                        "SELECT row_no FROM a_seen WHERE k=?", (key,),
                    ).fetchone()[0]
                    raise ValueError(
                        f"File A has duplicate primary key {key_label!r} "
                        f"at rows {previous} and {row_no}."
                    ) from exc
                matching = connection.execute(
                    "SELECT row_no, payload FROM b_rows WHERE k=?", (key,),
                ).fetchone()
                if matching is None:
                    findings.add(
                        "Row", "Missing data row", key_label,
                        f"{get_column_letter(positions_a[0] + 1)}{row_no}", key_label,
                        "—", "No matching key in File B", "File A primary key absent in File B",
                    )
                else:
                    b_row_no, payload = matching
                    b_row = pickle.loads(payload)
                    connection.execute("UPDATE b_rows SET matched=1 WHERE k=?", (key,))
                    matched_batch.append(
                        (row, b_row, row_no, b_row_no,
                         f"Primary key {key_label} (A row {row_no}, B row {b_row_no})")
                    )
                    _compare_keyed_row_layout(
                        layout_a, layout_b, row_no, b_row_no, key_label, findings,
                    )
                count_a += 1
                if len(matched_batch) >= batch_size:
                    _compare_matched_rows(
                        matched_batch, pairs, cache_a, cache_b, changed_style_cache, findings,
                    )
                    matched_batch.clear()
            if matched_batch:
                _compare_matched_rows(
                    matched_batch, pairs, cache_a, cache_b, changed_style_cache, findings,
                )
            connection.commit()

            for (b_row_no, payload) in connection.execute(
                "SELECT row_no, payload FROM b_rows WHERE matched=0 ORDER BY row_no"
            ):
                b_row = pickle.loads(payload)
                key_label = _key_label(b_row, positions_b, headers_b)
                findings.add(
                    "Row", "Additional data row", key_label, "—", "No matching key in File A",
                    f"{get_column_letter(positions_b[0] + 1)}{b_row_no}", key_label,
                    "File B primary key absent in File A",
                )
        finally:
            connection.close()
    return count_a, count_b


def compare_workbooks(
    file_a: Path,
    file_b: Path,
    sheet_a: str | None = None,
    sheet_b: str | None = None,
    header_row: int = 1,
    batch_size: int = 256,
    max_details: int = DEFAULT_MAX_DETAILS,
    key_a: Sequence[str] | str | None = None,
    key_b: Sequence[str] | str | None = None,
    temp_dir: Path | None = None,
) -> Comparison:
    if header_row < 1 or batch_size < 1 or max_details < 1:
        raise ValueError("Header row, batch size, and maximum detail rows must be positive.")
    if max_details > MAX_EXCEL_ROWS - 36:
        raise ValueError(f"Maximum detail rows per sheet cannot exceed {MAX_EXCEL_ROWS - 36:,}.")
    findings = Findings(max_details=max_details)
    wb_a = load_workbook(file_a, read_only=True, data_only=False, keep_links=False)
    wb_b = load_workbook(file_b, read_only=True, data_only=False, keep_links=False)
    try:
        name_a = sheet_a or wb_a.sheetnames[0]
        name_b = sheet_b or wb_b.sheetnames[0]
        if name_a not in wb_a or name_b not in wb_b:
            raise ValueError(
                f"Available worksheets: A={wb_a.sheetnames!r}; B={wb_b.sheetnames!r}."
            )
        ws_a, ws_b = wb_a[name_a], wb_b[name_b]
        if len(wb_a.sheetnames) != len(wb_b.sheetnames):
            findings.add("Workbook", "Worksheet count", "Workbook",
                         "Workbook A", len(wb_a.sheetnames),
                         "Workbook B", len(wb_b.sheetnames),
                         "The two workbooks contain different numbers of worksheets")
        if name_a != name_b:
            findings.add("Workbook", "Worksheet name", "Worksheet",
                         "Worksheet A", name_a, "Worksheet B", name_b,
                         "Selected worksheet names differ")

        head_a = tuple(next(ws_a.iter_rows(min_row=header_row, max_row=header_row), ()))
        head_b = tuple(next(ws_b.iter_rows(min_row=header_row, max_row=header_row), ()))
        headers_a, headers_b = _column_names(head_a), _column_names(head_b)
        key_positions = _resolve_key_columns(headers_a, headers_b, key_a, key_b)
        cache_a: dict[int, dict[str, Any]] = {}
        cache_b: dict[int, dict[str, Any]] = {}
        changed_style_cache: dict[tuple[int, int], list[tuple[str, Any, Any]]] = {}
        (
            pairs, missing, extra, case_changes, order_changes, order_note_a, order_note_b
        ) = _compare_headers(
            head_a, head_b, header_row, cache_a, cache_b, changed_style_cache, findings,
            key_positions,
        )
        # XML metadata is read separately to avoid loading the entire workbook
        # into memory just to obtain row heights and column widths.
        layout_a = read_sheet_layout(file_a, name_a)
        layout_b = read_sheet_layout(file_b, name_b)
        if key_positions:
            # Header layout and column layout are positional; data-row height
            # and visibility are compared at the corresponding key-matched rows.
            _compare_layout(layout_a, layout_b, header_row, 0, pairs, findings)
            count_a, count_b = _compare_keyed_data(
                ws_a, ws_b, header_row, key_positions, headers_a, headers_b, pairs,
                cache_a, cache_b, changed_style_cache, layout_a, layout_b, findings,
                batch_size, temp_dir,
            )
        else:
            source_rows = zip_longest(
                ws_a.iter_rows(min_row=header_row + 1, values_only=False),
                ws_b.iter_rows(min_row=header_row + 1, values_only=False),
                fillvalue=None,
            )
            count_a = count_b = offset = 0
            first_header = headers_a[0] if headers_a else "First column"
            for batch in _batches(source_rows, batch_size):
                ca, cb = _compare_data_batch(
                    batch, pairs, header_row, offset, cache_a, cache_b,
                    changed_style_cache, findings, first_header,
                )
                count_a += ca
                count_b += cb
                offset += len(batch)
            _compare_layout(layout_a, layout_b, header_row, min(count_a, count_b), pairs, findings)
        return Comparison(
            file_a, file_b, name_a, name_b,
            len(wb_a.sheetnames), len(wb_b.sheetnames),
            sum(bool(value) for value in headers_a),
            sum(bool(value) for value in headers_b),
            count_a, count_b, missing, extra, case_changes, order_changes, findings,
            order_note_a, order_note_b,
            tuple(headers_a[a] for a, _ in key_positions),
            tuple(headers_b[b] for _, b in key_positions),
        )
    finally:
        wb_a.close()
        wb_b.close()


def _report_cell_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, str):
        # Excel would otherwise interpret a leading '=' as a formula.
        return "'" + value if value.startswith(("=", "+", "-", "@")) else value
    return value


def _fill(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _style_report(
    workbook: Workbook,
    summary: Any,
    data_differences: Any,
    last_data_row: int,
    last_other_row: int,
) -> None:
    for ws in (summary, data_differences):
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.sheet_format.defaultRowHeight = 24
        ws["A2"].font = Font(name="Arial", size=15, bold=True, color=COLORS["navy"])
        ws["A3"].font = Font(name="Arial", size=10, italic=True, color=COLORS["muted"])
    summary.sheet_properties.tabColor = COLORS["navy"]
    data_differences.freeze_panes = "A5"

    for col, width in {"A": 29, "B": 23, "C": 26, "D": 22, "E": 34, "F": 48}.items():
        summary.column_dimensions[col].width = width
    for col, width in {
        "A": 7, "B": 16, "C": 23, "D": 25, "E": 17,
        "F": 25, "G": 17, "H": 25, "I": 48,
    }.items():
        data_differences.column_dimensions[col].width = width

    for ws, header_row, max_col in (
        (summary, 12, 5), (summary, 36, 6), (data_differences, 4, 9),
    ):
        ws.row_dimensions[header_row].height = 30
        for row in ws.iter_rows(min_row=header_row, max_row=header_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.fill = _fill(COLORS["navy"])
                cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    summary["A11"].font = Font(name="Arial", size=11, bold=True, color=COLORS["navy"])
    summary["A27"].font = Font(name="Arial", size=10, bold=True, color=COLORS["navy"])
    summary["A34"].font = Font(name="Arial", size=11, bold=True, color=COLORS["navy"])
    summary["A35"].font = Font(name="Arial", size=10, italic=True, color=COLORS["muted"])
    for cell_address in ("B8", "D8"):
        if summary[cell_address].value == "NO":
            summary[cell_address].fill = _fill(COLORS["pale_red"])
    if "NOT IDENTICAL" in summary["E8"].value:
        summary["E8"].fill = _fill(COLORS["pale_amber"])
    for cell in summary[8][:5]:
        cell.font = Font(name="Arial", size=10, bold=True, color=COLORS["navy"])
    for row in summary.iter_rows(min_row=13, max_row=23, min_col=1, max_col=5):
        for cell in row:
            cell.alignment = Alignment(vertical="center")
        if row[3].value == "DIFFERENT":
            row[3].fill = _fill(COLORS["pale_amber"])
    for cell in summary[24][:5]:
        cell.fill = _fill(COLORS["pale_red"])
        cell.font = Font(name="Arial", size=10, bold=True, color=COLORS["navy"])
    for row in summary.iter_rows(min_row=13, max_row=24, min_col=4, max_col=5):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    if last_data_row >= 5:
        for row in data_differences.iter_rows(
            min_row=5, max_row=last_data_row, min_col=1, max_col=9,
        ):
            for cell in row:
                cell.font = Font(name="Arial", size=10, color="263238")
                cell.alignment = Alignment(vertical="center", wrap_text=cell.column == 9)
            for index in (1, 5, 7):
                row[index - 1].alignment = Alignment(horizontal="center", vertical="center")
            color = {
                "Value": COLORS["pale_red"],
                "Row": COLORS["pale_amber"],
                "Data type": COLORS["pale_amber"],
            }.get(row[1].value)
            if color:
                row[1].fill = _fill(color)
            data_differences.row_dimensions[row[0].row].height = 28
        data_differences.auto_filter.ref = f"A4:I{last_data_row}"

    if last_other_row >= 37:
        for row in summary.iter_rows(
            min_row=37, max_row=last_other_row, min_col=1, max_col=6,
        ):
            for cell in row:
                cell.font = Font(name="Arial", size=10, color="263238")
                cell.alignment = Alignment(vertical="center", wrap_text=cell.column >= 4)
            color = {
                "Header": COLORS["pale_blue"],
                "Formatting": COLORS["gray"],
                "Workbook": COLORS["pale_blue"],
            }.get(row[0].value)
            if color:
                row[0].fill = _fill(color)
            summary.row_dimensions[row[0].row].height = 36
        summary.auto_filter.ref = f"A36:F{last_other_row}"


def _source_cell_text(cell: str, value: Any) -> str:
    rendered = str(_display_value(value))
    return f"{cell}: {rendered}" if cell and cell != "—" else rendered


def write_report(result: Comparison, output_path: Path) -> None:
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    data_differences = wb.create_sheet("Data Differences")
    summary["A2"] = "Excel Comparison Report"
    summary["A3"] = "Comparison results for the selected worksheets"
    summary["A5"], summary["B5"], summary["C5"], summary["D5"] = (
        "File A", result.file_a.name, "File B", result.file_b.name
    )
    summary["A6"], summary["B6"], summary["C6"], summary["D6"] = (
        "Worksheet A", result.sheet_a, "Worksheet B", result.sheet_b
    )
    summary["A8"], summary["B8"], summary["C8"], summary["D8"], summary["E8"] = (
        "Data identical?", "YES" if result.data_identical else "NO",
        "Formatting identical?", "YES" if result.formatting_identical else "NO",
        "Overall: " + ("IDENTICAL (CHECKED SCOPE)" if result.identical else "NOT IDENTICAL"),
    )
    summary["A11"] = "Comparison checks"
    summary.append([])  # Header and summary values below use fixed locations.
    headers = ["Check", "File A", "File B", "Result", "Difference count"]
    for col, value in enumerate(headers, 1):
        summary.cell(12, col, value)

    findings = result.findings
    first_case_finding = next(
        (row for row in findings.other_rows if row.issue == "Case change"), None,
    )
    mapped_header_names = [
        (name_a, name_b)
        for name_a, name_b in zip(result.key_columns_a, result.key_columns_b, strict=True)
        if name_a.casefold() != name_b.casefold()
    ]
    unmatched_a = [*result.missing_headers, *(a for a, _ in mapped_header_names)]
    unmatched_b = [*result.additional_headers, *(b for _, b in mapped_header_names)]
    line_items = [
        ("Worksheet count", result.sheet_count_a, result.sheet_count_b,
         result.sheet_count_a == result.sheet_count_b,
         findings.issue_counts["Worksheet count"]),
        ("Worksheet name", result.sheet_a, result.sheet_b,
         result.sheet_a == result.sheet_b,
         findings.issue_counts["Worksheet name"]),
        ("Header count", result.header_count_a, result.header_count_b,
         result.header_count_a == result.header_count_b,
         abs(result.header_count_a - result.header_count_b)),
        ("Unmatched / mapped headers" if mapped_header_names else "Unmatched headers",
         ", ".join(unmatched_a) or "None", ", ".join(unmatched_b) or "None",
         not (unmatched_a or unmatched_b),
         findings.issue_counts["Missing header"] + findings.issue_counts["Additional header"]
         + findings.issue_counts["Mapped key header"]),
        ("Header case",
         first_case_finding.a_value if first_case_finding else f"{result.case_changes} case changes",
         first_case_finding.b_value if first_case_finding else f"{result.case_changes} case changes",
         not result.case_changes, result.case_changes),
        ("Column order", result.order_notes_a or "Same",
         result.order_notes_b or "Same", not result.order_changes, result.order_changes),
        ("Data row count", result.row_count_a, result.row_count_b,
         result.row_count_a == result.row_count_b,
         "See additional rows" if result.row_count_a != result.row_count_b else 0),
        ("Changed cell values", "Compared matched rows", "Compared matched rows",
         findings.counts["Value"] == 0, findings.counts["Value"]),
        ("Different cell types", "Types checked", "Types checked",
         findings.counts["Data type"] == 0, findings.counts["Data type"]),
        ("Formatting properties", "Cell and layout styles", "Cell and layout styles",
         findings.counts["Formatting"] == 0, findings.counts["Formatting"]),
        ("Additional or missing rows", result.row_count_a, result.row_count_b,
         findings.counts["Row"] == 0, findings.counts["Row"]),
        ("Total unique findings", "", "", result.identical, findings.total),
    ]
    for row_index, (label, a_value, b_value, passed, count) in enumerate(line_items, 13):
        for col_index, value in enumerate(
            [label, a_value, b_value,
             ("IDENTICAL" if passed else "NOT IDENTICAL") if row_index == 24
             else ("MATCH" if passed else "DIFFERENT"),
             count], 1
        ):
            summary.cell(row_index, col_index, _report_cell_value(value))

    summary["A27"] = "Method"
    if result.key_columns_a:
        summary["A28"] = "Rows matched by a unique primary key; row order does not affect matching."
        summary["A29"] = (
            f"Key A: {', '.join(result.key_columns_a)}; "
            f"Key B: {', '.join(result.key_columns_b)}. Header lookup ignores case."
        )
        summary["A31"] = "Differently named key headers are reported, but their data columns are paired."
    else:
        summary["A28"] = "Rows matched by Excel row number; columns matched by name ignoring case."
        summary["A29"] = "Unmatched headers are not guessed. Formatting changes count per property."
    summary["A30"] = "Data values are case-sensitive. Spaces, data types and formulas are compared."
    if findings.truncated:
        summary["A32"] = (
            f"Showing {len(findings.data_rows):,}/{findings.data_total:,} data findings and "
            f"{len(findings.other_rows):,}/{findings.other_total:,} other findings. "
            "Summary totals include all findings."
        )
        summary["A32"].font = Font(name="Arial", size=10, bold=True, color="9C2D2D")

    summary["A34"] = "Other differences"
    summary["A35"] = "Header, workbook and formatting differences."
    for column, name in enumerate(
        ["Category", "Issue", "Column / Field", "File A (cell: value)",
         "File B (cell: value)", "Notes"], 1,
    ):
        summary.cell(36, column, name)
    for index, finding in enumerate(findings.other_rows, 37):
        values = [
            finding.category, finding.issue, finding.column,
            _source_cell_text(finding.a_cell, finding.a_value),
            _source_cell_text(finding.b_cell, finding.b_value),
            finding.note,
        ]
        for column, value in enumerate(values, 1):
            summary.cell(index, column, _report_cell_value(value))
    if not findings.other_rows:
        summary["A37"] = "No other differences found"
        summary["A37"].font = Font(name="Arial", size=10, italic=True, color=COLORS["muted"])

    data_differences["A2"] = "Data Differences"
    data_differences["A3"] = "Changed values, data types, and missing or additional records."
    for column, name in enumerate(
        ["No.", "Category", "Issue", "Column / Field", "File A Cell",
         "File A", "File B Cell", "File B", "Notes"], 1
    ):
        data_differences.cell(4, column, name)
    for index, finding in enumerate(findings.data_rows, 1):
        values = [
            index, finding.category, finding.issue, finding.column,
            finding.a_cell, finding.a_value, finding.b_cell, finding.b_value, finding.note,
        ]
        for column, value in enumerate(values, 1):
            data_differences.cell(index + 4, column, _report_cell_value(value))
    if not findings.data_rows:
        data_differences["A5"] = "No data differences found"
        data_differences["A5"].font = Font(
            name="Arial", size=10, italic=True, color=COLORS["muted"],
        )

    _style_report(
        wb, summary, data_differences,
        4 + len(findings.data_rows),
        36 + len(findings.other_rows),
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare two Excel worksheets by row position or optional primary key.",
    )
    parser.add_argument("file_a", type=Path, help="Path to the first .xlsx workbook")
    parser.add_argument("file_b", type=Path, help="Path to the second .xlsx workbook")
    parser.add_argument("-o", "--output", type=Path, default=Path("Comparison_Report.xlsx"),
                        help="Destination for the comparison report")
    parser.add_argument("--sheet-a", help="Worksheet name in File A (default: first sheet)")
    parser.add_argument("--sheet-b", help="Worksheet name in File B (default: first sheet)")
    parser.add_argument("--key-a", help="File A primary key, e.g. 'a1' or 'a2,a3'")
    parser.add_argument("--key-b", help="File B primary key, e.g. 'b1' or 'b2,b3'")
    parser.add_argument("--temp-dir", type=Path,
                        help="Folder for the disk-backed key index (default: system temporary folder)")
    parser.add_argument("--header-row", type=int, default=1, help="Excel row with the column headers")
    parser.add_argument("--batch-size", type=int, default=256, help="Rows processed per Polars batch")
    parser.add_argument("--max-details", type=int, default=DEFAULT_MAX_DETAILS,
                        help="Maximum number of detailed findings written into Excel")
    parser.add_argument("--overwrite", action="store_true", help="Allow replacing an existing report")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    files = (args.file_a.expanduser().resolve(), args.file_b.expanduser().resolve())
    output = args.output.expanduser().resolve()
    try:
        if files[0] == files[1]:
            raise ValueError("Provide two different source file paths.")
        for path in files:
            if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
                raise ValueError(f"Input must be an existing .xlsx or .xlsm file: {path}")
        if output in files:
            raise ValueError("The report path must differ from both input file paths.")
        if output.suffix.lower() != ".xlsx":
            raise ValueError("The output report must have the .xlsx extension.")
        if output.exists() and not args.overwrite:
            raise FileExistsError(f"Report already exists: {output}; specify --overwrite to replace it.")
        result = compare_workbooks(
            files[0], files[1], args.sheet_a, args.sheet_b,
            args.header_row, args.batch_size, args.max_details,
            args.key_a, args.key_b, args.temp_dir,
        )
        write_report(result, output)
    except (OSError, ValueError, KeyError, ET.ParseError) as exc:
        print(f"Comparison failed: {exc}", file=sys.stderr)
        return 2
    print(f"Report: {output}")
    print(f"Data identical: {result.data_identical}")
    print(f"Formatting identical: {result.formatting_identical}")
    print(f"Identical in checked scope: {result.identical}")
    print(f"Total findings: {result.findings.total:,}")
    if result.findings.truncated:
        print(
            f"Detail rows capped at {len(result.findings.data_rows):,} data and "
            f"{len(result.findings.other_rows):,} other findings."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
