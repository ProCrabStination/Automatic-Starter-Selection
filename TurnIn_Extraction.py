"""Read the cells listed in TurnIn_Information_to_extract.json from the order turn-in workbook.

Excel opens the turn-in READ-ONLY, recalculates it (so formulas such as =DAYS(E22, TODAY())/7 are current) and
hands the requested cell values straight back to Python. No Excel file is created; the only file written is the
filled JSON "<turn-in name> - turn-in values.json" in the Working folder.
If Excel cannot be used, openpyxl reads the turn-in read-only instead (formula cells then hold the values from the
last time the turn-in was saved, and a warning is logged).

read_cells()          the requested cells -> {(sheet key, row, column): value}
fill_cell_map()       a copy of TurnIn_Information_to_extract.json with every value / table row / status filled in
extract_turnin_info() both of the above + save the filled JSON; this is what Main.py calls
"""
import copy
import datetime
import json
import os
import re
import subprocess
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import from_excel

from GUI import EXCEL_EXTENSIONS, WORKING_DIR, map_nested

APP_DIR = Path(__file__).resolve().parent
INFO_FILE = APP_DIR / "TurnIn_Information_to_extract.json"
VALUES_FILE_SUFFIX = " - turn-in values"   # filled JSON: "<turn-in name> - turn-in values.json"

# Opens the workbook READ-ONLY in a hidden Excel, recalculates, and prints the requested cells as JSON.
# Inputs come through environment variables (XL_SRC = workbook, XL_REQ = [{"key": sheet key, "range": "E10"}]) so
# spaces, quotes, & etc. in paths are safe. Sheets are matched by name ignoring case and extra spaces.
_READ_CELLS_PS = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$requests = $env:XL_REQ | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$cells = @()
$names = @()
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($env:XL_SRC, 0, $true)
    try {
        $xl.CalculateFull()
        $byKey = @{}
        foreach ($ws in $wb.Worksheets) {
            $names += $ws.Name
            $byKey[(($ws.Name.Trim() -split '\s+') -join ' ').ToLower()] = $ws
        }
        foreach ($r in $requests) {
            $ws = $byKey[$r.key]
            if (-not $ws) { continue }
            foreach ($c in $ws.Range($r.range).Cells) {
                $cells += [pscustomobject]@{ key = $r.key; row = $c.Row; col = $c.Column; value = $c.Value2;
                                             format = [string]$c.NumberFormat }
            }
        }
    } finally { $wb.Close($false) }
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
[pscustomobject]@{ sheets = $names; cells = $cells } | ConvertTo-Json -Depth 4 -Compress
"""


# ----------------------------------------------------------------------------------------------- values
def _has_value(value) -> bool:
    """False for empty cells and cells holding only spaces."""
    return value is not None and not (isinstance(value, str) and not value.strip())


def _sheet_key(name) -> str:
    """Sheet names compared ignoring case and extra spaces: "Recip Rating " finds "recip rating"."""
    return " ".join(str(name).split()).lower()


def _as_shown(value, number_format):
    """A number rounded the way its plain number format shows it in Excel: "0" -> 8.2857 becomes 8,
    "0.00" -> 2 decimals, "#,##0.0" -> 1 decimal. Halves round up like Excel (2.5 -> 3), not to even.
    Other formats (General, %, dates, currency, ...) and non-numbers are returned unchanged.
    `value` may be a nested list (same-shape results back)."""
    if isinstance(value, (list, tuple)):
        return map_nested(lambda one: _as_shown(one, number_format), value)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return value
    match = re.fullmatch(r"0+(?:\.(0+))?", str(number_format or "").split(";")[0].replace("#", "").replace(",", ""))
    if not match:
        return value
    places = len(match.group(1) or "")
    rounded = Decimal(str(value)).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    return int(rounded) if places == 0 else float(rounded)


def _json_value(value):
    """A cell value JSON can hold: dates/times -> ISO text, text trimmed, 8.0 -> 8, empty -> None."""
    if not _has_value(value):
        return None
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, (int, float, bool)):
        return value
    return str(value)


def _convert(value, number_format):
    """Raw cell value + its number format -> the value as the turn-in shows it: Excel date serials in a date
    format become dates, plain numbers are rounded to their format. `value` may be a nested list."""
    if isinstance(value, (list, tuple)):
        return map_nested(lambda one: _convert(one, number_format), value)
    if isinstance(value, (int, float)) and not isinstance(value, bool) and is_date_format(number_format or "General"):
        try:
            value = from_excel(value)
        except (ValueError, OverflowError):
            pass
    return _json_value(_as_shown(value, number_format))


# ------------------------------------------------------------------------------------------------ reading
def load_cell_map(path=INFO_FILE) -> dict:
    """Read TurnIn_Information_to_extract.json (the list of sheets / cells / tables to pull from the turn-in).
    `path` may be a nested list of files (same-shape results back)."""
    if isinstance(path, (list, tuple)):
        return map_nested(load_cell_map, path)
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _skip_rule(block: dict):
    """(sheet, cell) of a block's "skip_if_zero" rule, or None. "E2" = own sheet; {"sheet", "cell"} = other sheet."""
    rule = block.get("skip_if_zero")
    if not rule:
        return None
    if isinstance(rule, dict):
        return rule["sheet"], rule["cell"].strip().upper()
    return block["sheet"], str(rule).strip().upper()


def _requests(cell_map: dict) -> list:
    """Every cell / range the cell map needs: [{"key": sheet key, "range": "E10" or "A1:C55"}], no repeats."""
    wanted = []
    for block in cell_map.get("sheets", []):
        key = _sheet_key(block["sheet"])
        wanted += [(key, f["cell"].strip().upper()) for f in block.get("fields", [])]
        if "table" in block:
            wanted.append((key, block["table"]["range"].strip().upper()))
        rule = _skip_rule(block)
        if rule:
            wanted.append((_sheet_key(rule[0]), rule[1]))
    return [{"key": key, "range": rng} for key, rng in dict.fromkeys(wanted)]


def _read_with_excel(source: Path, requests: list):
    """Requested cells through Excel (read-only, recalculated). Returns (sheet names, {(key, row, col): (value,
    number format)}). Raises RuntimeError if Excel fails."""
    env = dict(os.environ, XL_SRC=str(source), XL_REQ=json.dumps(requests))
    proc = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", _READ_CELLS_PS],
                          env=env, capture_output=True, encoding="utf-8", errors="replace", timeout=300,
                          creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if proc.returncode != 0 or not proc.stdout.strip():
        raise RuntimeError((proc.stderr or proc.stdout or "Excel read failed").strip())
    data = json.loads(proc.stdout.strip().lstrip("﻿"))
    cells = data.get("cells") or []
    if isinstance(cells, dict):        # PowerShell writes a single item as an object, not a list
        cells = [cells]
    names = data.get("sheets") or []
    names = [names] if isinstance(names, str) else names
    return names, {(c["key"], int(c["row"]), int(c["col"])): (c.get("value"), c.get("format")) for c in cells}


def _read_with_openpyxl(source: Path, requests: list):
    """Requested cells through openpyxl, file opened read-only ("rb"). Formula cells hold the values saved in the
    file. Same return shape as _read_with_excel."""
    with open(source, "rb") as handle:
        workbook = load_workbook(handle, data_only=True)
    try:
        by_key = {_sheet_key(ws.title): ws for ws in workbook.worksheets}
        cells = {}
        for request in requests:
            ws = by_key.get(request["key"])
            if ws is None:
                continue
            min_col, min_row, max_col, max_row = range_boundaries(request["range"])
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    cells[(request["key"], cell.row, cell.column)] = (cell.value, cell.number_format)
        return list(workbook.sheetnames), cells
    finally:
        workbook.close()


def read_cells(excel_path, cell_map=None, log=print) -> dict:
    """Read every cell the cell map asks for from the turn-in, without changing or copying the turn-in.
    Returns {"source": Path, "sheets": [sheet names], "cells": {(sheet key, row, col): value as shown}}.
    `excel_path` may be a nested list of workbooks (same-shape results back)."""
    if isinstance(excel_path, (list, tuple)):
        return map_nested(lambda one: read_cells(one, cell_map, log), excel_path)
    source = Path(excel_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Turn-in workbook not found: {source}")
    if source.suffix.lower() not in EXCEL_EXTENSIONS:
        raise ValueError(f"Turn-in must be an Excel workbook ({', '.join(EXCEL_EXTENSIONS)}): {source}")
    requests = _requests(load_cell_map() if cell_map is None else cell_map)

    log(f"Opening the turn-in read-only in Excel and recalculating: {source}")
    try:
        names, raw = _read_with_excel(source, requests)
    except Exception as exc:  # noqa: BLE001 - fall back to the saved values, but say so
        if source.suffix.lower() == ".xls":
            raise RuntimeError(f"Could not read the .xls turn-in through Excel: {exc}") from exc
        log(f"WARNING: Excel could not be used ({exc}). Reading the saved values instead: "
            "formula cells (e.g. TODAY()) are from the last time the turn-in was saved.")
        names, raw = _read_with_openpyxl(source, requests)
    log(f"Sheets in the turn-in: {', '.join(names)}")
    return {"source": source, "sheets": names,
            "cells": {key: _convert(value, fmt) for key, (value, fmt) in raw.items()}}


# ------------------------------------------------------------------------------------------------ filling
def _is_zero(value) -> bool:
    """True for 0, "0", "0.0" and blank (no quantity entered)."""
    if not _has_value(value):
        return True
    try:
        return float(str(value).strip()) == 0
    except ValueError:
        return False


def _skip_check(block: dict, present: set, lookup: dict):
    """The skip reason for a sheet block, or None to read it. The sheet is skipped when its "skip_if_zero" cell is
    0 or blank. If the sheet holding that cell is missing, the block is read rather than silently skipped."""
    rule = _skip_rule(block)
    if not rule:
        return None
    sheet, cell = rule
    key = _sheet_key(sheet)
    if key not in present or not _is_zero(lookup.get((key, *coordinate_to_tuple(cell)))):
        return None
    return f"skipped: {sheet} {cell} is 0 or blank"


def _read_table(lookup: dict, key: str, table: dict) -> list:
    """Rows of a table block ({"range": "A1:C55", "header_row": true}) as a list of {header: value}.
    Blank headers become the column letter, repeated headers get the column letter added, empty rows are left out.
    Each row also gets "Excel Row" so it can be traced back to the sheet."""
    min_col, min_row, max_col, max_row = range_boundaries(table["range"])
    columns = range(min_col, max_col + 1)
    if table.get("header_row", True):
        headers, used = [], set()
        for c in columns:
            text = lookup.get((key, min_row, c))
            header = str(text) if text is not None else get_column_letter(c)
            if header in used:
                header = f"{header} ({get_column_letter(c)})"
            used.add(header)
            headers.append(header)
        first_data_row = min_row + 1
    else:
        headers, first_data_row = [get_column_letter(c) for c in columns], min_row
    rows = []
    for r in range(first_data_row, max_row + 1):
        values = [lookup.get((key, r, c)) for c in columns]
        if any(v is not None for v in values):
            rows.append({"Excel Row": r, **dict(zip(headers, values))})
    return rows


def fill_cell_map(cells, cell_map=None, log=print) -> dict:
    """A filled-in copy of the cell map:
    - every field "value" = that cell on that sheet (null if the cell is empty)
    - every table "rows"  = the table's rows as {header: value}
    - every sheet "status" = "read", "sheet not found" or "skipped: ..." (for "skip_if_zero")
    - "source_workbook" = the turn-in path
    `cells` is the result of read_cells(); it may be a nested list (same-shape results back).
    `cell_map` defaults to TurnIn_Information_to_extract.json; it is not changed."""
    if isinstance(cells, (list, tuple)):
        return map_nested(lambda one: fill_cell_map(one, cell_map, log), cells)
    filled = copy.deepcopy(load_cell_map() if cell_map is None else cell_map)
    filled["source_workbook"] = str(cells["source"])
    lookup = cells["cells"]
    present = {_sheet_key(name) for name in cells["sheets"]}

    for block in filled.get("sheets", []):
        key = _sheet_key(block["sheet"])
        skip = _skip_check(block, present, lookup)
        block["status"] = "sheet not found" if key not in present else (skip or "read")
        log(f"Sheet '{block['sheet']}': {block['status']}")
        seen = {}
        for field in block.get("fields", []):
            cell = field["cell"].strip().upper()
            if block["status"] == "read":
                field["value"] = lookup.get((key, *coordinate_to_tuple(cell)))
                same = f"   (same cell as '{seen[cell]}')" if cell in seen else ""
                log(f"  {cell:>5}  {field['name']}: {field['value']}{same}")
            seen.setdefault(cell, field["name"])
        if "table" in block and block["status"] == "read":
            block["table"]["rows"] = _read_table(lookup, key, block["table"])
            log(f"  {block['table']['range']}  table: {len(block['table']['rows'])} row(s)")
    return filled


def _write_json(data: dict, target: Path) -> Path:
    """Save JSON; if the file is locked, use a time-stamped name instead. Returns the path written."""
    text = json.dumps(data, indent=2, ensure_ascii=False)
    try:
        target.write_text(text, encoding="utf-8")
        return target
    except PermissionError:
        stamped = target.with_name(f"{target.stem} {datetime.datetime.now():%Y%m%d_%H%M%S}{target.suffix}")
        stamped.write_text(text, encoding="utf-8")
        return stamped


def extract_turnin_info(excel_path, log=print, stream=None, output_dir=None, info_file=INFO_FILE) -> dict:
    """Turn-in workbook -> filled copy of TurnIn_Information_to_extract.json, saved as
    "<turn-in name> - turn-in values.json" in the Working folder (or `output_dir`). The turn-in is only read.
    Returns {"values_json": Path, "values": filled JSON dict}.
    `excel_path` may be a nested list of workbooks: each is processed and the results come back in the same shape.
    `log(text)` receives progress lines; `stream` is accepted to match GUI.run_with_progress and is not used."""
    if isinstance(excel_path, (list, tuple)):
        return map_nested(lambda one: extract_turnin_info(one, log, stream, output_dir, info_file), excel_path)
    cell_map = load_cell_map(info_file)
    cells = read_cells(excel_path, cell_map, log)
    filled = fill_cell_map(cells, cell_map, log)

    folder = Path(output_dir) if output_dir else WORKING_DIR
    folder.mkdir(parents=True, exist_ok=True)
    values_json = _write_json(filled, folder / f"{cells['source'].stem}{VALUES_FILE_SUFFIX}.json")
    fields = [f for block in filled.get("sheets", []) if block["status"] == "read" for f in block.get("fields", [])]
    log(f"Found a value for {sum(f['value'] is not None for f in fields)} of {len(fields)} cell(s) on read sheets.")
    log(f"Saved filled JSON: {values_json}")
    return {"values_json": values_json, "values": filled}
