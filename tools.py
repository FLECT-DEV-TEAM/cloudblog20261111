"""Read-only exploration of Excel .xlsx/.xlsm workbooks (Python 3.10+).

Paths are absolute or relative to the process working directory. Ranges must be
one finite A1 rectangle (e.g. A1, A1:C3, or $A$1:$C$3), without a sheet prefix.
Values are underlying Python values, not Excel's formatted display strings.
Formula values come from saved caches; no formulas or VBA are executed.
Opened workbooks are reused for the process lifetime, separately for formula and
cached-value reads. Files are assumed unchanged while cached; no change detection
or parsed-sheet cache is provided. Restart the process to open updated files.
"""

import atexit
import csv
import io
import json
import re
from contextlib import contextmanager
from pathlib import Path
from threading import Lock
from typing import Annotated
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula
from pydantic import Field

_CellReadRange = Annotated[str, Field(description=(
    "A cell or finite A1 rectangle without a sheet prefix, at most 10 rows "
    "and 10 columns (e.g. A1:J10 or $B$2:$K$11). Larger ranges raise "
    "ValueError; split them into smaller requests."
))]

__all__ = [
    "get_value_area", "get_cell_values_csv", "get_cell_formulas", "get_named_regions",
    "get_sheet_names",
    "close_cached_workbooks",
]

_workbooks: dict[tuple[Path, bool], Workbook] = {}
# Protect opening, use, and closing so concurrent calls cannot duplicate loads
# or close a workbook while another call is reading it.
_workbook_lock = Lock()
_WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def close_cached_workbooks() -> None:
    """Close all cached workbooks. Used at shutdown and by library callers.

    This is a lifecycle helper, not an MCP tool. Subsequent reads reopen files.
    """
    with _workbook_lock:
        while _workbooks:
            _, workbook = _workbooks.popitem()
            workbook.close()


atexit.register(close_cached_workbooks)

_A1_RANGE = re.compile(
    r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?"
)


def _bounds(excel_range: str) -> dict[str, int]:
    if not isinstance(excel_range, str) or not _A1_RANGE.fullmatch(excel_range):
        raise ValueError("excel_range must be a cell or finite A1 range, e.g. A1:C3")
    min_col, min_row, max_col, max_row = range_boundaries(excel_range.upper())
    if not (1 <= min_col <= max_col <= 16384 and 1 <= min_row <= max_row <= 1048576):
        raise ValueError("excel_range is reversed or exceeds Excel worksheet limits")
    return dict(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)


def _cell_read_bounds(excel_range: str) -> dict[str, int]:
    bounds = _bounds(excel_range)
    if (bounds["max_row"] - bounds["min_row"] + 1 > 10
            or bounds["max_col"] - bounds["min_col"] + 1 > 10):
        raise ValueError("excel_range must contain at most 10 rows and 10 columns")
    return bounds


@contextmanager
def _workbook(file_name: str, *, data_only: bool):
    path = _workbook_path(file_name)
    with _workbook_lock:
        key = (path, data_only)
        workbook = _workbooks.get(key)
        if workbook is None:
            # Cache only successful loads. Workbooks remain strictly read-only.
            workbook = load_workbook(
                path, read_only=True, data_only=data_only, keep_links=False,
            )
            _workbooks[key] = workbook
        yield workbook


def _workbook_path(file_name: str) -> Path:
    path = Path(file_name).expanduser().resolve()
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        raise ValueError("Supported workbook formats: .xlsx, .xlsm, .xltx, .xltm")
    return path


@contextmanager
def _worksheet(file_name: str, sheet_name: str, *, data_only: bool):
    with _workbook(file_name, data_only=data_only) as workbook:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet {sheet_name!r} not found. Available sheets: "
                + ", ".join(workbook.sheetnames)
            )
        sheet = workbook[sheet_name]
        if not hasattr(sheet, "iter_rows"):
            raise ValueError(f"{sheet_name!r} is not a cell worksheet")
        # Some producers write incorrect used-range metadata. Each call supplies
        # explicit bounds, so ignore that metadata when streaming the cells.
        sheet.reset_dimensions()
        yield sheet


def get_sheet_names(file_name: str) -> str:
    """Return workbook sheet names, in workbook order, as a JSON array.

    Supports .xlsx/.xlsm and their template variants. Includes hidden sheets.
    Sheet names are returned exactly as stored in the workbook.
    """
    with _workbook(file_name, data_only=False) as workbook:
        return json.dumps(workbook.sheetnames, ensure_ascii=False, indent=2)


def get_named_regions(file_name: str) -> str:
    """Return all workbook-defined names as a JSON array.

    Supports .xlsx/.xlsm and their template variants. Each object has `name`,
    `scope` ("workbook" or the worksheet name), and `refers_to`, exactly as
    stored in the workbook. Names may refer to a cell/range, table reference,
    constant, or formula. Hidden and built-in names (such as print areas and
    print titles) are included and marked with `hidden`.
    """
    regions = []
    path = _workbook_path(file_name)
    with ZipFile(path) as archive:
        workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    namespace = {"main": _WORKBOOK_NS}
    sheets = [
        sheet.get("name")
        for sheet in workbook_xml.findall("main:sheets/main:sheet", namespace)
    ]
    for definition in workbook_xml.findall(
        "main:definedNames/main:definedName", namespace,
    ):
        local_sheet_id = definition.get("localSheetId")
        scope = "workbook"
        if local_sheet_id is not None:
            sheet_index = int(local_sheet_id)
            if not 0 <= sheet_index < len(sheets):
                raise ValueError(
                    f"Defined name {definition.get('name')!r} has invalid "
                    f"localSheetId {local_sheet_id}"
                )
            scope = sheets[sheet_index]
        regions.append({
            "name": definition.get("name"),
            "scope": scope,
            "refers_to": definition.text or "",
            "hidden": definition.get("hidden", "0").lower() in {"1", "true"},
        })
    return json.dumps(regions, ensure_ascii=False, indent=2)


def get_value_area(file_name: str, sheet_name: str, excel_range: str) -> str:
    """Return the smallest rectangle containing non-null cells inside excel_range.

    Supports .xlsx/.xlsm and their template variants. Use a finite A1 range.
    Zero, False, whitespace, errors, and formula cells count as occupied, even
    when a formula has no cached result. Formatting alone does not count.
    Merged cells contribute only their stored anchor value. Return "" if empty,
    "B2" for one occupied cell, or a range such as "B2:B3" otherwise.
    """
    bounds = _bounds(excel_range)
    left, top, right, bottom = 16385, 1048577, 0, 0
    with _worksheet(file_name, sheet_name, data_only=False) as sheet:
        for row in sheet.iter_rows(**bounds):
            for cell in row:
                if cell.value is not None:
                    left = min(left, cell.column)
                    right = max(right, cell.column)
                    top = min(top, cell.row)
                    bottom = max(bottom, cell.row)
    if not right:
        return ""
    first = f"{get_column_letter(left)}{top}"
    last = f"{get_column_letter(right)}{bottom}"
    return first if first == last else f"{first}:{last}"


def get_cell_values_csv(file_name: str, sheet_name: str, excel_range: _CellReadRange) -> str:
    """Return CSV for an A1 rectangle of at most 10 rows and 10 columns.

    excel_range must be a cell or finite A1 rectangle without a sheet prefix,
    e.g. A1:J10. Larger ranges raise ValueError; split them into smaller requests.

    Supports .xlsx/.xlsm and their template variants. No header or addresses are
    added. Empty cells become empty fields. Standard CSV quoting preserves commas,
    quotes, and line breaks. Records end in CRLF, including the final record.
    Dates/times use Python string representations; Excel display formatting is
    not applied. Formula cells return the last saved cached result, which may be
    stale or blank if absent. This tool does not recalculate formulas.
    """
    bounds = _cell_read_bounds(excel_range)
    output = io.StringIO(newline="")
    writer = csv.writer(output)
    written_rows = 0
    with _worksheet(file_name, sheet_name, data_only=True) as sheet:
        for row in sheet.iter_rows(**bounds, values_only=True):
            writer.writerow(row)
            written_rows += 1
    # The streaming reader can omit trailing rows beyond the XML's last row.
    # CSV must retain the exact requested rectangle, even on an empty sheet.
    empty_row = [None] * (bounds["max_col"] - bounds["min_col"] + 1)
    row_count = bounds["max_row"] - bounds["min_row"] + 1
    for _ in range(row_count - written_rows):
        writer.writerow(empty_row)
    return output.getvalue()


def get_cell_formulas(file_name: str, sheet_name: str, excel_range: _CellReadRange) -> str:
    """Return a JSON object mapping cell addresses to formula strings, including '='.

    excel_range must be a cell or finite A1 rectangle without a sheet prefix,
    with at most 10 rows and 10 columns, e.g. A1:J10. Larger ranges raise
    ValueError; split them into smaller requests.

    Supports .xlsx/.xlsm and their template variants.
    Constants are omitted, including literal text that begins with '='. Return
    '{}' when no formulas exist. Shared formulas are expanded by openpyxl.
    Array formulas appear at their anchor cell only. Excel what-if data-table
    definitions have no formula string and raise ValueError rather than silently
    disappearing. Formula expressions are returned without evaluation.
    """
    bounds = _cell_read_bounds(excel_range)
    formulas = {}
    with _worksheet(file_name, sheet_name, data_only=False) as sheet:
        for row in sheet.iter_rows(**bounds):
            for cell in row:
                if cell.data_type != "f":
                    continue
                formula = cell.value
                if isinstance(formula, ArrayFormula):
                    formula = formula.text
                if not isinstance(formula, str):
                    raise ValueError(
                        f"Formula at {cell.coordinate} has no textual expression "
                        "(e.g. an Excel what-if data table)"
                    )
                formulas[cell.coordinate] = formula
    return json.dumps(formulas, ensure_ascii=False, indent=2)
