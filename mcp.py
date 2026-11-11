"""Publish the workbook tools over MCP stdio.

Install: python -m pip install -r requirements.txt
Run:     python /absolute/path/to/mcp.py

Configure an MCP host to launch that command with the same Python environment.
Use absolute workbook paths when the host's working directory is unspecified.
Stdout is reserved for the MCP protocol; SDK diagnostics go to stderr.
"""

import sys
from pathlib import Path

from tools import (
    close_cached_workbooks, get_cell_formulas, get_cell_values_csv, get_named_regions,
    get_sheet_names, get_value_area,
)


def main() -> None:
    # The requested filename mcp.py shadows the installed mcp package. Exclude
    # this directory during the SDK import, including a cwd/PYTHONPATH duplicate.
    script_dir = Path(__file__).resolve().parent
    original_path = sys.path[:]
    try:
        sys.path[:] = [p for p in sys.path if Path(p or ".").resolve() != script_dir]
        from mcp.server.fastmcp import FastMCP
        from mcp.types import ToolAnnotations
    finally:
        sys.path[:] = original_path

    server = FastMCP(
        "Excel Explorer",
        instructions=(
            "Explore local Excel workbooks using finite A1 ranges and exact sheet "
            "names. Prefer absolute file paths. CSV formula values are saved "
            "caches and may be stale or absent; formulas are never calculated. "
            "Value areas include formula cells even without cached results. "
            "Opened workbooks are cached for this process. Files must remain "
            "unchanged during use; restart the server to read updated files."
        ),
    )
    for function in (
        get_value_area, get_cell_values_csv, get_cell_formulas, get_named_regions,
        get_sheet_names,
    ):
        server.tool(
            annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False),
        )(function)
    try:
        server.run(transport="stdio")
    finally:
        close_cached_workbooks()


if __name__ == "__main__":
    main()
