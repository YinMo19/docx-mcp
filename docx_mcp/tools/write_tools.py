from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from docx_mcp.services.write_ops import (
    add_heading,
    add_paragraph,
    add_table,
    search_and_replace,
)
from docx_mcp.tools.common import handle_tool_error
from docx_mcp.tools.compat import parse_bool, parse_int, parse_matrix, parse_optional_float


def register_write_tools(mcp: FastMCP) -> None:
    @mcp.tool(
        name="search_and_replace",
        description="Replace text occurrences in paragraphs and table cells, then save DOCX.",
    )
    def search_and_replace_tool(
        filename: str,
        find_text: str,
        replace_text: str,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            return {
                "ok": True,
                "result": search_and_replace(
                    filename=filename,
                    find_text=find_text,
                    replace_text=replace_text,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="add_paragraph",
        description="Append a paragraph to document with optional style and font formatting.",
    )
    def add_paragraph_tool(
        filename: str,
        text: str,
        style: str | None = None,
        font_name: str | None = None,
        font_size: float | int | str | None = None,
        bold: bool | str = False,
        italic: bool | str = False,
        color: str | None = None,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_font_size = parse_optional_float(font_size, field="font_size")
            normalized_bold = parse_bool(bold, field="bold")
            normalized_italic = parse_bool(italic, field="italic")
            return {
                "ok": True,
                "result": add_paragraph(
                    filename=filename,
                    text=text,
                    style=style,
                    font_name=font_name,
                    font_size=normalized_font_size,
                    bold=normalized_bold,
                    italic=normalized_italic,
                    color=color,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="add_heading",
        description="Append a heading paragraph with level 0-9 and optional font overrides.",
    )
    def add_heading_tool(
        filename: str,
        text: str,
        level: int | str = 1,
        font_name: str | None = None,
        font_size: float | int | str | None = None,
        bold: bool | str = False,
        italic: bool | str = False,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_level = parse_int(level, field="level")
            normalized_font_size = parse_optional_float(font_size, field="font_size")
            normalized_bold = parse_bool(bold, field="bold")
            normalized_italic = parse_bool(italic, field="italic")
            return {
                "ok": True,
                "result": add_heading(
                    filename=filename,
                    text=text,
                    level=normalized_level,
                    font_name=font_name,
                    font_size=normalized_font_size,
                    bold=normalized_bold,
                    italic=normalized_italic,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="add_table",
        description="Append a table with given rows/cols and optional initial data matrix.",
    )
    def add_table_tool(
        filename: str,
        rows: int | str,
        cols: int | str,
        data: list[list[str]] | str | None = None,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_rows = parse_int(rows, field="rows")
            normalized_cols = parse_int(cols, field="cols")
            normalized_data = parse_matrix(data, field="data")
            return {
                "ok": True,
                "result": add_table(
                    filename=filename,
                    rows=normalized_rows,
                    cols=normalized_cols,
                    data=normalized_data,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)
