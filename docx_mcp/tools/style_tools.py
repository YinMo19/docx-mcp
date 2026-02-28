from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from docx_mcp.services.style_ops import format_table, set_paragraph_format
from docx_mcp.tools.compat import (
    parse_auto_fit,
    parse_bool,
    parse_int,
    parse_optional_bool,
    parse_optional_float,
    parse_paragraph_indices,
    parse_shading,
)
from docx_mcp.tools.common import handle_tool_error


def register_style_tools(mcp: FastMCP) -> None:
    @mcp.tool(
        name="format_table",
        description=(
            "Format a table by index with border style, header styling, and optional alternating row shading."
        ),
    )
    def format_table_tool(
        filename: str,
        table_index: int | str,
        border_style: str = "single",
        has_header_row: bool | str = True,
        shading: list[str] | str | None = None,
        header_fill_color: str = "D9E2F3",
        header_text_color: str | None = None,
        auto_fit: bool | str | None = None,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_table_index = parse_int(table_index, field="table_index")
            normalized_has_header_row = parse_bool(
                has_header_row, field="has_header_row"
            )
            normalized_shading = parse_shading(shading)
            normalized_auto_fit = parse_auto_fit(auto_fit)
            return {
                "ok": True,
                "result": format_table(
                    filename=filename,
                    table_index=normalized_table_index,
                    border_style=border_style,
                    has_header_row=normalized_has_header_row,
                    shading=normalized_shading,
                    header_fill_color=header_fill_color,
                    header_text_color=header_text_color,
                    auto_fit=normalized_auto_fit,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="set_paragraph_format",
        description=(
            "Apply paragraph and run formatting to selected top-level paragraphs by indices or text match."
        ),
    )
    def set_paragraph_format_tool(
        filename: str,
        paragraph_indices: list[int] | str | None = None,
        contains_text: str | None = None,
        font_name: str | None = None,
        font_size: float | int | str | None = None,
        bold: bool | str | None = None,
        italic: bool | str | None = None,
        color: str | None = None,
        alignment: str | None = None,
        line_spacing: float | int | str | None = None,
        space_before_pt: float | int | str | None = None,
        space_after_pt: float | int | str | None = None,
        left_indent_pt: float | int | str | None = None,
        right_indent_pt: float | int | str | None = None,
        first_line_indent_pt: float | int | str | None = None,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_paragraph_indices = parse_paragraph_indices(paragraph_indices)
            normalized_font_size = parse_optional_float(font_size, field="font_size")
            normalized_bold = parse_optional_bool(bold, field="bold")
            normalized_italic = parse_optional_bool(italic, field="italic")
            normalized_line_spacing = parse_optional_float(
                line_spacing, field="line_spacing"
            )
            normalized_space_before_pt = parse_optional_float(
                space_before_pt, field="space_before_pt"
            )
            normalized_space_after_pt = parse_optional_float(
                space_after_pt, field="space_after_pt"
            )
            normalized_left_indent_pt = parse_optional_float(
                left_indent_pt, field="left_indent_pt"
            )
            normalized_right_indent_pt = parse_optional_float(
                right_indent_pt, field="right_indent_pt"
            )
            normalized_first_line_indent_pt = parse_optional_float(
                first_line_indent_pt, field="first_line_indent_pt"
            )
            return {
                "ok": True,
                "result": set_paragraph_format(
                    filename=filename,
                    paragraph_indices=normalized_paragraph_indices,
                    contains_text=contains_text,
                    font_name=font_name,
                    font_size=normalized_font_size,
                    bold=normalized_bold,
                    italic=normalized_italic,
                    color=color,
                    alignment=alignment,
                    line_spacing=normalized_line_spacing,
                    space_before_pt=normalized_space_before_pt,
                    space_after_pt=normalized_space_after_pt,
                    left_indent_pt=normalized_left_indent_pt,
                    right_indent_pt=normalized_right_indent_pt,
                    first_line_indent_pt=normalized_first_line_indent_pt,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)
