from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from docx_mcp.services.document_style_ops import apply_document_style
from docx_mcp.tools.common import handle_tool_error
from docx_mcp.tools.compat import (
    parse_bool,
    parse_int,
    parse_optional_float,
)


def register_document_style_tools(mcp: FastMCP) -> None:
    @mcp.tool(
        name="apply_document_style",
        description=(
            "Apply page layout and generic paragraph/heading style settings to a DOCX."
        ),
    )
    def apply_document_style_tool(
        filename: str,
        page_size: str = "A4",
        margin_top_cm: float | int | str = 2.54,
        margin_bottom_cm: float | int | str = 2.54,
        margin_left_cm: float | int | str = 2.54,
        margin_right_cm: float | int | str = 2.54,
        normal_font_name: str = "Calibri",
        normal_western_font_name: str = "Calibri",
        normal_font_size_pt: float | int | str = 11.0,
        normal_line_spacing: float | int | str = 1.15,
        normal_first_line_indent_pt: float | int | str = 0.0,
        normal_alignment: str = "left",
        heading_font_name: str = "Calibri",
        heading_western_font_name: str = "Calibri",
        heading_1_size_pt: float | int | str = 16.0,
        heading_2_size_pt: float | int | str = 14.0,
        heading_3_size_pt: float | int | str = 12.0,
        max_heading_level: int | str = 3,
        apply_to_existing_paragraphs: bool | str = True,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_margin_top_cm = parse_optional_float(
                margin_top_cm, field="margin_top_cm"
            )
            normalized_margin_bottom_cm = parse_optional_float(
                margin_bottom_cm, field="margin_bottom_cm"
            )
            normalized_margin_left_cm = parse_optional_float(
                margin_left_cm, field="margin_left_cm"
            )
            normalized_margin_right_cm = parse_optional_float(
                margin_right_cm, field="margin_right_cm"
            )
            normalized_normal_font_size = parse_optional_float(
                normal_font_size_pt, field="normal_font_size_pt"
            )
            normalized_normal_line_spacing = parse_optional_float(
                normal_line_spacing, field="normal_line_spacing"
            )
            normalized_normal_first_line_indent = parse_optional_float(
                normal_first_line_indent_pt, field="normal_first_line_indent_pt"
            )
            normalized_heading_1_size = parse_optional_float(
                heading_1_size_pt, field="heading_1_size_pt"
            )
            normalized_heading_2_size = parse_optional_float(
                heading_2_size_pt, field="heading_2_size_pt"
            )
            normalized_heading_3_size = parse_optional_float(
                heading_3_size_pt, field="heading_3_size_pt"
            )
            normalized_max_heading_level = parse_int(
                max_heading_level, field="max_heading_level"
            )
            normalized_apply_to_existing = parse_bool(
                apply_to_existing_paragraphs, field="apply_to_existing_paragraphs"
            )
            return {
                "ok": True,
                "result": apply_document_style(
                    filename=filename,
                    page_size=page_size,
                    margin_top_cm=normalized_margin_top_cm or 2.54,
                    margin_bottom_cm=normalized_margin_bottom_cm or 2.54,
                    margin_left_cm=normalized_margin_left_cm or 2.54,
                    margin_right_cm=normalized_margin_right_cm or 2.54,
                    normal_font_name=normal_font_name,
                    normal_western_font_name=normal_western_font_name,
                    normal_font_size_pt=normalized_normal_font_size or 11.0,
                    normal_line_spacing=normalized_normal_line_spacing or 1.15,
                    normal_first_line_indent_pt=(
                        normalized_normal_first_line_indent or 0.0
                    ),
                    normal_alignment=normal_alignment,
                    heading_font_name=heading_font_name,
                    heading_western_font_name=heading_western_font_name,
                    heading_1_size_pt=normalized_heading_1_size or 16.0,
                    heading_2_size_pt=normalized_heading_2_size or 14.0,
                    heading_3_size_pt=normalized_heading_3_size or 12.0,
                    max_heading_level=normalized_max_heading_level,
                    apply_to_existing_paragraphs=normalized_apply_to_existing,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)
