from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from docx_mcp.services.section_ops import set_headers_footers
from docx_mcp.tools.common import handle_tool_error
from docx_mcp.tools.compat import (
    parse_bool,
    parse_optional_bool,
    parse_optional_int,
    parse_paragraph_indices,
)


def register_section_tools(mcp: FastMCP) -> None:
    @mcp.tool(
        name="set_headers_footers",
        description=(
            "Set section headers/footers, optional PAGE field, and section page number start."
        ),
    )
    def set_headers_footers_tool(
        filename: str,
        section_indices: list[int] | str | None = None,
        header_text: str | None = None,
        footer_text: str | None = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
        include_page_number: bool | str = False,
        clear_existing: bool | str = True,
        start_page_number: int | str | None = None,
        different_first_page: bool | str | None = None,
        different_odd_even: bool | str | None = None,
        unlink_from_previous: bool | str = True,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_section_indices = parse_paragraph_indices(section_indices)
            normalized_include_page_number = parse_bool(
                include_page_number, field="include_page_number"
            )
            normalized_clear_existing = parse_bool(
                clear_existing, field="clear_existing"
            )
            normalized_start_page_number = parse_optional_int(
                start_page_number, field="start_page_number"
            )
            normalized_different_first_page = parse_optional_bool(
                different_first_page, field="different_first_page"
            )
            normalized_different_odd_even = parse_optional_bool(
                different_odd_even, field="different_odd_even"
            )
            normalized_unlink = parse_bool(
                unlink_from_previous, field="unlink_from_previous"
            )
            return {
                "ok": True,
                "result": set_headers_footers(
                    filename=filename,
                    section_indices=normalized_section_indices,
                    header_text=header_text,
                    footer_text=footer_text,
                    header_alignment=header_alignment,
                    footer_alignment=footer_alignment,
                    include_page_number=normalized_include_page_number,
                    clear_existing=normalized_clear_existing,
                    start_page_number=normalized_start_page_number,
                    different_first_page=normalized_different_first_page,
                    different_odd_even=normalized_different_odd_even,
                    unlink_from_previous=normalized_unlink,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)
