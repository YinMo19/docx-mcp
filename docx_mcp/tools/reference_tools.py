from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from docx_mcp.services.reference_ops import (
    add_bookmark_to_paragraph,
    add_sequence_caption,
    insert_ref_field,
    insert_table_of_contents,
)
from docx_mcp.tools.common import handle_tool_error
from docx_mcp.tools.compat import parse_bool, parse_int


def register_reference_tools(mcp: FastMCP) -> None:
    @mcp.tool(
        name="insert_table_of_contents",
        description="Insert a TOC field based on heading levels.",
    )
    def insert_table_of_contents_tool(
        filename: str,
        heading_start: int | str = 1,
        heading_end: int | str = 3,
        title_text: str | None = "Table of Contents",
        add_page_break_before: bool | str = False,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_heading_start = parse_int(heading_start, field="heading_start")
            normalized_heading_end = parse_int(heading_end, field="heading_end")
            normalized_add_page_break = parse_bool(
                add_page_break_before, field="add_page_break_before"
            )
            return {
                "ok": True,
                "result": insert_table_of_contents(
                    filename=filename,
                    heading_start=normalized_heading_start,
                    heading_end=normalized_heading_end,
                    title_text=title_text,
                    add_page_break_before=normalized_add_page_break,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="add_sequence_caption",
        description="Append a caption paragraph using a SEQ field (Figure/Table/Equation).",
    )
    def add_sequence_caption_tool(
        filename: str,
        label: str,
        caption_text: str,
        seq_identifier: str | None = None,
        separator: str = ": ",
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            return {
                "ok": True,
                "result": add_sequence_caption(
                    filename=filename,
                    label=label,
                    caption_text=caption_text,
                    seq_identifier=seq_identifier,
                    separator=separator,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="add_bookmark_to_paragraph",
        description="Add a bookmark start/end pair to a paragraph by index.",
    )
    def add_bookmark_to_paragraph_tool(
        filename: str,
        paragraph_index: int | str,
        bookmark_name: str,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_index = parse_int(paragraph_index, field="paragraph_index")
            return {
                "ok": True,
                "result": add_bookmark_to_paragraph(
                    filename=filename,
                    paragraph_index=normalized_index,
                    bookmark_name=bookmark_name,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="insert_ref_field",
        description="Append a REF field paragraph pointing to a bookmark.",
    )
    def insert_ref_field_tool(
        filename: str,
        bookmark_name: str,
        prefix_text: str | None = None,
        hyperlink: bool | str = True,
        output_filename: str | None = None,
    ) -> dict[str, Any]:
        try:
            normalized_hyperlink = parse_bool(hyperlink, field="hyperlink")
            return {
                "ok": True,
                "result": insert_ref_field(
                    filename=filename,
                    bookmark_name=bookmark_name,
                    prefix_text=prefix_text,
                    hyperlink=normalized_hyperlink,
                    output_filename=output_filename,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)
