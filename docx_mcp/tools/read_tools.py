from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from docx_mcp.services.read_ops import (
    find_text_in_document,
    get_document_info,
    get_document_outline,
    get_document_text,
    list_available_documents,
)
from docx_mcp.tools.common import handle_tool_error


def register_read_tools(mcp: FastMCP) -> None:
    @mcp.tool(
        name="list_available_documents",
        description="List .docx files under the given directory (or current working directory).",
    )
    def list_available_documents_tool(directory: str | None = None) -> dict[str, Any]:
        try:
            return {"ok": True, "result": list_available_documents(directory)}
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="get_document_info",
        description="Read document metadata and basic statistics from a .docx file.",
    )
    def get_document_info_tool(filename: str) -> dict[str, Any]:
        try:
            return {"ok": True, "result": get_document_info(filename)}
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="get_document_text",
        description="Extract plain text lines from a .docx file.",
    )
    def get_document_text_tool(filename: str) -> dict[str, Any]:
        try:
            return {"ok": True, "result": get_document_text(filename)}
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="get_document_outline",
        description="Extract paragraph/table outline from a .docx file.",
    )
    def get_document_outline_tool(filename: str) -> dict[str, Any]:
        try:
            return {"ok": True, "result": get_document_outline(filename)}
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)

    @mcp.tool(
        name="find_text_in_document",
        description="Find text occurrences in document paragraphs and tables.",
    )
    def find_text_in_document_tool(
        filename: str,
        text_to_find: str,
        match_case: bool = False,
        whole_word: bool = False,
    ) -> dict[str, Any]:
        try:
            return {
                "ok": True,
                "result": find_text_in_document(
                    filename=filename,
                    text_to_find=text_to_find,
                    match_case=match_case,
                    whole_word=whole_word,
                ),
            }
        except Exception as exc:  # pragma: no cover - thin wrapper
            return handle_tool_error(exc)
