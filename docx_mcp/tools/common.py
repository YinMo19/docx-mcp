from __future__ import annotations

from typing import Any

from docx_mcp.errors import DocxMCPError


def handle_tool_error(exc: Exception) -> dict[str, Any]:
    if isinstance(exc, DocxMCPError):
        return {
            "ok": False,
            "error": {
                "code": exc.code,
                "message": exc.message,
                "details": exc.details,
            },
        }
    return {
        "ok": False,
        "error": {
            "code": "UNEXPECTED_ERROR",
            "message": str(exc),
            "details": {},
        },
    }

