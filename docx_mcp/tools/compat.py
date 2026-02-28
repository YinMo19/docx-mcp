from __future__ import annotations

import json
import re
from typing import Any

from docx_mcp.errors import DocxMCPError

_TRUE_VALUES = {"1", "true", "yes", "y", "on"}
_FALSE_VALUES = {"0", "false", "no", "n", "off"}


def parse_bool(value: bool | str, field: str) -> bool:
    if isinstance(value, bool):
        return value
    normalized = value.strip().lower()
    if normalized in _TRUE_VALUES:
        return True
    if normalized in _FALSE_VALUES:
        return False
    raise DocxMCPError(
        code="INVALID_BOOLEAN",
        message=f"{field} must be boolean-like.",
        details={"field": field, "value": value},
    )


def parse_optional_bool(value: bool | str | None, field: str) -> bool | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    return parse_bool(value, field=field)


def parse_int(value: int | str, field: str) -> int:
    if isinstance(value, bool):
        raise DocxMCPError(
            code="INVALID_INTEGER",
            message=f"{field} must be an integer.",
            details={"field": field, "value": value},
        )
    if isinstance(value, int):
        return value
    raw = value.strip()
    if not raw:
        raise DocxMCPError(
            code="INVALID_INTEGER",
            message=f"{field} must be an integer.",
            details={"field": field, "value": value},
        )
    try:
        return int(raw)
    except Exception as exc:
        raise DocxMCPError(
            code="INVALID_INTEGER",
            message=f"{field} must be an integer.",
            details={"field": field, "value": value},
        ) from exc


def parse_optional_int(value: int | str | None, field: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    return parse_int(value, field=field)


def parse_optional_float(value: float | int | str | None, field: str) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        raise DocxMCPError(
            code="INVALID_NUMBER",
            message=f"{field} must be numeric.",
            details={"field": field, "value": value},
        )
    if isinstance(value, (int, float)):
        return float(value)
    raw = value.strip()
    if not raw:
        return None
    try:
        return float(raw)
    except Exception as exc:
        raise DocxMCPError(
            code="INVALID_NUMBER",
            message=f"{field} must be numeric.",
            details={"field": field, "value": value},
        ) from exc


def parse_paragraph_indices(value: list[int] | str | None) -> list[int] | None:
    if value is None:
        return None
    if isinstance(value, list):
        try:
            return [int(v) for v in value]
        except Exception as exc:
            raise DocxMCPError(
                code="INVALID_PARAGRAPH_INDICES",
                message="paragraph_indices must be a list of integers.",
                details={"value": value},
            ) from exc

    raw = value.strip()
    if not raw:
        return []

    # JSON array support: "[1,2,3]"
    if raw.startswith("["):
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise DocxMCPError(
                code="INVALID_PARAGRAPH_INDICES",
                message="paragraph_indices JSON string is invalid.",
                details={"value": value},
            ) from exc
        if not isinstance(parsed, list):
            raise DocxMCPError(
                code="INVALID_PARAGRAPH_INDICES",
                message="paragraph_indices JSON must be a list.",
                details={"value": value},
            )
        try:
            return [int(v) for v in parsed]
        except Exception as exc:
            raise DocxMCPError(
                code="INVALID_PARAGRAPH_INDICES",
                message="paragraph_indices JSON must contain integers only.",
                details={"value": value},
            ) from exc

    parts = [p for p in re.split(r"[\s,;]+", raw) if p]
    try:
        return [int(p) for p in parts]
    except Exception as exc:
        raise DocxMCPError(
            code="INVALID_PARAGRAPH_INDICES",
            message="paragraph_indices string must contain integers only.",
            details={"value": value},
        ) from exc


def parse_matrix(value: list[list[str]] | str | None, field: str) -> list[list[str]] | None:
    if value is None:
        return None
    if isinstance(value, list):
        result: list[list[str]] = []
        for row in value:
            if not isinstance(row, list):
                raise DocxMCPError(
                    code="INVALID_MATRIX",
                    message=f"{field} must be a 2D list.",
                    details={"field": field, "value": value},
                )
            result.append([str(cell) for cell in row])
        return result

    raw = value.strip()
    if not raw:
        return None

    # JSON array support: [["a","b"],["c","d"]]
    if raw.startswith("["):
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise DocxMCPError(
                code="INVALID_MATRIX",
                message=f"{field} JSON string is invalid.",
                details={"field": field, "value": value},
            ) from exc
        if not isinstance(parsed, list):
            raise DocxMCPError(
                code="INVALID_MATRIX",
                message=f"{field} JSON must be a 2D list.",
                details={"field": field, "value": value},
            )
        result: list[list[str]] = []
        for row in parsed:
            if not isinstance(row, list):
                raise DocxMCPError(
                    code="INVALID_MATRIX",
                    message=f"{field} JSON must be a 2D list.",
                    details={"field": field, "value": value},
                )
            result.append([str(cell) for cell in row])
        return result

    # Line-based fallback:
    # "a,b\nc,d" or "a\tb\nc\td"
    rows = [line for line in raw.splitlines() if line.strip()]
    result = []
    for row in rows:
        if "\t" in row:
            cells = [c.strip() for c in row.split("\t")]
        else:
            cells = [c.strip() for c in row.split(",")]
        result.append(cells)
    return result


def parse_shading(value: list[str] | str | None) -> list[str] | None:
    if value is None:
        return None
    if isinstance(value, list):
        return [str(v) for v in value if str(v).strip()]

    raw = value.strip()
    if not raw:
        return None
    if raw.startswith("["):
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise DocxMCPError(
                code="INVALID_SHADING",
                message="shading JSON string is invalid.",
                details={"value": value},
            ) from exc
        if not isinstance(parsed, list):
            raise DocxMCPError(
                code="INVALID_SHADING",
                message="shading JSON must be a list.",
                details={"value": value},
            )
        return [str(v) for v in parsed]

    return [part.strip() for part in raw.split(",") if part.strip()]


def parse_auto_fit(value: bool | str | None) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    normalized = value.strip().lower()
    if not normalized:
        return None
    if normalized in _TRUE_VALUES or normalized in {"content", "window"}:
        return True
    if normalized in _FALSE_VALUES or normalized == "fixed":
        return False
    raise DocxMCPError(
        code="INVALID_AUTO_FIT",
        message="auto_fit must be boolean-like or one of: content, window, fixed.",
        details={"value": value},
    )
