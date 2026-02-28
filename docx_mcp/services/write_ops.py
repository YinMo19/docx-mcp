from __future__ import annotations

from typing import Any

from docx_mcp.errors import DocxMCPError
from docx_mcp.services.document_io import (
    open_document,
    resolve_for_write,
    save_document,
)
from docx_mcp.services.format_utils import apply_run_format, iter_all_paragraphs


def search_and_replace(
    filename: str,
    find_text: str,
    replace_text: str,
    output_filename: str | None = None,
) -> dict[str, Any]:
    if not find_text:
        raise DocxMCPError(
            code="INVALID_QUERY",
            message="find_text cannot be empty.",
        )

    src_path, dst_path = resolve_for_write(filename, output_filename)
    doc = open_document(src_path)

    replacement_count = 0
    touched_paragraphs = 0

    for paragraph in iter_all_paragraphs(doc):
        original_text = paragraph.text or ""
        if find_text not in original_text:
            continue

        touched_paragraphs += 1
        replaced_in_runs = 0
        if paragraph.runs:
            for run in paragraph.runs:
                text = run.text or ""
                if find_text in text:
                    count = text.count(find_text)
                    run.text = text.replace(find_text, replace_text)
                    replaced_in_runs += count
            if replaced_in_runs > 0:
                replacement_count += replaced_in_runs
            else:
                count = original_text.count(find_text)
                paragraph.text = original_text.replace(find_text, replace_text)
                replacement_count += count
        else:
            count = original_text.count(find_text)
            paragraph.text = original_text.replace(find_text, replace_text)
            replacement_count += count

    save_document(doc, dst_path)
    return {
        "source_path": str(src_path),
        "output_path": str(dst_path),
        "find_text": find_text,
        "replace_text": replace_text,
        "replacement_count": replacement_count,
        "touched_paragraphs": touched_paragraphs,
    }


def add_paragraph(
    filename: str,
    text: str,
    style: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
    output_filename: str | None = None,
) -> dict[str, Any]:
    src_path, dst_path = resolve_for_write(filename, output_filename)
    doc = open_document(src_path)

    paragraph = doc.add_paragraph(text)
    if style:
        paragraph.style = style

    if paragraph.runs:
        run = paragraph.runs[0]
    else:
        run = paragraph.add_run(text)
    apply_run_format(
        run,
        font_name=font_name,
        font_size=font_size,
        bold=True if bold else None,
        italic=True if italic else None,
        color=color,
    )

    save_document(doc, dst_path)
    return {
        "source_path": str(src_path),
        "output_path": str(dst_path),
        "paragraph_index": len(doc.paragraphs) - 1,
        "text": text,
        "style": style,
    }


def add_heading(
    filename: str,
    text: str,
    level: int = 1,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool = False,
    italic: bool = False,
    output_filename: str | None = None,
) -> dict[str, Any]:
    if level < 0 or level > 9:
        raise DocxMCPError(
            code="INVALID_HEADING_LEVEL",
            message="Heading level must be between 0 and 9.",
            details={"level": level},
        )

    src_path, dst_path = resolve_for_write(filename, output_filename)
    doc = open_document(src_path)

    heading = doc.add_heading(text, level=level)
    if heading.runs:
        run = heading.runs[0]
    else:
        run = heading.add_run(text)
    apply_run_format(
        run,
        font_name=font_name,
        font_size=font_size,
        bold=True if bold else None,
        italic=True if italic else None,
    )

    save_document(doc, dst_path)
    return {
        "source_path": str(src_path),
        "output_path": str(dst_path),
        "paragraph_index": len(doc.paragraphs) - 1,
        "text": text,
        "level": level,
    }


def add_table(
    filename: str,
    rows: int,
    cols: int,
    data: list[list[str]] | None = None,
    output_filename: str | None = None,
) -> dict[str, Any]:
    if rows <= 0 or cols <= 0:
        raise DocxMCPError(
            code="INVALID_TABLE_DIMENSIONS",
            message="rows and cols must be greater than 0.",
            details={"rows": rows, "cols": cols},
        )

    src_path, dst_path = resolve_for_write(filename, output_filename)
    doc = open_document(src_path)

    table = doc.add_table(rows=rows, cols=cols)
    if data:
        for r_idx, row_values in enumerate(data[:rows]):
            for c_idx, value in enumerate(row_values[:cols]):
                table.cell(r_idx, c_idx).text = value

    save_document(doc, dst_path)
    return {
        "source_path": str(src_path),
        "output_path": str(dst_path),
        "table_index": len(doc.tables) - 1,
        "rows": rows,
        "cols": cols,
        "filled_rows": min(rows, len(data) if data else 0),
    }
