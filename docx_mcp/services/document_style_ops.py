from __future__ import annotations

import re
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

from docx_mcp.errors import DocxMCPError
from docx_mcp.services.document_io import open_document, resolve_for_write, save_document

_ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_PAGE_SIZE_MM = {
    "A4": (210.0, 297.0),
    "LETTER": (215.9, 279.4),
}


def _validate_positive(value: float, field: str) -> float:
    if value <= 0:
        raise DocxMCPError(
            code="INVALID_NUMBER",
            message=f"{field} must be greater than 0.",
            details={"field": field, "value": value},
        )
    return value


def _validate_non_negative(value: float, field: str) -> float:
    if value < 0:
        raise DocxMCPError(
            code="INVALID_NUMBER",
            message=f"{field} must be greater than or equal to 0.",
            details={"field": field, "value": value},
        )
    return value


def _get_heading_level(style_name: str) -> int | None:
    lower = style_name.lower()
    en_match = re.search(r"heading\s*([0-9]+)", lower)
    if en_match:
        return int(en_match.group(1))
    zh_match = re.search(r"标题\s*([0-9]+)", style_name)
    if zh_match:
        return int(zh_match.group(1))
    return None


def _set_style_rfonts(style, east_asia_font: str, western_font: str) -> None:
    style_el = style._element
    r_pr = style_el.find(qn("w:rPr"))
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        style_el.append(r_pr)

    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)

    r_fonts.set(qn("w:ascii"), western_font)
    r_fonts.set(qn("w:hAnsi"), western_font)
    r_fonts.set(qn("w:cs"), western_font)
    r_fonts.set(qn("w:eastAsia"), east_asia_font)


def _apply_style_font(style, east_asia_font: str, western_font: str, size_pt: float) -> None:
    style.font.name = western_font
    style.font.size = Pt(size_pt)
    _set_style_rfonts(style, east_asia_font=east_asia_font, western_font=western_font)


def _apply_run_font(
    run,
    *,
    east_asia_font: str,
    western_font: str,
    size_pt: float,
    bold: bool | None = None,
) -> None:
    run.font.name = western_font
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), western_font)
    r_fonts.set(qn("w:hAnsi"), western_font)
    r_fonts.set(qn("w:cs"), western_font)
    r_fonts.set(qn("w:eastAsia"), east_asia_font)


def apply_document_style(
    filename: str,
    *,
    page_size: str = "A4",
    margin_top_cm: float = 2.54,
    margin_bottom_cm: float = 2.54,
    margin_left_cm: float = 2.54,
    margin_right_cm: float = 2.54,
    normal_font_name: str = "Calibri",
    normal_western_font_name: str = "Calibri",
    normal_font_size_pt: float = 11.0,
    normal_line_spacing: float = 1.15,
    normal_first_line_indent_pt: float = 0.0,
    normal_alignment: str = "left",
    heading_font_name: str = "Calibri",
    heading_western_font_name: str = "Calibri",
    heading_1_size_pt: float = 16.0,
    heading_2_size_pt: float = 14.0,
    heading_3_size_pt: float = 12.0,
    max_heading_level: int = 3,
    apply_to_existing_paragraphs: bool = True,
    output_filename: str | None = None,
) -> dict[str, Any]:
    normalized_page_size = page_size.upper()
    if normalized_page_size not in _PAGE_SIZE_MM:
        raise DocxMCPError(
            code="INVALID_PAGE_SIZE",
            message=f"Unsupported page_size: {page_size}",
            details={"supported": sorted(_PAGE_SIZE_MM.keys())},
        )
    if max_heading_level < 1 or max_heading_level > 9:
        raise DocxMCPError(
            code="INVALID_HEADING_LEVEL",
            message="max_heading_level must be between 1 and 9.",
            details={"max_heading_level": max_heading_level},
        )

    _validate_positive(normal_font_size_pt, "normal_font_size_pt")
    _validate_positive(normal_line_spacing, "normal_line_spacing")
    _validate_non_negative(normal_first_line_indent_pt, "normal_first_line_indent_pt")
    _validate_positive(heading_1_size_pt, "heading_1_size_pt")
    _validate_positive(heading_2_size_pt, "heading_2_size_pt")
    _validate_positive(heading_3_size_pt, "heading_3_size_pt")
    _validate_positive(margin_top_cm, "margin_top_cm")
    _validate_positive(margin_bottom_cm, "margin_bottom_cm")
    _validate_positive(margin_left_cm, "margin_left_cm")
    _validate_positive(margin_right_cm, "margin_right_cm")

    normalized_alignment = _ALIGNMENT_MAP.get(normal_alignment.lower())
    if normalized_alignment is None:
        raise DocxMCPError(
            code="INVALID_ALIGNMENT",
            message=f"Unsupported normal_alignment: {normal_alignment}",
            details={"supported": sorted(_ALIGNMENT_MAP.keys())},
        )

    src_path, dst_path = resolve_for_write(filename, output_filename)
    doc = open_document(src_path)

    page_width_mm, page_height_mm = _PAGE_SIZE_MM[normalized_page_size]
    for section in doc.sections:
        section.page_width = Mm(page_width_mm)
        section.page_height = Mm(page_height_mm)
        section.top_margin = Cm(margin_top_cm)
        section.bottom_margin = Cm(margin_bottom_cm)
        section.left_margin = Cm(margin_left_cm)
        section.right_margin = Cm(margin_right_cm)

    normal_style = doc.styles["Normal"]
    _apply_style_font(
        normal_style,
        east_asia_font=normal_font_name,
        western_font=normal_western_font_name,
        size_pt=normal_font_size_pt,
    )
    normal_style.paragraph_format.line_spacing = normal_line_spacing
    normal_style.paragraph_format.first_line_indent = Pt(normal_first_line_indent_pt)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)

    heading_size_map = {
        1: heading_1_size_pt,
        2: heading_2_size_pt,
        3: heading_3_size_pt,
    }

    for level in range(1, max_heading_level + 1):
        style_name = f"Heading {level}"
        try:
            heading_style = doc.styles[style_name]
        except KeyError:
            continue
        _apply_style_font(
            heading_style,
            east_asia_font=heading_font_name,
            western_font=heading_western_font_name,
            size_pt=heading_size_map.get(level, heading_3_size_pt),
        )
        heading_style.font.bold = True

    paragraph_count = len(doc.paragraphs)
    heading_paragraph_count = 0
    body_paragraph_count = 0

    if apply_to_existing_paragraphs:
        for paragraph in doc.paragraphs:
            style_name = paragraph.style.name if paragraph.style else ""
            heading_level = _get_heading_level(style_name)
            if heading_level is not None and heading_level <= max_heading_level:
                heading_paragraph_count += 1
                paragraph.paragraph_format.first_line_indent = Pt(0)
                for run in paragraph.runs:
                    _apply_run_font(
                        run,
                        east_asia_font=heading_font_name,
                        western_font=heading_western_font_name,
                        size_pt=heading_size_map.get(heading_level, heading_3_size_pt),
                        bold=True,
                    )
                continue

            body_paragraph_count += 1
            paragraph.alignment = normalized_alignment
            p_format = paragraph.paragraph_format
            p_format.line_spacing = normal_line_spacing
            p_format.first_line_indent = Pt(normal_first_line_indent_pt)
            p_format.space_before = Pt(0)
            p_format.space_after = Pt(0)
            for run in paragraph.runs:
                _apply_run_font(
                    run,
                    east_asia_font=normal_font_name,
                    western_font=normal_western_font_name,
                    size_pt=normal_font_size_pt,
                )

    save_document(doc, dst_path)
    return {
        "source_path": str(src_path),
        "output_path": str(dst_path),
        "page_size": normalized_page_size,
        "section_count": len(doc.sections),
        "paragraph_count": paragraph_count,
        "heading_paragraph_count": heading_paragraph_count,
        "body_paragraph_count": body_paragraph_count,
        "apply_to_existing_paragraphs": apply_to_existing_paragraphs,
        "margins_cm": {
            "top": margin_top_cm,
            "bottom": margin_bottom_cm,
            "left": margin_left_cm,
            "right": margin_right_cm,
        },
    }
