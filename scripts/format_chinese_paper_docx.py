#!/usr/bin/env python
"""Apply the user's default Chinese academic DOCX formatting rules."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


BODY_FONT_PT = 12
INDENT_PT = BODY_FONT_PT * 2

HEADING_RE = re.compile(r"^(?:[一二三四五六七八九十]+、|\d+(?:\.\d+)*\s+|\d+[.、])")
REF_RE = re.compile(r"^\[\d+\]")


def has_drawing(paragraph) -> bool:
    return bool(paragraph._p.xpath(".//w:drawing")) or bool(paragraph._p.xpath(".//w:pict"))


def is_caption(text: str) -> bool:
    text = text.strip()
    return bool(re.match(r"^[图表]\s*\d+", text)) or text.startswith(("图 ", "表 "))


def is_heading(text: str) -> bool:
    text = text.strip()
    return bool(text and HEADING_RE.match(text))


def is_reference(text: str) -> bool:
    return bool(REF_RE.match(text.strip()))


def is_bibliography_lead(text: str) -> bool:
    text = text.strip()
    return "[J]." in text or "[R]." in text or "[EB/OL]" in text


def is_body(text: str, paragraph) -> bool:
    text = text.strip()
    if not text or has_drawing(paragraph):
        return False
    if is_caption(text) or is_heading(text) or is_reference(text) or is_bibliography_lead(text):
        return False
    return True


def iter_paragraphs(document: Document):
    yield from document.paragraphs
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def set_run_fonts(paragraph) -> None:
    for run in paragraph.runs:
        if run.text and all(char in "“”‘’" for char in run.text):
            set_quote_run_font(run)
        else:
            set_standard_run_font(run)
        if run.font.size is None:
            run.font.size = Pt(BODY_FONT_PT)


def set_standard_run_font(run) -> None:
    r_fonts = get_or_add_r_fonts(run)
    run.font.name = "Times New Roman"
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "宋体")


def set_quote_run_font(run) -> None:
    r_fonts = get_or_add_r_fonts(run)
    run.font.name = "宋体"
    r_fonts.set(qn("w:ascii"), "宋体")
    r_fonts.set(qn("w:hAnsi"), "宋体")
    r_fonts.set(qn("w:eastAsia"), "宋体")


def get_or_add_r_fonts(run):
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    return r_fonts


def clear_paragraph_content(paragraph) -> None:
    p_element = paragraph._p
    for child in list(p_element):
        if child.tag != qn("w:pPr"):
            p_element.remove(child)


def rebuild_runs_with_quote_fonts(paragraph, text: str) -> None:
    clear_paragraph_content(paragraph)
    buffer = []
    for char in text:
        if char in "“”‘’":
            if buffer:
                run = paragraph.add_run("".join(buffer))
                set_standard_run_font(run)
                run.font.size = Pt(BODY_FONT_PT)
                buffer = []
            run = paragraph.add_run(char)
            set_quote_run_font(run)
            run.font.size = Pt(BODY_FONT_PT)
        else:
            buffer.append(char)
    if buffer:
        run = paragraph.add_run("".join(buffer))
        set_standard_run_font(run)
        run.font.size = Pt(BODY_FONT_PT)


def process_quotes(paragraph) -> int:
    text = paragraph.text
    converted = curly_quotes(text)
    if converted != text or any(char in converted for char in "“”‘’"):
        rebuild_runs_with_quote_fonts(paragraph, converted)
    return text.count('"')


def curly_quotes(text: str) -> str:
    out = []
    opening = True
    for char in text:
        if char == '"':
            out.append("“" if opening else "”")
            opening = not opening
        else:
            out.append(char)
    return "".join(out)


def clear_shading(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for shd in list(tc_pr.findall(qn("w:shd"))):
        tc_pr.remove(shd)


def cell_border(cell, **borders) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge, settings in borders.items():
        element = tc_borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tc_borders.append(element)
        for key, value in settings.items():
            element.set(qn(f"w:{key}"), str(value))


def apply_three_line_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    line = {"val": "single", "sz": "12", "space": "0", "color": "000000"}
    mid = {"val": "single", "sz": "8", "space": "0", "color": "000000"}
    nil = {"val": "nil"}

    if not table.rows:
        return

    for row in table.rows:
        for cell in row.cells:
            clear_shading(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_border(cell, top=nil, left=nil, bottom=nil, right=nil, insideH=nil, insideV=nil)
            for paragraph in cell.paragraphs:
                process_quotes(paragraph)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.first_line_indent = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                set_run_fonts(paragraph)

    for cell in table.rows[0].cells:
        clear_shading(cell)
        cell_border(cell, top=line, bottom=mid, left=nil, right=nil)
    for cell in table.rows[-1].cells:
        cell_border(cell, bottom=line, left=nil, right=nil)


def apply_document_format(path: Path, output: Path) -> dict:
    document = Document(path)

    for section in document.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    quote_replacements = 0
    body_count = 0
    heading_count = 0
    reference_count = 0
    image_count = 0
    caption_count = 0

    for paragraph in document.paragraphs:
        quote_replacements += process_quotes(paragraph)

        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        set_run_fonts(paragraph)

        stripped = paragraph.text.strip()
        if has_drawing(paragraph):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Pt(0)
            image_count += 1
        elif is_caption(stripped):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Pt(0)
            caption_count += 1
        elif is_reference(stripped):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Pt(0)
            reference_count += 1
        elif is_heading(stripped):
            paragraph.alignment = paragraph.alignment or WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Pt(0)
            heading_count += 1
        elif is_body(stripped, paragraph):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Pt(INDENT_PT)
            body_count += 1

    for table in document.tables:
        apply_three_line_table(table)

    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)

    return {
        "output": str(output),
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "body_paragraphs": body_count,
        "headings": heading_count,
        "references": reference_count,
        "captions": caption_count,
        "image_paragraphs": image_count,
        "quote_replacements": quote_replacements,
    }


def validate_docx(path: Path) -> dict:
    result = {
        "zip_ok": zipfile.is_zipfile(path),
        "ascii_double_quotes": None,
        "paragraphs": None,
        "tables": None,
    }
    if not result["zip_ok"]:
        return result
    document = Document(path)
    text = "\n".join(paragraph.text for paragraph in iter_paragraphs(document))
    result.update(
        {
            "ascii_double_quotes": text.count('"'),
            "paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
        }
    )
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    output = args.output or args.input
    summary = apply_document_format(args.input, output)
    summary["validation"] = validate_docx(output)

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(f"Formatted: {output}")
        print(f"Paragraphs: {summary['paragraphs']}; tables: {summary['tables']}")
        print(f"ASCII double quotes after formatting: {summary['validation']['ascii_double_quotes']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
