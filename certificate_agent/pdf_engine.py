"""
PDF Certificate Engine.
Directly manipulates existing PDF certificate templates using PyMuPDF (fitz),
preserving original resolution, vectors, fonts, and dimensions.
"""

import os
import re
from typing import List, Dict, Any, Tuple, Optional
import pymupdf as fitz

def analyze_pdf_template(pdf_path: str) -> Dict[str, Any]:
    """
    Analyzes a PDF template to determine dimensions and placeholder locations.
    """
    doc = fitz.open(pdf_path)
    if len(doc) == 0:
        raise ValueError(f"PDF template '{pdf_path}' contains no pages.")

    page = doc[0]
    rect = page.rect
    width, height = rect.width, rect.height

    # Check for text placeholders
    name_matches = []
    for ph in ["{{NAME}}", "[NAME]", "<NAME>", "{{Name}}", "[Name]"]:
        rects = page.search_for(ph)
        if rects:
            name_matches.extend(rects)

    reg_matches = []
    for ph in ["{{REG_NO}}", "[REG NO]", "<REG NO>", "{{REGNO}}", "[REGNO]"]:
        rects = page.search_for(ph)
        if rects:
            reg_matches.extend(rects)

    if name_matches:
        mode = "placeholder"
        name_rect = name_matches[0]
        reg_rect = reg_matches[0] if reg_matches else None
    else:
        mode = "layout"
        # Search for preamble and completion text
        intro_rects = page.search_for("certify that")
        if not intro_rects:
            intro_rects = page.search_for("presented to")

        body_rects = page.search_for("completed")
        if not body_rects:
            body_rects = page.search_for("organized by")

        if intro_rects and body_rects:
            y_top = intro_rects[0].y1 + 10
            y_bottom = body_rects[0].y0 - 10
        else:
            y_top = height * 0.38
            y_bottom = height * 0.55

        center_x = width / 2.0
        name_rect = fitz.Rect(50, y_top, width - 50, (y_top + y_bottom) / 2.0 + 5)
        reg_rect = fitz.Rect(50, (y_top + y_bottom) / 2.0 + 5, width - 50, y_bottom)

    doc.close()
    return {
        "mode": mode,
        "width": width,
        "height": height,
        "name_rect": name_rect,
        "reg_rect": reg_rect
    }

def generate_pdf_for_student(
    template_path: str,
    output_pdf_path: str,
    student: Dict[str, Any],
    layout_info: Dict[str, Any]
):
    """
    Generates an individualized certificate PDF from a PDF template.
    """
    doc = fitz.open(template_path)
    page = doc[0]
    name = student["name"].strip()
    reg_no = student["reg_no"].strip()

    name_rect = layout_info["name_rect"]
    reg_rect = layout_info["reg_rect"]

    # If placeholder mode, redact placeholder boxes first
    if layout_info["mode"] == "placeholder":
        page.add_redact_annot(name_rect, fill=(1, 1, 1))
        if reg_rect:
            page.add_redact_annot(reg_rect, fill=(1, 1, 1))
        page.apply_redactions()

    # Dynamic font sizing
    if len(name) <= 22:
        name_fontsize = 24
    elif len(name) <= 32:
        name_fontsize = 20
    else:
        name_fontsize = 17

    # Center Name
    page.insert_textbox(
        name_rect,
        name,
        fontsize=name_fontsize,
        fontname="helv", # Standard Helvetica Bold
        fontfile=None,
        color=(0.1, 0.1, 0.1),
        align=fitz.TEXT_ALIGN_CENTER
    )

    # Center Reg No
    if reg_rect:
        reg_text = f"(Reg No: {reg_no})"
        page.insert_textbox(
            reg_rect,
            reg_text,
            fontsize=13,
            fontname="helv",
            color=(0.25, 0.25, 0.25),
            align=fitz.TEXT_ALIGN_CENTER
        )

    doc.save(output_pdf_path)
    doc.close()
