"""
DOCX Certificate Engine.
Analyzes DOCX template layout, performs mail merge with dynamic font scaling,
and exports high-fidelity PDFs via Word COM engine.
"""

import os
import re
import tempfile
import copy
from typing import List, Dict, Any, Tuple, Optional
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Regex patterns for explicit placeholders
PLACEHOLDER_NAME = re.compile(r'(\{\{\s*NAME\s*\}\}|\[\s*NAME\s*\]|<\s*NAME\s*>|\{\s*NAME\s*\})', re.IGNORECASE)
PLACEHOLDER_REG = re.compile(r'(\{\{\s*(REG[_\s]*NO|REGNO|ROLL[_\s]*NO)\s*\}\}|\[\s*(REG[_\s]*NO|REGNO|ROLL[_\s]*NO)\s*\]|<\s*(REG[_\s]*NO|REGNO|ROLL[_\s]*NO)\s*>|\{\s*(REG[_\s]*NO|REGNO|ROLL[_\s]*NO)\s*\})', re.IGNORECASE)

INTRO_KEYWORDS = [
    "certify that", "certifies that", "presented to", "awarded to", 
    "certificate is presented to", "proudly presented to", "is hereby awarded to"
]

COMPLETION_KEYWORDS = [
    "has successfully completed", "has completed", "for successfully completing",
    "in recognition of", "for participating", "for active participation",
    "value-added course", "organized by"
]

def detect_placeholders_in_doc(doc: docx.Document) -> Tuple[bool, bool]:
    """Checks if the document has explicit placeholders for name and reg_no."""
    has_name = False
    has_reg = False
    for p in doc.paragraphs:
        if PLACEHOLDER_NAME.search(p.text):
            has_name = True
        if PLACEHOLDER_REG.search(p.text):
            has_reg = True
    return has_name, has_reg

def replace_text_in_paragraph(p: docx.text.paragraph.Paragraph, search_pattern: re.Pattern, replacement: str):
    """Replaces text matching pattern in a paragraph while preserving basic formatting."""
    full_text = "".join(r.text for r in p.runs)
    if not search_pattern.search(full_text):
        return

    # If full text matches within single run
    for r in p.runs:
        if search_pattern.search(r.text):
            r.text = search_pattern.sub(replacement, r.text)
            return

    # If match spans multiple runs, replace in first run and clear others
    new_text = search_pattern.sub(replacement, full_text)
    if p.runs:
        first_run = p.runs[0]
        first_run.text = new_text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = new_text

def analyze_docx_structure(template_path: str) -> Dict[str, Any]:
    """
    Analyzes DOCX layout to determine placeholder presence or insertion point.
    """
    doc = docx.Document(template_path)
    has_name_ph, has_reg_ph = detect_placeholders_in_doc(doc)

    if has_name_ph:
        return {
            "mode": "placeholder",
            "has_name_placeholder": has_name_ph,
            "has_reg_placeholder": has_reg_ph
        }

    # Implicit structure detection
    intro_idx = -1
    completion_idx = -1

    for idx, p in enumerate(doc.paragraphs):
        text_lower = p.text.lower().strip()
        if intro_idx == -1:
            for kw in INTRO_KEYWORDS:
                if kw in text_lower:
                    intro_idx = idx
                    break
        elif completion_idx == -1:
            for kw in COMPLETION_KEYWORDS:
                if kw in text_lower:
                    completion_idx = idx
                    break

    target_idx = -1
    create_new = False

    if intro_idx != -1 and completion_idx != -1 and completion_idx > intro_idx:
        # Check if there is an empty paragraph between them
        for idx in range(intro_idx + 1, completion_idx):
            if not doc.paragraphs[idx].text.strip():
                target_idx = idx
                break
        if target_idx == -1:
            # No empty paragraph exists, target slot is right after intro
            target_idx = intro_idx + 1
            create_new = True
    elif intro_idx != -1:
        target_idx = intro_idx + 1
        if target_idx < len(doc.paragraphs) and not doc.paragraphs[target_idx].text.strip():
            create_new = False
        else:
            create_new = True
    else:
        # Fallback: middle of document
        target_idx = max(1, len(doc.paragraphs) // 2)
        create_new = False

    # Detect predominant body font from paragraphs following the intro
    detected_font = "Book Antiqua"
    candidate_fonts = []
    for p in doc.paragraphs[intro_idx + 1:]:
        for r in p.runs:
            if r.font.name:
                candidate_fonts.append(r.font.name)

    # If body has a specific font (e.g., Book Antiqua, Georgia, Times New Roman), use it
    serif_priority = ["Book Antiqua", "Georgia", "Garamond", "Times New Roman", "Cambria"]
    chosen_font = None
    for sf in serif_priority:
        if any(sf.lower() in f.lower() for f in candidate_fonts):
            chosen_font = sf
            break

    if not chosen_font and candidate_fonts:
        # Avoid Comic Sans for formal student name if other options or default exists
        filtered = [f for f in candidate_fonts if "comic" not in f.lower()]
        if filtered:
            chosen_font = filtered[0]

    detected_font = chosen_font or "Book Antiqua"

    return {
        "mode": "layout",
        "intro_index": intro_idx,
        "completion_index": completion_idx,
        "target_index": target_idx,
        "create_new": create_new,
        "recommended_font": detected_font
    }

def calculate_name_font_size(name: str) -> int:
    """Calculates responsive font size for student name to avoid overflow."""
    length = len(name.strip())
    if length <= 22:
        return 24
    elif length <= 30:
        return 21
    elif length <= 40:
        return 18
    else:
        return 16

def populate_docx_for_student(
    template_path: str,
    output_docx_path: str,
    student: Dict[str, Any],
    layout_info: Dict[str, Any]
):
    """
    Creates an individualized DOCX for a student by merging data into template.
    """
    doc = docx.Document(template_path)
    name = student["name"].strip()
    reg_no = student["reg_no"].strip()

    if layout_info["mode"] == "placeholder":
        for p in doc.paragraphs:
            if layout_info.get("has_name_placeholder"):
                replace_text_in_paragraph(p, PLACEHOLDER_NAME, name)
            if layout_info.get("has_reg_placeholder"):
                replace_text_in_paragraph(p, PLACEHOLDER_REG, reg_no)
        # Also check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if layout_info.get("has_name_placeholder"):
                            replace_text_in_paragraph(p, PLACEHOLDER_NAME, name)
                        if layout_info.get("has_reg_placeholder"):
                            replace_text_in_paragraph(p, PLACEHOLDER_REG, reg_no)
    else:
        # Layout mode
        target_idx = layout_info["target_index"]
        font_name = layout_info.get("recommended_font", "Book Antiqua")
        name_size = calculate_name_font_size(name)

        if layout_info["create_new"]:
            # Insert after target_idx - 1
            ref_p = doc.paragraphs[target_idx - 1]
            target_p = ref_p.insert_paragraph_before()
        else:
            if target_idx < len(doc.paragraphs):
                target_p = doc.paragraphs[target_idx]
            else:
                target_p = doc.add_paragraph()

        # Clear any existing text
        target_p.text = ""
        target_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        target_p.paragraph_format.space_before = Pt(6)
        target_p.paragraph_format.space_after = Pt(6)
        target_p.paragraph_format.line_spacing = 1.15

        # Run 1: Student Name
        r_name = target_p.add_run(f"{name}\n")
        r_name.font.name = font_name
        r_name.font.size = Pt(name_size)
        r_name.font.bold = True
        r_name.font.color.rgb = RGBColor(20, 20, 20)

        # Run 2: Reg No
        r_reg = target_p.add_run(f"(Reg No: {reg_no})")
        r_reg.font.name = font_name
        r_reg.font.size = Pt(13)
        r_reg.font.bold = False
        r_reg.font.color.rgb = RGBColor(60, 60, 60)

    doc.save(output_docx_path)

def convert_docx_batch_to_pdf(
    job_list: List[Tuple[str, str]]
) -> List[Tuple[str, bool, Optional[str]]]:
    """
    Converts a batch of (docx_path, pdf_path) pairs to PDF using Word COM.
    Returns: List of (pdf_path, success_flag, error_message)
    """
    import win32com.client

    results = []
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0 # wdAlertsNone

        for docx_path, pdf_path in job_list:
            abs_docx = os.path.abspath(docx_path)
            abs_pdf = os.path.abspath(pdf_path)
            try:
                wdoc = word.Documents.Open(abs_docx, ReadOnly=True)
                # 17 = wdFormatPDF
                wdoc.SaveAs(abs_pdf, FileFormat=17)
                wdoc.Close(SaveChanges=0)
                results.append((pdf_path, True, None))
            except Exception as e:
                results.append((pdf_path, False, str(e)))
    except Exception as e:
        # Fallback if COM fails completely
        for docx_path, pdf_path in job_list:
            results.append((pdf_path, False, f"Word COM initialization failed: {e}"))
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass

    return results
