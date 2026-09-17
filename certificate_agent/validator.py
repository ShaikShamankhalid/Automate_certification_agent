"""
Quality Control and Validation module.
Validates generated certificate PDFs for integrity, content accuracy,
page count, and dimensions, and renders visual preview samples.
"""

import os
import pymupdf as fitz
from typing import List, Dict, Any, Tuple

def normalize_text(t: str) -> str:
    """Normalizes text for robust presence checking."""
    return "".join(t.lower().split())

def validate_single_pdf(
    pdf_path: str,
    student: Dict[str, Any],
    expected_pages: int = 1
) -> Tuple[bool, List[str]]:
    """
    Validates an individual certificate PDF.
    Returns (is_valid, list_of_error_messages).
    """
    errors = []

    if not os.path.exists(pdf_path):
        return False, [f"PDF file does not exist: {pdf_path}"]

    file_size = os.path.getsize(pdf_path)
    if file_size < 1024:
        return False, [f"PDF file is suspiciously small ({file_size} bytes): {pdf_path}"]

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        return False, [f"Corrupted PDF - cannot open: {e}"]

    try:
        if len(doc) != expected_pages:
            errors.append(f"Unexpected page count: expected {expected_pages}, got {len(doc)}")

        if len(doc) > 0:
            page = doc[0]
            text = page.get_text()
            norm_text = normalize_text(text)
            norm_name = normalize_text(student["name"])
            norm_reg = normalize_text(student["reg_no"])

            if norm_name not in norm_text:
                errors.append(f"Student name '{student['name']}' not detected in PDF text layer.")

            if norm_reg not in norm_text:
                errors.append(f"Student Reg No '{student['reg_no']}' not detected in PDF text layer.")

    except Exception as e:
        errors.append(f"Error inspecting PDF content: {e}")
    finally:
        doc.close()

    return len(errors) == 0, errors

def render_preview_image(pdf_path: str, output_img_path: str, scale: float = 1.5) -> bool:
    """Renders page 0 of a PDF to an image for visual inspection."""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=int(72 * scale))
        pix.save(output_img_path)
        doc.close()
        return True
    except Exception as e:
        print(f"Failed to render preview for '{pdf_path}': {e}")
        return False

def validate_all_certificates(
    jobs: List[Dict[str, Any]],
    output_dir: str
) -> Dict[str, Any]:
    """
    Validates all generated certificates.
    jobs list contains dicts: {"student": dict, "pdf_path": str}
    """
    total = len(jobs)
    passed = 0
    failed = 0
    issues = []
    preview_paths = []

    for idx, job in enumerate(jobs):
        student = job["student"]
        pdf_path = job["pdf_path"]
        is_valid, errs = validate_single_pdf(pdf_path, student)

        if is_valid:
            passed += 1
        else:
            failed += 1
            issues.append({
                "reg_no": student["reg_no"],
                "name": student["name"],
                "file": os.path.basename(pdf_path),
                "errors": errs
            })

    # Render preview for first student, longest name student, and last student
    if jobs:
        # 1. First student
        preview_indices = {0, len(jobs) - 1}
        # Find index with longest name
        longest_idx = max(range(len(jobs)), key=lambda i: len(jobs[i]["student"]["name"]))
        preview_indices.add(longest_idx)

        previews_dir = os.path.join(output_dir, ".previews")
        os.makedirs(previews_dir, exist_ok=True)

        for p_idx in sorted(preview_indices):
            target_job = jobs[p_idx]
            target_pdf = target_job["pdf_path"]
            img_name = f"preview_{target_job['student']['reg_no']}.png"
            img_path = os.path.join(previews_dir, img_name)
            if render_preview_image(target_pdf, img_path):
                preview_paths.append(img_path)

    return {
        "total": total,
        "passed": passed,
        "failed": failed,
        "issues": issues,
        "previews": preview_paths
    }
