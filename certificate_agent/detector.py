"""
Automatic file detection module.
Locates the student Excel spreadsheet and the certificate sample/template file
in the target directory without hard-coded filenames.
"""

import os
import glob
from typing import Tuple, Optional, List
import pandas as pd

from .excel_reader import identify_columns

SUPPORTED_TEMPLATE_EXTS = {
    ".docx": "docx",
    ".pdf": "pdf",
    ".png": "image",
    ".jpg": "image",
    ".jpeg": "image"
}

EXCEL_EXTS = {".xlsx", ".xls"}

def is_temporary_or_system_file(filename: str) -> bool:
    """Checks if a file is an office temp file or system file."""
    base = os.path.basename(filename)
    if base.startswith("~$") or base.startswith(".") or base.startswith("temp_"):
        return True
    return False

def find_excel_file(directory: str = ".") -> str:
    """
    Finds the student Excel spreadsheet in the directory.
    If multiple Excel files exist, inspects columns to find the one with student data.
    """
    all_files = os.listdir(directory)
    excel_candidates = [
        os.path.join(directory, f) for f in all_files
        if os.path.splitext(f)[1].lower() in EXCEL_EXTS and not is_temporary_or_system_file(f)
    ]

    if not excel_candidates:
        raise FileNotFoundError(
            f"No Excel file (.xlsx, .xls) found in directory '{os.path.abspath(directory)}'."
        )

    if len(excel_candidates) == 1:
        return excel_candidates[0]

    # Multiple Excel files: score by column matches
    scored = []
    for candidate in excel_candidates:
        try:
            df = pd.read_excel(candidate, nrows=2)
            reg_col, name_col, _ = identify_columns(df.columns.tolist())
            score = 0
            if reg_col:
                score += 2
            if name_col:
                score += 2
            if "student" in os.path.basename(candidate).lower():
                score += 1
            scored.append((score, candidate))
        except Exception:
            scored.append((0, candidate))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_candidate = scored[0]

    if best_score >= 4:
        return best_candidate

    # Fallback to the first one if scoring couldn't differentiate
    return best_candidate

def find_template_file(directory: str = ".") -> Tuple[str, str]:
    """
    Finds the certificate template file in the directory.
    Returns:
        (template_path, template_type)
        where template_type is one of: 'docx', 'pdf', 'image'
    """
    all_files = os.listdir(directory)
    candidates = []

    for f in all_files:
        if is_temporary_or_system_file(f):
            continue
        # Avoid picking up files generated inside certificates/ or reports
        if f.lower() in ("generation_report.csv", "sample_template.pdf", "test_output.pdf"):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext in SUPPORTED_TEMPLATE_EXTS:
            candidates.append((os.path.join(directory, f), SUPPORTED_TEMPLATE_EXTS[ext]))

    if not candidates:
        raise FileNotFoundError(
            f"No certificate template (.docx, .pdf, .png, .jpg) found in directory '{os.path.abspath(directory)}'."
        )

    if len(candidates) == 1:
        return candidates[0]

    # Multiple candidates: score based on name relevance
    def score_template(item: Tuple[str, str]) -> int:
        path, type_ = item
        name = os.path.basename(path).lower()
        score = 0
        if "certificate" in name:
            score += 10
        if "template" in name:
            score += 8
        if "sample" in name:
            score += 6
        if type_ == "docx":
            score += 2
        elif type_ == "pdf":
            score += 3
        return score

    candidates.sort(key=score_template, reverse=True)
    return candidates[0]

def detect_input_files(
    directory: str = ".",
    excel_override: Optional[str] = None,
    template_override: Optional[str] = None
) -> Tuple[str, str, str]:
    """
    Detects both the Excel student file and the certificate template.
    Returns:
        (excel_path, template_path, template_type)
    """
    if excel_override:
        if not os.path.exists(excel_override):
            raise FileNotFoundError(f"Specified Excel file does not exist: '{excel_override}'")
        excel_path = excel_override
    else:
        excel_path = find_excel_file(directory)

    if template_override:
        if not os.path.exists(template_override):
            raise FileNotFoundError(f"Specified template file does not exist: '{template_override}'")
        ext = os.path.splitext(template_override)[1].lower()
        if ext not in SUPPORTED_TEMPLATE_EXTS:
            raise ValueError(f"Unsupported template format: '{ext}'. Supported: .docx, .pdf, .png, .jpg")
        template_path = template_override
        template_type = SUPPORTED_TEMPLATE_EXTS[ext]
    else:
        template_path, template_type = find_template_file(directory)

    return excel_path, template_path, template_type
