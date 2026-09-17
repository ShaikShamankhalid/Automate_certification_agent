"""
Filename sanitization and generation utilities.
"""

import re
from typing import Set

INVALID_CHARS_PATTERN = re.compile(r'[\\/*?:"<>|\x00-\x1f]')

def sanitize_component(text: str) -> str:
    """
    Sanitize text component for use in a file name:
    - Replaces invalid Windows filename characters
    - Normalizes multiple spaces/underscores
    - Strips leading/trailing spaces and dots
    """
    if text is None:
        return ""
    # Strip whitespace
    clean = str(text).strip()
    # Replace invalid chars with underscore
    clean = INVALID_CHARS_PATTERN.sub("_", clean)
    # Replace consecutive spaces or underscores with single underscore
    clean = re.sub(r'[\s_]+', '_', clean)
    # Strip leading/trailing underscores and dots
    clean = clean.strip('_.')
    return clean

def generate_safe_filename(reg_no: str, name: str, existing_names: Set[str] = None) -> str:
    """
    Generates a safe Windows filename for a certificate:
    Format: <RegNo>_<StudentName>.pdf
    Example: 261FB04001_Kishore_Sharma.pdf
    Handles duplicates by appending a sequence suffix if necessary.
    """
    clean_reg = sanitize_component(reg_no) or "UNKNOWN_REG"
    clean_name = sanitize_component(name) or "STUDENT"

    base_name = f"{clean_reg}_{clean_name}"
    # Truncate if filename is excessively long (Windows MAX_PATH safety)
    if len(base_name) > 120:
        base_name = base_name[:120].rstrip('_.')

    filename = f"{base_name}.pdf"

    if existing_names is not None:
        counter = 1
        candidate = filename
        while candidate.lower() in existing_names:
            candidate = f"{base_name}_{counter}.pdf"
            counter += 1
        existing_names.add(candidate.lower())
        return candidate

    return filename
