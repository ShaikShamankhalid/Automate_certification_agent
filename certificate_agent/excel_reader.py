"""
Excel reading and validation module.
Reads student details, normalizes column headers, and validates records.
"""

import os
import re
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd

REG_PATTERNS = [
    r"^reg(istration)?[\s_\.\-]*no\.?$",
    r"^reg[\s_\.\-]*num(ber)?$",
    r"^regno$",
    r"^roll[\s_\.\-]*no\.?$",
    r"^roll[\s_\.\-]*num(ber)?$",
    r"^student[\s_\.\-]*id$",
    r"^id[\s_\.\-]*no\.?$",
    r"^ht[\s_\.\-]*no\.?$",
    r"^hall[\s_\.\-]*ticket[\s_\.\-]*no\.?$",
    r"^enrollment[\s_\.\-]*no\.?$",
    r"^id$"
]

NAME_PATTERNS = [
    r"^student[\s_\.\-]*name$",
    r"^candidate[\s_\.\-]*name$",
    r"^full[\s_\.\-]*name$",
    r"^participant[\s_\.\-]*name$",
    r"^name[\s_\.\-]*of[\s_\.\-]*the[\s_\.\-]*student$",
    r"^name$"
]

def _normalize_header(header: str) -> str:
    """Normalizes header string for comparison: lowercase, stripped."""
    if not isinstance(header, str):
        header = str(header)
    return header.strip().lower()

def identify_columns(columns: List[str]) -> Tuple[Optional[str], Optional[str], Dict[str, str]]:
    """
    Identifies 'Reg No' and 'Name' columns based on regex matching.
    Returns:
        (reg_col_name, name_col_name, {all_standardized_mapping})
    """
    reg_col = None
    name_col = None
    mapping = {}

    for col in columns:
        norm = _normalize_header(col)
        
        # Check Reg No
        if not reg_col:
            for pattern in REG_PATTERNS:
                if re.search(pattern, norm):
                    reg_col = col
                    mapping["reg_no"] = col
                    break

        # Check Name
        if not name_col:
            for pattern in NAME_PATTERNS:
                if re.search(pattern, norm):
                    name_col = col
                    mapping["name"] = col
                    break

    return reg_col, name_col, mapping

def clean_cell_value(val: Any) -> str:
    """Cleans an individual cell value, formatting floats properly."""
    if pd.isna(val) or val is None:
        return ""
    # If float represents an exact integer (e.g., 202601.0), convert to int string
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    if isinstance(val, int):
        return str(val)
    return str(val).strip()

def read_student_records(excel_path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Reads an Excel file and returns a list of valid student records and metadata.
    
    Returns:
        (records, metadata)
        records: List of dicts with keys:
            - 'reg_no': str
            - 'name': str
            - 'row_index': int (1-based row in Excel)
            - 'extra_fields': dict of any other columns present
        metadata: Dict with info (total_rows, valid_rows, skipped_rows, duplicates)
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: '{excel_path}'")

    try:
        df = pd.read_excel(excel_path, dtype=object)
    except Exception as e:
        raise ValueError(f"Failed to read Excel file '{excel_path}': {e}") from e

    reg_col, name_col, mapping = identify_columns(df.columns.tolist())

    missing = []
    if not reg_col:
        missing.append("Registration Number ('Reg No' / 'RegNo' / 'Registration No' / 'Roll No')")
    if not name_col:
        missing.append("Student Name ('Name' / 'Student Name')")

    if missing:
        found_cols = ", ".join([f"'{c}'" for c in df.columns])
        raise ValueError(
            f"Excel validation error in '{os.path.basename(excel_path)}':\n"
            f"Required columns missing: {', '.join(missing)}.\n"
            f"Columns found in file: [{found_cols}]."
        )

    records = []
    seen_reg_nos = set()
    duplicate_count = 0
    skipped_rows = 0

    for idx, row in df.iterrows():
        excel_row_num = idx + 2 # Header is row 1
        reg_val = clean_cell_value(row[reg_col])
        name_val = clean_cell_value(row[name_col])

        # Skip completely blank rows
        if not reg_val and not name_val:
            skipped_rows += 1
            continue

        if not reg_val or not name_val:
            # Partial row (one field missing)
            raise ValueError(
                f"Row {excel_row_num} in '{os.path.basename(excel_path)}' is missing a required field: "
                f"Reg No='{reg_val}', Name='{name_val}'."
            )

        # Collect any extra columns
        extra_fields = {}
        for col in df.columns:
            if col not in (reg_col, name_col):
                extra_fields[str(col).strip()] = clean_cell_value(row[col])

        if reg_val in seen_reg_nos:
            duplicate_count += 1
        seen_reg_nos.add(reg_val)

        records.append({
            "reg_no": reg_val,
            "name": name_val,
            "row_index": excel_row_num,
            "extra_fields": extra_fields
        })

    if not records:
        raise ValueError(f"No valid student records found in '{os.path.basename(excel_path)}'.")

    metadata = {
        "total_rows_in_file": len(df),
        "valid_students": len(records),
        "skipped_blank_rows": skipped_rows,
        "duplicate_reg_nos": duplicate_count,
        "reg_col": reg_col,
        "name_col": name_col
    }

    return records, metadata
