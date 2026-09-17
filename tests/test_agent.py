"""
Unit tests for the Certificate Generation Agent modules.
"""

import os
import unittest
import tempfile
import shutil
import pandas as pd

from certificate_agent.naming import sanitize_component, generate_safe_filename
from certificate_agent.excel_reader import identify_columns, read_student_records
from certificate_agent.detector import is_temporary_or_system_file
from certificate_agent.docx_engine import calculate_name_font_size

class TestNaming(unittest.TestCase):
    def test_sanitize_component(self):
        self.assertEqual(sanitize_component("Kishore Sharma"), "Kishore_Sharma")
        self.assertEqual(sanitize_component("Phani/Kumar:Reddy?*"), "Phani_Kumar_Reddy")
        self.assertEqual(sanitize_component("   Sneha   Jhansi  "), "Sneha_Jhansi")
        self.assertEqual(sanitize_component(None), "")

    def test_generate_safe_filename_unique(self):
        seen = set()
        f1 = generate_safe_filename("202601", "Ravi Kumar", seen)
        f2 = generate_safe_filename("202601", "Ravi Kumar", seen)
        self.assertEqual(f1, "202601_Ravi_Kumar.pdf")
        self.assertEqual(f2, "202601_Ravi_Kumar_1.pdf")

class TestExcelReader(unittest.TestCase):
    def test_column_identification_standard(self):
        cols = ["REG NO", "NAME"]
        reg, name, _ = identify_columns(cols)
        self.assertEqual(reg, "REG NO")
        self.assertEqual(name, "NAME")

    def test_column_identification_variations(self):
        variations = [
            (["Registration No", "Student Name"], "Registration No", "Student Name"),
            (["RegNo", "FullName"], "RegNo", "FullName"),
            (["Roll No", "Candidate Name"], "Roll No", "Candidate Name"),
            (["student_id", "name"], "student_id", "name"),
        ]
        for cols, exp_reg, exp_name in variations:
            r, n, _ = identify_columns(cols)
            self.assertEqual(r, exp_reg)
            self.assertEqual(n, exp_name)

    def test_read_student_records(self):
        temp_dir = tempfile.mkdtemp()
        try:
            excel_path = os.path.join(temp_dir, "test.xlsx")
            df = pd.DataFrame({
                "Reg No": ["202601", "202602", None, "202603"],
                "Name": ["Alice", "Bob", None, "Charlie"]
            })
            df.to_excel(excel_path, index=False)
            records, meta = read_student_records(excel_path)
            self.assertEqual(len(records), 3)
            self.assertEqual(meta["skipped_blank_rows"], 1)
            self.assertEqual(records[0]["reg_no"], "202601")
            self.assertEqual(records[0]["name"], "Alice")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

class TestDocxTypography(unittest.TestCase):
    def test_dynamic_font_sizes(self):
        self.assertEqual(calculate_name_font_size("John Doe"), 24)
        self.assertEqual(calculate_name_font_size("A" * 28), 21)
        self.assertEqual(calculate_name_font_size("A" * 38), 18)
        self.assertEqual(calculate_name_font_size("A" * 50), 16)

class TestDetectorHelpers(unittest.TestCase):
    def test_temp_file_filter(self):
        self.assertTrue(is_temporary_or_system_file("~$certificate.docx"))
        self.assertTrue(is_temporary_or_system_file(".git"))
        self.assertFalse(is_temporary_or_system_file("certificate.docx"))
        self.assertFalse(is_temporary_or_system_file("student_details.xlsx"))

if __name__ == "__main__":
    unittest.main()
