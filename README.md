# Automated Certificate Generation Agent

A production-quality, autonomous mail merge agent that marries student records from Excel with certificate templates (Word `.docx`, `.pdf`, or raster images) to generate professional, individualized certificate PDFs.

---

## Architecture Overview

```
cerificationai/
├── certificate_agent/
│   ├── __init__.py           # Package definition
│   ├── main.py               # Orchestrator & CLI workflow
│   ├── detector.py           # Auto-detection of Excel & template files
│   ├── excel_reader.py       # Robust Excel parsing & column normalization
│   ├── naming.py             # Windows-safe filename generator & deduplication
│   ├── docx_engine.py        # DOCX layout analysis & Word COM batch PDF export
│   ├── pdf_engine.py         # PyMuPDF vector PDF template mail merge engine
│   ├── image_engine.py       # High-resolution image template engine
│   ├── validator.py          # Quality control, text verification & preview renders
│   └── reporter.py           # Generation audit reporter & summary printer
├── tests/
│   └── test_agent.py         # Automated test suite
├── main.py                   # Root entry point
├── requirements.txt          # Python dependencies
├── README.md                 # Complete documentation
└── certificates/             # Generated PDFs and generation_report.csv
    ├── 261FB04001_Kishore_Sharma.pdf
    ├── ...
    ├── generation_report.csv
    └── .previews/
```

---

## 1. Input Files

To run the agent, place your files in the working directory:

1. **Student Data File (Excel)**:
   - Format: `.xlsx` or `.xls`
   - Automatically detected without hardcoding.
   - Requires columns for Student Name and Registration Number.

2. **Certificate Sample/Template**:
   - Format: `.docx`, `.pdf`, or `.png`/`.jpg`
   - Automatically detected without hardcoding.
   - Design source of truth: background, borders, typography, page orientation, and dimensions are strictly preserved.

---

## 2. Dynamic Field Detection & Layout

### Explicit Placeholders
If the template uses explicit tags, they are automatically detected and replaced:
- Name: `{{NAME}}`, `[NAME]`, `<NAME>`, `{NAME}`
- Reg No: `{{REG_NO}}`, `[REG NO]`, `<REG NO>`, `{REG_NO}`

### Implicit Layout Analysis (No Placeholders Required)
If no explicit placeholders exist (as with raw certificate designs):
1. The agent locates preamble text (e.g., *"This is to certify that"*, *"presented to"*, *"awarded to"*).
2. The agent identifies the subsequent completion text (e.g., *"has Successfully completed..."*).
3. The dynamic slot is targeted between these two sections.
4. The student's Name and Registration Number are inserted with matching professional typography.

---

## 3. Font Sizing & Typography Preservation

- **Typography Style Matching**: The agent extracts surrounding fonts from the template (e.g., *Book Antiqua*, *Georgia*, *Times New Roman*) so that inserted credentials match the certificate's aesthetic.
- **Dynamic Font Scaling**:
  - `≤ 22 characters`: 24 pt bold
  - `23 - 30 characters`: 21 pt bold
  - `31 - 40 characters`: 18 pt bold
  - `> 40 characters`: 16 pt bold
- **No Overlap Guarantee**: Line spacing and margins are constrained to prevent text from overflowing to a second page or overlapping borders/signatures.

---

## 4. Excel Processing & Validation

The agent includes fuzzy header normalization to support different naming conventions:
- **Registration Number**: `Reg No`, `RegNo`, `REG NO`, `Registration No`, `Roll No`, `Student ID`, `ID`
- **Student Name**: `Name`, `Student Name`, `Candidate Name`, `Full Name`

Features:
- Skips completely blank rows.
- Converts numeric/float student IDs (e.g., `202601.0` → `202601`).
- Strips excess leading/trailing whitespace.
- Detects and logs duplicate registration numbers.
- Fails fast with clear error messages if required columns are missing.

---

## 5. Execution

### Installation
```bash
pip install -r requirements.txt
```

### Run from Current Working Directory
Simply run:
```bash
python main.py
```
or
```bash
python -m certificate_agent.main
```

### Optional Command-Line Arguments
```bash
python main.py --dir . --output certificates --excel student_details.xlsx --template certificate.docx
```

---

## 6. Output & Reports

Generated files are saved in the isolated `certificates/` directory:
- **Individual Certificates**: `<RegNo>_<StudentName>.pdf` (sanitized for Windows filename rules).
- **Audit Report**: `certificates/generation_report.csv`
  ```csv
  Reg No,Name,Status,Output File
  261FB04001,Kishore Sharma,Success,261FB04001_Kishore_Sharma.pdf
  261FB04002,Deva Reddy,Success,261FB04002_Deva_Reddy.pdf
  ...
  ```
- **Console Summary**:
  ```text
  ==================================================
  Certificate generation completed.
  ==================================================
  Students found:       24
  Certificates created: 24
  Failed:               0

  Output directory:
  D:\Mtech\Agentic AI\cerificationai\certificates
  ==================================================

  All certificates successfully generated.
  ```

---

## 7. Quality Control & Validation

Before marking completion, the agent performs automated QC:
1. **Integrity Check**: Verifies that every PDF opens without corruption and has a non-zero byte size.
2. **Page Count Check**: Ensures each certificate is strictly 1 page matching the template.
3. **Content Verification**: Verifies that the student's Name and Reg No are present in the PDF text layer.
4. **Visual Previews**: Automatically renders preview images for inspection in `certificates/.previews/`.

---

## 8. Running Automated Tests
```bash
python -m unittest discover -s tests -v
```
