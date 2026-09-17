"""
Main orchestrator for the Automated Certificate Generation Agent.
Runs the complete autonomous workflow:
1. Detect input files (Excel + Template)
2. Read & validate student data
3. Analyze template layout
4. Prepare clean output directory
5. Execute mail merge to generate PDFs
6. Validate all generated certificates (integrity, text presence, page count)
7. Write generation_report.csv
8. Display console summary report
"""

import os
import sys
import argparse
import tempfile
import shutil
from typing import List, Dict, Any, Set

from .detector import detect_input_files
from .excel_reader import read_student_records
from .naming import generate_safe_filename
from .docx_engine import analyze_docx_structure, populate_docx_for_student, convert_docx_batch_to_pdf
from .pdf_engine import analyze_pdf_template, generate_pdf_for_student
from .image_engine import generate_pdf_from_image
from .validator import validate_all_certificates
from .reporter import write_generation_report, print_summary_banner

def run_pipeline(
    work_dir: str = ".",
    excel_file: str = None,
    template_file: str = None,
    output_dir: str = "certificates",
    verbose: bool = True
) -> bool:
    """
    Executes the automated certificate generation pipeline.
    Returns True if all certificates succeeded, False otherwise.
    """
    if verbose:
        print("\n" + "=" * 60)
        print("  AUTOMATED CERTIFICATE GENERATION AGENT")
        print("=" * 60)

    # Step 1 & 2: Detect input files
    if verbose:
        print("\n[Step 1/8] Detecting input files...")
    excel_path, template_path, template_type = detect_input_files(
        directory=work_dir,
        excel_override=excel_file,
        template_override=template_file
    )
    if verbose:
        print(f"  [+] Excel Data File:      {os.path.basename(excel_path)}")
        print(f"  [+] Certificate Template: {os.path.basename(template_path)} (Type: {template_type.upper()})")

    # Step 3 & 4: Read student data and validate
    if verbose:
        print("\n[Step 2/8] Reading and validating student data...")
    records, meta = read_student_records(excel_path)
    if verbose:
        print(f"  [+] Total student records found: {meta['valid_students']}")
        print(f"  [+] Detected Reg No column:     '{meta['reg_col']}'")
        print(f"  [+] Detected Name column:       '{meta['name_col']}'")
        if meta["duplicate_reg_nos"] > 0:
            print(f"  [!] Note: {meta['duplicate_reg_nos']} duplicate registration numbers detected.")

    # Step 5: Analyze certificate template
    if verbose:
        print("\n[Step 3/8] Analyzing certificate template...")
    layout_info = {}
    if template_type == "docx":
        layout_info = analyze_docx_structure(template_path)
        if verbose:
            print(f"  [+] DOCX Mode: {layout_info['mode']}")
            if layout_info["mode"] == "layout":
                print(f"  [+] Insertion Target Paragraph: P{layout_info['target_index']}")
                print(f"  [+] Matched Typography Font:    '{layout_info['recommended_font']}'")
    elif template_type == "pdf":
        layout_info = analyze_pdf_template(template_path)
        if verbose:
            print(f"  [+] PDF Mode: {layout_info['mode']}")
            print(f"  [+] Template Dimensions: {layout_info['width']} x {layout_info['height']} pt")
    elif template_type == "image":
        if verbose:
            print("  [+] Image Template Mode: Raster image overlay")

    # Step 6: Prepare clean output directory
    if verbose:
        print(f"\n[Step 4/8] Preparing output directory: '{output_dir}'...")
    abs_output_dir = os.path.abspath(output_dir)
    os.makedirs(abs_output_dir, exist_ok=True)

    # Clean previous generated PDFs to prevent mixing runs
    existing_pdfs = [os.path.join(abs_output_dir, f) for f in os.listdir(abs_output_dir) if f.lower().endswith(".pdf")]
    for p in existing_pdfs:
        try:
            os.remove(p)
        except Exception:
            pass

    # Step 7: Perform Mail Merge
    if verbose:
        print(f"\n[Step 5/8] Generating {len(records)} individualized certificates...")

    assigned_names: Set[str] = set()
    planned_jobs = []
    for student in records:
        out_filename = generate_safe_filename(student["reg_no"], student["name"], assigned_names)
        out_pdf_path = os.path.join(abs_output_dir, out_filename)
        planned_jobs.append({
            "student": student,
            "filename": out_filename,
            "pdf_path": out_pdf_path
        })

    report_entries = []

    if template_type == "docx":
        # Process via Word COM batch
        temp_dir = tempfile.mkdtemp(prefix="cert_docx_")
        docx_to_pdf_pairs = []
        try:
            for job in planned_jobs:
                temp_docx_name = f"doc_{job['student']['reg_no']}.docx"
                temp_docx_path = os.path.join(temp_dir, temp_docx_name)
                populate_docx_for_student(
                    template_path=template_path,
                    output_docx_path=temp_docx_path,
                    student=job["student"],
                    layout_info=layout_info
                )
                docx_to_pdf_pairs.append((temp_docx_path, job["pdf_path"]))

            if verbose:
                print(f"  [+] Exporting batch to vector PDF via Word COM engine...")
            conversion_results = convert_docx_batch_to_pdf(docx_to_pdf_pairs)

            for job, (_, success, err_msg) in zip(planned_jobs, conversion_results):
                status = "Success" if success else f"Failed: {err_msg}"
                report_entries.append({
                    "reg_no": job["student"]["reg_no"],
                    "name": job["student"]["name"],
                    "status": status,
                    "output_file": job["filename"] if success else ""
                })

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    elif template_type == "pdf":
        for job in planned_jobs:
            try:
                generate_pdf_for_student(
                    template_path=template_path,
                    output_pdf_path=job["pdf_path"],
                    student=job["student"],
                    layout_info=layout_info
                )
                report_entries.append({
                    "reg_no": job["student"]["reg_no"],
                    "name": job["student"]["name"],
                    "status": "Success",
                    "output_file": job["filename"]
                })
            except Exception as e:
                report_entries.append({
                    "reg_no": job["student"]["reg_no"],
                    "name": job["student"]["name"],
                    "status": f"Failed: {e}",
                    "output_file": ""
                })

    elif template_type == "image":
        for job in planned_jobs:
            try:
                generate_pdf_from_image(
                    template_path=template_path,
                    output_pdf_path=job["pdf_path"],
                    student=job["student"]
                )
                report_entries.append({
                    "reg_no": job["student"]["reg_no"],
                    "name": job["student"]["name"],
                    "status": "Success",
                    "output_file": job["filename"]
                })
            except Exception as e:
                report_entries.append({
                    "reg_no": job["student"]["reg_no"],
                    "name": job["student"]["name"],
                    "status": f"Failed: {e}",
                    "output_file": ""
                })

    # Step 8: Quality Control & Validation
    if verbose:
        print("\n[Step 6/8] Performing Quality Control and integrity checks...")

    successful_jobs = [j for j in planned_jobs if any(r["output_file"] == j["filename"] for r in report_entries)]
    qc_results = validate_all_certificates(successful_jobs, abs_output_dir)

    if verbose:
        print(f"  [+] Verified PDFs: {qc_results['passed']} passed, {qc_results['failed']} failed")
        if qc_results["previews"]:
            print(f"  [+] Generated {len(qc_results['previews'])} visual QC preview images in '.previews/'")

    # Step 9: Generation Report
    if verbose:
        print("\n[Step 7/8] Generating audit report...")
    report_csv = write_generation_report(report_entries, abs_output_dir)
    if verbose:
        print(f"  [+] Report written to: {report_csv}")

    # Step 10: Final Summary Display
    if verbose:
        print("\n[Step 8/8] Workflow Completed.")
        print_summary_banner(
            total_found=len(records),
            created_count=qc_results["passed"],
            failed_count=len(records) - qc_results["passed"],
            output_dir=output_dir,
            issues=qc_results.get("issues")
        )

    return qc_results["passed"] == len(records)

def main():
    parser = argparse.ArgumentParser(description="Automated Certificate Generation Agent")
    parser.add_argument("--dir", default=".", help="Working directory containing Excel and template files")
    parser.add_argument("--excel", default=None, help="Explicit path to Excel file")
    parser.add_argument("--template", default=None, help="Explicit path to certificate template")
    parser.add_argument("--output", default="certificates", help="Output directory for generated PDFs")

    args = parser.parse_args()
    success = run_pipeline(
        work_dir=args.dir,
        excel_file=args.excel,
        template_file=args.template,
        output_dir=args.output,
        verbose=True
    )
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
