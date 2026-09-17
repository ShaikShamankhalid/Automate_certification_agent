"""
Reporting and output summary module.
Writes generation_report.csv and prints formatted console summaries.
"""

import os
import csv
from typing import List, Dict, Any

def write_generation_report(report_entries: List[Dict[str, Any]], output_dir: str) -> str:
    """
    Writes certificates/generation_report.csv with columns:
    Reg No,Name,Status,Output File
    """
    os.makedirs(output_dir, exist_ok=True)
    report_file = os.path.join(output_dir, "generation_report.csv")

    fieldnames = ["Reg No", "Name", "Status", "Output File"]

    with open(report_file, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for entry in report_entries:
            writer.writerow({
                "Reg No": entry.get("reg_no", ""),
                "Name": entry.get("name", ""),
                "Status": entry.get("status", "Success"),
                "Output File": entry.get("output_file", "")
            })

    return report_file

def print_summary_banner(
    total_found: int,
    created_count: int,
    failed_count: int,
    output_dir: str,
    issues: List[Dict[str, Any]] = None
):
    """
    Prints the clean summary output matching required specification.
    """
    print("\n" + "=" * 50)
    print("Certificate generation completed.")
    print("=" * 50)
    print(f"Students found:       {total_found}")
    print(f"Certificates created: {created_count}")
    print(f"Failed:               {failed_count}")
    print("\nOutput directory:")
    print(f"{os.path.abspath(output_dir)}")
    print("=" * 50)

    if failed_count == 0:
        print("\nAll certificates successfully generated.\n")
    else:
        print(f"\nCompleted with {failed_count} failures:\n")
        if issues:
            for issue in issues[:10]:
                print(f"  - [{issue['reg_no']}] {issue['name']}: {', '.join(issue.get('errors', []))}")
            if len(issues) > 10:
                print(f"  ... and {len(issues) - 10} more.")
        print()
