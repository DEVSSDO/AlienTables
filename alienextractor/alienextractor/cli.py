"""
alienextractor.cli
~~~~~~~~~~~~~~~~

Interactive command-line interface for the AlienExtractor package.
"""

import sys
from pathlib import Path
from datetime import datetime

from . import engine


def main():
    """Entry point for the alienextractor CLI."""
    print("=========================================")
    print("              AlienExtractor")
    print("=========================================")
    print("Rules")
    print("1. Enter the exact column names that exist in your PDF template.")
    print("2. If incorrect column names are entered, the output may contain blank or random values.")
    print("3. All PDFs should follow the same layout/template.")
    print("4. If multiple PDF templates exist, process each template separately.")
    print("5. Same PDF columns should kept together to run this Tool")
    print("6. OCR will only run when required, exactly like the existing script.")
    print("7: make input folder in same folder and put all PDF files in it as by default folder")

    try:
        pdf_folder_str = input("Enter PDF Folder location: > ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\n\nCancelled.")
        return
    if not pdf_folder_str:
        pdf_folder_str = "input"
    pdf_folder = Path(pdf_folder_str)

    try:
        output_excel_str = input("Enter Output Excel File (optional): > ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\n\nCancelled.")
        return
    if output_excel_str:
        output_dir = Path(output_excel_str).parent
    else:
        output_dir = Path(".")

    requested_columns = []
    while True:
        try:
            col = input("Enter Column Name: > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n\nCancelled.")
            return
        if col.upper() == "OK":
            print("Type OK when finished.")
            break
        if col:
            requested_columns.append(col)

    if not requested_columns:
        print("No column names entered. Exiting.", file=sys.stderr)
        return

    # Run extraction
    engine.run(pdf_folder, output_dir, requested_columns)


if __name__ == "__main__":
    main()
