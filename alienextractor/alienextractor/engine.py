"""
alienextractor.engine
~~~~~~~~~~~~~~~~~~~

Main workflow controller. Orchestrates PDF discovery, per-file processing,
order aggregation, and final Excel output.
"""

import re
import sys
from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd
from tqdm import tqdm

from .parser import process_pdf_file
from .utils import extract_custom_column


def run(
    pdf_folder: Path,
    output_dir: Path,
    requested_columns: List[str],
) -> Optional[Path]:
    """
    Main extraction workflow.

    Args:
        pdf_folder: Path to the directory containing PDF files.
        output_dir: Parent directory where Excel files should be saved.
        requested_columns: List of column names as entered by the user.

    Returns:
        Path to the parent output directory, or None if no files were processed.
    """
    if not pdf_folder.exists():
        print(f"\n  [ERROR] Folder '{pdf_folder}' does not exist.", file=sys.stderr)
        print(f"  Please create the folder and place your PDF files inside it.\n",
              file=sys.stderr)
        return None

    # Discover and sort PDF files (numeric sort like original)
    def numeric_key(path: Path):
        m = re.search(r'(\d+)', path.stem)
        return int(m.group(1)) if m else 0

    files = [p for p in pdf_folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    files = sorted(files, key=numeric_key)

    if not files:
        print(f"\n  [ERROR] No PDF files found in '{pdf_folder}'.", file=sys.stderr)
        return None

    print(f"\n  Found {len(files)} PDF file(s) in '{pdf_folder}'")
    print(f"  Extracting {len(requested_columns)} column(s): {', '.join(requested_columns)}")
    print()

    per_file_issues = {}

    for f in tqdm(files, desc="  Processing"):
        try:
            rows, file_issues = process_pdf_file(f)

            # Merge rows by Order No. for this file
            file_orders = {}
            fallback_count = 0
            for r in rows:
                order_no_key = r.get("Order No.", "") or f"__UNKNOWN__{fallback_count}"
                if order_no_key.startswith("__UNKNOWN__"):
                    fallback_count += 1

                if order_no_key in file_orders:
                    existing = file_orders[order_no_key]

                    def extend_field(field, sep=" || "):
                        a = existing.get(field, "")
                        b = r.get(field, "")
                        if not b:
                            return a
                        if not a:
                            return b
                        items = a.split(sep)
                        for it in b.split(sep):
                            if it and it not in items:
                                items.append(it)
                        return sep.join(items)

                    existing["Product Name"] = extend_field("Product Name", " || ")

                    def extend_semicolon(field):
                        a = existing.get(field, "")
                        b = r.get(field, "")
                        if not b:
                            return a
                        if not a:
                            return b
                        aset = [x for x in a.split(";") if x]
                        for it in b.split(";"):
                            if it and it not in aset:
                                aset.append(it)
                        return ";".join(aset)

                    existing["Shop SKU"] = extend_semicolon("Shop SKU")
                    existing["Seller SKU"] = extend_semicolon("Seller SKU")
                    existing["Variant / Attributes"] = extend_semicolon("Variant / Attributes")
                    existing["Quantity"] = extend_semicolon("Quantity")
                    existing["Unit Price"] = extend_semicolon("Unit Price")
                    existing["_block_text"] = existing.get("_block_text", "") + "\n\n" + r.get("_block_text", "")

                    for tcol in ("Shipping Cost", "Voucher", "Total", "Payment Method"):
                        if r.get(tcol):
                            existing[tcol] = r.get(tcol)

                    ex_issues = (
                        set(existing.get("Issues", "").split("; "))
                        if existing.get("Issues") else set()
                    )
                    new_issues = (
                        set(r.get("Issues", "").split("; "))
                        if r.get("Issues") else set()
                    )
                    combined = sorted(x for x in (ex_issues.union(new_issues)) if x)
                    existing["Issues"] = "; ".join(combined)

                    try:
                        existing_conf = float(existing.get("Extraction Confidence") or 0)
                        new_conf = float(r.get("Extraction Confidence") or 0)
                        existing["Extraction Confidence"] = round(
                            min(1.0, max(existing_conf, new_conf)), 2
                        )
                    except Exception:
                        pass

                    file_orders[order_no_key] = existing
                else:
                    file_orders[order_no_key] = r

            # Extract user's requested columns for each merged order of this file
            output_rows = []
            for order_key, standard_row in file_orders.items():
                out_row = {}
                block_text = standard_row.get("_block_text", "")
                for col_name in requested_columns:
                    out_row[col_name] = extract_custom_column(block_text, col_name, standard_row, requested_columns)
                output_rows.append(out_row)

            # Write the output Excel file for this file
            if output_rows:
                df = pd.DataFrame(output_rows, columns=requested_columns)
                output_xlsx = output_dir / f"{f.stem}.xlsx"
                output_xlsx.parent.mkdir(parents=True, exist_ok=True)
                df.to_excel(output_xlsx, index=False)

            if file_issues:
                per_file_issues[f.name] = file_issues

        except Exception as e:
            print(f"\n  Error processing {f.name}: {e}", file=sys.stderr)
            per_file_issues[f.name] = [str(e)]

    # Write review_needed.csv if there were issues
    if per_file_issues:
        review_rows = []
        for fname, issues in per_file_issues.items():
            review_rows.append({"Source File": fname, "Issues": "; ".join(issues)})
        review_path = output_dir / "review_needed.csv"
        pd.DataFrame(review_rows).to_csv(str(review_path), index=False)
        print(f"\n  Wrote review_needed.csv for {len(per_file_issues)} file(s) with issues")

    return output_dir
