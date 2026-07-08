"""
alienextractor.parser
~~~~~~~~~~~~~~~~~~~

PDF reading and page processing.
Handles text extraction (PyMuPDF), OCR fallback, order block splitting,
and per-file processing. Logic preserved verbatim from the original script.
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple

import fitz
import pandas as pd

from .ocr import ocr_page, USE_OCR_FALLBACK
from .extractor import (
    RE_ORDER_HEADER, RE_ORDER_NO,
    parse_order_block_agg,
)
from .utils import _normalize_for_matching


# --------------- Audit file paths (written to CWD) ---------------
PAGE_AUDIT_CSV = Path("page_audit.csv")
ORPHAN_PAGES_CSV = Path("orphan_pages.csv")


# --------------- Page Splitting (verbatim from original) ---------------

def split_pages_into_order_blocks(pages: List[Dict]) -> List[Dict]:
    """
    Given pages as list of dicts {'text':..., 'ocr':bool},
    return list of order blocks with page numbers.
    """
    blocks = []
    current = None

    for idx, p in enumerate(pages, start=1):
        txt = p.get('text', '') or ''
        norm = _normalize_for_matching(txt)
        header_m = RE_ORDER_HEADER.search(norm)
        order_no = None

        if header_m:
            if header_m.lastindex and header_m.group(header_m.lastindex):
                order_no = header_m.group(header_m.lastindex).strip()
        if not order_no:
            mno = RE_ORDER_NO.search(norm)
            if mno:
                order_no = mno.group(1).strip()

        if header_m or order_no:
            if current:
                blocks.append(current)
            current = {
                'order_no': order_no or '',
                'pages_idx': [idx],
                'pages_text': [txt],
                'used_ocr': p.get('ocr', False)
            }
        else:
            if current:
                current['pages_idx'].append(idx)
                current['pages_text'].append(txt)
                current['used_ocr'] = current['used_ocr'] or p.get('ocr', False)
            else:
                current = {
                    'order_no': '',
                    'pages_idx': [idx],
                    'pages_text': [txt],
                    'used_ocr': p.get('ocr', False)
                }
    if current:
        blocks.append(current)
    return blocks


# --------------- Text Extraction (verbatim from original) ---------------

def extract_text_pages(pdf_path: Path) -> List[Dict]:
    """
    Return per-page dicts: {'text':..., 'ocr':bool} using PyMuPDF;
    fallback to OCR when needed.
    """
    doc = fitz.open(str(pdf_path))
    pages = []
    for i in range(len(doc)):
        page = doc.load_page(i)
        txt = page.get_text("text") or ""
        used_ocr = False
        if USE_OCR_FALLBACK:
            norm = _normalize_for_matching(txt)
            header = RE_ORDER_HEADER.search(norm) or RE_ORDER_NO.search(norm)
            if not norm.strip() or not header:
                ocr_txt = ocr_page(pdf_path, i + 1)
                if ocr_txt:
                    txt = (txt + "\n\n" + ocr_txt).strip()
                    used_ocr = True
        pages.append({'text': txt, 'ocr': used_ocr})
    return pages


# --------------- Per-File Processing (verbatim from original) ---------------

def process_pdf_file(pdf_path: Path) -> Tuple[List[Dict], List[str]]:
    """
    Process a single PDF file. Returns (rows, file_issues).
    Each row is a dict keyed by COLUMNS.
    """
    pages = extract_text_pages(pdf_path)

    # Audit pages
    page_audit = []
    for idx, p in enumerate(pages, start=1):
        norm = _normalize_for_matching(p.get('text', ''))
        has_header = bool(RE_ORDER_HEADER.search(norm) or RE_ORDER_NO.search(norm))
        page_audit.append({
            'Source File': pdf_path.name,
            'Page': idx,
            'Has_Text': bool(norm.strip()),
            'Has_Order_Header': has_header,
            'Used_OCR': bool(p.get('ocr', False))
        })

    # Write audit rows (append mode)
    try:
        df_audit = pd.DataFrame(page_audit)
        if PAGE_AUDIT_CSV.exists():
            df_audit.to_csv(PAGE_AUDIT_CSV, mode='a', index=False, header=False)
        else:
            df_audit.to_csv(PAGE_AUDIT_CSV, index=False)
    except Exception:
        pass

    blocks = split_pages_into_order_blocks(pages)
    rows = []
    file_issues = []
    orphan_pages = []

    for b in blocks:
        order_no = b.get('order_no', '')
        page_nos = b.get('pages_idx', [])
        block_text = "\n\n".join(b.get('pages_text', []))
        used_ocr = b.get('used_ocr', False)

        if not order_no:
            m = RE_ORDER_NO.search(_normalize_for_matching(block_text))
            if m:
                order_no = m.group(1).strip()

        if not order_no and all(not t.strip() for t in b.get('pages_text', [])):
            for p in page_nos:
                orphan_pages.append({
                    'Source File': pdf_path.name,
                    'Page': p,
                    'Reason': 'Blank page or unreadable'
                })
        elif not order_no:
            pass

        row = parse_order_block_agg(order_no, block_text, page_nos, pdf_path.name, used_ocr)
        rows.append(row)
        if row.get('Issues'):
            file_issues.append(
                f"{row.get('Order No.') or '(no-order)'}: {row.get('Issues')}"
            )

    # Save orphan pages
    try:
        if orphan_pages:
            df_orphan = pd.DataFrame(orphan_pages)
            if ORPHAN_PAGES_CSV.exists():
                df_orphan.to_csv(ORPHAN_PAGES_CSV, mode='a', index=False, header=False)
            else:
                df_orphan.to_csv(ORPHAN_PAGES_CSV, index=False)
    except Exception:
        pass

    return rows, file_issues
