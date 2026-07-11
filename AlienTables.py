#!/usr/bin/env python3
"""
AlienExtractor — interactive CLI version of the Daraz PDF extractor.
Preserves existing output columns and parsing logic, but adds:
- CLI prompts for input folder, output file/folder, and column names.
- Writes one Excel file per PDF (named same as the PDF).
- Auditing CSVs: Audit.csv, orphan_pages.csv, and review.csv for issues.
Dependencies: PyMuPDF, pandas, openpyxl, python-dateutil, tqdm, pdf2image, pillow, pytesseract.
"""
import re
import sys
import unicodedata
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple, Optional
import fitz
import pandas as pd
from dateutil import parser as dateparser
from tqdm import tqdm

# Optional OCR fallback
from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageFilter
import pytesseract

# --------------- CONFIG ---------------
# Default paths (will be overridden by user input)
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # update if needed
POPPLER_PATH = r"C:\poppler-25.12.0\Library\bin"               # update if needed
OCR_DPI = 300
USE_OCR_FALLBACK = True
PAGE_AUDIT_CSV = Path("Audit.csv")
ORPHAN_PAGES_CSV = Path("orphan_pages.csv")
REVIEW_CSV = Path("review.csv")

if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

# --------------- PARSING PATTERNS ---------------
RE_ORDER_HEADER = re.compile(
    r"(Daraz\s*(?:Purchase|Order)?\s*Summary|Purchase\s+Summary|Order\s+Summary)[\s\-–—:]*(?:Order\s*No\.?\s*[:\u00A0\s]*)?([0-9]{4,12})?",
    re.IGNORECASE | re.DOTALL
)
RE_ORDER_NO = re.compile(r'\b(?:Order\s*(?:No\.?|Number)|Order\s*#)\s*[:\u00A0\s]*([0-9]{4,12})', re.IGNORECASE)
RE_DATE = re.compile(r'([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}|\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4}|\d{4}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{1,2})')
RE_PHONE = re.compile(r'\b(92\d{9}|0\d{9,10}|\+?\d{11,13})\b')
RE_SHOP_SKU = re.compile(r'([0-9]{4,}_PK-[0-9]{4,})', re.IGNORECASE)
RE_PRICE = re.compile(r'([0-9]{1,3}(?:,[0-9]{3})*(?:\.\d{1,2})?)')
RE_TOTAL = re.compile(r'\bTotal\s*[:\-\u00A0]?\s*([0-9,\.]+)', re.IGNORECASE)
RE_SUBTOTAL = re.compile(r'\bSubtotal\s*[:\-\u00A0]?\s*([0-9,\.]+)', re.IGNORECASE)
RE_SHIPPING = re.compile(r'\bShipping Cost\s*[:\-\u00A0]?\s*([0-9,\.]+)', re.IGNORECASE)
RE_VOUCHER = re.compile(r'\bVoucher\s*[:\-\u00A0]?\s*[-]?([0-9,\.]+)', re.IGNORECASE)
RE_PAYMENT = re.compile(r'Payment Method\s*[:\-\u00A0]?\s*([A-Z0-9_ ]+)', re.IGNORECASE)

# --------------- UTILITY FUNCTIONS ---------------
def normalize_whitespace(s: str) -> str:
    if not s:
        return ""
    s = s.replace('\r', '\n')
    s = re.sub(r'\t+', ' ', s)
    s = re.sub(r' +', ' ', s)
    s = re.sub(r'\n{2,}', '\n\n', s)
    return s.strip()

def parse_currency_number(s: str) -> str:
    if not s:
        return ""
    return s.replace(',', '').strip()


COLUMN_ALIASES = {
    "order no": "Order No.",
    "order number": "Order No.",
    "order": "Order No.",
    "order num": "Order No.",
    "order id": "Order No.",
    "date": "Date",
    "order date": "Date",
    "bill to customer name": "Bill to Customer Name",
    "bill to name": "Bill to Customer Name",
    "bill customer name": "Bill to Customer Name",
    "bill name": "Bill to Customer Name",
    "bill to": "Bill to Customer Name",
    "billing name": "Bill to Customer Name",
    "customer name": "Bill to Customer Name",
    "bill to address": "Bill To Address",
    "bill address": "Bill To Address",
    "billing address": "Bill To Address",
    "address": "Bill To Address",
    "bill to phone": "Bill To Phone",
    "bill phone": "Bill To Phone",
    "billing phone": "Bill To Phone",
    "phone": "Bill To Phone",
    "phone number": "Bill To Phone",
    "deliver to customer name": "DELIVER to Customer Name",
    "deliver to name": "DELIVER to Customer Name",
    "deliver customer name": "DELIVER to Customer Name",
    "deliver name": "DELIVER to Customer Name",
    "deliver to": "DELIVER to Customer Name",
    "delivery name": "DELIVER to Customer Name",
    "shipping name": "DELIVER to Customer Name",
    "deliver to address": "DELIVER To Address",
    "deliver address": "DELIVER To Address",
    "delivery address": "DELIVER To Address",
    "shipping address": "DELIVER To Address",
    "deliver to phone": "DELIVER To Phone",
    "deliver phone": "DELIVER To Phone",
    "delivery phone": "DELIVER To Phone",
    "shipping phone": "DELIVER To Phone",
    "product name": "Product Name",
    "product": "Product Name",
    "item name": "Product Name",
    "item": "Product Name",
    "shop sku": "Shop SKU",
    "shopsku": "Shop SKU",
    "seller sku": "Seller SKU",
    "sellersku": "Seller SKU",
    "sku": "Seller SKU",
    "variant": "Variant / Attributes",
    "variant attributes": "Variant / Attributes",
    "attributes": "Variant / Attributes",
    "variations": "Variant / Attributes",
    "quantity": "Quantity",
    "qty": "Quantity",
    "unit price": "Unit Price",
    "price": "Unit Price",
    "unitprice": "Unit Price",
    "shipping cost": "Shipping Cost",
    "shipping": "Shipping Cost",
    "delivery cost": "Shipping Cost",
    "voucher": "Voucher",
    "discount": "Voucher",
    "coupon": "Voucher",
    "total": "Total",
    "grand total": "Total",
    "amount": "Total",
    "payment method": "Payment Method",
    "payment": "Payment Method",
    "pay method": "Payment Method",
    "source file": "Source File",
    "file": "Source File",
    "filename": "Source File",
    "page no": "Page No",
    "page": "Page No",
    "page number": "Page No",
    "ocr engine": "OCR Engine",
    "ocr": "OCR Engine",
    "extraction confidence": "Extraction Confidence",
    "confidence": "Extraction Confidence",
    "issues": "Issues",
    "parsed timestamp": "Parsed_Timestamp",
    "timestamp": "Parsed_Timestamp",
}

def normalize_column_name(name: str) -> str:
    if not name:
        return ""
    name = name.strip().lower()
    name = re.sub(r'[.:\/_\-#]', ' ', name)
    name = re.sub(r'\s+', ' ', name)
    return name.strip()

def resolve_column(user_input: str) -> Optional[str]:
    normalized = normalize_column_name(user_input)
    return COLUMN_ALIASES.get(normalized)

def search_column_in_text(text: str, col_name: str, all_cols: List[str] = None) -> str:
    if not text or not col_name:
        return ""
    norm_text = normalize_whitespace(text)
    col_escaped = re.escape(col_name)
    
    for use_boundary in [True, False]:
        boundary = r'\b' if use_boundary else ''
        pattern = re.compile(
            boundary + col_escaped + boundary + r'[\s\-–—:=:\u00A0]*([^\n]*)',
            re.IGNORECASE
        )
        for match in pattern.finditer(norm_text):
            val = match.group(1).strip()
            if val:
                val = re.sub(r'^[\s\-–—:=:\u00A0]+', '', val).strip()
                if val:
                    return val
            
            start_pos = match.end()
            remaining = norm_text[start_pos:].strip()
            if remaining:
                next_line = remaining.split('\n')[0].strip()
                if next_line:
                    is_other_col = False
                    if all_cols:
                        for other in all_cols:
                            if other.lower() != col_name.lower() and other.lower() in next_line.lower():
                                is_other_col = True
                                break
                    is_label = False
                    if ':' in next_line:
                        parts = next_line.split(':', 1)
                        if len(parts[0]) < 30 and not any(c.isdigit() for c in parts[0]):
                            is_label = True
                    
                    if not is_other_col and not is_label:
                        return next_line
    return ""

def extract_custom_column(block_text: str, col_name: str, standard_row: dict, all_requested_cols: list) -> str:
    internal_col = resolve_column(col_name)
    if internal_col and internal_col in standard_row:
        val = standard_row[internal_col]
        if val:
            return val
    return search_column_in_text(block_text, col_name, all_requested_cols)


def preprocess_image_for_ocr(img: Image.Image) -> Image.Image:
    try:
        img = img.convert("L")
        img = ImageOps.autocontrast(img)
        w, h = img.size
        if max(w, h) < 1200:
            scale = int(1200 / max(w, h) + 0.999)
            img = img.resize((w * scale, h * scale), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=3))
    except Exception:
        pass
    return img

def ocr_page(pdf_path: Path, page_num: int) -> str:
    """OCR a single page of the PDF (1-based page index)."""
    try:
        imgs = convert_from_path(str(pdf_path), dpi=OCR_DPI,
                                 first_page=page_num, last_page=page_num,
                                 poppler_path=POPPLER_PATH)
    except Exception:
        return ""
    if not imgs:
        return ""
    im = preprocess_image_for_ocr(imgs[0])
    return pytesseract.image_to_string(im)

def _normalize_for_matching(s: str) -> str:
    # Normalize unicode and collapse whitespace for regex matching
    if not s:
        return ""
    return unicodedata.normalize('NFKC', s)

# --------------- PARSING BUILDING BLOCKS ---------------
def split_pages_into_order_blocks(pages: List[Dict]) -> List[Dict]:
    """
    Group pages into order blocks based on detecting order headers.
    Each block corresponds to one order (possibly spanning pages).
    """
    blocks = []
    current = None
    for idx, p in enumerate(pages, start=1):
        txt = p.get('text','') or ''
        norm = _normalize_for_matching(txt)
        header_m = RE_ORDER_HEADER.search(norm)
        order_no = None
        if header_m:
            # Capture order number if present in header
            if header_m.lastindex and header_m.group(header_m.lastindex):
                order_no = header_m.group(header_m.lastindex).strip()
        if not order_no:
            # Fallback: find explicit "Order No" anywhere
            mno = RE_ORDER_NO.search(norm)
            if mno:
                order_no = mno.group(1).strip()
        if header_m or order_no:
            # Start a new block (if we were accumulating one, save it first)
            if current:
                blocks.append(current)
            current = {
                'order_no': order_no or '',
                'pages_idx': [idx],
                'pages_text': [txt],
                'used_ocr': p.get('ocr', False)
            }
        else:
            # Continuation of current block, or a new unknown block if none open
            if current:
                current['pages_idx'].append(idx)
                current['pages_text'].append(txt)
                current['used_ocr'] = current['used_ocr'] or p.get('ocr', False)
            else:
                # No current block: start an UNKNOWN block so pages aren't lost
                current = {
                    'order_no': '',
                    'pages_idx': [idx],
                    'pages_text': [txt],
                    'used_ocr': p.get('ocr', False)
                }
    if current:
        blocks.append(current)
    return blocks

def extract_bill_deliver(block_text: str) -> Dict[str, str]:
    """
    Extract BILL TO and DELIVER TO sections from the order block text.
    Returns dict with bill_name, bill_addr, bill_phone, deliver_name, deliver_addr, deliver_phone.
    """
    res = {"bill_name": "", "bill_addr": "", "bill_phone": "",
           "deliver_name": "", "deliver_addr": "", "deliver_phone": ""}
    text = block_text
    # Inline BILL TO and DELIVER TO on same line (rare layout)
    m_inline = re.search(r'BILL TO\s*[:\u00A0]?\s*([^\n\r]+?)\s+DELIVER TO\s*[:\u00A0]?\s*([^\n\r]+)',
                         text, re.IGNORECASE)
    if m_inline:
        res["bill_name"] = m_inline.group(1).strip()
        res["deliver_name"] = m_inline.group(2).strip()
    # Find ADDRESS fields for BILL and DELIVER
    addrs = list(re.finditer(r'ADDRESS\s*[:\u00A0]?', text, re.IGNORECASE))
    if addrs:
        # Helper to get chunk of address text up to next PHONE or marker
        def get_chunk(startidx):
            start = addrs[startidx].end()
            next_phone = re.search(r'\bPHONE\s*[:\u00A0]?', text[start:], re.IGNORECASE)
            next_marker = re.search(r'Your Ordered Items|Your Ordered Items:', text[start:], re.IGNORECASE)
            cutoff = len(text)
            if next_phone:
                cutoff = min(cutoff, start + next_phone.start())
            if next_marker:
                cutoff = min(cutoff, start + next_marker.start())
            cutoff = min(cutoff, start + 800)
            return normalize_whitespace(text[start:cutoff])
        if len(addrs) >= 1:
            res["bill_addr"] = get_chunk(0)
        if len(addrs) >= 2:
            res["deliver_addr"] = get_chunk(1)
    # Extract PHONE fields
    phones = re.findall(r'PHONE\s*[:\u00A0]?\s*([0-9\+\-\s]{7,})', text, re.IGNORECASE)
    if phones:
        # Clean whitespace from phone numbers
        res["bill_phone"] = re.sub(r'\s+', '', phones[0])
        if len(phones) > 1:
            res["deliver_phone"] = re.sub(r'\s+', '', phones[1])
    else:
        # Fallback: search for numeric patterns anywhere (Pakistan code or 0 leading)
        phs = RE_PHONE.findall(text)
        if phs:
            res["bill_phone"] = phs[0]
            if len(phs) > 1:
                res["deliver_phone"] = phs[1]
    return res

def extract_totals_and_payment(block_text: str) -> Dict[str, str]:
    """
    Extract subtotal, shipping, voucher, total, and payment method from the order block.
    """
    r = {"subtotal": "", "shipping": "", "voucher": "", "total": "", "payment_method": ""}
    m = RE_SUBTOTAL.search(block_text)
    if m:
        r["subtotal"] = parse_currency_number(m.group(1))
    m = RE_SHIPPING.search(block_text)
    if m:
        r["shipping"] = parse_currency_number(m.group(1))
    m = RE_VOUCHER.search(block_text)
    if m:
        r["voucher"] = parse_currency_number(m.group(1))
    m = RE_TOTAL.search(block_text)
    if m:
        r["total"] = parse_currency_number(m.group(1))
    m = RE_PAYMENT.search(block_text)
    if m:
        r["payment_method"] = m.group(1).strip()
    return r

def parse_products_from_block(block_text: str) -> List[Dict]:
    """
    Parse individual product lines from the order block.
    Returns a list of product dictionaries (name, SKUs, variant, qty, price).
    """
    products = []
    lines = [ln.strip() for ln in block_text.splitlines() if ln.strip()]
    n = len(lines)
    for i, ln in enumerate(lines):
        # Identify line containing Shop SKU (pattern)
        m_shop = RE_SHOP_SKU.search(ln)
        if m_shop:
            shop_sku = m_shop.group(1)
            name_parts = []
            # Look backwards for product name lines (up to 5 lines back, stopping at markers)
            j = i - 1
            while j >= 0 and len(name_parts) < 6:
                cand = lines[j]
                if re.search(r'\b(BILL TO|DELIVER TO|Subtotal|Total|Your Ordered Items)\b', cand, re.IGNORECASE):
                    break
                if RE_PRICE.findall(cand) and len(cand.split()) <= 3:
                    break
                name_parts.insert(0, cand)
                j -= 1
            product_name = " ".join(name_parts).strip()
            seller_sku = ""
            # Try to find a seller SKU (6+ chars alphanumeric) on same line (excluding shop_sku)
            seller_match = re.findall(r'\b[A-Z0-9\-]{6,}\b', ln)
            seller_match = [s for s in seller_match if s.upper() not in (shop_sku.upper(),)]
            if seller_match:
                seller_sku = seller_match[0]
            else:
                # Fallback: search a few lines below for seller SKU
                for k in range(i+1, min(n, i+4)):
                    cand2 = re.findall(r'\b[A-Z0-9\-]{6,}\b', lines[k])
                    if cand2:
                        seller_sku = cand2[0]
                        break
            # Capture variant attributes (Color, Size, etc.)
            variant = ""
            for k in range(i, min(n, i+6)):
                if 'Color' in lines[k] or 'Size' in lines[k] or 'Color family' in lines[k]:
                    variant += (" " + lines[k])
            variant = variant.strip()
            # Quantity, Unit Price, and Item Total extraction
            unit_price = ""
            item_total = ""
            qty = "1"
            for k in range(i, min(n, i+8)):
                nums = RE_PRICE.findall(lines[k])
                if nums:
                    item_total = parse_currency_number(nums[-1])
                    unit_price = parse_currency_number(nums[0])
                    break
            products.append({
                "Product Name": product_name,
                "Shop SKU": shop_sku,
                "Seller SKU": seller_sku,
                "Variant": variant,
                "Quantity": qty,
                "Unit Price": unit_price,
                "Item Total": item_total
            })
    # If no products found via SKU, attempt a fallback: look for any line with numbers (price)
    if not products:
        for i, ln in enumerate(lines):
            nums = RE_PRICE.findall(ln)
            if nums:
                prodname = lines[i-1] if i-1 >= 0 else ln
                products.append({
                    "Product Name": prodname,
                    "Shop SKU": "",
                    "Seller SKU": "",
                    "Variant": "",
                    "Quantity": "1",
                    "Unit Price": parse_currency_number(nums[0]),
                    "Item Total": parse_currency_number(nums[-1])
                })
    return products

# --------------- ORDER-LEVEL PARSER ---------------
def parse_order_block_agg(order_no: str, block_text: str, page_nos: List[int],
                          source_file: str, used_ocr: bool) -> Dict:
    """
    Given an order block text, extract all relevant fields and return a row dict.
    Aggregates all product names, SKUs, etc. into concatenated strings.
    """
    text = normalize_whitespace(block_text)
    date_val = ""
    dm = RE_DATE.search(text)
    if dm:
        try:
            date_val = dateparser.parse(dm.group(1), fuzzy=True).strftime("%Y-%m-%d")
        except Exception:
            date_val = dm.group(1).strip()
    totals = extract_totals_and_payment(text)
    bill = extract_bill_deliver(text)
    prods = parse_products_from_block(text)
    # Prepare lists for concatenation
    prod_names = []
    shop_skus = []
    seller_skus = []
    variants = []
    quantities = []
    unit_prices = []
    item_totals = []
    for p in prods:
        prod_names.append(p.get("Product Name", ""))
        shop_skus.append(p.get("Shop SKU", ""))
        seller_skus.append(p.get("Seller SKU", ""))
        variants.append(p.get("Variant", ""))
        quantities.append(str(p.get("Quantity", "1")))
        unit_prices.append(str(p.get("Unit Price", "")))
        item_totals.append(str(p.get("Item Total", "")))
    # Initialize output row with all original columns
    COLUMNS = [
        "Source File", "Page No", "Order No.", "Date",
        "Bill to Customer Name", "Bill To Address", "Bill To Phone",
        "DELIVER to Customer Name", "DELIVER To Address", "DELIVER To Phone",
        "Product Name", "Shop SKU", "Seller SKU", "Variant / Attributes",
        "Quantity", "Unit Price", "Shipping Cost", "Voucher", "Total",
        "Payment Method", "OCR Engine", "Extraction Confidence", "Issues", "Parsed_Timestamp"
    ]
    row = {c: "" for c in COLUMNS}
    row["Source File"] = source_file
    row["Page No"] = ";".join(map(str, page_nos)) if page_nos else ""
    row["Order No."] = order_no or ""
    row["Date"] = date_val
    row["Bill to Customer Name"] = bill.get("bill_name", "")
    row["Bill To Address"] = bill.get("bill_addr", "")
    row["Bill To Phone"] = bill.get("bill_phone", "")
    row["DELIVER to Customer Name"] = bill.get("deliver_name", "")
    row["DELIVER To Address"] = bill.get("deliver_addr", "")
    row["DELIVER To Phone"] = bill.get("deliver_phone", "")
    row["Product Name"] = " || ".join([p for p in prod_names if p])
    row["Shop SKU"] = ";".join([s for s in shop_skus if s])
    row["Seller SKU"] = ";".join([s for s in seller_skus if s])
    row["Variant / Attributes"] = ";".join([v for v in variants if v])
    row["Quantity"] = ";".join(quantities) if quantities else ""
    row["Unit Price"] = ";".join(unit_prices) if unit_prices else ""
    row["Shipping Cost"] = totals.get("shipping", "")
    row["Voucher"] = totals.get("voucher", "")
    row["Total"] = totals.get("total", "")
    row["Payment Method"] = totals.get("payment_method", "")
    row["OCR Engine"] = "Tesseract" if used_ocr else "PyMuPDF"
    # Estimate extraction confidence
    conf = 0.0
    if row["Order No."]:
        conf += 0.4
    if row["Product Name"]:
        conf += 0.35
    if row["Total"]:
        conf += 0.25
    row["Extraction Confidence"] = round(min(conf, 1.0), 2)
    issues = []
    if not row["Order No."]:
        issues.append("missing Order No.")
    if not row["Product Name"]:
        issues.append("missing Product Name")
    if not row["Total"]:
        issues.append("missing Total")
    row["Issues"] = "; ".join(issues)
    row["Parsed_Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return row

# --------------- PDF PAGE TEXT EXTRACTION ---------------
def extract_text_pages(pdf_path: Path) -> List[Dict]:
    """
    Extract text from each page of the PDF. Use PyMuPDF first;
    if a page has no text or no expected header, fall back to OCR for that page.
    Returns a list of {'text': text, 'ocr': bool} for each page.
    """
    doc = fitz.open(str(pdf_path))
    pages = []
    for i in range(len(doc)):
        page = doc.load_page(i)
        txt = page.get_text("text") or ""
        used_ocr = False
        if USE_OCR_FALLBACK:
            norm = _normalize_for_matching(txt)
            has_header = bool(RE_ORDER_HEADER.search(norm) or RE_ORDER_NO.search(norm))
            if not norm.strip() or not has_header:
                ocr_txt = ocr_page(pdf_path, i+1)
                if ocr_txt:
                    txt = (txt + "\n\n" + ocr_txt).strip()
                    used_ocr = True
        pages.append({'text': txt, 'ocr': used_ocr})
    return pages

def process_pdf_file(p: Path, user_columns: List[str]) -> Tuple[List[Dict], List[str]]:
    """
    Process one PDF file: extract pages, split into orders, parse each order.
    Returns a list of row dicts and a list of issues for review.
    """
    pages = extract_text_pages(p)
    # Audit: one row per physical page
    page_audit_rows = []
    for idx, pg in enumerate(pages, start=1):
        norm = _normalize_for_matching(pg.get('text',''))
        has_header = bool(RE_ORDER_HEADER.search(norm) or RE_ORDER_NO.search(norm))
        page_audit_rows.append({
            'Source File': p.name,
            'Page': idx,
            'Has_Text': bool(norm.strip()),
            'Has_Order_Header': has_header,
            'Used_OCR': pg.get('ocr', False)
        })
    # Append to Audit.csv
    try:
        df_audit = pd.DataFrame(page_audit_rows)
        if PAGE_AUDIT_CSV.exists():
            df_audit.to_csv(PAGE_AUDIT_CSV, mode='a', index=False, header=False)
        else:
            df_audit.to_csv(PAGE_AUDIT_CSV, index=False)
    except Exception:
        pass

    blocks = split_pages_into_order_blocks(pages)
    
    # Merge blocks by Order No. for this file to handle multi-page/multi-item orders
    file_orders = {}
    fallback_count = 0
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
            for pg in page_nos:
                orphan_pages.append({'Source File': p.name, 'Page': pg, 'Reason': 'Blank or unreadable'})
                
        row = parse_order_block_agg(order_no, block_text, page_nos, p.name, used_ocr)
        
        # Store block text in row temporarily for custom column extraction
        row["_block_text"] = block_text
        
        order_no_key = row.get("Order No.", "") or f"__UNKNOWN__{fallback_count}"
        if order_no_key.startswith("__UNKNOWN__"):
            fallback_count += 1
            
        if order_no_key in file_orders:
            existing = file_orders[order_no_key]
            
            def extend_field(field, sep=" || "):
                a = existing.get(field, "")
                b = row.get(field, "")
                if not b: return a
                if not a: return b
                items = a.split(sep)
                for it in b.split(sep):
                    if it and it not in items:
                        items.append(it)
                return sep.join(items)
                
            existing["Product Name"] = extend_field("Product Name", " || ")
            
            def extend_semicolon(field):
                a = existing.get(field, "")
                b = row.get(field, "")
                if not b: return a
                if not a: return b
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
            existing["_block_text"] = existing.get("_block_text", "") + "\n\n" + row.get("_block_text", "")
            
            for tcol in ("Shipping Cost", "Voucher", "Total", "Payment Method"):
                if row.get(tcol):
                    existing[tcol] = row.get(tcol)
            
            ex_issues = set(existing.get("Issues", "").split("; ")) if existing.get("Issues") else set()
            new_issues = set(row.get("Issues", "").split("; ")) if row.get("Issues") else set()
            combined = sorted(x for x in (ex_issues.union(new_issues)) if x)
            existing["Issues"] = "; ".join(combined)
            
            try:
                existing_conf = float(existing.get("Extraction Confidence") or 0)
                new_conf = float(row.get("Extraction Confidence") or 0)
                existing["Extraction Confidence"] = round(min(1.0, max(existing_conf, new_conf)), 2)
            except Exception:
                pass
            file_orders[order_no_key] = existing
        else:
            file_orders[order_no_key] = row

    # Now extract the custom columns for each merged order
    merged_rows = []
    for order_key, standard_row in file_orders.items():
        block_text = standard_row.get("_block_text", "")
        # Clean up temporary field
        standard_row.pop("_block_text", None)
        
        # Populate custom/requested columns
        for col_name in user_columns:
            standard_row[col_name] = extract_custom_column(block_text, col_name, standard_row, user_columns)
            
        merged_rows.append(standard_row)
        if standard_row.get('Issues'):
            file_issues.append(f"{standard_row.get('Order No.') or '(no-order)'}: {standard_row.get('Issues')}")

    # Append orphan pages to CSV
    try:
        if orphan_pages:
            df_orphan = pd.DataFrame(orphan_pages)
            if ORPHAN_PAGES_CSV.exists():
                df_orphan.to_csv(ORPHAN_PAGES_CSV, mode='a', index=False, header=False)
            else:
                df_orphan.to_csv(ORPHAN_PAGES_CSV, index=False)
    except Exception:
        pass

    return merged_rows, file_issues

def main():
    # Print banner and rules
    print("="*40)
    print("        AlienExtractor")
    print("="*40)
    print("Rules:")
    print("1. Enter the exact column names that exist in your PDF template.")
    print("2. If incorrect column names are entered, the output may contain blank or random values.")
    print("3. All PDFs should follow the same layout/template.")
    print("4. If multiple PDF templates exist, process each template separately.")
    print("5. Same PDF columns should be kept together to run this Tool.")
    print("6. OCR will only run when required, exactly like the existing script.")
    print("7. Make an 'input' folder and put all PDF files there (default).")
    print()

    # Get user inputs
    folder = input("Enter PDF Folder location (default 'input'): ").strip()
    if not folder:
        input_dir = Path("input")
    else:
        input_dir = Path(folder)
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Error: Folder '{input_dir}' not found.", file=sys.stderr)
        return

    out_entry = input("Enter Output Excel File (optional): ").strip()
    if out_entry:
        out_path = Path(out_entry)
        if out_path.suffix and out_path.suffix.lower() == ".xlsx":
            out_folder = None
            if out_path.is_dir():
                out_folder = out_path
            else:
                out_file = out_path
        else:
            # Treat as directory
            out_folder = out_path
            try:
                out_folder.mkdir(parents=True, exist_ok=True)
            except Exception:
                print(f"Warning: Could not create folder '{out_folder}', using input folder for output.")
                out_folder = None
    else:
        out_folder = None
        out_file = None

    print("\nEnter Column Names (one per line). Type 'OK' when finished:")
    user_columns = []
    while True:
        col = input("Enter Column Name: ").strip()
        if col.lower() == "ok":
            break
        if col:
            user_columns.append(col)

    if not user_columns:
        print("Error: No column names entered. Exiting.", file=sys.stderr)
        return

    # Initialize/clear audit and orphan files
    if PAGE_AUDIT_CSV.exists():
        PAGE_AUDIT_CSV.unlink()
    if ORPHAN_PAGES_CSV.exists():
        ORPHAN_PAGES_CSV.unlink()
    if REVIEW_CSV.exists():
        REVIEW_CSV.unlink()

    # Process each PDF
    files = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    if not files:
        print("No PDF files found in the input folder.", file=sys.stderr)
        return
    for f in tqdm(files, desc="Processing PDFs"):
        try:
            rows, file_issues = process_pdf_file(f, user_columns)
        except Exception as e:
            print(f"Error processing {f.name}: {e}", file=sys.stderr)
            file_issues = [str(e)]
            rows = []
        # Create DataFrame for this file's orders
        if rows:
            df = pd.DataFrame(rows)
            # Determine final output columns: always include Source File and Page No
            base_cols = ["Source File", "Page No"]
            extra_cols = ["OCR Engine", "Extraction Confidence", "Issues", "Parsed_Timestamp"]
            # Intersect user_columns with existing columns in df (some user-specified names might not exist)
            out_cols = base_cols + [col for col in user_columns] + extra_cols
            # Ensure DataFrame has all these columns (fill with blanks if missing)
            for col in out_cols:
                if col not in df.columns:
                    df[col] = ""
            df = df[out_cols]
            # Determine output path
            if 'out_file' in locals() and out_file and len(files) == 1:
                output_path = out_file
            else:
                # Default: same folder as PDF (or user-specified folder)
                if out_folder:
                    output_path = out_folder / f"{f.stem}.xlsx"
                else:
                    output_path = f.with_suffix(".xlsx")
            try:
                df.to_excel(output_path, index=False)
                print(f"Saved: {output_path.name} ({len(df)} orders)")
            except Exception as e:
                print(f"Error saving Excel for {f.name}: {e}", file=sys.stderr)
        else:
            print(f"No orders extracted from {f.name}.")
        # Collect review issues
        if file_issues:
            with_review = "; ".join(file_issues)
            # Append to review list
            if REVIEW_CSV.exists():
                pd.DataFrame([{"Source File": f.name, "Issues": with_review}]).to_csv(
                    REVIEW_CSV, mode='a', index=False, header=False)
            else:
                pd.DataFrame([{"Source File": f.name, "Issues": with_review}]).to_csv(
                    REVIEW_CSV, index=False)

if __name__ == "__main__":
    main()
