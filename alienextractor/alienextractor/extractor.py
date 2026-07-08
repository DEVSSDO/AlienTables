"""
alienextractor.extractor
~~~~~~~~~~~~~~~~~~~~~~

All extraction rules, regex patterns, and parsing logic.
This module contains the core business logic — preserved verbatim from the original script.
"""

import re
from datetime import datetime
from typing import List, Dict

from dateutil import parser as dateparser

from .utils import normalize_whitespace, parse_currency_number


# --------------- COLUMNS (unchanged from original) ---------------

COLUMNS = [
    "Source File", "Page No", "Order No.", "Date",
    "Bill to Customer Name", "Bill To Address", "Bill To Phone",
    "DELIVER to Customer Name", "DELIVER To Address", "DELIVER To Phone",
    "Product Name", "Shop SKU", "Seller SKU", "Variant / Attributes",
    "Quantity", "Unit Price", "Shipping Cost", "Voucher", "Total", "Payment Method",
    "OCR Engine", "Extraction Confidence", "Issues", "Parsed_Timestamp"
]


# --------------- REGEX PATTERNS (verbatim from original) ---------------

RE_ORDER_HEADER = re.compile(
    r"(Daraz\s*(?:Purchase|Order)?\s*Summary|Purchase\s+Summary|Order\s+Summary)"
    r"[\s\-–—:]*"
    r"(?:Order\s*No\.?\s*[:\u00A0\s]*)?"
    r"([0-9]{4,12})?",
    re.IGNORECASE | re.DOTALL
)

RE_ORDER_NO = re.compile(
    r'\b(?:Order\s*(?:No\.?|Number)|Order\s*#)\s*[:\u00A0\s]*([0-9]{4,12})',
    re.IGNORECASE
)

RE_DATE = re.compile(
    r'([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}|\d{1,2}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{2,4}|\d{4}[\/\-\.\s]\d{1,2}[\/\-\.\s]\d{1,2})'
)

RE_PHONE = re.compile(r'\b(92\d{9}|0\d{9,10}|\+?\d{11,13})\b')

RE_SHOP_SKU = re.compile(r'([0-9]{4,}_PK-[0-9]{4,})', re.IGNORECASE)

RE_PRICE = re.compile(r'([0-9]{1,3}(?:,[0-9]{3})*(?:\.\d{1,2})?)')

RE_TOTAL = re.compile(r'\bTotal\s*[:\-\u00A0]?\s*([0-9,\.]+)', re.IGNORECASE)

RE_SUBTOTAL = re.compile(r'\bSubtotal\s*[:\-\u00A0]?\s*([0-9,\.]+)', re.IGNORECASE)

RE_SHIPPING = re.compile(r'\bShipping Cost\s*[:\-\u00A0]?\s*([0-9,\.]+)', re.IGNORECASE)

RE_VOUCHER = re.compile(r'\bVoucher\s*[:\-\u00A0]?\s*[-]?([0-9,\.]+)', re.IGNORECASE)

RE_PAYMENT = re.compile(r'Payment Method\s*[:\-\u00A0]?\s*([A-Z0-9_ ]+)', re.IGNORECASE)


# --------------- Extraction Functions (verbatim from original) ---------------

def extract_bill_deliver(block_text: str) -> Dict[str, str]:
    """Extract bill-to and deliver-to information from an order block."""
    res = {
        "bill_name": "", "bill_addr": "", "bill_phone": "",
        "deliver_name": "", "deliver_addr": "", "deliver_phone": ""
    }
    text = block_text
    m_inline = re.search(
        r'BILL TO\s*[:\u00A0]?\s*([^\n\r]+?)\s+DELIVER TO\s*[:\u00A0]?\s*([^\n\r]+)',
        text, re.IGNORECASE
    )
    if m_inline:
        res["bill_name"] = m_inline.group(1).strip()
        res["deliver_name"] = m_inline.group(2).strip()

    addrs = list(re.finditer(r'ADDRESS\s*[:\u00A0]?', text, re.IGNORECASE))
    if addrs:
        def get_chunk(startidx):
            start = addrs[startidx].end()
            next_phone = re.search(r'\bPHONE\s*[:\u00A0]?', text[start:], re.IGNORECASE)
            next_marker = re.search(
                r'Your Ordered Items|Your Ordered Items:|Your Ordered Items',
                text[start:], re.IGNORECASE
            )
            cutoff = len(text)
            if next_phone:
                cutoff = min(cutoff, start + next_phone.start())
            if next_marker:
                cutoff = min(cutoff, start + next_marker.start())
            cutoff = min(cutoff, start + 800)
            return normalize_whitespace(text[start:cutoff])

        res["bill_addr"] = get_chunk(0) if len(addrs) >= 1 else ""
        if len(addrs) >= 2:
            res["deliver_addr"] = get_chunk(1)

    phones = re.findall(r'PHONE\s*[:\u00A0]?\s*([0-9\+\-\s]{7,})', text, re.IGNORECASE)
    if phones:
        res["bill_phone"] = re.sub(r'\s+', '', phones[0])
        if len(phones) > 1:
            res["deliver_phone"] = re.sub(r'\s+', '', phones[1])
    else:
        phs = RE_PHONE.findall(text)
        if phs:
            res["bill_phone"] = phs[0]
            if len(phs) > 1:
                res["deliver_phone"] = phs[1]
    return res


def extract_totals_and_payment(block_text: str) -> Dict[str, str]:
    """Extract subtotal, shipping, voucher, total, and payment method."""
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
    """Parse product information from an order block."""
    products = []
    lines = [ln.strip() for ln in block_text.splitlines() if ln.strip()]
    n = len(lines)
    for i, ln in enumerate(lines):
        m_shop = RE_SHOP_SKU.search(ln)
        if m_shop:
            shop_sku = m_shop.group(1)
            name_parts = []
            j = i - 1
            while j >= 0 and len(name_parts) < 6:
                cand = lines[j]
                if re.search(
                    r'\b(BILL TO|DELIVER TO|Subtotal|Total|Your Ordered Items)\b',
                    cand, re.IGNORECASE
                ):
                    break
                if RE_PRICE.findall(cand) and len(cand.split()) <= 3:
                    break
                name_parts.insert(0, cand)
                j -= 1
            product_name = " ".join(name_parts).strip()

            seller_sku = ""
            seller_match = re.findall(r'\b[A-Z0-9\-]{6,}\b', ln)
            seller_match = [s for s in seller_match if s.upper() not in (shop_sku.upper(),)]
            if seller_match:
                seller_sku = seller_match[0]
            else:
                for k in range(i + 1, min(n, i + 4)):
                    cand2 = re.findall(r'\b[A-Z0-9\-]{6,}\b', lines[k])
                    if cand2:
                        seller_sku = cand2[0]
                        break

            variant = ""
            for k in range(i, min(n, i + 6)):
                if 'Color' in lines[k] or 'Size' in lines[k] or 'Color family' in lines[k]:
                    variant += (" " + lines[k])
            variant = variant.strip()

            unit_price = ""
            item_total = ""
            qty = "1"
            for k in range(i, min(n, i + 8)):
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

    if not products:
        for i, ln in enumerate(lines):
            nums = RE_PRICE.findall(ln)
            if nums:
                prodname = lines[i - 1] if i - 1 >= 0 else ln
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


# --------------- Order-Level Parser (aggregates products) ---------------

def parse_order_block_agg(
    order_no: str,
    block_text: str,
    page_nos: List[int],
    source_file: str,
    used_ocr: bool
) -> Dict:
    """Parse a full order block and aggregate product data into a single row."""
    text = normalize_whitespace(block_text)

    # Extract date
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
    row["_block_text"] = block_text
    return row
