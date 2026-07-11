"""
alienextractor.utils
~~~~~~~~~~~~~~~~~~

Shared helper functions and column name normalization/alias mapping.
"""

import re
import unicodedata
from typing import Optional, List


# --------------- Text Helpers (verbatim from original) ---------------

def normalize_whitespace(s: str) -> str:
    """Collapse whitespace and normalize line endings."""
    if not s:
        return ""
    s = s.replace('\r', '\n')
    s = re.sub(r'\t+', ' ', s)
    s = re.sub(r' +', ' ', s)
    s = re.sub(r'\n{2,}', '\n\n', s)
    return s.strip()


def parse_currency_number(s: str) -> str:
    """Strip commas from currency strings."""
    if not s:
        return ""
    return s.replace(',', '').strip()


def _normalize_for_matching(s: str) -> str:
    """Normalize unicode and collapse whitespace to improve regex matching."""
    if not s:
        return ""
    return unicodedata.normalize('NFKC', s)


# --------------- Column Name Normalization ---------------

def normalize_column_name(name: str) -> str:
    """
    Normalize a user-entered column name for fuzzy matching.

    Strips whitespace, lowercases, collapses multiple spaces,
    and removes common punctuation (periods, colons, slashes, etc.).
    """
    if not name:
        return ""
    # Strip leading/trailing whitespace
    name = name.strip()
    # Lowercase
    name = name.lower()
    # Remove common punctuation: . : / - _ #
    name = re.sub(r'[.:\/_\-#]', ' ', name)
    # Collapse multiple spaces
    name = re.sub(r'\s+', ' ', name)
    # Final strip
    return name.strip()


# --------------- Column Alias Mapping ---------------
# Maps normalized user input -> internal COLUMNS field names.
# Multiple aliases can map to the same internal field.

COLUMN_ALIASES = {
    # Order Number
    "order no": "Order No.",
    "order number": "Order No.",
    "order": "Order No.",
    "order num": "Order No.",
    "order id": "Order No.",

    # Date
    "date": "Date",
    "order date": "Date",

    # Bill To - Customer Name
    "bill to customer name": "Bill to Customer Name",
    "bill to name": "Bill to Customer Name",
    "bill customer name": "Bill to Customer Name",
    "bill name": "Bill to Customer Name",
    "bill to": "Bill to Customer Name",
    "billing name": "Bill to Customer Name",
    "customer name": "Bill to Customer Name",

    # Bill To - Address
    "bill to address": "Bill To Address",
    "bill address": "Bill To Address",
    "billing address": "Bill To Address",
    "address": "Bill To Address",

    # Bill To - Phone
    "bill to phone": "Bill To Phone",
    "bill phone": "Bill To Phone",
    "billing phone": "Bill To Phone",
    "phone": "Bill To Phone",
    "phone number": "Bill To Phone",

    # Deliver To - Customer Name
    "deliver to customer name": "DELIVER to Customer Name",
    "deliver to name": "DELIVER to Customer Name",
    "deliver customer name": "DELIVER to Customer Name",
    "deliver name": "DELIVER to Customer Name",
    "deliver to": "DELIVER to Customer Name",
    "delivery name": "DELIVER to Customer Name",
    "shipping name": "DELIVER to Customer Name",

    # Deliver To - Address
    "deliver to address": "DELIVER To Address",
    "deliver address": "DELIVER To Address",
    "delivery address": "DELIVER To Address",
    "shipping address": "DELIVER To Address",

    # Deliver To - Phone
    "deliver to phone": "DELIVER To Phone",
    "deliver phone": "DELIVER To Phone",
    "delivery phone": "DELIVER To Phone",
    "shipping phone": "DELIVER To Phone",

    # Product Name
    "product name": "Product Name",
    "product": "Product Name",
    "item name": "Product Name",
    "item": "Product Name",

    # Shop SKU
    "shop sku": "Shop SKU",
    "shopsku": "Shop SKU",

    # Seller SKU
    "seller sku": "Seller SKU",
    "sellersku": "Seller SKU",
    "sku": "Seller SKU",

    # Variant / Attributes
    "variant": "Variant / Attributes",
    "variant attributes": "Variant / Attributes",
    "attributes": "Variant / Attributes",
    "variations": "Variant / Attributes",

    # Quantity
    "quantity": "Quantity",
    "qty": "Quantity",

    # Unit Price
    "unit price": "Unit Price",
    "price": "Unit Price",
    "unitprice": "Unit Price",

    # Shipping Cost
    "shipping cost": "Shipping Cost",
    "shipping": "Shipping Cost",
    "delivery cost": "Shipping Cost",

    # Voucher
    "voucher": "Voucher",
    "discount": "Voucher",
    "coupon": "Voucher",

    # Total
    "total": "Total",
    "grand total": "Total",
    "amount": "Total",

    # Payment Method
    "payment method": "Payment Method",
    "payment": "Payment Method",
    "pay method": "Payment Method",

    # Metadata columns
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


def resolve_column(user_input: str) -> Optional[str]:
    """
    Resolve a user-entered column name to an internal COLUMNS field name.

    Returns the internal field name if a match is found, or None if no match exists.
    """
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

