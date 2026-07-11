"""
alienextractor.ocr
~~~~~~~~~~~~~~~~

OCR functionality using Tesseract and pdf2image.
All OCR logic is preserved verbatim from the original script.
"""

import os
from pathlib import Path

from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageFilter
import pytesseract


# --------------- OCR Configuration ---------------
# Read from environment variables with sensible defaults.
# Users can set TESSERACT_CMD and POPPLER_PATH environment variables
# to override auto-detection.

OCR_DPI = 300
USE_OCR_FALLBACK = True

_tesseract_cmd = os.environ.get("TESSERACT_CMD", "")
_poppler_path = os.environ.get("POPPLER_PATH", "")

# Windows default paths as fallbacks
if not _tesseract_cmd:
    _default_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.isfile(_default_tesseract):
        _tesseract_cmd = _default_tesseract

if not _poppler_path:
    # Check common Poppler install locations on Windows
    for candidate in [
        r"C:\poppler-25.12.0\Library\bin",
        r"C:\poppler\Library\bin",
        r"C:\Program Files\poppler\Library\bin",
    ]:
        if os.path.isdir(candidate):
            _poppler_path = candidate
            break

if _tesseract_cmd:
    pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd

# Expose for use in parser.py
POPPLER_PATH = _poppler_path if _poppler_path else None


# --------------- OCR Functions (verbatim from original) ---------------

def preprocess_image_for_ocr(img: Image.Image) -> Image.Image:
    """Preprocess an image to improve OCR accuracy."""
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
    """OCR a single page number (1-based)."""
    try:
        kwargs = {
            "dpi": OCR_DPI,
            "first_page": page_num,
            "last_page": page_num,
        }
        if POPPLER_PATH:
            kwargs["poppler_path"] = POPPLER_PATH
        imgs = convert_from_path(str(pdf_path), **kwargs)
    except Exception:
        return ""
    if not imgs:
        return ""
    im = preprocess_image_for_ocr(imgs[0])
    return pytesseract.image_to_string(im)
