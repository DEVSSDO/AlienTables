"""
alienextractor
~~~~~~~~~~~~

A professional, rule-based PDF data extraction package.
Extracts structured data from PDF files using PyMuPDF, Tesseract OCR,
regex patterns, and configurable extraction rules.

Usage (CLI):
    $ alienextractor

Usage (Programmatic):
    >>> from alienextractor.engine import run
    >>> from pathlib import Path
    >>> run(Path("./input"), Path("output.xlsx"), ["Order No.", "Date", "Total"])
"""

__version__ = "1.0.0"
__author__ = "AlienExtractor Team"

from .engine import run
from .extractor import COLUMNS

__all__ = ["run", "COLUMNS", "__version__"]
