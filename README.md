<div align="center">

# 👽 AlienTables

**Turn thousands of PDFs into clean Excel data in minutes.**

Enterprise-grade PDF → Excel extraction engine with adaptive OCR, template-aware parsing, and automated validation — built for blazing-fast, production-scale batch processing.

Secure · Runs Locally · Adaptive OCR · Smart Validation · Zero Manual Entry

</div>

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [How It Works](#how-it-works)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Rules & Best Practices](#rules--best-practices)
- [Output](#output)
- [Project Structure](#project-structure)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Dependencies](#dependencies)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

AlienTables is a local, privacy-first tool that converts structured PDF documents (invoices, purchase orders, shipping manifests, etc.) into clean, organized Excel spreadsheets — no cloud upload, no manual data entry, no per-page babysitting.

It combines rule-based text parsing with an OCR fallback, so it works reliably on both digitally-generated PDFs and scanned/image-based documents, as long as all files in a batch share the same template layout.

## Features

- 🖥️ **Interactive CLI** — Prompts for input folder, output folder, and column names, no config files required.
- 🧠 **Rule-Based Parsing** — Extracts key fields (Order No, Date, Customer Details, Product Name, SKUs, Pricing, and more) using regex and positional logic.
- 📊 **Excel Output** — Automatically writes one `.xlsx` file per PDF, named to match the source file.
- 🔍 **OCR Fallback** — Integrates `pytesseract` and `pdf2image` for pages that fail standard text extraction, maximizing data recovery. OCR only kicks in when it's actually needed, keeping runs fast.
- 🕵️ **Detailed Auditing** — Every run generates audit CSVs so you always know what happened:
  - `Audit.csv` — extraction status for every page processed.
  - `orphan_pages.csv` — pages that couldn't be associated with a valid order.
  - `review.csv` — parsing issues or low-confidence extractions flagged for manual review.
- 🔒 **Fully Local** — All processing happens on your machine. No data ever leaves your environment.
- ⚡ **Batch-Scale** — Built to process thousands of PDFs in a single run.

## How It Works

1. You point AlienTables at a folder of PDFs that share the same layout/template.
2. It attempts direct text extraction first, falling back to OCR only on pages where extraction fails.
3. Regex and positional rules map the extracted text to the column names you provide.
4. Each PDF is written out as its own Excel file, and every page's outcome is logged to the audit CSVs.

## Prerequisites

Before installing the Python dependencies, make sure the following system-level tools are installed:

### 1. [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
Required for OCR fallback on non-text-extractable pages.
- **Windows**: Install via the [UB-Mannheim installer](https://github.com/UB-Mannheim/tesseract/wiki), then either add it to your system `PATH` or update the `TESSERACT_CMD` variable in `AlienTables.py`.
- **macOS**: `brew install tesseract`
- **Linux (Debian/Ubuntu)**: `sudo apt install tesseract-ocr`

### 2. [Poppler](https://poppler.freedesktop.org/)
Required by `pdf2image` for rendering PDF pages.
- **Windows**: Download the latest binaries and update the `POPPLER_PATH` variable in `AlienTables.py` to point to the `bin` directory.
- **macOS**: `brew install poppler`
- **Linux (Debian/Ubuntu)**: `sudo apt install poppler-utils`

## Installation

This project uses `pyproject.toml` for standard Python packaging. **Python 3.8+** is required.

```bash
# Clone the repository
git clone https://github.com/alienextractor/alienextractor.git
cd alienextractor

# Create a virtual environment (optional but recommended)
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\Scripts\activate

# Install the package and dependencies
pip install .
```

## Usage

Run the interactive CLI using the installed script:

```bash
alienextractor
```

Or run the main script directly:

```bash
python AlienTables.py
```

You'll be prompted to:

1. Provide the **input directory** containing your PDF files (defaults to an `input` folder).
2. Provide the **output directory** for the generated Excel files.
3. Enter the **exact column names** that exist in your PDF template.

## Rules & Best Practices

Following these rules ensures accurate, predictable extraction:

1. **Enter exact column names** — they must match what actually exists in your PDF template.
2. **Wrong column names = bad data** — incorrect names can produce blank or random values in the output.
3. **One template per run** — all PDFs in a batch should follow the same layout/template.
4. **Multiple templates?** — process each template separately, in its own run.
5. **Group consistently** — keep PDFs with the same columns together for a single run.
6. **OCR is automatic** — it only triggers when standard text extraction fails, keeping processing fast.
7. **Default input location** — create an `input` folder and place all PDFs there unless you specify a custom path.

## Output

After processing, AlienTables generates:

| File | Description |
|---|---|
| `<filename>.xlsx` | One Excel file per source PDF, containing extracted structured data |
| `Audit.csv` | Extraction status for every page processed |
| `orphan_pages.csv` | Pages that couldn't be matched to a valid order |
| `review.csv` | Low-confidence or problematic extractions flagged for manual review |

## Project Structure

```
alienextractor/
├── AlienTables.py       # Main extraction script / CLI entry point
├── pyproject.toml       # Package metadata and dependencies
├── input/               # Default folder for source PDFs
└── output/               # Default folder for generated Excel files
```

## Configuration

- **`TESSERACT_CMD`** — path to your Tesseract executable (edit in `AlienTables.py` if not in `PATH`).
- **`POPPLER_PATH`** — path to your Poppler `bin` directory (edit in `AlienTables.py` if not in `PATH`).
- **Column names** — provided interactively at runtime; must exactly match your PDF template's fields.

## Troubleshooting

| Issue | Likely Cause | Fix |
|---|---|---|
| Blank or random values in output | Column names don't match the PDF template | Double-check exact spelling/casing of column names |
| Slow processing | OCR fallback triggering frequently | Verify PDFs are text-based, not scanned images, where possible |
| `TesseractNotFoundError` | Tesseract not installed or not in `PATH` | Install Tesseract and/or set `TESSERACT_CMD` |
| `PDFInfoNotInstalledError` | Poppler not installed or not in `PATH` | Install Poppler and/or set `POPPLER_PATH` |
| Pages showing up in `orphan_pages.csv` | Page couldn't be linked to a valid order | Review the page manually; template may differ from the rest of the batch |

## Dependencies

- [`PyMuPDF`](https://pypi.org/project/PyMuPDF/) — PDF text extraction
- [`pandas`](https://pypi.org/project/pandas/) — data handling
- [`openpyxl`](https://pypi.org/project/openpyxl/) — Excel file generation
- [`python-dateutil`](https://pypi.org/project/python-dateutil/) — date parsing
- [`tqdm`](https://pypi.org/project/tqdm/) — progress bars
- [`pdf2image`](https://pypi.org/project/pdf2image/) — PDF-to-image conversion for OCR
- [`Pillow`](https://pypi.org/project/Pillow/) — image processing
- [`pytesseract`](https://pypi.org/project/pytesseract/) — OCR engine wrapper

## Contributing

Issues and pull requests are welcome. Please open an issue first to discuss significant changes.

## License

This project is licensed under the [MIT License](LICENSE).
