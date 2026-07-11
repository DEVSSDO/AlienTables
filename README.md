# AlienTables (AlienExtractor)

A professional, rule-based PDF data extraction tool with an interactive CLI, specifically tailored for Daraz Purchase/Order Summaries.

## Features

- **Interactive CLI**: Prompts for input folder, output folder, and column names.
- **Rule-Based Parsing**: Efficiently extracts key information like Order No, Date, Customer Details, Product Name, SKUs, Pricing, and more using Regex and positional logic.
- **Excel Output**: Automatically writes one Excel file per PDF (named the same as the PDF).
- **OCR Fallback**: Integrates `pytesseract` and `pdf2image` for pages that fail text extraction, ensuring maximum data recovery.
- **Detailed Auditing**: Automatically generates auditing CSVs:
  - `Audit.csv`: Logs extraction status for each page.
  - `orphan_pages.csv`: Logs pages that couldn't be associated with a valid order.
  - `review.csv`: Logs parsing issues or low-confidence extractions for manual review.

## Prerequisites

Before installing the python dependencies, ensure you have the following system requirements installed:

1. **[Tesseract OCR](https://github.com/tesseract-ocr/tesseract)**
   - Windows: Install using the [installer](https://github.com/UB-Mannheim/tesseract/wiki) and ensure it's in your system PATH or update the `TESSERACT_CMD` variable in `AlienTables.py`.
2. **[Poppler](https://poppler.freedesktop.org/)** (Required for `pdf2image`)
   - Windows: Download the latest binaries and update the `POPPLER_PATH` variable in `AlienTables.py` to point to the `bin` directory.

## Installation

This project uses `pyproject.toml` for standard Python package installation. Python 3.8 or higher is required.

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

You can run the interactive CLI tool using the installed script:

```bash
alienextractor
```

Or you can run the main script directly:

```bash
python AlienTables.py
```

Follow the on-screen interactive prompts to:
1. Provide the input directory containing your PDF files.
2. Provide the output directory for the Excel files.
3. Configure any column mapping preferences.

## Output

The tool processes the PDFs and generates:
- Individual `.xlsx` files corresponding to each parsed `.pdf` file.
- `Audit.csv`, `orphan_pages.csv`, and `review.csv` in your working directory for quality assurance.

## Dependencies

- `PyMuPDF`
- `pandas`
- `openpyxl`
- `python-dateutil`
- `tqdm`
- `pdf2image`
- `Pillow`
- `pytesseract`

## License

This project is licensed under the MIT License.
