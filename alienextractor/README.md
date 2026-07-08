# AlienExtractor

A professional, rule-based PDF data extraction tool that converts structured PDF documents into Excel spreadsheets.

## Features

- **Rule-Based Extraction** — Uses regex patterns and string matching for reliable, deterministic data extraction
- **Interactive CLI** — User-friendly command-line interface with guided prompts
- **Flexible Column Selection** — Extract only the columns you need, in any order
- **Fuzzy Column Matching** — Handles variations in column names (case, spacing, punctuation)
- **OCR Fallback** — Automatically uses Tesseract OCR when PDF text extraction fails
- **Batch Processing** — Process entire folders of PDF files at once
- **Excel Output** — Clean `.xlsx` output with only your requested columns

## Prerequisites

Before installing, ensure these system dependencies are available:

### Tesseract OCR
- **Windows**: Download from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt install tesseract-ocr`

### Poppler (for pdf2image)
- **Windows**: Download from [poppler releases](https://github.com/osesam/poppler/releases) and add to PATH
- **macOS**: `brew install poppler`
- **Linux**: `sudo apt install poppler-utils`

## Installation

```bash
pip install alienextractor
```

Or install from source:

```bash
git clone https://github.com/alienextractor/alienextractor.git
cd alienextractor
pip install -e .
```

## Usage

### CLI

Run the interactive extractor:

```bash
alienextractor
```

You'll be guided through:

1. **PDF Folder** — Enter the path to your PDF files (default: `./input`)
2. **Output File** — Name your Excel output file (auto-generated if blank)
3. **Column Names** — Enter column names one at a time, type `OK` when finished

### Example Session

```
=========================================
  AlienExtractor
=========================================

  Enter PDF Folder location: (default: ./input)
  > ./input

  Enter Output Excel File (optional):
  > my_orders.xlsx

  Enter Column Name:
  > Order No.

  Enter Column Name:
  > Date

  Enter Column Name:
  > Customer Name

  Enter Column Name:
  > Total

  Enter Column Name:
  > OK

  Processing 5 PDF file(s)...
  Done! Output saved to: my_orders.xlsx
```

### Programmatic Usage

```python
from pathlib import Path
from alienextractor import run

result = run(
    pdf_folder=Path("./input"),
    output_file=Path("output.xlsx"),
    requested_columns=["Order No.", "Date", "Total"]
)
```

## Supported Fields

The extraction engine supports these fields (and common aliases):

| Field | Aliases |
|-------|---------|
| Order No. | order number, order, order id |
| Date | order date |
| Bill to Customer Name | customer name, bill to, bill name |
| Bill To Address | address, billing address |
| Bill To Phone | phone, phone number |
| DELIVER to Customer Name | deliver to, delivery name |
| DELIVER To Address | delivery address, shipping address |
| DELIVER To Phone | deliver phone, delivery phone |
| Product Name | product, item name, item |
| Shop SKU | shopsku |
| Seller SKU | sellersku, sku |
| Variant / Attributes | variant, attributes |
| Quantity | qty |
| Unit Price | price, unitprice |
| Shipping Cost | shipping, delivery cost |
| Voucher | discount, coupon |
| Total | grand total, amount |
| Payment Method | payment |

**Any column name** is accepted. Unrecognized names will appear in the output with blank values.

## Configuration

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `TESSERACT_CMD` | Path to Tesseract executable | Auto-detected |
| `POPPLER_PATH` | Path to Poppler bin directory | Auto-detected |

### Example (Windows)

```powershell
$env:TESSERACT_CMD = "C:\Program Files\Tesseract-OCR\tesseract.exe"
$env:POPPLER_PATH = "C:\poppler-25.12.0\Library\bin"
alienextractor
```

## Project Structure

```
alienextractor/
├── __init__.py      # Package exports and version
├── cli.py           # Interactive command-line interface
├── engine.py        # Main workflow controller
├── parser.py        # PDF reading and page processing
├── ocr.py           # OCR functionality (Tesseract)
├── extractor.py     # Extraction rules and regex patterns
├── excel.py         # Excel output generation
└── utils.py         # Shared helpers and column mapping
```

## License

MIT License — see [LICENSE](LICENSE) for details.
