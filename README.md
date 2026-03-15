# Bank Statement PDF Extractor

A Python tool that extracts transactions from bank statement PDFs into a structured Excel file. Supports both native-text and scanned/image-based PDFs.

---

## Features

- Detects native text layer — extracts directly with `pdfplumber` (fast, no API cost)
- Falls back to **Azure Document Intelligence** for scanned/image PDFs
- Extracts: Date, Description, Debit, Credit, Running Balance
- Flags low-confidence fields (< 85%) with **red cell fill** in Excel
- Validates: Opening Balance + Credits − Debits = Closing Balance
- Reports exact variance amount if balance check fails
- One Excel sheet per PDF, one row per transaction

---

## Project Structure

```
project/
├── main.py                    # Entry point (CLI)
├── extractor.py               # PDF extraction logic (native + Azure)
├── excel_writer.py            # Excel output with formatting & flags
├── validator.py               # Balance validation logic
├── sample_pdfs/
│   ├── native_statement.pdf   # Text-layer PDF (pdfplumber path)
│   └── scanned_statement.pdf  # Image-only PDF (Azure path)
├── requirements.txt
└── README.md
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Poppler (required by `pdf2image` for sample generation only)

| OS      | Command                                                                 |
|---------|-------------------------------------------------------------------------|
| Windows | Download from [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases), add `bin/` to PATH |
| Linux   | `sudo apt install poppler-utils`                                        |
| Mac     | `brew install poppler`                                                  |

### 3. Set Azure credentials

```bash
# Option A — environment variables (recommended)
export AZURE_ENDPOINT="https://your-resource.cognitiveservices.azure.com/"
export AZURE_API_KEY="your-api-key-here"

# Option B — pass via CLI flags (see Usage below)
```

---

## Usage

```bash
python main.py --folder sample_pdfs --output output/results.xlsx
```

### CLI Options

| Flag               | Description                                      | Default                  |
|--------------------|--------------------------------------------------|--------------------------|
| `--folder` / `-f`  | Path to folder containing PDF files             | *(required)*             |
| `--output` / `-o`  | Output Excel file path                           | `output/results.xlsx`    |
| `--azure-endpoint` | Azure Document Intelligence endpoint URL        | `$AZURE_ENDPOINT` env var|
| `--azure-key`      | Azure Document Intelligence API key             | `$AZURE_API_KEY` env var |

---

## Sample PDFs

Two sample bank statement PDFs are included in the `sample_pdfs/` folder for testing:

- `native_statement.pdf` — generated using `reportlab` with a real text layer, processed by `pdfplumber`
- `scanned_statement.pdf` — created by converting the native PDF pages to images using `pdf2image` and re-embedding them into a new PDF using `img2pdf`, which strips the text layer entirely and triggers the Azure fallback

---

## Excel Output Format

Each PDF gets its own sheet with:

| Column       | Notes                                         |
|--------------|-----------------------------------------------|
| Date         | Transaction date                              |
| Description  | Narration / transaction detail                |
| Debit ($)    | Amount debited (blank = `—`)                  |
| Credit ($)   | Amount credited (blank = `—`)                 |
| Balance ($)  | Running balance                               |
| Flags        | Lists low-confidence fields, or `✓` if clean  |

**Red cell fill** = Azure confidence score below 85% for that specific field.

A **Balance Validation Summary** is appended at the bottom of each sheet showing opening balance, totals, expected vs actual closing balance, and variance if any.

---

## Assumptions

1. The first transaction row's balance is treated as the **opening balance**.
2. The last transaction row's balance is treated as the **closing balance**.
3. Rows with both zero debit and zero credit (e.g., opening/closing rows) are excluded from totals.
4. A floating-point tolerance of **$0.01** is allowed in balance validation.
5. For native PDFs, `pdfplumber` table extraction is attempted first; regex parsing is used as a fallback for unstructured text layouts.
6. Azure's `prebuilt-layout` model is used for scanned PDFs — it provides per-cell confidence scores.
