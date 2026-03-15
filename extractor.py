
import re
import pdfplumber
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from dataclasses import dataclass, field
from typing import Optional


# ── Data Models ────────────────────────────────────────────────────────────────

@dataclass
class Transaction:
    date: str
    description: str
    debit: Optional[float]
    credit: Optional[float]
    balance: Optional[float]
    date_conf: Optional[float] = None
    description_conf: Optional[float] = None
    debit_conf: Optional[float] = None
    credit_conf: Optional[float] = None
    balance_conf: Optional[float] = None
    source: str = "native"  # "native" or "azure"


@dataclass
class ExtractionResult:
    pdf_name: str
    transactions: list = field(default_factory=list)
    opening_balance: Optional[float] = None
    closing_balance: Optional[float] = None
    extraction_method: str = "native"


# ── Helpers ────────────────────────────────────────────────────────────────────

def _parse_amount(value: str) -> Optional[float]:
    """Converts a string like '$1,234.56' or '1234.56' to float. Returns None if not parseable."""
    if not value or value.strip() in ("", "—", "-", "N/A"):
        return None
    cleaned = re.sub(r"[^\d.]", "", value)
    try:
        return float(cleaned)
    except ValueError:
        return None


def _has_text_layer(pdf_path: str) -> bool:

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if len(text.strip()) > 20:
                    return True
    except Exception:
        pass
    return False


# ── Native Text Extraction ─────────────────────────────────────────────────────

def _extract_native(pdf_path: str, pdf_name: str) -> ExtractionResult:

    result = ExtractionResult(pdf_name=pdf_name, extraction_method="native")

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            parsed_from_table = False

            for table in tables:
                if not table or len(table) < 2:
                    continue

                header = [str(c).lower().strip() if c else "" for c in table[0]]
                col_map = _detect_columns(header)

                if not col_map:
                    continue  # Not a transaction table

                parsed_from_table = True
                for row in table[1:]:
                    if not row or all(c is None or str(c).strip() == "" for c in row):
                        continue

                    def get(key):
                        idx = col_map.get(key)
                        return str(row[idx]).strip() if idx is not None and idx < len(row) and row[idx] else ""

                    date        = get("date")
                    description = get("description")
                    debit       = _parse_amount(get("debit"))
                    credit      = _parse_amount(get("credit"))
                    balance     = _parse_amount(get("balance"))

                    if not date and not description:
                        continue

                    txn = Transaction(
                        date=date,
                        description=description,
                        debit=debit,
                        credit=credit,
                        balance=balance,
                        source="native",
                    )
                    result.transactions.append(txn)

            # ── Fallback: regex on raw text ────────────────────────────────────
            if not parsed_from_table:
                text = page.extract_text() or ""
                _parse_text_with_regex(text, result)

    _infer_opening_closing(result)
    return result


def _detect_columns(header: list) -> dict:
    """
    Maps column names to indices based on keywords in the header row.
    Returns empty dict if it doesn't look like a transaction table.
    """
    col_map = {}
    keywords = {
        "date":        ["date", "txn date", "transaction date", "value date"],
        "description": ["description", "narration", "details", "particulars", "remarks"],
        "debit":       ["debit", "withdrawal", "dr", "debit amount"],
        "credit":      ["credit", "deposit", "cr", "credit amount"],
        "balance":     ["balance", "running balance", "closing balance", "amount"],
    }
    for i, col in enumerate(header):
        for field_name, variants in keywords.items():
            if any(v in col for v in variants):
                col_map[field_name] = i
                break

    # Must have at least date + one amount column to be valid
    if "date" not in col_map:
        return {}
    if not any(k in col_map for k in ("debit", "credit", "balance")):
        return {}

    return col_map


def _parse_text_with_regex(text: str, result: ExtractionResult):
    """
    Regex fallback for when pdfplumber can't detect a table.
    Matches lines like: 04 Mar 2025   Salary Deposit   3,500.00   8,500.00
    """
    DATE_PATTERN = r"(\d{1,2}\s(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s\d{4})"
    AMOUNT       = r"([\d,]+\.\d{2})"
    DASH         = r"(—|-)"

    pattern = re.compile(
        DATE_PATTERN + r"\s+(.+?)\s+" +
        r"(?:" + AMOUNT + r"|" + DASH + r")\s+" +
        r"(?:" + AMOUNT + r"|" + DASH + r")\s+" +
        AMOUNT,
        re.IGNORECASE
    )

    for match in pattern.finditer(text):
        date        = match.group(1).strip()
        description = match.group(2).strip()
        debit_str   = match.group(3) or ""
        credit_str  = match.group(5) or ""
        balance_str = match.group(7) or ""

        result.transactions.append(Transaction(
            date=date,
            description=description,
            debit=_parse_amount(debit_str),
            credit=_parse_amount(credit_str),
            balance=_parse_amount(balance_str),
            source="native",
        ))


# ── Azure Extraction ───────────────────────────────────────────────────────────

def _extract_azure(pdf_path: str, pdf_name: str, endpoint: str, api_key: str) -> ExtractionResult:
    """
    Extracts transactions from a scanned/image PDF using Azure Document Intelligence.
    Uses the prebuilt-layout model to detect tables and extract cell content + confidence.
    """
    result = ExtractionResult(pdf_name=pdf_name, extraction_method="azure")

    client = DocumentAnalysisClient(
        endpoint=endpoint,
        credential=AzureKeyCredential(api_key)
    )

    with open(pdf_path, "rb") as f:
        poller = client.begin_analyze_document("prebuilt-layout", document=f)
    doc = poller.result()


    word_map = {}  # span_offset -> confidence
    for page in doc.pages:
        for word in page.words:
            if word.span:
                word_map[word.span.offset] = word.confidence

    def _cell_confidence(cell) -> Optional[float]:
        """
        Derives cell confidence from the minimum confidence of all
        words whose span offsets fall within the cell's spans.
        Returns None if cell is empty or no words matched.
        """
        if not cell.content.strip() or cell.content.strip() in ("-", "—"):
            return None
        confs = []
        for span in (cell.spans or []):
            for offset, conf in word_map.items():
                if span.offset <= offset < span.offset + span.length:
                    confs.append(conf)
        return min(confs) if confs else None

    for table in doc.tables:
        grid = {}
        for cell in table.cells:
            grid[(cell.row_index, cell.column_index)] = cell

        if table.row_count < 2:
            continue

        header = []
        for col in range(table.column_count):
            cell = grid.get((0, col))
            header.append(cell.content.lower().strip() if cell else "")

        col_map = _detect_columns(header)
        if not col_map:
            continue

        for row_idx in range(1, table.row_count):
            def get_cell(key):
                idx = col_map.get(key)
                if idx is None:
                    return None, None
                cell = grid.get((row_idx, idx))
                if not cell:
                    return None, None
                return cell.content.strip(), _cell_confidence(cell)

            date,        date_conf        = get_cell("date")
            description, description_conf = get_cell("description")
            debit_str,   debit_conf       = get_cell("debit")
            credit_str,  credit_conf      = get_cell("credit")
            balance_str, balance_conf     = get_cell("balance")

            if not date and not description:
                continue

            result.transactions.append(Transaction(
                date=date or "",
                description=description or "",
                debit=_parse_amount(debit_str or ""),
                credit=_parse_amount(credit_str or ""),
                balance=_parse_amount(balance_str or ""),
                date_conf=date_conf,
                description_conf=description_conf,
                debit_conf=debit_conf,
                credit_conf=credit_conf,
                balance_conf=balance_conf,
                source="azure",
            ))

    _infer_opening_closing(result)
    return result


# ── Opening / Closing Balance Inference ───────────────────────────────────────

def _infer_opening_closing(result: ExtractionResult):
    """Sets opening and closing balance from first and last transaction balances."""
    balances = [t.balance for t in result.transactions if t.balance is not None]
    if balances:
        result.opening_balance = balances[0]
        result.closing_balance = balances[-1]


# ── Public Entry Point ─────────────────────────────────────────────────────────

def extract_pdf(pdf_path: str, pdf_name: str, azure_endpoint: str, azure_key: str) -> ExtractionResult:
    """
    Main extraction function.
    Checks for native text layer first; falls back to Azure if not found.
    """
    if _has_text_layer(pdf_path):
        print(f"  [{pdf_name}] Native text layer detected → using pdfplumber")
        return _extract_native(pdf_path, pdf_name)
    else:
        print(f"  [{pdf_name}] No text layer detected → using Azure Document Intelligence")
        return _extract_azure(pdf_path, pdf_name, azure_endpoint, azure_key)