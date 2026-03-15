from dataclasses import dataclass
from typing import Optional
from extractor import ExtractionResult


@dataclass
class ValidationResult:
    pdf_name: str
    opening_balance: Optional[float]
    total_credits: float
    total_debits: float
    expected_closing: Optional[float]
    actual_closing: Optional[float]
    is_valid: bool
    variance: Optional[float]  # None if valid, exact difference if not


def validate(result: ExtractionResult) -> ValidationResult:
    """
    Runs balance validation on an ExtractionResult.

    Formula: opening_balance + total_credits - total_debits = closing_balance
    Variance = expected_closing - actual_closing
    """
    transactions = result.transactions

    middle = transactions[1:-1] if len(transactions) > 2 else transactions

    total_debits  = sum(t.debit  for t in middle if t.debit  is not None)
    total_credits = sum(t.credit for t in middle if t.credit is not None)

    opening = result.opening_balance
    closing = result.closing_balance

    if opening is not None:
        expected_closing = round(opening + total_credits - total_debits, 2)
    else:
        expected_closing = None

    if expected_closing is not None and closing is not None:
        variance = round(expected_closing - closing, 2)
        is_valid = abs(variance) <= 0.01
    else:
        variance = None
        is_valid = False

    return ValidationResult(
        pdf_name=result.pdf_name,
        opening_balance=opening,
        total_credits=total_credits,
        total_debits=total_debits,
        expected_closing=expected_closing,
        actual_closing=closing,
        is_valid=is_valid,
        variance=variance if not is_valid else None,
    )


def print_validation_report(vr: ValidationResult):
    """Prints a human-readable validation report to console."""
    print(f"\n  {'✅' if vr.is_valid else '❌'} Balance Validation — {vr.pdf_name}")
    print(f"     Opening Balance : ${vr.opening_balance:,.2f}" if vr.opening_balance else "     Opening Balance : N/A")
    print(f"     Total Credits   : ${vr.total_credits:,.2f}")
    print(f"     Total Debits    : ${vr.total_debits:,.2f}")
    print(f"     Expected Closing: ${vr.expected_closing:,.2f}" if vr.expected_closing else "     Expected Closing: N/A")
    print(f"     Actual Closing  : ${vr.actual_closing:,.2f}"  if vr.actual_closing  else "     Actual Closing  : N/A")
    if not vr.is_valid and vr.variance is not None:
        print(f"     VARIANCE     : ${vr.variance:,.2f}")
    else:
        print("     Balance checks out ✓")