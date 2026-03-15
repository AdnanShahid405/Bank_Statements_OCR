import os
import sys
import argparse
import glob
from extractor import extract_pdf
from excel_writer import write_excel
from validator import validate, print_validation_report
from dotenv import load_dotenv

# ── CLI Arguments ──────────────────────────────────────────────────────────────
load_dotenv()
def parse_args():
    parser = argparse.ArgumentParser(
        description="Extract bank statement transactions from PDFs into Excel."
    )
    parser.add_argument(
        "--folder", "-f",
        required=True,
        help="Path to folder containing PDF files."
    )
    parser.add_argument(
        "--output", "-o",
        default="output/results.xlsx",
        help="Output Excel file path. (default: output/results.xlsx)"
    )
    parser.add_argument(
        "--azure-endpoint",
        default=os.getenv("AZURE_ENDPOINT", ""),
        help="Azure Document Intelligence endpoint URL. Can also set AZURE_ENDPOINT env var."
    )
    parser.add_argument(
        "--azure-key",
        default=os.getenv("AZURE_API_KEY", ""),
        help="Azure Document Intelligence API key. Can also set AZURE_API_KEY env var."
    )
    return parser.parse_args()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()

    # ── Validate folder ───────────────────────────────────────────────────────
    if not os.path.isdir(args.folder):
        print(f"Error: Folder not found: {args.folder}")
        sys.exit(1)

    pdf_files = sorted(glob.glob(os.path.join(args.folder, "*.pdf")))
    if not pdf_files:
        print(f"Error: No PDF files found in: {args.folder}")
        sys.exit(1)

    print(f"\nFound {len(pdf_files)} PDF(s) in '{args.folder}'")
    print(f"Output: {args.output}\n")
    print("=" * 60)

    results     = []
    validations = []

    for pdf_path in pdf_files:
        pdf_name = os.path.basename(pdf_path)
        print(f"\nProcessing: {pdf_name}")

        # ── Extract ───────────────────────────────────────────────────────────
        try:
            result = extract_pdf(
                pdf_path=pdf_path,
                pdf_name=pdf_name,
                azure_endpoint=args.azure_endpoint,
                azure_key=args.azure_key,
            )
        except Exception as e:
            print(f"  Extraction failed for {pdf_name}: {e}")
            continue

        if not result.transactions:
            print(f"  No transactions found in {pdf_name}. Skipping.")
            continue

        print(f"  Extracted {len(result.transactions)} transaction(s)")

        # ── Validate ──────────────────────────────────────────────────────────
        validation = validate(result)
        print_validation_report(validation)

        results.append(result)
        validations.append(validation)

    # ── Write Excel ───────────────────────────────────────────────────────────
    if results:
        print("\n" + "=" * 60)
        write_excel(results, validations, args.output)
    else:
        print("\nNo data extracted. Excel file not created.")

    print("\nDone.\n")


if __name__ == "__main__":
    main()