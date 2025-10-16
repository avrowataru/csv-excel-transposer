#!/usr/bin/env python3

##csv_to_excel.py

##Convert a CSV file to an Excel workbook. Pass --transpose to swap rows and columns.
  
##pip install pandas openpyxl

import argparse
from pathlib import Path

try:
    import pandas as pd
except ImportError as exc:
    raise SystemExit(
    "pandas is required for this script. Install with: pip install pandas openpyxl"
     ) from exc


def parse_args() -> argparse.Namespace:
      parser = argparse.ArgumentParser(
          description="Convert CSV files to Excel (.xlsx) with optional transposition."
      )
      parser.add_argument("csv_path", type=Path, help="Path to the source CSV file.")
      parser.add_argument("excel_path", type=Path, help="Path where the Excel file is written.")
      parser.add_argument(
          "--transpose",
          action="store_true",
          help="Transpose the table so rows become columns (and vice versa).",
      )
      parser.add_argument(
          "--delimiter",
          default=",",
          help="Field delimiter used in the CSV file (default: ',').",
      )
      parser.add_argument(
          "--encoding",
          default="utf-8-sig",
          help="Encoding of the CSV file (default: utf-8-sig).",
      )
      parser.add_argument(
          "--sheet-name",
          default="Sheet1",
          help="Worksheet name inside the Excel file (default: Sheet1).",
      )
      parser.add_argument(
          "--no-header",
          action="store_true",
          help="Treat the CSV as having no header row and write default column labels.",
      )
      parser.add_argument(
          "--keep-index",
          action="store_true",
          help="Keep the DataFrame index as the first column in Excel (default: drop it).",
      )
      return parser.parse_args()


def main() -> None:
      args = parse_args()

      if not args.csv_path.exists():
          raise SystemExit(f"CSV file not found: {args.csv_path}")

      header = None if args.no_header else "infer"
      df = pd.read_csv(
          args.csv_path,
          sep=args.delimiter,
          encoding=args.encoding,
          header=header,
      )

  if args.transpose:
          df = df.transpose()
          df.index.name = None  # Remove index label after transpose for cleaner output.

   args.excel_path.parent.mkdir(parents=True, exist_ok=True)
      df.to_excel(
          args.excel_path,
          sheet_name=args.sheet_name,
          index=args.keep_index,
          engine="openpyxl",
      )


if __name__ == "__main__":
    main()
