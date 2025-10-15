"""Command-line interface for the csv-excel-converter package."""
from __future__ import annotations

import argparse
import sys
import textwrap
from pathlib import Path
from typing import Union

from .converter import csv_to_excel, excel_to_csv
from .exceptions import ConversionError

SheetSelector = Union[str, int]


def _sheet_identifier(value: str) -> SheetSelector:
    """Coerce sheet identifiers that look like integers into zero-based indexes."""
    value = value.strip()
    if value.isdigit():
        return int(value)
    return value


def _build_parser() -> argparse.ArgumentParser:
    examples = """
        Examples:
          csv-excel-converter csv-to-excel data/quarterly.csv reports/q1.xlsx
          csv-excel-converter csv-to-excel data/people.csv reports/wide.xlsx --transpose --sheet-name WideView
          csv-excel-converter excel-to-csv reports/q1.xlsx data/q1.csv --sheet-name Q1 --transpose --encoding latin-1 --delimiter ";"
          csv-excel-converter excel-to-csv matrix.xlsx matrix.csv --no-header --include-index
    """
    parser = argparse.ArgumentParser(
        prog="csv-excel-converter",
        description="Convert CSV files to Excel workbooks and back again.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(examples),
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    csv_to_excel_parser = subparsers.add_parser(
        "csv-to-excel",
        help="Convert a CSV file into an Excel workbook.",
    )
    csv_to_excel_parser.add_argument(
        "csv_path",
        type=Path,
        help="Path to the source CSV file.",
    )
    csv_to_excel_parser.add_argument(
        "excel_path",
        type=Path,
        help="Destination path for the Excel workbook.",
    )
    csv_to_excel_parser.add_argument(
        "--sheet-name",
        default="Sheet1",
        help="Worksheet name to create (default: Sheet1).",
    )
    csv_to_excel_parser.add_argument(
        "--delimiter",
        default=",",
        help="Input CSV delimiter (default: ',').",
    )
    csv_to_excel_parser.add_argument(
        "--encoding",
        default="utf-8",
        help="Encoding used to read the CSV (default: utf-8).",
    )
    csv_to_excel_parser.add_argument(
        "--no-header",
        action="store_true",
        help="Treat the first row as data instead of column names.",
    )
    csv_to_excel_parser.add_argument(
        "--transpose",
        action="store_true",
        help="Swap rows and columns before writing to Excel.",
    )
    csv_to_excel_parser.add_argument(
        "--include-index",
        action="store_true",
        help="Persist the row index in the worksheet.",
    )
    csv_to_excel_parser.set_defaults(func=_command_csv_to_excel)

    excel_to_csv_parser = subparsers.add_parser(
        "excel-to-csv",
        help="Convert an Excel worksheet into a CSV file.",
    )
    excel_to_csv_parser.add_argument(
        "excel_path",
        type=Path,
        help="Path to the source Excel workbook.",
    )
    excel_to_csv_parser.add_argument(
        "csv_path",
        type=Path,
        help="Destination path for the CSV file.",
    )
    excel_to_csv_parser.add_argument(
        "--sheet-name",
        type=_sheet_identifier,
        default=0,
        help="Worksheet name or zero-based index (default: 0).",
    )
    excel_to_csv_parser.add_argument(
        "--delimiter",
        default=",",
        help="Delimiter to use in the output CSV (default: ',').",
    )
    excel_to_csv_parser.add_argument(
        "--encoding",
        default="utf-8",
        help="Encoding to use for the output CSV (default: utf-8).",
    )
    excel_to_csv_parser.add_argument(
        "--transpose",
        action="store_true",
        help="Swap rows and columns before writing to CSV.",
    )
    excel_to_csv_parser.add_argument(
        "--include-index",
        action="store_true",
        help="Include the DataFrame index as the first column.",
    )
    excel_to_csv_parser.add_argument(
        "--no-header",
        action="store_true",
        help="Do not write column names to the CSV output.",
    )
    excel_to_csv_parser.set_defaults(func=_command_excel_to_csv)

    return parser


def _command_csv_to_excel(args: argparse.Namespace) -> int:
    try:
        csv_to_excel(
            csv_path=args.csv_path,
            excel_path=args.excel_path,
            sheet_name=args.sheet_name,
            delimiter=args.delimiter,
            encoding=args.encoding,
            has_header=not args.no_header,
            transpose=args.transpose,
            include_index=args.include_index,
        )
    except ConversionError as exc:
        print(f"conversion failed: {exc}", file=sys.stderr)
        return 1
    return 0


def _command_excel_to_csv(args: argparse.Namespace) -> int:
    try:
        excel_to_csv(
            excel_path=args.excel_path,
            csv_path=args.csv_path,
            sheet_name=args.sheet_name,
            delimiter=args.delimiter,
            encoding=args.encoding,
            transpose=args.transpose,
            include_index=args.include_index,
            include_header=not args.no_header,
        )
    except ConversionError as exc:
        print(f"conversion failed: {exc}", file=sys.stderr)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    parsed = parser.parse_args(argv)
    return parsed.func(parsed)


if __name__ == "__main__":
    sys.exit(main())
