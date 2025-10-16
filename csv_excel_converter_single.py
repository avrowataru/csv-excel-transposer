"""Standalone CSV ↔ Excel converter CLI combining package modules."""
from __future__ import annotations

import argparse
import sys
import textwrap
from pathlib import Path
from typing import Union

import pandas as pd


class ConversionError(Exception):
    """Raised when an import or export action fails."""


SheetSelector = Union[str, int]


def csv_to_excel(
    csv_path: Path,
    excel_path: Path,
    *,
    sheet_name: SheetSelector = "Sheet1",
    delimiter: str = ",",
    encoding: str = "utf-8",
    has_header: bool = True,
    transpose: bool = False,
    include_index: bool = False,
) -> None:
    """Convert a CSV file into an Excel workbook."""
    csv_path = Path(csv_path)
    excel_path = Path(excel_path)

    if not csv_path.exists():
        raise ConversionError(f"CSV file not found: {csv_path}")

    header = 0 if has_header else None
    try:
        frame = pd.read_csv(csv_path, sep=delimiter, encoding=encoding, header=header)
    except Exception as exc:
        raise ConversionError(f"Unable to read CSV '{csv_path}': {exc}") from exc

    if transpose:
        frame = frame.transpose(copy=True)

    try:
        excel_path.parent.mkdir(parents=True, exist_ok=True)
        frame.to_excel(
            excel_path,
            sheet_name=sheet_name,
            index=include_index,
            engine="openpyxl",
        )
    except Exception as exc:
        raise ConversionError(f"Unable to write Excel '{excel_path}': {exc}") from exc


def excel_to_csv(
    excel_path: Path,
    csv_path: Path,
    *,
    sheet_name: SheetSelector = 0,
    delimiter: str = ",",
    encoding: str = "utf-8",
    transpose: bool = False,
    include_index: bool = False,
    include_header: bool = True,
) -> None:
    """Convert an Excel worksheet into a CSV file."""
    excel_path = Path(excel_path)
    csv_path = Path(csv_path)

    if not excel_path.exists():
        raise ConversionError(f"Excel workbook not found: {excel_path}")

    try:
        frame = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            engine="openpyxl",
        )
    except Exception as exc:
        raise ConversionError(
            f"Unable to read worksheet '{sheet_name}' from '{excel_path}': {exc}"
        ) from exc

    if transpose:
        frame = frame.transpose(copy=True)

    try:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        frame.to_csv(
            csv_path,
            sep=delimiter,
            encoding=encoding,
            index=include_index,
            header=include_header,
            line_terminator="\n",
        )
    except Exception as exc:
        raise ConversionError(f"Unable to write CSV '{csv_path}': {exc}") from exc


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
