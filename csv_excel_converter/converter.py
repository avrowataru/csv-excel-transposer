"""Conversion helpers for CSV ↔ Excel."""
from __future__ import annotations

from pathlib import Path
from typing import Union

import pandas as pd

from .exceptions import ConversionError

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
