"""Regression tests for csv_excel_converter."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from csv_excel_converter.converter import csv_to_excel, excel_to_csv


def test_csv_to_excel_roundtrip(tmp_path: Path) -> None:
    csv_path = tmp_path / "scores.csv"
    csv_path.write_text("name,score\nAlice,10\nBob,12\n", encoding="utf-8")
    excel_path = tmp_path / "scores.xlsx"
    csv_to_excel(csv_path, excel_path, sheet_name="Scores")

    roundtrip_csv = tmp_path / "scores_roundtrip.csv"
    excel_to_csv(excel_path, roundtrip_csv, sheet_name="Scores")

    original = pd.read_csv(csv_path)
    rebuilt = pd.read_csv(roundtrip_csv)
    assert original.equals(rebuilt)


def test_transpose_roundtrip(tmp_path: Path) -> None:
    csv_path = tmp_path / "matrix.csv"
    csv_path.write_text("metric,Q1,Q2\nA,1,2\nB,3,4\n", encoding="utf-8")
    excel_path = tmp_path / "matrix.xlsx"
    csv_to_excel(csv_path, excel_path, sheet_name="Matrix", transpose=True)

    csv_back = tmp_path / "matrix_back.csv"
    excel_to_csv(
        excel_path,
        csv_back,
        sheet_name="Matrix",
        transpose=True,
        include_header=True,
    )

    original = pd.read_csv(csv_path)
    rebuilt = pd.read_csv(csv_back)
    assert original.equals(rebuilt)
