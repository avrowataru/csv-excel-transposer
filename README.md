# CSV Excel Converter

A small command-line tool that converts CSV files into Excel workbooks and back again. It is particularly handy when you need to transpose matrix-style data or switch between wide and tall orientation while keeping everything scriptable.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -e .[dev]
```

If you only need the runtime, drop the `.[dev]` extras suffix.

## Command-line usage

The installer exposes the `csv-excel-converter` command (or run `python -m csv_excel_converter`). Run with `-h` for full help:

```text
usage: csv-excel-converter [-h] {csv-to-excel,excel-to-csv} ...

Convert CSV files to Excel workbooks and back again.

positional arguments:
  {csv-to-excel,excel-to-csv}
    csv-to-excel         Convert a CSV file into an Excel workbook.
    excel-to-csv         Convert an Excel worksheet into a CSV file.

optional arguments:
  -h, --help             show this help message and exit

Examples:
  csv-excel-converter csv-to-excel data/quarterly.csv reports/q1.xlsx
  csv-excel-converter csv-to-excel data/people.csv reports/wide.xlsx --transpose --sheet-name WideView
  csv-excel-converter excel-to-csv reports/q1.xlsx data/q1.csv --sheet-name Q1 --transpose --encoding latin-1 --delimiter ";"
  csv-excel-converter excel-to-csv matrix.xlsx matrix.csv --no-header --include-index
```

### CSV → Excel

```bash
csv-excel-converter csv-to-excel data/quarterly.csv reports/q1.xlsx \
  --sheet-name Finance --transpose --delimiter ';' --encoding utf-8
```

Key options:

- `--transpose` rotates rows/columns before writing. Perfect for turning a list of metrics into a cross-tab.
- `--no-header` treats the first row as data so transposed output is not crippled by headings.
- `--include-index` keeps the DataFrame index if you need a lookup column in Excel.

### Excel → CSV

```bash
csv-excel-converter excel-to-csv reports/q1.xlsx exports/q1.csv \
  --sheet-name Finance --delimiter ';' --transpose --no-header
```

Highlights:

- `--sheet-name` accepts either the worksheet name or a zero-based index (e.g. `--sheet-name 1` for the second sheet).
- `--no-header` suppresses the column header row when exporting, useful for raw matrix data.
- `--include-index` keeps row labels that originated as Excel index columns.

### Validation tips

- Pandas follows Python’s zero-based row numbering. For headerless numeric data, combine `--no-header` with `--include-index` to preserve both axes.
- When working with locale-specific CSVs (e.g. semicolon separated, Latin-1 encoding), set both `--delimiter` and `--encoding` so pandas parses values correctly.
- If Excel complains about scientific notation after transpose, prepend a single quote in your CSV input to force string preservation or post-process via pandas dtype controls.

## Development

Run the tests:

```bash
pytest
```

The suite covers round-tripping to ensure you do not lose data during conversion, including transposed scenarios. Add more fixtures if you handle unusual encodings or multi-sheet workflows.
