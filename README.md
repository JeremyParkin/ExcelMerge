# Excel & CSV Row Stacker

A small Streamlit app for stacking matching rows from Excel worksheets and CSV files
into one merged workbook.

The app:

- Reads each uploaded CSV file and every sheet from each uploaded Excel workbook.
- Lets you choose which CSV files or Excel worksheets to include.
- Stacks included sources into one output worksheet.
- Suggests English/French column matches such as `Titre` to `Title`.
- Lets you filter, include, exclude, and rename columns before merge.
- Adds a `Source_File` column so merged rows remain traceable.
- Exports the merged workbook as `merged_sources.xlsx`.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run main.py
```

## Test

```bash
python3 -m unittest
```

## Notes

- Large workbooks can have many source worksheets. Use the `Include` checkbox
  to choose only the worksheet you want from each workbook before merging.
- Use the source filter to find matching file or sheet names, then select or clear
  just the visible rows.
- Use the column filter and `Include` checkbox to keep only
  the columns you need.
- When two source columns map to the same output field in one sheet, the app keeps the first non-empty value from left to right.
