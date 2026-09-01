# Multi-Sheet Excel Merger

A small Streamlit app for combining multiple `.xlsx` workbooks into one merged workbook.

The app:

- Reads every sheet from each uploaded workbook.
- Lets you group source sheets into compiled output sheets.
- Suggests English/French column matches such as `Titre` to `Title`.
- Lets you edit column mappings before merge.
- Adds a `Source_File` column so merged rows remain traceable.
- Exports the merged workbook as `merged_workbooks.xlsx`.

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

- Compiled output sheets are limited to five.
- Excel sheet names cannot contain `: \ / ? * [ ]`.
- Excel sheet names are limited to 31 characters.
- When two source columns map to the same output field in one sheet, the app keeps the first non-empty value from left to right.
