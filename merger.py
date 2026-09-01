import io
import re
import unicodedata
from collections import Counter

import pandas as pd


MAX_COMPILED_SHEETS = 5
EXCEL_SHEET_NAME_LIMIT = 31
INVALID_SHEET_NAME_CHARS = re.compile(r"[:\\/?*\[\]]")

BILINGUAL_COLUMN_HINTS = {
    "titre": "Title",
    "title": "Title",
    "auteur": "Author",
    "author": "Author",
    "date": "Date",
    "date de publication": "Publication Date",
    "publication date": "Publication Date",
    "nom": "Name",
    "name": "Name",
    "description": "Description",
    "resume": "Summary",
    "summary": "Summary",
    "url": "URL",
    "lien": "URL",
    "source": "Source",
    "langue": "Language",
    "language": "Language",
    "categorie": "Category",
    "category": "Category",
    "mots cles": "Keywords",
    "keywords": "Keywords",
    "id": "ID",
}


def normalize_text(value: str) -> str:
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def get_suggested_column_group(column_name: str) -> str:
    normalized = normalize_text(column_name)
    return BILINGUAL_COLUMN_HINTS.get(normalized, column_name)


def get_suggested_sheet_groups(sheet_names: list[str]) -> dict[str, str]:
    normalized_seen = {}
    suggestions = {}
    for sheet_name in sorted(sheet_names):
        normalized = normalize_text(sheet_name)
        if normalized and normalized not in normalized_seen:
            normalized_seen[normalized] = sheet_name
        suggestions[sheet_name] = normalized_seen.get(normalized, sheet_name)
    return suggestions


def resolve_mapping(
    mapping_df: pd.DataFrame, source_column: str, target_column: str
) -> dict[str, str]:
    resolved = {}
    for _, row in mapping_df.iterrows():
        original = str(row[source_column]).strip()
        mapped = str(row[target_column]).strip() if pd.notna(row[target_column]) else ""
        if not original:
            continue
        resolved[original] = mapped or original
    return resolved


def clean_sheet_name(name: object) -> str:
    if pd.isna(name):
        return ""
    return str(name).strip()


def validate_compiled_sheet_names(sheet_names: list[object]) -> list[str]:
    errors = []
    stripped_names = [clean_sheet_name(name) for name in sheet_names]
    clean_names = [name for name in stripped_names if name]

    invalid_names = sorted(name for name in clean_names if INVALID_SHEET_NAME_CHARS.search(name))
    if invalid_names:
        errors.append(
            "Compiled sheet names cannot contain these Excel characters: : \\ / ? * [ ]. "
            f"Invalid names: {', '.join(invalid_names)}."
        )

    blank_count = len(stripped_names) - len(clean_names)
    if blank_count:
        errors.append("Every source sheet must map to a non-empty compiled sheet name.")

    unique_names = list(dict.fromkeys(clean_names))
    truncated_names = [name[:EXCEL_SHEET_NAME_LIMIT] for name in unique_names]
    truncation_collisions = sorted(
        name for name, count in Counter(truncated_names).items() if count > 1
    )
    if truncation_collisions:
        errors.append(
            "Some compiled sheet names become duplicates after Excel's 31-character limit. "
            f"Shorten these names: {', '.join(truncation_collisions)}."
        )

    return errors


def coalesce_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    duplicate_columns = [
        column for column in df.columns if list(df.columns).count(column) > 1
    ]
    if not duplicate_columns:
        return df

    coalesced_columns = {}
    ordered_columns = list(dict.fromkeys(df.columns))
    for column in ordered_columns:
        matching = df.loc[:, df.columns == column]
        if matching.shape[1] == 1:
            coalesced_columns[column] = matching.iloc[:, 0]
        else:
            coalesced_columns[column] = matching.bfill(axis=1).iloc[:, 0]

    return pd.DataFrame(coalesced_columns)


def merge_dataframes(
    dataframes: list[pd.DataFrame], column_mapping: dict[str, str]
) -> pd.DataFrame:
    renamed_dfs = []
    for df in dataframes:
        rename_map = {
            col: column_mapping.get(col, col)
            for col in df.columns
            if col != "Source_File"
        }
        renamed_dfs.append(coalesce_duplicate_columns(df.rename(columns=rename_map)))

    return pd.concat(renamed_dfs, axis=0, join="outer", ignore_index=True)


def parse_workbook(
    file_name: str, file_bytes: bytes
) -> tuple[dict[str, pd.DataFrame], str | None]:
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        parsed_sheets = {}
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            df["Source_File"] = file_name
            parsed_sheets[sheet] = df
        return parsed_sheets, None
    except Exception as exc:
        return {}, f"{file_name}: {exc}"
