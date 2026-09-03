import io
import re
import unicodedata

import pandas as pd


EXCEL_SHEET_NAME_LIMIT = 31
MERGED_SHEET_NAME = "Merged Data"
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


def worksheet_matches_query(source_file: str, sheet_name: str, query: str) -> bool:
    normalized_query = normalize_text(query)
    if not normalized_query:
        return True

    searchable_text = normalize_text(f"{source_file} {sheet_name}")
    return all(token in searchable_text for token in normalized_query.split())


def column_matches_query(column_name: str, output_column: str, query: str) -> bool:
    normalized_query = normalize_text(query)
    if not normalized_query:
        return True

    searchable_text = normalize_text(f"{column_name} {output_column}")
    return all(token in searchable_text for token in normalized_query.split())


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
    dataframes: list[pd.DataFrame],
    column_mapping: dict[str, str],
    included_columns: set[str] | None = None,
) -> pd.DataFrame:
    renamed_dfs = []
    for df in dataframes:
        if included_columns is None:
            columns_to_keep = list(df.columns)
        else:
            columns_to_keep = [
                col for col in df.columns if col == "Source_File" or col in included_columns
            ]

        rename_map = {
            col: column_mapping.get(col, col)
            for col in columns_to_keep
            if col != "Source_File"
        }
        renamed_dfs.append(
            coalesce_duplicate_columns(df[columns_to_keep].rename(columns=rename_map))
        )

    return pd.concat(renamed_dfs, axis=0, join="outer", ignore_index=True)


def parse_uploaded_file(
    file_name: str, file_bytes: bytes
) -> tuple[dict[str, pd.DataFrame], str | None]:
    try:
        if file_name.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
            df["Source_File"] = file_name
            return {file_name: df}, None

        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        parsed_sheets = {}
        for sheet in xls.sheet_names:
            df = xls.parse(sheet)
            df["Source_File"] = file_name
            parsed_sheets[sheet] = df
        return parsed_sheets, None
    except Exception as exc:
        return {}, f"{file_name}: {exc}"
