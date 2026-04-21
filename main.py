import io
import re
import unicodedata

import pandas as pd
import streamlit as st


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


st.title("Multi-Sheet Excel Merger")


def normalize_column_name(value: str) -> str:
    """Normalize a column name to improve matching across language/casing differences."""
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def get_suggested_group(column_name: str) -> str:
    normalized = normalize_column_name(column_name)
    return BILINGUAL_COLUMN_HINTS.get(normalized, column_name)


def resolve_column_mapping(mapping_df: pd.DataFrame) -> dict[str, str]:
    resolved = {}
    for _, row in mapping_df.iterrows():
        original = str(row["Original Column"])
        mapped = str(row["Match Group"]).strip() if pd.notna(row["Match Group"]) else ""
        resolved[original] = mapped or original
    return resolved


uploaded_files = st.file_uploader(
    "Upload multiple Excel workbooks", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    sheet_data: dict[str, list[pd.DataFrame]] = {}

    for file in uploaded_files:
        xls = pd.ExcelFile(file, engine="openpyxl")
        for sheet in xls.sheet_names:
            sheet_data.setdefault(sheet, [])
            df = xls.parse(sheet)
            df["Source_File"] = file.name
            sheet_data[sheet].append(df)

    default_selected_sheets = {"CLEAN TRAD", "Authors", "Top Stories", "Clean Social"}

    all_sheets = list(sheet_data.keys())
    preselected_sheets = [sheet for sheet in all_sheets if sheet in default_selected_sheets]

    selected_sheets = st.multiselect(
        "Select sheets to merge:", options=all_sheets, default=preselected_sheets
    )

    merged_sheets = {}

    if selected_sheets:
        st.subheader("Column Matching & Grouping")
        st.caption(
            "Map columns that should be treated as the same field (e.g., English/French names)."
        )

        sheet_column_mappings: dict[str, dict[str, str]] = {}

        for sheet in selected_sheets:
            with st.expander(f"Column mapping for '{sheet}'", expanded=False):
                all_columns = []
                for df in sheet_data[sheet]:
                    all_columns.extend([col for col in df.columns if col != "Source_File"])

                unique_columns = sorted(set(all_columns))
                mapping_seed = pd.DataFrame(
                    {
                        "Original Column": unique_columns,
                        "Match Group": [get_suggested_group(col) for col in unique_columns],
                    }
                )

                st.write(
                    "Update **Match Group** values so columns with the same meaning share the same group name."
                )
                edited_mapping = st.data_editor(
                    mapping_seed,
                    hide_index=True,
                    use_container_width=True,
                    key=f"mapping_editor_{sheet}",
                )

                sheet_column_mappings[sheet] = resolve_column_mapping(edited_mapping)

        st.subheader("Merged Sheet Previews")
        tabs = st.tabs(selected_sheets)

        for tab, sheet in zip(tabs, selected_sheets):
            with tab:
                renamed_dfs = []
                for df in sheet_data[sheet]:
                    rename_map = {
                        col: sheet_column_mappings[sheet].get(col, col)
                        for col in df.columns
                        if col != "Source_File"
                    }
                    renamed_dfs.append(df.rename(columns=rename_map))

                merged_sheets[sheet] = pd.concat(
                    renamed_dfs,
                    axis=0,
                    join="outer",
                    ignore_index=True,
                )
                st.write(merged_sheets[sheet].head(10))

    if selected_sheets and merged_sheets and st.button("Merge & Download"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in merged_sheets.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)

        st.download_button(
            label="Download Merged Excel File",
            data=output,
            file_name="merged_workbooks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
