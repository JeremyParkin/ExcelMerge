import io
import re
import unicodedata

import pandas as pd
import streamlit as st


MAX_COMPILED_SHEETS = 5

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


def resolve_mapping(mapping_df: pd.DataFrame, source_column: str, target_column: str) -> dict[str, str]:
    resolved = {}
    for _, row in mapping_df.iterrows():
        original = str(row[source_column]).strip()
        mapped = str(row[target_column]).strip() if pd.notna(row[target_column]) else ""
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

    all_sheets = sorted(sheet_data.keys())

    st.subheader("Sheet Matching & Grouping")
    st.caption(
        f"Map source sheet name variations into up to {MAX_COMPILED_SHEETS} compiled output sheets."
    )

    suggested_sheet_groups = get_suggested_sheet_groups(all_sheets)
    sheet_mapping_seed = pd.DataFrame(
        {
            "Original Sheet": all_sheets,
            "Compiled Sheet": [suggested_sheet_groups[sheet] for sheet in all_sheets],
        }
    )

    edited_sheet_mapping = st.data_editor(
        sheet_mapping_seed,
        hide_index=True,
        use_container_width=True,
        key="sheet_mapping_editor",
    )

    sheet_group_mapping = resolve_mapping(
        edited_sheet_mapping, source_column="Original Sheet", target_column="Compiled Sheet"
    )

    grouped_sheet_data: dict[str, list[pd.DataFrame]] = {}
    for source_sheet, dfs in sheet_data.items():
        compiled_sheet = sheet_group_mapping.get(source_sheet, source_sheet)
        grouped_sheet_data.setdefault(compiled_sheet, []).extend(dfs)

    compiled_sheet_names = [
        name for name in edited_sheet_mapping["Compiled Sheet"].astype(str).str.strip().tolist() if name
    ]
    unique_compiled_sheet_names = list(dict.fromkeys(compiled_sheet_names))

    if len(unique_compiled_sheet_names) > MAX_COMPILED_SHEETS:
        st.error(
            f"You mapped {len(unique_compiled_sheet_names)} compiled sheets. Please reduce this to {MAX_COMPILED_SHEETS} or fewer."
        )
        st.stop()

    selected_compiled_sheets = st.multiselect(
        "Select compiled sheets to merge:",
        options=unique_compiled_sheet_names,
        default=unique_compiled_sheet_names,
    )

    merged_sheets = {}

    if selected_compiled_sheets:
        st.subheader("Column Matching & Grouping")
        st.caption(
            "Map columns that should be treated as the same field (for example, English/French labels)."
        )

        compiled_sheet_column_mappings: dict[str, dict[str, str]] = {}

        for compiled_sheet in selected_compiled_sheets:
            with st.expander(f"Column mapping for '{compiled_sheet}'", expanded=False):
                all_columns = []
                for df in grouped_sheet_data[compiled_sheet]:
                    all_columns.extend([col for col in df.columns if col != "Source_File"])

                unique_columns = sorted(set(all_columns))
                mapping_seed = pd.DataFrame(
                    {
                        "Original Column": unique_columns,
                        "Match Group": [get_suggested_column_group(col) for col in unique_columns],
                    }
                )

                edited_mapping = st.data_editor(
                    mapping_seed,
                    hide_index=True,
                    use_container_width=True,
                    key=f"column_mapping_editor_{compiled_sheet}",
                )

                compiled_sheet_column_mappings[compiled_sheet] = resolve_mapping(
                    edited_mapping,
                    source_column="Original Column",
                    target_column="Match Group",
                )

        st.subheader("Merged Sheet Previews")
        tabs = st.tabs(selected_compiled_sheets)

        for tab, compiled_sheet in zip(tabs, selected_compiled_sheets):
            with tab:
                renamed_dfs = []
                for df in grouped_sheet_data[compiled_sheet]:
                    rename_map = {
                        col: compiled_sheet_column_mappings[compiled_sheet].get(col, col)
                        for col in df.columns
                        if col != "Source_File"
                    }
                    renamed_dfs.append(df.rename(columns=rename_map))

                merged_sheets[compiled_sheet] = pd.concat(
                    renamed_dfs,
                    axis=0,
                    join="outer",
                    ignore_index=True,
                )
                st.write(merged_sheets[compiled_sheet].head(10))

    if selected_compiled_sheets and merged_sheets and st.button("Merge & Download"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in merged_sheets.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        output.seek(0)

        st.download_button(
            label="Download Merged Excel File",
            data=output,
            file_name="merged_workbooks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
