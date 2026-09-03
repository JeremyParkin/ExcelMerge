<<<<<<< ours
import hashlib
=======
>>>>>>> theirs
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

<<<<<<< ours
import pandas as pd
import streamlit as st

from merger import (
    EXCEL_SHEET_NAME_LIMIT,
    MAX_COMPILED_SHEETS,
    get_suggested_column_group,
    get_suggested_sheet_groups,
    merge_dataframes,
    parse_workbook,
    resolve_mapping,
    validate_compiled_sheet_names,
    column_matches_query,
    worksheet_matches_query,
)


cached_parse_workbook = st.cache_data(show_spinner=False)(parse_workbook)


st.title("Multi-Sheet Excel Merger")

=======
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


def get_or_init_mapping_df(
    state_key: str,
    seed_df: pd.DataFrame,
    source_column: str,
    target_column: str,
) -> pd.DataFrame:
    """Keep editable mapping state stable across reruns while picking up new source fields."""
    if state_key not in st.session_state:
        st.session_state[state_key] = seed_df.copy()
        return st.session_state[state_key]

    current_df = st.session_state[state_key].copy()
    current_df[source_column] = current_df[source_column].astype(str).str.strip()

    current_mapping = {
        row[source_column]: row[target_column]
        for _, row in current_df.iterrows()
        if str(row[source_column]).strip()
    }

    refreshed_rows = []
    for _, row in seed_df.iterrows():
        source_value = str(row[source_column]).strip()
        default_target = row[target_column]
        refreshed_rows.append(
            {
                source_column: source_value,
                target_column: current_mapping.get(source_value, default_target),
            }
        )

    refreshed_df = pd.DataFrame(refreshed_rows)
    st.session_state[state_key] = refreshed_df
    return refreshed_df


>>>>>>> theirs
uploaded_files = st.file_uploader(
    "Upload multiple Excel workbooks", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
<<<<<<< ours
    source_sheet_data: dict[str, pd.DataFrame] = {}
    source_sheet_records = []

    parse_errors = []

    for file_index, file in enumerate(uploaded_files):
        file_bytes = file.getvalue()
        file_digest = hashlib.sha256(file_bytes).hexdigest()[:12]
        parsed_sheets, parse_error = cached_parse_workbook(file.name, file_bytes)
        if parse_error:
            parse_errors.append(parse_error)
            continue

        for sheet, df in parsed_sheets.items():
            source_key = f"sheet_{file_index}_{file_digest}_{len(source_sheet_records)}"
            source_sheet_data[source_key] = df
            source_sheet_records.append(
                {
                    "source_key": source_key,
                    "source_file": file.name,
                    "original_sheet": sheet,
                }
            )

    for parse_error in parse_errors:
        st.error(parse_error)

    if not source_sheet_data:
        st.stop()

    all_sheets = sorted({record["original_sheet"] for record in source_sheet_records})

    st.subheader("Worksheet Selection & Grouping")
    st.caption(
        f"Choose which uploaded worksheets to include, then map them into up to {MAX_COMPILED_SHEETS} compiled output sheets."
    )

    suggested_sheet_groups = get_suggested_sheet_groups(all_sheets)

    selected_by_source_key = st.session_state.setdefault("selected_by_source_key", {})
    compiled_sheet_by_source_key = st.session_state.setdefault(
        "compiled_sheet_by_source_key", {}
    )
    active_source_keys = {record["source_key"] for record in source_sheet_records}
    for source_key in list(selected_by_source_key):
        if source_key not in active_source_keys:
            del selected_by_source_key[source_key]
    for source_key in list(compiled_sheet_by_source_key):
        if source_key not in active_source_keys:
            del compiled_sheet_by_source_key[source_key]

    for record in source_sheet_records:
        selected_by_source_key.setdefault(record["source_key"], False)
        compiled_sheet_by_source_key.setdefault(
            record["source_key"], suggested_sheet_groups[record["original_sheet"]]
        )

    worksheet_query = st.text_input(
        "Filter worksheets",
        placeholder="Type a sheet or file name",
        key="worksheet_filter_query",
    )
    visible_source_sheet_records = [
        record
        for record in source_sheet_records
        if worksheet_matches_query(
            record["source_file"], record["original_sheet"], worksheet_query
        )
    ]

    selected_count = sum(
        1 for record in source_sheet_records if selected_by_source_key[record["source_key"]]
    )
    st.caption(
        f"Showing {len(visible_source_sheet_records)} of {len(source_sheet_records)} worksheets. "
        f"{selected_count} selected."
    )

    select_visible, clear_visible, select_all, clear_all = st.columns(4)
    with select_visible:
        if st.button("Select visible"):
            for record in visible_source_sheet_records:
                selected_by_source_key[record["source_key"]] = True
            st.session_state["sheet_mapping_editor_revision"] = (
                st.session_state.get("sheet_mapping_editor_revision", 0) + 1
            )
            st.rerun()
    with clear_visible:
        if st.button("Clear visible"):
            for record in visible_source_sheet_records:
                selected_by_source_key[record["source_key"]] = False
            st.session_state["sheet_mapping_editor_revision"] = (
                st.session_state.get("sheet_mapping_editor_revision", 0) + 1
            )
            st.rerun()
    with select_all:
        if st.button("Select all"):
            for record in source_sheet_records:
                selected_by_source_key[record["source_key"]] = True
            st.session_state["sheet_mapping_editor_revision"] = (
                st.session_state.get("sheet_mapping_editor_revision", 0) + 1
            )
            st.rerun()
    with clear_all:
        if st.button("Clear all"):
            for record in source_sheet_records:
                selected_by_source_key[record["source_key"]] = False
            st.session_state["sheet_mapping_editor_revision"] = (
                st.session_state.get("sheet_mapping_editor_revision", 0) + 1
            )
            st.rerun()

    sheet_mapping_rows = [
        {
            "Include": selected_by_source_key[record["source_key"]],
            "Source File": record["source_file"],
            "Original Sheet": record["original_sheet"],
            "Compiled Sheet": compiled_sheet_by_source_key[record["source_key"]],
            "Source Key": record["source_key"],
        }
        for record in sorted(
            visible_source_sheet_records,
            key=lambda record: (
                record["source_file"].lower(),
                record["original_sheet"].lower(),
            ),
        )
    ]
    sheet_mapping_seed = pd.DataFrame(
        sheet_mapping_rows,
        columns=[
            "Include",
            "Source File",
            "Original Sheet",
            "Compiled Sheet",
            "Source Key",
        ],
    )

    sheet_mapping_editor_key = (
        "sheet_mapping_editor_"
        f"{worksheet_query}_"
        f"{st.session_state.get('sheet_mapping_editor_revision', 0)}"
    )

    edited_sheet_mapping = st.data_editor(
        sheet_mapping_seed,
        column_config={
            "Include": st.column_config.CheckboxColumn("Include"),
            "Source File": st.column_config.TextColumn("Source File", disabled=True),
            "Original Sheet": st.column_config.TextColumn("Original Sheet", disabled=True),
            "Compiled Sheet": st.column_config.TextColumn("Compiled Sheet", required=True),
            "Source Key": None,
        },
        disabled=["Source File", "Original Sheet"],
        hide_index=True,
        use_container_width=True,
        key=sheet_mapping_editor_key,
    )

    for _, row in edited_sheet_mapping.iterrows():
        source_key = row["Source Key"]
        selected_by_source_key[source_key] = bool(row["Include"])
        compiled_sheet_by_source_key[source_key] = row["Compiled Sheet"]

    included_sheet_mapping = pd.DataFrame(
        [
            {
                "Source Key": record["source_key"],
                "Compiled Sheet": compiled_sheet_by_source_key[record["source_key"]],
            }
            for record in source_sheet_records
            if selected_by_source_key[record["source_key"]]
        ]
    )

    if included_sheet_mapping.empty:
        st.info("Select at least one worksheet to merge.")
        st.stop()

    sheet_group_mapping = resolve_mapping(
        included_sheet_mapping, source_column="Source Key", target_column="Compiled Sheet"
    )

    grouped_sheet_data: dict[str, list[pd.DataFrame]] = {}
    for source_key, compiled_sheet in sheet_group_mapping.items():
        grouped_sheet_data.setdefault(compiled_sheet, []).append(source_sheet_data[source_key])

    compiled_sheet_names = [
        name
        for name in included_sheet_mapping["Compiled Sheet"].astype(str).str.strip().tolist()
        if name
    ]
    unique_compiled_sheet_names = list(dict.fromkeys(compiled_sheet_names))
    sheet_name_errors = validate_compiled_sheet_names(
        included_sheet_mapping["Compiled Sheet"].tolist()
    )

    for sheet_name_error in sheet_name_errors:
        st.error(sheet_name_error)

    if sheet_name_errors:
        st.stop()

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
        compiled_sheet_included_columns: dict[str, set[str]] = {}
        compiled_sheets_without_columns = []

        for compiled_sheet in selected_compiled_sheets:
            with st.expander(f"Column mapping for '{compiled_sheet}'", expanded=False):
                all_columns = []
                for df in grouped_sheet_data[compiled_sheet]:
                    all_columns.extend([col for col in df.columns if col != "Source_File"])

                unique_columns = sorted(set(all_columns))
                included_column_key = f"included_columns_{compiled_sheet}"
                output_column_key = f"output_columns_{compiled_sheet}"
                included_columns = st.session_state.setdefault(included_column_key, {})
                output_columns = st.session_state.setdefault(output_column_key, {})

                active_columns = set(unique_columns)
                for column in list(included_columns):
                    if column not in active_columns:
                        del included_columns[column]
                for column in list(output_columns):
                    if column not in active_columns:
                        del output_columns[column]

                for column in unique_columns:
                    included_columns.setdefault(column, True)
                    output_columns.setdefault(column, get_suggested_column_group(column))

                column_query = st.text_input(
                    "Filter columns",
                    placeholder="Type a source or output column name",
                    key=f"column_filter_query_{compiled_sheet}",
                )
                visible_columns = [
                    column
                    for column in unique_columns
                    if column_matches_query(column, output_columns[column], column_query)
                ]

                selected_column_count = sum(
                    1 for column in unique_columns if included_columns[column]
                )
                st.caption(
                    f"Showing {len(visible_columns)} of {len(unique_columns)} columns. "
                    f"{selected_column_count} included."
                )

                (
                    select_visible_columns,
                    clear_visible_columns,
                    select_all_columns,
                    clear_all_columns,
                ) = st.columns(4)
                with select_visible_columns:
                    if st.button("Select visible", key=f"select_visible_columns_{compiled_sheet}"):
                        for column in visible_columns:
                            included_columns[column] = True
                        st.session_state[f"column_mapping_revision_{compiled_sheet}"] = (
                            st.session_state.get(f"column_mapping_revision_{compiled_sheet}", 0) + 1
                        )
                        st.rerun()
                with clear_visible_columns:
                    if st.button("Clear visible", key=f"clear_visible_columns_{compiled_sheet}"):
                        for column in visible_columns:
                            included_columns[column] = False
                        st.session_state[f"column_mapping_revision_{compiled_sheet}"] = (
                            st.session_state.get(f"column_mapping_revision_{compiled_sheet}", 0) + 1
                        )
                        st.rerun()
                with select_all_columns:
                    if st.button("Select all", key=f"select_all_columns_{compiled_sheet}"):
                        for column in unique_columns:
                            included_columns[column] = True
                        st.session_state[f"column_mapping_revision_{compiled_sheet}"] = (
                            st.session_state.get(f"column_mapping_revision_{compiled_sheet}", 0) + 1
                        )
                        st.rerun()
                with clear_all_columns:
                    if st.button("Clear all", key=f"clear_all_columns_{compiled_sheet}"):
                        for column in unique_columns:
                            included_columns[column] = False
                        st.session_state[f"column_mapping_revision_{compiled_sheet}"] = (
                            st.session_state.get(f"column_mapping_revision_{compiled_sheet}", 0) + 1
                        )
                        st.rerun()

                mapping_seed = pd.DataFrame(
                    [
                        {
                            "Include": included_columns[column],
                            "Original Column": column,
                            "Output Column": output_columns[column],
                        }
                        for column in visible_columns
                    ],
                    columns=["Include", "Original Column", "Output Column"],
                )

                edited_mapping = st.data_editor(
                    mapping_seed,
                    column_config={
                        "Include": st.column_config.CheckboxColumn("Include"),
                        "Original Column": st.column_config.TextColumn(
                            "Original Column", disabled=True
                        ),
                        "Output Column": st.column_config.TextColumn("Output Column", required=True),
                    },
                    disabled=["Original Column"],
                    hide_index=True,
                    use_container_width=True,
                    key=(
                        f"column_mapping_editor_{compiled_sheet}_{column_query}_"
                        f"{st.session_state.get(f'column_mapping_revision_{compiled_sheet}', 0)}"
                    ),
                )

                for _, row in edited_mapping.iterrows():
                    column = row["Original Column"]
                    included_columns[column] = bool(row["Include"])
                    output_columns[column] = row["Output Column"]

                included_mapping = pd.DataFrame(
                    [
                        {
                            "Original Column": column,
                            "Output Column": output_columns[column],
                        }
                        for column in unique_columns
                        if included_columns[column]
                    ]
                )

                if included_mapping.empty:
                    st.info("Select at least one column for this compiled sheet.")
                    compiled_sheet_column_mappings[compiled_sheet] = {}
                    compiled_sheet_included_columns[compiled_sheet] = set()
                    compiled_sheets_without_columns.append(compiled_sheet)
                else:
                    compiled_sheet_column_mappings[compiled_sheet] = resolve_mapping(
                        included_mapping,
                        source_column="Original Column",
                        target_column="Output Column",
                    )
                    compiled_sheet_included_columns[compiled_sheet] = set(
                        included_mapping["Original Column"]
                    )

        if compiled_sheets_without_columns:
            st.stop()

        st.subheader("Merged Sheet Previews")
        tabs = st.tabs(selected_compiled_sheets)

        for tab, compiled_sheet in zip(tabs, selected_compiled_sheets):
            with tab:
                merged_sheets[compiled_sheet] = merge_dataframes(
                    grouped_sheet_data[compiled_sheet],
                    compiled_sheet_column_mappings[compiled_sheet],
                    compiled_sheet_included_columns[compiled_sheet],
                )
                st.write(merged_sheets[compiled_sheet].head(10))

    if selected_compiled_sheets and merged_sheets:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in merged_sheets.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name[:EXCEL_SHEET_NAME_LIMIT])
=======
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

    sheet_mapping_state_key = "sheet_mapping_df"
    editable_sheet_mapping = get_or_init_mapping_df(
        state_key=sheet_mapping_state_key,
        seed_df=sheet_mapping_seed,
        source_column="Original Sheet",
        target_column="Compiled Sheet",
    )

    if st.button("Reset sheet mapping to suggestions"):
        st.session_state[sheet_mapping_state_key] = sheet_mapping_seed.copy()
        editable_sheet_mapping = st.session_state[sheet_mapping_state_key]

    edited_sheet_mapping = st.data_editor(
        editable_sheet_mapping,
        hide_index=True,
        use_container_width=True,
        key="sheet_mapping_editor",
        disabled=["Original Sheet"],
    )
    st.session_state[sheet_mapping_state_key] = edited_sheet_mapping

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

                column_mapping_state_key = f"column_mapping_df_{compiled_sheet}"
                editable_column_mapping = get_or_init_mapping_df(
                    state_key=column_mapping_state_key,
                    seed_df=mapping_seed,
                    source_column="Original Column",
                    target_column="Match Group",
                )

                if st.button(
                    f"Reset column mapping for '{compiled_sheet}'",
                    key=f"reset_column_mapping_{compiled_sheet}",
                ):
                    st.session_state[column_mapping_state_key] = mapping_seed.copy()
                    editable_column_mapping = st.session_state[column_mapping_state_key]

                edited_mapping = st.data_editor(
                    editable_column_mapping,
                    hide_index=True,
                    use_container_width=True,
                    key=f"column_mapping_editor_{compiled_sheet}",
                    disabled=["Original Column"],
                )
                st.session_state[column_mapping_state_key] = edited_mapping

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
>>>>>>> theirs
        output.seek(0)

        st.download_button(
            label="Compile & Download Excel File",
            data=output,
            file_name="merged_workbooks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
