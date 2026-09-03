import hashlib
import io

import pandas as pd
import streamlit as st

from merger import (
    EXCEL_SHEET_NAME_LIMIT,
    MAX_COMPILED_SHEETS,
    column_matches_query,
    get_suggested_column_group,
    get_suggested_sheet_groups,
    merge_dataframes,
    parse_workbook,
    resolve_mapping,
    validate_compiled_sheet_names,
    worksheet_matches_query,
)


SINGLE_OUTPUT_SHEET_NAME = "Merged Data"

cached_parse_workbook = st.cache_data(show_spinner=False)(parse_workbook)


def bump_revision(key: str) -> None:
    st.session_state[key] = st.session_state.get(key, 0) + 1


def read_checkbox_value(value: object) -> bool:
    return bool(value) if pd.notna(value) else False


st.title("Multi-Sheet Excel Merger")

uploaded_files = st.file_uploader(
    "Upload multiple Excel workbooks", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
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

    output_grouping = st.radio(
        "Default output grouping",
        options=[
            "Stack selected worksheets into one output sheet",
            "Keep matching worksheet names separate",
        ],
        horizontal=True,
        key="output_grouping",
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

    previous_grouping = st.session_state.get("previous_output_grouping")
    if previous_grouping != output_grouping:
        compiled_sheet_by_source_key.clear()
        st.session_state["previous_output_grouping"] = output_grouping
        bump_revision("sheet_mapping_editor_revision")

    for record in source_sheet_records:
        source_key = record["source_key"]
        selected_by_source_key.setdefault(source_key, False)
        default_compiled_sheet = (
            SINGLE_OUTPUT_SHEET_NAME
            if output_grouping == "Stack selected worksheets into one output sheet"
            else suggested_sheet_groups[record["original_sheet"]]
        )
        compiled_sheet_by_source_key.setdefault(source_key, default_compiled_sheet)

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
        1
        for record in source_sheet_records
        if selected_by_source_key[record["source_key"]]
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
            bump_revision("sheet_mapping_editor_revision")
            st.rerun()
    with clear_visible:
        if st.button("Clear visible"):
            for record in visible_source_sheet_records:
                selected_by_source_key[record["source_key"]] = False
            bump_revision("sheet_mapping_editor_revision")
            st.rerun()
    with select_all:
        if st.button("Select all"):
            for record in source_sheet_records:
                selected_by_source_key[record["source_key"]] = True
            bump_revision("sheet_mapping_editor_revision")
            st.rerun()
    with clear_all:
        if st.button("Clear all"):
            for record in source_sheet_records:
                selected_by_source_key[record["source_key"]] = False
            bump_revision("sheet_mapping_editor_revision")
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
        selected_by_source_key[source_key] = read_checkbox_value(row["Include"])
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
            "Choose which columns to keep, and map columns that should be treated as the same field."
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
                        bump_revision(f"column_mapping_revision_{compiled_sheet}")
                        st.rerun()
                with clear_visible_columns:
                    if st.button("Clear visible", key=f"clear_visible_columns_{compiled_sheet}"):
                        for column in visible_columns:
                            included_columns[column] = False
                        bump_revision(f"column_mapping_revision_{compiled_sheet}")
                        st.rerun()
                with select_all_columns:
                    if st.button("Select all", key=f"select_all_columns_{compiled_sheet}"):
                        for column in unique_columns:
                            included_columns[column] = True
                        bump_revision(f"column_mapping_revision_{compiled_sheet}")
                        st.rerun()
                with clear_all_columns:
                    if st.button("Clear all", key=f"clear_all_columns_{compiled_sheet}"):
                        for column in unique_columns:
                            included_columns[column] = False
                        bump_revision(f"column_mapping_revision_{compiled_sheet}")
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
                    included_columns[column] = read_checkbox_value(row["Include"])
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
        output.seek(0)

        st.download_button(
            label="Compile & Download Excel File",
            data=output,
            file_name="merged_workbooks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
