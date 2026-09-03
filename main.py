import hashlib
import io

import pandas as pd
import streamlit as st

from merger import (
    EXCEL_SHEET_NAME_LIMIT,
    MERGED_SHEET_NAME,
    column_matches_query,
    get_suggested_column_group,
    merge_dataframes,
    parse_uploaded_file,
    resolve_mapping,
    worksheet_matches_query,
)


cached_parse_uploaded_file = st.cache_data(show_spinner=False)(parse_uploaded_file)


def bump_revision(key: str) -> None:
    st.session_state[key] = st.session_state.get(key, 0) + 1


def read_checkbox_value(value: object) -> bool:
    return bool(value) if pd.notna(value) else False


st.title("Excel & CSV Row Stacker")

uploaded_files = st.file_uploader(
    "Upload Excel workbooks or CSV files",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    source_sheet_data: dict[str, pd.DataFrame] = {}
    source_sheet_records = []
    parse_errors = []

    for file_index, file in enumerate(uploaded_files):
        file_bytes = file.getvalue()
        file_digest = hashlib.sha256(file_bytes).hexdigest()[:12]
        parsed_sheets, parse_error = cached_parse_uploaded_file(file.name, file_bytes)
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

    st.subheader("Source Selection")
    st.caption(
        "Choose the CSV files or Excel worksheets to stack into one downloaded worksheet."
    )

    selected_by_source_key = st.session_state.setdefault("selected_by_source_key", {})

    active_source_keys = {record["source_key"] for record in source_sheet_records}
    for source_key in list(selected_by_source_key):
        if source_key not in active_source_keys:
            del selected_by_source_key[source_key]

    for record in source_sheet_records:
        source_key = record["source_key"]
        selected_by_source_key.setdefault(source_key, False)

    worksheet_query = st.text_input(
        "Filter sources",
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
        f"Showing {len(visible_source_sheet_records)} of {len(source_sheet_records)} sources. "
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
            "Sheet": record["original_sheet"],
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
            "Sheet",
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
            "Sheet": st.column_config.TextColumn("Sheet", disabled=True),
            "Source Key": None,
        },
        disabled=["Source File", "Sheet"],
        hide_index=True,
        use_container_width=True,
        key=sheet_mapping_editor_key,
    )

    for _, row in edited_sheet_mapping.iterrows():
        source_key = row["Source Key"]
        selected_by_source_key[source_key] = read_checkbox_value(row["Include"])

    included_source_records = [
        record
        for record in source_sheet_records
        if selected_by_source_key[record["source_key"]]
    ]

    if not included_source_records:
        st.info("Select at least one source to merge.")
        st.stop()

    selected_dataframes = [
        source_sheet_data[record["source_key"]]
        for record in included_source_records
    ]

    st.subheader("Column Matching")
    st.caption(
        "Choose which columns to keep, and map columns that should be treated as the same field."
    )

    all_columns = []
    for df in selected_dataframes:
        all_columns.extend([col for col in df.columns if col != "Source_File"])

    unique_columns = sorted(set(all_columns))
    included_columns = st.session_state.setdefault("included_columns", {})
    output_columns = st.session_state.setdefault("output_columns", {})

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
        key="column_filter_query",
    )
    visible_columns = [
        column
        for column in unique_columns
        if column_matches_query(column, output_columns[column], column_query)
    ]

    selected_column_count = sum(1 for column in unique_columns if included_columns[column])
    st.caption(
        f"Showing {len(visible_columns)} of {len(unique_columns)} columns. "
        f"{selected_column_count} included."
    )

    select_visible_columns, clear_visible_columns, select_all_columns, clear_all_columns = st.columns(4)
    with select_visible_columns:
        if st.button("Select visible", key="select_visible_columns"):
            for column in visible_columns:
                included_columns[column] = True
            bump_revision("column_mapping_revision")
            st.rerun()
    with clear_visible_columns:
        if st.button("Clear visible", key="clear_visible_columns"):
            for column in visible_columns:
                included_columns[column] = False
            bump_revision("column_mapping_revision")
            st.rerun()
    with select_all_columns:
        if st.button("Select all", key="select_all_columns"):
            for column in unique_columns:
                included_columns[column] = True
            bump_revision("column_mapping_revision")
            st.rerun()
    with clear_all_columns:
        if st.button("Clear all", key="clear_all_columns"):
            for column in unique_columns:
                included_columns[column] = False
            bump_revision("column_mapping_revision")
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
            f"column_mapping_editor_{column_query}_"
            f"{st.session_state.get('column_mapping_revision', 0)}"
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
        st.info("Select at least one column to merge.")
        st.stop()

    column_mapping = resolve_mapping(
        included_mapping,
        source_column="Original Column",
        target_column="Output Column",
    )
    included_column_names = set(included_mapping["Original Column"])

    st.subheader("Merged Preview")
    merged_sheet = merge_dataframes(
        selected_dataframes,
        column_mapping,
        included_column_names,
    )
    st.write(merged_sheet.head(10))

    if not merged_sheet.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged_sheet.to_excel(
                writer,
                index=False,
                sheet_name=MERGED_SHEET_NAME[:EXCEL_SHEET_NAME_LIMIT],
            )
        output.seek(0)

        st.download_button(
            label="Compile & Download Excel File",
            data=output,
            file_name="merged_sources.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
