import hashlib
import io

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
    worksheet_matches_query,
)


cached_parse_workbook = st.cache_data(show_spinner=False)(parse_workbook)


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
                    column_config={
                        "Original Column": st.column_config.TextColumn(
                            "Original Column", disabled=True
                        ),
                        "Match Group": st.column_config.TextColumn("Match Group", required=True),
                    },
                    disabled=["Original Column"],
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
                merged_sheets[compiled_sheet] = merge_dataframes(
                    grouped_sheet_data[compiled_sheet],
                    compiled_sheet_column_mappings[compiled_sheet],
                )
                st.write(merged_sheets[compiled_sheet].head(10))

    if selected_compiled_sheets and merged_sheets and st.button("Merge & Download"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in merged_sheets.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name[:EXCEL_SHEET_NAME_LIMIT])
        output.seek(0)

        st.download_button(
            label="Download Merged Excel File",
            data=output,
            file_name="merged_workbooks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
