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

    for file in uploaded_files:
        parsed_sheets, parse_error = cached_parse_workbook(file.name, file.getvalue())
        if parse_error:
            parse_errors.append(parse_error)
            continue

        for sheet, df in parsed_sheets.items():
            source_key = f"sheet_{len(source_sheet_records)}"
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
    sheet_mapping_seed = pd.DataFrame(
        [
            {
                "Include": False,
                "Source File": record["source_file"],
                "Original Sheet": record["original_sheet"],
                "Compiled Sheet": suggested_sheet_groups[record["original_sheet"]],
                "Source Key": record["source_key"],
            }
            for record in sorted(
                source_sheet_records,
                key=lambda record: (
                    record["source_file"].lower(),
                    record["original_sheet"].lower(),
                ),
            )
        ]
    )

    select_all, clear_all = st.columns(2)
    with select_all:
        if st.button("Select all worksheets"):
            sheet_mapping_seed["Include"] = True
            st.session_state["sheet_mapping_editor_revision"] = (
                st.session_state.get("sheet_mapping_editor_revision", 0) + 1
            )
    with clear_all:
        if st.button("Clear worksheet selections"):
            sheet_mapping_seed["Include"] = False
            st.session_state["sheet_mapping_editor_revision"] = (
                st.session_state.get("sheet_mapping_editor_revision", 0) + 1
            )

    sheet_mapping_editor_key = (
        f"sheet_mapping_editor_{st.session_state.get('sheet_mapping_editor_revision', 0)}"
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

    included_sheet_mapping = edited_sheet_mapping[
        edited_sheet_mapping["Include"].fillna(False)
    ]

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
