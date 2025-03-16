import streamlit as st
import pandas as pd
import io

# Streamlit UI
st.title("Multi-Sheet Excel Merger")

# Step 1: File Upload
uploaded_files = st.file_uploader("Upload multiple Excel workbooks", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    sheet_data = {}  # Store raw dataframes per sheet name

    # Step 2: Extract sheet names from all uploaded files (no merging yet)
    for file in uploaded_files:
        xls = pd.ExcelFile(file, engine="openpyxl")
        for sheet in xls.sheet_names:
            if sheet not in sheet_data:
                sheet_data[sheet] = []
            df = xls.parse(sheet)
            df["Source_File"] = file.name  # Track which file the data came from
            sheet_data[sheet].append(df)

    # Default sheets to be pre-selected if they exist
    default_selected_sheets = {"CLEAN TRAD", "Authors", "Top Stories", "Clean Social"}

    # Step 3: Multiselect for choosing sheets (pre-select the default ones if they exist)
    all_sheets = list(sheet_data.keys())
    preselected_sheets = [sheet for sheet in all_sheets if sheet in default_selected_sheets]

    selected_sheets = st.multiselect(
        "Select sheets to merge:", options=all_sheets, default=preselected_sheets
    )

    # Step 4: Show previews in tabs (Only merge on-demand)
    if selected_sheets:
        st.subheader("Merged Sheet Previews")
        tab_titles = selected_sheets  # Use sheet names as tab titles
        tabs = st.tabs(tab_titles)

        merged_sheets = {}

        for tab, sheet in zip(tabs, selected_sheets):
            with tab:
                # Merge only when the tab is opened
                if sheet not in merged_sheets:
                    merged_sheets[sheet] = pd.concat(sheet_data[sheet], axis=0, join="outer", ignore_index=True)

                # Display merged preview (only 10 rows to optimize performance)
                st.write(merged_sheets[sheet].head(10))

    # Step 5: Create downloadable Excel file only when needed
    if st.button("Merge & Download"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df in merged_sheets.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        output.seek(0)

        st.download_button(
            label="Download Merged Excel File",
            data=output,
            file_name="merged_workbooks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
