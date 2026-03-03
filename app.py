# ================================
# SIMPLE EXCEL MERGE APP
# ================================

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Merger", layout="wide")

st.title("📊 Excel File Merger")
st.write("Upload multiple Excel files and merge them into one master file.")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []
    progress = st.progress(0)

    for i, file in enumerate(uploaded_files):

        try:
            # Read all sheets
            excel_data = pd.read_excel(file, sheet_name=None)

            for sheet_name, df in excel_data.items():
                df["Source_File"] = file.name
                df["Sheet_Name"] = sheet_name
                all_data.append(df)

        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

        progress.progress((i + 1) / len(uploaded_files))

    if all_data:

        merged_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Merge completed successfully!")
        st.write("Preview:")
        st.dataframe(merged_df.head(20), use_container_width=True)

        st.write(f"Total rows: {len(merged_df)}")
        st.write(f"Total columns: {len(merged_df.columns)}")

        # Export Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Merged_Data")

        output.seek(0)

        st.download_button(
            label="📥 Download Merged Excel File",
            data=output,
            file_name="Merged_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload at least one Excel file.")
