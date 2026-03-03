# ==========================================
# EXCEL MERGE APP - STABLE FULL VERSION
# ==========================================

import streamlit as st
import pandas as pd
import io

# ----------------------------
# PAGE CONFIG
# ----------------------------
st.set_page_config(
    page_title="Excel Merge App",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel File Merger")
st.write("Upload multiple Excel files and merge them into one master Excel file.")

st.divider()

# ----------------------------
# FILE UPLOADER
# ----------------------------
uploaded_files = st.file_uploader(
    "Upload Excel files (.xlsx or .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# ----------------------------
# MAIN LOGIC
# ----------------------------
if uploaded_files:

    st.info(f"Total files uploaded: {len(uploaded_files)}")

    all_data = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):

        try:
            # Read all sheets inside each Excel file
            excel_data = pd.read_excel(file, sheet_name=None)

            for sheet_name, df in excel_data.items():

                # Add source tracking columns
                df["Source_File"] = file.name
                df["Sheet_Name"] = sheet_name

                all_data.append(df)

        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

        progress_bar.progress((i + 1) / len(uploaded_files))

    if all_data:

        # Merge all dataframes
        merged_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Merge completed successfully!")

        # Preview
        st.subheader("Preview Merged Data")
        st.dataframe(merged_df.head(20), use_container_width=True)

        st.write(f"Total Rows: {len(merged_df)}")
        st.write(f"Total Columns: {len(merged_df.columns)}")

        # ----------------------------
        # EXPORT TO EXCEL
        # ----------------------------
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
        st.warning("No valid data to merge.")

else:
    st.info("Please upload at least one Excel file.")
