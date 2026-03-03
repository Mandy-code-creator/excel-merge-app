import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Merge Excel Files", layout="wide")
st.title("📊 Merge Multiple Excel Files (.xlsx only)")

uploaded_files = st.file_uploader(
    "Chọn nhiều file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

dfs = []

if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, engine="openpyxl")
            dfs.append(df)
        except Exception as e:
            st.error(f"Lỗi khi đọc file {file.name}: {e}")

    if dfs:
        df_all = pd.concat(dfs, ignore_index=True)
        st.success(f"✅ Gộp thành công {len(dfs)} file, tổng {df_all.shape[0]} dòng")
        st.dataframe(df_all)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_all.to_excel(writer, index=False, sheet_name="Sheet1")
        output.seek(0)

        st.download_button(
            label="📥 Tải file Excel đã gộp",
            data=output,
            file_name="gop_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
