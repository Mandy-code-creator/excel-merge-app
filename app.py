import streamlit as st
import pandas as pd
import gc

st.set_page_config(page_title="Merge Excel Files", layout="wide")
st.title("📊 Merge Multiple Excel Files (Bản Xuất CSV Chống Sập)")

st.warning("⚠️ Đã chuyển đổi sang định dạng CSV để bảo vệ RAM server không bị sập.")

# Upload nhiều file Excel
uploaded_files = st.file_uploader(
    "Chọn nhiều file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if 'csv_ready' not in st.session_state:
    st.session_state['csv_ready'] = False
    st.session_state['csv_data'] = None

if uploaded_files:
    if st.button("🚀 Tiến hành gộp file"):
        with st.spinner("Đang xử lý dữ liệu..."):
            dfs = []
            for file in uploaded_files:
                try:
                    # Đọc bằng calamine siêu nhẹ
                    df = pd.read_excel(file, engine="calamine")
                    dfs.append(df)
                except Exception as e:
                    st.error(f"Lỗi khi đọc file {file.name}: {e}")

            if dfs:
                df_all = pd.concat(dfs, ignore_index=True)
                st.success(f"✅ Gộp thành công {len(dfs)} file, tổng {df_all.shape[0]} dòng")
                
                # Xem trước 50 dòng
                st.dataframe(df_all.head(50))

                # 🔥 XUẤT THẲNG RA CSV - KHÔNG DÙNG OPENPYXL NÊN CỰC KỲ KHỎE RAM
                # utf-8-sig giúp khi mở file CSV này bằng phần mềm Excel trên máy tính không bị lỗi font Tiếng Việt/Trung
                st.session_state['csv_data'] = df_all.to_csv(index=False).encode('utf-8-sig')
                st.session_state['csv_ready'] = True

                # Giải phóng bộ nhớ ngay lập tức
                del dfs
                del df_all
                gc.collect()

# Nút download an toàn
if st.session_state['csv_ready'] and st.session_state['csv_data']:
    st.download_button(
        label="📥 Tải file kết quả đã gộp (.CSV)",
        data=st.session_state['csv_data'],
        file_name="gop_file_hoanthanh.csv",
        mime="text/csv"
    )
