import streamlit as st
import pandas as pd
import gc
import os

st.set_page_config(page_title="Merge Excel Files", layout="wide")
st.title("📊 Merge Multiple Excel Files (Bản Siêu Tối Ưu)")

# Upload nhiều file Excel
uploaded_files = st.file_uploader(
    "Chọn nhiều file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Sử dụng Session State để lưu trạng thái, tránh việc code chạy lại từ đầu khi bấm Download
if 'file_ready' not in st.session_state:
    st.session_state['file_ready'] = False

if uploaded_files:
    # Nút bấm để chủ động gộp, không gộp tự động
    if st.button("🚀 Bắt đầu gộp file"):
        with st.spinner("Đang xử lý và tối ưu bộ nhớ, vui lòng đợi..."):
            dfs = []
            for file in uploaded_files:
                try:
                    # Đọc bằng calamine cực nhẹ RAM
                    df = pd.read_excel(file, engine="calamine")
                    dfs.append(df)
                except Exception as e:
                    st.error(f"Lỗi khi đọc file {file.name}: {e}")

            if dfs:
                # Gộp tất cả DataFrame
                df_all = pd.concat(dfs, ignore_index=True)
                st.success(f"✅ Gộp thành công {len(dfs)} file, tổng {df_all.shape[0]} dòng")
                
                # 🔥 LƯU THẲNG RA Ổ ĐĨA TẠM CỦA SERVER THAY VÌ GIỮ TRONG RAM
                temp_file_path = "gop_file_hoanthanh.csv"
                df_all.to_csv(temp_file_path, index=False, encoding='utf-8-sig')
                
                # 🔥 DỌN SẠCH RAM NGAY LẬP TỨC TRƯỚC KHI TẠO NÚT TẢI
                del dfs
                del df_all
                gc.collect() 

                # Đánh dấu là file đã tạo xong
                st.session_state['file_ready'] = True
                st.session_state['file_path'] = temp_file_path

# Nút download chỉ hiện ra khi file đã được ghi xong vào ổ cứng
if st.session_state.get('file_ready') and os.path.exists(st.session_state.get('file_path', '')):
    st.info("💡 File đã sẵn sàng để tải về. RAM đã được giải phóng để chống sập server.")
    
    # Đọc trực tiếp từ ổ cứng, Streamlit xử lý việc này rất tiết kiệm RAM
    with open(st.session_state['file_path'], "rb") as file_to_download:
        st.download_button(
            label="📥 Tải file đã gộp (Định dạng .CSV)",
            data=file_to_download,
            file_name="gop_file_hoanthanh.csv",
            mime="text/csv"
        )
