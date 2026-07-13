import streamlit as st
import pandas as pd
import io
import gc

st.set_page_config(page_title="Merge Excel Files", layout="wide")
st.title("📊 Merge Multiple Excel Files (.xlsx only)")

# Upload nhiều file Excel
uploaded_files = st.file_uploader(
    "Chọn nhiều file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Tạo bộ nhớ tạm để giữ nút download không làm sập app
if 'file_ready' not in st.session_state:
    st.session_state['file_ready'] = False
    st.session_state['output_excel'] = None

if uploaded_files:
    # Đưa việc gộp file vào 1 nút bấm để kiểm soát
    if st.button("🚀 Tiến hành gộp file"):
        with st.spinner("Đang xử lý dữ liệu..."):
            dfs = []
            for file in uploaded_files:
                try:
                    # Dùng calamine ở đây để đọc siêu nhanh và không tốn RAM
                    df = pd.read_excel(file, engine="calamine")
                    dfs.append(df)
                except Exception as e:
                    st.error(f"Lỗi khi đọc file {file.name}: {e}")

            if dfs:
                # Gộp tất cả DataFrame (Giữ nguyên toàn bộ các cột như code gốc)
                df_all = pd.concat(dfs, ignore_index=True)
                st.success(f"✅ Gộp thành công {len(dfs)} file, tổng {df_all.shape[0]} dòng")
                
                # Chỉ hiện 100 dòng đầu để trình duyệt không bị đơ
                st.dataframe(df_all.head(100))

                # Xuất file Excel trong bộ nhớ
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_all.to_excel(writer, index=False, sheet_name="Sheet1")
                
                # Lưu file đã tạo vào session_state
                st.session_state['output_excel'] = output.getvalue()
                st.session_state['file_ready'] = True

                # 🔥 DỌN DẸP BỘ NHỚ RAM NGAY LẬP TỨC
                del dfs
                del df_all
                gc.collect()

# Nút download tách biệt hoàn toàn, không làm app chạy lại code bên trên
if st.session_state['file_ready'] and st.session_state['output_excel']:
    st.download_button(
        label="📥 Tải file Excel đã gộp (.xlsx)",
        data=st.session_state['output_excel'],
        file_name="gop_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
