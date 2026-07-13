import streamlit as st
import pandas as pd
import io
import gc  # Thư viện Garbage Collector để ép hệ thống xóa bộ nhớ rác

st.set_page_config(page_title="Merge Excel Files", layout="wide")
st.title("📊 Merge Multiple Excel Files (Bản Tối Ưu RAM)")

st.info("💡 Mẹo dành cho file lớn: Ứng dụng khuyên dùng xuất ra file dạng `.CSV` để tải về nhanh hơn và tránh làm sập hệ thống.")

# Upload nhiều file Excel
uploaded_files = st.file_uploader(
    "Chọn nhiều file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    dfs = []
    for file in uploaded_files:
        try:
            # Đọc bằng calamine rất nhẹ RAM
            df = pd.read_excel(file, engine="calamine")
            dfs.append(df)
        except Exception as e:
            st.error(f"Lỗi khi đọc file {file.name}: {e}")

    if dfs:
        # Gộp tất cả DataFrame
        df_all = pd.concat(dfs, ignore_index=True)
        st.success(f"✅ Gộp thành công {len(dfs)} file, tổng {df_all.shape[0]} dòng")
        
        # Chỉ hiển thị 50 dòng đầu để tránh đơ trình duyệt
        st.write("Xem trước 50 dòng đầu tiên của file đã gộp:")
        st.dataframe(df_all.head(50))

        # 🔥 BƯỚC QUAN TRỌNG: GIẢI PHÓNG RAM LẬP TỨC
        del dfs       # Xóa danh sách các file lẻ cũ khỏi bộ nhớ
        gc.collect()  # Ép server dọn rác và giải phóng RAM ngay lập tức

        # -------------------------------------------------------------
        # PHƯƠNG ÁN 1 (KHUYÊN DÙNG): Xuất ra file CSV (Tốn ít RAM, không bao giờ treo)
        # -------------------------------------------------------------
        # utf-8-sig giúp file CSV khi mở bằng Excel trên máy tính không bị lỗi font Tiếng Việt/Trung
        csv_data = df_all.to_csv(index=False).encode('utf-8-sig')
        
        st.download_button(
            label="📥 Tải file đã gộp (Định dạng .CSV - Khuyên dùng cho file nặng)",
            data=csv_data,
            file_name="gop_file_hoanthanh.csv",
            mime="text/csv"
        )

        # -------------------------------------------------------------
        # PHƯƠNG ÁN 2 (TÙY CHỌN): Nếu bắt buộc phải lấy file .XLSX (Excel)
        # -------------------------------------------------------------
        with st.expander("Bạn bắt buộc phải lấy định dạng Excel (.xlsx)?"):
            st.warning("⚠️ Lưu ý: Tạo file Excel cho dữ liệu > 100MB rất dễ làm sập server Streamlit Cloud do giới hạn RAM 1GB.")
            if st.button("Bấm vào đây để cố gắng tạo file Excel"):
                with st.spinner("Đang xử lý tạo file Excel... Xin vui lòng đợi."):
                    try:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            df_all.to_excel(writer, index=False, sheet_name="Sheet1")
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Tải file định dạng Excel (.xlsx)",
                            data=output,
                            file_name="gop_file_hoanthanh.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as excel_error:
                        st.error(f"Không thể tạo file Excel do server cạn kiệt RAM: {excel_error}. Hãy dùng nút tải file .CSV ở trên!")
