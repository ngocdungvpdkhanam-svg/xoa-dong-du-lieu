import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Công cụ lọc & xóa dòng Excel", layout="centered")

st.title("📂 Công cụ xử lý Excel hàng loạt")
st.write("Tải lên các file Excel, nhập cột và giá trị cần xóa. Kết quả sẽ được nén vào file ZIP.")

# 1. Nhập cấu hình
with st.sidebar:
    st.header("Cấu hình lọc")
    column_name = st.text_input("Tên cột cần kiểm tra:", placeholder="Ví dụ: TrangThai")
    filter_value = st.text_input("Giá trị cần xóa:", placeholder="Ví dụ: Loi")
    
# 2. Upload file
uploaded_files = st.file_uploader("Chọn các file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and column_name and filter_value:
    if st.button("Bắt đầu xử lý"):
        # Tạo một bộ nhớ đệm để chứa file ZIP
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            success_count = 0
            
            for uploaded_file in uploaded_files:
                try:
                    # Đọc file
                    df = pd.read_excel(uploaded_file)
                    
                    if column_name in df.columns:
                        # Thực hiện xóa các dòng khớp với giá trị lọc
                        # Chuyển cả cột và giá trị về string để so sánh khớp tuyệt đối
                        initial_rows = len(df)
                        df = df[df[column_name].astype(str) != str(filter_value)]
                        rows_removed = initial_rows - len(df)
                        
                        # Tạo tên file mới
                        base_name = uploaded_file.name.rsplit('.', 1)[0]
                        new_filename = f"{base_name}_{filter_value}.xlsx"
                        
                        # Lưu file vào bộ nhớ đệm Excel
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)
                        
                        # Thêm vào file ZIP
                        zip_file.writestr(new_filename, output.getvalue())
                        st.success(f"Đã xử lý: {uploaded_file.name} (Xóa {rows_removed} dòng)")
                        success_count += 1
                    else:
                        st.warning(f"File {uploaded_file.name} không có cột '{column_name}'")
                        
                except Exception as e:
                    st.error(f"Lỗi xử lý file {uploaded_file.name}: {e}")

        # Cho phép tải xuống file ZIP nếu có file thành công
        if success_count > 0:
            st.divider()
            st.write(f"🎉 Đã xử lý xong {success_count} file!")
            st.download_button(
                label="📥 Tải xuống tất cả kết quả (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"Ket_qua_loc_{filter_value}.zip",
                mime="application/zip"
            )
else:
    st.info("Vui lòng nhập đầy đủ tên cột, giá trị lọc và tải file lên.")
