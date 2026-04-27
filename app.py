import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Công cụ lọc Excel thông minh", layout="wide")

st.title("📂 Công cụ xóa dòng Excel hàng loạt")
st.write("Hướng dẫn: Tải file -> Chọn cột -> Chọn các giá trị muốn xóa -> Tải về.")

# 1. Tải file lên
uploaded_files = st.file_uploader("Bước 1: Chọn các file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # Lấy danh sách cột từ file đầu tiên để người dùng chọn
    # Dùng file đầu tiên làm mẫu (Template)
    try:
        sample_df = pd.read_excel(uploaded_files[0])
        columns = sample_df.columns.tolist()
        
        st.divider()
        col1, col2 = st.columns(2)
        
        with col1:
            # Bước 2: Chọn cột
            selected_column = st.selectbox("Bước 2: Chọn cột muốn lọc dữ liệu:", options=columns)

        if selected_column:
            # Lấy các giá trị duy nhất trong cột đó từ file đầu tiên (hoặc gộp tất cả file nếu cần)
            # Ở đây lấy từ file đầu tiên để nhanh, hoặc gộp toàn bộ để đầy đủ:
            all_unique_values = sample_df[selected_column].dropna().unique().tolist()
            
            with col2:
                # Bước 3: Chọn các giá trị muốn xóa (Cho phép chọn nhiều)
                values_to_delete = st.multiselect(
                    f"Bước 3: Chọn các giá trị trong cột '{selected_column}' để XOÁ:",
                    options=[str(v) for v in all_unique_values]
                )

        if values_to_delete:
            st.divider()
            if st.button("🚀 Bắt đầu xử lý tất cả file"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for uploaded_file in uploaded_files:
                        try:
                            # Đọc từng file
                            df = pd.read_excel(uploaded_file)
                            
                            if selected_column in df.columns:
                                # Logic: Giữ lại những dòng KHÔNG nằm trong danh sách chọn (nghĩa là xóa dòng đã chọn)
                                initial_rows = len(df)
                                # Chuyển cột về string để so sánh khớp với multiselect
                                df = df[~df[selected_column].astype(str).isin(values_to_delete)]
                                rows_removed = initial_rows - len(df)
                                
                                # Tạo tên file mới: TenCu_GiaTri1_GiaTri2...
                                base_name = uploaded_file.name.rsplit('.', 1)[0]
                                suffix = "_".join(values_to_delete)[:50] # Giới hạn độ dài tên file
                                new_filename = f"{base_name}_{suffix}.xlsx"
                                
                                # Ghi vào bộ nhớ đệm
                                output = io.BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    df.to_excel(writer, index=False)
                                
                                # Thêm vào ZIP
                                zip_file.writestr(new_filename, output.getvalue())
                                st.info(f"Đã xử lý: {uploaded_file.name} (Xóa {rows_removed} dòng)")
                        
                        except Exception as e:
                            st.error(f"Lỗi khi xử lý file {uploaded_file.name}: {e}")

                # Nút tải xuống file ZIP
                st.success("✅ Hoàn tất!")
                st.download_button(
                    label="📥 Tải xuống tất cả file đã xử lý (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Ket_qua_xoa_du_lieu.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"Không thể đọc file. Vui lòng kiểm tra định dạng Excel. Lỗi: {e}")

else:
    st.info("Vui lòng tải ít nhất một file Excel để bắt đầu.")
