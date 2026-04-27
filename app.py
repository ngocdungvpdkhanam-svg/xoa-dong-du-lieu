import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Công cụ xử lý Excel chuyên nghiệp", layout="wide")

# Hàm chuyển số thành chữ cái Excel (0 -> A, 1 -> B, ...)
def index_to_letter(n):
    return get_column_letter(n + 1)

st.title("📂 Công cụ lọc & xóa dòng Excel nâng cao")
st.write("Hỗ trợ giữ dòng tiêu đề, lọc theo cột chữ cái (A, B, C...) và xử lý hàng loạt.")

# 1. Tải file lên
uploaded_files = st.file_uploader("Bước 1: Chọn các file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        # Đọc file đầu tiên để lấy mẫu cấu trúc
        sample_df = pd.read_excel(uploaded_files[0], header=None)
        num_cols = len(sample_df.columns)
        col_letters = [index_to_letter(i) for i in range(num_cols)]
        
        st.divider()
        # Giao diện cấu hình
        row1_col1, row1_col2 = st.columns(2)
        
        with row1_col1:
            skip_rows = st.number_input("Bước 2: Số dòng tiêu đề muốn GIỮ LẠI (không lọc):", 
                                        min_value=0, max_value=len(sample_df), value=1, step=1)
            st.caption(f"Dòng 1 đến dòng {skip_rows} sẽ luôn được giữ lại.")

        with row1_col2:
            selected_letter = st.selectbox("Bước 3: Chọn cột dựa trên (A, B, C...):", options=col_letters)
            col_index = col_letters.index(selected_letter)

        # Lấy dữ liệu mẫu để chọn giá trị (loại bỏ các dòng tiêu đề khi lấy danh sách giá trị)
        data_only_sample = sample_df.iloc[skip_rows:]
        unique_vals = data_only_sample[col_index].dropna().unique().tolist()
        unique_vals = [str(v) for v in unique_vals]

        st.divider()
        row2_col1, row2_col2 = st.columns(2)

        with row2_col1:
            selected_values = st.multiselect(f"Bước 4: Chọn giá trị trong cột {selected_letter}:", options=unique_vals)

        with row2_col2:
            action = st.radio("Bước 5: Chọn hành động với giá trị đã chọn:", 
                              ("Xóa các dòng chứa giá trị này", 
                               "Giữ lại giá trị này (Xóa tất cả các dòng khác)"))

        if selected_values:
            st.divider()
            if st.button("🚀 Bắt đầu xử lý hàng loạt"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for uploaded_file in uploaded_files:
                        try:
                            # Đọc toàn bộ file
                            df = pd.read_excel(uploaded_file, header=None)
                            
                            # Tách phần tiêu đề và phần dữ liệu
                            header_part = df.head(skip_rows)
                            data_part = df.iloc[skip_rows:]
                            
                            # Thực hiện lọc trên phần dữ liệu
                            column_data = data_part.iloc[:, col_index].astype(str)
                            
                            if "Xóa các dòng" in action:
                                filtered_data = data_part[~column_data.isin(selected_values)]
                                mode_tag = "Xoa"
                            else:
                                filtered_data = data_part[column_data.isin(selected_values)]
                                mode_tag = "GiuLai"
                            
                            # Ghép tiêu đề và dữ liệu đã lọc lại với nhau
                            final_df = pd.concat([header_part, filtered_data], ignore_index=True)
                            
                            rows_removed = len(data_part) - len(filtered_data)
                            
                            # Tạo tên file mới
                            base_name = uploaded_file.name.rsplit('.', 1)[0]
                            info_str = "_".join(selected_values)[:30]
                            new_filename = f"{base_name}_{mode_tag}_{info_str}.xlsx"
                            
                            # Ghi file vào bộ nhớ
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                final_df.to_excel(writer, index=False, header=False)
                            
                            zip_file.writestr(new_filename, output.getvalue())
                            st.info(f"Đã xử lý: {uploaded_file.name} (Đã xử lý {rows_removed} dòng)")
                            
                        except Exception as e:
                            st.error(f"Lỗi file {uploaded_file.name}: {e}")

                st.success("✅ Đã xử lý xong!")
                st.download_button(
                    label="📥 Tải xuống kết quả (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="ket_qua_loc_du_lieu.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"Lỗi: {e}")
else:
    st.info("Vui lòng tải file Excel lên để bắt đầu.")
