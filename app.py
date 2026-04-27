import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Công cụ xử lý Excel theo vị trí cột", layout="wide")

# Hàm chuyển số thành chữ cái Excel (0 -> A, 1 -> B, ...)
def index_to_letter(n):
    return get_column_letter(n + 1)

st.title("📂 Công cụ lọc Excel theo Cột (A, B, C...)")
st.write("Công cụ này giúp bạn xóa hoặc giữ dòng dựa trên vị trí cột mặc định của Excel.")

# 1. Tải file lên
uploaded_files = st.file_uploader("Bước 1: Chọn các file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        # Đọc file đầu tiên để lấy cấu trúc cột
        # Dùng header=None để đọc toàn bộ dữ liệu, bao gồm cả hàng 1
        sample_df = pd.read_excel(uploaded_files[0], header=None)
        num_cols = len(sample_df.columns)
        
        # Tạo danh sách chữ cái A, B, C... dựa trên số cột của file
        col_letters = [index_to_letter(i) for i in range(num_cols)]
        
        st.divider()
        c1, c2, c3 = st.columns([1, 2, 2])
        
        with c1:
            # Bước 2: Chọn cột theo chữ cái
            selected_letter = st.selectbox("Bước 2: Chọn cột (A, B, C...):", options=col_letters)
            col_index = col_letters.index(selected_letter)

        # Lấy các giá trị duy nhất từ cột đã chọn (trong file đầu tiên) để gợi ý
        unique_vals = sample_df.iloc[:, col_index].dropna().unique().tolist()
        unique_vals = [str(v) for v in unique_vals]

        with c2:
            # Bước 3: Chọn giá trị
            selected_values = st.multiselect(f"Bước 3: Chọn các giá trị tại cột {selected_letter}:", options=unique_vals)

        with c3:
            # Bước 4: Chọn hành động
            action = st.radio("Bước 4: Chọn hành động:", 
                              ("Xóa các dòng chứa giá trị đã chọn", 
                               "Giữ lại giá trị đã chọn (Xóa các giá trị khác)"))

        if selected_values:
            st.divider()
            if st.button("🚀 Bắt đầu xử lý hàng loạt"):
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for uploaded_file in uploaded_files:
                        try:
                            # Đọc file (không dùng header để giữ nguyên hàng 1)
                            df = pd.read_excel(uploaded_file, header=None)
                            
                            initial_rows = len(df)
                            # Chuyển dữ liệu cột chọn sang string để so sánh
                            column_data = df.iloc[:, col_index].astype(str)
                            
                            if "Xóa các dòng" in action:
                                # Lọc bỏ (xóa) những dòng có giá trị trong danh sách chọn
                                df_result = df[~column_data.isin(selected_values)]
                                mode_name = "Xoa"
                            else:
                                # Giữ lại những dòng có giá trị trong danh sách chọn
                                df_result = df[column_data.isin(selected_values)]
                                mode_name = "GiuLai"

                            rows_removed = initial_rows - len(df_result)
                            
                            # Tạo tên file mới
                            base_name = uploaded_file.name.rsplit('.', 1)[0]
                            info = "_".join(selected_values)[:30]
                            new_filename = f"{base_name}_{mode_name}_{info}.xlsx"
                            
                            # Lưu vào bộ nhớ
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df_result.to_excel(writer, index=False, header=False)
                            
                            zip_file.writestr(new_filename, output.getvalue())
                            st.info(f"Đã xử lý: {uploaded_file.name} (Đã xử lý {rows_removed} dòng)")
                            
                        except Exception as e:
                            st.error(f"Lỗi file {uploaded_file.name}: {e}")

                st.success("✅ Đã xử lý xong tất cả các file!")
                st.download_button(
                    label="📥 Tải xuống kết quả (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="ket_qua_xu_ly.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
else:
    st.info("Vui lòng tải file lên để bắt đầu.")
