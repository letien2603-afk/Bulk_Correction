import streamlit as st
import pandas as pd
import re
from io import BytesIO

# --- HÀM MỚI: XỬ LÝ CHUỖI SIÊU TỐC (VECTORIZED) ---
def clean_invoice_series(series):
    """ Dùng sức mạnh xử lý mảng của Pandas thay vì chạy vòng lặp từng dòng """
    s = series.astype(str).str.strip()
    s = s.str.replace(r'\.0$', '', regex=True)
    base = s.str.replace(r'[- ]*(COR|REV)\d*$', '', regex=True, flags=re.IGNORECASE)
    cleaned = base.str.lstrip('0')
    return cleaned.where(cleaned != '', base).replace('nan', pd.NA)

def get_invoice_rank(inv):
    if pd.isna(inv): return -1
    inv_str = str(inv).upper()
    match = re.search(r'[- ]*(COR|REV)(\d*)$', inv_str)
    if not match: return 0 
    type_str, num_str = match.group(1), match.group(2)
    num = int(num_str) if num_str else 1
    type_score = 2 if type_str == 'COR' else 1
    return (num * 10) + type_score

def update_suffix(val, target_suffix):
    if pd.isna(val): return val
    val_str = str(val).strip()
    match = re.search(r'[- ]*(COR|REV)(\d*)$', val_str, flags=re.IGNORECASE)
    
    if match:
        current_num_str = match.group(2)
        base_str = val_str[:match.start()]
        num = int(current_num_str) if current_num_str else 1
        return f"{base_str}-{target_suffix}{num + 1}"
    else:
        return f"{val_str}-{target_suffix}"

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Invoice Processing Tool", layout="centered")
st.title("Chương Trình Xử Lý Invoice (Đa Định Dạng)")

# --- 1. NHẬP THÔNG TIN CASE ---
st.subheader("1. Thông tin Case (Bắt buộc)")
col1, col2 = st.columns(2)
with col1:
    case_number = st.text_input("Nhập Case Number", placeholder="VD: 2605-10535218")
with col2:
    impacted_month = st.text_input("Nhập Impacted Month", placeholder="VD: April")

# --- 2. UPLOAD TỆP DỮ LIỆU ---
st.subheader("2. Tải lên dữ liệu")
correction_file = st.file_uploader("Tải lên 'Requested Correction File' (Excel)", type=['xlsx', 'xls', 'xlsb'])
atf_file = st.file_uploader("Tải lên 'ATF File' (Excel)", type=['xlsx', 'xls', 'xlsb'])
postal_ref_file = st.file_uploader("Tải lên 'Postal Codes Ref File' (Excel)", type=['xlsx', 'xls', 'xlsb'])

# --- 3. XỬ LÝ DỮ LIỆU ---
if st.button("Bắt Đầu Xử Lý", type="primary"):
    if not case_number or not impacted_month:
        st.warning("⚠️ Vui lòng nhập đầy đủ 'Case Number' và 'Impacted Month'!")
    elif not correction_file or not atf_file or not postal_ref_file:
        st.warning("⚠️ Vui lòng tải lên đầy đủ cả 3 tệp trước khi xử lý!")
    else:
        progress_bar = st.progress(0, text="Khởi động quy trình...")
        
        try:
            # --- BƯỚC 1: XỬ LÝ CORRECTION FILE (20%) ---
            progress_bar.progress(20, text="Đang nạp toàn bộ sheet vào RAM...")
            
            all_sheets = pd.read_excel(correction_file, sheet_name=None)
            corr_original_invs = set()
            rep_mapping = {}
            ignore_sheets = ['Participant Earning Summary MAR']

            for sheet_name, df_sheet in all_sheets.items():
                sales_rep = str(sheet_name).strip()
                if sales_rep.upper() in [x.upper() for x in ignore_sheets]:
                    continue
                    
                if 'Invoice Number' in df_sheet.columns:
                    cleaned_series = clean_invoice_series(df_sheet['Invoice Number']).dropna()
                    
                    for orig_inv in cleaned_series:
                        corr_original_invs.add(orig_inv)
                        rep_mapping[orig_inv] = sales_rep.upper()

            # --- BƯỚC 2: ĐỌC POSTAL CODES REF FILE (40%) ---
            progress_bar.progress(40, text="Đang ánh xạ mã bưu điện...")
            
            df_postal = pd.read_excel(postal_ref_file)
            first_col = df_postal.iloc[:, 0].astype(str).str.strip().str.upper()
            second_col = df_postal.iloc[:, 1].astype(str).str.strip()
            postal_mapping = dict(zip(first_col, second_col))

            # --- BƯỚC 3: XỬ LÝ ATF FILE VÀ LỌC INVOICE (65%) ---
            progress_bar.progress(65, text="Đang quét dữ liệu ATF...")
            
            df_atf = pd.read_excel(atf_file)

            if 'Invoice Number' in df_atf.columns:
                df_atf['Original Invoice'] = clean_invoice_series(df_atf['Invoice Number'])
                
                matched_atf = df_atf[df_atf['Original Invoice'].isin(corr_original_invs)].copy()
                
                matched_atf['Priority_Rank'] = matched_atf['Invoice Number'].apply(get_invoice_rank)
                latest_invoices_df = matched_atf.sort_values('Priority_Rank', ascending=False) \
                                                .drop_duplicates(subset=['Original Invoice'], keep='first')
                latest_invoices_df = latest_invoices_df.drop(columns=['Priority_Rank'])
                
                df_cor = latest_invoices_df.copy()
                df_rev = latest_invoices_df.copy()
                
                # --- BƯỚC 4: UPDATE DATA COR, REV, UPLOAD (85%) ---
                progress_bar.progress(85, text="Đang tạo dữ liệu hợp nhất (Upload)...")
                
                # UPDATE COR
                if 'Transaction Number' in df_cor.columns:
                    df_cor['Transaction Number'] = df_cor['Transaction Number'].apply(lambda x: update_suffix(x, 'COR')) 
                if 'Transaction Type' in df_cor.columns:
                    df_cor['Transaction Type'] = 'MANUAL_ADJ' 
                if 'Invoice Number' in df_cor.columns:
                    df_cor['Invoice Number'] = df_cor['Invoice Number'].apply(lambda x: update_suffix(x, 'COR')) 
                if 'Other Postal Code' in df_cor.columns:
                    mapped_reps = df_cor['Original Invoice'].map(rep_mapping)
                    mapped_postals = mapped_reps.dropna().map(postal_mapping)
                    df_cor['Other Postal Code'] = mapped_postals.combine_first(df_cor['Other Postal Code'])
                
                # Cập nhật Comments
                comment_value = f"{case_number.strip()} Eric Hayes bulk ({impacted_month.strip()} Impact)"
                df_cor['Comments'] = comment_value

                # UPDATE REV
                if 'Transaction Number' in df_rev.columns:
                    df_rev['Transaction Number'] = df_rev['Transaction Number'].apply(lambda x: update_suffix(x, 'REV'))
                if 'Transaction Type' in df_rev.columns:
                    df_rev['Transaction Type'] = 'MANUAL_ADJ'
                if 'Invoice Number' in df_rev.columns:
                    df_rev['Invoice Number'] = df_rev['Invoice Number'].apply(lambda x: update_suffix(x, 'REV'))
                    
                cols_to_invert = ['Transaction Amount', 'EUR Value', 'CAD Value', 'GBP Value', 'Native Currency', 'AUD Value']
                for col in cols_to_invert:
                    if col in df_rev.columns:
                        df_rev[col] = pd.to_numeric(df_rev[col], errors='coerce') * -1

                df_cor = df_cor.drop(columns=['Original Invoice'])
                df_rev = df_rev.drop(columns=['Original Invoice'])

                # TẠO DATAFRAME TỔNG HỢP (UPLOAD)
                df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

                # --- BƯỚC 5: TẠO FILE EXCEL VÀ CSV (100%) ---
                progress_bar.progress(95, text="Đang đóng gói file Excel và CSV...")
                
                # 1. Khởi tạo file Excel (CHỈ LƯU 1 SHEET UPLOAD)
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_upload.to_excel(writer, sheet_name='Upload', index=False)
                excel_buffer.seek(0)
                
                # 2. Khởi tạo file CSV
                csv_data = df_upload.to_csv(index=False).encode('utf-8')
                
                progress_bar.progress(100, text="Hoàn tất quy trình xử lý!")
                
                st.success("✅ Xử lý thành công! Nhấn các nút bên dưới để tải file kết quả.")
                st.info(f"📊 Thống kê nhanh: Đã trích xuất {len(latest_invoices_df)} hóa đơn hợp lệ (Tổng cộng {len(df_upload)} dòng).")
                
                # --- HIỂN THỊ 2 NÚT DOWNLOAD CẠNH NHAU ---
                col_btn1, col_btn2 = st.columns(2)
                
                with col_btn1:
                    st.download_button(
                        label="📥 Tải xuống Excel (.xlsx)",
                        data=excel_buffer,
                        file_name="Matched_Latest_Invoices_Result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                with col_btn2:
                    st.download_button(
                        label="📥 Tải xuống CSV (.csv)",
                        data=csv_data,
                        file_name="Matched_Latest_Invoices_Result.csv",
                        mime="text/csv",
                        type="primary"
                    )

            else:
                st.error("❌ Lỗi: Không tìm thấy cột 'Invoice Number' trong file ATF.")
                progress_bar.empty()

        except Exception as e:
            st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý: {e}")
            progress_bar.empty()
