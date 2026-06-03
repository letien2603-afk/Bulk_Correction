import streamlit as st
import pandas as pd
import re
from io import BytesIO

# --- CÁC HÀM HỖ TRỢ XỬ LÝ CHUỖI ---
def clean_invoice(inv):
    """ Xóa hậu tố và đồng bộ số 0 ở đầu """
    if pd.isna(inv): return inv
    inv_str = str(inv).strip()
    if inv_str.endswith('.0'): 
        inv_str = inv_str[:-2]
        
    base = re.sub(r'[- ]*(COR|REV)\d*$', '', inv_str, flags=re.IGNORECASE)
    cleaned = base.lstrip('0')
    return cleaned if cleaned else base

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
st.title("Chương Trình Xử Lý Invoice (COR/REV)")

# --- 1. NHẬP THÔNG TIN CASE ---
st.subheader("1. Thông tin Case (Bắt buộc)")
col1, col2 = st.columns(2)
with col1:
    case_number = st.text_input("Nhập Case Number", placeholder="VD: 2605-10535218")
with col2:
    impacted_month = st.text_input("Nhập Impacted Month", placeholder="VD: April")

# --- 2. UPLOAD TỆP DỮ LIỆU ---
st.subheader("2. Tải lên dữ liệu")
st.write("Vui lòng tải lên 3 tệp dữ liệu cần thiết bên dưới để hệ thống xử lý.")

correction_file = st.file_uploader("Tải lên 'Requested Correction File' (Excel)", type=['xlsx', 'xls', 'xlsb'])
atf_file = st.file_uploader("Tải lên 'ATF File' (Excel)", type=['xlsx', 'xls', 'xlsb'])
postal_ref_file = st.file_uploader("Tải lên 'Postal Codes Ref File' (Excel)", type=['xlsx', 'xls', 'xlsb'])

# --- 3. XỬ LÝ DỮ LIỆU ---
if st.button("Bắt Đầu Xử Lý", type="primary"):
    # Kiểm tra người dùng đã nhập đủ Case Number và Impacted Month chưa
    if not case_number or not impacted_month:
        st.warning("⚠️ Vui lòng nhập đầy đủ 'Case Number' và 'Impacted Month' trước khi xử lý!")
    elif not correction_file or not atf_file or not postal_ref_file:
        st.warning("⚠️ Vui lòng tải lên đầy đủ cả 3 tệp trước khi xử lý!")
    else:
        # Khởi tạo thanh tiến trình
        progress_text = "Khởi động quy trình..."
        progress_bar = st.progress(0, text=progress_text)
        
        try:
            # --- BƯỚC 1: XỬ LÝ CORRECTION FILE (20%) ---
            progress_bar.progress(20, text="Đang đọc và phân tích 'Requested Correction File'...")
            
            corr_xls = pd.ExcelFile(correction_file)
            corr_original_invs = set()
            rep_mapping = {}
            ignore_sheets = ['Participant Earning Summary MAR']

            for sheet_name in corr_xls.sheet_names:
                sales_rep = str(sheet_name).strip()
                if sales_rep.upper() in [x.upper() for x in ignore_sheets]:
                    continue
                    
                df_sheet = pd.read_excel(corr_xls, sheet_name=sheet_name)
                if 'Invoice Number' in df_sheet.columns:
                    for _, row in df_sheet.iterrows():
                        orig_inv = clean_invoice(row['Invoice Number'])
                        if pd.notna(orig_inv) and str(orig_inv).strip() != 'nan':
                            corr_original_invs.add(orig_inv)
                            rep_mapping[orig_inv] = sales_rep.upper()

            # --- BƯỚC 2: ĐỌC POSTAL CODES REF FILE (40%) ---
            progress_bar.progress(40, text="Đang ánh xạ mã bưu điện từ 'Postal Codes Ref File'...")
            
            df_postal = pd.read_excel(postal_ref_file)
            first_col = df_postal.iloc[:, 0].astype(str).str.strip().str.upper()
            second_col = df_postal.iloc[:, 1].astype(str).str.strip()
            postal_mapping = dict(zip(first_col, second_col))

            # --- BƯỚC 3: XỬ LÝ ATF FILE VÀ LỌC INVOICE (65%) ---
            progress_bar.progress(65, text="Đang xử lý 'ATF File' và trích xuất hóa đơn mới nhất...")
            
            df_atf = pd.read_excel(atf_file)

            if 'Invoice Number' in df_atf.columns:
                df_atf['Original Invoice'] = df_atf['Invoice Number'].apply(clean_invoice)
                matched_atf = df_atf[df_atf['Original Invoice'].isin(corr_original_invs)].copy()
                
                matched_atf['Priority_Rank'] = matched_atf['Invoice Number'].apply(get_invoice_rank)
                latest_invoices_df = matched_atf.sort_values('Priority_Rank', ascending=False) \
                                                .drop_duplicates(subset=['Original Invoice'], keep='first')
                latest_invoices_df = latest_invoices_df.drop(columns=['Priority_Rank'])
                
                df_cor = latest_invoices_df.copy()
                df_rev = latest_invoices_df.copy()
                
                # --- BƯỚC 4: UPDATE DATA COR, REV, UPLOAD (85%) ---
                progress_bar.progress(85, text="Đang tính toán các chỉ số và tạo dữ liệu cho COR, REV, Upload...")
                
                # UPDATE SHEET "COR"
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
                
                # CẬP NHẬT CỘT COMMENTS TRONG SHEET COR VỚI INPUT CỦA USER
                comment_value = f"{case_number.strip()} Eric Hayes bulk ({impacted_month.strip()} Impact)"
                df_cor['Comments'] = comment_value

                # UPDATE SHEET "REV"
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

                # Dọn dẹp cột phụ
                df_cor = df_cor.drop(columns=['Original Invoice'])
                df_rev = df_rev.drop(columns=['Original Invoice'])

                # TẠO SHEET "UPLOAD"
                df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

                # --- BƯỚC 5: LƯU FILE KẾT QUẢ VÀO BỘ NHỚ (100%) ---
                progress_bar.progress(95, text="Đang xuất file Excel...")
                
                output_buffer = BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    df_cor.to_excel(writer, sheet_name='COR', index=False)
                    df_rev.to_excel(writer, sheet_name='REV', index=False)
                    df_upload.to_excel(writer, sheet_name='Upload', index=False)
                
                output_buffer.seek(0)
                
                # Hoàn tất tiến trình
                progress_bar.progress(100, text="Hoàn tất quy trình xử lý!")
                
                st.success("✅ Xử lý thành công! Nhấn nút bên dưới để tải file kết quả.")
                st.info(f"📊 Thống kê nhanh: Đã xử lý {len(latest_invoices_df)} hóa đơn hợp lệ.")
                
                st.download_button(
                    label="📥 Tải xuống Matched_Latest_Invoices_Result.xlsx",
                    data=output_buffer,
                    file_name="Matched_Latest_Invoices_Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            else:
                st.error("❌ Lỗi: Không tìm thấy cột 'Invoice Number' trong file ATF.")
                progress_bar.empty()

        except Exception as e:
            st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý: {e}")
            progress_bar.empty()
