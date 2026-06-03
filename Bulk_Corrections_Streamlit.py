import streamlit as st
import pandas as pd
import re
from io import BytesIO

# --- HÀM MỚI: XỬ LÝ CHUỖI SIÊU TỐC (VECTORIZED) ---
def clean_invoice_series(series):
    """ Dùng sức mạnh xử lý mảng của Pandas thay vì chạy vòng lặp từng dòng """
    # Chuyển về chuỗi và xóa khoảng trắng
    s = series.astype(str).str.strip()
    # Xóa đuôi .0
    s = s.str.replace(r'\.0$', '', regex=True)
    # Xóa hậu tố COR/REV (lưu lại base để dùng nếu lstrip ra chuỗi rỗng)
    base = s.str.replace(r'[- ]*(COR|REV)\d*$', '', regex=True, flags=re.IGNORECASE)
    # Xóa số 0 ở đầu
    cleaned = base.str.lstrip('0')
    # Nếu xóa 0 xong bị rỗng, lấy lại base. Đổi 'nan' thành NA thực thụ để drop
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
st.title("Bulk Correction - Eric Hayes")

# --- 1. NHẬP THÔNG TIN CASE ---
st.subheader("1. Required information")
col1, col2 = st.columns(2)
with col1:
    case_number = st.text_input("Case Number", placeholder="VD: 2605-10535218")
with col2:
    impacted_month = st.text_input("Impacted Month", placeholder="VD: April")

# --- 2. UPLOAD TỆP DỮ LIỆU ---
st.subheader("2. Upload section")
correction_file = st.file_uploader("Requested Correction File", type=['xlsx', 'xls', 'xlsb'])
atf_file = st.file_uploader("ATF File", type=['xlsx', 'xls', 'xlsb'])
postal_ref_file = st.file_uploader("Postal Codes Ref File", type=['xlsx', 'xls', 'xlsb'])

# --- 3. XỬ LÝ DỮ LIỆU ---
if st.button("Data Processing", type="primary"):
    if not case_number or not impacted_month:
        st.warning("⚠️ You must enter a case number and the impacted month!")
    elif not correction_file or not atf_file or not postal_ref_file:
        st.warning("⚠️ You must upload 3 required documents!")
    else:
        progress_bar = st.progress(0, text="Data processing...")
        
        try:
            # --- BƯỚC 1: XỬ LÝ CORRECTION FILE (20%) ---
            progress_bar.progress(20, text="Đang nạp toàn bộ sheet vào RAM...")
            
            # TỐI ƯU 1: Đọc sheet_name=None để load 1 lần duy nhất thay vì I/O nhiều lần
            all_sheets = pd.read_excel(correction_file, sheet_name=None)
            corr_original_invs = set()
            rep_mapping = {}
            ignore_sheets = ['Participant Earning Summary MAR']

            for sheet_name, df_sheet in all_sheets.items():
                sales_rep = str(sheet_name).strip()
                if sales_rep.upper() in [x.upper() for x in ignore_sheets]:
                    continue
                    
                if 'Invoice Number' in df_sheet.columns:
                    # TỐI ƯU 2: Dùng Vectorization thay cho vòng lặp apply
                    cleaned_series = clean_invoice_series(df_sheet['Invoice Number']).dropna()
                    
                    # Ánh xạ nhanh vào Dictionary
                    for orig_inv in cleaned_series:
                        corr_original_invs.add(orig_inv)
                        rep_mapping[orig_inv] = sales_rep.upper()

            # --- BƯỚC 2: ĐỌC POSTAL CODES REF FILE (40%) ---
            progress_bar.progress(40, text="Analyzing Postal Codes file...")
            
            df_postal = pd.read_excel(postal_ref_file)
            first_col = df_postal.iloc[:, 0].astype(str).str.strip().str.upper()
            second_col = df_postal.iloc[:, 1].astype(str).str.strip()
            postal_mapping = dict(zip(first_col, second_col))

            # --- BƯỚC 3: XỬ LÝ ATF FILE VÀ LỌC INVOICE (65%) ---
            progress_bar.progress(65, text="Data Analyzing...")
            
            df_atf = pd.read_excel(atf_file)

            if 'Invoice Number' in df_atf.columns:
                # TỐI ƯU 2: Xử lý chuỗi nguyên 1 cột cùng lúc, cực nhanh
                df_atf['Original Invoice'] = clean_invoice_series(df_atf['Invoice Number'])
                
                # TỐI ƯU 3: Lọc dữ liệu TRƯỚC, rồi mới tính toán Priority Rank (chỉ tính trên tập nhỏ)
                matched_atf = df_atf[df_atf['Original Invoice'].isin(corr_original_invs)].copy()
                
                matched_atf['Priority_Rank'] = matched_atf['Invoice Number'].apply(get_invoice_rank)
                latest_invoices_df = matched_atf.sort_values('Priority_Rank', ascending=False) \
                                                .drop_duplicates(subset=['Original Invoice'], keep='first')
                latest_invoices_df = latest_invoices_df.drop(columns=['Priority_Rank'])
                
                df_cor = latest_invoices_df.copy()
                df_rev = latest_invoices_df.copy()
                
                # --- BƯỚC 4: UPDATE DATA COR, REV, UPLOAD (85%) ---
                progress_bar.progress(85, text="Analyzing to create correction file...")
                
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
                
                # Cập nhật Comments dựa trên Input
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

                df_cor = df_cor.drop(columns=['Original Invoice'])
                df_rev = df_rev.drop(columns=['Original Invoice'])

                # TẠO SHEET "UPLOAD"
                df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

                # --- BƯỚC 5: LƯU FILE KẾT QUẢ VÀO BỘ NHỚ (100%) ---
                progress_bar.progress(95, text="Analyzing data...")
                
                output_buffer = BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    df_cor.to_excel(writer, sheet_name='COR', index=False)
                    df_rev.to_excel(writer, sheet_name='REV', index=False)
                    df_upload.to_excel(writer, sheet_name='Upload', index=False)
                
                output_buffer.seek(0)
                
                progress_bar.progress(100, text="Hoàn tất quy trình xử lý!")
                
                st.success("✅ Data processing is compelted. Click Download button below for Correction file.")
                st.download_button(
                    label="📥 Correction File Download.xlsx",
                    data=output_buffer,
                    file_name="Correction File.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            else:
                st.error("❌ Error: Can't find the Invoice Number in the file ATF.")
                progress_bar.empty()

        except Exception as e:
            st.error(f"❌ Data processing error: {e}")
            progress_bar.empty()
