import streamlit as st
import pandas as pd
import re
from io import BytesIO

# --- HÀM XỬ LÝ CHUỖI SIÊU TỐC (VECTORIZED) ---
def clean_invoice_series(series): 
    s = series.astype(str).str.strip() 
    s = s.str.replace(r'\.0$', '', regex=True) 
    # [CẬP NHẬT]: Dùng [\s\-]* để quét sạch mọi dấu gạch ngang (bao gồm --COR)
    base = s.str.replace(r'[\s\-]*(COR|REV)\d*$', '', regex=True, flags=re.IGNORECASE) 
    # Xoá số 0 ở đầu (Giải quyết triệt để lỗi mất số 0 leading zero giữa 2 file)
    cleaned = base.str.lstrip('0') 
    return cleaned.where(cleaned != '', base).replace('nan', pd.NA)

def get_invoice_rank(inv): 
    if pd.isna(inv): 
        return (0, 0, 0) 
    inv_str = str(inv).upper().strip() 
    # [CẬP NHẬT]: Tách phần dấu gạch ngang ra để đếm số lượng
    match = re.search(r'([\s\-]*)(COR|REV)(\d*)$', inv_str) 
    if not match: 
        return (0, 0, 0) 
    
    separator = match.group(1)
    type_str = match.group(2)
    num_str = match.group(3)
    
    type_val = 2 if type_str == 'COR' else 1
    num_val = int(num_str) if num_str else 1
    dash_count = separator.count('-') 
    
    # Trả về bộ 3 giá trị (Số đuôi, Loại, Số lượng gạch ngang)
    return (num_val, type_val, dash_count)

def update_suffix(val, target_suffix): 
    if pd.isna(val): 
        return val 
    val_str = str(val).strip() 
    match = re.search(r'(?i)(.*?)(?:[\s\-]*)(COR|REV)(\d*)$', val_str) 
    if match:
        prefix = match.group(1).rstrip('- ')
        num_str = match.group(3)
        current_num = int(num_str) if num_str else 1
        next_num = current_num + 1
        return f"{prefix}-{target_suffix}{next_num}"
    else:
        return f"{val_str.rstrip('- ')}-{target_suffix}"

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Invoice Processing Tool", layout="centered")
st.title("Bulk Corrections")

# --- KHỞI TẠO BỘ NHỚ TRẠNG THÁI (SESSION STATE) ---
if 'processing_done' not in st.session_state:
    st.session_state.processing_done = False
    st.session_state.excel_data = None
    st.session_state.csv_data = None
    st.session_state.success_msg = ""

# --- 1. NHẬP THÔNG TIN CASE ---
st.subheader("1. Required Information")
col1, col2 = st.columns(2)
with col1:
    case_number = st.text_input("Case Number", placeholder="VD: 2605-10535218")
with col2:
    impacted_month = st.text_input("Impacted Month", placeholder="VD: April")

# --- 2. UPLOAD TỆP DỮ LIỆU ---
st.subheader("2. Upload required documents")
correction_file = st.file_uploader("Requested Correction File", type=['xlsx', 'xls', 'xlsb'])
atf_file = st.file_uploader("ATF File", type=['xlsx', 'xls', 'xlsb'])
postal_ref_file = st.file_uploader("Postal Codes Ref File", type=['xlsx', 'xls', 'xlsb'])

# --- 3. XỬ LÝ DỮ LIỆU ---
if st.button("Start Processing", type="primary"):
    if not case_number or not impacted_month:
        st.warning("⚠️ You must enter a Case Number and Impacted Month!")
    elif not correction_file or not atf_file or not postal_ref_file:
        st.warning("⚠️ You must upload 3 required documents!")
    else:
        progress_bar = st.progress(0, text="Start Data Processing...")
        
        try:
            # Bước 1
            progress_bar.progress(20, text="Uploading sheets to RAM...")
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

            # Bước 2
            progress_bar.progress(40, text="Reading Postal Codes Ref...")
            df_postal = pd.read_excel(postal_ref_file)
            first_col = df_postal.iloc[:, 0].astype(str).str.strip().str.upper()
            second_col = df_postal.iloc[:, 1].astype(str).str.strip()
            postal_mapping = dict(zip(first_col, second_col))

#-------------------------------
            # Bước 3
            progress_bar.progress(65, text="Scanning ATF file...")
            df_atf = pd.read_excel(atf_file)

            if 'Invoice Number' in df_atf.columns:
                df_atf['Original Invoice'] = clean_invoice_series(df_atf['Invoice Number'])
                matched_atf = df_atf[df_atf['Original Invoice'].isin(corr_original_invs)].copy()
                
                matched_atf['Priority_Rank'] = matched_atf['Invoice Number'].apply(get_invoice_rank)
                
                # [QUAN TRỌNG]: Bỏ hàm drop_duplicates và dùng Temp_Amount để gom nhóm trị tuyệt đối
                # Việc này giúp giữ LẠI ĐÚNG VÀ ĐỦ 13 dòng màu vàng cho invoice 0043905960
                if 'Transaction Amount' in matched_atf.columns:
                    matched_atf['Temp_Amount'] = pd.to_numeric(matched_atf['Transaction Amount'], errors='coerce').abs()
                    max_ranks = matched_atf.groupby(['Original Invoice', 'Temp_Amount'], dropna=False)['Priority_Rank'].transform('max')
                    matched_atf.drop(columns=['Temp_Amount'], inplace=True)
                else:
                    max_ranks = matched_atf.groupby('Original Invoice')['Priority_Rank'].transform('max')
                    
                # Lọc ra tất cả các dòng đạt thứ hạng cao nhất (thay vì chỉ lấy dòng đầu tiên)
                latest_invoices_df = matched_atf[matched_atf['Priority_Rank'] == max_ranks].copy()
                latest_invoices_df = latest_invoices_df.drop(columns=['Priority_Rank'])
                
                df_cor = latest_invoices_df.copy()
                df_rev = latest_invoices_df.copy()

#----------------------
                
                # Bước 4
                progress_bar.progress(85, text="Creating the correction file...")
                
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
                
                df_cor['Comments'] = f"{case_number.strip()} Eric Hayes bulk ({impacted_month.strip()} Impact)"
                
                if 'Comments' in df_cor.columns:
                    df_rev['Comments'] = df_cor['Comments']
                elif 'Comment' in df_cor.columns:
                    df_rev['Comment'] = df_cor['Comment']

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
                df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

                # Bước 5: LƯU VÀO BỘ NHỚ STATE
                progress_bar.progress(95, text="Preparing the correction file...")
                
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_upload.to_excel(writer, sheet_name='Upload', index=False)
                
                # Lưu file dưới dạng bộ nhớ đệm (Bytes) vào Session State
                st.session_state.excel_data = excel_buffer.getvalue()
                st.session_state.csv_data = df_upload.to_csv(index=False).encode('utf-8')
                st.session_state.success_msg = f"✅ Data Processing is completed."
                st.session_state.processing_done = True
                
                progress_bar.progress(100, text="Data Processing is completed!")

            else:
                st.error("❌ Error: Can't find the Invoice Number column in file ATF.")
                progress_bar.empty()

        except Exception as e:
            st.error(f"❌ Data Processing Error: {e}")
            progress_bar.empty()

# --- 4. HIỂN THỊ NÚT TẢI XUỐNG DỰA VÀO SESSION STATE ---
if st.session_state.processing_done:
    st.success(st.session_state.success_msg)
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        st.download_button(
            label="📥 Download Correction.xlsx",
            data=st.session_state.excel_data,
            file_name="Correction.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
    with col_btn2:
        st.download_button(
            label="📥 Download Correction.CSV",
            data=st.session_state.csv_data,
            file_name="Correction.csv",
            mime="text/csv",
            type="primary"
        )
