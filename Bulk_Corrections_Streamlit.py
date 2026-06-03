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
st.title("Bulk Corrections-Eric Hayes")
st.write("Upload required documents below.")

# Upload Files (ĐÃ BỔ SUNG ĐỊNH DẠNG xlsb)
correction_file = st.file_uploader("1. Eric Haye's File", type=['xlsx', 'xls', 'xlsb'])
atf_file = st.file_uploader("2. ATF File", type=['xlsx', 'xls', 'xlsb'])
postal_ref_file = st.file_uploader("3. Postal Codes Ref File", type=['xlsx', 'xls', 'xlsb'])

if st.button("Start to process", type="primary"):
    if not correction_file or not atf_file or not postal_ref_file:
        st.warning("⚠️ Upload all required documents before processing")
    else:
        with st.spinner("⏳ Processing data, please wait for a few seconds..."):
            try:
                # --- BƯỚC 3a: XỬ LÝ CORRECTION FILE ---
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

                # --- BƯỚC ĐỌC POSTAL CODES REF FILE ---
                df_postal = pd.read_excel(postal_ref_file)
                first_col = df_postal.iloc[:, 0].astype(str).str.strip().str.upper()
                second_col = df_postal.iloc[:, 1].astype(str).str.strip()
                postal_mapping = dict(zip(first_col, second_col))

                # --- BƯỚC 3b & 3c: XỬ LÝ ATF FILE VÀ LỌC INVOICE ---
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
                    
                    # --- BƯỚC 5: UPDATE SHEET "COR" ---
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

                    # --- BƯỚC 6: UPDATE SHEET "REV" ---
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
                    latest_invoices_df = latest_invoices_df.drop(columns=['Original Invoice'])
                    df_cor = df_cor.drop(columns=['Original Invoice'])
                    df_rev = df_rev.drop(columns=['Original Invoice'])

                    # --- BƯỚC TẠO DATA CHO SHEET "UPLOAD" ---
                    df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

                    # --- LƯU FILE KẾT QUẢ VÀO BỘ NHỚ (BYTESIO) ---
                    output_buffer = BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        latest_invoices_df.to_excel(writer, sheet_name='Sheet1', index=False)
                        df_cor.to_excel(writer, sheet_name='COR', index=False)
                        df_rev.to_excel(writer, sheet_name='REV', index=False)
                        df_upload.to_excel(writer, sheet_name='Upload', index=False)
                    
                    output_buffer.seek(0)
                    
                    st.success("✅ Data processing is completed! Click the button below to download the upload file")
                    
                    st.download_button(
                        label="📥 Download the Upload file",
                        data=output_buffer,
                        file_name="Upload file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

                else:
                    st.error("❌ Lỗi: Can't find the Invoice Number in the ATF")

            except Exception as e:
                st.error(f"❌ Data processing error: {e}")
