import streamlit as st
import pandas as pd
import numpy as np

# Set Konfigurasi Halaman
st.set_page_config(
    page_title="Sistem Rekonsiliasi Belanja Modal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size:2rem; font-weight:700; color:#1E293B; }
    .sub-header { font-size:1rem; color:#64748B; margin-bottom:20px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏛️ Mesin Rekonsiliasi Belanja Modal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Pencocokan Transaksi Presisi Berdasarkan Acuan RAK Belanja Modal</div>', unsafe_allow_html=True)
st.divider()

# --- SIDEBAR UPLOAD ---
st.sidebar.header("📁 Upload File Data")
file_rak = st.sidebar.file_uploader("1. RAK Rekening Belanja Modal (Acuan)", type=['xlsx', 'xls', 'csv'])
file_sipd = st.sidebar.file_uploader("2. Data Realisasi SIPD (Foto 1)", type=['xlsx', 'xls', 'csv'])
file_skpd = st.sidebar.file_uploader("3. Data Entry SKPD / Rincian Aset (Foto 2)", type=['xlsx', 'xls', 'csv'])

# Helper pembersihan nominal Rupiah
def clean_currency(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if val_str in ['-', '', 'NaN', 'nan', 'None']:
        return 0.0
    val_str = val_str.replace('Rp', '').replace(' ', '')
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str and '.' not in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

# Helper membaca Excel SKPD (Foto 2)
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if 'KODE' in row_str or 'PENGADAAN' in row_str or 'URAIAN' in row_str or 'HARGA' in row_str or 'NILAI' in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Hapus baris total akumulasi
    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH|TOTAL').any(), axis=1)
    df = df[~mask_total]

    return df

# Helper membaca Excel SIPD (Foto 1)
def parse_sipd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if 'DEBIT' in row_str or 'SKPD' in row_str or 'NO. BUKTI' in row_str or 'URAIAN' in row_str:
            header_idx = idx
            break
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- PROSES UTAMA ---
if file_rak and file_sipd and file_skpd:
    try:
        df_rak = pd.read_excel(file_rak) if file_rak.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_rak)
        df_sipd = parse_sipd_data(file_sipd)
        df_skpd = parse_skpd_data(file_skpd)

        # 1. AMBIL KODE REKENING DARI RAK
        col_rek_rak = [c for c in df_rak.columns if 'REKENING' in str(c).upper() or 'KODE' in str(c).upper()]
        valid_rekening_list = df_rak[col_rek_rak[0]].dropna().astype(str).str.strip().tolist() if col_rek_rak else df_rak.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        valid_rekening_set = set(valid_rekening_list)

        # 2. FILTER SKPD TARGET DARI SIDEBAR
        col_skpd_sipd = [c for c in df_sipd.columns if 'skpd' in c.lower()]
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Target Rekonsiliasi")
        
        if col_skpd_sipd:
            list_skpd = sorted(df_sipd[col_skpd_sipd[0]].dropna().unique().tolist())
            default_index = 0
            for idx, name in enumerate(list_skpd):
                if 'BPKPD' in name.upper() or 'KEUANGAN' in name.upper():
                    default_index = idx
                    break
                    
            selected_skpd_target = st.sidebar.selectbox("Pilih SKPD:", options=list_skpd, index=default_index)
            df_sipd_filtered = df_sipd[df_sipd[col_skpd_sipd[0]] == selected_skpd_target].copy()
        else:
            df_sipd_filtered = df_sipd.copy()
            selected_skpd_target = "Semua SKPD"

        # 3. FILTER TRANSAKSI SIPD SESUAI RAK
        col_text_sipd = [c for c in df_sipd_filtered.columns if 'uraian' in c.lower() or 'ref' in c.lower() or 'rekening' in c.lower()]
        
        def match_rak_rekening(row):
            row_content = ' '.join([str(val) for val in row[col_text_sipd].values if pd.notna(val)])
            for rek in valid_rekening_set:
                if rek in row_content:
                    return True
            return False

        if col_text_sipd:
            df_sipd_filtered['Is_Belanja_Modal'] = df_sipd_filtered.apply(match_rak_rekening, axis=1)
            df_sipd_bm = df_sipd_filtered[df_sipd_filtered['Is_Belanja_Modal']].copy()
        else:
            df_sipd_bm = df_sipd_filtered.copy()

        # 4. HITUNG NOMINAL SIPD
        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-3]
        df_sipd_bm['Nominal_Clean'] = df_sipd_bm[sipd_target_col].apply(clean_currency)
        total_sipd_bm = df_sipd_bm['Nominal_Clean'].sum()

        # 5. PILIH KOLOM NOMINAL SKPD MANUAL DENGAN PEMILIH DI LAYAR UTAMA
        st.markdown("### 🛠️ Konfigurasi Kolom Entry SKPD")
        
        # Opsi pilihan default (mencari kolom dengan sum terbesar yang tidak melebihi 10 miliar)
        default_idx = 0
        max_sum = -1
        for i, col in enumerate(df_skpd.columns):
            s = df_skpd[col].apply(clean_currency).sum()
            if 100000 < s < 10000000000 and s > max_sum:
                max_sum = s
                default_idx = i

        selected_skpd_col = st.selectbox(
            "👉 **PILIH KOLOM NOMINAL PENERIMAAN / PENGADAAN (Foto 2):**",
            options=list(df_skpd.columns),
            index=default_idx
        )

        df_skpd['Nominal_Clean'] = df_skpd[selected_skpd_col].apply(clean_currency)
        df_skpd_bm = df_skpd[df_skpd['Nominal_Clean'] > 0].copy()

        total_skpd_bm = df_skpd_bm['Nominal_Clean'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        st.success(f"✅ Rekonsiliasi selesai untuk **{selected_skpd_target}**!")

        # Dashboard Metrik
        col1, col2, col3 = st.columns(3)
        col1.metric("Realisasi SIPD (Belanja Modal)", f"Rp {total_sipd_bm:,.2f}")
        col2.metric("Entry SKPD (Rincian Aset)", f"Rp {total_skpd_bm:,.2f}")
        col3.metric(
            "Selisih Rekonsiliasi Belanja Modal", 
            f"Rp {total_selisih:,.2f}", 
            delta=f"{-total_selisih:,.2f}", 
            delta_color="inverse"
        )

        st.markdown("---")

        # TAMPILAN TAB
        tab1, tab2, tab3 = st.tabs(["🔍 Detail Realisasi SIPD", "📋 Detail Entry SKPD (Rincian Aset)", "📁 Transaksi Dieliminasi"])

        with tab1:
            st.subheader(f"Transaksi Belanja Modal SIPD - {selected_skpd_target}")
            st.dataframe(df_sipd_bm, use_container_width=True)

        with tab2:
            st.subheader(f"Rincian Pengadaan SKPD Terdeteksi ({len(df_skpd_bm)} Item)")
            st.info(f"Kolom acuan saat ini: **{selected_skpd_col}** (Total: Rp {total_skpd_bm:,.2f})")
            st.dataframe(df_skpd_bm, use_container_width=True)

        with tab3:
            st.subheader("Transaksi Non-Belanja Modal yang Dieliminasi")
            df_sipd_non_bm = df_sipd_filtered[~df_sipd_filtered['Is_Belanja_Modal']].copy() if col_text_sipd else pd.DataFrame()
            st.dataframe(df_sipd_non_bm, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
