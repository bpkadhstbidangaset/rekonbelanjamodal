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

# Helper normalisasi string kode agar bebas dari spasi tersembunyi/karakter non-breaking
def clean_text_code(val):
    if pd.isna(val):
        return ""
    return str(val).replace('\xa0', ' ').strip()

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

# Helper membaca Acuan RAK (Foto Excel Acuan)
def parse_rak_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if 'KODE REKENING' in row_str or 'KODE KATEGORI' in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

# Helper membaca Excel SKPD (Foto 2)
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if ('5.2' in row_str or '5.3' in row_str or 'BELANJA' in row_str or 'HARGA' in row_str or 'NILAI' in row_str or 'KODE' in row_str) and 'REKAPITULASI' not in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Abaikan baris akumulasi TOTAL / JUMLAH
    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH TOTAL|GRAND TOTAL|SUBTOTAL').any(), axis=1)
    df = df[~mask_total]
    return df

# Helper membaca Excel SIPD (Foto 1)
def parse_sipd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if ('DEBIT' in row_str or 'SKPD' in row_str or 'URAIAN' in row_str) and 'REKAPITULASI' not in row_str:
            header_idx = idx
            break
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- PROSES UTAMA ---
if file_rak and file_sipd and file_skpd:
    try:
        df_rak = parse_rak_data(file_rak)
        df_sipd = parse_sipd_data(file_sipd)
        df_skpd = parse_skpd_data(file_skpd)

        # 1. AMBIL KODE REKENING & KODE KATEGORI DARI ACUAN RAK
        col_kode_rek = [c for c in df_rak.columns if 'KODE' in c and 'REK' in c]
        col_kode_kat = [c for c in df_rak.columns if 'KODE' in c and 'KAT' in c]

        col_rek_target = col_kode_rek[0] if col_kode_rek else df_rak.columns[-1]
        col_kat_target = col_kode_kat[0] if col_kode_kat else df_rak.columns[0]

        list_kode_rekening = df_rak[col_rek_target].dropna().apply(clean_text_code).tolist()
        list_kode_kategori = df_rak[col_kat_target].dropna().apply(clean_text_code).tolist()

        valid_acuan_set = {k for k in (list_kode_rekening + list_kode_kategori) if len(k) > 3}

        # 2. FILTER SKPD TARGET PADA SIPD
        col_skpd_sipd = [c for c in df_sipd.columns if any(k in c.lower() for k in ['skpd', 'dinas', 'opd'])]
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Target Rekonsiliasi")
        
        if col_skpd_sipd:
            list_skpd = sorted(df_sipd[col_skpd_sipd[0]].dropna().unique().tolist())
            selected_skpd_target = st.sidebar.selectbox("Pilih SKPD:", options=list_skpd)
            df_sipd_filtered = df_sipd[df_sipd[col_skpd_sipd[0]] == selected_skpd_target].copy()
        else:
            df_sipd_filtered = df_sipd.copy()
            selected_skpd_target = "Semua SKPD"

        # 3. FILTER SIPD BERDASARKAN ACUAN RAK
        def match_sipd_rak(row):
            row_str = ' '.join([clean_text_code(v) for v in row.values if pd.notna(v)])
            for code in valid_acuan_set:
                if code in row_str:
                    return True
            return False

        df_sipd_filtered['Is_Acuan_RAK'] = df_sipd_filtered.apply(match_sipd_rak, axis=1)
        df_sipd_bm = df_sipd_filtered[df_sipd_filtered['Is_Acuan_RAK']].copy()

        # 4. HITUNG NOMINAL REALISASI SIPD
        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-3]
        df_sipd_bm['Nominal_Clean'] = df_sipd_bm[sipd_target_col].apply(clean_currency)
        total_sipd_bm = df_sipd_bm['Nominal_Clean'].sum()

        # 5. FILTER SKPD BERDASARKAN ACUAN RAK
        def match_skpd_rak(row):
            row_str = ' '.join([clean_text_code(v) for v in row.values if pd.notna(v)])
            for code in valid_acuan_set:
                if code in row_str:
                    return True
            return False

        df_skpd['Is_Acuan_RAK'] = df_skpd.apply(match_skpd_rak, axis=1)
        df_skpd_bm = df_skpd[df_skpd['Is_Acuan_RAK']].copy()

        # 6. HITUNG NOMINAL SKPD
        cols_to_exclude = [df_skpd.columns[0], df_skpd.columns[1], 'Is_Acuan_RAK', 'SEMUA']
        potential_num_cols = [c for c in df_skpd.columns if c not in cols_to_exclude]

        if potential_num_cols:
            df_skpd_bm['Nominal_Clean'] = df_skpd_bm[potential_num_cols].apply(lambda s: s.map(clean_currency)).sum(axis=1)
        else:
            df_skpd_bm['Nominal_Clean'] = df_skpd_bm.iloc[:, -1].apply(clean_currency)

        total_skpd_bm = df_skpd_bm['Nominal_Clean'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        # TAMPILAN DASHBOARD
        st.success(f"✅ Rekonsiliasi selesai untuk **{selected_skpd_target}**!")

        col1, col2, col3 = st.columns(3)
        col1.metric("Realisasi SIPD (Sesuai RAK)", f"Rp {total_sipd_bm:,.2f}")
        col2.metric("Entry SKPD (Sesuai RAK)", f"Rp {total_skpd_bm:,.2f}")
        col3.metric(
            "Selisih Rekonsiliasi", 
            f"Rp {total_selisih:,.2f}", 
            delta=f"{-total_selisih:,.2f}", 
            delta_color="inverse"
        )

        st.markdown("---")

        tab1, tab2, tab3 = st.tabs(["🔍 Detail Realisasi SIPD", "📋 Detail Entry SKPD", "📁 Data Dieliminasi SKPD"])

        with tab1:
            st.subheader(f"Transaksi SIPD Cocok Acuan ({len(df_sipd_bm)} Baris)")
            st.dataframe(df_sipd_bm, use_container_width=True)

        with tab2:
            st.subheader(f"Rincian SKPD Cocok Acuan ({len(df_skpd_bm)} Baris)")
            st.caption("Nilai nominal dihitung dari kolom sumber dana terkait.")
            st.dataframe(df_skpd_bm, use_container_width=True)

        with tab3:
            st.subheader("Baris SKPD di Luar Acuan RAK")
            df_skpd_elim = df_skpd[~df_skpd['Is_Acuan_RAK']].copy()
            st.dataframe(df_skpd_elim, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
