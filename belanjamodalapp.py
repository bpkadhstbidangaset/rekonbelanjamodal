import streamlit as st
import pandas as pd
import numpy as np
import re

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

# Helper normalisasi kode (Hapus titik & karakter non-alfanumerik agar format fleksibel)
def normalize_code(val):
    if pd.isna(val):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(val).strip())

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

# Helper membaca Acuan RAK
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

# Helper membaca Excel SKPD
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if any(k in row_str for k in ['5.2', '5.3', 'JENIS BELANJA', 'REKENING', 'SEMUA']) and 'REKAPITULASI' not in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Hapus baris akumulasi TOTAL / JUMLAH
    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH TOTAL|GRAND TOTAL|SUBTOTAL').any(), axis=1)
    df = df[~mask_total]
    return df

# Helper membaca Excel SIPD
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

        # 1. AMBIL KODE DARI FILE RAK ACUAN
        col_kode_rek = [c for c in df_rak.columns if 'KODE' in c and 'REK' in c]
        col_kode_kat = [c for c in df_rak.columns if 'KODE' in c and 'KAT' in c]

        col_rek_target = col_kode_rek[0] if col_kode_rek else df_rak.columns[-1]
        col_kat_target = col_kode_kat[0] if col_kode_kat else df_rak.columns[0]

        list_kode_rekening = df_rak[col_rek_target].dropna().astype(str).tolist()
        list_kode_kategori = df_rak[col_kat_target].dropna().astype(str).tolist()

        # Simpan raw list & normalized set
        raw_acuan_list = [k.strip() for k in (list_kode_rekening + list_kode_kategori) if len(k.strip()) > 3]
        norm_acuan_set = {normalize_code(k) for k in raw_acuan_list if len(normalize_code(k)) >= 5}

        # 2. FILTER PILIHAN SKPD PADA SIPD
        col_skpd_sipd = [c for c in df_sipd.columns if any(k in c.lower() for k in ['skpd', 'dinas', 'opd', 'unit'])]
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Target Rekonsiliasi")
        
        if col_skpd_sipd:
            list_skpd = sorted(df_sipd[col_skpd_sipd[0]].dropna().unique().tolist())
            selected_skpd_target = st.sidebar.selectbox("Pilih SKPD Target:", options=list_skpd)
            df_sipd_filtered = df_sipd[df_sipd[col_skpd_sipd[0]] == selected_skpd_target].copy()
        else:
            df_sipd_filtered = df_sipd.copy()
            selected_skpd_target = "Semua Data SIPD"

        # 3. FILTER SIPD DENGAN ACUAN RAK
        def match_sipd_row(row):
            row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
            row_norm = normalize_code(row_str)
            
            # Cek string asli
            for code in raw_acuan_list:
                if code in row_str:
                    return True
            # Cek string ternormalisasi (mengatasi beda format titik)
            for n_code in norm_acuan_set:
                if n_code in row_norm:
                    return True
            return False

        df_sipd_filtered['Is_Acuan_RAK'] = df_sipd_filtered.apply(match_sipd_row, axis=1)
        df_sipd_bm = df_sipd_filtered[df_sipd_filtered['Is_Acuan_RAK']].copy()

        # 4. HITUNG NOMINAL REALISASI SIPD
        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-3]
        df_sipd_bm['Nominal_Clean'] = df_sipd_bm[sipd_target_col].apply(clean_currency)
        total_sipd_bm = df_sipd_bm['Nominal_Clean'].sum()

        # 5. FILTER SKPD DENGAN ACUAN RAK
        def match_skpd_row(row):
            first_cols_str = ' '.join([str(v) for v in row.iloc[:3].values if pd.notna(v)])
            norm_first = normalize_code(first_cols_str)

            for code in raw_acuan_list:
                if code in first_cols_str:
                    return True
            for n_code in norm_acuan_set:
                if n_code in norm_first:
                    return True
            return False

        df_skpd['Is_Acuan_RAK'] = df_skpd.apply(match_skpd_row, axis=1)
        df_skpd_bm = df_skpd[df_skpd['Is_Acuan_RAK']].copy()

        # 6. HITUNG NOMINAL SKPD (Mencegah Double Counting)
        # Prioritaskan kolom 'SEMUA' atau 'TOTAL' atau 'JUMLAH'
        col_semua = [c for c in df_skpd_bm.columns if c.upper() in ['SEMUA', 'TOTAL', 'JUMLAH', 'NILAI']]
        
        if col_semua:
            df_skpd_bm['Nominal_Clean'] = df_skpd_bm[col_semua[0]].apply(clean_currency)
        else:
            # Jika tidak ada kolom 'SEMUA', ambil kolom angka paling kanan (bukan semua kolom dijumlah)
            df_skpd_bm['Nominal_Clean'] = df_skpd_bm.iloc[:, -2].apply(clean_currency)

        total_skpd_bm = df_skpd_bm['Nominal_Clean'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        # TAMPILAN DASHBOARD
        st.success(f"✅ Rekonsiliasi selesai untuk: **{selected_skpd_target}**")

        # Peringatan jika file SKPD tidak sesuai dengan pilihan dropdown
        if "KESEHATAN" in selected_skpd_target.upper() and ("PUPR" in file_skpd.name.upper() or "PEKERJAAN" in file_skpd.name.upper()):
            st.warning("⚠️ **Perhatian**: Target SKPD yang dipilih adalah Bidang Kesehatan, tetapi file SKPD yang diunggah adalah file PUPR. Pastikan file SKPD sesuai dengan unit yang dipilih.")

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
            st.caption("Nominal diambil langsung dari kolom utama 'SEMUA' / total per baris.")
            st.dataframe(df_skpd_bm, use_container_width=True)

        with tab3:
            st.subheader("Baris SKPD di Luar Acuan RAK")
            df_skpd_elim = df_skpd[~df_skpd['Is_Acuan_RAK']].copy()
            st.dataframe(df_skpd_elim, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
