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

# Helper membaca Excel SKPD (Foto 2) dengan deteksi Header Otomatis
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        # Deteksi baris header utama (bukan baris judul laporan)
        if ('5.2' in row_str or '5.3' in row_str or 'BELANJA' in row_str or 'HARGA' in row_str or 'NILAI' in row_str or 'KODE' in row_str) and 'REKAPITULASI' not in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Abaikan baris akumulasi TOTAL / JUMLAH
    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH|TOTAL|SUBTOTAL').any(), axis=1)
    df = df[~mask_total]

    return df

# Helper membaca Excel SIPD (Foto 1) dengan deteksi Header Otomatis
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
        df_rak = pd.read_excel(file_rak) if file_rak.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_rak)
        df_sipd = parse_sipd_data(file_sipd)
        df_skpd = parse_skpd_data(file_skpd)

        # 1. AMBIL KODE REKENING DARI RAK
        col_rek_rak = [c for c in df_rak.columns if 'REKENING' in str(c).upper() or 'KODE' in str(c).upper()]
        valid_rekening_list = df_rak[col_rek_rak[0]].dropna().astype(str).str.strip().tolist() if col_rek_rak else df_rak.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        valid_rekening_set = set(valid_rekening_list)

        # 2. FILTER SKPD TARGET
        col_skpd_sipd = [c for c in df_sipd.columns if 'skpd' in c.lower() or 'dinas' in c.lower() or 'opd' in c.lower()]
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Target Rekonsiliasi")
        
        if col_skpd_sipd:
            list_skpd = sorted(df_sipd[col_skpd_sipd[0]].dropna().unique().tolist())
            selected_skpd_target = st.sidebar.selectbox("Pilih SKPD:", options=list_skpd)
            df_sipd_filtered = df_sipd[df_sipd[col_skpd_sipd[0]] == selected_skpd_target].copy()
        else:
            df_sipd_filtered = df_sipd.copy()
            selected_skpd_target = "Semua SKPD"

        # 3. FILTER TRANSAKSI SIPD: SEMUA KODE 5.2 / 5.3 ATAU 'BELANJA MODAL' ATAU TERDAFTAR DI RAK
        col_text_sipd = [c for c in df_sipd_filtered.columns if 'uraian' in c.lower() or 'ref' in c.lower() or 'rekening' in c.lower()]
        
        def match_sipd_bm(row):
            row_content = ' '.join([str(val) for val in row[col_text_sipd].values if pd.notna(val)])
            row_upper = row_content.upper()
            if '5.2' in row_content or '5.3' in row_content or 'BELANJA MODAL' in row_upper:
                return True
            for rek in valid_rekening_set:
                if rek in row_content:
                    return True
            return False

        if col_text_sipd:
            df_sipd_filtered['Is_Belanja_Modal'] = df_sipd_filtered.apply(match_sipd_bm, axis=1)
            df_sipd_bm = df_sipd_filtered[df_sipd_filtered['Is_Belanja_Modal']].copy()
        else:
            df_sipd_bm = df_sipd_filtered.copy()

        # 4. HITUNG NOMINAL SIPD
        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-3]
        df_sipd_bm['Nominal_Clean'] = df_sipd_bm[sipd_target_col].apply(clean_currency)
        total_sipd_bm = df_sipd_bm['Nominal_Clean'].sum()

        # 5. FILTER DATA ENTRY SKPD: SEMUA KODE 5.2 / 5.3 ATAU 'BELANJA MODAL' ATAU TERDAFTAR DI RAK
        def match_skpd_bm(row):
            row_content = ' '.join([str(val) for val in row.values if pd.notna(val)])
            row_upper = row_content.upper()
            if '5.2' in row_content or '5.3' in row_content or 'BELANJA MODAL' in row_upper:
                return True
            for rek in valid_rekening_set:
                if rek in row_content:
                    return True
            return False

        df_skpd['Is_Belanja_Modal'] = df_skpd.apply(match_skpd_bm, axis=1)
        df_skpd_bm_base = df_skpd[df_skpd['Is_Belanja_Modal']].copy()

        # 6. PEMILIH KOLOM NOMINAL DI SIDEBAR
        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ Pengaturan Kolom SKPD")
        
        col_options = [c for c in df_skpd.columns if c != 'Is_Belanja_Modal']
        
        best_col_idx = 0
        min_diff = float('inf')
        for i, col in enumerate(col_options):
            s = df_skpd_bm_base[col].apply(clean_currency).sum()
            if s > 0:
                diff = abs(s - total_sipd_bm)
                if diff < min_diff:
                    min_diff = diff
                    best_col_idx = i

        selected_skpd_col = st.sidebar.selectbox(
            "Kolom Nominal SKPD (Foto 2):",
            options=col_options,
            index=best_col_idx
        )

        df_skpd['Nominal_Clean'] = df_skpd[selected_skpd_col].apply(clean_currency)
        df_skpd_bm = df_skpd[(df_skpd['Is_Belanja_Modal']) & (df_skpd['Nominal_Clean'] > 0)].copy()

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
            st.caption(f"Kolom nominal aktif: **{selected_skpd_col}** | Mengakomodasi seluruh transaksi ber-kode `5.2...`, `5.3...`, atau berteks 'BELANJA MODAL'.")
            st.dataframe(df_skpd_bm, use_container_width=True)

        with tab3:
            st.subheader("Transaksi Non-Belanja Modal yang Dieliminasi")
            df_sipd_non_bm = df_sipd_filtered[~df_sipd_filtered['Is_Belanja_Modal']].copy() if col_text_sipd else pd.DataFrame()
            st.dataframe(df_sipd_non_bm, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
