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
        if ('5.2' in row_str or '5.3' in row_str or 'BELANJA' in row_str or 'HARGA' in row_str or 'NILAI' in row_str or 'KODE' in row_str) and 'REKAPITULASI' not in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Hanya buang baris TOTAL/JUMLAH gabungan di bagian paling bawah
    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH TOTAL|GRAND TOTAL').any(), axis=1)
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

        # 1. FILTER SKPD TARGET DARI SIDEBAR
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

        # 2. AMBIL SELURUH DATA SIPD (TANPA MEMBUANG BARIS)
        df_sipd_bm = df_sipd_filtered.copy()
        df_sipd_bm['Is_Belanja_Modal'] = True

        # 3. HITUNG NOMINAL SIPD
        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-3]
        df_sipd_bm['Nominal_Clean'] = df_sipd_bm[sipd_target_col].apply(clean_currency)
        total_sipd_bm = df_sipd_bm['Nominal_Clean'].sum()

        # 4. AMBIL SELURUH DATA SKPD / FOTO 2 (TANPA FILTER KODE / REKENING)
        df_skpd['Is_Belanja_Modal'] = True

        # 5. PEMILIH KOLOM NOMINAL DI SIDEBAR (OTOMATIS PILIH UNNAMED: 3 ATAU YAG ADA ISINYA)
        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ Pengaturan Kolom SKPD")
        
        col_options = [c for c in df_skpd.columns if c not in ['Is_Belanja_Modal', 'Nominal_Clean']]
        
        # Cari kolom Unnamed: 3 atau kolom harga perolehan
        default_idx = 0
        for i, col in enumerate(col_options):
            if '3' in str(col) or 'HARGA' in str(col).upper() or 'NILAI' in str(col).upper():
                default_idx = i
                break

        selected_skpd_col = st.sidebar.selectbox(
            "Kolom Nominal SKPD (Foto 2):",
            options=col_options,
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
            st.caption(f"Kolom nominal aktif: **{selected_skpd_col}** | Seluruh baris di Foto 2 dimuat secara utuh.")
            st.dataframe(df_skpd_bm, use_container_width=True)

        with tab3:
            st.subheader("Transaksi Non-Belanja Modal yang Dieliminasi")
            st.write("Tidak ada transaksi yang dieliminasi (Semua baris diikutsertakan).")

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
