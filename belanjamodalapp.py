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

# Custom Styling
st.markdown("""
<style>
    .main-header { font-size:2rem; font-weight:700; color:#1E293B; }
    .sub-header { font-size:1rem; color:#64748B; margin-bottom:20px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏛️ Mesin Rekonsiliasi & Penelusuran Selisih Belanja Modal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Pencocokan Otomatis antara Realisasi SIPD, Entry SKPD, dan RAK Rekening</div>', unsafe_allow_html=True)
st.divider()

# --- SIDEBAR UPLOAD ---
st.sidebar.header("📁 Upload File Data")
file_rak = st.sidebar.file_uploader("1. RAK Rekening Belanja Modal (Acuan)", type=['xlsx', 'xls', 'csv'])
file_sipd = st.sidebar.file_uploader("2. Data Realisasi SIPD (Foto 1)", type=['xlsx', 'xls', 'csv'])
file_skpd = st.sidebar.file_uploader("3. Data Entry SKPD / Rincian Aset (Foto 2)", type=['xlsx', 'xls', 'csv'])

# Helper pembersihan nominal
def clean_currency(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if val_str in ['-', '', 'NaN', 'nan']:
        return 0.0
    # Pembersihan format angka Indonesia (1.000.000,00)
    val_str = val_str.replace('Rp', '').replace(' ', '')
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str and '.' not in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

# Helper membaca Excel dengan deteksi header otomatis
def read_excel_smart(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    
    # Baca mentah untuk cari baris header
    df_raw = pd.read_excel(file, header=None)
    header_row = 0
    for idx, row in df_raw.iterrows():
        row_str = row.astype(str).str.lower().tolist()
        # Cari baris yang mengandung nama kolom kunci
        if any(key in ' '.join(row_str) for key in ['kode', 'uraian', 'debit', 'pengadaan', 'rekening']):
            header_row = idx
            break
            
    df = pd.read_excel(file, skiprows=header_row)
    df.columns = df.columns.astype(str).str.strip()
    return df

# --- PROSES DATA ---
if file_rak and file_sipd and file_skpd:
    try:
        df_rak = read_excel_smart(file_rak)
        df_sipd = read_excel_smart(file_sipd)
        df_skpd = read_excel_smart(file_skpd)

        st.success("✅ File berhasil dibaca dengan penyesuaian header otomatis!")

        # 1. EVUASI SIPD (Foto 1: Ambil kolom 'Debit', hindari 'Saldo')
        col_debit = [c for c in df_sipd.columns if c.lower() == 'debit']
        if not col_debit:
            col_debit = [c for c in df_sipd.columns if 'debit' in c.lower() or 'realisasi' in c.lower()]
        
        sipd_col_target = col_debit[0] if col_debit else df_sipd.columns[-3]
        df_sipd['Nominal_Clean'] = df_sipd[sipd_col_target].apply(clean_currency)

        # 2. EVALUASI SKPD (Foto 2: Ambil 'PENGADAAN' / 'ASET', abaikan kolom nomor urut '5')
        col_skpd_target = [c for c in df_skpd.columns if c.lower() in ['pengadaan', 'aset']]
        if not col_skpd_target:
            col_skpd_target = [c for c in df_skpd.columns if 'pengadaan' in c.lower() or 'aset' in c.lower()]
            
        skpd_col_selected = col_skpd_target[0] if col_skpd_target else df_skpd.columns[2]
        df_skpd['Nominal_Clean'] = df_skpd[skpd_col_selected].apply(clean_currency)

        # Filter baris angka urutan header jika ada (misal baris berisi angka 1, 2, 3, 4, 5)
        df_skpd = df_skpd[~df_skpd[skpd_col_selected].astype(str).str.strip().isin(['1', '2', '3', '4', '5'])]

        # Hitung Nilai KPI
        total_sipd = df_sipd['Nominal_Clean'].sum()
        total_skpd = df_skpd['Nominal_Clean'].sum()
        total_selisih = total_sipd - total_skpd

        # Tampilan KPI
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Realisasi SIPD (Debit)", f"Rp {total_sipd:,.2f}")
        col2.metric("Total Entry SKPD", f"Rp {total_skpd:,.2f}")
        col3.metric(
            "Total Selisih (SIPD - SKPD)", 
            f"Rp {total_selisih:,.2f}", 
            delta=f"{-total_selisih:,.2f}", 
            delta_color="inverse"
        )

        st.markdown("---")

        # TAB
        tab1, tab2, tab3 = st.tabs(["🔍 Hasil Penelusuran & Selisih", "📋 Acuan RAK Rekening", "📁 Data Mentah"])

        with tab1:
            st.subheader("Data Realisasi SIPD Terdeteksi")
            st.caption(f"Kolom acuan angka yang digunakan: **{sipd_col_target}**")
            
            # Filter SKPD
            skpd_cols = [c for c in df_sipd.columns if 'skpd' in c.lower()]
            if skpd_cols:
                unit_list = df_sipd[skpd_cols[0]].dropna().unique()
                selected_units = st.multiselect("Filter SKPD / Unit Kerja:", options=unit_list, default=unit_list)
                st.dataframe(df_sipd[df_sipd[skpd_cols[0]].isin(selected_units)], use_container_width=True)
            else:
                st.dataframe(df_sipd, use_container_width=True)

        with tab2:
            st.subheader("Master RAK Rekening Belanja Modal")
            st.dataframe(df_rak, use_container_width=True)

        with tab3:
            st.subheader("Deteksi Kolom Data")
            c1, c2 = st.columns(2)
            with c1:
                st.caption(f"SIPD - Kolom Digunakan: {sipd_col_target}")
                st.dataframe(df_sipd[['Nominal_Clean'] + list(df_sipd.columns[:-1])].head(10))
            with c2:
                st.caption(f"SKPD - Kolom Digunakan: {skpd_col_selected}")
                st.dataframe(df_skpd[['Nominal_Clean'] + list(df_skpd.columns[:-1])].head(10))

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
