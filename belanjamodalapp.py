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

# Custom Styling UI
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

# Helper pembersihan nominal Rupiah secara robust
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

# Helper pembacaan Excel khusus Data Entry SKPD (Foto 2)
def parse_skpd_data(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        # Baca mentah
        df_raw = pd.read_excel(file, header=None)
        header_row = None
        
        # Cari baris yang mengandung kata 'PENGADAAN' atau 'ASET'
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
            if 'PENGADAAN' in row_str or 'ASET' in row_str or 'URAIAN' in row_str:
                header_row = idx
                break
        
        if header_row is not None:
            df = pd.read_excel(file, skiprows=header_row)
        else:
            df = pd.read_excel(file)
            
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

# Helper pembacaan Excel Data Realisasi SIPD (Foto 1)
def parse_sipd_data(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        df_raw = pd.read_excel(file, header=None)
        header_row = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
            if 'DEBIT' in row_str or 'REK' in row_str or 'NO. BUKTI' in row_str:
                header_row = idx
                break
        df = pd.read_excel(file, skiprows=header_row)
        
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- PROSES UTAMA ---
if file_rak and file_sipd and file_skpd:
    try:
        df_rak = pd.read_excel(file_rak) if file_rak.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_rak)
        df_sipd = parse_sipd_data(file_sipd)
        df_skpd = parse_skpd_data(file_skpd)

        st.success("✅ Semua file berhasil diunggah dan dibaca!")

        # 1. EVALUASI SIPD (Debit)
        col_debit = [c for c in df_sipd.columns if c.lower() == 'debit']
        if not col_debit:
            col_debit = [c for c in df_sipd.columns if 'debit' in c.lower()]
        
        sipd_col_target = col_debit[0] if col_debit else df_sipd.columns[-3]
        df_sipd['Nominal_Clean'] = df_sipd[sipd_col_target].apply(clean_currency)

        # 2. EVALUASI SKPD (Pengadaan / Aset)
        col_pengadaan = [c for c in df_skpd.columns if 'PENGADAAN' in c]
        col_aset = [c for c in df_skpd.columns if 'ASET' in c]
        
        if col_pengadaan:
            skpd_col_selected = col_pengadaan[0]
        elif col_aset:
            skpd_col_selected = col_aset[0]
        else:
            # Mengambil kolom ke-3 atau ke-4 jika nama tidak terdeteksi
            skpd_col_selected = df_skpd.columns[2] if len(df_skpd.columns) > 2 else df_skpd.columns[-1]

        # Bersihkan baris yang mengandung angka penomoran header '1', '2', '3', '4', '5'
        df_skpd = df_skpd[~df_skpd[skpd_col_selected].astype(str).str.strip().isin(['1', '2', '3', '4', '5'])]
        df_skpd['Nominal_Clean'] = df_skpd[skpd_col_selected].apply(clean_currency)

        # Hitung Nilai Total KPI
        total_sipd = df_sipd['Nominal_Clean'].sum()
        total_skpd = df_skpd['Nominal_Clean'].sum()
        total_selisih = total_sipd - total_skpd

        # Dashboard Metrik
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

        # TAMPILAN TAB
        tab1, tab2, tab3 = st.tabs(["🔍 Hasil Penelusuran & Selisih", "📋 Acuan RAK Rekening", "📁 Data Mentah"])

        with tab1:
            st.subheader("Data Realisasi SIPD Terdeteksi")
            st.caption(f"Kolom acuan nominal SIPD: **{sipd_col_target}** | Kolom acuan nominal SKPD: **{skpd_col_selected}**")
            
            # Filter unit SKPD
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
            st.subheader("Pemeriksaan Kolom Terbaca")
            c1, c2 = st.columns(2)
            with c1:
                st.caption(f"Data SIPD (Target: {sipd_col_target})")
                st.dataframe(df_sipd[['Nominal_Clean'] + [c for c in df_sipd.columns if c != 'Nominal_Clean']].head(10))
            with c2:
                st.caption(f"Data SKPD (Target: {skpd_col_selected})")
                st.dataframe(df_skpd[['Nominal_Clean'] + [c for c in df_skpd.columns if c != 'Nominal_Clean']].head(10))

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
