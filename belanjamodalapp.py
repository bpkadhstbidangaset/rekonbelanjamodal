import streamlit as st
import pandas as pd
import numpy as np

# Set Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Sistem Rekonsiliasi Belanja Modal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk Tampilan Modern & Clean
st.markdown("""
<style>
    .main-header {
        font-size:2rem;
        font-weight:700;
        color:#1E293B;
        margin-bottom:0px;
    }
    .sub-header {
        font-size:1rem;
        color:#64748B;
        margin-bottom:20px;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12);
    }
</style>
""", unsafe_allow_html=True)

# --- HEADER SECTION ---
st.markdown('<div class="main-header">🏛️ Mesin Rekonsiliasi & Penelusuran Selisih Belanja Modal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Pencocokan Otomatis antara Realisasi SIPD, Entry SKPD, dan RAK Rekening</div>', unsafe_allow_html=True)
st.divider()

# --- SIDEBAR: UPLOAD FILE ---
st.sidebar.header("📁 Upload File Data")
st.sidebar.info("Unggah file Excel/CSV sesuai data acuan Anda.")

file_rak = st.sidebar.file_uploader("1. RAK Rekening Belanja Modal (Acuan)", type=['xlsx', 'xls', 'csv'])
file_sipd = st.sidebar.file_uploader("2. Data Realisasi SIPD (Foto 1)", type=['xlsx', 'xls', 'csv'])
file_skpd = st.sidebar.file_uploader("3. Data Entry SKPD / Rincian Aset (Foto 2)", type=['xlsx', 'xls', 'csv'])

# Fungsi pembantu pembersihan nilai nominal / angka
def clean_currency(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).replace('Rp', '').replace(' ', '').strip()
    # Format Rupiah Indonesia: ribuan titik, desimal koma
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str and '.' not in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

# --- ISI UTAMA APLIKASI ---
if file_rak and file_sipd and file_skpd:
    try:
        # Load Data
        df_rak = pd.read_excel(file_rak) if file_rak.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_rak)
        df_sipd = pd.read_excel(file_sipd) if file_sipd.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_sipd)
        df_skpd = pd.read_excel(file_skpd) if file_skpd.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_skpd)

        st.success("✅ Semua data berhasil diunggah!")

        # Bersihkan Nama Kolom
        df_rak.columns = df_rak.columns.str.strip()
        df_sipd.columns = df_sipd.columns.str.strip()
        df_skpd.columns = df_skpd.columns.str.strip()

        # Deteksi Kolom Nominal SIPD (Debit / Realisasi)
        col_debit_sipd = [c for c in df_sipd.columns if 'debit' in c.lower() or 'realisasi' in c.lower() or 'saldo' in c.lower()]
        debit_col = col_debit_sipd[0] if col_debit_sipd else df_sipd.columns[-1]
        df_sipd['Nominal_SIPD'] = df_sipd[debit_col].apply(clean_currency)

        # Deteksi Kolom Nominal SKPD (Pengadaan / Aset)
        col_pengadaan_skpd = [c for c in df_skpd.columns if 'pengadaan' in c.lower() or 'aset' in c.lower()]
        skpd_col = col_pengadaan_skpd[0] if col_pengadaan_skpd else df_skpd.columns[-1]
        df_skpd['Nominal_SKPD'] = df_skpd[skpd_col].apply(clean_currency)

        # Ringkasan KPI Total
        total_sipd = df_sipd['Nominal_SIPD'].sum()
        total_skpd = df_skpd['Nominal_SKPD'].sum()
        total_selisih = total_sipd - total_skpd

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Realisasi SIPD", f"Rp {total_sipd:,.2f}")
        col2.metric("Total Entry SKPD", f"Rp {total_skpd:,.2f}")
        col3.metric(
            "Total Selisih (Varian)", 
            f"Rp {total_selisih:,.2f}", 
            delta=f"{-total_selisih:,.2f}", 
            delta_color="inverse"
        )

        st.markdown("---")

        # TAB LAYOUT
        tab1, tab2, tab3 = st.tabs(["🔍 Hasil Penelusuran & Selisih", "📋 Acuan RAK Rekening", "📁 Data Mentah"])

        with tab1:
            st.subheader("Detail Transaksi Realisasi SIPD")
            
            # Filter SKPD jika ada
            col_skpd_name = [c for c in df_sipd.columns if 'skpd' in c.lower()]
            if col_skpd_name:
                skpd_list = df_sipd[col_skpd_name[0]].dropna().unique()
                selected_skpd = st.multiselect("Filter SKPD / Unit Kerja:", options=skpd_list, default=skpd_list)
                filtered_sipd = df_sipd[df_sipd[col_skpd_name[0]].isin(selected_skpd)]
            else:
                filtered_sipd = df_sipd

            st.dataframe(filtered_sipd, use_container_width=True, height=350)

        with tab2:
            st.subheader("Master RAK Rekening Belanja Modal & Pemetaan KIB")
            st.dataframe(df_rak, use_container_width=True)

        with tab3:
            st.subheader("Pemeriksaan Data Mentah (Raw Data)")
            c1, c2 = st.columns(2)
            with c1:
                st.caption("Data Realisasi SIPD")
                st.dataframe(df_sipd.head(20), height=300)
            with c2:
                st.caption("Data Entry SKPD")
                st.dataframe(df_skpd.head(20), height=300)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
        st.info("Pastikan kolom pada file Excel/CSV sesuai dengan format data SIPD, SKPD, dan RAK Rekening Anda.")
else:
    # Tampilan Awal
    st.warning("👈 Silakan unggah ketiga file data pada menu di samping kiri untuk memulai proses pencocokan.")
    
    st.markdown("""
    ### 📌 Ketentuan File Input:
    1. **RAK Rekening Belanja Modal**: Berisi kolom *KODE REKENING*, *NAMA REKENING*, dan *KATEGORI ASET (PEMETAAN KIB)*.
    2. **Data Realisasi SIPD**: Berisi kolom *Tanggal*, *Uraian*, *Ref*, *No. Bukti*, *SKPD*, *Debit*, dll.
    3. **Data Entry SKPD**: Berisi rincian pengadaan aset (*KODE*, *URAIAN*, *PENGADAAN*, *ASET*, *SELISIH*).
    """)
