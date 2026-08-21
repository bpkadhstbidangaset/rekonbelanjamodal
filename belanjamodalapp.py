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
st.markdown('<div class="sub-header">Pencocokan Otomatis Per SKPD (Realisasi SIPD vs Entry SKPD vs RAK Rekening)</div>', unsafe_allow_html=True)
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

# Helper membaca Excel SKPD (Foto 2) secara fleksibel
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if 'KODE' in row_str and ('PENGADAAN' in row_str or 'ASET' in row_str or 'URAIAN' in row_str):
            header_idx = idx
            break
            
    if header_idx is not None:
        df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    else:
        df = pd.read_excel(file) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file)
        
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

# Helper membaca Excel SIPD (Foto 1)
def parse_sipd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if 'DEBIT' in row_str or 'SKPD' in row_str or 'NO. BUKTI' in row_str:
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

        # 1. Dapatkan Daftar SKPD dari File SIPD
        col_skpd_sipd = [c for c in df_sipd.columns if 'skpd' in c.lower()]
        
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Target Rekonsiliasi")
        
        if col_skpd_sipd:
            list_skpd = sorted(df_sipd[col_skpd_sipd[0]].dropna().unique().tolist())
            
            # Cari default BPKPD / Badan Pengelola Keuangan
            default_index = 0
            for idx, name in enumerate(list_skpd):
                if 'BPKPD' in name.upper() or 'KEUANGAN' in name.upper():
                    default_index = idx
                    break
                    
            selected_skpd_target = st.sidebar.selectbox(
                "Pilih SKPD yang disandingkan:",
                options=list_skpd,
                index=default_index
            )
            
            # Filter Data SIPD khusus SKPD Terpilih
            df_sipd_filtered = df_sipd[df_sipd[col_skpd_sipd[0]] == selected_skpd_target].copy()
        else:
            df_sipd_filtered = df_sipd.copy()
            selected_skpd_target = "Semua SKPD"

        st.success(f"✅ Data diproses khusus untuk: **{selected_skpd_target}**")

        # 2. EVALUASI NOMINAL SIPD
        col_debit = [c for c in df_sipd_filtered.columns if c.lower() == 'debit']
        if not col_debit:
            col_debit = [c for c in df_sipd_filtered.columns if 'debit' in c.lower()]
        sipd_target_col = col_debit[0] if col_debit else df_sipd_filtered.columns[-3]
        df_sipd_filtered['Nominal_Clean'] = df_sipd_filtered[sipd_target_col].apply(clean_currency)

        # 3. EVALUASI NOMINAL SKPD
        # Deteksi kolom angka nominal di Data Entry SKPD
        candidate_cols = [c for c in df_skpd.columns if any(k in c for k in ['PENGADAAN', 'ASET', '3', '4'])]
        if candidate_cols:
            skpd_target_col = candidate_cols[0]
        else:
            skpd_target_col = df_skpd.columns[2] if len(df_skpd.columns) > 2 else df_skpd.columns[-1]

        # Abaikan baris penomoran '1', '2', '3', '4', '5'
        df_skpd_clean = df_skpd[~df_skpd[skpd_target_col].astype(str).str.strip().isin(['1', '2', '3', '4', '5'])].copy()
        df_skpd_clean['Nominal_Clean'] = df_skpd_clean[skpd_target_col].apply(clean_currency)

        # Hitung Totals
        total_sipd = df_sipd_filtered['Nominal_Clean'].sum()
        total_skpd = df_skpd_clean['Nominal_Clean'].sum()
        total_selisih = total_sipd - total_skpd

        # Dashboard Metrik
        col1, col2, col3 = st.columns(3)
        col1.metric(f"Realisasi SIPD ({selected_skpd_target})", f"Rp {total_sipd:,.2f}")
        col2.metric("Entry SKPD (Rincian Aset)", f"Rp {total_skpd:,.2f}")
        col3.metric(
            "Selisih Rekonsiliasi", 
            f"Rp {total_selisih:,.2f}", 
            delta=f"{-total_selisih:,.2f}", 
            delta_color="inverse"
        )

        st.markdown("---")

        # TAMPILAN TAB
        tab1, tab2, tab3 = st.tabs(["🔍 Detail Transaksi & Pencocokan", "📋 Acuan RAK Rekening", "📁 Preview Data Mentah"])

        with tab1:
            st.subheader(f"Transaksi Realisasi SIPD - {selected_skpd_target}")
            st.dataframe(df_sipd_filtered, use_container_width=True)

        with tab2:
            st.subheader("Master RAK Rekening Belanja Modal")
            st.dataframe(df_rak, use_container_width=True)

        with tab3:
            st.subheader("Pemeriksaan Kolom & Pembersihan")
            c1, c2 = st.columns(2)
            with c1:
                st.caption(f"Data Realisasi SIPD Terfilter ({len(df_sipd_filtered)} baris)")
                st.dataframe(df_sipd_filtered[['Nominal_Clean'] + [c for c in df_sipd_filtered.columns if c != 'Nominal_Clean']].head(10))
            with c2:
                st.caption(f"Data Entry SKPD Terdeteksi ({len(df_skpd_clean)} baris)")
                st.dataframe(df_skpd_clean[['Nominal_Clean'] + [c for c in df_skpd_clean.columns if c != 'Nominal_Clean']].head(10))

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
