import streamlit as st
import pandas as pd
import numpy as np
import re
import io

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
st.markdown('<div class="sub-header">Pencocokan Transaksi Presisi & Analisis Selisih Berdasarkan Acuan RAK</div>', unsafe_allow_html=True)
st.divider()

# --- SIDEBAR UPLOAD ---
st.sidebar.header("📁 Upload File Data")
file_rak = st.sidebar.file_uploader("1. RAK Rekening Belanja Modal (Acuan)", type=['xlsx', 'xls', 'csv'])
file_sipd = st.sidebar.file_uploader("2. Data Realisasi SIPD (Foto 1)", type=['xlsx', 'xls', 'csv'])
file_skpd = st.sidebar.file_uploader("3. Data Entry SKPD / Rincian Aset (Foto 2)", type=['xlsx', 'xls', 'csv'])

# Helper normalisasi kode
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

    # Abaikan baris akumulasi TOTAL / JUMLAH
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

        # 1. AMBIL KODE REKENING & MAPPING URAIAN DARI ACUAN RAK
        col_kode_rek = [c for c in df_rak.columns if 'KODE' in c and 'REK' in c]
        col_uraian_rek = [c for c in df_rak.columns if 'URAIAN' in c and 'REK' in c]
        col_kode_kat = [c for c in df_rak.columns if 'KODE' in c and 'KAT' in c]

        col_rek_target = col_kode_rek[0] if col_kode_rek else df_rak.columns[-1]
        col_uraian_target = col_uraian_rek[0] if col_uraian_rek else df_rak.columns[-2]
        col_kat_target = col_kode_kat[0] if col_kode_kat else df_rak.columns[0]

        rak_lookup = {}
        for _, r in df_rak.iterrows():
            k_rek = str(r[col_rek_target]).strip() if pd.notna(r[col_rek_target]) else ""
            u_rek = str(r[col_uraian_target]).strip() if pd.notna(r[col_uraian_target]) else ""
            if k_rek and len(k_rek) > 3:
                rak_lookup[k_rek] = u_rek
                rak_lookup[normalize_code(k_rek)] = (k_rek, u_rek)

        valid_raw_codes = [k for k in df_rak[col_rek_target].dropna().astype(str).str.strip().tolist() if len(k) > 3]
        if col_kat_target in df_rak.columns:
            valid_raw_codes += [k for k in df_rak[col_kat_target].dropna().astype(str).str.strip().tolist() if len(k) > 3]

        norm_acuan_set = {normalize_code(k) for k in valid_raw_codes if len(normalize_code(k)) >= 5}

        # 2. FILTER SKPD TARGET PADA SIPD
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

        # 3. FILTER SIPD & EKSTRAK KODE REKENING
        def get_matched_code(row_str):
            row_norm = normalize_code(row_str)
            for raw_c in valid_raw_codes:
                if raw_c in row_str:
                    return raw_c
            for n_c in norm_acuan_set:
                if n_c in row_norm:
                    if n_c in rak_lookup and isinstance(rak_lookup[n_c], tuple):
                        return rak_lookup[n_c][0]
                    return n_c
            return None

        def match_sipd(row):
            row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
            matched = get_matched_code(row_str)
            return matched if matched else ""

        df_sipd_filtered['Kode_Rekening_Cocok'] = df_sipd_filtered.apply(match_sipd, axis=1)
        df_sipd_bm = df_sipd_filtered[df_sipd_filtered['Kode_Rekening_Cocok'] != ""].copy()

        # Hitung Nominal Realisasi SIPD
        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-4]
        df_sipd_bm['Nominal_Clean'] = df_sipd_bm[sipd_target_col].apply(clean_currency)
        total_sipd_bm = df_sipd_bm['Nominal_Clean'].sum()

        # 4. FILTER SKPD & EKSTRAK KODE REKENING
        def match_skpd(row):
            first_cols_str = ' '.join([str(v) for v in row.iloc[:3].values if pd.notna(v)])
            matched = get_matched_code(first_cols_str)
            return matched if matched else ""

        df_skpd['Kode_Rekening_Cocok'] = df_skpd.apply(match_skpd, axis=1)
        df_skpd_bm = df_skpd[df_skpd['Kode_Rekening_Cocok'] != ""].copy()

        # Hitung Nominal SKPD (Kolom SEMUA)
        col_semua = [c for c in df_skpd_bm.columns if c.upper() in ['SEMUA', 'TOTAL', 'JUMLAH', 'NILAI']]
        if col_semua:
            df_skpd_bm['Nominal_Clean'] = df_skpd_bm[col_semua[0]].apply(clean_currency)
        else:
            df_skpd_bm['Nominal_Clean'] = df_skpd_bm.iloc[:, -3].apply(clean_currency)

        total_skpd_bm = df_skpd_bm['Nominal_Clean'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        # 5. TABEL ANALISIS SELISIH PER REKENING
        grp_sipd = df_sipd_bm.groupby('Kode_Rekening_Cocok')['Nominal_Clean'].sum().rename('Realisasi_SIPD')
        grp_skpd = df_skpd_bm.groupby('Kode_Rekening_Cocok')['Nominal_Clean'].sum().rename('Entry_SKPD')

        df_rekon = pd.concat([grp_sipd, grp_skpd], axis=1).fillna(0)
        df_rekon['Selisih'] = df_rekon['Realisasi_SIPD'] - df_rekon['Entry_SKPD']
        df_rekon['Status'] = df_rekon['Selisih'].apply(lambda x: '✅ Cocok' if abs(x) < 1 else ('🔻 Kurang Catat di SIPD' if x < 0 else '🔺 Belum Terinput SKPD'))
        df_rekon = df_rekon.reset_index().rename(columns={'index': 'Kode_Rekening', 'Kode_Rekening_Cocok': 'Kode_Rekening'})

        # Tambahkan Deskripsi Uraian
        df_rekon['Uraian_Rekening'] = df_rekon['Kode_Rekening'].apply(lambda x: rak_lookup.get(x, ''))
        
        # Urutkan Kolom
        df_rekon = df_rekon[['Kode_Rekening', 'Uraian_Rekening', 'Realisasi_SIPD', 'Entry_SKPD', 'Selisih', 'Status']]
        df_rekon = df_rekon.sort_values(by='Selisih', key=abs, ascending=False)

        # DASHBOARD METRIK
        st.success(f"✅ Rekonsiliasi selesai untuk: **{selected_skpd_target}**")

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

        # TAMPILAN TAB UTAMA
        tab_rekon, tab1, tab2, tab3 = st.tabs([
            "⚖️ Rekonsiliasi & Analisis Selisih", 
            "🔍 Detail Realisasi SIPD", 
            "📋 Detail Entry SKPD", 
            "📁 Data SKPD Dieliminasi"
        ])

        with tab_rekon:
            st.subheader("📊 Rincian Selisih Antara SIPD & SKPD Per Rekening")
            st.caption("Daftar diurutkan dari rekening dengan selisih nominal terbesar.")
            
            # Format Rupiah untuk Tampilan
            df_rekon_display = df_rekon.copy()
            df_rekon_display['Realisasi_SIPD'] = df_rekon_display['Realisasi_SIPD'].apply(lambda x: f"Rp {x:,.2f}")
            df_rekon_display['Entry_SKPD'] = df_rekon_display['Entry_SKPD'].apply(lambda x: f"Rp {x:,.2f}")
            df_rekon_display['Selisih'] = df_rekon_display['Selisih'].apply(lambda x: f"Rp {x:,.2f}")
            
            st.dataframe(df_rekon_display, use_container_width=True)

            # Tombol Download Hasil Rekonsiliasi
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rekon.to_excel(writer, index=False, sheet_name='Rekonsiliasi')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Unduh Laporan Selisih (Excel)",
                data=excel_data,
                file_name=f"Hasil_Rekonsiliasi_{selected_skpd_target}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab1:
            st.subheader(f"Transaksi SIPD Terkait ({len(df_sipd_bm)} Baris)")
            st.dataframe(df_sipd_bm, use_container_width=True)

        with tab2:
            st.subheader(f"Rincian SKPD Terkait ({len(df_skpd_bm)} Baris)")
            st.dataframe(df_skpd_bm, use_container_width=True)

        with tab3:
            st.subheader("Baris SKPD di Luar Acuan RAK")
            df_skpd_elim = df_skpd[df_skpd['Kode_Rekening_Cocok'] == ""].copy()
            st.dataframe(df_skpd_elim, use_container_width=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
