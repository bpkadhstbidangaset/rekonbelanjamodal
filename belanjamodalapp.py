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

# Helper pembersihan nominal Rupiah ke float
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

# Helper format ke Rupiah string
def format_rupiah(val):
    try:
        return f"Rp {float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "Rp 0,00"

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

# Helper membaca Excel SKPD (Presisi untuk Berbagai Format Header)
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    
    # Cari baris tabel yang memiliki kolom KODE/URAIAN dan PENGADAAN/ASET/SEMUA
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        # Hindari baris metadata seperti 'JENIS BELANJA : SEMUA'
        if ('KODE' in row_str and 'URAIAN' in row_str) or ('PENGADAAN' in row_str and 'ASET' in row_str):
            header_idx = idx
            break
        elif 'JENIS BELANJA' in row_str and ':' in row_str:
            continue
        elif any(k in row_str for k in ['5.2', '5.3']) and 'PEMERINTAH' not in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    # Abaikan baris akumulasi TOTAL / JUMLAH / Tanda Tangan Pejabat
    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH TOTAL|GRAND TOTAL|SUBTOTAL|T O T A L|PENGURUS BARANG').any(), axis=1)
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

        # 1. MAPPING KODE & URAIAN DARI RAK ACUAN
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

        # 3. IDENTIFIKASI & HITUNG SKPD TERLEBIH DAHULU
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

        # Filter baris rekening spesifik di SKPD (Hanya baris akun leaf 5.2... atau yang cocok di RAK)
        def match_skpd_row(row):
            first_col = str(row.iloc[0]).strip()
            # Abaikan nomor urut biasa (1, 2, 3), Program (5.02.01), Kegiatan (5.02.01.2.09), dll
            if first_col.count('.') < 3 and not (first_col.startswith('5.2') or first_col.startswith('5.3')):
                return ""
            if '5.1.01' in first_col:
                return ""
            
            first_cols_str = ' '.join([str(v) for v in row.iloc[:3].values if pd.notna(v)])
            matched = get_matched_code(first_cols_str)
            return matched if matched else ""

        df_skpd['Kode Rekening (Acuan)'] = df_skpd.apply(match_skpd_row, axis=1)
        df_skpd_bm = df_skpd[df_skpd['Kode Rekening (Acuan)'] != ""].copy()

        # Deteksi kolom nominal SKPD secara cerdas (Cari 'PENGADAAN', 'ASET', 'SEMUA', 'NILAI', atau kolom nominal pertama)
        col_nom_skpd = [c for c in df_skpd_bm.columns if any(k in c.upper() for k in ['PENGADAAN', 'ASET', 'SEMUA', 'TOTAL', 'JUMLAH', 'NILAI'])]
        
        if col_nom_skpd:
            df_skpd_bm['Nominal SKPD'] = df_skpd_bm[col_nom_skpd[0]].apply(clean_currency)
        else:
            # Fallback ambil kolom numerik terakhir
            df_skpd_bm['Nominal SKPD'] = df_skpd_bm.iloc[:, -1].apply(clean_currency)

        # Kumpulan kode rekening belanja modal yang ada di SKPD
        active_skpd_codes = set(df_skpd_bm['Kode Rekening (Acuan)'].unique())

        # 4. IDENTIFIKASI & HITUNG REALISASI SIPD
        def match_sipd_row(row):
            row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
            if '5.1.01' in row_str or 'GAJI' in row_str.upper() or 'IURAN JAMINAN' in row_str.upper():
                return ""
            matched = get_matched_code(row_str)
            if matched and matched in active_skpd_codes:
                return matched
            return ""

        df_sipd_filtered['Kode Rekening (Acuan)'] = df_sipd_filtered.apply(match_sipd_row, axis=1)
        df_sipd_bm = df_sipd_filtered[df_sipd_filtered['Kode Rekening (Acuan)'] != ""].copy()

        col_debit = [c for c in df_sipd_bm.columns if c.lower() == 'debit']
        sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-4]
        df_sipd_bm['Nominal Realisasi'] = df_sipd_bm[sipd_target_col].apply(clean_currency)

        # 5. TABEL ANALISIS REKONSILIASI
        grp_sipd = df_sipd_bm.groupby('Kode Rekening (Acuan)')['Nominal Realisasi'].sum().rename('Realisasi SIPD')
        grp_skpd = df_skpd_bm.groupby('Kode Rekening (Acuan)')['Nominal SKPD'].sum().rename('Entry SKPD')

        df_rekon = pd.concat([grp_sipd, grp_skpd], axis=1).fillna(0)
        df_rekon = df_rekon[df_rekon['Entry SKPD'] > 0].copy()

        df_rekon['Selisih'] = df_rekon['Realisasi SIPD'] - df_rekon['Entry SKPD']
        df_rekon['Status'] = df_rekon['Selisih'].apply(
            lambda x: '✅ Sesuai (Balance)' if abs(x) < 1 else ('🔻 Realisasi Lebih Kecil' if x < 0 else '🔺 Realisasi Lebih Besar')
        )
        df_rekon = df_rekon.reset_index().rename(columns={'Kode Rekening (Acuan)': 'Kode Rekening'})
        df_rekon['Uraian Rekening (RAK)'] = df_rekon['Kode Rekening'].apply(lambda x: rak_lookup.get(x, 'Belanja Modal'))

        df_rekon = df_rekon[['Kode Rekening', 'Uraian Rekening (RAK)', 'Realisasi SIPD', 'Entry SKPD', 'Selisih', 'Status']]
        df_rekon = df_rekon.sort_values(by='Selisih', key=abs, ascending=False)

        # Total Metrik
        total_sipd_bm = df_rekon['Realisasi SIPD'].sum()
        total_skpd_bm = df_rekon['Entry SKPD'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        # DASHBOARD METRIK
        st.success(f"✅ Rekonsiliasi selesai untuk: **{selected_skpd_target}**")

        col1, col2, col3 = st.columns(3)
        col1.metric("Realisasi SIPD (Belanja Modal)", format_rupiah(total_sipd_bm))
        col2.metric("Entry SKPD (Belanja Modal)", format_rupiah(total_skpd_bm))
        col3.metric(
            "Selisih Rekonsiliasi", 
            format_rupiah(total_selisih), 
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
            st.subheader("📊 Tabel Komparasi Belanja Modal (SIPD vs SKPD)")
            st.caption("Menampilkan akun Belanja Modal RAK yang di-entry oleh SKPD beserta nilai realisasinya di SIPD.")

            df_rekon_view = df_rekon.copy()
            df_rekon_view['Realisasi SIPD'] = df_rekon_view['Realisasi SIPD'].apply(format_rupiah)
            df_rekon_view['Entry SKPD'] = df_rekon_view['Entry SKPD'].apply(format_rupiah)
            df_rekon_view['Selisih'] = df_rekon_view['Selisih'].apply(format_rupiah)

            st.dataframe(df_rekon_view, use_container_width=True, hide_index=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rekon.to_excel(writer, index=False, sheet_name='Rekonsiliasi')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Unduh Rekapitulasi Rekonsiliasi (Excel)",
                data=excel_data,
                file_name=f"Rekonsiliasi_{selected_skpd_target}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab1:
            st.subheader(f"Transaksi Realisasi SIPD Terkait ({len(df_sipd_bm)} Baris)")
            cols_sipd_clean = [c for c in df_sipd_bm.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal Realisasi']]
            df_sipd_view = df_sipd_bm[['Kode Rekening (Acuan)'] + cols_sipd_clean + ['Nominal Realisasi']].copy()
            df_sipd_view['Nominal Realisasi'] = df_sipd_view['Nominal Realisasi'].apply(format_rupiah)
            st.dataframe(df_sipd_view, use_container_width=True, hide_index=True)

        with tab2:
            st.subheader(f"Rincian Pengadaan SKPD Terkait ({len(df_skpd_bm)} Baris)")
            df_skpd_view = df_skpd_bm.copy()
            
            # Format kolom agar rapi
            cols_to_keep = [c for c in df_skpd_view.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal SKPD']]
            df_skpd_display = df_skpd_view[['Kode Rekening (Acuan)'] + cols_to_keep + ['Nominal SKPD']].copy()
            df_skpd_display['Nominal SKPD'] = df_skpd_display['Nominal SKPD'].apply(format_rupiah)
            df_skpd_display = df_skpd_display.rename(columns={'Nominal SKPD': 'Nilai Bersih (Rupiah)'})

            st.dataframe(df_skpd_display, use_container_width=True, hide_index=True)

        with tab3:
            st.subheader("Baris SKPD di Luar Belanja Modal (Induk / Non-BM)")
            st.caption("Baris-baris ini merupakan header program/kegiatan atau belanja operasional lainnya.")
            
            df_skpd_elim = df_skpd[df_skpd['Kode Rekening (Acuan)'] == ""].copy()
            cols_elim_clean = [c for c in df_skpd_elim.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)']]
            df_elim_display = df_skpd_elim[cols_elim_clean].copy()
            
            st.dataframe(df_elim_display, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
