import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Aplikasi Rekap & Pembanding Belanja Modal (SIPD)", layout="wide")

st.title("🏛️ Aplikasi Rekap & Pembanding Belanja Modal (SIPD)")
st.write("Unggah file Realisasi LRA, Master RAK, dan Data Aplikasi Belanja Modal untuk membandingkan ketiganya secara otomatis.")

# Sidebar File Uploaders
st.sidebar.header("📁 Unggah File")
file_lra = st.sidebar.file_uploader("1. File Realisasi LRA (Excel)", type=["xlsx", "xls"])
file_rak = st.sidebar.file_uploader("2. File Master RAK Belanja Modal (Excel)", type=["xlsx", "xls"])
file_app_bm = st.sidebar.file_uploader("3. File Data Aplikasi Belanja Modal (Excel)", type=["xlsx", "xls"])

def clean_number_indonesia(series):
    """Membersihkan nilai angka format Indonesia (Contoh: 1.500.000,00 atau Rp 1.500.000)"""
    if series is None:
        return pd.Series(0)
    
    s = series.astype(str).str.strip()
    # Hapus teks Rp, spasi, dll
    s = s.str.replace('Rp', '', regex=False).str.replace(' ', '', regex=False)
    
    # Jika mengandung koma sebagai desimal (contoh: 1500000,00)
    # Ubah titik ribuan menjadi hilang, dan koma desimal menjadi titik
    def parse_val(val):
        try:
            if ',' in val and '.' in val:
                val = val.replace('.', '').replace(',', '.')
            elif ',' in val:
                val = val.replace(',', '.')
            elif '.' in val and len(val.split('.')[-1]) != 2: # Titik sebagai ribuan
                val = val.replace('.', '')
            return float(val)
        except:
            return 0.0

    return s.apply(parse_val).fillna(0.0)

def find_header_and_read(file):
    """Mencari baris header tabel yang benar secara otomatis"""
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df_raw = pd.read_excel(file, sheet_name=sheet)
    
    header_row = 0
    for idx, row in df_raw.head(20).iterrows():
        row_vals = [str(v).lower() for v in row.dropna().values]
        if any('kode' in v or 'rekening' in v or 'realisasi' in v or 'pagu' in v for v in row_vals):
            header_row = idx
            break
            
    df_clean = pd.read_excel(file, sheet_name=sheet, skiprows=header_row)
    # Hapus baris yang seluruhnya kosong
    df_clean = df_clean.dropna(how='all')
    return df_clean

def find_best_col(df, target_keywords, exclude_keywords=[]):
    """Mencari nama kolom berdasarkan kata kunci"""
    for col in df.columns:
        col_str = str(col).lower()
        if any(ex in col_str for ex in exclude_keywords):
            continue
        if any(kw in col_str for kw in target_keywords):
            return col
    return None

if file_lra and file_rak and file_app_bm:
    try:
        # 1. READ RAK
        df_rak = find_header_and_read(file_rak)
        col_rak_kode = find_best_col(df_rak, ['kode', 'rekening']) or df_rak.columns[0]
        col_rak_nama = find_best_col(df_rak, ['nama', 'uraian', 'keterangan']) or df_rak.columns[1]
        col_rak_pagu = find_best_col(df_rak, ['pagu', 'anggaran', 'rak', 'jumlah'], exclude_keywords=['no', 'satuan', 'volume']) or df_rak.columns[-1]

        df_rak['Kode_Clean'] = df_rak[col_rak_kode].astype(str).str.strip()
        df_rak['Anggaran_RAK_Num'] = clean_number_indonesia(df_rak[col_rak_pagu])

        # 2. READ LRA
        df_lra = find_header_and_read(file_lra)
        col_lra_kode = find_best_col(df_lra, ['kode', 'rekening']) or df_lra.columns[0]
        col_lra_nama = find_best_col(df_lra, ['nama', 'uraian']) or df_lra.columns[1]
        col_lra_nilai = find_best_col(df_lra, ['realisasi', 'nilai', 'sp2d', 'jumlah'], exclude_keywords=['no', 'satuan', 'volume', 'sisa']) or df_lra.columns[-1]
        col_lra_skpd = find_best_col(df_lra, ['skpd', 'opd', 'dinas'])

        df_lra['Kode_Clean'] = df_lra[col_lra_kode].astype(str).str.strip()
        df_lra['Nilai_LRA_Num'] = clean_number_indonesia(df_lra[col_lra_nilai])

        # Filter Belanja Modal (5.2)
        df_bm_lra = df_lra[df_lra['Kode_Clean'].str.startswith('5.2')].copy()
        if len(df_bm_lra) == 0:
            df_bm_lra = df_lra.copy()

        rekap_lra = df_bm_lra.groupby('Kode_Clean').agg(
            Realisasi_LRA=('Nilai_LRA_Num', 'sum'),
            Transaksi_LRA=('Nilai_LRA_Num', 'count')
        ).reset_index()

        # 3. READ APLIKASI BELANJA MODAL
        df_app = find_header_and_read(file_app_bm)
        col_app_kode = find_best_col(df_app, ['kode', 'rekening', 'barang']) or df_app.columns[0]
        col_app_nilai = find_best_col(df_app, ['nilai', 'harga', 'total', 'jumlah', 'realisasi'], exclude_keywords=['no', 'satuan', 'qty', 'jumlah barang']) or df_app.columns[-1]

        df_app['Kode_Clean'] = df_app[col_app_kode].astype(str).str.strip()
        df_app['Nilai_App_Num'] = clean_number_indonesia(df_app[col_app_nilai])

        rekap_app = df_app.groupby('Kode_Clean').agg(
            Nilai_Aplikasi_BM=('Nilai_App_Num', 'sum')
        ).reset_index()

        # MERGE DATA
        df_compare = pd.merge(df_rak[['Kode_Clean', col_rak_nama, 'Anggaran_RAK_Num']], 
                              rekap_lra, on='Kode_Clean', how='outer')
        df_compare = pd.merge(df_compare, rekap_app, on='Kode_Clean', how='outer')

        df_compare['Anggaran_RAK'] = df_compare['Anggaran_RAK_Num'].fillna(0)
        df_compare['Realisasi_LRA'] = df_compare['Realisasi_LRA'].fillna(0)
        df_compare['Nilai_Aplikasi_BM'] = df_compare['Nilai_Aplikasi_BM'].fillna(0)

        df_compare['Selisih (LRA vs App BM)'] = df_compare['Realisasi_LRA'] - df_compare['Nilai_Aplikasi_BM']
        df_compare['Sisa_Anggaran_RAK'] = df_compare['Anggaran_RAK'] - df_compare['Realisasi_LRA']

        st.sidebar.success("✅ Berhasil memproses data!")

        # DISPLAY METRICS
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Pagu RAK", f"Rp {df_compare['Anggaran_RAK'].sum():,.0f}")
        m2.metric("Total Realisasi LRA", f"Rp {df_compare['Realisasi_LRA'].sum():,.0f}")
        m3.metric("Total Aplikasi BM", f"Rp {df_compare['Nilai_Aplikasi_BM'].sum():,.0f}")
        m4.metric("Selisih (LRA vs App)", f"Rp {df_compare['Selisih (LRA vs App BM)'].sum():,.0f}")

        # TABS DISPLAY
        tab1, tab2, tab3 = st.tabs(["⚖️ Pembanding 3 Data", "🏢 Realisasi LRA per SKPD", "📑 Detail Transaksi"])

        with tab1:
            st.subheader("Tabel Pembanding: RAK vs Realisasi LRA vs Aplikasi Belanja Modal")
            st.dataframe(df_compare[['Kode_Clean', col_rak_nama, 'Anggaran_RAK', 'Realisasi_LRA', 'Nilai_Aplikasi_BM', 'Selisih (LRA vs App BM)', 'Sisa_Anggaran_RAK']], use_container_width=True)

        with tab2:
            st.subheader("Rekap Realisasi Belanja Modal per SKPD")
            if col_lra_skpd:
                group_skpd = df_bm_lra.groupby(col_lra_skpd).agg(
                    Total_Realisasi=('Nilai_LRA_Num', 'sum'),
                    Jumlah_Transaksi=('Nilai_LRA_Num', 'count')
                ).reset_index().sort_values(by='Total_Realisasi', ascending=False)
                st.dataframe(group_skpd, use_container_width=True)
            else:
                st.info("Kolom SKPD tidak terdeteksi pada file LRA.")

        with tab3:
            st.subheader("Detail Transaksi LRA")
            st.dataframe(df_bm_lra, use_container_width=True)

        # DOWNLOAD EXCEL
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_compare.to_excel(writer, index=False, sheet_name='Pembanding_3_Data')
            if col_lra_skpd:
                group_skpd.to_excel(writer, index=False, sheet_name='Rekap_SKPD')
            df_bm_lra.to_excel(writer, index=False, sheet_name='Detail_LRA')

        st.download_button(
            label="📥 Download Hasil Rekon 3 Data (Excel)",
            data=output.getvalue(),
            file_name="REKON_BELANJA_MODAL_3_DATA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
elif file_lra or file_rak or file_app_bm:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri untuk mem mulai pembandingan.")
else:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri.")
