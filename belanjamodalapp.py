import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Aplikasi Rekap & Pembanding Belanja Modal (SIPD)", layout="wide")

st.title("🏛️ Aplikasi Rekap & Pembanding Belanja Modal (SIPD)")
st.write("Unggah file Realisasi LRA, Master RAK, dan Data Aplikasi Belanja Modal untuk membandingkan ketiganya secara otomatis.")

# Sidebar File Uploaders (3 File)
st.sidebar.header("📁 Unggah File")
file_lra = st.sidebar.file_uploader("1. File Realisasi LRA (Excel)", type=["xlsx", "xls"])
file_rak = st.sidebar.file_uploader("2. File Master RAK Belanja Modal (Excel)", type=["xlsx", "xls"])
file_app_bm = st.sidebar.file_uploader("3. File Data Aplikasi Belanja Modal (Excel)", type=["xlsx", "xls"])

def find_column(df, keywords):
    for col in df.columns:
        for kw in keywords:
            if kw.lower() in str(col).lower():
                return col
    return None

def clean_numeric(series):
    # Mengubah data teks/string menjadi angka bersih (float)
    if series is None:
        return 0
    return pd.to_numeric(
        series.astype(str)
        .str.replace('Rp', '', regex=False)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .str.strip(), 
        errors='coerce'
    ).fillna(0)

if file_lra and file_rak and file_app_bm:
    try:
        # 1. Load RAK Belanja Modal
        df_rak = pd.read_excel(file_rak)
        col_rak_kode = find_column(df_rak, ['kode', 'rekening']) or df_rak.columns[0]
        col_rak_nama = find_column(df_rak, ['nama', 'uraian', 'keterangan']) or (df_rak.columns[1] if len(df_rak.columns)>1 else df_rak.columns[0])
        col_rak_pagu = find_column(df_rak, ['pagu', 'anggaran', 'nilai', 'rak', 'jumlah']) or (df_rak.columns[2] if len(df_rak.columns)>2 else None)
        
        df_rak['Kode_Clean'] = df_rak[col_rak_kode].astype(str).str.strip()
        if col_rak_pagu:
            df_rak['Anggaran_RAK_Num'] = clean_numeric(df_rak[col_rak_pagu])
        else:
            df_rak['Anggaran_RAK_Num'] = 0

        # 2. Load LRA Data
        xl_lra = pd.ExcelFile(file_lra)
        sheet_lra = xl_lra.sheet_names[0]
        df_raw_lra = pd.read_excel(file_lra, sheet_name=sheet_lra)
        
        header_idx = 0
        for idx, row in df_raw_lra.head(15).iterrows():
            row_str = ' '.join(row.dropna().astype(str)).lower()
            if 'kode' in row_str or 'rekening' in row_str or 'realisasi' in row_str:
                header_idx = idx
                break
                
        if header_idx > 0:
            df_lra = pd.read_excel(file_lra, sheet_name=sheet_lra, skiprows=header_idx)
        else:
            df_lra = df_raw_lra.copy()

        col_lra_kode = find_column(df_lra, ['kode rekening', 'kode', 'rekening']) or df_lra.columns[0]
        col_lra_nama = find_column(df_lra, ['nama rekening', 'nama', 'uraian']) or df_lra.columns[1]
        col_lra_nilai = find_column(df_lra, ['nilai realisasi', 'realisasi', 'nilai', 'jumlah']) or df_lra.columns[-1]
        col_lra_skpd = find_column(df_lra, ['nama skpd', 'skpd', 'dinas', 'opd'])

        df_lra['Kode_Clean'] = df_lra[col_lra_kode].astype(str).str.strip()
        df_lra['Nilai_LRA_Num'] = clean_numeric(df_lra[col_lra_nilai])

        df_bm_lra = df_lra[df_lra['Kode_Clean'].str.startswith('5.2')].copy()
        if len(df_bm_lra) == 0:
            df_bm_lra = df_lra.copy()

        rekap_lra = df_bm_lra.groupby(['Kode_Clean', col_lra_nama]).agg(
            Realisasi_LRA=('Nilai_LRA_Num', 'sum'),
            Transaksi_LRA=('Nilai_LRA_Num', 'count')
        ).reset_index()

        # 3. Load Data Aplikasi Belanja Modal
        xl_app = pd.ExcelFile(file_app_bm)
        sheet_app = xl_app.sheet_names[0]
        df_app_raw = pd.read_excel(file_app_bm, sheet_name=sheet_app)
        
        col_app_kode = find_column(df_app_raw, ['kode', 'rekening', 'barang']) or df_app_raw.columns[0]
        col_app_nilai = find_column(df_app_raw, ['nilai', 'jumlah', 'harga', 'realisasi', 'total']) or df_app_raw.columns[-1]
        
        df_app_raw['Kode_Clean'] = df_app_raw[col_app_kode].astype(str).str.strip()
        df_app_raw['Nilai_App_Num'] = clean_numeric(df_app_raw[col_app_nilai])

        rekap_app = df_app_raw.groupby('Kode_Clean').agg(
            Nilai_Aplikasi_BM=('Nilai_App_Num', 'sum')
        ).reset_index()

        # Merge 3 Data
        df_compare = pd.merge(df_rak[['Kode_Clean', col_rak_nama, 'Anggaran_RAK_Num']], 
                              rekap_lra[['Kode_Clean', 'Realisasi_LRA', 'Transaksi_LRA']], 
                              on='Kode_Clean', how='outer')
        df_compare = pd.merge(df_compare, rekap_app, on='Kode_Clean', how='outer')
        
        df_compare['Anggaran_RAK'] = df_compare['Anggaran_RAK_Num'].fillna(0)
        df_compare['Realisasi_LRA'] = df_compare['Realisasi_LRA'].fillna(0)
        df_compare['Nilai_Aplikasi_BM'] = df_compare['Nilai_Aplikasi_BM'].fillna(0)
        
        # Perhitungan Selisih (Aman dari String Error)
        df_compare['Selisih (LRA vs App BM)'] = df_compare['Realisasi_LRA'] - df_compare['Nilai_Aplikasi_BM']
        df_compare['Sisa_Anggaran_RAK'] = df_compare['Anggaran_RAK'] - df_compare['Realisasi_LRA']

        st.sidebar.success("✅ Ketiga file berhasil diproses!")

        # Metrics Overview
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Pagu RAK", f"Rp {df_compare['Anggaran_RAK'].sum():,.0f}")
        m2.metric("Total Realisasi LRA", f"Rp {df_compare['Realisasi_LRA'].sum():,.0f}")
        m3.metric("Total Aplikasi BM", f"Rp {df_compare['Nilai_Aplikasi_BM'].sum():,.0f}")
        m4.metric("Selisih (LRA vs App)", f"Rp {df_compare['Selisih (LRA vs App BM)'].sum():,.0f}")

        # Data Tabs
        tab1, tab2, tab3 = st.tabs(["⚖️ Pembanding 3 Data", "🏢 Realisasi LRA per SKPD", "📑 Detail Realisasi LRA"])
        
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
                st.info("Kolom SKPD tidak ditemukan pada file LRA.")
            
        with tab3:
            st.subheader("Detail Transaksi LRA")
            st.dataframe(df_bm_lra, use_container_width=True)

        # Download Result
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
    st.info("Silakan unggah ketiga file Excel di sidebar sebelah kiri untuk mulai membandingkan.")
else:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri.")
