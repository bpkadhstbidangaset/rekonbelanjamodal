import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Aplikasi Rekap & Pembanding Belanja Modal (SIPD)", layout="wide")

st.title("🏛️ Aplikasi Rekap & Pembanding Belanja Modal (SIPD)")
st.write("Unggah file Realisasi LRA, Master RAK, dan Data Aplikasi Belanja Modal untuk membandingkan ketiganya secara otomatis.")

# Sidebar File Uploaders
st.sidebar.header("📁 1. Unggah File")
file_lra = st.sidebar.file_uploader("1. File Realisasi LRA (Excel)", type=["xlsx", "xls"])
file_rak = st.sidebar.file_uploader("2. File Master RAK Belanja Modal (Excel)", type=["xlsx", "xls"])
file_app_bm = st.sidebar.file_uploader("3. File Data Aplikasi Belanja Modal (Excel)", type=["xlsx", "xls"])

def clean_number_indonesia(series):
    if series is None:
        return pd.Series(0)
    
    s = series.astype(str).str.strip()
    s = s.str.replace('Rp', '', regex=False).str.replace(' ', '', regex=False)
    
    def parse_val(val):
        try:
            if ',' in val and '.' in val:
                val = val.replace('.', '').replace(',', '.')
            elif ',' in val:
                val = val.replace(',', '.')
            elif '.' in val and len(val.split('.')[-1]) != 2:
                val = val.replace('.', '')
            return float(val)
        except:
            return 0.0

    return s.apply(parse_val).fillna(0.0)

def find_header_and_read(file):
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
    df_clean = df_clean.dropna(how='all')
    return df_clean

if file_lra and file_rak and file_app_bm:
    try:
        # Load raw dataframes
        df_lra = find_header_and_read(file_lra)
        df_rak = find_header_and_read(file_rak)
        df_app = find_header_and_read(file_app_bm)

        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ 2. Pemetaan Kolom Excel")

        # 1. Map RAK
        st.sidebar.subheader("Master RAK")
        col_rak_kode = st.sidebar.selectbox("Kolom Kode Rekening RAK", df_rak.columns, index=0)
        col_rak_nama = st.sidebar.selectbox("Kolom Nama Rekening RAK", df_rak.columns, index=min(1, len(df_rak.columns)-1))
        col_rak_pagu = st.sidebar.selectbox("Kolom Pagu Anggaran RAK", df_rak.columns, index=len(df_rak.columns)-1)

        # 2. Map LRA
        st.sidebar.subheader("Realisasi LRA")
        col_lra_kode = st.sidebar.selectbox("Kolom Kode Rekening LRA", df_lra.columns, index=0)
        col_lra_nama = st.sidebar.selectbox("Kolom Nama Rekening LRA", df_lra.columns, index=min(1, len(df_lra.columns)-1))
        col_lra_nilai = st.sidebar.selectbox("Kolom Realisasi Rp LRA", df_lra.columns, index=len(df_lra.columns)-1)
        col_lra_skpd = st.sidebar.selectbox("Kolom SKPD (Opsional)", [None] + list(df_lra.columns), index=0)

        # 3. Map App BM
        st.sidebar.subheader("Aplikasi Belanja Modal")
        col_app_kode = st.sidebar.selectbox("Kolom Kode Rekening App BM", df_app.columns, index=0)
        col_app_nilai = st.sidebar.selectbox("Kolom Nilai Rp App BM", df_app.columns, index=len(df_app.columns)-1)

        # Process Data
        df_rak['Kode_Clean'] = df_rak[col_rak_kode].astype(str).str.strip()
        df_rak['Anggaran_RAK_Num'] = clean_number_indonesia(df_rak[col_rak_pagu])

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

        df_app['Kode_Clean'] = df_app[col_app_kode].astype(str).str.strip()
        df_app['Nilai_App_Num'] = clean_number_indonesia(df_app[col_app_nilai])

        rekap_app = df_app.groupby('Kode_Clean').agg(
            Nilai_Aplikasi_BM=('Nilai_App_Num', 'sum')
        ).reset_index()

        # Merge
        df_compare = pd.merge(df_rak[['Kode_Clean', col_rak_nama, 'Anggaran_RAK_Num']], 
                              rekap_lra, on='Kode_Clean', how='outer')
        df_compare = pd.merge(df_compare, rekap_app, on='Kode_Clean', how='outer')

        df_compare['Anggaran_RAK'] = df_compare['Anggaran_RAK_Num'].fillna(0)
        df_compare['Realisasi_LRA'] = df_compare['Realisasi_LRA'].fillna(0)
        df_compare['Nilai_Aplikasi_BM'] = df_compare['Nilai_Aplikasi_BM'].fillna(0)

        df_compare['Selisih (LRA vs App BM)'] = df_compare['Realisasi_LRA'] - df_compare['Nilai_Aplikasi_BM']
        df_compare['Sisa_Anggaran_RAK'] = df_compare['Anggaran_RAK'] - df_compare['Realisasi_LRA']

        st.sidebar.success("✅ Berhasil memproses data!")

        # Metrics
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Pagu RAK", f"Rp {df_compare['Anggaran_RAK'].sum():,.0f}")
        m2.metric("Total Realisasi LRA", f"Rp {df_compare['Realisasi_LRA'].sum():,.0f}")
        m3.metric("Total Aplikasi BM", f"Rp {df_compare['Nilai_Aplikasi_BM'].sum():,.0f}")
        m4.metric("Selisih (LRA vs App)", f"Rp {df_compare['Selisih (LRA vs App BM)'].sum():,.0f}")

        # Tabs
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
                st.info("Pilih kolom SKPD di sidebar kiri untuk menampilkan rekap per SKPD.")

        with tab3:
            st.subheader("Detail Transaksi LRA")
            st.dataframe(df_bm_lra, use_container_width=True)

        # Download
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
    st.info("Silakan unggah ketiga file Excel di sidebar kiri untuk memulai pembandingan.")
else:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri.")
