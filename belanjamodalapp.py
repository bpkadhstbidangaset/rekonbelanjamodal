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

if file_lra and file_rak and file_app_bm:
    try:
        # 1. Load RAK Belanja Modal
        df_rak = pd.read_excel(file_rak)
        rak_code_col = df_rak.columns[0]
        rak_nama_col = df_rak.columns[1]
        rak_pagu_col = df_rak.columns[2] if len(df_rak.columns) > 2 else None
        
        df_rak['Kode_Clean'] = df_rak[rak_code_col].astype(str).str.strip()

        # 2. Load LRA Data
        xl_lra = pd.ExcelFile(file_lra)
        sheet_name = xl_lra.sheet_names[0]
        df_raw_lra = pd.read_excel(file_lra, sheet_name=sheet_name)
        
        header_row_idx = None
        for idx, row in df_raw_lra.head(10).iterrows():
            if 'Kode Rekening' in row.values:
                header_row_idx = idx
                break
                
        if header_row_idx is not None:
            df_lra = pd.read_excel(file_lra, sheet_name=sheet_name, skiprows=header_row_idx)
        else:
            df_lra = df_raw_lra.copy()
            
        df_lra['Kode_Clean'] = df_lra['Kode Rekening'].astype(str).str.strip()
        df_bm_lra = df_lra[df_lra['Kode_Clean'].str.startswith('5.2')].copy()
        
        rekap_lra = df_bm_lra.groupby(['Kode_Clean', 'Nama Rekening']).agg(
            Realisasi_LRA=('Nilai Realisasi', 'sum'),
            Transaksi_LRA=('Nilai Realisasi', 'count')
        ).reset_index()

        # 3. Load Data Aplikasi Belanja Modal
        xl_app = pd.ExcelFile(file_app_bm)
        sheet_app = xl_app.sheet_names[0]
        df_app_raw = pd.read_excel(file_app_bm, sheet_name=sheet_app)
        
        # Deteksi otomatis kolom kode rekening & nilai di aplikasi BM
        app_code_col = [c for c in df_app_raw.columns if 'kode' in str(c).lower() or 'rekening' in str(c).lower()]
        app_val_col = [c for c in df_app_raw.columns if 'nilai' in str(c).lower() or 'jumlah' in str(c).lower() or 'realisasi' in str(c).lower() or 'harga' in str(c).lower()]
        
        col_code = app_code_col[0] if app_code_col else df_app_raw.columns[0]
        col_val = app_val_col[0] if app_val_col else df_app_raw.columns[-1]
        
        df_app_raw['Kode_Clean'] = df_app_raw[col_code].astype(str).str.strip()
        rekap_app = df_app_raw.groupby('Kode_Clean').agg(
            Nilai_Aplikasi_BM=(col_val, 'sum')
        ).reset_index()

        # Merge 3 Data
        df_compare = pd.merge(df_rak[['Kode_Clean', rak_nama_col, rak_pagu_col]], rekap_lra[['Kode_Clean', 'Realisasi_LRA', 'Transaksi_LRA']], on='Kode_Clean', how='outer')
        df_compare = pd.merge(df_compare, rekap_app, on='Kode_Clean', how='outer')
        
        df_compare['Anggaran_RAK'] = df_compare[rak_pagu_col].fillna(0) if rak_pagu_col else 0
        df_compare['Realisasi_LRA'] = df_compare['Realisasi_LRA'].fillna(0)
        df_compare['Nilai_Aplikasi_BM'] = df_compare['Nilai_Aplikasi_BM'].fillna(0)
        
        # Perhitungan Selisih
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
            st.dataframe(df_compare[['Kode_Clean', rak_nama_col, 'Anggaran_RAK', 'Realisasi_LRA', 'Nilai_Aplikasi_BM', 'Selisih (LRA vs App BM)', 'Sisa_Anggaran_RAK']], use_container_width=True)
            
        with tab2:
            st.subheader("Rekap Realisasi Belanja Modal per SKPD")
            group_skpd = df_bm_lra.groupby(['Kode SKPD', 'Nama SKPD']).agg(
                Total_Realisasi=('Nilai Realisasi', 'sum'),
                Jumlah_Transaksi=('Nilai Realisasi', 'count')
            ).reset_index().sort_values(by='Total_Realisasi', ascending=False)
            st.dataframe(group_skpd, use_container_width=True)
            
        with tab3:
            st.subheader("Detail Transaksi LRA")
            st.dataframe(df_bm_lra, use_container_width=True)

        # Download Result
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_compare.to_excel(writer, index=False, sheet_name='Pembanding_3_Data')
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
