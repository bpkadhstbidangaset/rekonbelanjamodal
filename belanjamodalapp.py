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

def clean_number_id(series):
    """Mengubah format rupiah Indonesia menjadi Float"""
    if series is None:
        return pd.Series(0.0)
    
    s = series.astype(str).str.strip()
    s = s.str.replace('Rp.', '', regex=False).str.replace('Rp', '', regex=False).str.replace(' ', '', regex=False)
    
    def parse_val(val):
        if not val or str(val).lower() in ['nan', 'none', '-']:
            return 0.0
        try:
            val_str = str(val)
            if ',' in val_str and '.' in val_str:
                val_str = val_str.replace('.', '').replace(',', '.')
            elif ',' in val_str:
                val_str = val_str.replace(',', '.')
            elif '.' in val_str:
                parts = val_str.split('.')
                if len(parts[-1]) != 2:
                    val_str = val_str.replace('.', '')
            return float(val_str)
        except:
            return 0.0

    return s.apply(parse_val).fillna(0.0)

def read_lra_file(file):
    xl = pd.ExcelFile(file)
    df_raw = pd.read_excel(file, sheet_name=xl.sheet_names[0])
    
    header_idx = 0
    for idx, row in df_raw.head(15).iterrows():
        # Memastikan seluruh elemen diubah ke string terlebih dahulu
        row_str = ' '.join([str(v) for v in row.dropna().values]).lower()
        if 'kode rekening' in row_str or 'nilai sp2d' in row_str or 'realisasi' in row_str:
            header_idx = idx
            break
            
    df = pd.read_excel(file, sheet_name=xl.sheet_names[0], skiprows=header_idx)
    df.columns = [str(c) for c in df.columns]
    
    col_kode = [c for c in df.columns if 'kode rekening' in str(c).lower()]
    col_nama = [c for c in df.columns if 'nama rekening' in str(c).lower()]
    col_nilai = [c for c in df.columns if 'nilai sp2d' in str(c).lower() or 'nilai realisasi' in str(c).lower()]
    col_skpd = [c for c in df.columns if 'nama skpd' in str(c).lower()]

    kode_col = col_kode[0] if col_kode else df.columns[0]
    nama_col = col_nama[0] if col_nama else (df.columns[1] if len(df.columns)>1 else df.columns[0])
    nilai_col = col_nilai[0] if col_nilai else df.columns[-1]
    skpd_col = col_skpd[0] if col_skpd else None

    df['Kode_Clean'] = df[kode_col].astype(str).str.strip()
    df['Nilai_LRA_Num'] = clean_number_id(df[nilai_col])
    
    return df, kode_col, nama_col, nilai_col, skpd_col

def read_rak_file(file):
    xl = pd.ExcelFile(file)
    df = pd.read_excel(file, sheet_name=xl.sheet_names[0])
    df.columns = [str(c) for c in df.columns]
    
    col_kode = [c for c in df.columns if 'kode' in str(c).lower()]
    col_nama = [c for c in df.columns if 'nama' in str(c).lower() or 'rekening' in str(c).lower()]
    col_pagu = [c for c in df.columns if 'pagu' in str(c).lower() or 'anggaran' in str(c).lower() or 'nilai' in str(c).lower()]

    kode_col = col_kode[0] if col_kode else df.columns[0]
    nama_col = col_nama[0] if col_nama else (df.columns[1] if len(df.columns)>1 else df.columns[0])
    pagu_col = col_pagu[0] if col_pagu else None

    df['Kode_Clean'] = df[kode_col].astype(str).str.strip()
    if pagu_col:
        df['Anggaran_RAK_Num'] = clean_number_id(df[pagu_col])
    else:
        df['Anggaran_RAK_Num'] = 0.0
        
    return df, kode_col, nama_col, pagu_col

def read_app_bm_file(file):
    xl = pd.ExcelFile(file)
    df_raw = pd.read_excel(file, sheet_name=xl.sheet_names[0])
    
    header_idx = 0
    for idx, row in df_raw.head(15).iterrows():
        row_str = ' '.join([str(v) for v in row.dropna().values]).lower()
        if 'kode' in row_str and ('uraian' in row_str or 'pengadaan' in row_str):
            header_idx = idx
            break
            
    df = pd.read_excel(file, sheet_name=xl.sheet_names[0], skiprows=header_idx)
    df.columns = [str(c) for c in df.columns]
    
    col_kode = [c for c in df.columns if 'kode' in str(c).lower()]
    col_nilai = [c for c in df.columns if 'pengadaan' in str(c).lower() or 'aset' in str(c).lower()]

    kode_col = col_kode[0] if col_kode else df.columns[0]
    nilai_col = col_nilai[0] if col_nilai else df.columns[2]

    df['Kode_Clean'] = df[kode_col].astype(str).str.strip()
    df['Nilai_App_Num'] = clean_number_id(df[nilai_col])
    
    df = df[df['Kode_Clean'].str.contains(r'\.', na=False)].copy()
    
    return df, kode_col, nilai_col

if file_lra and file_rak and file_app_bm:
    try:
        df_lra, lra_k, lra_n, lra_val, col_lra_skpd = read_lra_file(file_lra)
        df_rak, rak_k, rak_n, rak_p = read_rak_file(file_rak)
        df_app, app_k, app_val = read_app_bm_file(file_app_bm)

        df_bm_lra = df_lra[df_lra['Kode_Clean'].str.startswith('5.2')].copy()
        if len(df_bm_lra) == 0:
            df_bm_lra = df_lra.copy()

        rekap_lra = df_bm_lra.groupby('Kode_Clean').agg(
            Realisasi_LRA=('Nilai_LRA_Num', 'sum'),
            Transaksi_LRA=('Nilai_LRA_Num', 'count')
        ).reset_index()

        rekap_app = df_app.groupby('Kode_Clean').agg(
            Nilai_Aplikasi_BM=('Nilai_App_Num', 'sum')
        ).reset_index()

        df_compare = pd.merge(df_rak[['Kode_Clean', rak_n, 'Anggaran_RAK_Num']], 
                              rekap_lra, on='Kode_Clean', how='outer')
        df_compare = pd.merge(df_compare, rekap_app, on='Kode_Clean', how='outer')

        df_compare['Anggaran_RAK'] = df_compare['Anggaran_RAK_Num'].fillna(0)
        df_compare['Realisasi_LRA'] = df_compare['Realisasi_LRA'].fillna(0)
        df_compare['Nilai_Aplikasi_BM'] = df_compare['Nilai_Aplikasi_BM'].fillna(0)

        df_compare['Selisih (LRA vs App BM)'] = df_compare['Realisasi_LRA'] - df_compare['Nilai_Aplikasi_BM']
        df_compare['Sisa_Anggaran_RAK'] = df_compare['Anggaran_RAK'] - df_compare['Realisasi_LRA']

        df_compare = df_compare[df_compare['Kode_Clean'].str.contains(r'\.', na=False)].copy()

        st.sidebar.success("✅ Berhasil memproses data!")

        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Pagu RAK", f"Rp {df_compare['Anggaran_RAK'].sum():,.0f}")
        m2.metric("Total Realisasi LRA", f"Rp {df_compare['Realisasi_LRA'].sum():,.0f}")
        m3.metric("Total Aplikasi BM", f"Rp {df_compare['Nilai_Aplikasi_BM'].sum():,.0f}")
        m4.metric("Selisih (LRA vs App)", f"Rp {df_compare['Selisih (LRA vs App BM)'].sum():,.0f}")

        tab1, tab2, tab3 = st.tabs(["⚖️ Pembanding 3 Data", "🏢 Realisasi LRA per SKPD", "📑 Detail Transaksi LRA"])

        with tab1:
            st.subheader("Tabel Pembanding: RAK vs Realisasi LRA vs Aplikasi Belanja Modal")
            st.dataframe(
                df_compare[['Kode_Clean', rak_n, 'Anggaran_RAK', 'Realisasi_LRA', 'Nilai_Aplikasi_BM', 'Selisih (LRA vs App BM)', 'Sisa_Anggaran_RAK']], 
                use_container_width=True
            )

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
    st.info("Silakan unggah ketiga file Excel di sidebar kiri untuk mulai membandingkan.")
else:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri.")
