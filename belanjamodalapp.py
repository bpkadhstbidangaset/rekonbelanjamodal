import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Aplikasi Rekap & Pembanding Belanja Modal (SIPD)", layout="wide")

st.title("🏛️ Aplikasi Rekap & Pembanding Belanja Modal (SIPD)")
st.write("Unggah file Realisasi LRA, Master RAK, dan Data Aplikasi Belanja Modal untuk membandingkan ketiganya secara otomatis.")

st.sidebar.header("📁 Unggah File")
file_lra = st.sidebar.file_uploader("1. File Realisasi LRA (Excel)", type=["xlsx", "xls"])
file_rak = st.sidebar.file_uploader("2. File Master RAK Belanja Modal (Excel)", type=["xlsx", "xls"])
file_app_bm = st.sidebar.file_uploader("3. File Data Aplikasi Belanja Modal (Excel)", type=["xlsx", "xls"])

def clean_kode_rekening(kode):
    """Pembersih Kode Rekening Ekstrem: Menghapus spasi tersembunyi, petik, dan karakter aneh"""
    if pd.isna(kode) or kode is None:
        return ""
    s = str(kode).replace('\xa0', '').replace("'", "").replace('"', '').strip()
    s = re.sub(r'\s+', '', s)
    return s

def parse_indonesian_number(val):
    """Mengkonversi teks angka format Indonesia (misal: 796.755.869,00) ke Float"""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['nan', 'none', '-', '']:
        return 0.0
    
    # Jika teks terlihat seperti kode rekening (punya banyak titik dan angka berpola), jangan jadikan nilai
    if len(re.findall(r'\.', val_str)) >= 2:
        return 0.0

    val_clean = re.sub(r'[^\d\.,]', '', val_str)
    if not val_clean:
        return 0.0

    try:
        if '.' in val_clean and ',' in val_clean:
            val_clean = val_clean.replace('.', '').replace(',', '.')
        elif ',' in val_clean:
            val_clean = val_clean.replace(',', '.')
        elif '.' in val_clean:
            parts = val_clean.split('.')
            if len(parts[-1]) != 2:
                val_clean = val_clean.replace('.', '')
        return float(val_clean)
    except:
        return 0.0

def find_header_and_read(file, keywords):
    xl = pd.ExcelFile(file)
    df_raw = pd.read_excel(file, sheet_name=xl.sheet_names[0], header=None)
    
    header_idx = 0
    for idx, row in df_raw.head(25).iterrows():
        row_str = ' '.join([str(v).lower() for v in row.dropna().values])
        if any(kw in row_str for kw in keywords):
            header_idx = idx
            break
            
    df = pd.read_excel(file, sheet_name=xl.sheet_names[0], header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_lra_file(file):
    df = find_header_and_read(file, ['kode rekening', 'nilai sp2d', 'realisasi'])
    
    col_kode = [c for c in df.columns if 'kode rekening' in c.lower()]
    col_nama = [c for c in df.columns if 'nama rekening' in c.lower()]
    col_nilai = [c for c in df.columns if 'nilai sp2d' in c.lower() or 'nilai realisasi' in c.lower() or 'realisasi' in c.lower()]
    col_skpd = [c for c in df.columns if 'nama skpd' in c.lower() or 'skpd' in c.lower()]

    kode_col = col_kode[0] if col_kode else df.columns[0]
    nama_col = col_nama[0] if col_nama else (df.columns[1] if len(df.columns)>1 else df.columns[0])
    nilai_col = col_nilai[0] if col_nilai else df.columns[-1]
    skpd_col = col_skpd[0] if col_skpd else None

    df['Kode_Clean'] = df[kode_col].apply(clean_kode_rekening)
    df['Nilai_LRA_Num'] = df[nilai_col].apply(parse_indonesian_number)
    
    return df, kode_col, nama_col, nilai_col, skpd_col

def read_rak_file(file):
    df = find_header_and_read(file, ['kode', 'rekening', 'kategori aset'])
    
    col_kode = [c for c in df.columns if 'kode' in c.lower()]
    col_nama = [c for c in df.columns if 'nama' in c.lower() or 'rekening' in c.lower() or 'uraian' in c.lower()]
    col_pagu = [c for c in df.columns if 'pagu' in c.lower() or 'anggaran' in c.lower() or 'nilai' in c.lower() or 'jumlah' in c.lower()]

    kode_col = col_kode[0] if col_kode else df.columns[0]
    nama_col = col_nama[0] if col_nama else (df.columns[1] if len(df.columns)>1 else df.columns[0])
    pagu_col = col_pagu[0] if col_pagu else None

    df['Kode_Clean'] = df[kode_col].apply(clean_kode_rekening)
    df['Anggaran_RAK_Num'] = df[pagu_col].apply(parse_indonesian_number) if pagu_col else 0.0
        
    return df, kode_col, nama_col, pagu_col

def read_app_bm_file(file):
    df = find_header_and_read(file, ['kode', 'uraian', 'pengadaan', 'nilai'])
    
    col_kode = [c for c in df.columns if 'kode' in c.lower()]
    
    # Cari spesifik kolom nilai/pengadaan, abaikan kolom kode
    col_nilai = [c for c in df.columns if ('pengadaan' in c.lower() or 'aset' in c.lower() or 'nilai' in c.lower() or 'harga' in c.lower() or 'total' in c.lower()) and 'kode' not in c.lower()]

    kode_col = col_kode[0] if col_kode else df.columns[0]
    # Jika tidak ketemu, default ambil kolom paling kanan/terakhir
    nilai_col = col_nilai[0] if col_nilai else df.columns[-1]

    df['Kode_Clean'] = df[kode_col].apply(clean_kode_rekening)
    df['Nilai_App_Num'] = df[nilai_col].apply(parse_indonesian_number)
    
    df = df[df['Kode_Clean'].str.contains(r'\.', na=False)].copy()
    return df, kode_col, nilai_col

if file_lra and file_rak and file_app_bm:
    try:
        df_lra, lra_k, lra_n, lra_val, col_lra_skpd = read_lra_file(file_lra)
        df_rak, rak_k, rak_n, rak_p = read_rak_file(file_rak)
        df_app, app_k, app_val = read_app_bm_file(file_app_bm)

        # Filter transaksi LRA (Kode diawali 5)
        df_bm_lra = df_lra[df_lra['Kode_Clean'].str.startswith('5')].copy()
        if len(df_bm_lra) == 0:
            df_bm_lra = df_lra.copy()

        rekap_lra = df_bm_lra.groupby('Kode_Clean').agg(
            Realisasi_LRA=('Nilai_LRA_Num', 'sum'),
            Transaksi_LRA=('Nilai_LRA_Num', 'count')
        ).reset_index()

        rekap_app = df_app.groupby('Kode_Clean').agg(
            Nilai_Aplikasi_BM=('Nilai_App_Num', 'sum')
        ).reset_index()

        rekap_rak = df_rak.groupby('Kode_Clean').agg(
            Anggaran_RAK=('Anggaran_RAK_Num', 'sum')
        ).reset_index()

        # Ambil master kode dari gabungan ketiganya
        master_kode = pd.DataFrame({'Kode_Clean': list(set(rekap_rak['Kode_Clean']).union(set(rekap_lra['Kode_Clean'])).union(set(rekap_app['Kode_Clean'])))})
        
        # Gabungkan Nama Rekening dari RAK / LRA
        nama_map = pd.concat([
            df_rak[['Kode_Clean', rak_n]].rename(columns={rak_n: 'Nama_Rekening'}),
            df_lra[['Kode_Clean', lra_n]].rename(columns={lra_n: 'Nama_Rekening'})
        ]).drop_duplicates(subset=['Kode_Clean'])

        df_compare = pd.merge(master_kode, nama_map, on='Kode_Clean', how='left')
        df_compare = pd.merge(df_compare, rekap_rak, on='Kode_Clean', how='left')
        df_compare = pd.merge(df_compare, rekap_lra, on='Kode_Clean', how='left')
        df_compare = pd.merge(df_compare, rekap_app, on='Kode_Clean', how='left')

        df_compare['Anggaran_RAK'] = df_compare['Anggaran_RAK'].fillna(0)
        df_compare['Realisasi_LRA'] = df_compare['Realisasi_LRA'].fillna(0)
        df_compare['Nilai_Aplikasi_BM'] = df_compare['Nilai_Aplikasi_BM'].fillna(0)

        df_compare['Selisih (LRA vs App BM)'] = df_compare['Realisasi_LRA'] - df_compare['Nilai_Aplikasi_BM']
        df_compare['Sisa_Anggaran_RAK'] = df_compare['Anggaran_RAK'] - df_compare['Realisasi_LRA']

        # Filter hanya kode valid dengan titik
        df_compare = df_compare[df_compare['Kode_Clean'].str.contains(r'\.', na=False)].sort_values('Kode_Clean').copy()

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
                df_compare[['Kode_Clean', 'Nama_Rekening', 'Anggaran_RAK', 'Realisasi_LRA', 'Nilai_Aplikasi_BM', 'Selisih (LRA vs App BM)', 'Sisa_Anggaran_RAK']], 
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
