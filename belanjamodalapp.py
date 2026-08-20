import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Hasil Rekonsiliasi Belanja Modal", layout="wide")

st.sidebar.header("Upload File Excel")
file_rak = st.sidebar.file_uploader("1. RAK BELANJA MODAL.xlsx", type=["xlsx", "xls"])
file_lra = st.sidebar.file_uploader("2. DATA REALISASI (LRA/SIPD).xlsx", type=["xlsx", "xls"])
file_app_bm = st.sidebar.file_uploader("3. DATA APLIKASI BELANJA MODAL.xlsx", type=["xlsx", "xls"])

def clean_kode_rekening(kode):
    if pd.isna(kode) or kode is None:
        return ""
    s = str(kode).replace('\xa0', '').replace("'", "").replace('"', '').strip()
    s = re.sub(r'\s+', '', s)
    if s.lower() in ['nan', 'none', '0', ''] or not re.search(r'\d', s):
        return ""
    return s

def parse_indonesian_number(val):
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['nan', 'none', '-', '']:
        return 0.0

    val_clean = re.sub(r'[^\d\.,-]', '', val_str)
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
    for idx, row in df_raw.head(30).iterrows():
        row_str = ' '.join([str(v).lower() for v in row.dropna().values])
        if any(kw in row_str for kw in keywords):
            header_idx = idx
            break
            
    df = pd.read_excel(file, sheet_name=xl.sheet_names[0], header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def format_rupiah(val):
    if val < 0:
        return f"Rp -{abs(val):,.0f}".replace(',', '.')
    return f"Rp {val:,.0f}".replace(',', '.')

def get_status(row):
    lra = row['Total_LRA']
    app = row['Total_Aplikasi_BM']
    
    if lra > 0 and app == 0:
        return "Belum Input di Aplikasi Belanja Modal"
    elif lra == 0 and app > 0:
        return "Tidak Ada Realisasi di LRA"
    elif abs(lra - app) > 1:
        return "Beda Nilai Realisasi"
    else:
        return "Sesuai"

def get_column_safe(df, keywords, default_idx=0):
    cols = [c for c in df.columns if any(kw in c.lower() for kw in keywords)]
    if cols:
        return cols[0]
    if len(df.columns) > abs(default_idx):
        return df.columns[default_idx]
    return df.columns[0]

if file_rak and file_lra and file_app_bm:
    try:
        # 1. Baca File LRA
        df_lra = find_header_and_read(file_lra, ['kode', 'rekening', 'sp2d', 'realisasi', 'nilai'])
        col_k_lra = get_column_safe(df_lra, ['kode'], 0)
        col_n_lra = get_column_safe(df_lra, ['nama', 'uraian', 'rekening'], 1)
        col_v_lra = get_column_safe(df_lra, ['sp2d', 'realisasi', 'nilai', 'jumlah', 'kredit', 'debet'], -1)
        col_skpd = [c for c in df_lra.columns if 'skpd' in c.lower()]
        
        df_lra['Kode_Clean'] = df_lra[col_k_lra].apply(clean_kode_rekening)
        df_lra['Nilai_Num'] = df_lra[col_v_lra].apply(parse_indonesian_number)
        df_lra = df_lra[df_lra['Kode_Clean'] != ""].copy()
        df_lra = df_lra[~df_lra[col_k_lra].astype(str).str.lower().str.contains('total|jumlah', na=False)].copy()

        # 2. Sidebar Filter SKPD
        selected_skpd = "Semua SKPD"
        if col_skpd:
            skpd_col_name = col_skpd[0]
            list_skpd = ["Semua SKPD"] + sorted([str(x) for x in df_lra[skpd_col_name].dropna().unique()])
            st.sidebar.markdown("---")
            st.sidebar.header("Filter SKPD")
            selected_skpd = st.sidebar.selectbox("Pilih SKPD yang ingin direkonsiliasi:", list_skpd)
            
            if selected_skpd != "Semua SKPD":
                df_lra = df_lra[df_lra[skpd_col_name].astype(str) == selected_skpd].copy()

        # 3. Baca File RAK
        df_rak = find_header_and_read(file_rak, ['kode', 'rekening', 'pagu', 'anggaran'])
        col_k_rak = get_column_safe(df_rak, ['kode'], 0)
        col_n_rak = get_column_safe(df_rak, ['nama', 'uraian', 'rekening'], 1)
        col_v_rak = get_column_safe(df_rak, ['pagu', 'anggaran', 'nilai', 'jumlah'], -1)
        
        df_rak['Kode_Clean'] = df_rak[col_k_rak].apply(clean_kode_rekening)
        df_rak['Anggaran_Num'] = df_rak[col_v_rak].apply(parse_indonesian_number)

        # 4. Baca File Aplikasi BM (Dengan Validasi Ketat Kolom Nilai)
        df_app = find_header_and_read(file_app_bm, ['kode', 'uraian', 'pengadaan', 'nilai', 'harga', 'total'])
        col_k_app = get_column_safe(df_app, ['kode'], 0)
        
        # Cari kolom nilai rupiah murni yang BUKAN kolom kode
        candidate_cols = [c for c in df_app.columns if any(kw in c.lower() for kw in ['pengadaan', 'aset', 'nilai', 'harga', 'total', 'jumlah', 'realisasi']) and 'kode' not in c.lower()]
        
        if candidate_cols:
            col_v_app = candidate_cols[0]
        else:
            # Ambil kolom numerik paling kanan yang isinya bukan pola kode rekening
            col_v_app = df_app.columns[-1]
            for col in reversed(df_app.columns):
                if col != col_k_app:
                    sample_vals = df_app[col].dropna().astype(str).head(10).tolist()
                    # Jika isi kolom bukan kode rekening (tidak diawali 5.0 atau 5.2)
                    if not any(v.strip().startswith(('5.', '50', '52')) for v in sample_vals):
                        col_v_app = col
                        break

        df_app['Kode_Clean'] = df_app[col_k_app].apply(clean_kode_rekening)
        df_app['Nilai_App_Num'] = df_app[col_v_app].apply(parse_indonesian_number)

        # Proteksi Tambahan: Netralkan nilai jika tidak sengaja membaca kode rekening sebagai rupiah
        for idx, row in df_app.iterrows():
            clean_k = str(row['Kode_Clean']).replace('.', '')
            val_num = str(int(row['Nilai_App_Num'])) if row['Nilai_App_Num'] > 0 else ""
            if clean_k and val_num and (clean_k == val_num or val_num.startswith('5020') or val_num.startswith('5202')):
                df_app.at[idx, 'Nilai_App_Num'] = 0.0

        # 5. Agregasi Rekap
        rekap_lra = df_lra.groupby('Kode_Clean').agg(Total_LRA=('Nilai_Num', 'sum')).reset_index()
        rekap_app = df_app.groupby('Kode_Clean').agg(Total_Aplikasi_BM=('Nilai_App_Num', 'sum')).reset_index()
        rekap_rak = df_rak.groupby('Kode_Clean').agg(Total_RAK=('Anggaran_Num', 'sum')).reset_index()

        # Merge Data
        master_kode = pd.DataFrame({'Kode_Clean': list(set(rekap_lra['Kode_Clean']).union(set(rekap_app['Kode_Clean'])).union(set(rekap_rak['Kode_Clean'])))})
        
        nama_map = pd.concat([
            df_rak[['Kode_Clean', col_n_rak]].rename(columns={col_n_rak: 'Nama_Rekening'}),
            df_lra[['Kode_Clean', col_n_lra]].rename(columns={col_n_lra: 'Nama_Rekening'})
        ]).drop_duplicates(subset=['Kode_Clean'])

        df_final = pd.merge(master_kode, nama_map, on='Kode_Clean', how='left')
        df_final = pd.merge(df_final, rekap_lra, on='Kode_Clean', how='left')
        df_final = pd.merge(df_final, rekap_app, on='Kode_Clean', how='left')
        
        df_final['Total_LRA'] = df_final['Total_LRA'].fillna(0)
        df_final['Total_Aplikasi_BM'] = df_final['Total_Aplikasi_BM'].fillna(0)
        df_final['Selisih'] = df_final['Total_LRA'] - df_final['Total_Aplikasi_BM']
        df_final['Status'] = df_final.apply(get_status, axis=1)

        # Filter Kode Rekening Valid
        df_final = df_final[df_final['Kode_Clean'].str.contains(r'\.', na=False)].copy()
        df_final = df_final[(df_final['Total_LRA'] > 0) | (df_final['Total_Aplikasi_BM'] > 0)].sort_values('Kode_Clean')

        # 6. Tampilan Dashboard Utama
        st.write("Pemerintah Kabupaten Hulu Sungai Tengah")
        st.title(f"Hasil Rekonsiliasi SKPD: {selected_skpd}")

        tot_lra = df_final['Total_LRA'].sum()
        tot_app = df_final['Total_Aplikasi_BM'].sum()
        tot_selisih = tot_lra - tot_app

        c1, c2, c3 = st.columns(3)
        c1.metric("Total LRA", format_rupiah(tot_lra))
        c2.metric("Total Aplikasi Belanja Modal", format_rupiah(tot_app))
        c3.metric("Total Selisih", format_rupiah(tot_selisih))

        st.subheader("Detail Rekonsiliasi Per Kode Rekening")

        df_display = df_final.copy()
        df_display.rename(columns={'Kode_Clean': 'Kode_Rekening'}, inplace=True)
        df_display['Nama_Rekening'] = df_display['Nama_Rekening'].fillna("0")
        
        df_display['Total_LRA'] = df_display['Total_LRA'].apply(format_rupiah)
        df_display['Total_Aplikasi_BM'] = df_display['Total_Aplikasi_BM'].apply(format_rupiah)
        df_display['Selisih'] = df_display['Selisih'].apply(format_rupiah)

        st.dataframe(
            df_display[['Kode_Rekening', 'Nama_Rekening', 'Total_LRA', 'Total_Aplikasi_BM', 'Selisih', 'Status']], 
            use_container_width=True,
            hide_index=True
        )

        # Download Excel
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Hasil_Rekonsiliasi')

        st.download_button(
            label="📥 Download Hasil Rekonsiliasi (.xlsx)",
            data=output.getvalue(),
            file_name=f"Hasil_Rekonsiliasi_{selected_skpd}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
elif file_rak or file_lra or file_app_bm:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri untuk mulai membandingkan.")
else:
    st.info("Silakan unggah ketiga file Excel di sidebar kiri.")
