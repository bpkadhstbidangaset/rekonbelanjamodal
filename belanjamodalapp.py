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
    if s.lower() in ['nan', 'none', '0', '1', '2', '3', '4', '5', ''] or not re.search(r'\d', s):
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
        if ',' in val_clean:
            val_clean = val_clean.replace('.', '').replace(',', '.')
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
    for idx, row in df_raw.head(35).iterrows():
        row_str = ' '.join([str(v).lower() for v in row.dropna().values])
        # Pastikan mencari kata kunci murni dan sanasi header
        if any(kw in row_str for kw in keywords) and not all(str(v).strip().isdigit() for v in row.dropna().values):
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

if file_rak and file_lra and file_app_bm:
    try:
        # 1. Baca File LRA
        df_lra = find_header_and_read(file_lra, ['kode', 'rekening', 'sp2d', 'realisasi'])
        
        # Cari Kolom
        col_k_lra = [c for c in df_lra.columns if 'kode' in c.lower()][0]
        col_n_lra = [c for c in df_lra.columns if any(kw in c.lower() for kw in ['nama', 'uraian', 'rekening']) and 'kode' not in c.lower()][0]
        col_v_lra = [c for c in df_lra.columns if any(kw in c.lower() for kw in ['sp2d', 'realisasi', 'nilai', 'jumlah', 'kredit'])][-1]
        
        # Cari Kolom SKPD (Deteksi Kata 'SKPD', 'UNIT', atau 'ORGANISASI')
        col_skpd_list = [c for c in df_lra.columns if any(kw in c.lower() for kw in ['skpd', 'unit', 'organisasi'])]
        col_skpd = col_skpd_list[0] if col_skpd_list else None

        df_lra['Kode_Clean'] = df_lra[col_k_lra].apply(clean_kode_rekening)
        df_lra['Nilai_Num'] = df_lra[col_v_lra].apply(parse_indonesian_number)
        
        # Filter Baris Sampah
        df_lra = df_lra[df_lra['Kode_Clean'] != ""].copy()
        df_lra = df_lra[~df_lra[col_k_lra].astype(str).str.lower().str.contains('total|jumlah', na=False)].copy()

        # 2. Filter SKPD di Sidebar
        selected_skpd = "Semua SKPD"
        st.sidebar.markdown("---")
        st.sidebar.header("Filter SKPD")
        
        if col_skpd:
            list_skpd = ["Semua SKPD"] + sorted([str(x).strip() for x in df_lra[col_skpd].dropna().unique() if str(x).strip() != ""])
            selected_skpd = st.sidebar.selectbox("Pilih SKPD yang ingin direkonsiliasi:", list_skpd)
            
            if selected_skpd != "Semua SKPD":
                df_lra = df_lra[df_lra[col_skpd].astype(str).str.strip() == selected_skpd].copy()
        else:
            st.sidebar.warning("Kolom SKPD tidak terdeteksi di file LRA.")

        # 3. Baca File RAK
        df_rak = find_header_and_read(file_rak, ['kode', 'rekening', 'pagu', 'anggaran'])
        col_k_rak = [c for c in df_rak.columns if 'kode' in c.lower()][0]
        col_n_rak = [c for c in df_rak.columns if any(kw in c.lower() for kw in ['nama', 'uraian', 'rekening']) and 'kode' not in c.lower()][0]
        col_v_rak = [c for c in df_rak.columns if any(kw in c.lower() for kw in ['pagu', 'anggaran', 'nilai', 'jumlah'])][-1]

        df_rak['Kode_Clean'] = df_rak[col_k_rak].apply(clean_kode_rekening)
        df_rak['Anggaran_Num'] = df_rak[col_v_rak].apply(parse_indonesian_number)

        # 4. Baca File Aplikasi Belanja Modal
        df_app = find_header_and_read(file_app_bm, ['kode', 'uraian', 'pengadaan', 'aset'])
        
        col_k_app = [c for c in df_app.columns if 'kode' in c.lower()][0]
        col_n_app = [c for c in df_app.columns if 'uraian' in c.lower()][0]
        
        # Tentukan Kolom Nilai PENGADAAN
        col_v_app_candidates = [c for c in df_app.columns if 'pengadaan' in c.lower()]
        col_v_app = col_v_app_candidates[0] if col_v_app_candidates else df_app.columns[2]

        df_app['Kode_Clean'] = df_app[col_k_app].apply(clean_kode_rekening)
        df_app['Nilai_App_Num'] = df_app[col_v_app].apply(parse_indonesian_number)

        # Buang baris nomor urut header 1, 2, 3, 4, 5 pada Aplikasi BM
        df_app = df_app[~df_app[col_k_app].astype(str).str.strip().isin(['1', '2', '3', '4', '5'])].copy()
        df_app = df_app[df_app['Kode_Clean'] != ""].copy()

        # Netralkan jika kode rekening terbaca sebagai nilai rupiah
        for idx, row in df_app.iterrows():
            k_digits = re.sub(r'\D', '', str(row['Kode_Clean']))
            v_val = str(int(row['Nilai_App_Num'])) if row['Nilai_App_Num'] > 0 else ""
            if k_digits and v_val and k_digits == v_val:
                df_app.at[idx, 'Nilai_App_Num'] = 0.0

        # 5. Agregasi & Rekapitulasi
        rekap_lra = df_lra.groupby('Kode_Clean').agg(Total_LRA=('Nilai_Num', 'sum')).reset_index()
        rekap_app = df_app.groupby('Kode_Clean').agg(Total_Aplikasi_BM=('Nilai_App_Num', 'sum')).reset_index()
        rekap_rak = df_rak.groupby('Kode_Clean').agg(Total_RAK=('Anggaran_Num', 'sum')).reset_index()

        # Gabungkan Master Kode
        master_kode = pd.DataFrame({'Kode_Clean': list(set(rekap_lra['Kode_Clean']).union(set(rekap_app['Kode_Clean'])).union(set(rekap_rak['Kode_Clean'])))})
        master_kode = master_kode[~master_kode['Kode_Clean'].isin(['1', '2', '3', '4', '5', ''])].copy()

        nama_map = pd.concat([
            df_rak[['Kode_Clean', col_n_rak]].rename(columns={col_n_rak: 'Nama_Rekening'}),
            df_lra[['Kode_Clean', col_n_lra]].rename(columns={col_n_lra: 'Nama_Rekening'}),
            df_app[['Kode_Clean', col_n_app]].rename(columns={col_n_app: 'Nama_Rekening'})
        ]).drop_duplicates(subset=['Kode_Clean'])

        df_final = pd.merge(master_kode, nama_map, on='Kode_Clean', how='left')
        df_final = pd.merge(df_final, rekap_lra, on='Kode_Clean', how='left')
        df_final = pd.merge(df_final, rekap_app, on='Kode_Clean', how='left')
        
        df_final['Total_LRA'] = df_final['Total_LRA'].fillna(0)
        df_final['Total_Aplikasi_BM'] = df_final['Total_Aplikasi_BM'].fillna(0)
        df_final['Selisih'] = df_final['Total_LRA'] - df_final['Total_Aplikasi_BM']
        df_final['Status'] = df_final.apply(get_status, axis=1)

        # Ambil hanya kode rekening yang memiliki nilai / transaksi
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
        df_display['Nama_Rekening'] = df_display['Nama_Rekening'].fillna("Tidak Ada Uraian")
        
        df_display['Total_LRA'] = df_display['Total_LRA'].apply(format_rupiah)
        df_display['Total_Aplikasi_BM'] = df_display['Total_Aplikasi_BM'].apply(format_rupiah)
        df_display['Selisih'] = df_display['Selisih'].apply(format_rupiah)

        st.dataframe(
            df_display[['Kode_Rekening', 'Nama_Rekening', 'Total_LRA', 'Total_Aplikasi_BM', 'Selisih', 'Status']], 
            use_container_width=True,
            hide_index=True
        )

        # Download File Excel
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
