import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Aplikasi Rekap & Pembanding Belanja Modal (SIPD)", layout="wide")

st.title("🏛️ Aplikasi Rekap & Pembanding Belanja Modal (SIPD)")
st.write("Unggah file Realisasi LRA dan Master RAK Belanja Modal untuk membandingkan anggaran (pagu RAK) dengan realisasi.")

# Sidebar File Uploaders
st.sidebar.header("📁 Unggah File")
file_lra = st.sidebar.file_uploader("1. File Realisasi LRA (Excel)", type=["xlsx", "xls"])
file_rak = st.sidebar.file_uploader("2. File Master RAK Belanja Modal (Excel)", type=["xlsx", "xls"])

if file_lra and file_rak:
    try:
        # Load RAK Belanja Modal
        df_rak = pd.read_excel(file_rak)
        st.sidebar.success("✅ File RAK Belanja Modal Berhasil Dimuat")
        
        # Load LRA Data
        xl_lra = pd.ExcelFile(file_lra)
        sheet_name = xl_lra.sheet_names[0]
        df_raw = pd.read_excel(file_lra, sheet_name=sheet_name)
        
        # Detect header row dynamically
        header_row_idx = None
        for idx, row in df_raw.head(10).iterrows():
            if 'Kode Rekening' in row.values:
                header_row_idx = idx
                break
                
        if header_row_idx is not None:
            df_lra = pd.read_excel(file_lra, sheet_name=sheet_name, skiprows=header_row_idx)
        else:
            df_lra = df_raw.copy()
            
        df_lra['Kode_Clean'] = df_lra['Kode Rekening'].astype(str).str.strip()
        
        # Filter Belanja Modal (5.2)
        df_bm = df_lra[df_lra['Kode_Clean'].str.startswith('5.2')].copy()
        
        # Standarisasi RAK
        rak_code_col = df_rak.columns[0]
        rak_nama_col = df_rak.columns[1]
        rak_pagu_col = df_rak.columns[2] if len(df_rak.columns) > 2 else None
        
        df_rak['Kode_Clean'] = df_rak[rak_code_col].astype(str).str.strip()
        
        # Rekap Realisasi per Kode Rekening
        rekap_realisasi = df_bm.groupby(['Kode_Clean', 'Nama Rekening']).agg(
            Total_Realisasi=('Nilai Realisasi', 'sum'),
            Jumlah_Transaksi=('Nilai Realisasi', 'count')
        ).reset_index()
        
        # Merge RAK dan Realisasi untuk Pembanding
        if rak_pagu_col and pd.api.types.is_numeric_dtype(df_rak[rak_pagu_col]):
            df_compare = pd.merge(df_rak[['Kode_Clean', rak_nama_col, rak_pagu_col]], 
                                  rekap_realisasi[['Kode_Clean', 'Total_Realisasi', 'Jumlah_Transaksi']], 
                                  on='Kode_Clean', how='outer')
            
            df_compare['Total_Realisasi'] = df_compare['Total_Realisasi'].fillna(0)
            df_compare['Anggaran_RAK'] = df_compare[rak_pagu_col].fillna(0)
            df_compare['Sisa_Anggaran'] = df_compare['Anggaran_RAK'] - df_compare['Total_Realisasi']
            df_compare['Persentase_Capaian (%)'] = (df_compare['Total_Realisasi'] / df_compare['Anggaran_RAK'].replace(0, 1)) * 100
        else:
            df_compare = pd.merge(rekap_realisasi, df_rak[['Kode_Clean']], on='Kode_Clean', how='left')
            df_compare['Anggaran_RAK'] = 0
            df_compare['Sisa_Anggaran'] = 0
            df_compare['Persentase_Capaian (%)'] = 0

        st.success(f"Ditemukan **{len(df_bm)}** transaksi realisasi Belanja Modal (5.2)")
        
        # Dashboard Overview
        st.markdown("---")
        m1, m2, m3, m4 = st.columns(4)
        total_pagu = df_compare['Anggaran_RAK'].sum()
        total_real = df_compare['Total_Realisasi'].sum()
        sisa_pagu = total_pagu - total_real
        persen_total = (total_real / total_pagu * 100) if total_pagu > 0 else 0
        
        m1.metric("Total Anggaran (RAK)", f"Rp {total_pagu:,.0f}")
        m2.metric("Total Realisasi", f"Rp {total_real:,.0f}")
        m3.metric("Sisa Anggaran", f"Rp {sisa_pagu:,.0f}")
        m4.metric("Persentase Realisasi", f"{persen_total:.2f}%")
        
        # Tabs Tampilan
        tab1, tab2, tab3 = st.tabs(["⚖️ Pembanding RAK vs Realisasi", "🏢 Realisasi per SKPD", "📑 Detail Transaksi"])
        
        with tab1:
            st.subheader("Tabel Pembanding Anggaran RAK vs Realisasi")
            st.dataframe(df_compare[['Kode_Clean', rak_nama_col, 'Anggaran_RAK', 'Total_Realisasi', 'Sisa_Anggaran', 'Persentase_Capaian (%)']], use_container_width=True)
            
        with tab2:
            st.subheader("Rekap Realisasi Belanja Modal per SKPD")
            group_skpd = df_bm.groupby(['Kode SKPD', 'Nama SKPD']).agg(
                Total_Realisasi=('Nilai Realisasi', 'sum'),
                Jumlah_Transaksi=('Nilai Realisasi', 'count')
            ).reset_index().sort_values(by='Total_Realisasi', ascending=False)
            st.dataframe(group_skpd, use_container_width=True)
            
        with tab3:
            st.subheader("Detail Transaksi LRA")
            st.dataframe(df_bm, use_container_width=True)

        # Download Result
        st.markdown("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_compare.to_excel(writer, index=False, sheet_name='Pembanding_RAK_vs_Realisasi')
            group_skpd.to_excel(writer, index=False, sheet_name='Rekap_SKPD')
            df_bm.to_excel(writer, index=False, sheet_name='Detail_Transaksi')
            
        st.download_button(
            label="📥 Download Hasil Perbandingan (Excel)",
            data=output.getvalue(),
            file_name="HASIL_PEMBANDING_BELANJA_MODAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
else:
    st.info("Silakan unggah kedua file Excel di sidebar kiri.")