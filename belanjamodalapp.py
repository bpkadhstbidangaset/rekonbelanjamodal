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

# Custom Styling Bersih & Informatif
st.markdown("""
<style>
    .main-header { font-size:1.8rem; font-weight:700; color:#1E293B; margin-bottom: 2px; }
    .sub-header { font-size:0.95rem; color:#64748B; margin-bottom:18px; }
    .action-box-red {
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 15px;
    }
    .action-box-blue {
        background-color: #EFF6FF;
        border-left: 5px solid #3B82F6;
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 15px;
    }
    .action-box-green {
        background-color: #F0FDF4;
        border-left: 5px solid #22C55E;
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏛️ Mesin Rekonsiliasi Belanja Modal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Pencocokan Presisi & Diagnosa Tindakan Selisih Transaksi</div>', unsafe_allow_html=True)
st.divider()

# --- SIDEBAR UPLOAD ---
st.sidebar.header("📁 Upload File Data")
file_rak = st.sidebar.file_uploader("1. RAK Rekening Belanja Modal (Acuan)", type=['xlsx', 'xls', 'csv'])
file_sipd = st.sidebar.file_uploader("2. Data Realisasi SIPD", type=['xlsx', 'xls', 'csv'])
file_skpd = st.sidebar.file_uploader("3. Data Entry SKPD / Rincian Aset", type=['xlsx', 'xls', 'csv'])

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

# Helper membaca Excel SKPD
def parse_skpd_data(file):
    df_raw = pd.read_excel(file, header=None) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
        if 'RINCIAN PENGADAAN' in row_str or 'PEMERINTAH' in row_str or 'JENIS BELANJA :' in row_str:
            continue
        if ('KODE' in row_str and 'URAIAN' in row_str) or ('PENGADAAN' in row_str and 'ASET' in row_str):
            header_idx = idx
            break
        elif any(k in row_str for k in ['JENIS BELANJA', 'REKENING', 'SEMUA']) and ':' not in row_str:
            header_idx = idx
            break

    df = pd.read_excel(file, skiprows=header_idx) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, skiprows=header_idx)
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    if len(df) > 0:
        first_row_vals = [str(v).strip() for v in df.iloc[0].values if pd.notna(v)]
        if all(v.isdigit() for v in first_row_vals if v):
            df = df.iloc[1:].reset_index(drop=True)

    mask_total = df.apply(lambda row: row.astype(str).str.upper().str.contains('JUMLAH TOTAL|GRAND TOTAL|SUBTOTAL|T O T A L|PENGURUS BARANG|NIP.').any(), axis=1)
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
        col_skpd_sipd = [c for c in df_sipd.columns if any(k in c.lower() for k in ['skpd', 'dinas', 'opd', 'unit', 'nama_skpd'])]
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Target Rekonsiliasi")
        
        if col_skpd_sipd:
            list_skpd = sorted(df_sipd[col_skpd_sipd[0]].dropna().unique().tolist())
            selected_skpd_target = st.sidebar.selectbox("Pilih SKPD Target:", options=list_skpd)
            df_sipd_filtered = df_sipd[df_sipd[col_skpd_sipd[0]] == selected_skpd_target].copy()
        else:
            df_sipd_filtered = df_sipd.copy()
            selected_skpd_target = "Semua Data SIPD"

        # 3. IDENTIFIKASI & HITUNG SKPD
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

        current_matched_rek = ""
        skpd_rek_list = []
        
        for idx, row in df_skpd.iterrows():
            first_val = str(row.iloc[0]).strip()
            first_cols_str = ' '.join([str(v) for v in row.iloc[:3].values if pd.notna(v)])
            matched = get_matched_code(first_cols_str)
            
            if matched and (first_val.startswith('5.2') or first_val.startswith('5.3') or 'BELANJA MODAL' in first_cols_str.upper()):
                if not first_val.startswith('5.0') and '5.1.01' not in first_val:
                    current_matched_rek = matched
                    skpd_rek_list.append(matched)
                    continue

            if current_matched_rek and pd.isna(row.iloc[0]) and any(pd.notna(v) for v in row.iloc[1:4]):
                skpd_rek_list.append(current_matched_rek)
            else:
                skpd_rek_list.append("")

        df_skpd['Kode Rekening (Acuan)'] = skpd_rek_list

        def get_skpd_nominal(row):
            for col in ['PENGADAAN', 'ASET', 'SEMUA', 'TOTAL', 'JUMLAH', 'NILAI']:
                for c in df_skpd.columns:
                    if c == col or (col in c and 'RINCIAN' not in c and 'KODE' not in c and 'URAIAN' not in c):
                        val = clean_currency(row[c])
                        if val > 0:
                            return val
            row_vals = [clean_currency(v) for v in row.values if isinstance(v, (int, float)) or (isinstance(v, str) and any(d.isdigit() for d in v))]
            return max(row_vals) if row_vals else 0.0

        df_skpd['Nominal SKPD'] = df_skpd.apply(get_skpd_nominal, axis=1)
        
        def is_main_leaf(row):
            first_val = str(row.iloc[0]).strip()
            return (first_val.startswith('5.2') or first_val.startswith('5.3')) and row['Kode Rekening (Acuan)'] != ""

        df_skpd_main_leaves = df_skpd[df_skpd.apply(is_main_leaf, axis=1)].copy()
        active_skpd_codes = set(df_skpd_main_leaves['Kode Rekening (Acuan)'].unique())

        # 4. IDENTIFIKASI DATA SIPD
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
        grp_skpd = df_skpd_main_leaves.groupby('Kode Rekening (Acuan)')['Nominal SKPD'].sum().rename('Entry SKPD')

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

        # Total Metrik Dashboard
        total_sipd_bm = df_rekon['Realisasi SIPD'].sum()
        total_skpd_bm = df_rekon['Entry SKPD'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        # TAMPILAN HEADER DASHBOARD
        st.success(f"✅ Rekonsiliasi selesai untuk: **{selected_skpd_target}**")

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Realisasi SIPD", format_rupiah(total_sipd_bm))
        c2.metric("Total Entry SKPD", format_rupiah(total_skpd_bm))
        c3.metric(
            "Selisih Bersih", 
            format_rupiah(total_selisih), 
            delta=f"{-total_selisih:,.2f}", 
            delta_color="inverse"
        )

        st.markdown("---")

        # TAMPILAN TAB UTAMA
        tab_rekon, tab_investigasi, tab1, tab2, tab3 = st.tabs([
            "⚖️ Rekonsiliasi & Analisis Selisih",
            "🔎 Diagnosa & Rincian Selisih",
            "🔍 Detail Realisasi SIPD", 
            "📋 Detail Entry SKPD", 
            "📁 Data SKPD Dieliminasi"
        ])

        # === TAB 1: REKONSILIASI ===
        with tab_rekon:
            st.subheader("📊 Tabel Komparasi Belanja Modal (SIPD vs SKPD)")
            
            f_col1, f_col2 = st.columns([2, 2])
            with f_col1:
                filter_status = st.multiselect(
                    "Filter Status:",
                    options=sorted(df_rekon['Status'].unique()),
                    default=sorted(df_rekon['Status'].unique()),
                    key="f_status_rekon"
                )
            with f_col2:
                search_rekon = st.text_input("🔍 Cari Kode / Uraian:", "", key="s_rekon")

            df_rekon_filtered = df_rekon[df_rekon['Status'].isin(filter_status)].copy()
            if search_rekon:
                df_rekon_filtered = df_rekon_filtered[
                    df_rekon_filtered['Kode Rekening'].str.contains(search_rekon, case=False, na=False) |
                    df_rekon_filtered['Uraian Rekening (RAK)'].str.contains(search_rekon, case=False, na=False)
                ]

            df_rekon_view = df_rekon_filtered.copy()
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

        # === TAB 2: DIAGNOSA & RINCIAN SELISIH (BAHASA LUGAS & ACTIONABLE) ===
        with tab_investigasi:
            st.subheader("🔎 Diagnosa & Langkah Tindakan Selisih")
            
            df_selisih_only = df_rekon[abs(df_rekon['Selisih']) > 1].copy()
            
            if len(df_selisih_only) > 0:
                opsi_akun = [f"{r['Kode Rekening']} - {r['Uraian Rekening (RAK)']} (Selisih: {format_rupiah(r['Selisih'])})" for _, r in df_selisih_only.iterrows()]
                selected_opsi = st.selectbox("🎯 Pilih Akun yang Ingin Didorong Solusinya:", options=opsi_akun)
                
                selected_code_target = selected_opsi.split(" - ")[0].strip()
                row_info = df_selisih_only[df_selisih_only['Kode Rekening'] == selected_code_target].iloc[0]

                # --- KOTAK KESIMPULAN & ACTION PLAN ---
                if row_info['Selisih'] < 0:
                    st.markdown(f"""
                    <div class="action-box-red">
                        <h4 style="margin:0 0 6px 0; color:#991B1B;">⚠️ KESIMPULAN: Realisasi SIPD Lebih Kecil Rp {format_rupiah(abs(row_info['Selisih']))[3:]}</h4>
                        <b>Penyebab:</b> SKPD mencatat pengadaan sebesar <b>{format_rupiah(row_info['Entry SKPD'])}</b>, namun dana yang cair di kas daerah (SIPD) baru sebesar <b>{format_rupiah(row_info['Realisasi SIPD'])}</b>.<br>
                        👉 <b>Langkah Tindakan SKPD:</b><br>
                        1. Cek apakah paket pekerjaan ini <b>berupa kontrak bertahap (termin)</b> yang SP2D pelunasannya belum terbit.<br>
                        2. Jika pekerjaan belum selesai/SP2D belum cair, pastikan pencatatan di aplikasi aset belanja modal disesuaikan dengan realisasi SP2D yang sudah terbit saja.
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="action-box-blue">
                        <h4 style="margin:0 0 6px 0; color:#1E40AF;">ℹ️ KESIMPULAN: Realisasi SIPD Lebih Besar Rp {format_rupiah(row_info['Selisih'])[3:]}</h4>
                        <b>Penyebab:</b> Ada SP2D yang sudah cair di kas daerah sebesar <b>{format_rupiah(row_info['Realisasi SIPD'])}</b>, namun baru diinput di SKPD sebesar <b>{format_rupiah(row_info['Entry SKPD'])}</b>.<br>
                        👉 <b>Langkah Tindakan SKPD:</b><br>
                        Operator/Pengurus Barang SKPD <b>wajib menginput sisa SP2D yang belum tercatat</b> ke dalam aplikasi belanja modal.
                    </div>
                    """, unsafe_allow_html=True)

                # Ambil data SIPD & SKPD untuk akun ini
                df_sipd_target = df_sipd_bm[df_sipd_bm['Kode Rekening (Acuan)'] == selected_code_target].copy()
                df_skpd_target = df_skpd[df_skpd['Kode Rekening (Acuan)'] == selected_code_target].copy()
                
                # Rincian pengadaan di bawah header akun
                df_skpd_rincian = df_skpd_target[~df_skpd_target.apply(is_main_leaf, axis=1)].copy()
                if len(df_skpd_rincian) == 0:
                    df_skpd_rincian = df_skpd_target.copy()

                # --- PENCOCOKAN NOMINAL ---
                sipd_nominals = df_sipd_target['Nominal Realisasi'].tolist()
                skpd_nominals = df_skpd_rincian['Nominal SKPD'].tolist()

                sipd_counts = pd.Series(sipd_nominals).value_counts().to_dict()
                skpd_counts = pd.Series(skpd_nominals).value_counts().to_dict()

                unmatched_skpd = []
                for _, r in df_skpd_rincian.iterrows():
                    nom = r['Nominal SKPD']
                    if nom > 0 and (nom not in sipd_counts or skpd_counts.get(nom, 0) > sipd_counts.get(nom, 0)):
                        unmatched_skpd.append(r)

                unmatched_sipd = []
                for _, r in df_sipd_target.iterrows():
                    nom = r['Nominal Realisasi']
                    if nom > 0 and (nom not in skpd_counts or sipd_counts.get(nom, 0) > skpd_counts.get(nom, 0)):
                        unmatched_sipd.append(r)

                df_unmatched_skpd = pd.DataFrame(unmatched_skpd) if unmatched_skpd else pd.DataFrame()
                df_unmatched_sipd = pd.DataFrame(unmatched_sipd) if unmatched_sipd else pd.DataFrame()

                tot_u_skpd = df_unmatched_skpd['Nominal SKPD'].sum() if len(df_unmatched_skpd) > 0 else 0.0
                tot_u_sipd = df_unmatched_sipd['Nominal Realisasi'].sum() if len(df_unmatched_sipd) > 0 else 0.0

                c_sel1, c_sel2 = st.columns(2)
                
                # SISI KIRI: ITEM HANYA DI SKPD
                with c_sel1:
                    st.markdown("##### 📁 Rincian Paket di SKPD (Belum/Beda di SIPD)")
                    st.caption(f"Total: **{len(df_unmatched_skpd)} Paket** | Nominal: **`{format_rupiah(tot_u_skpd)}`**")
                    
                    if len(df_unmatched_skpd) > 0:
                        cols_text = [c for c in df_unmatched_skpd.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal SKPD']]
                        
                        # Ambil kolom uraian yang memiliki teks terpanjang (deskripsi paket)
                        best_col = cols_text[0] if cols_text else df_unmatched_skpd.columns[1]
                        for c in cols_text:
                            if df_unmatched_skpd[c].astype(str).str.len().mean() > df_unmatched_skpd[best_col].astype(str).str.len().mean():
                                best_col = c

                        df_u_skpd_clean = pd.DataFrame({
                            'Nama Pekerjaan / Paket Pengadaan SKPD': df_unmatched_skpd[best_col].astype(str),
                            'Nilai Tercatat SKPD': df_unmatched_skpd['Nominal SKPD'].apply(format_rupiah)
                        })
                        st.dataframe(df_u_skpd_clean, use_container_width=True, hide_index=True)
                    else:
                        st.success("✅ Semua paket pengadaan di SKPD cocok dengan SP2D SIPD.")

                # SISI KANAN: ITEM HANYA DI SIPD
                with c_sel2:
                    st.markdown("##### 🏛️ Rincian SP2D SIPD (Belum Tercatat di SKPD)")
                    st.caption(f"Total: **{len(df_unmatched_sipd)} SP2D** | Nominal: **`{format_rupiah(tot_u_sipd)}`**")
                    
                    if len(df_unmatched_sipd) > 0:
                        col_tgl = [c for c in df_unmatched_sipd.columns if 'TANGGAL' in c.upper() or 'TGL' in c.upper()]
                        col_ket = [c for c in df_unmatched_sipd.columns if 'KETERANGAN' in c.upper() or 'URAIAN' in c.upper() or 'BUKTI' in c.upper()]
                        
                        tgl_series = df_unmatched_sipd[col_tgl[0]] if col_tgl else df_unmatched_sipd.iloc[:, 0]
                        ket_series = df_unmatched_sipd[col_ket[0]] if col_ket else df_unmatched_sipd.iloc[:, 1]
                        
                        df_u_sipd_clean = pd.DataFrame({
                            'Tanggal SP2D': tgl_series.astype(str),
                            'Keterangan Transaksi SP2D': ket_series.astype(str),
                            'Nilai Cair SP2D': df_unmatched_sipd['Nominal Realisasi'].apply(format_rupiah)
                        })
                        st.dataframe(df_u_sipd_clean, use_container_width=True, hide_index=True)
                    else:
                        st.success("✅ Tidak ada SP2D di SIPD yang tertinggal diinput.")

            else:
                st.markdown("""
                <div class="action-box-green">
                    <h4 style="margin:0; color:#15803D;">🎉 SEMUA AKUN BELANJA MODAL SUDAH BALANCE 100%!</h4>
                    Tidak ada perbedaan angka antara realisasi SP2D di SIPD dan pencatatan entry di SKPD.
                </div>
                """, unsafe_allow_html=True)

        # === TAB 3: DETAIL REALISASI SIPD ===
        with tab1:
            st.subheader(f"Transaksi Realisasi SIPD Terkait ({len(df_sipd_bm)} Baris)")
            
            f_sipd_col1, f_sipd_col2 = st.columns([2, 2])
            with f_sipd_col1:
                list_rek_sipd = sorted(df_sipd_bm['Kode Rekening (Acuan)'].unique().tolist())
                selected_rek_sipd = st.multiselect("Filter Kode Rekening (SIPD):", options=list_rek_sipd, key="f_rek_sipd")
            with f_sipd_col2:
                search_sipd_text = st.text_input("🔍 Cari Keterangan / No. Bukti SIPD:", "", key="s_sipd")

            df_sipd_display = df_sipd_bm.copy()
            if selected_rek_sipd:
                df_sipd_display = df_sipd_display[df_sipd_display['Kode Rekening (Acuan)'].isin(selected_rek_sipd)]
            if search_sipd_text:
                mask_s = df_sipd_display.apply(lambda r: r.astype(str).str.contains(search_sipd_text, case=False).any(), axis=1)
                df_sipd_display = df_sipd_display[mask_s]

            cols_sipd_clean = [c for c in df_sipd_display.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal Realisasi']]
            df_sipd_view = df_sipd_display[['Kode Rekening (Acuan)'] + cols_sipd_clean + ['Nominal Realisasi']].copy()
            df_sipd_view['Nominal Realisasi'] = df_sipd_view['Nominal Realisasi'].apply(format_rupiah)
            
            st.dataframe(df_sipd_view, use_container_width=True, hide_index=True)

        # === TAB 4: DETAIL ENTRY SKPD ===
        with tab2:
            st.subheader(f"Rincian Pengadaan SKPD Terkait ({len(df_skpd_main_leaves)} Akun Utama)")
            
            f_skpd_col1, f_skpd_col2 = st.columns([2, 2])
            with f_skpd_col1:
                list_rek_skpd = sorted(df_skpd['Kode Rekening (Acuan)'].replace('', np.nan).dropna().unique().tolist())
                selected_rek_skpd = st.multiselect("Filter Kode Rekening (SKPD):", options=list_rek_skpd, key="f_rek_skpd")
            with f_skpd_col2:
                search_skpd_text = st.text_input("🔍 Cari Uraian Pengadaan SKPD:", "", key="s_skpd")

            df_skpd_display = df_skpd[df_skpd['Kode Rekening (Acuan)'] != ""].copy()
            if selected_rek_skpd:
                df_skpd_display = df_skpd_display[df_skpd_display['Kode Rekening (Acuan)'].isin(selected_rek_skpd)]
            if search_skpd_text:
                mask_skpd = df_skpd_display.apply(lambda r: r.astype(str).str.contains(search_skpd_text, case=False).any(), axis=1)
                df_skpd_display = df_skpd_display[mask_skpd]

            cols_to_keep = [c for c in df_skpd_display.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal SKPD']]
            df_skpd_view = df_skpd_display[['Kode Rekening (Acuan)'] + cols_to_keep + ['Nominal SKPD']].copy()
            df_skpd_view['Nominal SKPD'] = df_skpd_view['Nominal SKPD'].apply(format_rupiah)
            df_skpd_view = df_skpd_view.rename(columns={'Nominal SKPD': 'Nilai Bersih (Rupiah)'})

            st.dataframe(df_skpd_view, use_container_width=True, hide_index=True)

        # === TAB 5: DATA DIELIMINASI ===
        with tab3:
            st.subheader("Baris SKPD di Luar Belanja Modal (Induk / Non-BM)")
            df_skpd_elim = df_skpd[df_skpd['Kode Rekening (Acuan)'] == ""].copy()
            cols_elim_clean = [c for c in df_skpd_elim.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)']]
            df_elim_display = df_skpd_elim[cols_elim_clean].copy()
            
            st.dataframe(df_elim_display, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.warning("👈 Unggah ketiga file Excel/CSV di menu sebelah kiri untuk memulai rekonsiliasi.")
