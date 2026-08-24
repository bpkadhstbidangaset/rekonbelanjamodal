import streamlit as st
import pandas as pd
import numpy as np
import re
import io

# Set Konfigurasi Halaman
st.set_page_config(
    page_title="Mesin Rekon Belanja Modal",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Khusus Bubble Card & Header Modern
st.markdown("""
<style>
    .app-topbar {
        background: #FFFFFF;
        padding: 16px 24px;
        border-radius: 16px;
        box-shadow: 0 2px 6px -1px rgba(0, 0, 0, 0.05);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 20px;
        border: 1px solid #E2E8F0;
    }
    .breadcrumb-title {
        color: #64748B;
        font-size: 0.85rem;
        font-weight: 500;
        margin-bottom: 2px;
    }
    .breadcrumb-active {
        color: #0F172A;
        font-weight: 800;
        font-size: 1.25rem;
    }

    .bubble-card {
        background: #FFFFFF;
        border-radius: 20px;
        padding: 20px 22px;
        box-shadow: 0 4px 12px -2px rgba(0, 0, 0, 0.05);
        border: 1px solid #E2E8F0;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .card-top-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
    }
    .card-title-label {
        font-size: 0.78rem;
        color: #64748B;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .icon-bubble {
        width: 36px;
        height: 36px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.15rem;
    }
    .icon-blue { background-color: #EFF6FF; }
    .icon-purple { background-color: #FAF5FF; }
    .icon-red { background-color: #FEF2F2; }
    .icon-green { background-color: #F0FDF4; }

    .card-val-rp {
        font-size: 0.95rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: -2px;
    }
    .card-val-number {
        font-size: 1.35rem;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 12px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    
    .pill-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 700;
        width: fit-content;
    }
    .pill-blue { background: #DBEAFE; color: #1E40AF; }
    .pill-purple { background: #F3E8FF; color: #6B21A8; }
    .pill-red { background: #FEE2E2; color: #991B1B; }
    .pill-green { background: #DCFCE7; color: #15803D; }

    .action-box-red {
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }
    .action-box-blue {
        background-color: #EFF6FF;
        border-left: 5px solid #3B82F6;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }
    .action-box-green {
        background-color: #F0FDF4;
        border-left: 5px solid #22C55E;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }

    .calc-box {
        background-color: #0F172A;
        color: #FFFFFF;
        padding: 12px 18px;
        border-radius: 10px;
        font-size: 0.9rem;
        margin-top: 10px;
        display: flex;
        gap: 20px;
        align-items: center;
    }
</style>
""", unsafe_allow_html=True)

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

# Helper ekstraksi tanggal dari teks kontrak SIMBADA (jika ada)
def extract_date_from_text(text):
    if not isinstance(text, str):
        return "-"
    # Pola: tanggal 23 April 2026 atau 23/04/2026 atau 2026-04-23
    match_tgl = re.search(r'(\d{1,2}\s+(?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember)\s+\d{4})', text, re.IGNORECASE)
    if match_tgl:
        return match_tgl.group(1)
    match_num = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{4})', text)
    if match_num:
        return match_num.group(1)
    return "-"

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

# Helper membaca Excel SIMBADA
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

# Helper membaca Excel LRA (Mendukung Multi-Sheet SEMESTER 1 & SEMESTER 2)
def parse_lra_multi_sheet(file, selected_semester):
    if file.name.endswith(('.xlsx', '.xls')):
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        has_sem_sheets = any('SEMESTER' in s.upper() for s in sheet_names)
        
        sheets_to_read = []
        if has_sem_sheets:
            if selected_semester == "Semester 1":
                sheets_to_read = [s for s in sheet_names if '1' in s or 'I' in s]
            elif selected_semester == "Semester 2":
                sheets_to_read = [s for s in sheet_names if '2' in s or 'II' in s]
            else:
                sheets_to_read = sheet_names
        else:
            sheets_to_read = [sheet_names[0]]

        df_list = []
        for s_name in sheets_to_read:
            df_raw = pd.read_excel(file, sheet_name=s_name, header=None)
            header_idx = 0
            for idx, row in df_raw.iterrows():
                row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
                if ('NAMA SKPD' in row_str and 'KODE REKENING' in row_str) or ('NILAI REALISASI' in row_str):
                    header_idx = idx
                    break
                elif ('DEBIT' in row_str or 'SKPD' in row_str or 'URAIAN' in row_str) and 'REKAPITULASI' not in row_str:
                    header_idx = idx
                    break
            
            df_sheet = pd.read_excel(file, sheet_name=s_name, skiprows=header_idx)
            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
            df_sheet['__SHEET_SOURCE__'] = s_name
            df_list.append(df_sheet)

        df_combined = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    else:
        df_raw = pd.read_csv(file, header=None)
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(v).upper() for v in row.values if pd.notna(v)])
            if ('DEBIT' in row_str or 'SKPD' in row_str or 'URAIAN' in row_str) and 'REKAPITULASI' not in row_str:
                header_idx = idx
                break
        df_combined = pd.read_csv(file, skiprows=header_idx)
        df_combined.columns = [str(c).strip() for c in df_combined.columns]

    if selected_semester != "Semua Periode (Semester 1 & 2)":
        col_bulan = [c for c in df_combined.columns if c.lower() in ['bulan', 'periode spd']]
        if col_bulan:
            c_b = col_bulan[0]
            if selected_semester == "Semester 1":
                df_combined = df_combined[df_combined[c_b].astype(str).str.upper().str.contains('JAN|FEB|MAR|APR|MEI|JUN|1|I')].copy()
            elif selected_semester == "Semester 2":
                df_combined = df_combined[df_combined[c_b].astype(str).str.upper().str.contains('JUL|AGU|SEP|OKT|NOP|NOV|DES|2|II')].copy()

    return df_combined

# --- SIDEBAR MENU & UPLOAD ---
with st.sidebar:
    st.markdown("### 🏛️ **Mesin Rekon Belanja Modal**")
    st.caption("Aplikasi Rekonsiliasi Aset & Realisasi Kasda")
    st.markdown("---")
    
    st.markdown("##### 📁 **Upload Dokumen Sumber**")
    file_rak = st.file_uploader("1. Acuan RAK Belanja Modal", type=['xlsx', 'xls', 'csv'])
    file_sipd = st.file_uploader("2. Data Realisasi LRA / Kasda", type=['xlsx', 'xls', 'csv'])
    file_skpd = st.file_uploader("3. Data Entry SIMBADA", type=['xlsx', 'xls', 'csv'])

    st.markdown("---")
    st.markdown("##### ⏱️ **Filter Periode Rekonsiliasi**")
    selected_semester = st.selectbox(
        "Pilih Periode Semester LRA:",
        options=["Semua Periode (Semester 1 & 2)", "Semester 1", "Semester 2"],
        index=0
    )

# --- PROSES UTAMA ---
if file_rak and file_sipd and file_skpd:
    try:
        df_rak = parse_rak_data(file_rak)
        df_sipd = parse_lra_multi_sheet(file_sipd, selected_semester)
        df_skpd = parse_skpd_data(file_skpd)

        # 1. MAPPING KODE & URAIAN RAK
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

        # 2. FILTER SKPD TARGET
        col_skpd_lra = [c for c in df_sipd.columns if c.strip().lower() in ['nama skpd', 'skpd', 'nama_skpd', 'dinas', 'opd']]
        if not col_skpd_lra:
            col_skpd_lra = [c for c in df_sipd.columns if any(k in c.lower() for k in ['skpd', 'dinas', 'opd'])]
        
        with st.sidebar:
            st.markdown("---")
            st.markdown("##### 🎯 **Filter SKPD Target**")
            if col_skpd_lra:
                col_skpd_active = col_skpd_lra[0]
                list_skpd = sorted(df_sipd[col_skpd_active].dropna().unique().tolist())
                selected_skpd_target = st.selectbox("Pilih Satuan Kerja / SKPD:", options=list_skpd)
                df_sipd_filtered = df_sipd[df_sipd[col_skpd_active] == selected_skpd_target].copy()
            else:
                df_sipd_filtered = df_sipd.copy()
                selected_skpd_target = "Semua SKPD"

        # 3. IDENTIFIKASI DATA SIMBADA
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

        # 4. IDENTIFIKASI DATA REALISASI LRA
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

        col_nom_lra = [c for c in df_sipd_bm.columns if c.strip().lower() in ['nilai realisasi', 'nilai_realisasi', 'nilai sp2d', 'debit']]
        if col_nom_lra:
            sipd_target_col = col_nom_lra[0]
        else:
            col_debit = [c for c in df_sipd_bm.columns if 'debit' in c.lower() or 'realisasi' in c.lower()]
            sipd_target_col = col_debit[0] if col_debit else df_sipd_bm.columns[-4]

        df_sipd_bm['Nominal Realisasi'] = df_sipd_bm[sipd_target_col].apply(clean_currency)

        # 5. TABEL ANALISIS REKONSILIASI
        grp_sipd = df_sipd_bm.groupby('Kode Rekening (Acuan)')['Nominal Realisasi'].sum().rename('Realisasi LRA')
        grp_skpd = df_skpd_main_leaves.groupby('Kode Rekening (Acuan)')['Nominal SKPD'].sum().rename('Entry SIMBADA')

        df_rekon = pd.concat([grp_sipd, grp_skpd], axis=1).fillna(0)
        df_rekon = df_rekon[df_rekon['Entry SIMBADA'] > 0].copy()

        df_rekon['Selisih'] = df_rekon['Realisasi LRA'] - df_rekon['Entry SIMBADA']
        df_rekon['Status'] = df_rekon['Selisih'].apply(
            lambda x: '✅ Balance' if abs(x) < 1 else ('🔻 Realisasi LRA Lebih Kecil' if x < 0 else '🔺 Realisasi LRA Lebih Besar')
        )
        df_rekon = df_rekon.reset_index().rename(columns={'Kode Rekening (Acuan)': 'Kode Rekening'})
        df_rekon['Uraian Rekening (RAK)'] = df_rekon['Kode Rekening'].apply(lambda x: rak_lookup.get(x, 'Belanja Modal'))

        df_rekon = df_rekon[['Kode Rekening', 'Uraian Rekening (RAK)', 'Realisasi LRA', 'Entry SIMBADA', 'Selisih', 'Status']]
        df_rekon = df_rekon.sort_values(by='Selisih', key=abs, ascending=False)

        # Total Metrik
        total_sipd_bm = df_rekon['Realisasi LRA'].sum()
        total_skpd_bm = df_rekon['Entry SIMBADA'].sum()
        total_selisih = total_sipd_bm - total_skpd_bm

        # === TOP BAR HEADER BUBBLE ===
        st.markdown(f"""
        <div class="app-topbar">
            <div>
                <div class="breadcrumb-title">🏠 Beranda / Penatausahaan Aset / Rekonsiliasi Belanja Modal ({selected_semester})</div>
                <div class="breadcrumb-active">{selected_skpd_target}</div>
            </div>
            <div>
                <span class="pill-badge {'pill-green' if abs(total_selisih) < 1 else 'pill-red'}" style="font-size:0.85rem; padding: 6px 14px;">
                    {'● STATUS: BALANCE (SESUAI)' if abs(total_selisih) < 1 else f'● SELISIH: {format_rupiah(total_selisih)}'}
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # === 4 BUBBLE STAT CARDS ===
        k1, k2, k3, k4 = st.columns(4)
        
        with k1:
            st.markdown(f"""
            <div class="bubble-card">
                <div class="card-top-row">
                    <span class="card-title-label">Realisasi LRA</span>
                    <div class="icon-bubble icon-blue">🏛️</div>
                </div>
                <div>
                    <div class="card-val-rp">Rp</div>
                    <div class="card-val-number">{f"{total_sipd_bm:,.2f}"[0:-3].replace(',', '.')}</div>
                </div>
                <span class="pill-badge pill-blue">Kasda: {selected_semester}</span>
            </div>
            """, unsafe_allow_html=True)

        with k2:
            st.markdown(f"""
            <div class="bubble-card">
                <div class="card-top-row">
                    <span class="card-title-label">Entry SIMBADA (Aset)</span>
                    <div class="icon-bubble icon-purple">📋</div>
                </div>
                <div>
                    <div class="card-val-rp">Rp</div>
                    <div class="card-val-number">{f"{total_skpd_bm:,.2f}"[0:-3].replace(',', '.')}</div>
                </div>
                <span class="pill-badge pill-purple">Total Terdata SKPD</span>
            </div>
            """, unsafe_allow_html=True)

        with k3:
            st.markdown(f"""
            <div class="bubble-card">
                <div class="card-top-row">
                    <span class="card-title-label">Selisih Rekonsiliasi</span>
                    <div class="icon-bubble {'icon-green' if abs(total_selisih) < 1 else 'icon-red'}">⚖️</div>
                </div>
                <div>
                    <div class="card-val-rp">Rp</div>
                    <div class="card-val-number">{f"{total_selisih:,.2f}"[0:-3].replace(',', '.')}</div>
                </div>
                <span class="pill-badge {'pill-green' if abs(total_selisih) < 1 else 'pill-red'}">
                    {'Nol (0) Selisih' if abs(total_selisih) < 1 else 'Perlu Penyesuaian'}
                </span>
            </div>
            """, unsafe_allow_html=True)

        with k4:
            persen_match = (1.0 - abs(total_selisih) / max(total_sipd_bm, total_skpd_bm, 1)) * 100
            st.markdown(f"""
            <div class="bubble-card">
                <div class="card-top-row">
                    <span class="card-title-label">Akurasi Pencocokan</span>
                    <div class="icon-bubble icon-green">📊</div>
                </div>
                <div>
                    <div class="card-val-number" style="margin-top: 14px;">{persen_match:.1f}%</div>
                </div>
                <span class="pill-badge pill-green">{len(df_rekon)} Rekening Dianalisis</span>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

        # === TAMPILAN TAB UTAMA ===
        tab_rekon, tab_investigasi, tab_grafik, tab1, tab2 = st.tabs([
            "⚖️ Tabel Rekonsiliasi",
            "🔎 Diagnosa & Rincian Selisih",
            "📊 Grafik Komparasi Rekening",
            "🏛️ Data Realisasi LRA", 
            "📋 Data SIMBADA"
        ])

        # === TAB 1: REKONSILIASI ===
        with tab_rekon:
            st.markdown(f"#### **Tabel Komparasi Belanja Modal ({selected_semester})**")
            
            f_col1, f_col2 = st.columns([2, 2])
            with f_col1:
                filter_status = st.multiselect(
                    "Status Rekonsiliasi:",
                    options=sorted(df_rekon['Status'].unique()),
                    default=sorted(df_rekon['Status'].unique()),
                    key="f_status_rekon"
                )
            with f_col2:
                search_rekon = st.text_input("🔍 Cari Kode atau Nama Rekening:", "", key="s_rekon")

            df_rekon_filtered = df_rekon[df_rekon['Status'].isin(filter_status)].copy()
            if search_rekon:
                df_rekon_filtered = df_rekon_filtered[
                    df_rekon_filtered['Kode Rekening'].str.contains(search_rekon, case=False, na=False) |
                    df_rekon_filtered['Uraian Rekening (RAK)'].str.contains(search_rekon, case=False, na=False)
                ]

            df_rekon_view = df_rekon_filtered.copy()
            df_rekon_view['Realisasi LRA'] = df_rekon_view['Realisasi LRA'].apply(format_rupiah)
            df_rekon_view['Entry SIMBADA'] = df_rekon_view['Entry SIMBADA'].apply(format_rupiah)
            df_rekon_view['Selisih'] = df_rekon_view['Selisih'].apply(format_rupiah)

            st.dataframe(df_rekon_view, use_container_width=True, hide_index=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rekon.to_excel(writer, index=False, sheet_name='Rekonsiliasi')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Unduh Rekapitulasi Excel (.xlsx)",
                data=excel_data,
                file_name=f"Rekonsiliasi_{selected_skpd_target}_{selected_semester}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # === TAB 2: DIAGNOSA DETAIL SELISIH ===
        with tab_investigasi:
            st.markdown(f"#### **Diagnosa Per Kode Rekening yang Berselisih ({selected_semester})**")
            
            df_selisih_only = df_rekon[abs(df_rekon['Selisih']) > 1].copy()
            
            if len(df_selisih_only) > 0:
                opsi_akun = [f"{r['Kode Rekening']} - {r['Uraian Rekening (RAK)']} (Selisih: {format_rupiah(r['Selisih'])})" for _, r in df_selisih_only.iterrows()]
                selected_opsi = st.selectbox("Pilih Rekening yang Ingin Dilihat Rincian Selisihnya:", options=opsi_akun)
                
                selected_code_target = selected_opsi.split(" - ")[0].strip()
                row_info = df_selisih_only[df_selisih_only['Kode Rekening'] == selected_code_target].iloc[0]

                # Kotak Diagnosa
                if row_info['Selisih'] < 0:
                    st.markdown(f"""
                    <div class="action-box-red">
                        <h4 style="margin:0 0 6px 0; color:#DC2626;">📌 {row_info['Kode Rekening']} — {row_info['Uraian Rekening (RAK)']}</h4>
                        <div style="font-size:1.05rem; font-weight:700; color:#0F172A; margin-bottom:8px;">
                            Realisasi LRA Lebih Kecil Rp {format_rupiah(abs(row_info['Selisih']))[3:]}
                        </div>
                        <p style="color:#475569; margin:0 0 8px 0; font-size:0.9rem;">
                            <b>Analisis:</b> Total entry pada <b>SIMBADA</b> tercatat <b>{format_rupiah(row_info['Entry SIMBADA'])}</b>, sedangkan SP2D yang cair di <b>LRA Kasda ({selected_semester})</b> baru tercatat sebesar <b>{format_rupiah(row_info['Realisasi LRA'])}</b>.
                        </p>
                        <div style="background:#FEE2E2; padding:8px 12px; border-radius:8px; font-size:0.85rem; color:#991B1B; font-weight:600;">
                            👉 Catatan: Paket SIMBADA di bawah kemungkinan merupakan pekerjaan termin atau SP2D-nya baru cair pada semester berikutnya.
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="action-box-blue">
                        <h4 style="margin:0 0 6px 0; color:#2563EB;">📌 {row_info['Kode Rekening']} — {row_info['Uraian Rekening (RAK)']}</h4>
                        <div style="font-size:1.05rem; font-weight:700; color:#0F172A; margin-bottom:8px;">
                            Realisasi LRA Lebih Besar Rp {format_rupiah(row_info['Selisih'])[3:]}
                        </div>
                        <p style="color:#475569; margin:0 0 8px 0; font-size:0.9rem;">
                            <b>Analisis:</b> SP2D yang cair di <b>LRA Kasda ({selected_semester})</b> sebesar <b>{format_rupiah(row_info['Realisasi LRA'])}</b>, namun baru tercatat di <b>SIMBADA</b> sebesar <b>{format_rupiah(row_info['Entry SIMBADA'])}</b>.
                        </p>
                        <div style="background:#DBEAFE; padding:8px 12px; border-radius:8px; font-size:0.85rem; color:#1E40AF; font-weight:600;">
                            👉 Langkah: Pengurus Barang SKPD wajib menginput SP2D yang belum tercatat ke aplikasi SIMBADA.
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                # Ambil data SIPD & SKPD untuk akun ini
                df_sipd_target = df_sipd_bm[df_sipd_bm['Kode Rekening (Acuan)'] == selected_code_target].copy()
                df_skpd_target = df_skpd[df_skpd['Kode Rekening (Acuan)'] == selected_code_target].copy()
                
                df_skpd_rincian = df_skpd_target[~df_skpd_target.apply(is_main_leaf, axis=1)].copy()
                if len(df_skpd_rincian) == 0:
                    df_skpd_rincian = df_skpd_target.copy()

                # Pencocokan Nominal
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

                st.caption("💡 **Info**: Tanggal pada tabel SIMBADA diekstrak otomatis dari teks kontrak/BAP. Gunakan fitur centang baris untuk menghitung akumulasi nominal.")

                c_sel1, c_sel2 = st.columns(2)
                
                # SISI KIRI: DATA ENTRY SIMBADA
                with c_sel1:
                    st.markdown(f"""
                    <div style="font-weight:700; color:#0F172A; margin-bottom:4px;">
                        📁 Data Entry SIMBADA ({len(df_unmatched_skpd)} Paket)
                    </div>
                    <div style="font-size:0.85rem; color:#64748B; margin-bottom:8px;">
                        Total Nilai: <b>{format_rupiah(tot_u_skpd)}</b>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if len(df_unmatched_skpd) > 0:
                        def format_simbada_desc(row):
                            text_parts = []
                            for c in df_unmatched_skpd.columns:
                                if c not in ['Kode Rekening (Acuan)', 'Nominal SKPD', 'Is_Acuan_RAK', '__SHEET_SOURCE__']:
                                    v = row[c]
                                    if pd.notna(v) and str(v).strip() not in ['', 'nan', 'None', 'NaN']:
                                        if not (isinstance(v, (int, float)) or (isinstance(v, str) and v.isdigit() and len(v) > 6)):
                                            text_parts.append(str(v).strip())
                            return ' — '.join(text_parts) if text_parts else "Rincian Pengadaan SIMBADA"

                        full_desc_series = df_unmatched_skpd.apply(format_simbada_desc, axis=1)
                        tgl_simbada_series = full_desc_series.apply(extract_date_from_text)

                        df_u_skpd_clean = pd.DataFrame({
                            'Tgl Kontrak': tgl_simbada_series,
                            'Rincian Pengadaan SIMBADA': full_desc_series,
                            'Nilai Tercatat': df_unmatched_skpd['Nominal SKPD'].apply(format_rupiah),
                            'Nilai_Nominal_Angka': df_unmatched_skpd['Nominal SKPD'].values
                        })
                        
                        skpd_selection = st.dataframe(
                            df_u_skpd_clean[['Tgl Kontrak', 'Rincian Pengadaan SIMBADA', 'Nilai Tercatat']], 
                            use_container_width=True, 
                            hide_index=True,
                            on_select="rerun",
                            selection_mode="multi-row",
                            key="skpd_sel_box"
                        )
                        
                        sel_skpd_rows = skpd_selection.selection.rows
                        if sel_skpd_rows:
                            sel_sum_skpd = df_u_skpd_clean.iloc[sel_skpd_rows]['Nilai_Nominal_Angka'].sum()
                            st.markdown(f"""
                            <div class="calc-box">
                                <div>📌 <b>{len(sel_skpd_rows)} Baris Terpilih</b></div>
                                <div>∑ Total: <b>{format_rupiah(sel_sum_skpd)}</b></div>
                            </div>
                            """, unsafe_allow_html=True)

                        out_skpd = io.BytesIO()
                        with pd.ExcelWriter(out_skpd, engine='openpyxl') as writer:
                            df_export_skpd = df_u_skpd_clean[['Tgl Kontrak', 'Rincian Pengadaan SIMBADA', 'Nilai_Nominal_Angka']].rename(columns={'Nilai_Nominal_Angka': 'Nilai Tercatat (Rp)'})
                            df_export_skpd.to_excel(writer, index=False, sheet_name='SIMBADA_Belum_LRA')
                        
                        st.download_button(
                            label="📥 Unduh Rincian SIMBADA (Excel .xlsx)",
                            data=out_skpd.getvalue(),
                            file_name=f"SIMBADA_Belum_LRA_{selected_code_target}_{selected_semester}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_down_skpd"
                        )
                    else:
                        st.success("✅ Semua data SIMBADA sudah klop dengan LRA.")

                # SISI KANAN: DATA REALISASI LRA / SP2D
                with c_sel2:
                    st.markdown(f"""
                    <div style="font-weight:700; color:#0F172A; margin-bottom:4px;">
                        🏛️ Data Realisasi LRA / SP2D ({len(df_unmatched_sipd)} SP2D)
                    </div>
                    <div style="font-size:0.85rem; color:#64748B; margin-bottom:8px;">
                        Total Nilai: <b>{format_rupiah(tot_u_sipd)}</b>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if len(df_unmatched_sipd) > 0:
                        col_tgl = [c for c in df_unmatched_sipd.columns if any(k in c.lower() for k in ['tanggal sp2d', 'tanggal dokumen', 'tanggal'])]
                        col_ket = [c for c in df_unmatched_sipd.columns if any(k in c.lower() for k in ['keterangan dokumen', 'keterangan', 'uraian'])]
                        
                        tgl_series = df_unmatched_sipd[col_tgl[0]] if col_tgl else df_unmatched_sipd.iloc[:, 0]
                        ket_series = df_unmatched_sipd[col_ket[0]] if col_ket else df_unmatched_sipd.iloc[:, 1]
                        
                        df_u_sipd_clean = pd.DataFrame({
                            'Tanggal SP2D': tgl_series.astype(str),
                            'Keterangan Realisasi SP2D': ket_series.astype(str),
                            'Nilai SP2D': df_unmatched_sipd['Nominal Realisasi'].apply(format_rupiah),
                            'Nilai_Nominal_Angka': df_unmatched_sipd['Nominal Realisasi'].values
                        })
                        
                        sipd_selection = st.dataframe(
                            df_u_sipd_clean[['Tanggal SP2D', 'Keterangan Realisasi SP2D', 'Nilai SP2D']], 
                            use_container_width=True, 
                            hide_index=True,
                            on_select="rerun",
                            selection_mode="multi-row",
                            key="sipd_sel_box"
                        )

                        sel_sipd_rows = sipd_selection.selection.rows
                        if sel_sipd_rows:
                            sel_sum_sipd = df_u_sipd_clean.iloc[sel_sipd_rows]['Nilai_Nominal_Angka'].sum()
                            st.markdown(f"""
                            <div class="calc-box">
                                <div>📌 <b>{len(sel_sipd_rows)} Baris Terpilih</b></div>
                                <div>∑ Total: <b>{format_rupiah(sel_sum_sipd)}</b></div>
                            </div>
                            """, unsafe_allow_html=True)

                        out_sipd = io.BytesIO()
                        with pd.ExcelWriter(out_sipd, engine='openpyxl') as writer:
                            df_export_sipd = df_u_sipd_clean[['Tanggal SP2D', 'Keterangan Realisasi SP2D', 'Nilai_Nominal_Angka']].rename(columns={'Nilai_Nominal_Angka': 'Nilai SP2D (Rp)'})
                            df_export_sipd.to_excel(writer, index=False, sheet_name='LRA_Belum_SIMBADA')
                        
                        st.download_button(
                            label="📥 Unduh Rincian SP2D LRA (Excel .xlsx)",
                            data=out_sipd.getvalue(),
                            file_name=f"LRA_Belum_SIMBADA_{selected_code_target}_{selected_semester}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_down_sipd"
                        )
                    else:
                        st.success("✅ Semua realisasi LRA SP2D sudah tercatat di SIMBADA.")

            else:
                st.success("🎉 Semua Akun Belanja Modal Sudah Balance 100%!")

        # === TAB 3: GRAFIK KOMPARASI ===
        with tab_grafik:
            st.markdown(f"#### **Grafik Perbandingan Anggaran: Realisasi LRA vs SIMBADA ({selected_semester})**")
            df_chart = df_rekon[['Kode Rekening', 'Realisasi LRA', 'Entry SIMBADA']].copy()
            df_chart = df_chart.set_index('Kode Rekening')
            st.bar_chart(df_chart, height=380)

        # === TAB 4: DETAIL REALISASI LRA ===
        with tab1:
            st.markdown(f"#### **Data Mentah Realisasi LRA Kasda ({len(df_sipd_bm)} Baris - {selected_semester})**")
            cols_sipd_clean = [c for c in df_sipd_bm.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal Realisasi', '__SHEET_SOURCE__']]
            df_sipd_view = df_sipd_bm[['Kode Rekening (Acuan)'] + cols_sipd_clean + ['Nominal Realisasi']].copy()
            df_sipd_view['Nominal Realisasi'] = df_sipd_view['Nominal Realisasi'].apply(format_rupiah)
            st.dataframe(df_sipd_view, use_container_width=True, hide_index=True)

        # === TAB 5: DETAIL SIMBADA ===
        with tab2:
            st.markdown(f"#### **Data Mentah Entry SIMBADA ({len(df_skpd_main_leaves)} Rekening)**")
            cols_to_keep = [c for c in df_skpd.columns if not str(c).startswith('Unnamed') and c not in ['Kode Rekening (Acuan)', 'Nominal SKPD']]
            df_skpd_view = df_skpd[['Kode Rekening (Acuan)'] + cols_to_keep + ['Nominal SKPD']].copy()
            df_skpd_view['Nominal SKPD'] = df_skpd_view['Nominal SKPD'].apply(format_rupiah)
            st.dataframe(df_skpd_view, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan pemrosesan data: {e}")
else:
    st.info("👋 **Selamat Datang di Mesin Rekon Belanja Modal.** Silakan unggah ketiga file data di menu samping untuk memulai pencocokan.")
