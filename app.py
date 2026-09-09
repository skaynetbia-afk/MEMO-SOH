# =========================================================================
# BAGIAN PALING ATAS SKRIP (TEMPAT IMPORT LIBRARY)
# =========================================================================
import streamlit as st
import pandas as pd
import io  # 🟢 TAMBAHKAN BARIS INI DI PALING ATAS AGAR ANTI-ERROR NameError
from datetime import datetime

# 1. Konfigurasi Halaman Dasbor Sentral SNJ
st.set_page_config(page_title="Dasbor 11 Gudang Sewa", layout="wide")
st.title("❄️ Dasbor Stock on Hand (SOH) 11 Gudang Sewa")
st.subheader("PT. Suri Nusantara Jaya - Centralized Warehouse Dashboard")
st.markdown("---")

total_box_gabungan = 0
total_kg_gabungan = 0
jumlah_gudang_aktif = 0

# =========================================================================
# 2. PANEL MULTI-UPLOAD TERPISAH (GUDANG 1 SAMPAI 11)
# =========================================================================
st.subheader("📥 Area Unggah Berkas Laporan Gudang")
with st.expander("📂 Buka / Tutup Panel Unggah Berkas Excel SOH", expanded=True):
    kol_up1, kol_up2, kol_up3 = st.columns(3)
    
    with kol_up1:
        st.markdown("**Kelompok Gudang A:**")
        file_bosco = st.file_uploader("Unggah SOH Gudang Bosco", type=["xlsx", "xls", "csv"], key="up_bosco")
        file_einichi = st.file_uploader("Unggah SOH Gudang Einichi", type=["xlsx", "xls", "csv"], key="up_einichi")
        file_kiat = st.file_uploader("Unggah SOH Gudang Kiat (Gudang 3)", type=["xlsx", "xls", "csv"], key="up_kiat")
        file_mega = st.file_uploader("Unggah SOH Gudang Mega (Gudang 4)", type=["xlsx", "xls", "csv"], key="up_mega")
        
    with kol_up2:
        st.markdown("**Kelompok Gudang B:**")
        file_mtp = st.file_uploader("Unggah SOH Gudang MTP (Gudang 5)", type=["xlsx", "xls", "csv"], key="up_mtp")
        file_rpi = st.file_uploader("Unggah SOH Gudang RPI (Gudang 6)", type=["xlsx", "xls", "csv"], key="up_rpi")
        file_tcl = st.file_uploader("Unggah SOH Gudang TCL (Gudang 7)", type=["xlsx", "xls", "csv"], key="up_tcl")
        file_gudang8 = st.file_uploader("Unggah SOH Gudang Marina (Gudang 8)", type=["xlsx", "xls", "csv"], key="up_g8")
        
    with kol_up3:
        st.markdown("**Kelompok Gudang C:**")
        file_gudang9 = st.file_uploader("Unggah SOH Gudang PUM ( GUDANG 9)", type=["xlsx", "xls", "csv"], key="up_g9")
        file_gudang10 = st.file_uploader("Unggah SOH Gudang Kawanishi (Gudang 10)", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="up_g10")
        file_gudang11 = st.file_uploader("Unggah SOH Gudang 11 (DPR)", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="up_g11")

st.markdown("---")

# =========================================================================
# 3. PROSES DATA & PENYUSUNAN STRUKTUR MANDIRI PER GUDANG
# =========================================================================
gudang_data = {}

## --- PROSES GUDANG BOSCO (TAMBAH KOLOM VISUAL BARU EXPIRE DATE) ---
if file_bosco is not None:
    try:
        df_raw = pd.read_csv(file_bosco) if file_bosco.name.endswith('.csv') else pd.read_excel(file_bosco)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        nama_kolom_bosco = [
            'No', 'PID', 'Item #', 'No Polisi', 'Spec Product', 'Tanggal Masuk', 
            'No Container', 'Supplier Name', 'Brand', 'Comments', 'Expire Date',
            'Stock_Awal_Ctn', 'Stock_Awal_Kg', 'In_Ctn', 'In_Kg', 'Out_Ctn', 'Out_Kg', 
            'Stock_Akhir_Ctn', 'Stock_Akhir_Kg'
        ]
        if len(df_raw.columns) >= len(nama_kolom_bosco):
            df_raw = df_raw.iloc[:, :len(nama_kolom_bosco)]
            df_raw.columns = nama_kolom_bosco
        
        df_asli = df_raw[pd.to_numeric(df_raw['No'], errors='coerce').notnull()].copy()
        df_asli['Stock_Akhir_Ctn'] = pd.to_numeric(df_asli['Stock_Akhir_Ctn'], errors='coerce').fillna(0).astype(int)
        df_asli['Stock_Akhir_Kg'] = pd.to_numeric(df_asli['Stock_Akhir_Kg'], errors='coerce').fillna(0)
        
        # Format Tanggal Expire Date agar bersih dari jam bawaan Excel (diambil tanggalnya saja YYYY-MM-DD)
        df_asli['Expire_Clean'] = pd.to_datetime(df_asli['Expire Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli['Expire Date'].astype(str))
        
        df_memo = pd.DataFrame()
        
        # 1. Kolom BATCH dibersihkan total dari spasi internal maupun ujung
        batch_clean = df_asli['No Container'].astype(str).str.replace(r'\s+', '', regex=True)
        
        # 2. Kolom Produk dan Brand mempertahankan spasi kalimat naturalnya
        produk_normal = df_asli['Spec Product'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        brand_normal = df_asli['Brand'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        box_normal = df_asli['Stock_Akhir_Ctn'].astype(str).str.strip()
        
        # 3. Format MEMO standar rapat (BATCH//PRODUK-BRAND=BOXBOX)
        df_memo['MEMO'] = batch_clean + "//" + produk_normal + "-" + brand_normal + "=" + box_normal + "BOX"
        
        # 4. 🟢 PEMBENTUKAN STRUKTUR KOLOM BARU DI TABEL REKAPITULASI (Termasuk Expire Date)
        df_memo['BATCH'] = df_asli['No Container']
        df_memo['PRODUK'] = df_asli['Spec Product']
        df_memo['BOX'] = df_asli['Stock_Akhir_Ctn']
        df_memo['KG'] = df_asli['Stock_Akhir_Kg']
        df_memo['EXPIRE DATE'] = df_asli['Expire_Clean'] # 🟢 Kolom Baru Resmi Ditambahkan di Sini
        
        gudang_data["GUDANG BOSCO"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Bosco: {e}")


# --- PROSES GUDANG EINICHI (REVISI MEMO: HANYA BATCH BERSIH SPASI & TAMBAH KOLOM EXPIRED DATE) ---
if file_einichi is not None:
    try:
        df_raw = pd.read_csv(file_einichi) if file_einichi.name.endswith('.csv') else pd.read_excel(file_einichi)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        # PENGAMAN KOLOM 'No.' / 'No': Menyesuaikan variasi titik pada header Excel
        kolom_no_einichi = 'No.' if 'No.' in df_raw.columns else ('No' if 'No' in df_raw.columns else df_raw.columns[0])
        df_asli = df_raw[df_raw[kolom_no_einichi].notnull()].copy()
        
        df_asli['Qty'] = pd.to_numeric(df_asli['Qty'], errors='coerce').fillna(0).astype(int)
        df_asli['Qty Kg'] = pd.to_numeric(df_asli['Qty Kg'], errors='coerce').fillna(0)
        
        kolom_do = 'Cust DO number' if 'Cust DO number' in df_asli.columns else 'CUST DO NUMBER'
        
        # PENGAMAN KOLOM EXPIRED DATE: Menyesuaikan variasi nama kolom expired di Excel Einichi
        kolom_exp_einichi = 'Expired Date' if 'Expired Date' in df_asli.columns else ([c for c in df_asli.columns if 'EXP' in c.upper() or 'DATE' in c.upper()]+['Expired Date'])[0]
        
        # Format Tanggal Expired agar bersih dari jam bawaan Excel (diambil tanggalnya saja YYYY-MM-DD)
        if kolom_exp_einichi in df_asli.columns:
            df_asli['Expired_Clean'] = pd.to_datetime(df_asli[kolom_exp_einichi], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_einichi].astype(str))
        else:
            df_asli['Expired_Clean'] = "-"

        df_memo = pd.DataFrame()
        
        # 1. Bersihkan total seluruh spasi internal khusus pada kolom BATCH saja
        batch_clean = df_asli[kolom_do].astype(str).str.replace(r'\s+', '', regex=True)
        
        # 2. Ambil produk, brand, dan box berspasi normal kalimat (hanya distrip ujungnya)
        produk_normal = df_asli['Commodity Name'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        brand_normal = df_asli['Brand'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        box_normal = df_asli['Qty'].astype(str).str.strip()
        
        # 3. Gabungkan menjadi string MEMO rapat di antara operator pemisah (BATCH//PRODUK-BRAND=BOXBOX)
        df_memo['MEMO'] = batch_clean + "//" + produk_normal + "-" + brand_normal + "=" + box_normal + "BOX"
        
        # 4. PEMBENTUKAN STRUKTUR KOLOM VISUAL DI TABEL REKAPITULASI (Termasuk Expired Date)
        df_memo['BATCH'] = df_asli[kolom_do]
        df_memo['PRODUK'] = df_asli['Commodity Name']
        df_memo['BOX'] = df_asli['Qty']
        df_memo['KG'] = df_asli['Qty Kg']
        df_memo['EXPIRED DATE'] = df_asli['Expired_Clean'] # 🟢 Kolom Baru Resmi Ditambahkan di Sini
        
        gudang_data["GUDANG EINICHI"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Einichi: {e}")


# =========================================================================
# --- PROSES GUDANG KIAT (GUDANG 3) ---
# =========================================================================
if file_kiat is not None:
    try:
        df_asli = pd.read_csv(file_kiat) if file_kiat.name.endswith('.csv') else pd.read_excel(file_kiat)
        df_asli.columns = [str(c).strip().lstrip('_').upper() for c in df_asli.columns]
        
        df_asli['SOH'] = pd.to_numeric(df_asli['SOH'], errors='coerce').fillna(0).astype(int)
        df_asli['SOH_KG'] = pd.to_numeric(df_asli['SOH_KG'], errors='coerce').fillna(0)
        
        kolom_exp_kiat = 'EXP_DATE' if 'EXP_DATE' in df_asli.columns else 'EXP DATE'
        kolom_usage_kiat = 'DESC_USAGE' if 'DESC_USAGE' in df_asli.columns else 'DESC USAGE'
        
        if kolom_exp_kiat in df_asli.columns:
            df_asli['EXP_DATE_CLEAN'] = pd.to_datetime(df_asli[kolom_exp_kiat], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_kiat].astype(str))
        else:
            df_asli['EXP_DATE_CLEAN'] = "-"
            
        if kolom_usage_kiat not in df_asli.columns:
            df_asli[kolom_usage_kiat] = "-"

        # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
        batch_clean = df_asli['BATCH'].astype(str).str.replace(r'\s+', '', regex=True)
        produk_normal = df_asli['DESCRIPTION'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.replace(r'\s*-\s*PT\s*\.?\s*SURI\s*NUSANTARA\s*JAYA', '', regex=True, case=False)
        box_normal = df_asli['SOH'].astype(str).str.strip()
        
        df_memo = pd.DataFrame()
        df_memo['MEMO'] = batch_clean + "//" + produk_normal + "=" + box_normal + "BOX"
        df_memo['BATCH'] = df_asli['BATCH']
        df_memo['PRODUK'] = df_asli['DESCRIPTION']
        df_memo['BOX'] = df_asli['SOH']
        df_memo['KG'] = df_asli['SOH_KG']
        df_memo['EXP_DATE'] = df_asli['EXP_DATE_CLEAN']
        df_memo['DESC_USAGE'] = df_asli[kolom_usage_kiat]
        
        gudang_data["GUDANG KIAT"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Gudang Kiat: {e}")


# =========================================================================
# --- PROSES GUDANG MEGA (GUDANG 4) ---
# =========================================================================
if file_mega is not None:
    try:
        df_asli = pd.read_csv(file_mega) if file_mega.name.endswith('.csv') else pd.read_excel(file_mega)
        df_asli.columns = [str(c).strip() for c in df_asli.columns]
        
        kolom_desc = 'descriptionname' if 'descriptionname' in df_asli.columns else 'Deskripsi SNJ'
        kolom_lot = 'lotno' if 'lotno' in df_asli.columns else 'Kode SNJ'
        kolom_sub_gedung = 'gedung' if 'gedung' in df_asli.columns else 'gedung'
        
        kolom_exp_mega = 'expired' if 'expired' in df_asli.columns else (
            'Expired Date' if 'Expired Date' in df_asli.columns else (
                [c for c in df_asli.columns if 'EXP' in c.upper() or 'DATE' in c.upper()]+['expired']
            )[0]
        )
        
        df_asli[kolom_desc] = df_asli[kolom_desc].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_asli[kolom_lot] = df_asli[kolom_lot].astype(str).str.strip()
        df_asli[kolom_sub_gedung] = df_asli[kolom_sub_gedung].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.upper()
        
        df_asli['qty'] = pd.to_numeric(df_asli['qty'], errors='coerce').fillna(0).astype(int)
        df_asli['weight'] = pd.to_numeric(df_asli['weight'], errors='coerce').fillna(0)
        
        if kolom_exp_mega in df_asli.columns:
            df_asli['Expired_Clean'] = pd.to_datetime(df_asli[kolom_exp_mega], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_mega].astype(str))
        else:
            df_asli['Expired_Clean'] = "-"
            
        # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
        batch_clean = df_asli[kolom_lot].astype(str).str.replace(r'\s+', '', regex=True)
        produk_normal = df_asli[kolom_desc].astype(str).str.strip()
        sub_gedung_normal = df_asli[kolom_sub_gedung].astype(str).str.strip()
        box_normal = df_asli['qty'].astype(str).str.strip()
        
        df_memo = pd.DataFrame()
        # Formasi Rapat Operator, Teks Bawaan Tetap Alami Berspasi
        df_memo['MEMO'] = batch_clean + "//" + produk_normal + "(" + sub_gedung_normal + ")" + "=" + box_normal + "BOX"
        df_memo['BATCH'] = df_asli[kolom_lot]
        df_memo['PRODUK'] = df_asli[kolom_desc]
        df_memo['BOX'] = df_asli['qty']
        df_memo['KG'] = df_asli['weight']
        df_memo['SUB_GEDUNG'] = df_asli[kolom_sub_gedung]
        df_memo['EXPIRED'] = df_asli['Expired_Clean']
        
        gudang_data["GUDANG MEGA"] = {"asli": df_asli, "rekap": df_memo, "kolom_sub": kolom_sub_gedung}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Gudang Mega: {e}")


# =========================================================================
# --- PROSES GUDANG MTP (GUDANG 5) ---
# =========================================================================
if file_mtp is not None:
    try:
        df_raw = pd.read_csv(file_mtp) if file_mtp.name.endswith('.csv') else pd.read_excel(file_mtp)
        df_raw.columns = [str(c).strip().upper() for c in df_raw.columns]
        
        df_asli = df_raw.copy()
        df_asli['AVAILABLE QTY'] = pd.to_numeric(df_asli['AVAILABLE QTY'], errors='coerce').fillna(0).astype(int)
        
        kolom_kg_mtp = 'STOCK KG' if 'STOCK KG' in df_asli.columns else ('STOCK_KG' if 'STOCK_KG' in df_asli.columns else df_asli.columns)
        df_asli['STOCK_KG_CLEAN'] = pd.to_numeric(df_asli[kolom_kg_mtp], errors='coerce').fillna(0)
        
        kolom_sku = 'SKU' if 'SKU' in df_asli.columns else 'sku'
        kolom_exp_mtp = 'EXP_DATE' if 'EXP_DATE' in df_asli.columns else ('EXP DATE' if 'EXP DATE' in df_asli.columns else 'EXPIRED')
        
        if kolom_exp_mtp in df_asli.columns:
            df_asli['EXP_DATE_CLEAN'] = pd.to_datetime(df_asli[kolom_exp_mtp], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_mtp].astype(str))
        else:
            df_asli['EXP_DATE_CLEAN'] = "-"
            
        item_wajib_sku = ["FLANK", "FOREQUARTER", "FQ ROLL", "SHIN SHANK", "TRIMMING 90 VL"]
        list_produk_final = []
        for index, row in df_asli.iterrows():
            desc_raw = str(row['DESCRIPTION']).strip()
            sku_raw = str(row[kolom_sku]).strip()
            cocok = any(item.upper() in desc_raw.upper() for item in item_wajib_sku)
            if cocok: 
                list_produk_final.append(f"{sku_raw} - {desc_raw}")
            else: 
                list_produk_final.append(desc_raw)
                
        df_asli['PRODUK_ADAPTIF'] = list_produk_final
        
        df_memo = pd.DataFrame()
        # Logika Spasi Memo MTP: Sesuai Instruksi, Tetap Menjaga Spasi Renggang
        df_memo['MEMO'] = (
            df_asli['BATCH'].astype(str).str.strip() + " // " + 
            df_asli['PRODUK_ADAPTIF'].astype(str) + " - " + 
            df_asli['LOTTABLE03'].astype(str) + " - " + 
            df_asli['LOTTABLE02'].astype(str) + " = " + 
            df_asli['AVAILABLE QTY'].astype(str) + " BOX"
        )
        df_memo['BATCH'] = df_asli['BATCH']
        df_memo['PRODUK'] = df_asli['PRODUK_ADAPTIF']
        df_memo['BOX'] = df_asli['AVAILABLE QTY']
        df_memo['KG'] = df_asli['STOCK_KG_CLEAN']
        df_memo['EXP DATE'] = df_asli['EXP_DATE_CLEAN']
        
        gudang_data["GUDANG MTP"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Gudang MTP: {e}")
# =========================================================================
# --- PROSES GUDANG RPI (GUDANG 6) ---
# =========================================================================
if file_rpi is not None:
    try:
        df_raw = pd.read_csv(file_rpi) if file_rpi.name.endswith('.csv') else pd.read_excel(file_rpi)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        # Saring baris total bawaan Excel RPI
        df_clean = df_raw[df_raw['No.'].notnull() & df_raw['Location'].notnull()].copy()
        
        kolom_wajib_rpi = [
            'No.', 'Area', 'Location', 'Item #', 'Description', 'Lot #',
            'Production #', 'Exp. Date', 'PID', 'Pallet ID', 'Stock Qty (Box)',
            'Stock Qty (Kg)', 'Storage Date'
        ]
        kolom_tersedia = [c for c in kolom_wajib_rpi if c in df_clean.columns]
        df_asli = df_clean[kolom_tersedia].copy()
        
        df_asli['Stock_Qty_Box_Clean'] = pd.to_numeric(df_asli['Stock Qty (Box)'], errors='coerce').fillna(0).astype(int)
        df_asli['Stock_Qty_Kg_Clean'] = pd.to_numeric(df_asli['Stock Qty (Kg)'], errors='coerce').fillna(0)
        
        # 🟢 PERBAIKAN: Cari 1 nama kolom expired terbaik secara aman (mengembalikan string tunggal, bukan list)
        kolom_exp_rpi = 'Exp. Date'
        for c in df_asli.columns:
            if 'EXP' in str(c).upper() or 'DATE' in str(c).upper():
                kolom_exp_rpi = c
                break
        
        if kolom_exp_rpi in df_asli.columns:
            df_asli['Exp_Clean_Visual'] = pd.to_datetime(df_asli[kolom_exp_rpi], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_rpi].astype(str))
        else:
            df_asli['Exp_Clean_Visual'] = "-"
            
        df_memo = pd.DataFrame()
        # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
        batch_clean = df_asli['Production #'].astype(str).str.replace(r'\s+', '', regex=True)
        desc_normal = df_asli['Description'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        lot_normal = df_asli['Lot #'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        box_normal = df_asli['Stock_Qty_Box_Clean'].astype(str).str.strip()
        
        df_memo['MEMO'] = batch_clean + "//" + desc_normal + "-" + lot_normal + "=" + box_normal + "BOX"
        df_memo['BATCH'] = df_asli['Production #']
        df_memo['PRODUK'] = df_asli['Description']
        df_memo['BOX'] = df_asli['Stock_Qty_Box_Clean']
        df_memo['KG'] = df_asli['Stock_Qty_Kg_Clean']
        df_memo['EXP. DATE'] = df_asli['Exp_Clean_Visual']
        
        gudang_data["GUDANG RPI"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Gudang RPI: {e}")


# =========================================================================
# --- PROSES GUDANG TCL (GUDANG 7) ---
# =========================================================================
if file_tcl is not None:
    try:
        df_raw = pd.read_excel(file_tcl, sheet_name=0) if file_tcl.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_tcl)
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        
        df_asli = df_raw[df_raw['Pallet ID'].notnull()].copy()
        kolom_box_tcl = 'CAR / PLT' if 'CAR / PLT' in df_asli.columns else ('CAR/PLT' if 'CAR/PLT' in df_asli.columns else 'Total Qty')
        
        kolom_exp_tcl = 'Exp Date'
        for c in df_asli.columns:
            if 'EXP' in str(c).upper() or 'DATE' in str(c).upper():
                kolom_exp_tcl = c
                break
                
        df_asli['Box_Clean'] = pd.to_numeric(df_asli[kolom_box_tcl], errors='coerce').fillna(0).astype(int)
        df_asli['Gross Wt'] = pd.to_numeric(df_asli['Gross Wt'], errors='coerce').fillna(0)
        
        df_asli['Description'] = df_asli['Description'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_asli['LotNo'] = df_asli['LotNo'].astype(str).str.strip()
        
        if kolom_exp_tcl in df_asli.columns:
            df_asli['Exp_Clean'] = pd.to_datetime(df_asli[kolom_exp_tcl], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_tcl].astype(str))
        else:
            df_asli['Exp_Clean'] = "-"
            
        df_memo = pd.DataFrame()
        # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
        batch_clean = df_asli['LotNo'].astype(str).str.replace(r'\s+', '', regex=True)
        produk_normal = df_asli['Description'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        box_normal = df_asli['Box_Clean'].astype(str).str.strip()
        
        df_memo['MEMO'] = batch_clean + "//" + produk_normal + "=" + box_normal + "BOX"
        df_memo['BATCH'] = df_asli['LotNo']
        df_memo['PRODUK'] = df_asli['Description']
        df_memo['BOX'] = df_asli['Box_Clean']
        df_memo['KG'] = df_asli['Gross Wt']
        df_memo['EXP. DATE'] = df_asli['Exp_Clean']
        
        gudang_data["GUDANG TCL"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Gudang TCL: {e}")
# =========================================================================
# --- PROSES GUDANG MARINA (GUDANG 8) ---
# =========================================================================
if file_gudang8 is not None:
    try:
        df_raw = pd.read_csv(file_gudang8) if file_gudang8.name.endswith('.csv') else pd.read_excel(file_gudang8)
        df_raw.columns = [str(c).strip().lower() for c in df_raw.columns]
        
        # Pengaman: Buang baris ringkasan Grand Total sekunder milik Excel
        df_clean = df_raw[df_raw['lotno'].notnull() & df_raw['qty'].notnull()].copy()
        df_clean = df_clean[df_clean['lotno'] != 'lotno']
        
        kolom_exp_marina = 'expired'
        for c in df_clean.columns:
            if 'exp' in str(c) or 'date' in str(c) or 'expiry' in str(c):
                kolom_exp_marina = c
                break
        
        df_clean['qty'] = pd.to_numeric(df_clean['qty'], errors='coerce').fillna(0).astype(int)
        df_clean['weight'] = pd.to_numeric(df_clean['weight'], errors='coerce').fillna(0)
        df_clean['item description accurate'] = df_clean['item description accurate'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_clean['lotno'] = df_clean['lotno'].astype(str).str.strip()
        df_clean['prodnm'] = df_clean['prodnm'].astype(str).str.strip()
        
        if kolom_exp_marina in df_clean.columns:
            df_clean['exp_clean_date'] = pd.to_datetime(df_clean[kolom_exp_marina], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_clean[kolom_exp_marina].astype(str))
        else:
            df_clean['exp_clean_date'] = "-"
            
        # Smart Detection kelompok PPI vs Berdikari
        kondisi_ppi = (df_clean['prodnm'].str.contains('PPI', case=False, na=False)) | (df_clean['item description accurate'].str.contains('PPI', case=False, na=False))
        df_ppi = df_clean[kondisi_ppi].copy()
        df_berdikari = df_clean[~kondisi_ppi].copy()
        
        df_berdikari['SUMBER_SHEET'] = "MASTER SOH BERDIKARI"
        df_ppi['SUMBER_SHEET'] = "MASTER SOH PPI"
        df_asli_marina = pd.concat([df_berdikari, df_ppi], ignore_index=True)
        
        df_memo = pd.DataFrame()
        # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
        lot_clean = df_asli_marina['lotno'].astype(str).str.replace(r'\s+', '', regex=True)
        desc_normal = df_asli_marina['item description accurate'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        qty_normal = df_asli_marina['qty'].astype(str).str.strip()
        
        df_memo['MEMO'] = lot_clean + "//" + desc_normal + "=" + qty_normal + "BOX"
        df_memo['BATCH'] = df_asli_marina['lotno']
        df_memo['PRODUK'] = df_asli_marina['item description accurate']
        df_memo['BOX'] = df_asli_marina['qty']
        df_memo['KG'] = df_asli_marina['weight']
        df_memo['SUMBER'] = df_asli_marina['SUMBER_SHEET']
        df_memo['EXPIRED'] = df_asli_marina['exp_clean_date']
        
        gudang_data["GUDANG MARINA"] = {"asli": df_asli_marina, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas Gudang Marina: {e}")


# =========================================================================
# --- PROSES GUDANG KONSOLIDASI GUDANG PUM (GUDANG 9) ---
# =========================================================================
if file_gudang9 is not None:
    try:
        df_pum_berdikari = pd.read_excel(file_gudang9, sheet_name="Input Stock Berdikari", header=13)
        df_pum_ppi = pd.read_excel(file_gudang9, sheet_name="Input Stock PPI", header=13)
        
        list_pum_rekap = []
        list_pum_asli_clean = []
        
        for df_sheet, label_sumber in [(df_pum_berdikari, "BKR"), (df_pum_ppi, "PPI")]:
            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
            df_filter = df_sheet[pd.to_numeric(df_sheet['No'], errors='coerce').notnull()].copy()
            df_filter = df_filter[df_filter['Lot.No'].notnull() & (df_filter['Lot.No'].astype(str).str.strip() != '')]
            df_filter = df_filter[~df_filter['Spec Product'].astype(str).str.contains('Total|Grand', case=False, na=False)]
            
            if not df_filter.empty:
                df_std = pd.DataFrame()
                df_std['GUDANG'] = ["PUM"] * len(df_filter)
                df_std['PALLET_ID'] = df_filter['PID'].astype(str)
                df_std['PRODUK'] = df_filter['Spec Product'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
                df_std['BRAND'] = df_filter['Item #'].astype(str).str.strip()
                
                # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
                df_std['BATCH'] = df_filter['Lot.No'].astype(str).str.replace(r'\s+', '', regex=True)
                
                kolom_box_pum = 'MC.3' if 'MC.3' in df_filter.columns else df_filter.columns
                kolom_kg_pum = 'Kg.3' if 'Kg.3' in df_filter.columns else df_filter.columns
                
                df_std['BOX'] = pd.to_numeric(df_filter[kolom_box_pum], errors='coerce').fillna(0).astype(int)
                df_std['KG'] = pd.to_numeric(df_filter[kolom_kg_pum], errors='coerce').fillna(0)
                
                kolom_exp_pum = 'Expire Date'
                for c in df_filter.columns:
                    if 'EXP' in str(c).upper() or 'DATE' in str(c).upper():
                        kolom_exp_pum = c
                        break
                        
                if kolom_exp_pum in df_filter.columns:
                    df_std['Expired Date'] = pd.to_datetime(df_filter[kolom_exp_pum], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_filter[kolom_exp_pum].astype(str))
                else:
                    df_std['Expired Date'] = "-"
                    
                produk_clean_memo = df_std['PRODUK'].astype(str).str.replace(r'\s+', '', regex=True)
                box_clean_memo = df_std['BOX'].astype(str).str.strip()
                
                df_std['MEMO'] = df_std['BATCH'] + "//" + produk_clean_memo + "-" + label_sumber + "=" + box_clean_memo + "BOX"
                df_std['SUMBER_SHEET'] = label_sumber
                list_pum_rekap.append(df_std)
                
                df_tabel_asli_terpotong = df_filter.iloc[:, :18].copy()
                df_tabel_asli_terpotong.insert(0, 'SUMBER_SHEET', label_sumber)
                list_pum_asli_clean.append(df_tabel_asli_terpotong)
                
        if list_pum_rekap and list_pum_asli_clean:
            df_asli_pum_full = pd.concat(list_pum_rekap, ignore_index=True)
            df_excel_asli_gabungan = pd.concat(list_pum_asli_clean, ignore_index=True)
            
            gudang_data["GUDANG PUM"] = {
                "master_pum": df_asli_pum_full,
                "asli_gabungan": df_excel_asli_gabungan
            }
            total_box_gabungan += df_asli_pum_full['BOX'].sum()
            total_kg_gabungan += df_asli_pum_full['KG'].sum()
            jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses susunan berkas format baru Gudang PUM: {e}")


# =========================================================================
# --- PROSES GUDANG KAWANISHI (GUDANG 10) ---
# =========================================================================
if file_gudang10:
    try:
        list_df_kawanishi = []
        for file_kwn in file_gudang10:
            df_singel = pd.read_csv(file_kwn) if file_kwn.name.endswith('.csv') else pd.read_excel(file_kwn)
            df_singel.columns = [str(c).strip() for c in df_singel.columns]
            
            if 'Item Code' in df_singel.columns: 
                df_singel = df_singel.rename(columns={'Item Code': 'ITEM_CODE_UNIFIED'})
            elif 'Item' in df_singel.columns: 
                df_singel = df_singel.rename(columns={'Item': 'ITEM_CODE_UNIFIED'})
            list_df_kawanishi.append(df_singel)
            
        df_raw_kawanishi = pd.concat(list_df_kawanishi, ignore_index=True)
        df_asli = df_raw_kawanishi[df_raw_kawanishi['Item Name'].notnull() & df_raw_kawanishi['Lot No.'].notnull()].copy()
        df_asli = df_asli[df_asli['Item Name'] != 'Item Name']
        
        df_asli['End Stock'] = pd.to_numeric(df_asli['End Stock'], errors='coerce').fillna(0).astype(int)
        df_asli['Weight (Kg)'] = pd.to_numeric(df_asli['Weight (Kg)'], errors='coerce').fillna(0)
        df_asli['Item Name'] = df_asli['Item Name'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_asli['Lot No.'] = df_asli['Lot No.'].astype(str).str.strip()
        df_asli = df_asli.rename(columns={'ITEM_CODE_UNIFIED': 'Item Code'})
        
        kolom_exp_kwn = 'Expired Date'
        for c in df_asli.columns:
            if 'EXP' in str(c).upper() or 'DATE' in str(c).upper() or 'KEDALUWARSA' in str(c).upper():
                kolom_exp_kwn = c
                break
                
        if kolom_exp_kwn in df_asli.columns:
            df_asli['Exp_Clean'] = pd.to_datetime(df_asli[kolom_exp_kwn], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df_asli[kolom_exp_kwn].astype(str))
        else:
            df_asli['Exp_Clean'] = "-"
            
        df_memo = pd.DataFrame()
        # Logika Spasi Memo: Hanya BATCH yang bersih dari spasi internal
        lot_clean = df_asli['Lot No.'].astype(str).str.replace(r'\s+', '', regex=True)
        name_normal = df_asli['Item Name'].astype(str).str.strip()
        stock_normal = df_asli['End Stock'].astype(str).str.strip()
        
        df_memo['MEMO'] = lot_clean + "//" + name_normal + "=" + stock_normal + "BOX"
        df_memo['BATCH'] = df_asli['Lot No.']
        df_memo['PRODUK'] = df_asli['Item Name']
        df_memo['BOX'] = df_asli['End Stock']
        df_memo['KG'] = df_asli['Weight (Kg)']
        df_memo['EXPIRED DATE'] = df_asli['Exp_Clean']
        
        gudang_data["GUDANG KAWANISHI"] = {"asli": df_asli, "rekap": df_memo}
        total_box_gabungan += df_memo['BOX'].sum()
        total_kg_gabungan += df_memo['KG'].sum()
        jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses berkas gabungan Gudang Kawanishi: {e}")


# =========================================================================
# --- PROSES GUDANG DPR (GUDANG 11) ---
# =========================================================================
if file_gudang11:
    try:
        list_df_asli_dpr = []
        list_df_memo_dpr = []
        kolom_wajib_dpr = ['No.', 'Location', 'Item #', 'Description', 'Lot #', 'Production #', 'PID', 'Pallet ID', 'Stock Qty (Box)', 'Stock Qty (Kg)', 'Storage Date']
        
        for file_dpr in file_gudang11:
            nama_file_lc = file_dpr.name.lower()
            if 'bdk' in nama_file_lc or 'berdikari' in nama_file_lc: 
                label_sumber = "BKR"
            elif 'ppi' in nama_file_lc: 
                label_sumber = "PPI"
            elif 'sapi' in nama_file_lc: 
                label_sumber = "SAPI"
            else: 
                label_sumber = "LAINNYA"
                
            dict_sheets = pd.read_excel(file_dpr, sheet_name=None) if not file_dpr.name.endswith('.csv') else {file_dpr.name: pd.read_csv(file_dpr)}
            for nama_sheet, df_raw in dict_sheets.items():
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                df_clean = df_raw[df_raw['No.'].notnull()].copy()
                
                if not df_clean.empty:
                    kolom_tersedia = [c for c in kolom_wajib_dpr if c in df_clean.columns]
                    df_clean = df_clean[kolom_tersedia]
                    
                    if 'Stock Qty (Box)' in df_clean.columns: 
                        df_clean['Stock Qty (Box)'] = pd.to_numeric(df_clean['Stock Qty (Box)'], errors='coerce').fillna(0).astype(int)
                    if 'Stock Qty (Kg)' in df_clean.columns: 
                        df_clean['Stock Qty (Kg)'] = pd.to_numeric(df_clean['Stock Qty (Kg)'], errors='coerce').fillna(0)
                        
                    if 'Description' in df_clean.columns: 
                        df_clean['Description'] = df_clean['Description'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
                    if 'Lot #' in df_clean.columns: 
                        df_clean['Lot #'] = df_clean['Lot #'].astype(str).str.strip()
                    if 'Production #' in df_clean.columns: 
                        df_clean['Production #'] = df_clean['Production #'].astype(str).str.strip()
                        
                    df_clean['SUMBER_FILE'] = label_sumber
                    list_df_asli_dpr.append(df_clean)
                    
                                        # --- PERBAIKAN LOGIKA REKAP GUDANG DPR (TAMBAH DATA PALLET ID) ---
                    df_memo = pd.DataFrame()
                    
                    # 1. Kolom BATCH (Production #) dibersihkan dari spasi internal
                    production_clean = df_clean['Production #'].astype(str).str.replace(r'\s+', '', regex=True)
                    
                    # 2. Kolom lainnya tetap mempertahankan spasi aslinya
                    desc_normal = df_clean['Description'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
                    lot_normal = df_clean['Lot #'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
                    box_normal = df_clean['Stock Qty (Box)'].astype(str).str.strip()
                    
                    # 3. Formula Memo tetap sama rapat
                    df_memo['MEMO'] = production_clean + "//" + desc_normal + "-" + lot_normal + "-" + label_sumber + "=" + box_normal + "BOX"
                    
                    # 4. Tambahkan kolom data asli Pallet ID ke tabel rekap
                    df_memo['BATCH'] = df_clean['Production #']
                    df_memo['PRODUK'] = df_clean['Description']
                    df_memo['BOX'] = df_clean['Stock Qty (Box)']
                    df_memo['KG'] = df_clean['Stock Qty (Kg)']
                    df_memo['PALLET ID'] = df_clean['Pallet ID'].astype(str).str.strip() # 🟢 Ambil kolom Pallet ID
                    df_memo['SUMBER_FILE'] = label_sumber
                    
                    list_df_memo_dpr.append(df_memo)

                    
        if list_df_asli_dpr and list_df_memo_dpr:
            gudang_data["GUDANG DPR"] = {
                "asli": pd.concat(list_df_asli_dpr, ignore_index=True), 
                "rekap": pd.concat(list_df_memo_dpr, ignore_index=True)
            }
            total_box_gabungan += gudang_data["GUDANG DPR"]["rekap"]['BOX'].sum()
            total_kg_gabungan += gudang_data["GUDANG DPR"]["rekap"]['KG'].sum()
            jumlah_gudang_aktif += 1
    except Exception as e:
        st.error(f"Gagal memproses susunan berkas gabungan Gudang DPR (Gudang 11): {e}")
# =========================================================================
# 4. TAMPILAN RINGKASAN UTAMA (KPI AKUMULASI GLOBAL)
# =========================================================================
if gudang_data:
    st.subheader("📊 Totals Akumulasi Operasional (11 Gudang Gabungan)")
    kpi1, kpi2, kpi3 = st.columns(3)
    with kpi1: 
        st.metric("Total Keseluruhan Stok (BOX)", f"{total_box_gabungan:,}")
    with kpi2: 
        st.metric("Total Berat Muatan Daging (KG)", f"{total_kg_gabungan:,.2f}")
    with kpi3: 
        st.metric("Jumlah Gudang Terisi Data", f"{jumlah_gudang_aktif} dari 11 Gudang")
    
    st.markdown("---")
  
    # =========================================================================
    # 5. TAMPILAN TABEL TERPISAH PER-GUDANG (TAB INTERAKTIF)
    # =========================================================================
    st.subheader("📋 Pemantauan Laporan Terpisah per Gudang")
    nama_tab_gudang = list(gudang_data.keys())
    tabs_gudang_sewa = st.tabs(nama_tab_gudang)
    
    for indeks, nama_gudang in enumerate(nama_tab_gudang):
        with tabs_gudang_sewa[indeks]:
            st.success(f"📦 Data untuk {nama_gudang} Berhasil Ditampilkan")
            
            # --- A. TAMPILAN GUDANG MEGA ---
            if nama_gudang == "GUDANG MEGA":
                df_rekap_ini = gudang_data[nama_gudang]["rekap"]
                df_asli_ini = gudang_data[nama_gudang]["asli"]
                
                col_sub1, col_sub2 = st.columns(2)
                col_sub1.metric("Total Stok GUDANG MEGA (BOX)", f"{df_rekap_ini['BOX'].sum():,}")
                col_sub2.metric("Total Berat GUDANG MEGA (KG)", f"{df_rekap_ini['KG'].sum():,.2f}")
                
                st.markdown("##### 📝 Tabel Rekapitulasi Memo - GUDANG MEGA")
                st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'SUB_GEDUNG', 'EXPIRED']], use_container_width=True, hide_index=True)
                
                csv_rekap = df_rekap_ini.to_csv(index=False).encode('utf-8')
                st.download_button(label="📥 Download Rekap Memo GUDANG MEGA (CSV)", data=csv_rekap, file_name="rekap_memo_gudang_mega.csv", mime="text/csv", key="dl_rekap_mega_gabungan")
                st.write("")
                st.markdown("##### 📋 Tabel Laporan Excel Asli - GUDANG MEGA")
                st.dataframe(df_asli_ini, use_container_width=True, hide_index=True)

            # --- B. TAMPILAN GUDANG MARINA ---
            elif nama_gudang == "GUDANG MARINA":
                df_rekap_ini = gudang_data[nama_gudang]["rekap"]
                df_asli_ini = gudang_data[nama_gudang]["asli"]
                
                col_sub1, col_sub2 = st.columns(2)
                col_sub1.metric("Total Stok GUDANG MARINA (BOX)", f"{df_rekap_ini['BOX'].sum():,}")
                col_sub2.metric("Total Berat GUDANG MARINA (KG)", f"{df_rekap_ini['KG'].sum():,.2f}")
                
                st.markdown("##### 📝 Tabel Rekapitulasi Memo - GUDANG MARINA")
                st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'SUMBER', 'EXPIRED']], use_container_width=True, hide_index=True)
                
                csv_rekap = df_rekap_ini.to_csv(index=False).encode('utf-8')
                st.download_button(label="📥 Download Rekap Memo GUDANG MARINA (CSV)", data=csv_rekap, file_name="rekap_memo_gudang_marina.csv", mime="text/csv", key="dl_marina_unified")
                st.write("")
                st.markdown("##### 📋 Tabel Laporan Excel Asli - GUDANG MARINA")
                st.dataframe(df_asli_ini, use_container_width=True, hide_index=True)

            # --- C. TAMPILAN GUDANG PUM ---
            elif nama_gudang == "GUDANG PUM":
                df_rekap_ini = gudang_data[nama_gudang]["master_pum"]
                df_asli_gabungan_ini = gudang_data[nama_gudang]["asli_gabungan"]
                
                col_sub1, col_sub2 = st.columns(2)
                col_sub1.metric("Total Stok GUDANG PUM (BOX)", f"{df_rekap_ini['BOX'].sum():,}")
                col_sub2.metric("Total Berat GUDANG PUM (KG)", f"{df_rekap_ini['KG'].sum():,.2f}")
                
                st.markdown("##### 📝 Tabel Rekapitulasi Memo Gabungan - GUDANG PUM")
                st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'SUMBER_SHEET', 'Expired Date']], use_container_width=True, hide_index=True)
                
                csv_rekap = df_rekap_ini.to_csv(index=False).encode('utf-8')
                st.download_button(label="📥 Download Rekap Memo GUDANG PUM (CSV)", data=csv_rekap, file_name="rekap_memo_gudang_pum.csv", mime="text/csv", key="dl_pum_unified")
                st.write("")
                st.markdown("##### 📋 Tabel Excel Asli Gabungan (BKR & PPI) - GUDANG PUM")
                st.dataframe(df_asli_gabungan_ini, use_container_width=True, hide_index=True)

            # --- D. TAMPILAN GUDANG KAWANISHI ---
            elif nama_gudang == "GUDANG KAWANISHI":
                df_rekap_ini = gudang_data[nama_gudang]["rekap"]
                df_asli_ini = gudang_data[nama_gudang]["asli"]
                
                col_sub1, col_sub2 = st.columns(2)
                col_sub1.metric("Total Stok GUDANG KAWANISHI (BOX)", f"{df_rekap_ini['BOX'].sum():,}")
                col_sub2.metric("Total Berat GUDANG KAWANISHI (KG)", f"{df_rekap_ini['KG'].sum():,.2f}")
                
                st.markdown("##### 📝 Tabel Rekapitulasi Memo - GUDANG KAWANISHI")
                st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'EXPIRED DATE']], use_container_width=True, hide_index=True)
                
                csv_rekap = df_rekap_ini.to_csv(index=False).encode('utf-8')
                st.download_button(label="📥 Download Rekap Memo KAWANISHI (CSV)", data=csv_rekap, file_name="rekap_memo_kawanishi.csv", mime="text/csv", key=f"dl_kwn_{indeks}")
                st.write("")
                st.markdown("##### 📋 Tabel Laporan Excel Asli - GUDANG KAWANISHI")
                st.dataframe(df_asli_ini, use_container_width=True, hide_index=True)

            # =========================================================================
            # TAMPILAN GUDANG DPR (GUDANG 11 - TAMBAH KOLOM VISUAL PALLET ID)
            # =========================================================================
            elif nama_gudang == "GUDANG DPR":
                df_rekap_ini = gudang_data[nama_gudang]["rekap"]
                df_asli_ini = gudang_data[nama_gudang]["asli"]
                
                col_sub1, col_sub2 = st.columns(2)
                col_sub1.metric("Total Stok GUDANG DPR (BOX)", f"{df_rekap_ini['BOX'].sum():,}")
                col_sub2.metric("Total Berat GUDANG DPR (KG)", f"{df_rekap_ini['KG'].sum():,.2f}")
                
                st.markdown("##### 📝 Tabel Rekapitulasi Memo - GUDANG DPR")
                # 🟢 DI SINI KITA SISIPKAN 'PALLET ID' DI DALAM ARRAY KOLOM TAMPILAN
                st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'PALLET ID', 'SUMBER_FILE']], use_container_width=True, hide_index=True)
                
                csv_rekap = df_rekap_ini.to_csv(index=False).encode('utf-8')
                st.download_button(label="📥 Download Rekap Memo GUDANG DPR (CSV)", data=csv_rekap, file_name="rekap_memo_gudang_dpr.csv", mime="text/csv", key=f"dl_dpr_{indeks}")
                st.write("")
                st.markdown("##### 📋 Tabel Laporan Excel Asli - GUDANG DPR")
                st.dataframe(df_asli_ini, use_container_width=True, hide_index=True)


             # --- F. TAMPILAN GUDANG STANDAR LAINNYA (BOSCO, EINICHI, KIAT, MTP, RPI, TCL) ---
            else:
                st.markdown(f"##### 📝 Tabel Rekapitulasi Memo - {nama_gudang}")
                df_rekap_ini = gudang_data[nama_gudang]["rekap"]
                df_asli_ini = gudang_data[nama_gudang]["asli"]
                
                # Pemanggilan dinamis daftar susunan kolom untuk kolom Expired yang berbeda-beda nama
                if nama_gudang == "GUDANG KIAT":
                    st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'EXP_DATE', 'DESC_USAGE']], use_container_width=True, hide_index=True)
                elif nama_gudang == "GUDANG BOSCO":
                    st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'EXPIRE DATE']], use_container_width=True, hide_index=True)
                elif nama_gudang in ["GUDANG EINICHI", "GUDANG MTP", "GUDANG RPI", "GUDANG TCL"]:
                    kolom_exp_std = 'EXPIRED DATE' if nama_gudang == "GUDANG EINICHI" else ('EXP DATE' if nama_gudang == "GUDANG MTP" else 'EXP. DATE')
                    st.dataframe(df_rekap_ini[['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', kolom_exp_std]], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(df_rekap_ini, use_container_width=True, hide_index=True)
                
                csv_rekap = df_rekap_ini.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download Rekap Memo {nama_gudang} (CSV)", 
                    data=csv_rekap, 
                    file_name=f"rekap_memo_{nama_gudang.lower().replace(' ', '')}.csv", 
                    mime="text/csv", 
                    key=f"dl_rekap_{indeks}"
                )
                
                st.write("") 
                st.markdown(f"##### 📋 Tabel Laporan Excel Asli - {nama_gudang}")
                
                # Pembersihan kolom visual tabel asli diletakkan di sini (Sesuai lekukan fungsi tampilan gudang standar)
                if "Stock_Qty_Box_Clean" in df_asli_ini.columns:
                    df_asli_ini = df_asli_ini.drop(columns=["Stock_Qty_Box_Clean", "Stock_Qty_Kg_Clean"])
                st.dataframe(df_asli_ini, use_container_width=True, hide_index=True)

    # =========================================================================
    # 📥 AREA TOMBOL DOWNLOAD GABUNGAN (DI SEBELAH LUAR LOOP TAB / PALING BAWAH HALAMAN)
    # =========================================================================
    st.markdown("---")
    st.markdown("#### 📥 Download Rekap Gabungan Seluruh Gudang")
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        # --- 1. TOMBOL DOWNLOAD GABUNGAN MEMO ---
        buffer_memo = io.BytesIO()
        with pd.ExcelWriter(buffer_memo, engine='openpyxl') as writer:
            for nama_gudang_excel, data in gudang_data.items():
                if "rekap" in data:
                    nama_sheet_bersih = nama_gudang_excel.replace("GUDANG ", "")[:31]
                    data["rekap"].to_excel(writer, sheet_name=nama_sheet_bersih, index=False)
                elif "master_pum" in data:
                    nama_sheet_bersih = "PUM_MEMO"
                    data["master_pum"][['MEMO', 'BATCH', 'PRODUK', 'BOX', 'KG', 'SUMBER_SHEET', 'Expired Date']].to_excel(writer, sheet_name=nama_sheet_bersih, index=False)
        buffer_memo.seek(0)
        st.download_button(
            label="🟢 Download Semua Rekap Memo (1 File Excel - Multi Sheet)",
            data=buffer_memo,
            file_name=f"Rekap_Memo_11_Gudang_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_global_excel_multisheet_memo"
        )
        
    with col_dl2:
        # --- 2. TOMBOL DOWNLOAD SEMUA LAPORAN EXCEL ASLI ---
        buffer_asli = io.BytesIO()
        with pd.ExcelWriter(buffer_asli, engine='openpyxl') as writer:
            for nama_gudang_excel, data in gudang_data.items():
                nama_sheet_bersih = nama_gudang_excel.replace("GUDANG ", "")[:31]
                if nama_gudang_excel == "GUDANG PUM":
                    if "asli_gabungan" in data and not data["asli_gabungan"].empty:
                        data["asli_gabungan"].to_excel(writer, sheet_name="PUM_ASLI", index=False)
                else:
                    if "asli" in data and not data["asli"].empty:
                        data["asli"].to_excel(writer, sheet_name=nama_sheet_bersih, index=False)
        buffer_asli.seek(0)
        st.download_button(
            label="🔵 Download Semua Laporan Excel Asli (1 File Excel - Multi Sheet)",
            data=buffer_asli,
            file_name=f"Laporan_Asli_11_Gudang_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_global_excel_multisheet_asli_fixed"
        )
else:
    # Penutupan blok utama paling bawah jika belum ada satu pun file Excel gudang yang di-upload
    st.info("💡 Sistem Dasbor Sentral Siap. Silakan unggah berkas Excel SOH pada masing-masing tombol nama gudang sewa di atas.")
