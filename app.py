import streamlit as st
import pandas as pd
from datetime import datetime
import base64
import json
import os

st.set_page_config(page_title="E-Nota Bakul Ayam Segar", layout="wide")

def get_img_as_base64(file_path):
    try:
        with open(file_path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except Exception:
        return ""

logo_base64 = get_img_as_base64("AST.jpeg")
watermark_base64 = get_img_as_base64("ASTremove.png")

st.markdown(f"""
    <style>
        header[data-testid="stHeader"] {{
            background: linear-gradient(90deg, #E53935 0%, #E53935 50%, #FFFFFF 50%, #FFFFFF 100%) !important;
        }}
        [data-testid="stSidebar"] {{
            background-color: #FFF4F4 !important;
            border-right: 3px solid #E53935 !important;
        }}
        .block-container {{
            background-color: rgba(255, 255, 255, 0.95);
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 450px;
            border-radius: 20px;
            padding: 3rem;
            box-shadow: 0px 4px 15px rgba(0,0,0,0.05);
        }}
        /* Sembunyikan elemen sidebar & kontrol saat diprint */
        @media print {{
            [data-testid="stSidebar"], .stFileUploader, div[data-baseweb="select"], .stDateInput, hr {{
                display: none !important;
            }}
            button {{
                display: none !important;
            }}
        }}
    </style>
""", unsafe_allow_html=True)

st.markdown(f"""
    <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
        <img src="data:image/jpeg;base64,{logo_base64}" style="width: 60px; height: 60px; border-radius: 50%; border: 2px solid #E53935;">
        <h1 style="margin: 0; font-size: 38px;">E-Nota Bakul Ayam Segar</h1>
    </div>
""", unsafe_allow_html=True)

# --- LOGIKA PENYIMPANAN HARGA MASTER ---
FILE_HARGA = "master_harga.json"

default_harga = {
    "glondong": 28500,
    "jeroan": 12000,
    "usus": 16500,
    "telur_a": 269000,
    "telur_b": 250000,
    "peti": 2000,
    "box": 28500
}

if os.path.exists(FILE_HARGA):
    try:
        with open(FILE_HARGA, "r") as f:
            saved_harga = json.load(f)
    except:
        saved_harga = default_harga
else:
    saved_harga = default_harga

# 1. Sidebar - Master Harga (Dengan Format Angka Rapi / Pemisah Ribuan)
st.sidebar.markdown("<h2>⚙️ Master Harga Harian</h2>", unsafe_allow_html=True)

harga_glondong = st.sidebar.number_input("Harga Glondong / Kg", value=int(saved_harga.get("glondong", 28500)), step=500, format="%d")
st.sidebar.caption(f"💡 Rp {harga_glondong:,.0f}".replace(",", "."))

harga_jeroan = st.sidebar.number_input("Harga Jeroan", value=int(saved_harga.get("jeroan", 12000)), step=500, format="%d")
st.sidebar.caption(f"💡 Rp {harga_jeroan:,.0f}".replace(",", "."))

harga_usus = st.sidebar.number_input("Harga Usus", value=int(saved_harga.get("usus", 16500)), step=500, format="%d")
st.sidebar.caption(f"💡 Rp {harga_usus:,.0f}".replace(",", "."))

harga_telur_a = st.sidebar.number_input("Harga Telur A", value=int(saved_harga.get("telur_a", 269000)), step=1000, format="%d")
st.sidebar.caption(f"💡 Rp {harga_telur_a:,.0f}".replace(",", "."))

harga_telur_b = st.sidebar.number_input("Harga Telur B", value=int(saved_harga.get("telur_b", 250000)), step=1000, format="%d")
st.sidebar.caption(f"💡 Rp {harga_telur_b:,.0f}".replace(",", "."))

harga_peti = st.sidebar.number_input("Harga Peti", value=int(saved_harga.get("peti", 2000)), step=100, format="%d")
st.sidebar.caption(f"💡 Rp {harga_peti:,.0f}".replace(",", "."))

harga_box = st.sidebar.number_input("Harga Box", value=int(saved_harga.get("box", 28500)), step=500, format="%d")
st.sidebar.caption(f"💡 Rp {harga_box:,.0f}".replace(",", "."))

current_harga = {
    "glondong": harga_glondong,
    "jeroan": harga_jeroan,
    "usus": harga_usus,
    "telur_a": harga_telur_a,
    "telur_b": harga_telur_b,
    "peti": harga_peti,
    "box": harga_box
}

if current_harga != saved_harga:
    with open(FILE_HARGA, "w") as f:
        json.dump(current_harga, f)
# ---------------------------------------

# 2. Upload
col_up1, col_up2 = st.columns([2, 1])
with col_up1:
    uploaded_file = st.file_uploader("Upload File Rekap Excel (.xlsx)", type=["xlsx"])
with col_up2:
    tanggal_transaksi = st.date_input("Tanggal Transaksi", value=datetime.today())

if uploaded_file is not None:
    xl = pd.ExcelFile(uploaded_file)
    sheets = [s for s in xl.sheet_names if s not in ['TOTAL TONASE', 'NOTA FR', 'Sheet1', 'ploting']]
    selected_sheet = st.selectbox("Pilih Group / Sheet", sheets)
    
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.upper().str.contains('NAMA').any():
            header_idx = idx
            break
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = [str(c).strip().upper() for c in df_raw.iloc[header_idx].values]
    name_col = next((c for c in df.columns if 'NAMA' in c), df.columns[1])
    df = df[df[name_col].notna()]
    
    selected_bakul = st.selectbox("Pilih Nama Bakul", df[name_col].unique())
    
    if 'last_bakul' not in st.session_state or st.session_state['last_bakul'] != selected_bakul:
        st.session_state['last_bakul'] = selected_bakul
        st.session_state['qty_box_val'] = 0.0

    if selected_bakul:
        row_bakul = df[df[name_col] == selected_bakul].iloc[0]
        
        def get_valid_float(col_name):
            try:
                if not col_name or col_name not in df.columns:
                    return 0.0
                val = row_bakul[col_name]
                num = pd.to_numeric(val, errors='coerce')
                return 0.0 if pd.isna(num) else float(num)
            except:
                return 0.0

        col_in1, col_in2 = st.columns(2)
        with col_in1:
            qty_peti = st.number_input("Jumlah Peti", value=0, step=1)
        with col_in2:
            qty_box = st.number_input("Jumlah Box (Manual)", key='qty_box_val', step=0.1)

        qty_tonase = get_valid_float(next((c for c in df.columns if 'TONASE' in c), ''))
        qty_jeroan = get_valid_float(next((c for c in df.columns if 'JEROAN' in c), ''))
        qty_usus = get_valid_float(next((c for c in df.columns if 'USUS' in c), ''))
        
        qty_telur_a = get_valid_float(next((c for c in df.columns if 'TELUR A' in c or 'TELUR' in c), ''))
        qty_telur_b = get_valid_float(next((c for c in df.columns if 'TELUR B' in c), ''))

        val_ket = get_valid_float(next((c for c in df.columns if 'KET' in c), ''))
        biaya_kresek = 7000 if (val_ket > 0 and not float(val_ket).is_integer()) else 0

        tot_glondong = qty_tonase * harga_glondong
        tot_jeroan = qty_jeroan * harga_jeroan
        tot_usus = qty_usus * harga_usus
        tot_telur_a = qty_telur_a * harga_telur_a
        tot_telur_b = qty_telur_b * harga_telur_b
        tot_peti = qty_peti * harga_peti
        tot_box = qty_box * harga_box

        total_bayar = tot_glondong + tot_jeroan + tot_usus + tot_telur_a + tot_telur_b + tot_peti + tot_box + biaya_kresek
        
        st.markdown("<hr style='border: 1px solid #ddd;'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color: #E53935; margin-bottom: 0px;'>🐔 AYAM SEGAR TUMPANG</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: #555; margin-bottom: 20px;'>Ds. Kambingan - Tumpang - Kab. Malang</p>", unsafe_allow_html=True)
        
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            st.write(f"**Nama Bakul:** {selected_bakul}")
            st.write(f"**Group:** {selected_sheet}")
        with col_n2:
            st.write(f"**Tanggal:** {tanggal_transaksi.strftime('%d-%m-%Y')}")
        
        st.markdown("<br>", unsafe_allow_html=True)

        items = [
            {"Nama Barang": "GLONDONG", "KG": qty_tonase, "Harga": harga_glondong, "Jumlah": tot_glondong},
            {"Nama Barang": "JEROAN", "KG": qty_jeroan, "Harga": harga_jeroan, "Jumlah": tot_jeroan},
            {"Nama Barang": "USUS B", "KG": qty_usus, "Harga": harga_usus, "Jumlah": tot_usus},
            {"Nama Barang": "TELUR A", "KG": qty_telur_a, "Harga": harga_telur_a, "Jumlah": tot_telur_a},
            {"Nama Barang": "TELUR B", "KG": qty_telur_b, "Harga": harga_telur_b, "Jumlah": tot_telur_b},
            {"Nama Barang": "PETI", "KG": qty_peti, "Harga": harga_peti, "Jumlah": tot_peti},
            {"Nama Barang": "BOX", "KG": qty_box, "Harga": harga_box, "Jumlah": tot_box},
            {"Nama Barang": "BIAYA KRESEK", "KG": 1 if biaya_kresek > 0 else 0, "Harga": 7000, "Jumlah": biaya_kresek},
        ]
        
        filtered_items = [i for i in items if i['KG'] > 0]
        
        if filtered_items:
            df_nota = pd.DataFrame(filtered_items)
            df_nota['KG'] = df_nota['KG'].apply(lambda x: f"{int(x)}" if isinstance(x, float) and x.is_integer() else f"{x:.2f}" if isinstance(x, float) else str(x))
            df_nota['Harga'] = df_nota['Harga'].map("Rp {:,.0f}".format)
            df_nota['Jumlah'] = df_nota['Jumlah'].map("Rp {:,.0f}".format)
            
            st.table(df_nota)
            
            col_t1, col_t2 = st.columns([2, 1])
            with col_t1:
                st.markdown("""
                    <button onclick="window.print()" style="background-color: #E53935; color: white; padding: 10px 20px; border: none; border-radius: 5px; font-weight: bold; cursor: pointer; font-size: 16px;">
                        🖨️ Cetak / Print Nota
                    </button>
                """, unsafe_allow_html=True)
            with col_t2:
                st.markdown(f"<h3 style='text-align: right; margin-top: 0px;'>TOTAL: Rp {total_bayar:,.0f}</h3>", unsafe_allow_html=True)
        else:
            st.warning("Tidak ada item pembelian untuk bakul ini.")