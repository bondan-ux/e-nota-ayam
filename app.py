import streamlit as st
import pandas as pd
from datetime import datetime
import base64

st.set_page_config(page_title="E-Nota Bakul Ayam Segar", layout="wide")

def get_img_as_base64(file_path):
    with open(file_path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

# Memuat logo untuk judul dan watermark tengah
logo_base64 = get_img_as_base64("AST.jpeg")
watermark_base64 = get_img_as_base64("ASTremove.PNG")

# --- CSS ---
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
            background-color: rgba(255, 255, 255, 0.92);
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 450px;
            border-radius: 20px;
            padding: 3rem;
        }}
    </style>
""", unsafe_allow_html=True)

st.markdown(f"""
    <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
        <img src="data:image/jpeg;base64,{logo_base64}" style="width: 60px; height: 60px; border-radius: 50%; border: 2px solid #E53935;">
        <h1 style="margin: 0; font-size: 38px; color: #1E1E1E;">E-Nota Bakul Ayam Segar</h1>
    </div>
""", unsafe_allow_html=True)

# 1. Sidebar - Master Harga
st.sidebar.header("⚙️ Master Harga Harian")
harga_glondong = st.sidebar.number_input("Harga Glondong / Kg", value=28500, step=500)
harga_jeroan = st.sidebar.number_input("Harga Jeroan", value=12000, step=500)
harga_usus = st.sidebar.number_input("Harga Usus", value=16500, step=500)
harga_telur = st.sidebar.number_input("Harga Telur A", value=269000, step=1000)
harga_peti = st.sidebar.number_input("Harga Peti", value=2000, step=100)
harga_box = st.sidebar.number_input("Harga Box", value=28500, step=500)

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
    
    # Input Manual Qty
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        qty_peti = st.number_input("Jumlah Peti", value=0, step=1)
    
    # Logic Box (Manual + Info)
    row_bakul = df[df[name_col] == selected_bakul].iloc[0]
    col_ket = next((c for c in df.columns if 'KET' in c), 'KET')
    auto_box = float(row_bakul[col_ket]) if pd.notna(row_bakul.get(col_ket)) else 0.0
    
    with col_in2:
        qty_box = st.number_input("Jumlah Box (Manual)", value=float(auto_box), step=0.1)
        st.info(f"ℹ️ Otomatis dari Excel: {auto_box} Box")

    if selected_bakul:
        qty_tonase = float(row_bakul[next((c for c in df.columns if 'TONASE' in c), 0)])
        qty_jeroan = float(row_bakul[next((c for c in df.columns if 'JEROAN' in c), 0)])
        qty_usus = float(row_bakul[next((c for c in df.columns if 'USUS' in c), 0)])
        qty_telur = float(row_bakul[next((c for c in df.columns if 'TELUR' in c), 0)])

        total_bayar = (qty_tonase * harga_glondong) + (qty_jeroan * harga_jeroan) + \
                      (qty_usus * harga_usus) + (qty_telur * harga_telur) + \
                      (qty_peti * harga_peti) + (qty_box * harga_box)
        
        st.markdown("---")
        st.markdown("### 🐔 AYAM SEGAR TUMPANG")
        st.write(f"**Nama Bakul:** {selected_bakul} | **Group:** {selected_sheet} | **Tanggal:** {tanggal_transaksi.strftime('%d-%m-%Y')}")
        
        items = [
            {"Nama Barang": "GLONDONG", "KG": qty_tonase, "Harga": harga_glondong, "Jumlah": qty_tonase * harga_glondong},
            {"Nama Barang": "JEROAN", "KG": qty_jeroan, "Harga": harga_jeroan, "Jumlah": qty_jeroan * harga_jeroan},
            {"Nama Barang": "USUS B", "KG": qty_usus, "Harga": harga_usus, "Jumlah": qty_usus * harga_usus},
            {"Nama Barang": "TELUR", "KG": qty_telur, "Harga": harga_telur, "Jumlah": qty_telur * harga_telur},
            {"Nama Barang": "PETI", "KG": qty_peti, "Harga": harga_peti, "Jumlah": qty_peti * harga_peti},
            {"Nama Barang": "BOX", "KG": qty_box, "Harga": harga_box, "Jumlah": qty_box * harga_box},
        ]
        
        df_nota = pd.DataFrame([i for i in items if i['KG'] > 0])
        df_nota['Harga'] = df_nota['Harga'].map("Rp {:,.0f}".format)
        df_nota['Jumlah'] = df_nota['Jumlah'].map("Rp {:,.0f}".format)
        st.table(df_nota)
        st.subheader(f"TOTAL JUMLAH: Rp {total_bayar:,.0f}")