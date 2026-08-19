import streamlit as st
import pandas as pd
from datetime import datetime
import base64

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

# --- CSS ANTI DARK-MODE ---
st.markdown(f"""
    <style>
        /* Paksa background utama jadi terang */
        .stApp {{
            background-color: #F8F9FA !important;
        }}
        
        /* Header Bendera Merah Putih */
        header[data-testid="stHeader"] {{
            background: linear-gradient(90deg, #E53935 0%, #E53935 50%, #FFFFFF 50%, #FFFFFF 100%) !important;
        }}
        
        /* Sidebar */
        [data-testid="stSidebar"] {{
            background-color: #FFF4F4 !important;
            border-right: 3px solid #E53935 !important;
        }}
        
        /* Paksa SEMUA teks di sidebar jadi gelap */
        [data-testid="stSidebar"] * {{
            color: #1E1E1E !important;
        }}

        /* Wadah konten utama dengan watermark */
        .block-container {{
            background-color: rgba(255, 255, 255, 0.95) !important;
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 450px;
            border-radius: 20px;
            padding: 3rem;
            box-shadow: 0px 4px 15px rgba(0,0,0,0.05);
        }}
        
        /* Paksa SEMUA teks di area kanan jadi gelap */
        .block-container * {{
            color: #1E1E1E !important;
        }}
    </style>
""", unsafe_allow_html=True)

st.markdown(f"""
    <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
        <img src="data:image/jpeg;base64,{logo_base64}" style="width: 60px; height: 60px; border-radius: 50%; border: 2px solid #E53935;">
        <h1 style="margin: 0; font-size: 38px; color: #1E1E1E !important;">E-Nota Bakul Ayam Segar</h1>
    </div>
""", unsafe_allow_html=True)

# 1. Sidebar - Master Harga
st.sidebar.markdown("<h2 style='color:#1E1E1E;'>⚙️ Master Harga Harian</h2>", unsafe_allow_html=True)
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
    
    if selected_bakul:
        row_bakul = df[df[name_col] == selected_bakul].iloc[0]
        
        # FUNGSI ANTI NAN (DIJAMIN AMAN DARI ERROR EXCEL)
        def get_valid_float(col_name):
            try:
                if not col_name or col_name not in df.columns:
                    return 0.0
                val = row_bakul[col_name]
                # Paksa ubah ke angka, kalau error/kosong langsung jadikan NaN, lalu diganti ke 0.0
                num = pd.to_numeric(val, errors='coerce')
                return 0.0 if pd.isna(num) else float(num)
            except:
                return 0.0

        # Ambil nilai otomatis box
        col_ket = next((c for c in df.columns if 'KET' in c), 'KET')
        auto_box = get_valid_float(col_ket)

        # Input Manual Qty Peti & Box
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            qty_peti = st.number_input("Jumlah Peti", value=0, step=1)
        with col_in2:
            qty_box = st.number_input("Jumlah Box (Manual)", value=round(auto_box, 2), step=0.1)
            st.info(f"ℹ️ Otomatis dari Excel: {auto_box:.2f} Box")

        # Tarik semua data kuantitas
        qty_tonase = get_valid_float(next((c for c in df.columns if 'TONASE' in c), ''))
        qty_jeroan = get_valid_float(next((c for c in df.columns if 'JEROAN' in c), ''))
        qty_usus = get_valid_float(next((c for c in df.columns if 'USUS' in c), ''))
        qty_telur = get_valid_float(next((c for c in df.columns if 'TELUR' in c), ''))

        # Hitung total
        tot_glondong = qty_tonase * harga_glondong
        tot_jeroan = qty_jeroan * harga_jeroan
        tot_usus = qty_usus * harga_usus
        tot_telur = qty_telur * harga_telur
        tot_peti = qty_peti * harga_peti
        tot_box = qty_box * harga_box

        total_bayar = tot_glondong + tot_jeroan + tot_usus + tot_telur + tot_peti + tot_box
        
        st.markdown("<hr style='border: 1px solid #ddd;'>", unsafe_allow_html=True)
        st.markdown("<h3 style='color: #E53935 !important; margin-bottom: 0px;'>🐔 AYAM SEGAR TUMPANG</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: #555 !important; margin-bottom: 20px;'>Ds. Kambingan - Tumpang - Kab. Malang</p>", unsafe_allow_html=True)
        
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
            {"Nama Barang": "TELUR", "KG": qty_telur, "Harga": harga_telur, "Jumlah": tot_telur},
            {"Nama Barang": "PETI", "KG": qty_peti, "Harga": harga_peti, "Jumlah": tot_peti},
            {"Nama Barang": "BOX", "KG": qty_box, "Harga": harga_box, "Jumlah": tot_box},
        ]
        
        filtered_items = [i for i in items if i['KG'] > 0]
        
        if filtered_items:
            df_nota = pd.DataFrame(filtered_items)
            df_nota['KG'] = df_nota['KG'].apply(lambda x: f"{int(x)}" if isinstance(x, float) and x.is_integer() else f"{x:.2f}" if isinstance(x, float) else str(x))
            df_nota['Harga'] = df_nota['Harga'].map("Rp {:,.0f}".format)
            df_nota['Jumlah'] = df_nota['Jumlah'].map("Rp {:,.0f}".format)
            
            # Kembalikan ke tabel standar Streamlit biar otomatis ngikut styling baru
            st.table(df_nota)
            st.markdown(f"<h3 style='color: #1E1E1E !important; text-align: left; margin-top: 15px;'>TOTAL JUMLAH: Rp {total_bayar:,.0f}</h3>", unsafe_allow_html=True)
        else:
            st.warning("Tidak ada item pembelian untuk bakul ini.")