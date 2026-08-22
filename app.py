import streamlit as st
import pandas as pd
from datetime import datetime
import base64
import json
import os

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Sistem Manajemen Ayam Segar", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- FUNGSI GAMBAR ---
def get_img_as_base64(file_path):
    try:
        with open(file_path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except Exception:
        return ""

watermark_base64 = get_img_as_base64("ASTremove.png")

# --- CSS STYLING ---
st.markdown(f"""
    <style>
        footer, #MainMenu {{
            visibility: hidden;
        }}
        [data-testid="stHeader"] {{
            background-color: transparent !important;
            z-index: 100 !important;
        }}
        [data-testid="stSidebar"] {{
            background-color: #FFEBEE !important;
            border-right: 3px solid #C62828 !important;
        }}
        [data-testid="stSidebar"] .stRadio label {{
            font-size: 15px !important;
            font-weight: bold !important;
            color: #262626 !important;
        }}
        .block-container {{
            background-color: #FFFFFF;
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 400px;
            padding-top: 1rem !important;
        }}
        @media print {{
            [data-testid="stSidebar"], [data-testid="stHeader"], .stFileUploader, .stSelectbox, .stDateInput, .stNumberInput, button {{
                display: none !important;
            }}
            .block-container {{
                padding: 0 !important;
                margin: 0 !important;
            }}
        }}
    </style>
""", unsafe_allow_html=True)

# --- 2. SISTEM LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = ""

def login():
    st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if watermark_base64:
            st.markdown(f"""
                <div style='text-align: center; margin-bottom: 15px;'>
                    <img src="data:image/png;base64,{watermark_base64}" style="width: 220px;">
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<h3 style='text-align: center; color: #C62828;'>Ayam Segar Tumpang</h3>", unsafe_allow_html=True)
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        
        if st.button("Masuk / Login", use_container_width=True):
            if username == "admin" and password == "admin123":
                st.session_state.logged_in = True
                st.session_state.role = "Admin"
                st.rerun()
            elif username == "kasir" and password == "kasir123":
                st.session_state.logged_in = True
                st.session_state.role = "Kasir"
                st.rerun()
            else:
                st.error("Username atau Password salah!")

if not st.session_state.logged_in:
    login()
    st.stop()


# --- 3. MASTER HARGA JSON ---
FILE_HARGA = "master_harga.json"
default_harga = {"glondong": 28500, "jeroan": 12000, "usus": 16500, "telur_a": 269000, "telur_b": 250000, "peti": 2000, "box": 28500}

if os.path.exists(FILE_HARGA):
    try:
        with open(FILE_HARGA, "r") as f:
            saved_harga = json.load(f)
    except:
        saved_harga = default_harga
else:
    saved_harga = default_harga


# --- 4. SIDEBAR NAVIGATION & SET HARGA ---
with st.sidebar:
    st.markdown("<h2 style='color: #C62828;'>🐔 AST SYSTEM</h2>", unsafe_allow_html=True)
    st.write(f"Logged in as: **{st.session_state.role}**")
    st.markdown("---")
    
    st.markdown("### 📌 NAVIGASI UTAMA")
    menu_options = ["📊 Dashboard", "🧾 Nota", "🛍️ Penjualan", "📦 Stock", "💵 Finance", "⏱️ Absensi & Jadwal"]
    selected_menu = st.radio("Pilih Halaman:", menu_options)

    sub_menu = None
    if selected_menu == "🧾 Nota":
        with st.expander("📂 Sub-Menu Nota", expanded=True):
            sub_menu = st.radio("Tipe Nota:", ["📑 Bakul", "🏬 Bedak", "🤝 Mitra"])

    elif selected_menu == "⏱️ Absensi & Jadwal":
        with st.expander("📂 Sub-Menu Absensi", expanded=True):
            sub_menu = st.radio("Tipe Absen:", ["📅 Atur Jadwal", "📌 Plotting & Input Absen", "📊 Rekap Bulanan"])

    # MASTER HARGA DITAMPILKAN DI SIDEBAR (KHUSUS HALAMAN NOTA)
    if selected_menu == "🧾 Nota":
        st.markdown("---")
        st.markdown("<h3 style='color: #C62828;'>⚙️ Master Harga</h3>", unsafe_allow_html=True)
        
        def input_harga(label, key_name, default_val, step_val):
            val = st.number_input(
                f"{label}", 
                value=int(saved_harga.get(key_name, default_val)), 
                step=step_val, 
                format="%d"
            )
            st.caption(f"➔ **Rp {val:,.0f}**".replace(",", "."))
            return val

        h_glondong = input_harga("Harga Glondong", "glondong", 28500, 500)
        h_jeroan = input_harga("Harga Jeroan", "jeroan", 12000, 500)
        h_usus = input_harga("Harga Usus", "usus", 16500, 500)
        h_telur_a = input_harga("Harga Telur A", "telur_a", 269000, 1000)
        h_telur_b = input_harga("Harga Telur B", "telur_b", 250000, 1000)
        h_peti = input_harga("Harga Peti", "peti", 2000, 100)
        h_box = input_harga("Harga Box", "box", 28500, 500)

        current_harga = {
            "glondong": h_glondong, 
            "jeroan": h_jeroan, 
            "usus": h_usus, 
            "telur_a": h_telur_a, 
            "telur_b": h_telur_b, 
            "peti": h_peti, 
            "box": h_box
        }
        if current_harga != saved_harga:
            with open(FILE_HARGA, "w") as f:
                json.dump(current_harga, f)

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.role = ""
        st.rerun()


# --- 5. AREA KONTEN UTAMA ---
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"<h2 style='color: #C62828; margin:0;'>Ayam Segar Tumpang - {selected_menu}</h2>", unsafe_allow_html=True)
with col_h2:
    st.markdown(f"<div style='text-align: right; font-weight: bold; padding-top: 10px;'>👤 {st.session_state.role}</div>", unsafe_allow_html=True)

st.markdown("<hr style='border: 1px solid #C62828; margin-top: 5px; margin-bottom: 20px;'>", unsafe_allow_html=True)

# LOGIKA PEMBACAAN EXCEL & PENAMPILAN DATA NOTA
if selected_menu == "📊 Dashboard":
    st.info("Visualisasi data omset, stok, dan performa harian akan ditampilkan di sini.")

elif selected_menu == "🧾 Nota":
    if sub_menu == "📑 Bakul" or sub_menu is None:
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
            no_col = next((c for c in df.columns if c in ['NO', 'NO.', 'NOMOR']), None)
            
            df = df[df[name_col].notna()].copy()
            
            bakul_options = []
            bakul_map = {}
            
            for idx, (real_idx, row) in enumerate(df.iterrows(), start=1):
                nama_bakul = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                if not nama_bakul:
                    continue
                
                if no_col and pd.notna(row[no_col]) and str(row[no_col]).strip() != '':
                    no_val = str(row[no_col]).strip()
                    if no_val.endswith('.0'):
                        no_val = no_val[:-2]
                    label = f"{no_val}. {nama_bakul}"
                else:
                    label = f"{idx}. {nama_bakul}"
                    
                bakul_options.append(label)
                bakul_map[label] = (nama_bakul, real_idx)

            selected_bakul_label = st.selectbox("Pilih Nama Bakul", bakul_options)
            selected_bakul_info = bakul_map.get(selected_bakul_label, None)

            if selected_bakul_info:
                selected_bakul, real_idx = selected_bakul_info
                
                if 'last_bakul' not in st.session_state or st.session_state['last_bakul'] != selected_bakul_label:
                    st.session_state['last_bakul'] = selected_bakul_label
                    st.session_state['qty_box_val'] = 0.0

                row_bakul = df.loc[real_idx]
                
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

                tot_glondong = qty_tonase * h_glondong
                tot_jeroan = qty_jeroan * h_jeroan
                tot_usus = qty_usus * h_usus
                tot_telur_a = qty_telur_a * h_telur_a
                tot_telur_b = qty_telur_b * h_telur_b
                tot_peti = qty_peti * h_peti
                tot_box = qty_box * h_box

                total_bayar = tot_glondong + tot_jeroan + tot_usus + tot_telur_a + tot_telur_b + tot_peti + tot_box + biaya_kresek
                
                st.markdown("<hr style='border: 2px dashed #C62828;'>", unsafe_allow_html=True)
                st.markdown("<h3 style='color: #C62828; margin-bottom: 0px;'>🐔 AYAM SEGAR TUMPANG</h3>", unsafe_allow_html=True)
                st.markdown("<p style='color: #333; margin-bottom: 20px;'>Ds. Kambingan - Tumpang - Kab. Malang</p>", unsafe_allow_html=True)
                
                col_n1, col_n2 = st.columns(2)
                with col_n1:
                    st.write(f"**Nama Bakul:** {selected_bakul}")
                    st.write(f"**Group:** {selected_sheet}")
                with col_n2:
                    st.write(f"**Tanggal:** {tanggal_transaksi.strftime('%d-%m-%Y')}")
                
                st.markdown("<br>", unsafe_allow_html=True)

                items = [
                    {"Nama Barang": "GLONDONG", "KG": qty_tonase, "Harga": h_glondong, "Jumlah": tot_glondong},
                    {"Nama Barang": "JEROAN", "KG": qty_jeroan, "Harga": h_jeroan, "Jumlah": tot_jeroan},
                    {"Nama Barang": "USUS B", "KG": qty_usus, "Harga": h_usus, "Jumlah": tot_usus},
                    {"Nama Barang": "TELUR A", "KG": qty_telur_a, "Harga": h_telur_a, "Jumlah": tot_telur_a},
                    {"Nama Barang": "TELUR B", "KG": qty_telur_b, "Harga": h_telur_b, "Jumlah": tot_telur_b},
                    {"Nama Barang": "PETI", "KG": qty_peti, "Harga": h_peti, "Jumlah": tot_peti},
                    {"Nama Barang": "BOX", "KG": qty_box, "Harga": h_box, "Jumlah": tot_box},
                    {"Nama Barang": "BIAYA KRESEK", "KG": 1 if biaya_kresek > 0 else 0, "Harga": 7000, "Jumlah": biaya_kresek},
                ]
                
                filtered_items = [i for i in items if i['KG'] > 0]
                
                if filtered_items:
                    df_nota = pd.DataFrame(filtered_items)
                    df_nota['KG'] = df_nota['KG'].apply(lambda x: f"{int(x)}" if isinstance(x, float) and x.is_integer() else f"{x:.2f}" if isinstance(x, float) else str(x))
                    df_nota['Harga'] = df_nota['Harga'].map("Rp {:,.0f}".format).str.replace(",", ".")
                    df_nota['Jumlah'] = df_nota['Jumlah'].map("Rp {:,.0f}".format).str.replace(",", ".")
                    
                    st.table(df_nota)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_t1, col_t2 = st.columns([1, 1])
                    with col_t1:
                        st.markdown("""
                            <button onclick="window.print()" style="background-color: #C62828; color: white; padding: 12px 25px; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 16px;">
                                🖨️ Cetak / Print Nota
                            </button>
                        """, unsafe_allow_html=True)
                    with col_t2:
                        st.markdown(f"<h2 style='text-align: right; margin: 0px; color: #C62828; font-weight: bold;'>TOTAL: Rp {total_bayar:,.0f}</h2>".replace(",", "."), unsafe_allow_html=True)
                else:
                    st.warning("Tidak ada item transaksi untuk bakul ini.")

    elif sub_menu == "🏬 Bedak":
        st.info("Fitur nota bedak sedang dikembangkan.")
    elif sub_menu == "🤝 Mitra":
        st.info("Fitur nota mitra sedang dikembangkan.")

elif selected_menu == "🛍️ Penjualan":
    st.info("Halaman Penjualan Harian & Bulanan.")
elif selected_menu == "📦 Stock":
    st.info("Halaman Inventori Stok Barang.")
elif selected_menu == "💵 Finance":
    st.info("Halaman Pencatatan Kas & Piutang.")
elif selected_menu == "⏱️ Absensi & Jadwal":
    st.subheader(f"Menu Absensi: {sub_menu if sub_menu else 'Atur Jadwal'}")