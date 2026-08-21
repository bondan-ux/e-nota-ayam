import streamlit as st
import pandas as pd
from datetime import datetime
import base64
import json
import os

# Konfigurasi Halaman Utama
st.set_page_config(page_title="Sistem Manajemen Ayam Segar", layout="wide", initial_sidebar_state="expanded")

# --- FUNGSI GAMBAR ---
def get_img_as_base64(file_path):
    try:
        with open(file_path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except Exception:
        return ""

logo_base64 = get_img_as_base64("AST.jpeg")
watermark_base64 = get_img_as_base64("ASTremove.png")

# --- CSS STYLING & CUSTOM NAVBAR ---
st.markdown(f"""
    <style>
        /* Sembunyikan Toolbar, MainMenu, dan Tombol "Manage App" bawaan Streamlit Cloud */
        [data-testid="stToolbar"], #MainMenu, header[data-testid="stHeader"] {{
            display: none !important;
        }}
        
        div[class*="viewerBadge"], .stAppViewerFooter, [data-testid="stStatusWidget"],
        [data-testid="manage-app-button"], iframe[title="streamlitApp"] ~ div {{
            display: none !important;
        }}

        /* Custom Navbar Merah di atas dengan Logo & Teks */
        .custom-navbar {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 70px;
            background-color: #C62828;
            z-index: 999999;
            display: flex;
            align-items: center;
            padding-left: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }}
        
        .custom-navbar img {{
            height: 55px;
            margin-right: 15px;
        }}
        .custom-navbar span {{
            color: white;
            font-size: 24px;
            font-weight: bold;
        }}

        /* Sidebar Styling */
        [data-testid="stSidebar"] {{
            background-color: #FFEBEE !important; 
            border-right: 4px solid #C62828 !important;
            margin-top: 70px;
        }}

        /* Memperbesar Font Navigasi Utama */
        [data-testid="stSidebar"] .stRadio label {{
            font-size: 18px !important;
            font-weight: bold !important;
            color: #262626 !important;
            padding: 6px 0px !important;
        }}
        
        /* Custom Styling Expander / Sub-Menu Slide */
        [data-testid="stSidebar"] .streamlit-expanderHeader {{
            font-size: 16px !important;
            font-weight: bold !important;
            color: #C62828 !important;
            background-color: #FFCDD2 !important;
            border-radius: 6px !important;
        }}

        /* Area Kerja Utama & Jarak Bawah Tambahan */
        .block-container {{
            margin-top: 50px;
            background-color: #FFFFFF;
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 450px;
            padding: 2rem 3rem 6rem 3rem !important; /* Bottom padding diperbesar agar footer/total tidak kepotong */
        }}
        
        /* Tombol Utama */
        .stButton>button {{
            background-color: #C62828 !important;
            color: white !important;
            font-weight: bold !important;
            border-radius: 8px !important;
            border: none !important;
        }}
        .stButton>button:hover {{
            background-color: #B71C1C !important;
            border: none !important;
        }}

        @media print {{
            [data-testid="stSidebar"], .custom-navbar, .stTabs, .stFileUploader, div[data-baseweb="select"], .stDateInput, hr, button {{
                display: none !important;
            }}
            .block-container {{
                margin-top: 0px !important;
                padding-bottom: 0px !important;
            }}
        }}
    </style>
    
    <!-- Elemen HTML Custom Navbar -->
    <div class="custom-navbar">
        <img src="data:image/png;base64,{watermark_base64}">
        <span>Ayam Segar Tumpang</span>
    </div>
""", unsafe_allow_html=True)

# --- 1. SISTEM LOGIN (SESSION STATE) ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = ""

def login():
    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown(f"""
            <div style='text-align: center; margin-bottom: 20px;'>
                <img src="data:image/png;base64,{watermark_base64}" style="width: 440px; filter: drop-shadow(0px 4px 6px rgba(0,0,0,0.1));">
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
            <div style='text-align: center; margin-bottom: 20px;'>
                <h3 style='color: #C62828; margin: 0; font-weight: bold;'>👋 Halo!</h3>
                <p style='color: #555; font-size: 15px; margin-top: 5px;'>Silakan login untuk memulai aktivitas hari ini</p>
            </div>
        """, unsafe_allow_html=True)
        
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
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

# Tampilkan Login jika belum masuk
if not st.session_state.logged_in:
    login()
    st.stop() 

# --- 2. MENU UTAMA & SIDEBAR ---

# Header Utama di dalam Dashboard
st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; border-bottom: 3px solid #C62828; padding-bottom: 10px;">
        <div style="display: flex; align-items: center; gap: 15px;">
            <h1 style="margin: 0; color: #C62828; font-size: 32px; font-weight: bold;">SISTEM MANAJEMEN</h1>
        </div>
        <div style="text-align: right;">
            <span style="font-size: 16px; font-weight: bold;">👤 {st.session_state.role}</span>
        </div>
    </div>
""", unsafe_allow_html=True)

# NAVIGASI UTAMA DI SIDEBAR
st.sidebar.markdown("<h2 style='color: #C62828; margin-bottom: 10px;'>📌 NAVIGASI UTAMA</h2>", unsafe_allow_html=True)

menu_options = ["📊 Dashboard", "🧾 Nota", "🛍️ Penjualan", "📦 Stock", "💵 Finance", "⏱️ Absensi & Jadwal"]
selected_menu = st.sidebar.radio("", menu_options)

# SUB-MENU SLIDE DOWN BERDASARKAN MENU YANG DIPILIH
sub_menu = None
if selected_menu == "🧾 Nota":
    with st.sidebar.expander("📂 Sub-Menu Nota", expanded=True):
        sub_menu = st.radio("Pilih Halaman Nota:", ["📑 Bakul", "🏬 Bedak", "🤝 Mitra"])

elif selected_menu == "⏱️ Absensi & Jadwal":
    with st.sidebar.expander("📂 Sub-Menu Absensi", expanded=True):
        sub_menu = st.radio("Pilih Halaman Absen:", ["📅 Atur Jadwal", "📌 Plotting & Input Absen", "📊 Rekap Bulanan"])

st.sidebar.markdown("---")

# MASTER HARGA HANYA MUNCUL DI DALAM MENU NOTA
if selected_menu == "🧾 Nota":
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

    st.sidebar.markdown("<h3 style='color: #C62828;'>⚙️ Master Harga</h3>", unsafe_allow_html=True)
    h_glondong = st.sidebar.number_input("Harga Glondong", value=int(saved_harga.get("glondong", 28500)), step=500, format="%d")
    st.sidebar.caption(f"💡 Rp {h_glondong:,.0f}".replace(",", "."))
    h_jeroan = st.sidebar.number_input("Harga Jeroan", value=int(saved_harga.get("jeroan", 12000)), step=500, format="%d")
    st.sidebar.caption(f"💡 Rp {h_jeroan:,.0f}".replace(",", "."))
    h_usus = st.sidebar.number_input("Harga Usus", value=int(saved_harga.get("usus", 16500)), step=500, format="%d")
    st.sidebar.caption(f"💡 Rp {h_usus:,.0f}".replace(",", "."))
    h_telur_a = st.sidebar.number_input("Harga Telur A", value=int(saved_harga.get("telur_a", 269000)), step=1000, format="%d")
    st.sidebar.caption(f"💡 Rp {h_telur_a:,.0f}".replace(",", "."))
    h_telur_b = st.sidebar.number_input("Harga Telur B", value=int(saved_harga.get("telur_b", 250000)), step=1000, format="%d")
    st.sidebar.caption(f"💡 Rp {h_telur_b:,.0f}".replace(",", "."))
    h_peti = st.sidebar.number_input("Harga Peti", value=int(saved_harga.get("peti", 2000)), step=100, format="%d")
    st.sidebar.caption(f"💡 Rp {h_peti:,.0f}".replace(",", "."))
    h_box = st.sidebar.number_input("Harga Box", value=int(saved_harga.get("box", 28500)), step=500, format="%d")
    st.sidebar.caption(f"💡 Rp {h_box:,.0f}".replace(",", "."))

    current_harga = {"glondong": h_glondong, "jeroan": h_jeroan, "usus": h_usus, "telur_a": h_telur_a, "telur_b": h_telur_b, "peti": h_peti, "box": h_box}
    if current_harga != saved_harga:
        with open(FILE_HARGA, "w") as f:
            json.dump(current_harga, f)

    st.sidebar.markdown("---")

# INFORMASI USER & LOGOUT
st.sidebar.markdown(f"**👤 User:** {st.session_state.role}")
if st.sidebar.button("🚪 Keluar / Logout", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.role = ""
    st.rerun()


# --- 3. LOGIKA HALAMAN BERDASARKAN MENU & SUB-MENU ---

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
            
            # Cari baris header yang berisi kata NAMA
            header_idx = 0
            for idx, row in df_raw.iterrows():
                if row.astype(str).str.upper().str.contains('NAMA').any():
                    header_idx = idx
                    break
                    
            df = df_raw.iloc[header_idx+1:].copy()
            df.columns = [str(c).strip().upper() for c in df_raw.iloc[header_idx].values]
            
            # Deteksi kolom Nama dan Nomor (jika ada)
            name_col = next((c for c in df.columns if 'NAMA' in c), df.columns[1])
            no_col = next((c for c in df.columns if c in ['NO', 'NO.', 'NOMOR']), None)
            
            df = df[df[name_col].notna()].copy()
            
            # Buat list bakul berpenomoran
            bakul_options = []
            bakul_map = {}
            
            for idx, (_, row) in enumerate(df.iterrows(), start=1):
                nama_bakul = str(row[name_col]).strip()
                
                # Menggunakan nomor dari kolom NO Excel jika ada, jika tidak pakai penomoran otomatis 1,2,3...
                if no_col and pd.notna(row[no_col]) and str(row[no_col]).strip() != '':
                    no_val = str(row[no_col]).strip()
                    if no_val.endswith('.0'):
                        no_val = no_val[:-2]
                    label = f"{no_val}. {nama_bakul}"
                else:
                    label = f"{idx}. {nama_bakul}"
                    
                if label not in bakul_options:
                    bakul_options.append(label)
                    bakul_map[label] = nama_bakul

            selected_bakul_label = st.selectbox("Pilih Nama Bakul", bakul_options)
            selected_bakul = bakul_map.get(selected_bakul_label, "")

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
                    
                    # Layout footer nota yang dinaikkan posisinya
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
                    
                    st.markdown("<br><br>", unsafe_allow_html=True)
                else:
                    st.warning("Tidak ada item pembelian untuk bakul ini.")

    elif sub_menu == "🏬 Bedak":
        st.info("Fitur manajemen dan cetak nota untuk bedak sedang dalam tahap pengembangan.")
    elif sub_menu == "🤝 Mitra":
        st.info("Fitur rekapitulasi dan nota khusus untuk mitra akan dikembangkan di sini.")

elif selected_menu == "🛍️ Penjualan":
    st.info("Modul rekap penjualan harian dan bulanan.")
elif selected_menu == "📦 Stock":
    st.info("Modul untuk memantau barang masuk dan keluar (Glondong, Telur, dll).")
elif selected_menu == "💵 Finance":
    st.info("Modul pencatatan arus kas (Cash Flow) dan Piutang Bakul.")
elif selected_menu == "⏱️ Absensi & Jadwal":
    if sub_menu == "📅 Atur Jadwal" or sub_menu is None:
        st.info("Pengaturan Master Shift dan Jatah Libur akan dibuat di sini.")
    elif sub_menu == "📌 Plotting & Input Absen":
        st.info("Pencatatan jam masuk & jam keluar pegawai (Tepat Waktu/Terlambat) akan ada di sini.")
    elif sub_menu == "📊 Rekap Bulanan":
        st.info("Tabel hasil rekap absen bulanan.")