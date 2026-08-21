import streamlit as st
import pandas as pd
from datetime import datetime
import base64

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

# --- CSS STYLING FIX ---
st.markdown(f"""
    <style>
        /* Sembunyikan footer dan elemen menu titik tiga bawaan, TAPI sembunyikan toolbar secara halus */
        footer, #MainMenu {{
            visibility: hidden;
        }}
        
        /* Pertahankan tombol toggle sidebar agar tetap bisa diklik */
        [data-testid="stHeader"] {{
            background-color: transparent !important;
            z-index: 100 !important;
        }}

        /* Styling Sidebar */
        [data-testid="stSidebar"] {{
            background-color: #FFEBEE !important;
            border-right: 3px solid #C62828 !important;
        }}
        
        /* Warna Teks Radio Sidebar */
        [data-testid="stSidebar"] .stRadio label {{
            font-size: 16px !important;
            font-weight: bold !important;
            color: #262626 !important;
        }}

        /* Watermark Background */
        .block-container {{
            background-color: #FFFFFF;
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 400px;
            padding-top: 1rem !important;
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


# --- 3. SIDEBAR NAVIGATION ---
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

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.role = ""
        st.rerun()


# --- 4. AREA KONTEN UTAMA ---
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"<h2 style='color: #C62828; margin:0;'>Ayam Segar Tumpang - {selected_menu}</h2>", unsafe_allow_html=True)
with col_h2:
    st.markdown(f"<div style='text-align: right; font-weight: bold; padding-top: 10px;'>👤 {st.session_state.role}</div>", unsafe_allow_html=True)

st.markdown("<hr style='border: 1px solid #C62828; margin-top: 5px; margin-bottom: 20px;'>", unsafe_allow_html=True)

if selected_menu == "📊 Dashboard":
    st.info("Visualisasi data omset, stok, dan performa harian akan ditampilkan di sini.")
elif selected_menu == "🧾 Nota":
    st.subheader(f"Menu Nota: {sub_menu if sub_menu else 'Bakul'}")
    st.file_uploader("Upload Rekap Transaksi (.xlsx)", type=["xlsx"])
elif selected_menu == "🛍️ Penjualan":
    st.info("Halaman Penjualan Harian & Bulanan.")
elif selected_menu == "📦 Stock":
    st.info("Halaman Inventori Stok Barang.")
elif selected_menu == "💵 Finance":
    st.info("Halaman Pencatatan Kas & Piutang.")
elif selected_menu == "⏱️ Absensi & Jadwal":
    st.subheader(f"Menu Absensi: {sub_menu if sub_menu else 'Atur Jadwal'}")