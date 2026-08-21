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
    initial_sidebar_state="expanded" # Memaksa sidebar langsung terbuka saat di-load
)

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
logout_base64 = get_img_as_base64("logout.png")

if logout_base64:
    wrapper_width = "42px"
    wrapper_width_mobile = "36px"
    bg_css = f"background-image: url('data:image/png;base64,{logout_base64}'); background-color: transparent !important;"
    p_css = "display: none !important;"
else:
    wrapper_width = "auto"
    wrapper_width_mobile = "auto"
    bg_css = "background-color: #FFFFFF !important; padding: 0 12px !important;"
    p_css = "color: #C62828 !important; font-weight: bold;"

# --- CSS FIXING SIDEBAR & NAVBAR ---
st.markdown(f"""
    <style>
        /* Sembunyikan Header Bawaan Streamlit & Footer */
        header[data-testid="stHeader"] {{
            background: transparent !important;
            z-index: 100 !important;
        }}
        [data-testid="stToolbar"], #MainMenu, footer {{
            display: none !important;
        }}

        /* ========================================================== */
        /* PERBAIKAN SIDEBAR Supaya MUNCUL & BISA DIKLIK              */
        /* ========================================================== */
        [data-testid="stSidebar"] {{
            background-color: #FFEBEE !important;
            border-right: 3px solid #C62828 !important;
            z-index: 999999 !important; /* Memastikan sidebar di atas segalanya */
            top: 0px !important;
        }}
        
        /* Tombol Buka/Tutup Sidebar Bawaan */
        [data-testid="stSidebarCollapseButton"], 
        [data-testid="stSidebarExpandButton"],
        button[aria-label="Toggle sidebar"] {{
            color: white !important;
            z-index: 1000000 !important;
        }}

        /* Custom Navbar Merah */
        .custom-navbar {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 60px;
            background-color: #C62828;
            z-index: 999900;
            display: flex;
            align-items: center;
            padding: 0 20px 0 80px; /* Space kiri untuk icon sidebar */
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }}
        
        .navbar-brand {{
            display: flex;
            align-items: center;
        }}

        .custom-navbar img {{
            height: 40px;
            margin-right: 12px;
        }}
        .custom-navbar span.brand-text {{
            color: white;
            font-size: 20px;
            font-weight: bold;
        }}

        .user-status-text {{
            position: absolute;
            right: 80px;
            color: white;
            font-weight: bold;
            font-size: 14px;
        }}

        /* Position Tombol Logout */
        div.element-container:has(#logout-target) + div.element-container {{
            position: fixed;
            top: 9px;
            right: 18px;
            z-index: 999995;
            width: {wrapper_width};
        }}
        
        div.element-container:has(#logout-target) + div.element-container button {{
            {bg_css}
            background-size: contain;
            background-position: center;
            background-repeat: no-repeat;
            border: none !important;
            border-radius: 8px !important;
            height: 42px !important;
            min-height: 42px !important;
            width: 100% !important;
            cursor: pointer;
        }}
        
        div.element-container:has(#logout-target) + div.element-container button p {{
            {p_css}
        }}

        /* Layout Konten Utama */
        .block-container {{
            margin-top: 50px;
            background-color: #FFFFFF;
            background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), 
                              url("data:image/png;base64,{watermark_base64}");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 400px;
            padding: 2rem !important;
        }}

        /* Styling Navigasi Radio di Sidebar */
        [data-testid="stSidebar"] .stRadio label {{
            font-size: 16px !important;
            font-weight: bold !important;
            color: #262626 !important;
        }}
    </style>
""", unsafe_allow_html=True)

# --- 2. SISTEM LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = ""

def login():
    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(f"""
            <div style='text-align: center; margin-bottom: 15px;'>
                <img src="data:image/png;base64,{watermark_base64}" style="width: 250px;">
            </div>
        """, unsafe_allow_html=True)
        
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


# --- 3. TOPBAR NAVBAR & SIDEBAR MENU ---

# Header Navbar
st.markdown(f"""
    <div class="custom-navbar">
        <div class="navbar-brand">
            <img src="data:image/png;base64,{watermark_base64}">
            <span class="brand-text">Ayam Segar Tumpang</span>
        </div>
        <div class="user-status-text">
            👤 {st.session_state.role}
        </div>
    </div>
""", unsafe_allow_html=True)

# Logout Target Marker
st.markdown('<div id="logout-target"></div>', unsafe_allow_html=True)
if st.button("Logout", key="top_logout"):
    st.session_state.logged_in = False
    st.session_state.role = ""
    st.rerun()

# --- MODUL NAVIGASI SIDEBAR ---
st.sidebar.markdown("<h2 style='color: #C62828;'>📌 NAVIGASI</h2>", unsafe_allow_html=True)

menu_options = ["📊 Dashboard", "🧾 Nota", "🛍️ Penjualan", "📦 Stock", "💵 Finance", "⏱️ Absensi & Jadwal"]
selected_menu = st.sidebar.radio("Pilih Menu:", menu_options)

sub_menu = None
if selected_menu == "🧾 Nota":
    with st.sidebar.expander("📂 Sub-Menu Nota", expanded=True):
        sub_menu = st.radio("Tipe Nota:", ["📑 Bakul", "🏬 Bedak", "🤝 Mitra"])

elif selected_menu == "⏱️ Absensi & Jadwal":
    with st.sidebar.expander("📂 Sub-Menu Absensi", expanded=True):
        sub_menu = st.radio("Tipe Absen:", ["📅 Atur Jadwal", "📌 Plotting & Input Absen", "📊 Rekap Bulanan"])


# --- 4. HALAMAN UTAMA ---

st.markdown(f"""
    <div style="margin-bottom: 20px; border-bottom: 3px solid #C62828; padding-bottom: 5px;">
        <h1 style="margin: 0; color: #C62828; font-size: 26px; font-weight: bold;">SISTEM MANAJEMEN</h1>
    </div>
""", unsafe_allow_html=True)

if selected_menu == "📊 Dashboard":
    st.info("Visualisasi data omset, stok, dan performa harian akan ditampilkan di sini.")
elif selected_menu == "🧾 Nota":
    st.subheader(f"Halaman Nota - {sub_menu if sub_menu else 'Bakul'}")
    st.write("Silakan upload data transaksi Excel milikmu.")
elif selected_menu == "🛍️ Penjualan":
    st.info("Halaman Penjualan Harian & Bulanan.")
elif selected_menu == "📦 Stock":
    st.info("Halaman Inventori Stok Barang.")
elif selected_menu == "💵 Finance":
    st.info("Halaman Pencatatan Kas & Piutang.")
elif selected_menu == "⏱️ Absensi & Jadwal":
    st.subheader(f"Halaman Absensi - {sub_menu if sub_menu else 'Atur Jadwal'}")