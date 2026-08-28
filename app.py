import streamlit as st
import pandas as pd
from datetime import datetime
import base64
import json
import os
from fpdf import FPDF

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Sistem Manajemen Ayam Segar", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- FUNGSI GAMBAR SAFE LOAD BASE64 ---
def get_img_as_base64(file_path):
    if file_path and os.path.exists(file_path):
        try:
            with open(file_path, "rb") as f:
                data = f.read()
            return base64.b64encode(data).decode()
        except Exception:
            return ""
    return ""

# Cari file logo di direktori aktif
logo_filename = None
for fname in ["ASTremove.PNG", "ASTremove.png", "logo.png", "logo.PNG", "AST.png", "AST.PNG"]:
    if os.path.exists(fname):
        logo_filename = fname
        break

watermark_base64 = get_img_as_base64(logo_filename) if logo_filename else ""

# --- CLASS GENERATOR PDF NOTA ---
class NotaPDF(FPDF):
    def __init__(self, logo_path=None):
        super().__init__(orientation='L', unit='mm', format='A5')
        self.logo_path = logo_path

    def generate(self, tgl, bakul, group, items, total_bayar):
        self.add_page()
        self.set_margins(12, 12, 12)
        
        # Border Luar Nota (Proporsional A5)
        self.rect(10, 10, 190, 128)

        # Header Logo & Judul
        if self.logo_path and os.path.exists(self.logo_path):
            self.image(self.logo_path, x=14, y=13, w=16)
            self.set_xy(32, 13)
        else:
            self.set_xy(14, 13)

        self.set_font("Helvetica", "B", 13)
        self.set_text_color(198, 40, 40)
        self.cell(90, 6, "AYAM SEGAR TUMPANG", ln=1)
        
        self.set_x(32 if (self.logo_path and os.path.exists(self.logo_path)) else 14)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(90, 90, 90)
        self.cell(90, 4, "Ds. Kambingan - Tumpang - Kab. Malang", ln=0)

        # Header Info Kanan
        self.set_xy(110, 13)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(0, 0, 0)
        self.cell(84, 4, f"Tanggal: {tgl}", ln=1, align='R')
        self.set_x(110)
        self.cell(84, 4, f"Pembeli / Bakul: {bakul}", ln=1, align='R')
        self.set_x(110)
        self.cell(84, 4, f"Group: {group}", ln=1, align='R')

        self.ln(6)

        # Tabel Header
        self.set_x(13)
        self.set_font("Helvetica", "B", 8.5)
        self.set_fill_color(240, 240, 240)
        self.cell(24, 7, "QTY / KG", 1, 0, 'C', fill=True)
        self.cell(80, 7, "BARANG", 1, 0, 'L', fill=True)
        self.cell(40, 7, "HARGA  ", 1, 0, 'R', fill=True)
        self.cell(40, 7, "JUMLAH  ", 1, 1, 'R', fill=True)

        # Isi Tabel
        self.set_font("Helvetica", "", 8.5)
        for item in items:
            self.set_x(13)
            kg_val = item['KG']
            kg_str = f"{int(kg_val)}" if isinstance(kg_val, float) and kg_val.is_integer() else f"{kg_val:.2f}" if isinstance(kg_val, float) else str(kg_val)
            h_str = f"Rp {item['Harga']:,.0f}  ".replace(",", ".")
            j_str = f"Rp {item['Jumlah']:,.0f}  ".replace(",", ".")

            self.cell(24, 7, kg_str, 1, 0, 'C')
            self.cell(80, 7, f" {item['Nama Barang']}", 1, 0, 'L')
            self.cell(40, 7, h_str, 1, 0, 'R')
            self.cell(40, 7, j_str, 1, 1, 'R')

        # Footer Area
        self.ln(8)
        y_footer = self.get_y()

        # Tanda Tangan
        self.set_xy(18, y_footer)
        self.cell(45, 4, "Penerima,", 0, 0, 'C')
        self.cell(45, 4, "Hormat Kami,", 0, 0, 'C')

        self.set_xy(18, y_footer + 18)
        self.cell(45, 4, "( ............................ )", 0, 0, 'C')
        self.cell(45, 4, "( ............................ )", 0, 0, 'C')

        # Total Kanan
        self.set_xy(110, y_footer + 2)
        self.set_font("Helvetica", "B", 10)
        self.cell(35, 8, "TOTAL :", 0, 0, 'R')
        
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(198, 40, 40)
        tot_str = f"Rp {total_bayar:,.0f}  ".replace(",", ".")
        self.cell(52, 8, tot_str, 0, 1, 'R')

        pdf_output = self.output()
        if isinstance(pdf_output, str):
            return pdf_output.encode('latin1')
        return bytes(pdf_output)

# --- CSS STYLING UTAMA ---
st.markdown(f"""
    <style>
        footer, #MainMenu {{ visibility: hidden; }}
        [data-testid="stHeader"] {{ background-color: transparent !important; z-index: 100 !important; }}
        [data-testid="stSidebar"] {{ background-color: #FFEBEE !important; border-right: 3px solid #C62828 !important; }}
        [data-testid="stSidebar"] .stRadio label {{ font-size: 15px !important; font-weight: bold !important; color: #262626 !important; }}
        .block-container {{
            background-color: #FFFFFF;
            {"background-image: linear-gradient(rgba(255, 255, 255, 0.95), rgba(255, 255, 255, 0.95)), url('data:image/png;base64," + watermark_base64 + "');" if watermark_base64 else ""}
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 400px;
            padding-top: 1rem !important;
        }}
        .nota-preview {{
            border: 1px solid #333;
            padding: 15px;
            background-color: #fff;
            border-radius: 6px;
            margin-bottom: 15px;
        }}
        .nota-preview table {{ width: 100%; border-collapse: collapse; }}
        .nota-preview th, .nota-preview td {{ border: 1px solid #000; padding: 6px; font-size: 13px; }}
        .nota-preview th {{ background-color: #f2f2f2; text-align: center; }}
    </style>
""", unsafe_allow_html=True)

# --- 2. LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = ""

def login():
    st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if watermark_base64:
            st.markdown(f"<div style='text-align: center; margin-bottom: 15px;'><img src='data:image/png;base64,{watermark_base64}' style='width: 220px;'></div>", unsafe_allow_html=True)
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
        with open(FILE_HARGA, "r") as f: saved_harga = json.load(f)
    except: saved_harga = default_harga
else: saved_harga = default_harga

# --- 4. SIDEBAR ---
with st.sidebar:
    logo_html = f"<img src='data:image/png;base64,{watermark_base64}' style='width: 65px; height: 65px; object-fit: contain;'>" if watermark_base64 else ""
    st.markdown(f"""
        <div style='display: flex; align-items: center; gap: 12px; margin-bottom: 10px;'>
            {logo_html}
            <h2 style='color: #C62828; margin: 0; font-size: 26px; font-weight: bold;'>AST SYSTEM</h2>
        </div>
    """, unsafe_allow_html=True)
    st.write(f"Logged in as: **{st.session_state.role}**")
    st.markdown("---")
    
    st.markdown("### 📌 NAVIGASI UTAMA")
    menu_options = ["📊 Dashboard", "🧾 Nota", "🛍️ Penjualan", "📦 Stock", "💵 Finance", "⏱️ Absensi & Jadwal"]
    selected_menu = st.radio("Pilih Halaman:", menu_options)

    sub_menu = None
    if selected_menu == "🧾 Nota":
        with st.expander("📂 Sub-Menu Nota", expanded=True):
            sub_menu = st.radio("Tipe Nota:", ["📑 Bakul", "🏬 Bedak", "🤝 Mitra"])

    if selected_menu == "🧾 Nota":
        st.markdown("---")
        st.markdown("<h3 style='color: #C62828;'>⚙️ Master Harga</h3>", unsafe_allow_html=True)
        
        def input_harga(label, key_name, default_val, step_val):
            val = st.number_input(f"{label}", value=int(saved_harga.get(key_name, default_val)), step=step_val, format="%d")
            st.caption(f"➔ **Rp {val:,.0f}**".replace(",", "."))
            return val

        h_glondong = input_harga("Harga Glondong", "glondong", 28500, 500)
        h_jeroan = input_harga("Harga Jeroan", "jeroan", 12000, 500)
        h_usus = input_harga("Harga Usus", "usus", 16500, 500)
        h_telur_a = input_harga("Harga Telur A", "telur_a", 269000, 1000)
        h_telur_b = input_harga("Harga Telur B", "telur_b", 250000, 1000)
        h_peti = input_harga("Harga Peti", "peti", 2000, 100)
        h_box = input_harga("Harga Box", "box", 28500, 500)

        current_harga = {"glondong": h_glondong, "jeroan": h_jeroan, "usus": h_usus, "telur_a": h_telur_a, "telur_b": h_telur_b, "peti": h_peti, "box": h_box}
        if current_harga != saved_harga:
            with open(FILE_HARGA, "w") as f: json.dump(current_harga, f)

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

if selected_menu == "🧾 Nota":
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
                if not nama_bakul: continue
                
                if no_col and pd.notna(row[no_col]) and str(row[no_col]).strip() != '':
                    no_val = str(row[no_col]).strip()
                    if no_val.endswith('.0'): no_val = no_val[:-2]
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
                        if not col_name or col_name not in df.columns: return 0.0
                        val = row_bakul[col_name]
                        num = pd.to_numeric(val, errors='coerce')
                        return 0.0 if pd.isna(num) else float(num)
                    except: return 0.0

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
                    # PREVIEW KERTAS NOTA DI WEB
                    rows_html = ""
                    for item in filtered_items:
                        kg_str = f"{int(item['KG'])}" if isinstance(item['KG'], float) and item['KG'].is_integer() else f"{item['KG']:.2f}" if isinstance(item['KG'], float) else str(item['KG'])
                        rows_html += f"<tr><td style='text-align: center;'>{kg_str}</td><td>{item['Nama Barang']}</td><td style='text-align: right;'>Rp {item['Harga']:,.0f}</td><td style='text-align: right;'>Rp {item['Jumlah']:,.0f}</td></tr>".replace(",", ".")

                    # CEK & RENDER LOGO HANYA JIKA FILE LOGO BENAR-BENAR ADA
                    img_tag = f'<img src="data:image/png;base64,{watermark_base64}" style="width: 50px; height: 50px; object-fit: contain;">' if watermark_base64 else ''
                    
                    st.markdown(f"""
                        <div class="nota-preview">
                            <table style="border:none; margin-bottom:10px;">
                                <tr>
                                    <td style="border:none;">
                                        <div style="display:flex; align-items:center; gap:10px;">
                                            {img_tag}
                                            <div>
                                                <h4 style="margin:0; color:#C62828;">AYAM SEGAR TUMPANG</h4>
                                                <small>Ds. Kambingan - Tumpang - Kab. Malang</small>
                                            </div>
                                        </div>
                                    </td>
                                    <td style="border:none; text-align:right;">
                                        <small><b>Tanggal:</b> {tanggal_transaksi.strftime("%d-%m-%Y")}<br>
                                        <b>Bakul:</b> {selected_bakul}<br>
                                        <b>Group:</b> {selected_sheet}</small>
                                    </td>
                                </tr>
                            </table>
                            <table>
                                <thead><tr><th>QTY / KG</th><th>BARANG</th><th>HARGA</th><th>JUMLAH</th></tr></thead>
                                <tbody>{rows_html}</tbody>
                            </table>
                            <div style="text-align:right; margin-top:10px; font-weight:bold;">
                                TOTAL: <span style="color:#C62828; font-size:18px;">Rp {total_bayar:,.0f}</span>
                            </div>
                        </div>
                    """.replace(",", "."), unsafe_allow_html=True)

                    # GENERATE PDF & TOMBOL DOWNLOAD
                    pdf_generator = NotaPDF(logo_path=logo_filename)
                    pdf_data = pdf_generator.generate(
                        tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                        bakul=selected_bakul,
                        group=selected_sheet,
                        items=filtered_items,
                        total_bayar=total_bayar
                    )

                    file_name = f"Nota_{selected_bakul}_{tanggal_transaksi.strftime('%Y%m%d')}.pdf"
                    
                    st.download_button(
                        label="📄 Download / Print PDF Nota",
                        data=pdf_data,
                        file_name=file_name,
                        mime="application/pdf",
                        use_container_width=True
                    )