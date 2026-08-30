import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os
import base64
import streamlit.components.v1 as components

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Sistem Manajemen Ayam Segar", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Cek nama file logo yang ada di repository
logo_filename = None
for fname in ["ASTremove.PNG", "ASTremove.png", "AST.jpeg"]:
    if os.path.exists(fname):
        logo_filename = fname
        break

# Function Encode Gambar ke Base64 (supaya logo muncul di HTML Printable)
def get_base64_logo(file_path):
    if file_path and os.path.exists(file_path):
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

# --- 2. CSS STYLING UTAMA ---
st.markdown("""
    <style>
        footer, #MainMenu { visibility: hidden; }
        [data-testid="stHeader"] { background-color: transparent !important; z-index: 100 !important; }
        [data-testid="stSidebar"] { background-color: #FFEBEE !important; border-right: 3px solid #C62828 !important; }
        [data-testid="stSidebar"] .stRadio label { font-size: 15px !important; font-weight: bold !important; color: #262626 !important; }
        .block-container {
            background-color: #FFFFFF;
            padding-top: 1rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- 3. LOGIN ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = ""

def login():
    st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if logo_filename:
            st.image(logo_filename, width=240)
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

# --- 4. MASTER HARGA JSON ---
FILE_HARGA = "master_harga.json"
default_harga = {"glondong": 28500, "jeroan": 12000, "usus": 16500, "telur_a": 269000, "telur_b": 250000, "peti": 2000, "box": 28500}

if os.path.exists(FILE_HARGA):
    try:
        with open(FILE_HARGA, "r") as f: saved_harga = json.load(f)
    except: saved_harga = default_harga
else: saved_harga = default_harga

# --- 5. SIDEBAR ---
with st.sidebar:
    if logo_filename:
        st.image(logo_filename, width=90)
    st.markdown("<h2 style='color: #C62828; margin: 0; font-size: 26px; font-weight: bold;'>AST SYSTEM</h2>", unsafe_allow_html=True)
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

# --- 6. AREA KONTEN UTAMA ---
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
                    st.session_state['qty_box_val'] = 0
                    st.session_state['qty_telur_b_val'] = 0

                row_bakul = df.loc[real_idx]
                
                def get_valid_float(col_name):
                    try:
                        if not col_name or col_name not in df.columns: return 0.0
                        val = row_bakul[col_name]
                        num = pd.to_numeric(val, errors='coerce')
                        return 0.0 if pd.isna(num) else float(num)
                    except: return 0.0

                col_in1, col_in2, col_in3 = st.columns(3)
                with col_in1:
                    qty_peti = st.number_input("Jumlah Peti", value=0, step=1, format="%d")
                with col_in2:
                    qty_box = st.number_input("Jumlah Box (Manual)", key='qty_box_val', value=0, step=1, format="%d")
                with col_in3:
                    qty_telur_b_manual = st.number_input("Jumlah Telur B (Manual)", key='qty_telur_b_val', value=0, step=1, format="%d")

                qty_tonase = get_valid_float(next((c for c in df.columns if 'TONASE' in c), ''))
                qty_jeroan = get_valid_float(next((c for c in df.columns if 'JEROAN' in c), ''))
                qty_usus = get_valid_float(next((c for c in df.columns if 'USUS' in c), ''))
                qty_telur_a = get_valid_float(next((c for c in df.columns if 'TELUR A' in c or 'TELUR' in c), ''))
                
                qty_telur_b_excel = get_valid_float(next((c for c in df.columns if 'TELUR B' in c), ''))
                qty_telur_b = qty_telur_b_excel if qty_telur_b_excel > 0 else qty_telur_b_manual

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
                    # PREVIEW NOTA DI SCREEN
                    with st.container(border=True):
                        col_left, col_right = st.columns([7, 3])
                        logo_b64 = get_base64_logo(logo_filename)
                        
                        with col_left:
                            if logo_b64:
                                st.markdown(f"""
                                    <div style="display: flex; align-items: center; gap: 14px;">
                                        <img src="data:image/png;base64,{logo_b64}" style="width: 100px; height: auto;">
                                        <div>
                                            <h3 style="margin: 0; color: #C62828; font-weight: bold; font-size: 20px;">AYAM SEGAR TUMPANG</h3>
                                            <span style="color: #555; font-size: 13px;">Ds. Kambingan - Tumpang - Kab. Malang</span>
                                        </div>
                                    </div>
                                """, unsafe_allow_html=True)
                            else:
                                st.markdown("""
                                    <div>
                                        <h3 style="margin: 0; color: #C62828; font-weight: bold; font-size: 20px;">AYAM SEGAR TUMPANG</h3>
                                        <span style="color: #555; font-size: 13px;">Ds. Kambingan - Tumpang - Kab. Malang</span>
                                    </div>
                                """, unsafe_allow_html=True)

                        with col_right:
                            st.markdown(f"<div style='text-align:right; font-size: 13px;'><b>Tanggal:</b> {tanggal_transaksi.strftime('%d-%m-%Y')}<br><b>Bakul:</b> {selected_bakul}<br><b>Group:</b> {selected_sheet}</div>", unsafe_allow_html=True)

                        rows_html = ""
                        for item in filtered_items:
                            kg_str = f"{int(item['KG'])}" if isinstance(item['KG'], (int, float)) and float(item['KG']).is_integer() else f"{item['KG']:.2f}" if isinstance(item['KG'], float) else str(item['KG'])
                            rows_html += f"<tr><td style='text-align: center;'>{kg_str}</td><td>{item['Nama Barang']}</td><td style='text-align: right;'>Rp {item['Harga']:,.0f}</td><td style='text-align: right;'>Rp {item['Jumlah']:,.0f}</td></tr>".replace(",", ".")

                        st.markdown(f"""
                            <table style="width:100%; border-collapse:collapse; margin-top:15px; font-size:13px;">
                                <thead><tr><th style="border:1px solid #000; padding:6px; background:#f2f2f2;">QTY / KG</th><th style="border:1px solid #000; padding:6px; background:#f2f2f2;">BARANG</th><th style="border:1px solid #000; padding:6px; background:#f2f2f2;">HARGA</th><th style="border:1px solid #000; padding:6px; background:#f2f2f2;">JUMLAH</th></tr></thead>
                                <tbody>{rows_html}</tbody>
                            </table>
                            <div style="text-align:right; margin-top:10px; font-weight:bold;">
                                TOTAL: <span style="color:#C62828; font-size:17px;">Rp {total_bayar:,.0f}</span>
                            </div>
                        """.replace(",", "."), unsafe_allow_html=True)

                    # --- TOMBOL CETAK NOTA (HTML PRINTABLE) ---
                    # Kamu bebas custom margin (@page margin: 5mm/10mm) atau font size di CSS di bawah ini!
                    html_table_rows = ""
                    for item in filtered_items:
                        kg_str = f"{int(item['KG'])}" if isinstance(item['KG'], (int, float)) and float(item['KG']).is_integer() else f"{item['KG']:.2f}" if isinstance(item['KG'], float) else str(item['KG'])
                        html_table_rows += f"""
                            <tr>
                                <td style="text-align:center; border:1px solid #333; padding:4px;">{kg_str}</td>
                                <td style="border:1px solid #333; padding:4px;">{item['Nama Barang']}</td>
                                <td style="text-align:right; border:1px solid #333; padding:4px;">Rp {item['Harga']:,.0f}</td>
                                <td style="text-align:right; border:1px solid #333; padding:4px;">Rp {item['Jumlah']:,.0f}</td>
                            </tr>
                        """.replace(",", ".")

                    printable_html = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <style>
                            /* KAMU BISA EDIT MARGIN CETAK KERTAS DI SINI */
                            @page {{
                                size: A4 portrait;
                                margin: 8mm; /* Bebas ubah misal: 5mm, 10mm, 0mm */
                            }}
                            body {{
                                font-family: Arial, sans-serif;
                                font-size: 11px;
                                margin: 0;
                                padding: 0;
                                color: #000;
                            }}
                            .nota-box {{
                                border: 1.5px solid #000;
                                padding: 10px;
                                max-width: 750px;
                                margin: 0 auto;
                            }}
                            .header-table {{
                                width: 100%;
                                border-collapse: collapse;
                                margin-bottom: 8px;
                            }}
                            .items-table {{
                                width: 100%;
                                border-collapse: collapse;
                                margin-top: 6px;
                            }}
                            .items-table th {{
                                border: 1px solid #333;
                                background-color: #eee;
                                padding: 4px;
                                font-size: 11px;
                            }}
                            .footer-table {{
                                width: 100%;
                                margin-top: 15px;
                            }}
                            .btn-print {{
                                background-color: #C62828;
                                color: white;
                                padding: 10px 20px;
                                font-size: 15px;
                                font-weight: bold;
                                border: none;
                                border-radius: 5px;
                                cursor: pointer;
                                width: 100%;
                            }}
                            .btn-print:hover {{
                                background-color: #B71C1C;
                            }}
                            @media print {{
                                .no-print {{ display: none !important; }}
                                .nota-box {{ border: 1.5px solid #000 !important; }}
                            }}
                        </style>
                    </head>
                    <body>
                        <div class="no-print" style="margin-bottom: 10px;">
                            <button class="btn-print" onclick="window.print()">🖨️ Cetak / Print Nota (Bisa Disimpan Sebagai PDF)</button>
                        </div>
                        
                        <div class="nota-box">
                            <table class="header-table">
                                <tr>
                                    <td style="vertical-align: top; width: 60%;">
                                        <div style="display: flex; align-items: center; gap: 10px;">
                                            {"<img src='data:image/png;base64," + logo_b64 + "' style='height: 42px;' />" if logo_b64 else ""}
                                            <div>
                                                <div style="font-size: 15px; font-weight: bold; color: #C62828;">AYAM SEGAR TUMPANG</div>
                                                <div style="font-size: 10px; color: #555;">Ds. Kambingan - Tumpang - Kab. Malang</div>
                                            </div>
                                        </div>
                                    </td>
                                    <td style="vertical-align: top; text-align: right; width: 40%; font-size: 10.5px;">
                                        <b>Tanggal:</b> {tanggal_transaksi.strftime('%d-%m-%Y')}<br>
                                        <b>Bakul:</b> {selected_bakul}<br>
                                        <b>Group:</b> {selected_sheet}
                                    </td>
                                </tr>
                            </table>

                            <table class="items-table">
                                <thead>
                                    <tr>
                                        <th style="width: 15%;">QTY / KG</th>
                                        <th style="width: 45%; text-align: left;">BARANG</th>
                                        <th style="width: 20%; text-align: right;">HARGA</th>
                                        <th style="width: 20%; text-align: right;">JUMLAH</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {html_table_rows}
                                </tbody>
                            </table>

                            <table class="footer-table">
                                <tr>
                                    <td style="vertical-align: top; width: 50%;">
                                        <table style="width: 100%; text-align: center; font-size: 10px;">
                                            <tr>
                                                <td>Penerima,</td>
                                                <td>Hormat Kami,</td>
                                            </tr>
                                            <tr><td colspan="2" style="height: 35px;"></td></tr>
                                            <tr>
                                                <td>( ............................ )</td>
                                                <td>( ............................ )</td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="vertical-align: bottom; text-align: right; width: 50%;">
                                        <span style="font-size: 12px; font-weight: bold;">TOTAL : </span>
                                        <span style="font-size: 15px; font-weight: bold; color: #C62828;">Rp {total_bayar:,.0f}</span>
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </body>
                    </html>
                    """.replace(",", ".")

                    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
                    components.html(printable_html, height=260, scrolling=True)