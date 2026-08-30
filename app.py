import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from PIL import Image, ImageDraw, ImageFont

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

# --- FUNGSI GENERATOR WORD (.DOCX) ---
def generate_word_nota(tgl, bakul, group, items, total_bayar, logo_path):
    doc = Document()

    # Atur Margin Halaman Word
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

    # Tabel Utama Bingkai Nota
    outer_table = doc.add_table(rows=1, cols=1)
    outer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = outer_table.cell(0, 0)
    
    # Atur border bingkai luar
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = parse_xml(r'''
        <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="12" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="12" w:space="0" w:color="000000"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)

    # Header Table (Logo + Info)
    header_table = cell.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False

    # Kolom Kiri Header (Logo & Judul)
    cell_left = header_table.cell(0, 0)
    p_hdr = cell_left.paragraphs[0]
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(0)
    
    if logo_path and os.path.exists(logo_path):
        run_img = p_hdr.add_run()
        run_img.add_picture(logo_path, width=Inches(0.9))
        run_space = p_hdr.add_run("  ")

    run_title = p_hdr.add_run("AYAM SEGAR TUMPANG\n")
    run_title.bold = True
    run_title.font.size = Pt(12)
    run_title.font.color.rgb = RGBColor(198, 40, 40)

    run_sub = p_hdr.add_run("Ds. Kambingan - Tumpang - Kab. Malang")
    run_sub.font.size = Pt(8.5)
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    # Kolom Kanan Header (Tanggal & Bakul)
    cell_right = header_table.cell(0, 1)
    p_info = cell_right.paragraphs[0]
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_info.paragraph_format.space_before = Pt(0)
    p_info.paragraph_format.space_after = Pt(0)

    run_info = p_info.add_run(f"Tanggal: {tgl}\nPembeli / Bakul: {bakul}\nGroup: {group}")
    run_info.font.size = Pt(9)

    # Spacing
    p_space = cell.add_paragraph()
    p_space.paragraph_format.space_before = Pt(6)
    p_space.paragraph_format.space_after = Pt(6)

    # Tabel Barang
    item_table = cell.add_table(rows=1, cols=4)
    item_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells = item_table.rows[0].cells
    headers = ["QTY / KG", "BARANG", "HARGA", "JUMLAH"]
    aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT]
    
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        p = hdr_cells[i].paragraphs[0]
        p.alignment = aligns[i]
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(8.5)
        # Background Abu Header Tabel
        shading = parse_xml(r'<w:shd {} w:fill="F0F0F0"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

    for item in items:
        row_cells = item_table.add_row().cells
        kg_str = f"{int(item['KG'])}" if isinstance(item['KG'], (int, float)) and float(item['KG']).is_integer() else f"{item['KG']:.2f}" if isinstance(item['KG'], float) else str(item['KG'])
        h_str = f"Rp {item['Harga']:,.0f}".replace(",", ".")
        j_str = f"Rp {item['Jumlah']:,.0f}".replace(",", ".")
        
        vals = [kg_str, item['Nama Barang'], h_str, j_str]
        for i, val in enumerate(vals):
            row_cells[i].text = val
            p = row_cells[i].paragraphs[0]
            p.alignment = aligns[i]
            p.runs[0].font.size = Pt(8.5)

    # Apply Border ke Seluruh Sel Tabel Barang
    for row in item_table.rows:
        for c in row.cells:
            tcPr = c._element.get_or_add_tcPr()
            tcBorders = parse_xml(r'''
                <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tcBorders>
            ''')
            tcPr.append(tcBorders)

    # Footer Table (Tanda Tangan & Total)
    footer_table = cell.add_table(rows=1, cols=2)
    footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Kiri Footer (Tanda Tangan)
    c_ft_left = footer_table.cell(0, 0)
    p_ft_l = c_ft_left.paragraphs[0]
    p_ft_l.paragraph_format.space_before = Pt(10)
    p_ft_l.add_run("Penerima,\t\tHormat Kami,\n\n\n( ............................ )\t( ............................ )").font.size = Pt(8.5)

    # Kanan Footer (Total)
    c_ft_right = footer_table.cell(0, 1)
    p_ft_r = c_ft_right.paragraphs[0]
    p_ft_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ft_r.paragraph_format.space_before = Pt(25)
    
    run_tot_lbl = p_ft_r.add_run("TOTAL : ")
    run_tot_lbl.bold = True
    run_tot_lbl.font.size = Pt(9.5)
    
    run_tot_val = p_ft_r.add_run(f"Rp {total_bayar:,.0f}".replace(",", "."))
    run_tot_val.bold = True
    run_tot_val.font.size = Pt(11)
    run_tot_val.font.color.rgb = RGBColor(198, 40, 40)

    # Output Buffer
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

# --- FUNGSI GENERATOR GAMBAR NOTA (.PNG) ---
def generate_image_nota(tgl, bakul, group, items, total_bayar, logo_path):
    width, height = 750, 450
    img = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(img)
    
    try:
        font_title = ImageFont.truetype("arial.ttf", 20)
        font_bold = ImageFont.truetype("arial.ttf", 13)
        font_regular = ImageFont.truetype("arial.ttf", 12)
        font_small = ImageFont.truetype("arial.ttf", 10)
    except:
        font_title = font_bold = font_regular = font_small = ImageFont.load_default()

    # Border Luar
    draw.rectangle([(10, 10), (width - 10, height - 10)], outline="black", width=2)
    
    # Logo
    x_offset = 20
    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path).convert("RGBA")
            logo_img.thumbnail((70, 70))
            img.paste(logo_img, (20, 20), logo_img)
            x_offset = 100
        except: pass

    # Header Teks
    draw.text((x_offset, 25), "AYAM SEGAR TUMPANG", fill="#C62828", font=font_title)
    draw.text((x_offset, 52), "Ds. Kambingan - Tumpang - Kab. Malang", fill="#555555", font=font_small)

    # Info Kanan
    draw.text((width - 200, 25), f"Tanggal: {tgl}", fill="black", font=font_small)
    draw.text((width - 200, 40), f"Bakul: {bakul}", fill="black", font=font_small)
    draw.text((width - 200, 55), f"Group: {group}", fill="black", font=font_small)

    # Tabel Items
    y_table = 90
    draw.rectangle([(20, y_table), (width - 20, y_table + 25)], fill="#F0F0F0", outline="black")
    draw.text((30, y_table + 5), "QTY / KG", fill="black", font=font_bold)
    draw.text((150, y_table + 5), "BARANG", fill="black", font=font_bold)
    draw.text((450, y_table + 5), "HARGA", fill="black", font=font_bold)
    draw.text((600, y_table + 5), "JUMLAH", fill="black", font=font_bold)

    y_curr = y_table + 25
    for item in items:
        draw.rectangle([(20, y_curr), (width - 20, y_curr + 22)], outline="black")
        kg_str = f"{int(item['KG'])}" if isinstance(item['KG'], (int, float)) and float(item['KG']).is_integer() else f"{item['KG']:.2f}" if isinstance(item['KG'], float) else str(item['KG'])
        h_str = f"Rp {item['Harga']:,.0f}".replace(",", ".")
        j_str = f"Rp {item['Jumlah']:,.0f}".replace(",", ".")

        draw.text((30, y_curr + 3), kg_str, fill="black", font=font_regular)
        draw.text((150, y_curr + 3), item['Nama Barang'], fill="black", font=font_regular)
        draw.text((450, y_curr + 3), h_str, fill="black", font=font_regular)
        draw.text((600, y_curr + 3), j_str, fill="black", font=font_regular)
        y_curr += 22

    # Footer Sign
    draw.text((40, height - 80), "Penerima,", fill="black", font=font_small)
    draw.text((180, height - 80), "Hormat Kami,", fill="black", font=font_small)
    draw.text((40, height - 30), "( ............................ )", fill="black", font=font_small)
    draw.text((180, height - 30), "( ............................ )", fill="black", font=font_small)

    # Total
    tot_str = f"TOTAL: Rp {total_bayar:,.0f}".replace(",", ".")
    draw.text((width - 240, height - 40), tot_str, fill="#C62828", font=font_title)

    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

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
                    # PREVIEW NOTA
                    with st.container(border=True):
                        col_left, col_right = st.columns([7, 3])
                        with col_left:
                            if logo_filename:
                                st.image(logo_filename, width=80)
                            st.markdown("<h3 style='margin: 0; color: #C62828;'>AYAM SEGAR TUMPANG</h3><span style='color: #555; font-size: 13px;'>Ds. Kambingan - Tumpang - Kab. Malang</span>", unsafe_allow_html=True)
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

                    # TOMBOL DOWNLOAD PILIHAN (WORD / GAMBAR PNG)
                    col_btn1, col_btn2 = st.columns(2)

                    # 1. Download File Word (.docx)
                    docx_bytes = generate_word_nota(
                        tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                        bakul=selected_bakul,
                        group=selected_sheet,
                        items=filtered_items,
                        total_bayar=total_bayar,
                        logo_path=logo_filename
                    )
                    with col_btn1:
                        st.download_button(
                            label="📝 Download File Word (.docx)",
                            data=docx_bytes,
                            file_name=f"Nota_{selected_bakul}_{tanggal_transaksi.strftime('%Y%m%d')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

                    # 2. Download File Gambar (.png)
                    img_bytes = generate_image_nota(
                        tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                        bakul=selected_bakul,
                        group=selected_sheet,
                        items=filtered_items,
                        total_bayar=total_bayar,
                        logo_path=logo_filename
                    )
                    with col_btn2:
                        st.download_button(
                            label="🖼️ Download Gambar Nota (.png)",
                            data=img_bytes,
                            file_name=f"Nota_{selected_bakul}_{tanggal_transaksi.strftime('%Y%m%d')}.png",
                            mime="image/png",
                            use_container_width=True
                        )