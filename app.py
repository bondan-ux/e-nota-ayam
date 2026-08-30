import streamlit as st
import pandas as pd
import io
import os
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- CONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="E-Nota Bakul Ayam Segar",
    page_icon="🐔",
    layout="wide"
)

# --- FUNGSI GENERATOR GAMBAR NOTA (.PNG) - DYNAMIC HEIGHT ---
def generate_image_nota(tgl, bakul, group, items, total_bayar, logo_path):
    width = 750
    # Hitung tinggi dinamis berdasarkan jumlah item
    table_header_h = 30
    row_h = 24
    footer_h = 90
    top_header_h = 85
    
    total_table_h = table_header_h + (len(items) * row_h)
    height = top_header_h + total_table_h + footer_h + 30 # total tinggi dinamis
    
    img = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(img)
    
    try:
        font_title = ImageFont.truetype("arial.ttf", 18)
        font_bold = ImageFont.truetype("arial.ttf", 12)
        font_regular = ImageFont.truetype("arial.ttf", 11)
        font_small = ImageFont.truetype("arial.ttf", 10)
    except:
        font_title = font_bold = font_regular = font_small = ImageFont.load_default()

    # Border Luar Bingkai Nota
    draw.rectangle([(10, 10), (width - 10, height - 10)], outline="black", width=2)
    
    # Logo & Header Title
    x_offset = 20
    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path).convert("RGBA")
            logo_img.thumbnail((60, 60))
            img.paste(logo_img, (20, 18), logo_img)
            x_offset = 90
        except: pass

    # Header Teks Kiri
    draw.text((x_offset, 20), "AYAM SEGAR TUMPANG", fill="#C62828", font=font_title)
    draw.text((x_offset, 45), "Ds. Kambingan - Tumpang - Kab. Malang", fill="#555555", font=font_small)

    # Info Kanan
    draw.text((width - 210, 20), f"Tanggal: {tgl}", fill="black", font=font_small)
    draw.text((width - 210, 35), f"Bakul: {bakul}", fill="black", font=font_small)
    draw.text((width - 210, 50), f"Group: {group}", fill="black", font=font_small)

    # Tabel Items Header
    y_table = 80
    draw.rectangle([(20, y_table), (width - 20, y_table + 24)], fill="#F0F0F0", outline="black")
    draw.text((30, y_table + 5), "QTY / KG", fill="black", font=font_bold)
    draw.text((150, y_table + 5), "BARANG", fill="black", font=font_bold)
    draw.text((450, y_table + 5), "HARGA", fill="black", font=font_bold)
    draw.text((600, y_table + 5), "JUMLAH", fill="black", font=font_bold)

    # Isi Baris Tabel
    y_curr = y_table + 24
    for item in items:
        draw.rectangle([(20, y_curr), (width - 20, y_curr + row_h)], outline="black")
        kg_str = f"{int(item['KG'])}" if isinstance(item['KG'], (int, float)) and float(item['KG']).is_integer() else f"{item['KG']:.2f}" if isinstance(item['KG'], float) else str(item['KG'])
        h_str = f"Rp {item['Harga']:,.0f}".replace(",", ".")
        j_str = f"Rp {item['Jumlah']:,.0f}".replace(",", ".")

        draw.text((30, y_curr + 4), kg_str, fill="black", font=font_regular)
        draw.text((150, y_curr + 4), item['Nama Barang'], fill="black", font=font_regular)
        draw.text((450, y_curr + 4), h_str, fill="black", font=font_regular)
        draw.text((600, y_curr + 4), j_str, fill="black", font=font_regular)
        y_curr += row_h

    # Footer Area (Langsung Rapat di Bawah Tabel)
    y_footer = y_curr + 15
    draw.text((40, y_footer), "Penerima,", fill="black", font=font_small)
    draw.text((180, y_footer), "Hormat Kami,", fill="black", font=font_small)
    
    draw.text((40, y_footer + 45), "( ............................ )", fill="black", font=font_small)
    draw.text((180, y_footer + 45), "( ............................ )", fill="black", font=font_small)

    # Total Kanan Bawah
    tot_str = f"TOTAL: Rp {total_bayar:,.0f}".replace(",", ".")
    draw.text((width - 240, y_footer + 35), tot_str, fill="#C62828", font=font_title)

    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()


# --- FUNGSI GENERATOR WORD (.DOCX) ---
def generate_word_nota(tgl, bakul, group, items, total_bayar, logo_path):
    doc = Document()

    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

    outer_table = doc.add_table(rows=1, cols=1)
    outer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = outer_table.cell(0, 0)
    
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

    # Header
    header_table = cell.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_left = header_table.cell(0, 0)
    p_hdr = cell_left.paragraphs[0]
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(0)
    
    if logo_path and os.path.exists(logo_path):
        try:
            run_img = p_hdr.add_run()
            run_img.add_picture(logo_path, width=Inches(0.8))
            p_hdr.add_run("  ")
        except: pass

    run_title = p_hdr.add_run("AYAM SEGAR TUMPANG\n")
    run_title.bold = True
    run_title.font.size = Pt(11)
    run_title.font.color.rgb = RGBColor(198, 40, 40)

    run_sub = p_hdr.add_run("Ds. Kambingan - Tumpang - Kab. Malang")
    run_sub.font.size = Pt(8)
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    cell_right = header_table.cell(0, 1)
    p_info = cell_right.paragraphs[0]
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_info.paragraph_format.space_before = Pt(0)
    p_info.paragraph_format.space_after = Pt(0)

    run_info = p_info.add_run(f"Tanggal: {tgl}\nBakul: {bakul}\nGroup: {group}")
    run_info.font.size = Pt(8.5)

    p_space = cell.add_paragraph()
    p_space.paragraph_format.space_before = Pt(4)
    p_space.paragraph_format.space_after = Pt(4)

    # Item Table
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

    # Footer Table
    footer_table = cell.add_table(rows=1, cols=2)
    footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    c_ft_left = footer_table.cell(0, 0)
    p_ft_l = c_ft_left.paragraphs[0]
    p_ft_l.paragraph_format.space_before = Pt(6)
    p_ft_l.add_run("Penerima,\t\tHormat Kami,\n\n\n( ............................ )\t( ............................ )").font.size = Pt(8)

    c_ft_right = footer_table.cell(0, 1)
    p_ft_r = c_ft_right.paragraphs[0]
    p_ft_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ft_r.paragraph_format.space_before = Pt(18)
    
    run_tot_lbl = p_ft_r.add_run("TOTAL : ")
    run_tot_lbl.bold = True
    run_tot_lbl.font.size = Pt(9)
    
    run_tot_val = p_ft_r.add_run(f"Rp {total_bayar:,.0f}".replace(",", "."))
    run_tot_val.bold = True
    run_tot_val.font.size = Pt(10.5)
    run_tot_val.font.color.rgb = RGBColor(198, 40, 40)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()


# --- INTERFACE UTAMA STREAMLIT ---
st.title("🐔 E-Nota Bakul Ayam Segar")
st.caption("Aplikasi pembuat nota digital praktis dari file Excel penjualan harian.")

col_input, col_preview = st.columns([1, 1])

with col_input:
    st.subheader("1. Upload File Excel")
    uploaded_file = st.file_uploader("Pilih file Excel (.xlsx / .xls)", type=["xlsx", "xls"])
    
    logo_path = "logo.png" # Lokasi default logo jika ada
    
    if uploaded_file:
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names
            
            st.subheader("2. Pilih Sheet & Bakul")
            selected_sheet = st.selectbox("Pilih Sheet (Tanggal/Group):", sheet_names)
            
            # Read sheet data
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            # Kolom parsing (menyesuaikan struktur file Excel)
            if 'Bakul' in df.columns:
                list_bakul = df['Bakul'].dropna().unique().tolist()
                selected_bakul = st.selectbox("Pilih Nama Bakul:", list_bakul)
                
                # Filter data berdasarkan bakul yang dipilih
                df_bakul = df[df['Bakul'] == selected_bakul]
                
                # Format list item
                items = []
                total_bayar = 0
                
                for _, row in df_bakul.iterrows():
                    nama_brg = row.get('Nama Barang', row.get('Barang', 'Ayam Segar'))
                    kg = row.get('KG', row.get('Qty', 0))
                    harga = row.get('Harga', 0)
                    jumlah = row.get('Jumlah', kg * harga)
                    
                    items.append({
                        'Nama Barang': str(nama_brg),
                        'KG': float(kg),
                        'Harga': float(harga),
                        'Jumlah': float(jumlah)
                    })
                    total_bayar += float(jumlah)
                
                st.success(f"Berhasil memuat {len(items)} barang untuk {selected_bakul}")
                
            else:
                st.error("Kolom 'Bakul' tidak ditemukan dalam sheet ini.")
                selected_bakul = None
                items = []
                total_bayar = 0

        except Exception as e:
            st.error(f"Gagal membaca file Excel: {e}")
            selected_bakul = None
            items = []
            total_bayar = 0

with col_preview:
    st.subheader("3. Preview & Download Nota")
    if uploaded_file and selected_bakul and items:
        # Generate gambar & docx
        img_bytes = generate_image_nota(
            tgl=selected_sheet, 
            bakul=selected_bakul, 
            group="Ayam Segar", 
            items=items, 
            total_bayar=total_bayar, 
            logo_path=logo_path
        )
        
        doc_bytes = generate_word_nota(
            tgl=selected_sheet, 
            bakul=selected_bakul, 
            group="Ayam Segar", 
            items=items, 
            total_bayar=total_bayar, 
            logo_path=logo_path
        )

        st.image(img_bytes, caption=f"Nota_{selected_bakul}_{selected_sheet}.png", use_column_width=True)
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📥 Download Gambar (.PNG)",
                data=img_bytes,
                file_name=f"Nota_{selected_bakul}_{selected_sheet}.png",
                mime="image/png"
            )
        with col_dl2:
            st.download_button(
                label="📄 Download Word (.DOCX)",
                data=doc_bytes,
                file_name=f"Nota_{selected_bakul}_{selected_sheet}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.info("Silakan upload file Excel dan pilih Bakul terlebih dahulu untuk melihat preview nota.")