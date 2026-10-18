import base64
from datetime import datetime
import io
import json
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Mm, Pt, RGBColor
import pandas as pd
import streamlit as st

# ==========================================
# 1. KONFIGURASI HALAMAN & UTILS
# ==========================================
st.set_page_config(
    page_title="Sistem Manajemen Ayam Segar",
    layout="wide",
    initial_sidebar_state="expanded",
)

logo_filename = None
for fname in ["ASTremove.PNG", "ASTremove.png", "AST.jpeg"]:
    if os.path.exists(fname):
        logo_filename = fname
        break


def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None


bg_base64 = get_base64_image("back.jpg")


# ==========================================
# 2. HELPER GENERATOR WORD & EXCEL NOTA
# ==========================================
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [
        ("top", top),
        ("bottom", bottom),
        ("left", left),
        ("right", right),
    ]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        '  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        '  <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        '  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        '  <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        '  <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        '  <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def append_word_nota_to_doc(doc, tgl, bakul, group, items, total_bayar, logo_path, no_bakul=""):
    """Menambahkan 1 Nota ke Word persis seperti contoh desain terlampir"""
    
    # 1. HEADER NOTA
    tbl_hdr = doc.add_table(rows=1, cols=2)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr.autofit = False

    cell_l, cell_r = tbl_hdr.rows[0].cells
    cell_l.width = Mm(75)
    cell_r.width = Mm(65)

    set_cell_margins(cell_l, top=0, bottom=0, left=0, right=0)
    set_cell_margins(cell_r, top=0, bottom=0, left=0, right=0)

    cell_l.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_r.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Brand / Logo Kiri
    tbl_brand = cell_l.add_table(rows=1, cols=2)
    tbl_brand.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_brand.autofit = False

    b_cell_img, b_cell_txt = tbl_brand.rows[0].cells
    b_cell_img.width = Mm(18)
    b_cell_txt.width = Mm(57)

    set_cell_margins(b_cell_img, top=0, bottom=0, left=0, right=0)
    set_cell_margins(b_cell_txt, top=0, bottom=0, left=0, right=0)

    b_cell_img.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    b_cell_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_img = b_cell_img.paragraphs[0]
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(0)
    if logo_path and os.path.exists(logo_path):
        run_img = p_img.add_run()
        run_img.add_picture(logo_path, width=Mm(15))

    p_txt = b_cell_txt.paragraphs[0]
    p_txt.paragraph_format.space_before = Pt(0)
    p_txt.paragraph_format.space_after = Pt(0)
    p_txt.paragraph_format.line_spacing = 1.15

    run_title = p_txt.add_run("AYAM SEGAR TUMPANG\n")
    run_title.bold = True
    run_title.font.name = "Courier New"
    run_title.font.size = Pt(9.5)
    run_title.font.color.rgb = RGBColor(180, 0, 0)

    run_sub = p_txt.add_run("Ds. Kambingan - Tumpang")
    run_sub.font.name = "Courier New"
    run_sub.font.size = Pt(8)
    run_sub.font.color.rgb = RGBColor(60, 60, 60)

    # Info Kanan
    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(0)
    p_r.paragraph_format.line_spacing = 1.15

    info_text = f"No: {no_bakul}\nTgl: {tgl}\nBakul: {bakul}\nGrp: {group}"
    r_info = p_r.add_run(info_text)
    r_info.font.name = "Courier New"
    r_info.font.size = Pt(8.5)

    # Spasi sebelum tabel
    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_before = Pt(12)
    p_sp1.paragraph_format.space_after = Pt(0)

    # 2. TABEL BARANG
    tbl_items = doc.add_table(rows=1, cols=4)
    tbl_items.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_items.autofit = False

    col_widths = [Mm(45), Mm(25), Mm(35), Mm(35)]
    headers = ["BARANG", "QTY", "HARGA", "JUMLAH"]
    alignments = [
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.CENTER,
    ]

    set_table_borders(tbl_items)

    hdr_cells = tbl_items.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        set_cell_margins(hdr_cells[i], top=80, bottom=80, left=60, right=60)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = alignments[i]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h_text)
        r.bold = True
        r.font.name = "Courier New"
        r.font.size = Pt(8.5)

    data_alignments = [
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
    ]

    for item in items:
        row_cells = tbl_items.add_row().cells
        kg_val = item["KG"]
        kg_str = (
            f"{int(kg_val)}"
            if isinstance(kg_val, float) and kg_val.is_integer()
            else (
                f"{kg_val:.2f}"
                if isinstance(kg_val, float)
                else str(kg_val)
            )
        )
        h_str = f"{item['Harga']:,.0f}".replace(",", ".")
        j_str = f"{item['Jumlah']:,.0f}".replace(",", ".")

        vals = [f" {item['Nama Barang']}", kg_str, h_str, j_str]
        for i, val in enumerate(vals):
            row_cells[i].width = col_widths[i]
            set_cell_margins(row_cells[i], top=100, bottom=100, left=60, right=60)
            p = row_cells[i].paragraphs[0]
            p.alignment = data_alignments[i]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Courier New"
            r.font.size = Pt(8.5)

    # 3. TOTAL (RAPAT KANAN DI LUAR TABEL)
    p_tot = doc.add_paragraph()
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_tot.paragraph_format.space_before = Pt(8)
    p_tot.paragraph_format.space_after = Pt(0)

    r_tlbl = p_tot.add_run("TOTAL: ")
    r_tlbl.bold = True
    r_tlbl.font.name = "Courier New"
    r_tlbl.font.size = Pt(9.5)

    tot_str = f"Rp {total_bayar:,.0f}".replace(",", ".")
    r_tval = p_tot.add_run(tot_str)
    r_tval.bold = True
    r_tval.font.name = "Courier New"
    r_tval.font.size = Pt(10)
    r_tval.font.color.rgb = RGBColor(180, 0, 0)

    # 4. FOOTER / TANDA TANGAN
    tbl_footer = doc.add_table(rows=2, cols=2)
    tbl_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_footer.autofit = False

    f_widths = [Mm(70), Mm(70)]
    for row in tbl_footer.rows:
        for i, cell in enumerate(row.cells):
            cell.width = f_widths[i]
            set_cell_margins(cell, top=0, bottom=0, left=0, right=0)

    cell_f0_l, cell_f0_r = tbl_footer.rows[0].cells

    p_f0_l = cell_f0_l.paragraphs[0]
    p_f0_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f0_l.paragraph_format.space_before = Pt(24)
    p_f0_l.paragraph_format.space_after = Pt(0)
    r_sl = p_f0_l.add_run("Penerima,")
    r_sl.font.name = "Courier New"
    r_sl.font.size = Pt(8.5)

    p_f0_r = cell_f0_r.paragraphs[0]
    p_f0_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f0_r.paragraph_format.space_before = Pt(24)
    p_f0_r.paragraph_format.space_after = Pt(0)
    r_sr = p_f0_r.add_run("Hormat Kami,")
    r_sr.font.name = "Courier New"
    r_sr.font.size = Pt(8.5)

    cell_f1_l, cell_f1_r = tbl_footer.rows[1].cells

    p_f1_l = cell_f1_l.paragraphs[0]
    p_f1_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f1_l.paragraph_format.space_before = Pt(24)
    p_f1_l.paragraph_format.space_after = Pt(0)
    r_tl = p_f1_l.add_run("( ....... )")
    r_tl.font.name = "Courier New"
    r_tl.font.size = Pt(8.5)

    p_f1_r = cell_f1_r.paragraphs[0]
    p_f1_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f1_r.paragraph_format.space_before = Pt(24)
    p_f1_r.paragraph_format.space_after = Pt(0)
    r_tr = p_f1_r.add_run("( ....... )")
    r_tr.font.name = "Courier New"
    r_tr.font.size = Pt(8.5)


def generate_word_nota(tgl, bakul, group, items, total_bayar, logo_path, no_bakul=""):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(5.5)
    section.top_margin = Mm(12)
    section.bottom_margin = Mm(10)
    section.left_margin = Mm(18)
    section.right_margin = Mm(18)

    append_word_nota_to_doc(doc, tgl, bakul, group, items, total_bayar, logo_path, no_bakul)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()


def write_single_excel_nota(ws, tgl, bakul, group, items, logo_path, no_bakul="", start_row=1):
    """Menyusun 1 Nota Data Asli di Excel dengan Format Presisi Sesuai Contoh"""
    font_family = "Courier New"
    f_title = Font(name=font_family, size=10, bold=True, color="B40000")
    f_sub = Font(name=font_family, size=8, color="3C3C3C")
    f_info = Font(name=font_family, size=8.5)
    f_hdr = Font(name=font_family, size=8.5, bold=True)
    f_data = Font(name=font_family, size=8.5)
    f_tot_lbl = Font(name=font_family, size=9.5, bold=True)
    f_tot_val = Font(name=font_family, size=10, bold=True, color="B40000")
    f_sig = Font(name=font_family, size=8.5)

    a_left = Alignment(horizontal="left", vertical="center")
    a_right = Alignment(horizontal="right", vertical="center")
    a_center = Alignment(horizontal="center", vertical="center")

    thin_side = Side(border_style="thin", color="000000")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16

    r = start_row

    ws.cell(row=r, column=1, value="AYAM SEGAR TUMPANG").font = f_title
    ws.cell(row=r, column=1).alignment = a_left
    ws.cell(row=r + 1, column=1, value="Ds. Kambingan - Tumpang").font = f_sub
    ws.cell(row=r + 1, column=1).alignment = a_left

    ws.cell(row=r, column=4, value=f"No: {no_bakul}").font = f_info
    ws.cell(row=r, column=4).alignment = a_right
    ws.cell(row=r + 1, column=4, value=f"Tgl: {tgl}").font = f_info
    ws.cell(row=r + 1, column=4).alignment = a_right
    ws.cell(row=r + 2, column=4, value=f"Bakul: {bakul}").font = f_info
    ws.cell(row=r + 2, column=4).alignment = a_right
    ws.cell(row=r + 3, column=4, value=f"Grp: {group}").font = f_info
    ws.cell(row=r + 3, column=4).alignment = a_right

    r += 5

    headers = ["BARANG", "QTY", "HARGA", "JUMLAH"]
    aligns = [a_center, a_center, a_center, a_center]
    for col_i, h_text in enumerate(headers, start=1):
        c = ws.cell(row=r, column=col_i, value=h_text)
        c.font = f_hdr
        c.alignment = aligns[col_i - 1]
        c.border = border_cell

    r += 1
    start_item_row = r

    for item in items:
        cA = ws.cell(row=r, column=1, value=item["Nama Barang"])
        cA.font = f_data
        cA.alignment = a_left
        cA.border = border_cell

        cB = ws.cell(row=r, column=2, value=float(item["KG"]))
        cB.font = f_data
        cB.alignment = a_center
        cB.border = border_cell
        cB.number_format = "#,##0.00" if isinstance(item["KG"], float) and not item["KG"].is_integer() else "#,##0"

        cC = ws.cell(row=r, column=3, value=float(item["Harga"]))
        cC.font = f_data
        cC.alignment = a_right
        cC.border = border_cell
        cC.number_format = '#,##0'

        cD = ws.cell(row=r, column=4, value=f"=B{r}*C{r}")
        cD.font = f_data
        cD.alignment = a_right
        cD.border = border_cell
        cD.number_format = '#,##0'

        r += 1

    end_item_row = r - 1

    c_tot_lbl = ws.cell(row=r, column=3, value="TOTAL:")
    c_tot_lbl.font = f_tot_lbl
    c_tot_lbl.alignment = a_right

    c_tot_val = ws.cell(row=r, column=4, value=f"=SUM(D{start_item_row}:D{end_item_row})")
    c_tot_val.font = f_tot_val
    c_tot_val.alignment = a_right
    c_tot_val.number_format = '"Rp "#,##0'

    r += 3

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    cL = ws.cell(row=r, column=1, value="Penerima,")
    cL.font = f_sig
    cL.alignment = a_center

    cR = ws.cell(row=r, column=3, value="Hormat Kami,")
    cR.font = f_sig
    cR.alignment = a_center

    r += 3

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    cTL = ws.cell(row=r, column=1, value="( ....... )")
    cTL.font = f_sig
    cTL.alignment = a_center

    cTR = ws.cell(row=r, column=3, value="( ....... )")
    cTR.font = f_sig
    cTR.alignment = a_center

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = False
    ws.sheet_view.showGridLines = True

    return r + 2


def generate_excel_nota_satuan(tgl, bakul, group, items, total_bayar, logo_path, no_bakul=""):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nota"

    write_single_excel_nota(
        ws=ws,
        tgl=tgl,
        bakul=bakul,
        group=group,
        items=items,
        logo_path=logo_path,
        no_bakul=no_bakul,
        start_row=1,
    )

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# 3. CSS STYLING UTAMA
# ==========================================
bg_css = ""
if bg_base64:
    bg_css = f"""
        .stApp {{
            background-image: url("data:image/jpeg;base64,{bg_base64}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }}
    """

st.markdown(
    f"""
    <style>
        footer, #MainMenu {{ visibility: hidden; }}
        [data-testid="stHeader"] {{ background-color: transparent !important; z-index: 100 !important; }}
        [data-testid="stSidebar"] {{ background-color: #FFEBEE !important; border-right: 3px solid #C62828 !important; }}
        [data-testid="stSidebar"] .stRadio label {{ font-size: 15px !important; font-weight: bold !important; color: #262626 !important; }}
        
        {bg_css}

        .block-container {{
            padding-top: 1.5rem !important;
        }}

        div[data-testid="stColumn"] > div:has(input[type="password"]) {{
            background-color: rgba(255, 255, 255, 0.98);
            padding: 30px;
            border-radius: 16px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.25);
            border: 1px solid #FFCDD2;
        }}
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 4. LOGIN SYSTEM
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = ""


def login():
    st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)
    _, col2, _ = st.columns([1, 1.8, 1])

    with col2:
        if logo_filename:
            _, c_img, _ = st.columns([1, 2, 1])
            with c_img:
                st.image(logo_filename, use_container_width=True)

        st.markdown(
            "<h3 style='text-align: center; color: #C62828; margin-top:"
            " 10px;'>Ayam Segar Tumpang</h3>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<p style='text-align: center; color: #666; font-size: 13px;'>Silakan"
            " login untuk mengakses sistem</p>",
            unsafe_allow_html=True,
        )

        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        if st.button("Masuk / Login", use_container_width=True, type="primary"):
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

# ==========================================
# 5. MASTER HARGA JSON
# ==========================================
FILE_HARGA = "master_harga.json"
default_harga = {
    "glondong": 28500,
    "jeroan": 10000,
    "usus": 16500,
    "telur_a": 269000,
    "telur_b": 150000,
    "telur_bt": 150000,
    "peti": 6000,
    "box": 160000,
}

if os.path.exists(FILE_HARGA):
    try:
        with open(FILE_HARGA, "r") as f:
            saved_harga = json.load(f)
    except Exception:
        saved_harga = default_harga
else:
    saved_harga = default_harga

# ==========================================
# 6. SIDEBAR MENU
# ==========================================
with st.sidebar:
    if logo_filename:
        st.image(logo_filename, width=90)
    st.markdown(
        "<h2 style='color: #C62828; margin: 0; font-size: 26px; font-weight:"
        " bold;'>AST SYSTEM</h2>",
        unsafe_allow_html=True,
    )
    st.write(f"Logged in as: **{st.session_state.role}**")
    st.markdown("---")

    st.markdown("### 📌 MENU UTAMA")
    menu_options = [
        "📊 Dashboard",
        "📄 Nota",
        "🛍️ Penjualan",
        "🚚 Pengiriman",
        "📦 Stock",
        "💵 Finance",
        "⏱️ Absensi & Jadwal",
    ]
    selected_menu = st.pills("Pilih Halaman:", menu_options, default="📊 Dashboard")
    if not selected_menu:
        selected_menu = "📊 Dashboard"

    sub_menu = None
    if selected_menu in ["📄 Nota", "🧾 Nota"]:
        with st.expander("📂 Sub-Menu Nota", expanded=True):
            sub_menu = st.pills("Tipe Nota:", ["📑 Bakul", "🏬 Bedak", "🤝 Mitra"])

    if selected_menu in ["📄 Nota", "🧾 Nota"]:
        st.markdown("---")
        st.markdown(
            "<h3 style='color: #C62828;'>⚙️ Master Harga</h3>",
            unsafe_allow_html=True,
        )

        def input_harga(label, key_name, default_val, step_val):
            val = st.number_input(
                f"{label}",
                value=int(saved_harga.get(key_name, default_val)),
                step=step_val,
                format="%d",
            )
            st.caption(f"➔ **Rp {val:,.0f}**".replace(",", "."))
            return val

        h_glondong = input_harga("Harga Glondong", "glondong", 28500, 500)
        h_jeroan = input_harga("Harga Jeroan", "jeroan", 12000, 500)
        h_usus = input_harga("Harga Usus", "usus", 16500, 500)
        h_telur_a = input_harga("Harga Telur A", "telur_a", 269000, 1000)
        h_telur_b = input_harga("Harga Telur B", "telur_b", 250000, 1000)
        h_telur_bt = input_harga("Harga Telur BT", "telur_bt", 150000, 1000)
        h_peti = input_harga("Harga Peti", "peti", 2000, 100)
        h_box = input_harga("Harga Box", "box", 28500, 500)

        current_harga = {
            "glondong": h_glondong,
            "jeroan": h_jeroan,
            "usus": h_usus,
            "telur_a": h_telur_a,
            "telur_b": h_telur_b,
            "telur_bt": h_telur_bt,
            "peti": h_peti,
            "box": h_box,
        }
        if current_harga != saved_harga:
            with open(FILE_HARGA, "w") as f:
                json.dump(current_harga, f)

    st.markdown("---")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.role = ""
        st.rerun()

# ==========================================
# 7. AREA KONTEN UTAMA
# ==========================================
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(
        f"<h2 style='color: #C62828; margin:0;'>Ayam Segar Tumpang -"
        f" {selected_menu}</h2>",
        unsafe_allow_html=True,
    )
with col_h2:
    st.markdown(
        "<div style='text-align: right; font-weight: bold; padding-top:"
        f" 10px;'>👤 {st.session_state.role}</div>",
        unsafe_allow_html=True,
    )

st.markdown(
    "<hr style='border: 1px solid #C62828; margin-top: 5px; margin-bottom:"
    " 20px;'>",
    unsafe_allow_html=True,
)

if selected_menu in ["📄 Nota", "🧾 Nota"]:
    if sub_menu == "📑 Bakul" or sub_menu is None:
        col_up1, col_up2 = st.columns([2, 1])
        with col_up1:
            uploaded_file = st.file_uploader(
                "Upload File Rekap Excel (.xlsx)", type=["xlsx"]
            )
        with col_up2:
            tanggal_transaksi = st.date_input(
                "Tanggal Transaksi", value=datetime.today()
            )

        if uploaded_file is not None:
            all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
            valid_sheets = [
                s
                for s in all_sheets.keys()
                if s not in ["TOTAL TONASE", "NOTA FR", "Sheet1", "ploting"]
            ]

            tab_satuan, tab_bulk = st.tabs(
                ["📄 Nota Satuan (Word/Excel)", "📦 Export All Nota"]
            )

            # ==========================================
            # TAB 1: NOTA SATUAN (PILIH 1 BAKUL)
            # ==========================================
            with tab_satuan:
                selected_sheet = st.selectbox("Pilih Group / Sheet", valid_sheets)
                df_raw = all_sheets[selected_sheet]

                header_idx = 0
                for idx, row in df_raw.iterrows():
                    if row.astype(str).str.upper().str.contains("NAMA").any():
                        header_idx = idx
                        break

                df = df_raw.iloc[header_idx + 1 :].copy()
                df.columns = [
                    str(c).strip().upper() for c in df_raw.iloc[header_idx].values
                ]

                name_col = next((c for c in df.columns if "NAMA" in c), df.columns[1])
                no_col = next(
                    (c for c in df.columns if c in ["NO", "NO.", "NOMOR"]), None
                )

                df = df[df[name_col].notna()].copy()

                bakul_options = []
                bakul_map = {}

                for idx, (real_idx, row) in enumerate(df.iterrows(), start=1):
                    nama_bakul = (
                        str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                    )
                    if not nama_bakul:
                        continue

                    if (
                        no_col
                        and pd.notna(row[no_col])
                        and str(row[no_col]).strip() != ""
                    ):
                        no_val = str(row[no_col]).strip()
                        if no_val.endswith(".0"):
                            no_val = no_val[:-2]
                        label = f"{no_val}. {nama_bakul}"
                    else:
                        no_val = str(idx)
                        label = f"{idx}. {nama_bakul}"

                    bakul_options.append(label)
                    bakul_map[label] = (nama_bakul, real_idx, no_val)

                selected_bakul_label = st.selectbox("Pilih Nama Bakul", bakul_options)
                selected_bakul_info = bakul_map.get(selected_bakul_label, None)

                if selected_bakul_info:
                    selected_bakul, real_idx, no_bakul = selected_bakul_info

                    if (
                        "last_bakul" not in st.session_state
                        or st.session_state["last_bakul"] != selected_bakul_label
                    ):
                        st.session_state["last_bakul"] = selected_bakul_label
                        st.session_state["qty_peti_val"] = 0.0
                        st.session_state["qty_box_val"] = 0.0
                        st.session_state["qty_telur_b_val"] = 0.0
                        st.session_state["qty_telur_bt_val"] = 0.0

                    row_bakul = df.loc[real_idx]

                    def get_valid_float(col_name):
                        try:
                            if not col_name or col_name not in df.columns:
                                return 0.0
                            val = row_bakul[col_name]
                            num = pd.to_numeric(val, errors="coerce")
                            return 0.0 if pd.isna(num) else float(num)
                        except Exception:
                            return 0.0

                    col_in1, col_in2, col_in3, col_in4 = st.columns(4)
                    with col_in1:
                        qty_peti = st.number_input(
                            "Jumlah Peti", key="qty_peti_val", step=1
                        )
                    with col_in2:
                        qty_box = st.number_input(
                            "Jumlah Box", key="qty_box_val", step=1
                        )
                    with col_in3:
                        qty_telur_b_manual = st.number_input(
                            "Jumlah Telur B", key="qty_telur_b_val", step=1
                        )
                    with col_in4:
                        qty_telur_bt_manual = st.number_input(
                            "Jumlah Telur BT", key="qty_telur_bt_val", step=1
                        )

                    qty_tonase = get_valid_float(
                        next((c for c in df.columns if "TONASE" in c), "")
                    )
                    qty_jeroan = get_valid_float(
                        next((c for c in df.columns if "JEROAN" in c), "")
                    )
                    qty_usus = get_valid_float(
                        next((c for c in df.columns if "USUS" in c), "")
                    )
                    qty_telur_a = get_valid_float(
                        next(
                            (c for c in df.columns if "TELUR A" in c or "TELUR" in c), ""
                        )
                    )

                    qty_telur_b_excel = get_valid_float(
                        next((c for c in df.columns if "TELUR B" in c), "")
                    )
                    qty_telur_b = (
                        qty_telur_b_excel
                        if qty_telur_b_excel > 0
                        else qty_telur_b_manual
                    )

                    qty_telur_bt_excel = get_valid_float(
                        next((c for c in df.columns if "TELUR BT" in c), "")
                    )
                    qty_telur_bt = (
                        qty_telur_bt_excel
                        if qty_telur_bt_excel > 0
                        else qty_telur_bt_manual
                    )

                    tot_glondong = qty_tonase * h_glondong
                    tot_jeroan = qty_jeroan * h_jeroan
                    tot_usus = qty_usus * h_usus
                    tot_telur_a = qty_telur_a * h_telur_a
                    tot_telur_b = qty_telur_b * h_telur_b
                    tot_telur_bt = qty_telur_bt * h_telur_bt
                    tot_peti = qty_peti * h_peti
                    tot_box = qty_box * h_box

                    total_bayar = (
                        tot_glondong
                        + tot_jeroan
                        + tot_usus
                        + tot_telur_a
                        + tot_telur_b
                        + tot_telur_bt
                        + tot_peti
                        + tot_box
                    )

                    items = [
                        {
                            "Nama Barang": "GLONDONG",
                            "KG": qty_tonase,
                            "Harga": h_glondong,
                            "Jumlah": tot_glondong,
                        },
                        {
                            "Nama Barang": "JEROAN",
                            "KG": qty_jeroan,
                            "Harga": h_jeroan,
                            "Jumlah": tot_jeroan,
                        },
                        {
                            "Nama Barang": "USUS B",
                            "KG": qty_usus,
                            "Harga": h_usus,
                            "Jumlah": tot_usus,
                        },
                        {
                            "Nama Barang": "TELUR A",
                            "KG": qty_telur_a,
                            "Harga": h_telur_a,
                            "Jumlah": tot_telur_a,
                        },
                        {
                            "Nama Barang": "TELUR B",
                            "KG": qty_telur_b,
                            "Harga": h_telur_b,
                            "Jumlah": tot_telur_b,
                        },
                        {
                            "Nama Barang": "TELUR BT",
                            "KG": qty_telur_bt,
                            "Harga": h_telur_bt,
                            "Jumlah": tot_telur_bt,
                        },
                        {
                            "Nama Barang": "PETI",
                            "KG": qty_peti,
                            "Harga": h_peti,
                            "Jumlah": tot_peti,
                        },
                        {
                            "Nama Barang": "BOX",
                            "KG": qty_box,
                            "Harga": h_box,
                            "Jumlah": tot_box,
                        },
                    ]

                    filtered_items = [i for i in items if i["KG"] > 0]

                    if filtered_items:
                        with st.container(border=True):
                            col_l, col_m, col_r = st.columns([1.5, 4, 3])
                            with col_l:
                                if logo_filename:
                                    st.image(logo_filename, width=110)
                            with col_m:
                                st.markdown(
                                    "<h3 style='margin:0; color:#B40000; font-weight: bold;"
                                    " font-size:24px;'>AYAM SEGAR TUMPANG</h3><small>Ds. Kambingan"
                                    " - Tumpang - Kab. Malang</small>",
                                    unsafe_allow_html=True,
                                )
                            with col_r:
                                st.markdown(
                                    f"<div style='text-align:right;'><small><b>No:</b> {no_bakul}<br>"
                                    f"<b>Tgl:</b> {tanggal_transaksi.strftime('%d-%m-%Y')}<br>"
                                    f"<b>Bakul:</b> {selected_bakul}<br>"
                                    f"<b>Grp:</b> {selected_sheet}</small></div>",
                                    unsafe_allow_html=True,
                                )

                            rows_html = ""
                            for item in filtered_items:
                                kg_str = (
                                    f"{int(item['KG'])}"
                                    if isinstance(item["KG"], float) and item["KG"].is_integer()
                                    else (
                                        f"{item['KG']:.2f}"
                                        if isinstance(item["KG"], float)
                                        else str(item["KG"])
                                    )
                                )
                                rows_html += f"<tr><td>{item['Nama Barang']}</td><td style='text-align: center;'>{kg_str}</td><td style='text-align: right;'>{item['Harga']:,.0f}</td><td style='text-align: right;'>{item['Jumlah']:,.0f}</td></tr>".replace(
                                    ",", "."
                                )

                            st.markdown(
                                f"""
                                <table style="width:100%; border-collapse:collapse; margin-top:10px;">
                                    <thead><tr>
                                        <th style="border:1px solid #000; padding:6px; text-align:center;">BARANG</th>
                                        <th style="border:1px solid #000; padding:6px; text-align:center;">QTY</th>
                                        <th style="border:1px solid #000; padding:6px; text-align:center;">HARGA</th>
                                        <th style="border:1px solid #000; padding:6px; text-align:center;">JUMLAH</th>
                                    </tr></thead>
                                    <tbody>{rows_html}</tbody>
                                </table>
                                <div style="text-align:right; margin-top:10px; font-weight:bold;">
                                    TOTAL: <span style="color:#B40000; font-size:20px; font-weight:bold;">Rp {total_bayar:,.0f}</span>
                                </div>
                            """.replace(",", "."),
                                unsafe_allow_html=True,
                            )

                        col_btn1, col_btn2 = st.columns(2)

                        with col_btn1:
                            word_data = generate_word_nota(
                                tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                                bakul=selected_bakul,
                                group=selected_sheet,
                                items=filtered_items,
                                total_bayar=total_bayar,
                                logo_path=logo_filename,
                                no_bakul=no_bakul,
                            )
                            st.download_button(
                                label="📝 Download Nota Word (.docx)",
                                data=word_data,
                                file_name=(
                                    f"Nota_{selected_bakul}_{tanggal_transaksi.strftime('%Y%m%d')}.docx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                ),
                                use_container_width=True,
                            )

                        with col_btn2:
                            excel_satuan_data = generate_excel_nota_satuan(
                                tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                                bakul=selected_bakul,
                                group=selected_sheet,
                                items=filtered_items,
                                total_bayar=total_bayar,
                                logo_path=logo_filename,
                                no_bakul=no_bakul,
                            )
                            st.download_button(
                                label="📊 Download Nota Excel (.xlsx)",
                                data=excel_satuan_data,
                                file_name=(
                                    f"Nota_{selected_bakul}_{tanggal_transaksi.strftime('%Y%m%d')}.xlsx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                ),
                                use_container_width=True,
                            )

            # ==========================================
            # TAB 2: EXPORT ALL NOTA (KE 1 FILE EXCEL ATAU 1 FILE WORD)
            # ==========================================
            with tab_bulk:
                st.markdown("### 📄 Export All Nota (Excel / Word)")
                st.caption("Centang nama bakul yang ingin diproses ke dalam 1 file utuh.")

                selected_bakul_by_sheet = {}

                for sheet_name in valid_sheets:
                    df_raw_sheet = all_sheets[sheet_name]

                    header_idx_sheet = 0
                    for idx, row in df_raw_sheet.iterrows():
                        if row.astype(str).str.upper().str.contains("NAMA").any():
                            header_idx_sheet = idx
                            break

                    df_s = df_raw_sheet.iloc[header_idx_sheet + 1 :].copy()
                    df_s.columns = [
                        str(c).strip().upper()
                        for c in df_raw_sheet.iloc[header_idx_sheet].values
                    ]

                    name_c = next(
                        (c for c in df_s.columns if "NAMA" in c), df_s.columns[1]
                    )
                    df_s = df_s[df_s[name_c].notna()].copy()

                    list_bakul = [
                        str(b).strip() for b in df_s[name_c].tolist() if str(b).strip()
                    ]

                    if list_bakul:
                        with st.expander(
                            f"📋 Sheet: {sheet_name} ({len(list_bakul)} Bakul)",
                            expanded=True,
                        ):
                            col_a, col_b = st.columns([1, 4])
                            with col_a:
                                select_all = st.checkbox(
                                    "Pilih Semua",
                                    value=True,
                                    key=f"all_xl_img_{sheet_name}",
                                )

                            default_selected = list_bakul if select_all else []
                            chosen = st.multiselect(
                                "Daftar Bakul Terpilih:",
                                options=list_bakul,
                                default=default_selected,
                                key=f"ms_xl_img_{sheet_name}",
                            )
                            selected_bakul_by_sheet[sheet_name] = (df_s, name_c, chosen)

                st.markdown("---")

                col_dl1, col_dl2 = st.columns(2)

                with col_dl1:
                    if st.button(
                        "🚀 Generate Excel All Nota (.xlsx)",
                        type="primary",
                        use_container_width=True,
                    ):
                        wb_bulk = openpyxl.Workbook()
                        wb_bulk.remove(wb_bulk.active)

                        total_processed = 0

                        for sheet_name, (
                            df_s,
                            name_c,
                            chosen_bakul,
                        ) in selected_bakul_by_sheet.items():
                            if not chosen_bakul:
                                continue

                            ws_bulk = wb_bulk.create_sheet(title=sheet_name[:31])

                            df_filtered = df_s[
                                df_s[name_c].astype(str).str.strip().isin(chosen_bakul)
                            ].copy()

                            no_col_sheet = next(
                                (c for c in df_s.columns if c in ["NO", "NO.", "NOMOR"]), None
                            )

                            current_start_row = 1

                            for idx, row in df_filtered.iterrows():
                                nama_bakul = str(row[name_c]).strip()

                                if (
                                    no_col_sheet
                                    and pd.notna(row[no_col_sheet])
                                    and str(row[no_col_sheet]).strip() != ""
                                ):
                                    no_val = str(row[no_col_sheet]).strip()
                                    if no_val.endswith(".0"):
                                        no_val = no_val[:-2]
                                else:
                                    no_val = ""

                                def get_val_from_row(c_keyword):
                                    try:
                                        col_found = next(
                                            (
                                                c
                                                for c in df_s.columns
                                                if c_keyword in str(c).upper()
                                            ),
                                            None,
                                        )
                                        if col_found is not None:
                                            val = pd.to_numeric(
                                                row[col_found], errors="coerce"
                                            )
                                            return 0.0 if pd.isna(val) else float(val)
                                    except Exception:
                                        pass
                                    return 0.0

                                kg_tonase = get_val_from_row("TONASE")
                                kg_jeroan = get_val_from_row("JEROAN")
                                kg_usus = get_val_from_row("USUS")

                                kg_telur_a = get_val_from_row("TELUR A")
                                if kg_telur_a == 0.0:
                                    kg_telur_a = get_val_from_row("TELUR")

                                kg_telur_b = get_val_from_row("TELUR B")
                                kg_telur_bt = get_val_from_row("TELUR BT")

                                tot_glondong = kg_tonase * h_glondong
                                tot_jeroan = kg_jeroan * h_jeroan
                                tot_usus = kg_usus * h_usus
                                tot_telur_a = kg_telur_a * h_telur_a
                                tot_telur_b = kg_telur_b * h_telur_b
                                tot_telur_bt = kg_telur_bt * h_telur_bt

                                items = [
                                    {
                                        "Nama Barang": "GLONDONG",
                                        "KG": kg_tonase,
                                        "Harga": h_glondong,
                                        "Jumlah": tot_glondong,
                                    },
                                    {
                                        "Nama Barang": "JEROAN",
                                        "KG": kg_jeroan,
                                        "Harga": h_jeroan,
                                        "Jumlah": tot_jeroan,
                                    },
                                    {
                                        "Nama Barang": "USUS B",
                                        "KG": kg_usus,
                                        "Harga": h_usus,
                                        "Jumlah": tot_usus,
                                    },
                                    {
                                        "Nama Barang": "TELUR A",
                                        "KG": kg_telur_a,
                                        "Harga": h_telur_a,
                                        "Jumlah": tot_telur_a,
                                    },
                                    {
                                        "Nama Barang": "TELUR B",
                                        "KG": kg_telur_b,
                                        "Harga": h_telur_b,
                                        "Jumlah": tot_telur_b,
                                    },
                                    {
                                        "Nama Barang": "TELUR BT",
                                        "KG": kg_telur_bt,
                                        "Harga": h_telur_bt,
                                        "Jumlah": tot_telur_bt,
                                    },
                                ]

                                filtered_items = [i for i in items if i["KG"] > 0]

                                if filtered_items:
                                    write_single_excel_nota(
                                        ws=ws_bulk,
                                        tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                                        bakul=nama_bakul,
                                        group=sheet_name,
                                        items=filtered_items,
                                        logo_path=logo_filename,
                                        no_bakul=no_val,
                                        start_row=current_start_row,
                                    )
                                    
                                    current_start_row += 16
                                    total_processed += 1

                        if total_processed > 0:
                            output_excel_bulk = io.BytesIO()
                            wb_bulk.save(output_excel_bulk)

                            st.success(f"✅ Berhasil mengeksport {total_processed} nota ke Excel!")
                            st.download_button(
                                label="📥 Download File Excel All Nota",
                                data=output_excel_bulk.getvalue(),
                                file_name=(
                                    "EXCEL_ALL_NOTA_"
                                    f"{tanggal_transaksi.strftime('%Y%m%d')}.xlsx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                ),
                                use_container_width=True,
                                key="dl_btn_excel_bulk"
                            )
                        else:
                            st.warning("⚠️ Tidak ada nota yang diproses.")

                with col_dl2:
                    if st.button(
                        "🚀 Generate Word All Nota (.docx)",
                        type="primary",
                        use_container_width=True,
                    ):
                        doc_bulk = Document()
                        section = doc_bulk.sections[0]
                        section.page_width = Inches(8.5)
                        section.page_height = Inches(5.5)
                        section.top_margin = Mm(12)
                        section.bottom_margin = Mm(10)
                        section.left_margin = Mm(18)
                        section.right_margin = Mm(18)

                        total_processed_word = 0

                        for sheet_name, (
                            df_s,
                            name_c,
                            chosen_bakul,
                        ) in selected_bakul_by_sheet.items():
                            if not chosen_bakul:
                                continue

                            df_filtered = df_s[
                                df_s[name_c].astype(str).str.strip().isin(chosen_bakul)
                            ].copy()

                            no_col_sheet = next(
                                (c for c in df_s.columns if c in ["NO", "NO.", "NOMOR"]), None
                            )

                            for idx, row in df_filtered.iterrows():
                                nama_bakul = str(row[name_c]).strip()

                                if (
                                    no_col_sheet
                                    and pd.notna(row[no_col_sheet])
                                    and str(row[no_col_sheet]).strip() != ""
                                ):
                                    no_val = str(row[no_col_sheet]).strip()
                                    if no_val.endswith(".0"):
                                        no_val = no_val[:-2]
                                else:
                                    no_val = ""

                                def get_val_from_row(c_keyword):
                                    try:
                                        col_found = next(
                                            (
                                                c
                                                for c in df_s.columns
                                                if c_keyword in str(c).upper()
                                            ),
                                            None,
                                        )
                                        if col_found is not None:
                                            val = pd.to_numeric(
                                                row[col_found], errors="coerce"
                                            )
                                            return 0.0 if pd.isna(val) else float(val)
                                    except Exception:
                                        pass
                                    return 0.0

                                kg_tonase = get_val_from_row("TONASE")
                                kg_jeroan = get_val_from_row("JEROAN")
                                kg_usus = get_val_from_row("USUS")

                                kg_telur_a = get_val_from_row("TELUR A")
                                if kg_telur_a == 0.0:
                                    kg_telur_a = get_val_from_row("TELUR")

                                kg_telur_b = get_val_from_row("TELUR B")
                                kg_telur_bt = get_val_from_row("TELUR BT")

                                tot_glondong = kg_tonase * h_glondong
                                tot_jeroan = kg_jeroan * h_jeroan
                                tot_usus = kg_usus * h_usus
                                tot_telur_a = kg_telur_a * h_telur_a
                                tot_telur_b = kg_telur_b * h_telur_b
                                tot_telur_bt = kg_telur_bt * h_telur_bt
                                tot_bayar = (
                                    tot_glondong
                                    + tot_jeroan
                                    + tot_usus
                                    + tot_telur_a
                                    + tot_telur_b
                                    + tot_telur_bt
                                )

                                items = [
                                    {
                                        "Nama Barang": "GLONDONG",
                                        "KG": kg_tonase,
                                        "Harga": h_glondong,
                                        "Jumlah": tot_glondong,
                                    },
                                    {
                                        "Nama Barang": "JEROAN",
                                        "KG": kg_jeroan,
                                        "Harga": h_jeroan,
                                        "Jumlah": tot_jeroan,
                                    },
                                    {
                                        "Nama Barang": "USUS B",
                                        "KG": kg_usus,
                                        "Harga": h_usus,
                                        "Jumlah": tot_usus,
                                    },
                                    {
                                        "Nama Barang": "TELUR A",
                                        "KG": kg_telur_a,
                                        "Harga": h_telur_a,
                                        "Jumlah": tot_telur_a,
                                    },
                                    {
                                        "Nama Barang": "TELUR B",
                                        "KG": kg_telur_b,
                                        "Harga": h_telur_b,
                                        "Jumlah": tot_telur_b,
                                    },
                                    {
                                        "Nama Barang": "TELUR BT",
                                        "KG": kg_telur_bt,
                                        "Harga": h_telur_bt,
                                        "Jumlah": tot_telur_bt,
                                    },
                                ]

                                filtered_items = [i for i in items if i["KG"] > 0]

                                if filtered_items:
                                    append_word_nota_to_doc(
                                        doc=doc_bulk,
                                        tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                                        bakul=nama_bakul,
                                        group=sheet_name,
                                        items=filtered_items,
                                        total_bayar=tot_bayar,
                                        logo_path=logo_filename,
                                        no_bakul=no_val,
                                    )
                                    p_space = doc_bulk.add_paragraph()
                                    p_space.paragraph_format.space_before = Pt(18)
                                    p_space.paragraph_format.space_after = Pt(0)
                                    total_processed_word += 1

                        if total_processed_word > 0:
                            output_word_bulk = io.BytesIO()
                            doc_bulk.save(output_word_bulk)

                            st.success(f"✅ Berhasil mengeksport {total_processed_word} nota ke Word!")
                            st.download_button(
                                label="📥 Download File Word All Nota",
                                data=output_word_bulk.getvalue(),
                                file_name=(
                                    "WORD_ALL_NOTA_"
                                    f"{tanggal_transaksi.strftime('%Y%m%d')}.docx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                ),
                                use_container_width=True,
                                key="dl_btn_word_bulk"
                            )
                        else:
                            st.warning("⚠️ Tidak ada nota yang diproses.")

        else:
            st.info("💡 Silakan upload file Rekap Excel untuk memproses nota.")
    else:
        st.info(f"Halaman {selected_menu} sedang dalam pengembangan.")
else:
    st.info(f"Halaman {selected_menu} sedang dalam pengembangan.")
