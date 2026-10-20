import base64
from datetime import datetime
import io
import json
import os
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Mm, Pt, RGBColor
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
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


def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
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


def generate_word_nota(tgl, bakul, group, items, total_bayar, logo_path, no_bakul=""):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(130)
    section.page_height = Mm(105)
    section.top_margin = Mm(8)      # Diperkecil agar lebih atas
    section.bottom_margin = Mm(3)   # Diperkecil agar lebih bawah
    section.left_margin = Mm(1)     # Mepet ke kiri (ubah dari Mm(2) ke Mm(1))
    section.right_margin = Mm(8)    # Sisi kanan dibuat lebih lega

    # 1. HEADER NOTA
    tbl_hdr = doc.add_table(rows=1, cols=2)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_hdr.autofit = False

    cell_l, cell_r = tbl_hdr.rows[0].cells
    cell_l.width = Mm(65)
    cell_r.width = Mm(55)

    set_cell_margins(cell_l, top=0, bottom=0, left=0, right=0)
    set_cell_margins(cell_r, top=0, bottom=0, left=0, right=0)

    cell_l.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    cell_r.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    tbl_brand = cell_l.add_table(rows=1, cols=2)
    tbl_brand.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_brand.autofit = False

    b_cell_img, b_cell_txt = tbl_brand.rows[0].cells
    b_cell_img.width = Mm(20)
    b_cell_txt.width = Mm(45)

    set_cell_margins(b_cell_img, top=0, bottom=0, left=0, right=0)
    set_cell_margins(b_cell_txt, top=0, bottom=0, left=0, right=0)

    b_cell_img.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    b_cell_txt.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    p_img = b_cell_img.paragraphs[0]
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(0)
    if logo_path and os.path.exists(logo_path):
        run_img = p_img.add_run()
        run_img.add_picture(logo_path, width=Mm(18))

    p_txt = b_cell_txt.paragraphs[0]
    p_txt.paragraph_format.space_before = Pt(0)
    p_txt.paragraph_format.space_after = Pt(0)
    p_txt.paragraph_format.line_spacing = 1.05

    run_title = p_txt.add_run("AYAM SEGAR TUMPANG\n")
    run_title.bold = True
    run_title.font.name = "Courier New"
    run_title.font.size = Pt(11)
    run_title.font.color.rgb = RGBColor(198, 40, 40)

    run_sub = p_txt.add_run("Ds. Kambingan - Tumpang - Malang")
    run_sub.font.name = "Courier New"
    run_sub.font.size = Pt(7.5)
    run_sub.font.color.rgb = RGBColor(90, 90, 90)

    # Info No, Tanggal, Bakul, Group di Kanan
    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(18)
    p_r.paragraph_format.space_after = Pt(0)
    p_r.paragraph_format.line_spacing = 1.05

    info_text = f"No: {no_bakul}\nTanggal: {tgl}\nPembeli / Bakul: {bakul}\nGroup: {group}"
    r_info = p_r.add_run(info_text)
    r_info.font.name = "Courier New"
    r_info.font.size = Pt(9)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(1)
    p_spacer.paragraph_format.space_after = Pt(1)

    # 2. TABEL BARANG (BARANG di kolom 1, QTY di kolom 2)
    tbl_items = doc.add_table(rows=1, cols=4)
    tbl_items.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_items.autofit = False

    col_widths = [Mm(46), Mm(20), Mm(27), Mm(27)]
    headers = ["BARANG", "QTY / KG", "HARGA", "JUMLAH"]
    alignments = [
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
    ]

    set_table_borders(tbl_items)

    hdr_cells = tbl_items.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "EAEAEA")
        set_cell_margins(hdr_cells[i], top=12, bottom=12, left=15, right=15)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = alignments[i]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h_text)
        r.bold = True
        r.font.name = "Courier New"
        r.font.size = Pt(8)

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
        h_str = f"Rp {item['Harga']:,.0f}".replace(",", ".")
        j_str = f"Rp {item['Jumlah']:,.0f}".replace(",", ".")

        vals = [f" {item['Nama Barang']}", kg_str, h_str, j_str]
        for i, val in enumerate(vals):
            row_cells[i].width = col_widths[i]
            set_cell_margins(row_cells[i], top=12, bottom=12, left=15, right=15)
            p = row_cells[i].paragraphs[0]
            p.alignment = alignments[i]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Courier New"
            r.font.size = Pt(8)

    # 3. TOTAL & TANDA TANGAN
    tbl_footer = doc.add_table(rows=3, cols=2)
    tbl_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_footer.autofit = False

    f_widths = [Mm(65), Mm(55)]
    for row in tbl_footer.rows:
        for i, cell in enumerate(row.cells):
            cell.width = f_widths[i]
            set_cell_margins(cell, top=1, bottom=1, left=2, right=2)

    # BARIS 0: TOTAL
    cell_f0_l, cell_f0_r = tbl_footer.rows[0].cells
    p_f0_r = cell_f0_r.paragraphs[0]
    p_f0_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_f0_r.paragraph_format.space_before = Pt(2)
    p_f0_r.paragraph_format.space_after = Pt(2)

    r_lbl = p_f0_r.add_run("TOTAL : ")
    r_lbl.bold = True
    r_lbl.font.name = "Courier New"
    r_lbl.font.size = Pt(9)

    tot_str = f"Rp {total_bayar:,.0f}".replace(",", ".")
    r_tot = p_f0_r.add_run(tot_str)
    r_tot.bold = True
    r_tot.font.name = "Courier New"
    r_tot.font.size = Pt(12)
    r_tot.font.color.rgb = RGBColor(198, 40, 40)

    # BARIS 1: PENERIMA & HORMAT KAMI
    cell_f1_l, cell_f1_r = tbl_footer.rows[1].cells

    p_f1_l = cell_f1_l.paragraphs[0]
    p_f1_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f1_l.paragraph_format.space_before = Pt(4)
    p_f1_l.paragraph_format.space_after = Pt(0)
    r_sl = p_f1_l.add_run("Penerima,")
    r_sl.font.name = "Courier New"
    r_sl.font.size = Pt(7.5)

    p_f1_r = cell_f1_r.paragraphs[0]
    p_f1_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f1_r.paragraph_format.space_before = Pt(4)
    p_f1_r.paragraph_format.space_after = Pt(0)
    r_sr = p_f1_r.add_run("Hormat Kami,")
    r_sr.font.name = "Courier New"
    r_sr.font.size = Pt(7.5)

    # BARIS 2: TANDA TANGAN
    cell_f2_l, cell_f2_r = tbl_footer.rows[2].cells

    p_f2_l = cell_f2_l.paragraphs[0]
    p_f2_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f2_l.paragraph_format.space_before = Pt(26)
    p_f2_l.paragraph_format.space_after = Pt(0)
    r_tl = p_f2_l.add_run("( ........... )")
    r_tl.font.name = "Courier New"
    r_tl.font.size = Pt(7.5)

    p_f2_r = cell_f2_r.paragraphs[0]
    p_f2_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f2_r.paragraph_format.space_before = Pt(26)
    p_f2_r.paragraph_format.space_after = Pt(0)
    r_tr = p_f2_r.add_run("( ........... )")
    r_tr.font.name = "Courier New"
    r_tr.font.size = Pt(7.5)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()
    
# ---------------------------------------------------------------
# GENERATOR EXCEL NOTA (pengganti JPG/PNG)
# Layout disamakan dengan generate_word_nota:
#   - Kertas landscape (ukuran ikut custom form printer), margin atas 8 / bawah 3 / kiri 1 / kanan 8 mm
#   - Lebar konten 120 mm  (46 + 20 + 27 + 27 mm, sama dengan tabel barang di Word)
#   - Font Courier New, ukuran font sama dengan Word
# Isi nota berupa SEL BENERAN (bukan gambar): JUMLAH = QTY x HARGA, TOTAL = SUM.
# ---------------------------------------------------------------
XL_FONT = "Courier New"
XL_COL_MM = [20, 26, 20, 27, 27]          # kolom A..E (total 120 mm)
XL_RP_FORMAT = '"Rp "#,##0'
XL_PAGE_AVAIL_PT = (105 - 8 - 3) / 25.4 * 72   # tinggi area cetak per halaman (pt)
XL_RED = "C62828"


def _xl_col_width(mm):
    px = mm / 25.4 * 96
    return round((px - 5) / 7, 2)


def _xl_prepare_logo(logo_path, w_mm=18.0, max_h_mm=19.0):
    """Siapkan logo: ukuran 18 mm (sama dengan Word), tapi dibatasi supaya
    tidak lebih tinggi dari area header (supaya tidak menimpa tabel)."""
    if not logo_path or not os.path.exists(logo_path):
        return None
    try:
        im = Image.open(logo_path).convert("RGBA")
    except Exception:
        return None
    iw, ih = im.size
    w, h = w_mm, w_mm * ih / iw
    if h > max_h_mm:
        h = max_h_mm
        w = h * iw / ih
    target_w = max(1, int(w / 25.4 * 300))
    target_h = max(1, int(target_w * ih / iw))
    im = im.resize((target_w, target_h), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue(), w, h


def _xl_style(ws, ref, value=None, size=8, bold=False, color="000000",
              h="left", v="center", wrap=False, shrink=False, fill=None, fmt=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = Font(name=XL_FONT, size=size, bold=bold, color=color)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap,
                            shrink_to_fit=shrink)
    if fill:
        c.fill = PatternFill("solid", start_color=fill, end_color=fill)
    if fmt:
        c.number_format = fmt
    return c


def _xl_box(ws, rng):
    thin = Side(style="thin", color="000000")
    for row in ws[rng]:
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def _xl_write_nota(ws, r0, nota, tgl, logo):
    """Tulis 1 nota mulai baris r0. Return (baris_terakhir, tinggi_total_pt)."""
    heights = {}

    def set_h(r, pt):
        ws.row_dimensions[r].height = pt
        heights[r] = pt

    # ---------- HEADER (4 baris): logo | judul + alamat | info ----------
    hdr_h = [15, 13, 13, 15]
    for i, hh in enumerate(hdr_h):
        set_h(r0 + i, hh)

    ws.merge_cells(f"B{r0}:C{r0}")
    _xl_style(ws, f"B{r0}", "AYAM SEGAR TUMPANG", size=11, bold=True,
              color=XL_RED, h="left", v="center", shrink=True)
    ws.merge_cells(f"B{r0 + 1}:C{r0 + 2}")
    _xl_style(ws, f"B{r0 + 1}", "Ds. Kambingan - Tumpang - Malang", size=7.5,
              color="5A5A5A", h="left", v="top", wrap=True)

    info = [
        f"No: {nota.get('no', '')}",
        f"Tanggal: {tgl}",
        f"Pembeli / Bakul: {nota['bakul']}",
        f"Group: {nota['group']}",
    ]
    for i, txt in enumerate(info):
        r = r0 + i
        ws.merge_cells(f"D{r}:E{r}")
        _xl_style(ws, f"D{r}", txt, size=9, h="right", v="center", shrink=True)

    if logo:
        png, w_mm, h_mm = logo
        img = XLImage(io.BytesIO(png))
        marker = AnchorMarker(col=0, colOff=int(0.5 * 36000),
                              row=r0 - 1, rowOff=int(0.3 * 36000))
        img.anchor = OneCellAnchor(
            _from=marker,
            ext=XDRPositiveSize2D(int(w_mm * 36000), int(h_mm * 36000)),
        )
        ws.add_image(img)

    # ---------- SPACER ----------
    r = r0 + 4
    set_h(r, 10)

    # ---------- TABEL BARANG ----------
    r += 1
    hdr_row = r
    set_h(r, 12)
    ws.merge_cells(f"A{r}:B{r}")
    _xl_style(ws, f"A{r}", "BARANG", bold=True, h="left", fill="EAEAEA")
    _xl_style(ws, f"C{r}", "QTY / KG", bold=True, h="center", fill="EAEAEA")
    _xl_style(ws, f"D{r}", "HARGA", bold=True, h="right", fill="EAEAEA")
    _xl_style(ws, f"E{r}", "JUMLAH", bold=True, h="right", fill="EAEAEA")
    _xl_box(ws, f"A{r}:E{r}")

    first_item = r + 1
    for item in nota["items"]:
        r += 1
        set_h(r, 12)
        kg = item["KG"]
        is_int = float(kg).is_integer()
        ws.merge_cells(f"A{r}:B{r}")
        _xl_style(ws, f"A{r}", item["Nama Barang"], h="left")
        _xl_style(ws, f"C{r}", int(kg) if is_int else float(kg), h="center",
                  fmt="0" if is_int else "0.00")
        _xl_style(ws, f"D{r}", float(item["Harga"]), h="right", fmt=XL_RP_FORMAT)
        _xl_style(ws, f"E{r}", f"=C{r}*D{r}", h="right", fmt=XL_RP_FORMAT)
        _xl_box(ws, f"A{r}:E{r}")
    last_item = r

    # ---------- TOTAL ----------
    r += 1
    set_h(r, 20)
    _xl_style(ws, f"C{r}", "TOTAL :", size=9, bold=True, h="right", v="center")
    ws.merge_cells(f"D{r}:E{r}")
    _xl_style(ws, f"D{r}", f"=SUM(E{first_item}:E{last_item})", size=12,
              bold=True, color=XL_RED, h="right", v="center", shrink=True,
              fmt=XL_RP_FORMAT)

    # ---------- PENERIMA / HORMAT KAMI ----------
    r += 1
    set_h(r, 14)
    ws.merge_cells(f"A{r}:C{r}")
    _xl_style(ws, f"A{r}", "Penerima,", size=7.5, h="center", v="bottom")
    ws.merge_cells(f"D{r}:E{r}")
    _xl_style(ws, f"D{r}", "Hormat Kami,", size=7.5, h="center", v="bottom")

    r += 1
    set_h(r, 36)
    ws.merge_cells(f"A{r}:C{r}")
    _xl_style(ws, f"A{r}", "( ........... )", size=7.5, h="center", v="bottom")
    ws.merge_cells(f"D{r}:E{r}")
    _xl_style(ws, f"D{r}", "( ........... )", size=7.5, h="center", v="bottom")

    return r, sum(heights.values())


def generate_excel_nota(notas, tgl, logo_path, sheet_title="NOTA"):
    """Bikin 1 file Excel, 1 sheet. `notas` = list of dict:
        {"bakul": str, "group": str, "no": str, "items": [ {Nama Barang, KG, Harga, Jumlah}, ... ]}
    Semua nota berjejer ke bawah; tiap nota = 1 halaman cetak (A6 landscape).
    Nota satuan = list berisi 1 nota."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.sheet_view.showGridLines = False

    for i, mm in enumerate(XL_COL_MM):
        ws.column_dimensions["ABCDE"[i]].width = _xl_col_width(mm)

    logo = _xl_prepare_logo(logo_path)

    r0 = 1
    last_row = 1
    n = len(notas)
    for idx, nota in enumerate(notas):
        end_row, h_pt = _xl_write_nota(ws, r0, nota, tgl, logo)
        last_row = end_row
        if idx < n - 1:
            # baris pemisah di bawah nota, lalu page break -> nota berikutnya halaman baru
            gap_h = max(0, min(12, XL_PAGE_AVAIL_PT - h_pt - 1))
            ws.row_dimensions[end_row + 1].height = max(gap_h, 1)
            ws.row_breaks.append(Break(id=end_row + 1))
            r0 = end_row + 2

    # ---------- PENGATURAN CETAK ----------
    ws.print_area = f"A1:E{last_row}"
    # Ukuran kertas SENGAJA tidak dikunci (jangan A6!). Form fisik 14 cm (arah kertas keluar)
    # x 9,5 cm harus dibuat sebagai Custom Form di printer, lalu dipilih di Excel/printer default.
    # Kalau mau dikunci ke ukuran standar Excel, isi kode kertasnya di sini, mis. ws.page_setup.paperSize = 70 (A6).
    ws.page_setup.orientation = "landscape"
    ws.page_setup.scale = 100               # jangan fit-to-page (manual page break harus jalan)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    ws.page_margins = PageMargins(
        left=1 / 25.4, right=8 / 25.4, top=8 / 25.4, bottom=3 / 25.4,
        header=0, footer=0,
    )
    ws.print_options.horizontalCentered = False

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


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
                                    "<h3 style='margin:0; color:#C62828; font-weight: bold;"
                                    " font-size:24px;'>AYAM SEGAR TUMPANG</h3><small>Ds. Kambingan"
                                    " - Tumpang - Kab. Malang</small>",
                                    unsafe_allow_html=True,
                                )
                            with col_r:
                                st.markdown(
                                    f"<div style='text-align:right;'><small><b>No:</b> {no_bakul}<br>"
                                    f"<b>Tanggal:</b> {tanggal_transaksi.strftime('%d-%m-%Y')}<br>"
                                    f"<b>Bakul:</b> {selected_bakul}<br>"
                                    f"<b>Group:</b> {selected_sheet}</small></div>",
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
                                rows_html += f"<tr><td>{item['Nama Barang']}</td><td style='text-align: center;'>{kg_str}</td><td style='text-align: right;'>Rp {item['Harga']:,.0f}</td><td style='text-align: right;'>Rp {item['Jumlah']:,.0f}</td></tr>".replace(
                                    ",", "."
                                )

                            st.markdown(
                                f"""
                                <table style="width:100%; border-collapse:collapse; margin-top:10px;">
                                    <thead><tr>
                                        <th style="border:1px solid #000; padding:6px; background:#f2f2f2;">BARANG</th>
                                        <th style="border:1px solid #000; padding:6px; background:#f2f2f2;">QTY / KG</th>
                                        <th style="border:1px solid #000; padding:6px; background:#f2f2f2;">HARGA</th>
                                        <th style="border:1px solid #000; padding:6px; background:#f2f2f2;">JUMLAH</th>
                                    </tr></thead>
                                    <tbody>{rows_html}</tbody>
                                </table>
                                <div style="text-align:right; margin-top:10px; font-weight:bold;">
                                    TOTAL: <span style="color:#C62828; font-size:22px; font-weight:bold;">Rp {total_bayar:,.0f}</span>
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
                            xl_data = generate_excel_nota(
                                notas=[
                                    {
                                        "bakul": selected_bakul,
                                        "group": selected_sheet,
                                        "no": no_bakul,
                                        "items": filtered_items,
                                    }
                                ],
                                tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                                logo_path=logo_filename,
                                sheet_title="NOTA",
                            )
                            st.download_button(
                                label="📊 Download Nota Excel (.xlsx)",
                                data=xl_data,
                                file_name=(
                                    f"Nota_{selected_bakul}_{tanggal_transaksi.strftime('%Y%m%d')}.xlsx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                ),
                                use_container_width=True,
                            )

            # ==========================================
            # TAB 2: EXPORT ALL NOTA (1 KOLOM BERURUTAN KE BAWAH)
            # ==========================================
            with tab_bulk:
                st.markdown("### 📄 Export All Nota ke 1 File Excel")
                st.caption("Centang nama bakul yang ingin diproses.")

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

                if st.button(
                    "🚀 Generate Excel All Nota (1 File, 1 Sheet)",
                    type="primary",
                    use_container_width=True,
                ):
                    semua_nota = []

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

                        # Cari kolom NO di sheet ini
                        no_col_sheet = next(
                            (c for c in df_s.columns if c in ["NO", "NO.", "NOMOR"]), None
                        )

                        for idx, row in df_filtered.iterrows():
                            nama_bakul = str(row[name_c]).strip()

                            # Ambil nomor urut dari Excel
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
                                semua_nota.append(
                                    {
                                        "bakul": nama_bakul,
                                        "group": sheet_name,
                                        "no": no_val,
                                        "items": filtered_items,
                                    }
                                )

                    if semua_nota:
                        output_excel = generate_excel_nota(
                            notas=semua_nota,
                            tgl=tanggal_transaksi.strftime("%d-%m-%Y"),
                            logo_path=logo_filename,
                            sheet_title="NOTA",
                        )

                        st.success(
                            f"✅ Berhasil menyusun {len(semua_nota)} nota berurutan ke bawah dalam 1 sheet Excel!"
                        )
                        st.download_button(
                            label="📥 Download Excel All Nota (.xlsx)",
                            data=output_excel,
                            file_name=(
                                "EXCEL_ALL_NOTA_"
                                f"{tanggal_transaksi.strftime('%Y%m%d')}.xlsx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            ),
                            use_container_width=True,
                        )
                    else:
                        st.warning(
                            "⚠️ Tidak ada nota yang diproses. Pastikan bakul yang dipilih"
                            " memiliki data transaksi."
                        )

        else:
            st.info("💡 Silakan upload file Rekap Excel untuk memproses nota.")
    else:
        st.info(f"Fitur untuk Nota {sub_menu} siap dikembangkan.")
else:
    st.info(f"Halaman {selected_menu} sedang dalam pengembangan.")
