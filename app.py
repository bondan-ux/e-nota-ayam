import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

st.set_page_config(page_title="E-Nota Bakul Ayam Segar", layout="centered")

st.title("E-Nota Bakul Ayam Segar")

# --- 1. INPUT DATA HEADER ---
col_h1, col_h2 = st.columns(2)
with col_h1:
    toko_name = "AYAM SEGAR TUMPANG"
    alamat = "Ds. Kambingan - Tumpang - Kab. Malang"
    st.write(f"**{toko_name}**")
    st.caption(alamat)

with col_h2:
    tgl = st.date_input("Tanggal", pd.to_datetime("today"))
    pembeli = st.text_input("Pembeli / Bakul", "IDA")
    group = st.text_input("Group", "FARHAN")

st.divider()

# --- 2. INPUT ITEM / TABEL ---
st.subheader("Detail Barang")

# Form input barang sederhana
if "items" not in st.session_state:
    st.session_state.items = [
        {"BARANG": "GLONDONG", "QTY / KG": 155, "HARGA": 29500},
        {"BARANG": "JEROAN", "QTY / KG": 5, "HARGA": 10000},
        {"BARANG": "BIAYA KRESEK", "QTY / KG": 1, "HARGA": 7000},
    ]

# Tampilan editor data agar bisa diubah dinamis di web
edited_df = st.data_editor(
    pd.DataFrame(st.session_state.items),
    num_rows="dynamic",
    use_container_width=True,
)

# Hitung JUMLAH & TOTAL pembayaran (Logika Utama)
if not edited_df.empty:
    edited_df["JUMLAH"] = edited_df["QTY / KG"] * edited_df["HARGA"]
    total_bayar = edited_df["JUMLAH"].sum()
else:
    total_bayar = 0

st.markdown(f"### **TOTAL: Rp {total_bayar:,.0f}**".replace(",", "."))


# --- 3. FUNGSI GENERATE FILE WORD (.DOCX) ---
def generate_word_nota(df_data, total_val, tgl_str, pembeli_str, group_str):
    doc = docx.Document()

    # Header Nota
    p_head = doc.add_paragraph()
    p_head.add_run(f"{toko_name}\n").bold = True
    p_head.add_run(f"{alamat}\n\n")
    p_head.add_run(
        f"Tanggal: {tgl_str}\nPembeli / Bakul: {pembeli_str}\nGroup: {group_str}"
    )

    # Tabel Barang
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    headers = ["QTY / KG", "BARANG", "HARGA", "JUMLAH"]
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True

    # Isi Data Tabel
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row["QTY / KG"])
        row_cells[1].text = str(row["BARANG"])
        row_cells[2].text = f"Rp {row['HARGA']:,.0f}".replace(",", ".")
        row_cells[3].text = f"Rp {row['JUMLAH']:,.0f}".replace(",", ".")

    # Total Pembayaran (Langsung ambil dari total_val)
    p_total = doc.add_paragraph()
    p_total.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_tot = p_total.add_run(
        f"\nTOTAL : Rp {total_val:,.0f}".replace(",", ".")
    )
    run_tot.font.bold = True
    run_tot.font.size = Pt(14)
    run_tot.font.color.rgb = RGBColor(192, 0, 0)

    # Save to buffer
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target


# --- 4. FUNGSI GENERATE GAMBAR (.PNG) ---
def generate_image_nota(df_data, total_val, tgl_str, pembeli_str, group_str):
    img = Image.new("RGB", (800, 600), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Menggunakan font bawaan/default
    try:
        font = ImageFont.truetype("arial.ttf", 16)
        font_bold = ImageFont.truetype("arialbd.ttf", 18)
    except IOError:
        font = font_bold = ImageFont.load_default()

    # Draw Header
    draw.text(
        (30, 20),
        f"{toko_name}\n{alamat}",
        fill=(180, 0, 0),
        font=font_bold,
    )
    draw.text(
        (500, 20),
        f"Tanggal: {tgl_str}\nPembeli: {pembeli_str}\nGroup: {group_str}",
        fill=(0, 0, 0),
        font=font,
    )

    # Draw Table Header
    y_pos = 120
    draw.rectangle([(30, y_pos), (770, y_pos + 30)], outline="black", fill=(240, 240, 240))
    draw.text((40, y_pos + 5), "QTY/KG", fill="black", font=font_bold)
    draw.text((150, y_pos + 5), "BARANG", fill="black", font=font_bold)
    draw.text((450, y_pos + 5), "HARGA", fill="black", font=font_bold)
    draw.text((620, y_pos + 5), "JUMLAH", fill="black", font=font_bold)

    # Draw Data Rows
    y_pos += 30
    for _, row in df_data.iterrows():
        draw.rectangle([(30, y_pos), (770, y_pos + 30)], outline="black")
        draw.text((40, y_pos + 5), str(row["QTY / KG"]), fill="black", font=font)
        draw.text((150, y_pos + 5), str(row["BARANG"]), fill="black", font=font)
        draw.text(
            (450, y_pos + 5),
            f"Rp {row['HARGA']:,.0f}".replace(",", "."),
            fill="black",
            font=font,
        )
        draw.text(
            (620, y_pos + 5),
            f"Rp {row['JUMLAH']:,.0f}".replace(",", "."),
            fill="black",
            font=font,
        )
        y_pos += 30

    # Draw TOTAL (Menggunakan total_val kalkulasi terbaru)
    y_pos += 20
    draw.text(
        (450, y_pos),
        "TOTAL :",
        fill="black",
        font=font_bold,
    )
    draw.text(
        (550, y_pos),
        f"Rp {total_val:,.0f}".replace(",", "."),
        fill=(192, 0, 0),
        font=font_bold,
    )

    # Save to buffer
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format="PNG")
    img_byte_arr.seek(0)
    return img_byte_arr


# --- 5. TOMBOL DOWNLOAD ---
st.divider()
col_btn1, col_btn2 = st.columns(2)

tgl_formatted = tgl.strftime("%d-%m-%Y")

with col_btn1:
    docx_file = generate_word_nota(
        edited_df, total_bayar, tgl_formatted, pembeli, group
    )
    st.download_button(
        label="📄 Download Nota Word (.docx)",
        data=docx_file,
        file_name=f"Nota_{pembeli}_{tgl_formatted}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

with col_btn2:
    img_file = generate_image_nota(
        edited_df, total_bayar, tgl_formatted, pembeli, group
    )
    st.download_button(
        label="🖼️ Download Nota Gambar (.png)",
        data=img_file,
        file_name=f"Nota_{pembeli}_{tgl_formatted}.png",
        mime="image/png",
    )