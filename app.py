import streamlit as st
import pandas as pd
from datetime import datetime
import base64

# Konfigurasi Halaman
st.set_page_config(page_title="E-Nota Bakul Ayam Segar", layout="wide")

# CSS Styling (Termasuk trik agar Print bersih)
st.markdown("""
    <style>
        /* Sembunyikan elemen bawaan Streamlit saat Print */
        @media print {
            header, [data-testid="stSidebar"], .stFileUploader, .stDateInput, 
            div[data-baseweb="select"], .stButton, .print-hidden {
                display: none !important;
            }
            .block-container {
                padding: 0 !important;
                max-width: 100% !important;
            }
            .printable-area {
                display: block !important;
                width: 100% !important;
            }
        }
    </style>
""", unsafe_allow_html=True)

# Judul Utama (Disembunyikan saat print pakai div class print-hidden)
st.markdown('<div class="print-hidden">', unsafe_allow_html=True)
st.title("E-Nota Bakul Ayam Segar")
st.markdown('</div>', unsafe_allow_html=True)

# 1. Sidebar - Input Master Harga
st.sidebar.markdown("<h2>⚙️ Master Harga</h2>", unsafe_allow_html=True)
harga_glondong = st.sidebar.number_input("Harga Glondong / Kg", value=28500, step=500)
harga_jeroan = st.sidebar.number_input("Harga Jeroan", value=12000, step=500)
harga_usus = st.sidebar.number_input("Harga Usus B", value=16500, step=500)
harga_telur = st.sidebar.number_input("Harga Telur", value=269000, step=1000)
harga_peti = st.sidebar.number_input("Harga Peti", value=2000, step=100)
harga_box = st.sidebar.number_input("Harga Box", value=28500, step=500)

# 2. Input Data Transaksi
st.markdown('<div class="print-hidden">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload File Rekap Excel", type=["xlsx"])
tanggal_transaksi = st.date_input("Tanggal Transaksi", value=datetime.today())
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file:
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
    df = df[df[name_col].notna()]
    
    selected_bakul = st.selectbox("Pilih Nama Bakul", df[name_col].unique())
    
    if selected_bakul:
        row_bakul = df[df[name_col] == selected_bakul].iloc[0]
        
        # Helper fungsi ambil angka
        def get_val(col_name):
            try:
                val = row_bakul[col_name]
                return float(val) if pd.notna(val) else 0.0
            except: return 0.0

        st.markdown('<div class="print-hidden">', unsafe_allow_html=True)
        qty_peti = st.number_input("Jumlah Peti", value=0, step=1)
        qty_box = st.number_input("Jumlah Box (Manual)", value=0.0, step=0.1)
        st.markdown('</div>', unsafe_allow_html=True)

        qty_tonase = get_val(next((c for c in df.columns if 'TONASE' in c), ''))
        qty_jeroan = get_val(next((c for c in df.columns if 'JEROAN' in c), ''))
        qty_usus = get_val(next((c for c in df.columns if 'USUS' in c), ''))
        qty_telur = get_val(next((c for c in df.columns if 'TELUR' in c), ''))

        total_bayar = (qty_tonase * harga_glondong) + (qty_jeroan * harga_jeroan) + \
                      (qty_usus * harga_usus) + (qty_telur * harga_telur) + \
                      (qty_peti * harga_peti) + (qty_box * harga_box)

        # --- AREA YANG AKAN DI-PRINT ---
        st.markdown('<div class="printable-area">', unsafe_allow_html=True)
        st.markdown("<h3>🐔 AYAM SEGAR TUMPANG</h3>", unsafe_allow_html=True)
        st.write("Ds. Kambingan - Tumpang - Kab. Malang")
        
        c1, c2 = st.columns(2)
        c1.write(f"**Nama Bakul:** {selected_bakul}")
        c1.write(f"**Group:** {selected_sheet}")
        c2.write(f"**Tanggal:** {tanggal_transaksi.strftime('%d-%m-%Y')}")
        
        data = [
            {"Barang": "GLONDONG", "KG": qty_tonase, "Harga": harga_glondong, "Jml": qty_tonase * harga_glondong},
            {"Barang": "JEROAN", "KG": qty_jeroan, "Harga": harga_jeroan, "Jml": qty_jeroan * harga_jeroan},
            {"Barang": "USUS B", "KG": qty_usus, "Harga": harga_usus, "Jml": qty_usus * harga_usus},
            {"Barang": "TELUR", "KG": qty_telur, "Harga": harga_telur, "Jml": qty_telur * harga_telur},
            {"Barang": "PETI", "KG": qty_peti, "Harga": harga_peti, "Jml": qty_peti * harga_peti},
            {"Barang": "BOX", "KG": qty_box, "Harga": harga_box, "Jml": qty_box * harga_box}
        ]
        
        df_nota = pd.DataFrame([d for d in data if d['KG'] > 0])
        df_nota['Harga'] = df_nota['Harga'].map("Rp {:,.0f}".format)
        df_nota['Jml'] = df_nota['Jml'].map("Rp {:,.0f}".format)
        
        st.table(df_nota)
        st.subheader(f"TOTAL: Rp {total_bayar:,.0f}")
        st.markdown('</div>', unsafe_allow_html=True) # Tutup printable-area

        # Tombol Cetak
        st.markdown('<div class="print-hidden">', unsafe_allow_html=True)
        if st.button("🖨️ Cetak / Print Nota"):
            st.write("Tekan Ctrl+P atau Cmd+P untuk menyimpan sebagai PDF.")
        st.markdown('</div>', unsafe_allow_html=True)