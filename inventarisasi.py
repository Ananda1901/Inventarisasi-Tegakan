# Inventarisasi Tegakan dengan Excel Download
import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook

st.set_page_config(page_title="Inventarisasi Tegakan", layout="wide")
st.markdown("""
    <style>
    .main {background-color: #e6f4ea;}
    div[data-testid="stSidebar"] {background-color: #a9d6ae;}
    h1 {color: #2e7d32;}
    </style>
""", unsafe_allow_html=True)

st.title("🌳 Inventarisasi Tegakan")

st.subheader("Input Data Pohon")

with st.form("form_inventaris"):
    col1, col2 = st.columns(2)
    with col1:
        tanggal = st.date_input("Tanggal")
        lokasi = st.text_input("Lokasi")
        jenis = st.text_input("Jenis Tanaman")
        diameter = st.number_input("Diameter (cm)", min_value=0.0, step=0.1)
        tinggi = st.number_input("Tinggi Total (m)", min_value=0.0, step=0.1)
        tbbc = st.number_input("Tinggi Bebas Cabang (m)", min_value=0.0, step=0.1)
    with col2:
        tajuk_utara = st.number_input("Lebar Tajuk Utara (m)", min_value=0.0, step=0.1)
        tajuk_selatan = st.number_input("Lebar Tajuk Selatan (m)", min_value=0.0, step=0.1)
        tajuk_barat = st.number_input("Lebar Tajuk Barat (m)", min_value=0.0, step=0.1)
        tajuk_timur = st.number_input("Lebar Tajuk Timur (m)", min_value=0.0, step=0.1)

    submitted = st.form_submit_button("Tambah Data")

# Inisialisasi session state
if "data" not in st.session_state:
    st.session_state.data = []

# Tambahkan data
if submitted:
    rata_tajuk = (tajuk_utara + tajuk_selatan + tajuk_barat + tajuk_timur) / 4
    volume = 0.00007854 * (diameter ** 2) * tinggi
    st.session_state.data.append({
        "Tanggal": tanggal,
        "Lokasi": lokasi,
        "Jenis": jenis,
        "Diameter (cm)": diameter,
        "Tinggi Total (m)": tinggi,
        "TBBC (m)": tbbc,
        "Tajuk U": tajuk_utara,
        "Tajuk S": tajuk_selatan,
        "Tajuk B": tajuk_barat,
        "Tajuk T": tajuk_timur,
        "Rata-rata Tajuk (m)": rata_tajuk,
        "Volume (m3)": volume
    })

# Tampilkan data
st.subheader("📋 Data Terekam")

if st.session_state.data:
    df = pd.DataFrame(st.session_state.data)

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        key="data_editor"
    )

    st.session_state.data = edited_df.to_dict("records")

    # Pilih baris untuk dihapus
    row_to_delete = st.number_input("Masukkan nomor baris yang ingin dihapus (mulai dari 0)", min_value=0, max_value=len(st.session_state.data)-1, step=1)
    if st.button("Hapus Baris"):
        st.session_state.data.pop(row_to_delete)
        st.experimental_rerun()

    # Simpan ke Excel
    df = pd.DataFrame(st.session_state.data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data Inventarisasi')
    processed_data = output.getvalue()

    st.download_button(
        label="⬇️ Unduh Data Excel",
        data=processed_data,
        file_name="data_inventarisasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Belum ada data dimasukkan.")
