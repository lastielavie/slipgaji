import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Aplikasi Slip Gaji", layout="wide")

st.title("Aplikasi Pemroses Gaji & Slip Karyawan")

# 1. Unggah File Data Gaji
uploaded_file = st.sidebar.file_uploader("Upload File Gaji (Excel/CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    # Membaca data berdasarkan format file
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("Pilih Sheet", xls.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    st.subheader("1. Preview Data Mentah")
    st.dataframe(df.head())

    # Deteksi kolom secara otomatis
    all_columns = df.columns.tolist()
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    text_cols = [c for c in all_columns if c not in num_cols]

    # Menentukan default kolom pendapatan dan potongan berdasar keyword
    default_pendapatan = [c for c in num_cols if any(k in c.lower() for k in ['gaji', 'tunjangan', 'bonus', 'insentif', 'lembur'])]
    default_potongan = [c for c in num_cols if any(k in c.lower() for k in ['potongan', 'bpjs', 'pajak', 'pph', 'denda', 'kasbon'])]

    st.sidebar.header("Pemetaan Kolom")
    id_col = st.sidebar.selectbox("Kolom Nama / ID Karyawan", options=all_columns, index=0)
    col_pendapatan = st.sidebar.multiselect("Kolom Pendapatan (Penambah)", options=num_cols, default=default_pendapatan)
    col_potongan = st.sidebar.multiselect("Kolom Potongan (Pengurang)", options=num_cols, default=default_potongan)

    # 2. Pembersihan Data (Handling NaN)
    df_clean = df.copy()
    df_clean[num_cols] = df_clean[num_cols].fillna(0)

    # 3. Hitung Pendapatan & Potongan Standar
    df_clean['Total Pendapatan'] = df_clean[col_pendapatan].sum(axis=1) if col_pendapatan else 0
    df_clean['Total Potongan'] = df_clean[col_potongan].sum(axis=1) if col_potongan else 0

    # 4. Penanganan Kasus Khusus
    st.sidebar.header("Logika Kasus Khusus")
    enable_custom_rule = st.sidebar.checkbox("Aktifkan Penyesuaian Kasus Khusus")

    df_clean['Penyesuaian Khusus'] = 0

    if enable_custom_rule:
        if text_cols:
            kriteria_col = st.sidebar.selectbox("Pilih Kolom Acuan Kriteria", options=text_cols)
            nilai_kriteria = st.sidebar.text_input("Nilai Kriteria (misal: 'Bonus Proyek' / 'Resign')", "")
            nominal_penyesuaian = st.sidebar.number_input(
                "Nominal Penyesuaian (Gunakan angka negatif untuk potongan)", value=0
            )

            if nilai_kriteria:
                # Pencocokan string acuan tanpa mempedulikan huruf besar/kecil
                mask = df_clean[kriteria_col].astype(str).str.strip().str.lower() == nilai_kriteria.strip().lower()
                df_clean['Penyesuaian Khusus'] = np.where(mask, nominal_penyesuaian, 0)
                
                # Memasukkan penyesuaian ke total pendapatan / potongan
                df_clean['Total Pendapatan'] += np.where(df_clean['Penyesuaian Khusus'] > 0, df_clean['Penyesuaian Khusus'], 0)
                df_clean['Total Potongan'] += np.where(df_clean['Penyesuaian Khusus'] < 0, np.abs(df_clean['Penyesuaian Khusus']), 0)
        else:
            st.sidebar.warning("Tidak ditemukan kolom teks untuk kriteria kasus khusus.")

    # 5. Hitung Gaji Bersih (Take Home Pay)
    df_clean['Gaji Bersih'] = df_clean['Total Pendapatan'] - df_clean['Total Potongan']

    st.subheader("2. Hasil Rekapitulasi Gaji")
    st.dataframe(df_clean)

    # Export Data Rekap ke Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_clean.to_excel(writer, index=False, sheet_name='Rekap Gaji')
    processed_data = output.getvalue()

    st.download_button(
        label="Download Rekap Gaji (Excel)",
        data=processed_data,
        file_name="rekap_gaji_diproses.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 6. Tampilan Detail Slip Gaji Individu
    st.markdown("---")
    st.subheader("3. Preview Slip Gaji Per Karyawan")

    selected_emp = st.selectbox("Pilih Karyawan", options=df_clean[id_col].unique())
    emp_row = df_clean[df_clean[id_col] == selected_emp].iloc[0]

    st.markdown(f"### **SLIP GAJI: {emp_row[id_col]}**")
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("#### **PENDAPATAN**")
        for c in col_pendapatan:
            st.write(f"• {c}: Rp {emp_row[c]:,.0f}")
        if emp_row['Penyesuaian Khusus'] > 0:
            st.write(f"• Penyesuaian Khusus: Rp {emp_row['Penyesuaian Khusus']:,.0f}")
        st.markdown(f"**Total Pendapatan: Rp {emp_row['Total Pendapatan']:,.0f}**")

    with col_right:
        st.markdown("#### **POTONGAN**")
        for c in col_potongan:
            st.write(f"• {c}: Rp {emp_row[c]:,.0f}")
        if emp_row['Penyesuaian Khusus'] < 0:
            st.write(f"• Penyesuaian Khusus (Potongan): Rp {abs(emp_row['Penyesuaian Khusus']):,.0f}")
        st.markdown(f"**Total Potongan: Rp {emp_row['Total Potongan']:,.0f}**")

    st.markdown(f"### **GAJI BERSIH (TAKE HOME PAY): Rp {emp_row['Gaji Bersih']:,.0f}**")

else:
    st.info("Silakan unggah file Excel atau CSV rekap gaji Anda di bilah samping (sidebar) sebelah kiri.")
