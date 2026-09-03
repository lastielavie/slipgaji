import io
import re
import zipfile
from pathlib import Path
import pandas as pd
import streamlit as st

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

st.set_page_config(page_title="Cetak Slip Gaji Teknisi (Fleksibel)", layout="wide", page_icon="🧾")

DIR_ASET = Path(__file__).parent / "assets"
LOGO_MADINAH = DIR_ASET / "logo-madinah.png"
LOGO_MFLASH = DIR_ASET / "logo-mflash.png"

# ---------------------------------------------------------------------------
# Utilitas & Parser
# ---------------------------------------------------------------------------
def rupiah(v) -> str:
    try:
        if pd.isna(v):
            v = 0.0
        v = float(v)
    except (TypeError, ValueError):
        v = 0.0
    s = f"{abs(v):,.0f}".replace(",", ".")
    return ("Rp (" + s + ")") if v < 0 else ("Rp " + s)


def _nama_berkas_aman(teks, cadangan='TANPA-NAMA'):
    aman = re.sub(r'[^A-Za-z0-9 _.-]', '-', str(teks)).strip(' .-')
    return (aman[:80] or cadangan)


def get_dynamic_categories(columns):
    EXCLUDE_OMZET = {'total', 'jasa (total)', 'jasa', '(total)', 'keseluruhan'}
    EXCLUDE_BH = {'(aturan)', 'aturan', 'total', 'nett', 'net', '(total)', 'keseluruhan'}
    
    categories = []
    for c in columns:
        c_strip = str(c).strip()
        c_low = c_strip.lower()
        if c_low.startswith('omzet '):
            cat = c_strip[6:].strip()
            if cat.lower() not in EXCLUDE_OMZET and cat not in categories:
                categories.append(cat)
        elif c_low.startswith('bagi hasil '):
            cat = c_strip[11:].strip()
            if cat.lower() not in EXCLUDE_BH and cat not in categories:
                categories.append(cat)
    return categories


def get_dynamic_potongan_cols(columns, categories):
    known_summary = [
        'nama teknisi', 'teknisi', 'cabang', 'baris',
        'omzet jasa (total)', 'omzet total', 'omzet jasa',
        'bagi hasil (aturan)', 'bagi hasil total', 'bagi_hasil',
        'total bagi hasil (aturan)', 'total bagi hasil',
        'pembanding 30%', 'selisih', 'efektif %',
        'total potongan', 'total_potongan',
        'gaji teknisi', 'gaji_teknisi',
        'nett bagi hasil', 'nett_bagi_hasil', 'nett',
        'cadangan 7 tahun / bulan', 'cadangan 7 tahun', 'total cadangan 7 tahun'
    ]
    
    cat_cols_lower = set()
    for cat in categories:
        cat_cols_lower.add(f"omzet {cat}".lower())
        cat_cols_lower.add(f"bagi hasil {cat}".lower())
        
    pot_cols = []
    for c in columns:
        c_str = str(c).strip()
        c_low = c_str.lower()
        if c_low in known_summary or c_low in cat_cols_lower:
            continue
        if c_low.startswith('unnamed:') or c_low.startswith('pembanding '):
            continue
        pot_cols.append(c_str)
        
    return pot_cols


def parse_excel_file(file_bytes, file_name):
    periode_str = ""
    sheet_used = ""

    if file_name.lower().endswith(('.csv', '.csv.gz')):
        df_raw = pd.read_csv(io.BytesIO(file_bytes))
        sheet_used = "CSV"
    else:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        target_sheet = 'RAW' if 'RAW' in xls.sheet_names else xls.sheet_names[0]
        for s in xls.sheet_names:
            if 'RAW' in s.upper():
                target_sheet = s
                break
        sheet_used = target_sheet
        
        df_temp = pd.read_excel(xls, sheet_name=target_sheet, header=None)
        header_row_idx = 0
        
        for i in range(min(10, len(df_temp))):
            row_str_vals = [str(v).strip() for v in df_temp.iloc[i].values]
            row_text = " ".join(row_str_vals)
            
            if "Periode:" in row_text and not periode_str:
                parts = row_text.split("Periode:")
                if len(parts) > 1:
                    periode_str = parts[1].split("·")[0].strip()
            
            if any('NAMA TEKNISI' in str(v).upper() for v in row_str_vals):
                header_row_idx = i
                break
                
        df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=header_row_idx)

    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    col_map = {}
    for c in df_raw.columns:
        cu = c.upper()
        if cu in ['NAMA TEKNISI', 'TEKNISI']:
            col_map[c] = 'TEKNISI'
        elif cu in ['CABANG']:
            col_map[c] = 'CABANG'
        elif cu in ['BAGI HASIL (ATURAN)', 'BAGI HASIL', 'BAGI_HASIL', 'BRUTO']:
            col_map[c] = 'BAGI_HASIL'
        elif cu in ['TOTAL POTONGAN', 'TOTAL_POTONGAN']:
            col_map[c] = 'TOTAL_POTONGAN'
        elif cu in ['NETT BAGI HASIL', 'NETT']:
            col_map[c] = 'NETT_BAGI_HASIL'
            
    df_data = df_raw.rename(columns=col_map)
    
    if 'TEKNISI' in df_data.columns:
        df_data = df_data[df_data['TEKNISI'].notna()]
        df_data = df_data[~df_data['TEKNISI'].astype(str).str.strip().str.upper().isin(
            ['TOTAL', 'NAN', 'NONE', '', 'UNNAMED: 0']
        )]
        
    return df_data, periode_str, sheet_used


def extract_slip_data_from_row(row, columns):
    categories = get_dynamic_categories(columns)
    pot_cols = get_dynamic_potongan_cols(columns, categories)
    
    per_kual = []
    for cat in categories:
        raw_omzet = row.get(f"Omzet {cat}", 0)
        raw_bh = row.get(f"Bagi Hasil {cat}", 0)
        
        omzet = 0.0 if pd.isna(raw_omzet) else float(raw_omzet or 0)
        bh = 0.0 if pd.isna(raw_bh) else float(raw_bh or 0)
        
        if omzet > 0 or bh != 0:
            akad_pct = bh / omzet if omzet > 0 else 0.0
            per_kual.append((cat, omzet, akad_pct, bh))
            
    topup_val = 0.0
    topup_raw = None
    for c in row.index:
        if str(c).strip().lower() == "gaji teknisi":
            topup_raw = row.get(c)
            break
            
    if topup_raw is not None:
        topup_val = 0.0 if pd.isna(topup_raw) else float(topup_raw or 0)
        if topup_val != 0:
            per_kual.append(("Topup Gaji", 0.0, 0.0, topup_val))
            
    raw_bruto = row.get('BAGI_HASIL', row.get('Bagi Hasil (Aturan)', 0))
    bruto = 0.0 if pd.isna(raw_bruto) else float(raw_bruto or 0)
    if not bruto and per_kual:
        bruto = sum(x[3] for x in per_kual)
    elif bruto != 0:
        bruto += topup_val
        
    pot = []
    for col in pot_cols:
        val_raw = row.get(col, 0)
        val = 0.0 if pd.isna(val_raw) else float(val_raw or 0)
        pot.append((col, val))
        
    raw_total_pot = row.get('TOTAL_POTONGAN', row.get('Total Potongan', 0))
    total_pot = 0.0 if pd.isna(raw_total_pot) else float(raw_total_pot or 0)
    if not total_pot and pot:
        total_pot = sum(x[1] for x in pot)
        
    nett = bruto - total_pot
        
    cadangan_bulan = 0.0
    cadangan_total = 0.0
    for c in row.index:
        c_low = str(c).strip().lower()
        if c_low == "cadangan 7 tahun / bulan":
            val = row.get(c)
            cadangan_bulan = 0.0 if pd.isna(val) else float(val or 0)
        elif c_low == "total cadangan 7 tahun":
            val = row.get(c)
            cadangan_total = 0.0 if pd.isna(val) else float(val or 0)
    
    return per_kual, bruto, pot, total_pot, nett, categories, pot_cols, cadangan_bulan, cadangan_total


# ---------------------------------------------------------------------------
# Generator PDF
# ---------------------------------------------------------------------------
def _gambar_slip(c, lebar, tinggi, nama, cabang, periode, angka, catatan):
    per_kual, bruto, pot, total_pot, nett, cadangan_bulan, cadangan_total = angka
    m = 18 * mm
    y = tinggi - 14 * mm

    if LOGO_MADINAH.exists():
        c.drawImage(ImageReader(str(LOGO_MADINAH)), m, y - 20 * mm, width=20 * mm, height=20 * mm, mask='auto')
    if LOGO_MFLASH.exists():
        c.drawImage(ImageReader(str(LOGO_MFLASH)), lebar - m - 34 * mm, y - 20 * mm, width=34 * mm, height=24 * mm, mask='auto', preserveAspectRatio=True, anchor='ne')
    y -= 26 * mm

    c.setFillColorRGB(0.12, 0.22, 0.39)
    c.setFont('Helvetica-Bold', 13)
    c.drawCentredString(lebar / 2, y, 'SLIP BAGI HASIL TEKNISI MADINAH FLASH')
    y -= 4 * mm
    c.setLineWidth(1.2)
    c.line(m, y, lebar - m, y)
    y -= 9 * mm

    c.setFillColorRGB(0, 0, 0)
    c.setFont('Helvetica', 9.5)
    for label, isi in [('Nama', nama), ('Jabatan', 'Teknisi'), ('Divisi', f'MFlash — {cabang}'), ('Periode', periode)]:
        c.setFont('Helvetica-Bold', 9.5)
        c.drawString(m, y, label)
        c.setFont('Helvetica', 9.5)
        c.drawString(m + 24 * mm, y, f': {isi}')
        y -= 5.4 * mm
    y -= 3 * mm

    def judul_tabel(teks, kolom_kanan=True):
        nonlocal y
        c.setFillColorRGB(0.12, 0.22, 0.39)
        c.rect(m, y - 5.6 * mm, lebar - 2 * m, 5.6 * mm, stroke=0, fill=1)
        c.setFillColorRGB(1, 1, 1)
        c.setFont('Helvetica-Bold', 8.5)
        c.drawString(m + 2 * mm, y - 4 * mm, teks)
        if kolom_kanan:
            c.drawRightString(lebar - m - 46 * mm, y - 4 * mm, 'OMZET')
            c.drawRightString(lebar - m - 30 * mm, y - 4 * mm, 'AKAD')
            c.drawRightString(lebar - m - 2 * mm, y - 4 * mm, 'BAGI HASIL')
        else:
            c.drawRightString(lebar - m - 2 * mm, y - 4 * mm, 'JUMLAH')
        c.setFillColorRGB(0, 0, 0)
        y -= 9 * mm

    judul_tabel('PENDAPATAN PER KUALIFIKASI')
    c.setFont('Helvetica', 8.5)
    if not per_kual:
        c.drawString(m + 2 * mm, y, '(rincian kualifikasi tidak ditemukan)')
        y -= 5 * mm
    for lbl, omzet, akad, bh in per_kual:
        c.drawString(m + 2 * mm, y, str(lbl))
        c.drawRightString(lebar - m - 46 * mm, y, rupiah(omzet))
        c.drawRightString(lebar - m - 30 * mm, y, f'{akad*100:.1f}'.replace('.', ',') + '%')
        c.drawRightString(lebar - m - 2 * mm, y, rupiah(bh))
        y -= 4.8 * mm

    y -= 1 * mm
    c.setLineWidth(0.6)
    c.line(lebar - m - 52 * mm, y + 1.5 * mm, lebar - m, y + 1.5 * mm)
    y -= 2.5 * mm
    c.setFont('Helvetica-Bold', 9)
    c.drawString(m + 2 * mm, y, 'Total Bruto Bagi Hasil')
    c.drawRightString(lebar - m - 2 * mm, y, rupiah(bruto))
    y -= 8 * mm

    judul_tabel('POTONGAN', kolom_kanan=False)
    c.setFont('Helvetica', 8.5)
    
    has_pot = False
    
    # ---------------------------------------------------------
    # HANYA PISAHKAN "TABUNGAN RUMAH" DARI LIST POTONGAN UTAMA
    # ---------------------------------------------------------
    pot_tampil = []
    tabungan_rumah_val = 0.0
    for label, nilai in pot:
        lbl_low = str(label).strip().lower()
        if lbl_low == "tabungan rumah":
            tabungan_rumah_val += nilai
        else:
            pot_tampil.append((label, nilai))
            
    # Tampilkan potongan sisanya (termasuk simpanan/tabungan lain dan nilai 0)
    for label, nilai in pot_tampil:
        c.drawString(m + 2 * mm, y, str(label))
        c.drawRightString(lebar - m - 2 * mm, y, rupiah(nilai))
        y -= 4.8 * mm
        has_pot = True
        
    if not has_pot:
        c.drawString(m + 2 * mm, y, '(tidak ada potongan)')
        y -= 4.8 * mm

    c.setLineWidth(0.6)
    c.line(lebar - m - 52 * mm, y + 1.5 * mm, lebar - m, y + 1.5 * mm)
    y -= 2.5 * mm
    c.setFont('Helvetica-Bold', 9)
    c.drawString(m + 2 * mm, y, 'Total Potongan')
    c.drawRightString(lebar - m - 2 * mm, y, rupiah(total_pot))
    y -= 8 * mm

    c.setFillColorRGB(0.86, 0.92, 0.84)
    c.rect(m, y - 3 * mm, lebar - 2 * m, 7.5 * mm, stroke=0, fill=1)
    c.setFillColorRGB(0.05, 0.35, 0.15)
    c.setFont('Helvetica-Bold', 10.5)
    c.drawString(m + 2 * mm, y, 'NETT BAGI HASIL')
    c.drawRightString(lebar - m - 2 * mm, y, rupiah(nett))
    c.setFillColorRGB(0, 0, 0)
    y -= 10 * mm

    # ---------------------------------------------------------
    # TAMPILKAN CADANGAN & TABUNGAN RUMAH DI BAWAH (HANYA JIKA != 0)
    # ---------------------------------------------------------
    if cadangan_bulan != 0 or cadangan_total != 0 or tabungan_rumah_val != 0:
        c.setFont('Helvetica', 8.5)
        if cadangan_bulan != 0:
            c.drawString(m + 2 * mm, y, 'Cadangan 7 Tahun / bulan')
            c.drawRightString(lebar - m - 2 * mm, y, rupiah(cadangan_bulan))
            y -= 4.5 * mm
        if cadangan_total != 0:
            c.drawString(m + 2 * mm, y, 'Total Cadangan 7 Tahun')
            c.drawRightString(lebar - m - 2 * mm, y, rupiah(cadangan_total))
            y -= 4.5 * mm
        # Tabungan Rumah hanya tampil di bawah jika nilainya tidak 0
        if tabungan_rumah_val != 0:
            c.drawString(m + 2 * mm, y, 'Tabungan Rumah')
            c.drawRightString(lebar - m - 2 * mm, y, rupiah(tabungan_rumah_val))
            y -= 4.5 * mm
        y -= 2 * mm
    else:
        y -= 2 * mm

    c.setFont('Helvetica-Bold', 8)
    c.drawString(m, y, 'Catatan')
    y -= 2.5 * mm
    tinggi_kotak = 16 * mm
    c.setLineWidth(0.6)
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.rect(m, y - tinggi_kotak, lebar - 2 * m, tinggi_kotak, stroke=1, fill=0)
    c.setFont('Helvetica', 8)
    baris_catatan = str(catatan or '').splitlines()
    yy = y - 4.5 * mm
    for baris in baris_catatan[:4]:
        c.drawString(m + 2 * mm, yy, baris[:110])
        yy -= 3.8 * mm
    y -= tinggi_kotak + 10 * mm

    c.setStrokeColorRGB(0, 0, 0)
    c.setFont('Helvetica', 8)
    for x, teks in ((m + 8 * mm, 'Teknisi'), (lebar / 2 - 12 * mm, 'Kepala Cabang'), (lebar - m - 40 * mm, 'Finance')):
        c.line(x, y, x + 32 * mm, y)
        c.drawCentredString(x + 16 * mm, y - 4 * mm, teks)


def generate_zip_slips(df_data, periode_txt, catatan_slip, zip_per_cabang=False):
    buf = io.BytesIO()
    ringkas = []
    
    cabang_col = 'CABANG' if 'CABANG' in df_data.columns else df_data.columns[1]
    
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as luar:
        for cab, grp in df_data.groupby(cabang_col):
            if grp.empty:
                continue
            cab_str = str(cab).strip()
            folder = _nama_berkas_aman(cab_str, 'CABANG')
            berkas = []
            
            for _, row in grp.iterrows():
                nama = str(row['TEKNISI']).strip()
                if not nama or nama.upper() in ['TOTAL', 'NAN', 'NONE', 'TIDAK ADA TEKNISI']:
                    continue
                    
                angka_data = extract_slip_data_from_row(row, df_data.columns)
                angka = angka_data[:5] + (angka_data[7], angka_data[8])
                
                pdf_buf = io.BytesIO()
                lebar, tinggi = A4
                c = canvas.Canvas(pdf_buf, pagesize=A4)
                c.setTitle(f'Slip Bagi Hasil — {nama} ({cab_str})')
                c.setAuthor('Madinah Flash')
                
                _gambar_slip(c, lebar, tinggi, nama, cab_str, periode_txt, angka, catatan_slip)
                c.showPage()
                c.save()
                pdf_buf.seek(0)
                
                berkas.append((f'{folder} - {_nama_berkas_aman(nama)}.pdf', pdf_buf.getvalue()))
                
            if not berkas:
                continue

            if zip_per_cabang:
                dalam = io.BytesIO()
                with zipfile.ZipFile(dalam, 'w', zipfile.ZIP_DEFLATED) as z2:
                    for nm, isi in berkas:
                        z2.writestr(nm, isi)
                luar.writestr(f'{folder}.zip', dalam.getvalue())
            else:
                for nm, isi in berkas:
                    luar.writestr(f'{folder}/{nm}', isi)
            ringkas.append({'Cabang': cab_str, 'Jumlah Slip': len(berkas)})
            
    buf.seek(0)
    return buf.getvalue(), pd.DataFrame(ringkas)

# ---------------------------------------------------------------------------
# Streamlit Web UI
# ---------------------------------------------------------------------------
st.title("🧾 Aplikasi Cetak Slip Gaji Teknisi (Fleksibel)")
st.caption("Generator PDF Slip Gaji Teknisi - Penyesuaian Omzet, Akad & Potongan Interaktif")

if not REPORTLAB_AVAILABLE:
    st.error("Library `reportlab` belum terinstal.")
    st.stop()

uploaded_file = st.file_uploader("Upload File Excel (sheet 'RAW')", type=['xlsx', 'xls', 'csv', 'gz'], key='main_uploader')

if uploaded_file is not None:
    try:
        if 'df_parsed' not in st.session_state or st.session_state.get('uploaded_filename') != uploaded_file.name:
            df_parsed, auto_periode, sheet_used = parse_excel_file(uploaded_file.getvalue(), uploaded_file.name)
            st.session_state['df_parsed'] = df_parsed
            st.session_state['auto_periode'] = auto_periode
            st.session_state['sheet_used'] = sheet_used
            st.session_state['uploaded_filename'] = uploaded_file.name
        
        df_parsed = st.session_state['df_parsed']
        auto_periode = st.session_state['auto_periode']
        sheet_used = st.session_state['sheet_used']
        
        detected_cats = get_dynamic_categories(df_parsed.columns)
        detected_pots = get_dynamic_potongan_cols(df_parsed.columns, detected_cats)
        
        st.success(f"Berhasil membaca sheet **'{sheet_used}'**! Ditemukan **{len(df_parsed)} data teknisi**.")
        
        col1, col2 = st.columns(2)
        with col1:
            periode_input = st.text_input("Label Periode Gaji", value=auto_periode if auto_periode else "24 Juli 2026 – 23 Agustus 2026")
            bentuk = st.radio("Format Pengelompokan File ZIP", ['Folder per cabang', 'ZIP per cabang'], horizontal=True)
        with col2:
            catatan_slip = st.text_area("Catatan pada Slip", value="", height=100, help="Dapat diisi catatan manual jika ada.")

        st.divider()

        # -------------------------------------------------------------------
        # Fitur Modifikasi Interaktif per Teknisi
        # -------------------------------------------------------------------
        st.subheader("⚙️ Adjust / Sesuaikan Data Omzet & Akad Teknisi")
        st.caption("Pilih nama teknisi dari daftar di bawah untuk mengubah angka Omzet, Bagi Hasil, atau Potongan secara langsung.")

        list_teknisi = df_parsed['TEKNISI'].tolist()
        selected_teknisi = st.selectbox("Pilih Teknisi untuk Disesuaikan:", list_teknisi)

        if selected_teknisi:
            idx = df_parsed[df_parsed['TEKNISI'] == selected_teknisi].index[0]
            row_data = df_parsed.loc[idx]

            st.write(f"**Mengedit Data:** `{selected_teknisi}` ({row_data.get('CABANG', '-')})")
            
            t1, t2 = st.tabs(["📊 Pendapatan / Kualifikasi", "💸 Potongan"])

            with t1:
                kual_data = []
                for cat in detected_cats:
                    omzet_val = float(row_data.get(f"Omzet {cat}", 0) or 0) if not pd.isna(row_data.get(f"Omzet {cat}", 0)) else 0.0
                    bh_val = float(row_data.get(f"Bagi Hasil {cat}", 0) or 0) if not pd.isna(row_data.get(f"Bagi Hasil {cat}", 0)) else 0.0
                    kual_data.append({
                        "Kategori": cat,
                        "Omzet": omzet_val,
                        "Bagi Hasil": bh_val
                    })
                
                df_kual_edit = pd.DataFrame(kual_data)
                edited_kual = st.data_editor(
                    df_kual_edit,
                    num_rows="fixed",
                    use_container_width=True,
                    key=f"editor_kual_{selected_teknisi}",
                    column_config={
                        "Omzet": st.column_config.NumberColumn("Omzet (Rp)", format="Rp %d"),
                        "Bagi Hasil": st.column_config.NumberColumn("Bagi Hasil (Rp)", format="Rp %d")
                    }
                )
                
                for _, e_row in edited_kual.iterrows():
                    cat = e_row["Kategori"]
                    df_parsed.at[idx, f"Omzet {cat}"] = e_row["Omzet"]
                    df_parsed.at[idx, f"Bagi Hasil {cat}"] = e_row["Bagi Hasil"]

            with t2:
                pot_data = []
                for p_col in detected_pots:
                    p_val = float(row_data.get(p_col, 0) or 0) if not pd.isna(row_data.get(p_col, 0)) else 0.0
                    pot_data.append({"Jenis Potongan": p_col, "Jumlah": p_val})
                
                df_pot_edit = pd.DataFrame(pot_data)
                edited_pot = st.data_editor(
                    df_pot_edit,
                    num_rows="fixed",
                    use_container_width=True,
                    key=f"editor_pot_{selected_teknisi}",
                    column_config={
                        "Jumlah": st.column_config.NumberColumn("Jumlah Potongan (Rp)", format="Rp %d")
                    }
                )

                for _, p_row in edited_pot.iterrows():
                    p_col = p_row["Jenis Potongan"]
                    df_parsed.at[idx, p_col] = p_row["Jumlah"]

            per_kual_new, bruto_new, pot_new, total_pot_new, nett_new, _, _, _, _ = extract_slip_data_from_row(df_parsed.loc[idx], df_parsed.columns)
            df_parsed.at[idx, 'BAGI_HASIL'] = bruto_new
            df_parsed.at[idx, 'TOTAL_POTONGAN'] = total_pot_new
            df_parsed.at[idx, 'NETT_BAGI_HASIL'] = nett_new

            st.success(f"Data {selected_teknisi} diperbarui secara otomatis di memori.")

        st.divider()

        # -------------------------------------------------------------------
        # Preview Semua Data Teknisi
        # -------------------------------------------------------------------
        st.subheader("Preview Ringkasan Semua Data Slip Gaji")
        preview_list = []
        for _, r in df_parsed.iterrows():
            _, bruto, _, total_pot, nett, _, _, _, _ = extract_slip_data_from_row(r, df_parsed.columns)
            preview_list.append({
                'Nama Teknisi': r.get('TEKNISI', '-'),
                'Cabang': r.get('CABANG', '-'),
                'Bruto Bagi Hasil': rupiah(bruto),
                'Total Potongan': rupiah(total_pot),
                'Nett Bagi Hasil': rupiah(nett)
            })
        st.dataframe(pd.DataFrame(preview_list), use_container_width=True, hide_index=True)
        
        st.divider()
        
        if st.button("🧾 Siapkan & Cetak Slip Gaji PDF", type="primary", use_container_width=True):
            with st.spinner("Menyusun berkas PDF slip gaji..."):
                zip_bytes, summary_df = generate_zip_slips(
                    df_parsed, periode_input, catatan_slip, zip_per_cabang=(bentuk == 'ZIP per cabang')
                )
                st.session_state['ready_zip'] = zip_bytes
                st.session_state['ready_summary'] = summary_df
                
        if st.session_state.get('ready_zip') is not None:
            st.subheader("Unduh Berkas Slip Gaji")
            st.download_button(
                label="⬇️ Unduh Semua Slip Gaji (.ZIP)",
                data=st.session_state['ready_zip'],
                file_name=f"slip_gaji_teknisi_{_nama_berkas_aman(periode_input)}.zip",
                mime="application/zip",
                use_container_width=True
            )
            
            summary_df = st.session_state.get('ready_summary')
            if summary_df is not None and not summary_df.empty:
                st.caption(f"Total **{int(summary_df['Jumlah Slip'].sum()):,} slip PDF** berhasil dibuat.")
                with st.expander("Rincian Jumlah Slip per Cabang"):
                    st.dataframe(summary_df, hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses file Excel: {e}")

else:
    st.info("Silakan upload file Excel yang berisi sheet **RAW** untuk memulai.")
