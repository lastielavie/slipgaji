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

st.set_page_config(page_title="Cetak Slip Gaji Teknisi", layout="wide", page_icon="🧾")

# ---------------------------------------------------------------------------
# Konfigurasi & Aset Logo (jika ada folder assets)
# ---------------------------------------------------------------------------
DIR_ASET = Path(__file__).parent / "assets"
LOGO_MADINAH = DIR_ASET / "logo-madinah.png"
LOGO_MFLASH = DIR_ASET / "logo-mflash.png"

KATEGORI_ORDER = ['Interface', 'Normal', 'Mati Total', 'Promo', 'Lainnya']

PETA_POTONGAN_SLIP = [
    ('Potongan Kasbon', ['Potongan Kasbon']),
    ('Potongan Refund', ['Potongan Refund']),
    ('Potongan AR', ['Potongan AR']),
    ('Potongan Terlambat', ['Keterlambatan']),
    ('Potongan Minus Audit', ['Potongan Minus Audit']),
    ('Potongan Audit Compliance', ['Potongan Audit Compliance']),
    ('Potongan Koperasi', ['Biaya Pendaftaran Koperasi', 'Simpanan Pokok', 'Simpanan Wajib']),
]

# ---------------------------------------------------------------------------
# Fungsi Pembantu & Parser Excel Sheet RAW
# ---------------------------------------------------------------------------
def rupiah(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        v = 0.0
    s = f"{abs(v):,.0f}".replace(",", ".")
    return ("Rp (" + s + ")") if v < 0 else ("Rp " + s)


def _nama_berkas_aman(teks, cadangan='TANPA-NAMA'):
    aman = re.sub(r'[^A-Za-z0-9 _.-]', '-', str(teks)).strip(' .-')
    return (aman[:80] or cadangan)


def parse_excel_file(file_bytes, file_name):
    """Membaca file Excel dan mendeteksi sheet RAW serta periode secara otomatis."""
    periode_str = ""
    sheet_used = ""

    if file_name.lower().endswith(('.csv', '.csv.gz')):
        df_raw = pd.read_csv(io.BytesIO(file_bytes))
        sheet_used = "CSV"
    else:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        
        # 1. Cari sheet 'RAW'
        target_sheet = None
        if 'RAW' in xls.sheet_names:
            target_sheet = 'RAW'
        else:
            for s in xls.sheet_names:
                if 'RAW' in s.upper():
                    target_sheet = s
                    break
            if not target_sheet:
                target_sheet = xls.sheet_names[0]
                
        sheet_used = target_sheet
        
        # 2. Cek posisi header & ekstrak teks Periode dari baris atas
        df_temp = pd.read_excel(xls, sheet_name=target_sheet, header=None)
        header_row_idx = 0
        
        for i in range(min(10, len(df_temp))):
            row_str_vals = [str(v).strip() for v in df_temp.iloc[i].values]
            row_text = " ".join(row_str_vals)
            
            # Deteksi Teks Periode (misal: "Periode: 24 Juli 2026 – 23 Agustus 2026")
            if "Periode:" in row_text and not periode_str:
                parts = row_text.split("Periode:")
                if len(parts) > 1:
                    periode_str = parts[1].split("·")[0].strip()
            
            # Deteksi Baris Header 'Nama Teknisi'
            if any('NAMA TEKNISI' in str(v).upper() for v in row_str_vals):
                header_row_idx = i
                break
                
        df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=header_row_idx)

    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    # Standarisasi nama kolom penting
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
        elif cu in ['GAJI TEKNISI']:
            col_map[c] = 'GAJI_TEKNISI'
            
    df_data = df_raw.rename(columns=col_map)
    
    # Filter hanya baris teknisi yang valid
    if 'TEKNISI' in df_data.columns:
        df_data = df_data[df_data['TEKNISI'].notna()]
        df_data = df_data[~df_data['TEKNISI'].astype(str).str.strip().str.upper().isin(
            ['TOTAL', 'NAN', 'NONE', '', 'UNNAMED: 0']
        )]
        
    return df_data, periode_str, sheet_used


def extract_slip_data_from_row(row):
    """Mengambil pendapatan per kualifikasi & potongan langsung dari baris sheet RAW."""
    per_kual = []
    for k in KATEGORI_ORDER:
        omzet_col = f"Omzet {k}"
        bh_col = f"Bagi Hasil {k}"
        
        omzet = float(row.get(omzet_col, 0) or 0)
        bh = float(row.get(bh_col, 0) or 0)
        
        if omzet > 0 or bh > 0:
            akad_pct = bh / omzet if omzet > 0 else 0.0
            per_kual.append((k, omzet, akad_pct, bh))
            
    bruto = float(row.get('BAGI_HASIL', 0) or 0)
    if not bruto and per_kual:
        bruto = sum(x[3] for x in per_kual)
        
    pot = []
    for label, cols in PETA_POTONGAN_SLIP:
        val = sum(float(row.get(c, 0) or 0) for c in cols)
        pot.append((label, val))
        
    total_pot = float(row.get('TOTAL_POTONGAN', 0) or 0)
    if not total_pot and pot:
        total_pot = sum(x[1] for x in pot)
        
    nett = float(row.get('NETT_BAGI_HASIL', 0) or (bruto - total_pot))
    
    return per_kual, bruto, pot, total_pot, nett

# ---------------------------------------------------------------------------
# Generator PDF Slip Gaji
# ---------------------------------------------------------------------------
def _gambar_slip(c, lebar, tinggi, nama, cabang, periode, angka, catatan):
    """Menggambar desain PDF Slip Gaji."""
    per_kual, bruto, pot, total_pot, nett = angka
    m = 18 * mm
    y = tinggi - 14 * mm

    if LOGO_MADINAH.exists():
        c.drawImage(ImageReader(str(LOGO_MADINAH)), m, y - 20 * mm, width=20 * mm,
                    height=20 * mm, mask='auto')
    if LOGO_MFLASH.exists():
        c.drawImage(ImageReader(str(LOGO_MFLASH)), lebar - m - 34 * mm, y - 20 * mm,
                    width=34 * mm, height=24 * mm, mask='auto',
                    preserveAspectRatio=True, anchor='ne')
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
    for label, isi in [('Nama', nama), ('Jabatan', 'Teknisi'),
                       ('Divisi', f'MFlash — {cabang}'), ('Periode', periode)]:
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
    c.setFont('Helvetica', 9)
    if not per_kual:
        c.drawString(m + 2 * mm, y, '(rincian kualifikasi tidak ditemukan)')
        y -= 5.4 * mm
    for lbl, omzet, akad, bh in per_kual:
        c.drawString(m + 2 * mm, y, lbl)
        c.drawRightString(lebar - m - 46 * mm, y, rupiah(omzet))
        c.drawRightString(lebar - m - 30 * mm, y, f'{akad*100:.1f}'.replace('.', ',') + '%')
        c.drawRightString(lebar - m - 2 * mm, y, rupiah(bh))
        y -= 5.4 * mm

    y -= 1 * mm
    c.setLineWidth(0.6)
    c.line(lebar - m - 52 * mm, y + 1.5 * mm, lebar - m, y + 1.5 * mm)
    y -= 3 * mm
    c.setFont('Helvetica-Bold', 9.5)
    c.drawString(m + 2 * mm, y, 'Total Bruto Bagi Hasil')
    c.drawRightString(lebar - m - 2 * mm, y, rupiah(bruto))
    y -= 9 * mm

    judul_tabel('POTONGAN', kolom_kanan=False)
    c.setFont('Helvetica', 9)
    for label, nilai in pot:
        c.drawString(m + 2 * mm, y, label)
        c.drawRightString(lebar - m - 2 * mm, y, rupiah(nilai))
        y -= 5.4 * mm
    c.setLineWidth(0.6)
    c.line(lebar - m - 52 * mm, y + 1.5 * mm, lebar - m, y + 1.5 * mm)
    y -= 3 * mm
    c.setFont('Helvetica-Bold', 9.5)
    c.drawString(m + 2 * mm, y, 'Total Potongan')
    c.drawRightString(lebar - m - 2 * mm, y, rupiah(total_pot))
    y -= 9 * mm

    c.setFillColorRGB(0.86, 0.92, 0.84)
    c.rect(m, y - 3 * mm, lebar - 2 * m, 8 * mm, stroke=0, fill=1)
    c.setFillColorRGB(0.05, 0.35, 0.15)
    c.setFont('Helvetica-Bold', 11)
    c.drawString(m + 2 * mm, y, 'NETT BAGI HASIL')
    c.drawRightString(lebar - m - 2 * mm, y, rupiah(nett))
    c.setFillColorRGB(0, 0, 0)
    y -= 14 * mm

    c.setFont('Helvetica-Bold', 8.5)
    c.drawString(m, y, 'Catatan')
    y -= 3 * mm
    tinggi_kotak = 20 * mm
    c.setLineWidth(0.6)
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.rect(m, y - tinggi_kotak, lebar - 2 * m, tinggi_kotak, stroke=1, fill=0)
    c.setFont('Helvetica', 8.5)
    baris_catatan = str(catatan or '').splitlines()
    yy = y - 5 * mm
    for baris in baris_catatan[:5]:
        c.drawString(m + 2 * mm, yy, baris[:110])
        yy -= 4.2 * mm
    y -= tinggi_kotak + 12 * mm

    c.setStrokeColorRGB(0, 0, 0)
    c.setFont('Helvetica', 8.5)
    for x, teks in ((m + 8 * mm, 'Teknisi'),
                    (lebar / 2 - 12 * mm, 'Kepala Cabang'),
                    (lebar - m - 40 * mm, 'Finance')):
        c.line(x, y, x + 32 * mm, y)
        c.drawCentredString(x + 16 * mm, y - 4.5 * mm, teks)


def generate_zip_slips(df_data, periode_txt, catatan_slip, zip_per_cabang=False):
    """Membuat file ZIP berisi seluruh file PDF Slip Gaji."""
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
                    
                angka = extract_slip_data_from_row(row)
                
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
# Antarmuka Aplikasi Streamlit
# ---------------------------------------------------------------------------
st.title("🧾 Aplikasi Cetak Slip Gaji Teknisi")
st.caption("Generator PDF Slip Gaji Teknisi Madinah Flash dari Sheet RAW Excel")

if not REPORTLAB_AVAILABLE:
    st.error("Library `reportlab` belum terinstal. Pastikan file `requirements.txt` berisi `reportlab` dan sudah di-commit ke GitHub.")
    st.stop()

# 1. Upload File Excel tunggal
uploaded_file = st.file_uploader(
    "Upload 1 File Excel (misal: '05 BAGI HASIL TEKNISI CILANGKAP AGUSTUS 2026.xlsx')",
    type=['xlsx', 'xls', 'csv', 'gz'],
    key='main_uploader'
)

if uploaded_file is not None:
    try:
        df_parsed, auto_periode, sheet_used = parse_excel_file(
            uploaded_file.getvalue(), uploaded_file.name
        )
        
        st.success(f"Berhasil membaca sheet **'{sheet_used}'**! Ditemukan **{len(df_parsed)} data teknisi**.")
        
        st.divider()
        
        # 2. Pengaturan Slip Gaji
        col1, col2 = st.columns(2)
        
        with col1:
            periode_input = st.text_input(
                "Label Periode Gaji",
                value=auto_periode if auto_periode else "24 Juli 2026 – 23 Agustus 2026",
                help="Otomatis terdeteksi dari isi sheet RAW atau bisa diubah manual."
            )
            bentuk = st.radio(
                "Format Pengelompokan File ZIP",
                ['Folder per cabang', 'ZIP per cabang'],
                horizontal=True
            )
            
        with col2:
            catatan_slip = st.text_area(
                "Catatan pada Slip",
                value="Slip gaji ini dikeluarkan otomatis oleh sistem dan sah tanpa tanda tangan basah.",
                height=100
            )

        # 3. Preview Ringkasan Data
        st.subheader("Preview Data Slip Gaji")
        preview_list = []
        for _, r in df_parsed.iterrows():
            _, bruto, _, total_pot, nett = extract_slip_data_from_row(r)
            preview_list.append({
                'Nama Teknisi': r.get('TEKNISI', '-'),
                'Cabang': r.get('CABANG', '-'),
                'Bruto Bagi Hasil': rupiah(bruto),
                'Total Potongan': rupiah(total_pot),
                'Nett Bagi Hasil': rupiah(nett)
            })
        st.dataframe(pd.DataFrame(preview_list), use_container_width=True, hide_index=True)
        
        st.divider()
        
        # 4. Tombol Generate PDF
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