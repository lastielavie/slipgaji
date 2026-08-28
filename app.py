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
# Konfigurasi & Aset
# ---------------------------------------------------------------------------
DIR_ASET = Path(__file__).parent / "assets"
LOGO_MADINAH = DIR_ASET / "logo-madinah.png"
LOGO_MFLASH = DIR_ASET / "logo-mflash.png"

KATEGORI_ORDER = ['Interface', 'Normal', 'Mati Total', 'Promo', 'Lainnya']
KOLOM_POTONGAN = [
    'Potongan Refund', 'Potongan AR', 'Potongan Kasbon', 'Keterlambatan',
    'Potongan Minus Audit', 'Potongan Audit Compliance',
    'Biaya Pendaftaran Koperasi', 'Simpanan Pokok', 'Simpanan Wajib'
]
KOLOM_CADANGAN = ['Cadangan 7 Tahun / bulan', 'Cadangan 7 Tahun']

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
# Fungsi Pembantu & Generator PDF
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


@st.cache_data(show_spinner="Membaca berkas potongan...")
def baca_potongan(isi: bytes):
    """Membaca data potongan dari Excel hasil pengisian finance."""
    hasil, terbaca = {}, 0
    xls = pd.ExcelFile(io.BytesIO(isi), engine='openpyxl')
    for sheet in xls.sheet_names:
        try:
            d = xls.parse(sheet, header=3)
        except Exception:
            continue
        if 'Nama Teknisi' not in d.columns or 'Cabang' not in d.columns:
            continue
        if not any(k in d.columns for k in KOLOM_POTONGAN):
            continue
        d = d[d['Nama Teknisi'].notna() & (d['Nama Teknisi'].astype(str) != 'TOTAL')]
        for _, r in d.iterrows():
            kunci = (str(r['Cabang']).strip().upper(),
                     str(r['Nama Teknisi']).strip().upper())
            isi_baris = {}
            for kol in KOLOM_POTONGAN + KOLOM_CADANGAN:
                v = r.get(kol)
                isi_baris[kol] = 0.0 if v is None or pd.isna(v) else float(v)
            hasil[kunci] = isi_baris
            terbaca += 1
    return hasil, terbaca


def _baris_slip(sub, potongan):
    """Menghitung angka pendapatan & potongan untuk satu slip teknisi."""
    per_kual = []
    for lbl in KATEGORI_ORDER:
        s = sub[sub['TARIF_LABEL'] == lbl] if 'TARIF_LABEL' in sub.columns else pd.DataFrame()
        if s.empty:
            continue
        omzet = s['TOTAL HARGA'].sum() if 'TOTAL HARGA' in s.columns else 0.0
        bh = s['BAGI_HASIL'].sum() if 'BAGI_HASIL' in s.columns else 0.0
        if omzet > 0 or bh > 0:
            per_kual.append((lbl, omzet, bh / omzet if omzet else 0.0, bh))
    
    bruto = sum(x[3] for x in per_kual) if per_kual else sub['BAGI_HASIL'].sum() if 'BAGI_HASIL' in sub.columns else 0.0

    pot = []
    for label, kolom in PETA_POTONGAN_SLIP:
        pot.append((label, sum(float(potongan.get(k, 0) or 0) for k in kolom)))
    total_pot = sum(x[1] for x in pot)

    return per_kual, bruto, pot, total_pot, bruto - total_pot


def _gambar_slip(c, lebar, tinggi, nama, cabang, periode, angka, catatan):
    """Menggambar layout slip gaji pada Canvas ReportLab."""
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
        c.drawString(m + 2 * mm, y, '(tidak ada rincian transaksi per kualifikasi)')
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


def buat_pdf_teknisi(sub, nama, cabang, potongan, catatan, periode):
    """Menghasilkan stream bytes PDF untuk satu teknisi."""
    buf = io.BytesIO()
    lebar, tinggi = A4
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f'Slip Bagi Hasil — {nama} ({cabang})')
    c.setAuthor('Madinah Flash')
    pot = potongan.get((str(cabang).strip().upper(), str(nama).strip().upper()), {})
    _gambar_slip(c, lebar, tinggi, nama, cabang, periode, _baris_slip(sub, pot), catatan)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()


def buat_zip_slip(df_sumber, potongan, catatan, periode, zip_per_cabang=False):
    """Menghasilkan file ZIP berisi seluruh PDF slip gaji teknisi."""
    d = df_sumber.copy()
    d['CABANG'] = d['CABANG'].astype(str).str.strip()
    d['TEKNISI'] = d['TEKNISI'].astype(str).str.strip()
    
    buf = io.BytesIO()
    ringkas = []
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as luar:
        for cab in sorted(d['CABANG'].unique()):
            sub_cab = d[d['CABANG'] == cab]
            if sub_cab.empty:
                continue
            folder = _nama_berkas_aman(cab, 'CABANG')
            berkas = []
            for nama in sorted(sub_cab['TEKNISI'].unique()):
                if nama.upper() in ['TIDAK ADA TEKNISI', 'NAN', 'NONE', '']:
                    continue
                sub_tek = sub_cab[sub_cab['TEKNISI'] == nama]
                isi = buat_pdf_teknisi(sub_tek, nama, cab, potongan, catatan, periode)
                berkas.append((f'{folder} - {_nama_berkas_aman(nama)}.pdf', isi))
            
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
            ringkas.append({'Cabang': cab, 'Slip': len(berkas)})

    buf.seek(0)
    return buf.getvalue(), pd.DataFrame(ringkas)

# ---------------------------------------------------------------------------
# Antarmuka Aplikasi Streamlit
# ---------------------------------------------------------------------------
st.title("🧾 Aplikasi Cetak Slip Gaji Teknisi")
st.caption("Generator PDF Slip Gaji Teknisi Madinah Flash per Cabang")

if not REPORTLAB_AVAILABLE:
    st.error("Library `reportlab` belum terinstal. Silakan jalankan `pip install reportlab` untuk menggunakan fitur cetak PDF.")
    st.stop()

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("1. Unggah Data Data Penjualan / Bagi Hasil")
    up_sales = st.file_uploader(
        "Upload file data penjualan / rekap (CSV atau Excel)",
        type=['csv', 'xlsx', 'gz'],
        key='up_sales_slip'
    )

with col_right:
    st.subheader("2. Unggah Berkas Potongan (Opsional)")
    up_pot = st.file_uploader(
        "Upload file Excel potongan (diisi oleh Finance)",
        type=['xlsx'],
        key='up_pot_slip',
        help="Berkas Excel hasil ekspor rekap yang sudah diisi kolom potongannya. Jika kosong, seluruh potongan dianggap Rp 0."
    )

st.divider()

# Pengaturan Informasi Slip
st.subheader("3. Pengaturan Slip Gaji")
c_p1, c_p2 = st.columns(2)

with c_p1:
    periode_input = st.text_input("Label Periode Gaji", value="24 Juni 2026 – 23 Juli 2026")
    bentuk = st.radio(
        "Format Pengelompokan Berkas ZIP",
        ['Folder per cabang', 'ZIP per cabang'],
        horizontal=True,
        key='bentuk_zip_app'
    )

with c_p2:
    catatan_slip = st.text_area(
        "Catatan pada Slip",
        value="Slip gaji ini dikeluarkan otomatis oleh sistem dan sah tanpa tanda tangan basah.",
        height=100
    )

# Proses Data Penjualan
df_jasa = pd.DataFrame()
if up_sales is not None:
    try:
        if up_sales.name.endswith('.csv') or up_sales.name.endswith('.csv.gz'):
            df_jasa = pd.read_csv(up_sales)
        else:
            df_jasa = pd.read_excel(up_sales)

        # Normalisasi Kolom Wajib
        col_map = {}
        for c in df_jasa.columns:
            cu = str(c).upper().strip()
            if cu in ['TEKNISI', 'NAMA TEKNISI', 'NAMA TEKNISI (FINAL)']:
                col_map[c] = 'TEKNISI'
            elif cu in ['CABANG']:
                col_map[c] = 'CABANG'
            elif cu in ['TOTAL HARGA', 'OMZET', 'OMZET JASA']:
                col_map[c] = 'TOTAL HARGA'
            elif cu in ['BAGI HASIL', 'BAGI_HASIL', 'BAGI HASIL (ATURAN)']:
                col_map[c] = 'BAGI_HASIL'
            elif cu in ['TARIF_LABEL', 'KATEGORI TARIF', 'KUALIFIKASI']:
                col_map[c] = 'TARIF_LABEL'
        
        df_jasa = df_jasa.rename(columns=col_map)
        
        # Validasi minimal kolom TEKNISI & CABANG
        if 'TEKNISI' not in df_jasa.columns or 'CABANG' not in df_jasa.columns:
            st.error("File data harus memiliki minimal kolom `TEKNISI` dan `CABANG`.")
            df_jasa = pd.DataFrame()
        else:
            if 'BAGI_HASIL' not in df_jasa.columns:
                df_jasa['BAGI_HASIL'] = df_jasa['TOTAL HARGA'] if 'TOTAL HARGA' in df_jasa.columns else 0.0
            st.success(f"Data terbaca: {len(df_jasa):,} baris transaksi.")
    except Exception as e:
        st.error(f"Gagal membaca file data: {e}")

# Baca Potongan
potongan_map = {}
if up_pot is not None:
    try:
        potongan_map, n_pot = baca_potongan(up_pot.getvalue())
        st.success(f"Potongan terbaca untuk {n_pot:,} baris teknisi.")
    except Exception as e:
        st.error(f"Berkas potongan tidak dapat diproses: {e}")

st.divider()

# Tombol Eksekusi
if not df_jasa.empty:
    if st.button("🧾 Siapkan Slip Gaji PDF", use_container_width=True, type="primary"):
        with st.spinner("Menyusun berkas PDF slip gaji..."):
            try:
                zip_bytes, ringkas_df = buat_zip_slip(
                    df_jasa, potongan_map, catatan_slip, periode_input,
                    zip_per_cabang=(bentuk == 'ZIP per cabang')
                )
                st.session_state['result_zip'] = zip_bytes
                st.session_state['result_summary'] = ringkas_df
            except Exception as e:
                st.error(f"Terjadi kesalahan saat membuat PDF: {e}")

if st.session_state.get('result_zip') is not None:
    st.subheader("4. Unduh Berkas Slip Gaji")
    st.download_button(
        label="⬇️ Unduh Semua Slip Gaji (.ZIP)",
        data=st.session_state['result_zip'],
        file_name=f"slip_gaji_teknisi_{periode_input.replace(' ', '_')}.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    rs = st.session_state.get('result_summary')
    if rs is not None and not rs.empty:
        st.caption(f"Total **{int(rs['Slip'].sum()):,} slip** berhasil dibuat untuk **{len(rs)} cabang**.")
        with st.expander("Rincian Jumlah Slip per Cabang"):
            st.dataframe(rs, hide_index=True, use_container_width=True)