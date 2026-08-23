"""
st_dinas_luar_builder.py — Mesin Pembuat Surat Tugas Dinas Luar
Dekan: Dr. Ir. Ockstan Jurike Kalesaran, M.Sc. | NIP. 196910241994032014
Semua font minimal 12 pt. Kop menggunakan header Word dengan logo di kiri.
"""
import io
import os
from dataclasses import dataclass
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_OFFICIAL = "Times New Roman"
DEKAN_NAMA    = "Dr. Ir. Ockstan Jurike Kalesaran, M.Sc."
DEKAN_NIP     = "196910241994032014"

_HERE     = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(_HERE, "logo_unsrat.png")
if not os.path.exists(LOGO_PATH):
    LOGO_PATH = "logo_unsrat.png"


@dataclass
class DataSuratTugas:
    nomor_surat: str  = "800/.../FPIK-UNSRAT/2026"
    nama: str         = ""
    nip: str          = ""
    jabatan: str      = ""
    unit_kerja: str   = "Fakultas Perikanan dan Ilmu Kelautan, Universitas Sam Ratulangi"
    jenis_tugas: str  = "Survei Awal Ekologi dan Pemetaan Dasar"
    lokasi: str       = "Pelabuhan Perikanan Samudera (PPS) Bitung, Sulawesi Utara"
    tanggal_mulai:    date = date(2026, 9, 15)
    tanggal_selesai:  date = date(2026, 9, 16)
    jumlah_hari: int  = 2
    klaim_8_oj: bool  = True
    peran_tim: str    = "Pembantu Peneliti"


def format_tanggal_indonesia(tgl: date) -> str:
    bulan_id = [
        "Januari","Februari","Maret","April","Mei","Juni",
        "Juli","Agustus","September","Oktober","November","Desember"
    ]
    return f"{tgl.day} {bulan_id[tgl.month - 1]} {tgl.year}"


def _set_tcW(cell, twips: int):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(twips))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def _no_border(cell):
    tcPr  = cell._tc.get_or_add_tcPr()
    tcBdr = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right","insideH","insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "none")
        el.set(qn("w:sz"),    "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        tcBdr.append(el)
    tcPr.append(tcBdr)


def _add_kop_surat(doc: Document):
    """
    Kop menggunakan header Word (bukan body) agar teks bebas menggunakan
    lebar penuh halaman. Layout:
      Tabel 1 baris x 2 kolom di dalam header:
        - Kol kiri  : logo 1.8 cm
        - Kol kanan : teks kementerian/universitas/fakultas/alamat, center
    Garis tebal hitam ditambahkan di bawah paragraph terakhir header.
    """
    section  = doc.sections[0]
    header   = section.header
    header.is_linked_to_previous = False

    # Pastikan header tidak punya paragraf default kosong berlebih
    # Kosongkan dulu isi header
    for p in header.paragraphs:
        p.clear()

    # Hitung lebar usable header dalam twips
    # Lebar halaman A4 = 11906 twips (21 cm)
    # margin kiri 2.5cm = 1418 twips, kanan 2.0cm = 1134 twips
    # usable = 11906 - 1418 - 1134 = 9354 twips
    TW_LOGO = 1020   # ~1.8 cm
    TW_TEKS = 8334   # ~14.7 cm

    tbl = header.add_table(rows=1, cols=2, width=Cm(16.5))
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    cl = tbl.cell(0, 0)
    ct = tbl.cell(0, 1)
    _set_tcW(cl, TW_LOGO)
    _set_tcW(ct, TW_TEKS)
    _no_border(cl)
    _no_border(ct)

    # --- Logo ---
    p_logo = cl.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after  = Pt(0)
    if os.path.exists(LOGO_PATH):
        p_logo.add_run().add_picture(LOGO_PATH, width=Cm(1.7))
    else:
        r = p_logo.add_run("[LOGO]")
        r.font.size = Pt(10)

    # --- Teks kop ---
    baris = [
        ("KEMENTERIAN PENDIDIKAN TINGGI, SAINS, DAN TEKNOLOGI", 11, True),
        ("UNIVERSITAS SAM RATULANGI",                           12, True),
        ("FAKULTAS PERIKANAN DAN ILMU KELAUTAN",               11, True),
        (
            "Alamat : Kampus UNSRAT Manado 95115  "
            "Laman : http://fpik.unsrat.ac.id ; Email : fpik@unsrat.ac.id",
            10, False
        ),
    ]

    first = True
    for (teks, ukuran, tebal) in baris:
        if first:
            p = ct.paragraphs[0]
            first = False
        else:
            p = ct.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        r = p.add_run(teks)
        r.font.name = FONT_OFFICIAL
        r.font.size = Pt(ukuran)
        r.bold      = tebal

    # --- Garis tebal bawah di header ---
    p_line = header.add_paragraph()
    p_line.paragraph_format.space_before = Pt(3)
    p_line.paragraph_format.space_after  = Pt(0)
    pPr  = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    btm  = OxmlElement("w:bottom")
    btm.set(qn("w:val"),   "single")
    btm.set(qn("w:sz"),    "18")
    btm.set(qn("w:space"), "1")
    btm.set(qn("w:color"), "000000")
    pBdr.append(btm)
    pPr.append(pBdr)


def generate_st_dinas_luar(data: DataSuratTugas) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.top_margin        = Cm(2.0)
        section.bottom_margin     = Cm(2.5)
        section.left_margin       = Cm(2.5)
        section.right_margin      = Cm(2.0)
        section.header_distance   = Cm(1.0)

    # 1. KOP di header Word
    _add_kop_surat(doc)

    # Jarak antara header dan konten
    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(6)
    p_gap.paragraph_format.space_after  = Pt(0)

    # 2. JUDUL
    p_judul = doc.add_paragraph()
    p_judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_judul.paragraph_format.space_before = Pt(6)
    p_judul.paragraph_format.space_after  = Pt(2)
    r_j = p_judul.add_run("SURAT TUGAS DINAS LUAR")
    r_j.font.name = FONT_OFFICIAL; r_j.font.size = Pt(12); r_j.bold = True; r_j.underline = True

    p_nomor = doc.add_paragraph()
    p_nomor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nomor.paragraph_format.space_after = Pt(12)
    r_n = p_nomor.add_run(f"Nomor: {data.nomor_surat}")
    r_n.font.name = FONT_OFFICIAL; r_n.font.size = Pt(12)

    # 3. PEMBUKA
    p_isi = doc.add_paragraph()
    p_isi.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_isi.paragraph_format.line_spacing = 1.5
    p_isi.paragraph_format.space_after  = Pt(6)
    r_i = p_isi.add_run(
        "Dekan Fakultas Perikanan dan Ilmu Kelautan Universitas Sam Ratulangi, "
        "menugaskan kepada:"
    )
    r_i.font.name = FONT_OFFICIAL; r_i.font.size = Pt(12)

    # 4. TABEL PERSONIL
    # usable body = 21 - 2.5 - 2.0 = 16.5 cm = 9355 twips
    TW_L = 2835   # ~5.0 cm
    TW_R = 6520   # ~11.5 cm
    tbl_p = doc.add_table(rows=5, cols=2)
    tbl_p.style = "Table Grid"
    for i, (label, val) in enumerate([
        ("Nama",               data.nama),
        ("NIP",                data.nip),
        ("Jabatan",            data.jabatan),
        ("Unit Kerja",         data.unit_kerja),
        ("Peran dalam Kegiatan", f"{data.peran_tim} \u2014 {data.jenis_tugas}"),
    ]):
        cl = tbl_p.cell(i, 0); cr = tbl_p.cell(i, 1)
        _set_tcW(cl, TW_L); _set_tcW(cr, TW_R)
        rl = cl.paragraphs[0].add_run(label)
        rl.bold = True; rl.font.name = FONT_OFFICIAL; rl.font.size = Pt(12)
        rr = cr.paragraphs[0].add_run(f": {val}")
        rr.font.name = FONT_OFFICIAL; rr.font.size = Pt(12)

    # 5. RINCIAN TUGAS
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_t.paragraph_format.line_spacing = 1.5
    p_t.paragraph_format.space_after  = Pt(4)
    r_t = p_t.add_run(
        "Untuk melaksanakan tugas dinas luar dalam rangka kegiatan Swakelola Tipe II: "
        "Relokasi dan Restorasi Terumbu Karang di Pelabuhan Perikanan Samudera (PPS) "
        "Bitung Tahun Anggaran 2026, dengan rincian sebagai berikut:"
    )
    r_t.font.name = FONT_OFFICIAL; r_t.font.size = Pt(12)

    tgl_str = (
        f"{format_tanggal_indonesia(data.tanggal_mulai)} "
        f"s.d. {format_tanggal_indonesia(data.tanggal_selesai)} "
        f"({data.jumlah_hari} hari kerja)"
    )
    tbl_r = doc.add_table(rows=3, cols=2)
    for i, (k, v) in enumerate([
        ("Jenis Tugas",  data.jenis_tugas),
        ("Lokasi Tugas", data.lokasi),
        ("Tanggal",      tgl_str),
    ]):
        cl = tbl_r.cell(i, 0); cr = tbl_r.cell(i, 1)
        _set_tcW(cl, TW_L); _set_tcW(cr, TW_R)
        _no_border(cl); _no_border(cr)
        rk = cl.paragraphs[0].add_run(k)
        rk.bold = True; rk.font.name = FONT_OFFICIAL; rk.font.size = Pt(12)
        rv = cr.paragraphs[0].add_run(f": {v}")
        rv.font.name = FONT_OFFICIAL; rv.font.size = Pt(12)

    # 6. KLAUSUL
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_k = doc.add_paragraph()
    p_k.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_k.paragraph_format.line_spacing = 1.5
    if data.klaim_8_oj and data.peran_tim == "Pembantu Peneliti":
        kteks = (
            "Surat tugas ini diterbitkan sebagai syarat mutlak pengajuan honorarium "
            "Pembantu Peneliti sebesar 8 (delapan) Orang-Jam (OJ) per hari kerja. "
            "Sebagai konsekuensinya, yang bersangkutan dibebaskan dari tugas mengajar "
            "dan tugas rutin kantor pada tanggal tersebut, serta tidak menerima uang "
            "lauk pauk/tunjangan makan dari unit kerja."
        )
    else:
        kteks = (
            "Surat tugas ini diterbitkan untuk melegalekan penugasan dinas luar dan "
            "menjadi dasar pengajuan Uang Harian Perjalanan Dinas (UHPD) sesuai dengan "
            "Standar Biaya Masukan (SBM) yang berlaku."
        )
    r_k = p_k.add_run(kteks)
    r_k.font.name = FONT_OFFICIAL; r_k.font.size = Pt(12); r_k.italic = True

    # 7. TTD
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    p_ttd = doc.add_paragraph()
    p_ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ttd.paragraph_format.line_spacing = 1.5

    def _rr(paragraph, text, bold=False, nl=True):
        r = paragraph.add_run(text + ("\n" if nl else ""))
        r.font.name = FONT_OFFICIAL; r.font.size = Pt(12); r.bold = bold

    _rr(p_ttd, f"Manado, {format_tanggal_indonesia(date.today())}")
    _rr(p_ttd, "Dekan,")
    for _ in range(4): _rr(p_ttd, "")
    _rr(p_ttd, DEKAN_NAMA, bold=True)
    _rr(p_ttd, f"NIP. {DEKAN_NIP}", nl=False)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()
