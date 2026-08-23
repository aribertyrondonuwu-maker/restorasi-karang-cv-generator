"""
st_dinas_luar_builder.py — Mesin Pembuat Surat Tugas Dinas Luar
Dekan: Dr. Ir. Ockstan Jurike Kalesaran, M.Sc. | NIP. 196910241994032014
Semua font minimal 12 pt.
"""
import io
from dataclasses import dataclass
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_OFFICIAL = "Times New Roman"
COLOR_HEADER = RGBColor(0x00, 0x33, 0x66)

DEKAN_NAMA = "Dr. Ir. Ockstan Jurike Kalesaran, M.Sc."
DEKAN_NIP = "196910241994032014"


@dataclass
class DataSuratTugas:
    nomor_surat: str = "800/.../FPIK-UNSRAT/2026"
    nama: str = ""
    nip: str = ""
    jabatan: str = ""
    unit_kerja: str = "Fakultas Perikanan dan Ilmu Kelautan, Universitas Sam Ratulangi"
    jenis_tugas: str = "Survei Awal Ekologi dan Pemetaan Dasar"
    lokasi: str = "Pelabuhan Perikanan Samudera (PPS) Bitung, Sulawesi Utara"
    tanggal_mulai: date = date(2026, 9, 15)
    tanggal_selesai: date = date(2026, 9, 16)
    jumlah_hari: int = 2
    klaim_8_oj: bool = True
    peran_tim: str = "Pembantu Peneliti"


def format_tanggal_indonesia(tgl: date) -> str:
    bulan_id = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    return f"{tgl.day} {bulan_id[tgl.month - 1]} {tgl.year}"


def add_horizontal_line(doc):
    p = doc.add_paragraph()
    p.space_before = Pt(2)
    p.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def generate_st_dinas_luar(data: DataSuratTugas) -> bytes:
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
    
    # 1. KOP SURAT
    p_kop1 = doc.add_paragraph()
    p_kop1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop1.space_after = Pt(0)
    
    r1 = p_kop1.add_run("KEMENTERIAN PENDIDIKAN TINGGI, SAINS, DAN TEKNOLOGI\n")
    r1.font.name = FONT_OFFICIAL; r1.font.size = Pt(12); r1.bold = True
    
    r2 = p_kop1.add_run("UNIVERSITAS SAM RATULANGI\n")
    r2.font.name = FONT_OFFICIAL; r2.font.size = Pt(12); r2.bold = True
    
    r3 = p_kop1.add_run("FAKULTAS PERIKANAN DAN ILMU KELAUTAN\n")
    r3.font.name = FONT_OFFICIAL; r3.font.size = Pt(12); r3.bold = True
    
    r4 = p_kop1.add_run("Kampus Unsrat Bahu, Kecamatan Malalayang, Kota Manado 95115")
    r4.font.name = FONT_OFFICIAL; r4.font.size = Pt(12)
    
    add_horizontal_line(doc)
    
    # 2. JUDUL
    p_judul = doc.add_paragraph()
    p_judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_judul.space_before = Pt(12)
    p_judul.space_after = Pt(2)
    r_judul = p_judul.add_run("SURAT TUGAS DINAS LUAR")
    r_judul.font.name = FONT_OFFICIAL; r_judul.font.size = Pt(12); r_judul.bold = True; r_judul.underline = True
    
    p_nomor = doc.add_paragraph()
    p_nomor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nomor.space_after = Pt(12)
    r_nomor = p_nomor.add_run(f"Nomor: {data.nomor_surat}")
    r_nomor.font.name = FONT_OFFICIAL; r_nomor.font.size = Pt(12)
    
    # 3. ISI
    p_isi = doc.add_paragraph()
    p_isi.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_isi.paragraph_format.line_spacing = 1.15
    p_isi.paragraph_format.space_after = Pt(8)
    
    r_isi = p_isi.add_run("Dekan Fakultas Perikanan dan Ilmu Kelautan Universitas Sam Ratulangi, menugaskan kepada:\n")
    r_isi.font.name = FONT_OFFICIAL; r_isi.font.size = Pt(12)
    
    tabel = doc.add_table(rows=5, cols=2)
    tabel.style = "Table Grid"
    data_personil = [
        ("Nama", data.nama), ("NIP", data.nip), ("Jabatan", data.jabatan),
        ("Unit Kerja", data.unit_kerja), ("Peran dalam Kegiatan", f"{data.peran_tim} — {data.jenis_tugas}"),
    ]
    
    for i, (label, val) in enumerate(data_personil):
        cell_l = tabel.cell(i, 0); cell_r = tabel.cell(i, 1)
        cell_l.width = Inches(1.5); cell_r.width = Inches(4.5)
        
        r_l = cell_l.paragraphs[0].add_run(label)
        r_l.bold = True; r_l.font.name = FONT_OFFICIAL; r_l.font.size = Pt(12)
        
        r_r = cell_r.paragraphs[0].add_run(f": {val}")
        r_r.font.name = FONT_OFFICIAL; r_r.font.size = Pt(12)
    
    doc.add_paragraph().space_after = Pt(6)
    
    p_tugas = doc.add_paragraph()
    p_tugas.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_tugas.paragraph_format.line_spacing = 1.15
    p_tugas.paragraph_format.space_after = Pt(8)
    
    r_tugas = p_tugas.add_run(
        f"Untuk melaksanakan tugas dinas luar dalam rangka kegiatan "
        f"Swakelola Tipe II: Relokasi dan Restorasi Terumbu Karang di "
        f"Pelabuhan Perikanan Samudera (PPS) Bitung Tahun Anggaran 2026, "
        f"dengan rincian sebagai berikut:\n\n"
        f"Jenis Tugas\t: {data.jenis_tugas}\n"
        f"Lokasi Tugas\t: {data.lokasi}\n"
        f"Tanggal\t\t: {format_tanggal_indonesia(data.tanggal_mulai)} "
        f"s.d. {format_tanggal_indonesia(data.tanggal_selesai)} "
        f"({data.jumlah_hari} hari kerja)\n"
    )
    r_tugas.font.name = FONT_OFFICIAL; r_tugas.font.size = Pt(12)
    
    # 4. KLAUSUL
    p_klausul = doc.add_paragraph()
    p_klausul.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_klausul.paragraph_format.line_spacing = 1.15
    p_klausul.paragraph_format.space_before = Pt(6)
    
    if data.klaim_8_oj and data.peran_tim == "Pembantu Peneliti":
        klausul_teks = (
            "SURAT TUGAS INI DITERBITKAN SEBAGAI SYARAT MUTLAK PENGAJUAN "
            "HONORARIUM PEMBANTU PENELITI SEBESAR 8 (DELAPAN) ORANG-JAM (OJ) "
            "PER HARI KERJA. SEBAGAI KONSEKUENSINYA, YANG BERSANGKUTAN "
            "DIBEBASKAN DARI TUGAS MENGAJAR DAN TUGAS RUTIN KANTOR PADA "
            "TANGGAL TERSEBUT, SERTA TIDAK MENERIMA UANG LAUK PAUK/TUNJANGAN "
            "MAKAN DARI UNIT KERJA."
        )
    else:
        klausul_teks = (
            "SURAT TUGAS INI DITERBITKAN UNTUK MELEGALEKAN PENUGASAN DINAS LUAR "
            "DAN MENJADI DASAR PENGAJUAN UANG HARIAN PERJALANAN DINAS (UHPD) "
            "SESUAI DENGAN STANDAR BIAYA MASUKAN (SBM) YANG BERLAKU."
        )
    
    r_klausul = p_klausul.add_run(klausul_teks)
    r_klausul.font.name = FONT_OFFICIAL; r_klausul.font.size = Pt(12); r_klausul.bold = True; r_klausul.italic = True
    
    # 5. TTD
    doc.add_paragraph().space_after = Pt(6)
    
    p_ttd = doc.add_paragraph()
    p_ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ttd.paragraph_format.line_spacing = 1.15
    
    r_ttd1 = p_ttd.add_run("Manado, "); r_ttd1.font.name = FONT_OFFICIAL; r_ttd1.font.size = Pt(12)
    r_ttd2 = p_ttd.add_run(format_tanggal_indonesia(date.today()) + "\n"); r_ttd2.font.name = FONT_OFFICIAL; r_ttd2.font.size = Pt(12)
    r_ttd3 = p_ttd.add_run("DEKAN FPIK UNSRAT,\n\n\n\n\n"); r_ttd3.font.name = FONT_OFFICIAL; r_ttd3.font.size = Pt(12)
    
    r_dekan = p_ttd.add_run(f"{DEKAN_NAMA}\n")
    r_dekan.font.name = FONT_OFFICIAL; r_dekan.font.size = Pt(12); r_dekan.bold = True
    
    r_nip = p_ttd.add_run(f"NIP. {DEKAN_NIP}")
    r_nip.font.name = FONT_OFFICIAL; r_nip.font.size = Pt(12)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()