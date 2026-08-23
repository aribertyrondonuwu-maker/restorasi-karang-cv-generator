"""
cv_builder.py — Mesin Pembangun CV Modern (DOCX)
Template: Tema Kelautan, Layout Profesional, Foto 3x4
Semua font minimal 12 pt.
"""
import io
import base64
from dataclasses import dataclass
from typing import Optional, List
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===== PALET WARNA TEMA KELAUTAN =====
COLOR_PRIMARY   = RGBColor(0x00, 0x3D, 0x66)
COLOR_SECONDARY = RGBColor(0x00, 0x77, 0xB6)
COLOR_ACCENT    = RGBColor(0x00, 0xB4, 0xD8)
COLOR_TEXT      = RGBColor(0x1A, 0x2B, 0x3C)
COLOR_LIGHT     = RGBColor(0x55, 0x55, 0x55)
COLOR_BG_LIGHT  = "E8F4F8"
FONT_MAIN = "Calibri"
FONT_HEADING = "Calibri"


@dataclass
class CVData:
    nama: str = "Dr. Ir. Ari Berty Rondonuwu, M.Si."
    jabatan: str = "Lektor (Penata)"
    nip: str = "196801291993031001"
    nidn: str = "0029016804"
    ttl: str = "Tareran, 29 Januari 1968"
    alamat: str = "Perumahan Duta Graha Blok B No. 14A, Manado-95136"
    email: str = "arirondonuwu@unsrat.ac.id"
    telepon: str = "081356033368"
    afiliasi: str = "Fakultas Perikanan dan Ilmu Kelautan, Universitas Sam Ratulangi"
    
    pendidikan: str = ""
    peran_tim: str = "Ketua Tim Pelaksana"
    peran_teknis: str = "Ahli Rehabilitasi Terumbu Karang (Team Leader)"
    penelitian: str = ""
    publikasi: str = ""
    restorasi: str = ""
    pengabdian: str = ""
    kebijakan: str = ""
    keahlian: str = ""
    sertifikasi: str = ""
    foto_b64: Optional[str] = None


def set_cell_shading(cell, color_hex: str):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    shading.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge, val in kwargs.items():
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), val.get("val", "single"))
        element.set(qn("w:sz"), val.get("sz", "4"))
        element.set(qn("w:color"), val.get("color", "auto"))
        element.set(qn("w:space"), val.get("space", "0"))
        tcBorders.append(element)
    tcPr.append(tcBorders)


def add_styled_heading(doc, text: str, level: int = 1):
    p = doc.add_paragraph()
    p.space_before = Pt(14)
    p.space_after = Pt(6)
    
    run = p.add_run(text.upper())
    run.font.size = Pt(12) # Minimal 12pt
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    run.font.name = FONT_HEADING
    
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "00B4D8")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def add_bullet(doc, text: str, bold_prefix: str = ""):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(1.0)
    
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.size = Pt(12) # Minimal 12pt
        run_b.font.name = FONT_MAIN
    
    run = p.add_run(text)
    run.font.size = Pt(12) # Minimal 12pt
    run.font.name = FONT_MAIN
    run.font.color.rgb = COLOR_TEXT
    return p


def add_highlight_box(doc, title: str, items: List[str]):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_shading(cell, COLOR_BG_LIGHT)
    set_cell_border(cell, 
        top={"val": "none"}, bottom={"val": "none"},
        left={"val": "single", "sz": "12", "color": "0077B6"},
        right={"val": "none"}
    )
    
    p_title = cell.paragraphs[0]
    p_title.space_before = Pt(4)
    run_t = p_title.add_run(f"📌 {title}")
    run_t.bold = True
    run_t.font.size = Pt(12) # Minimal 12pt
    run_t.font.color.rgb = COLOR_PRIMARY
    run_t.font.name = FONT_MAIN
    
    for item in items:
        p_item = cell.add_paragraph()
        p_item.paragraph_format.space_after = Pt(2)
        p_item.paragraph_format.left_indent = Cm(0.5)
        run_i = p_item.add_run(f"▸ {item}")
        run_i.font.size = Pt(12) # Minimal 12pt
        run_i.font.name = FONT_MAIN
        run_i.font.color.rgb = COLOR_TEXT
    
    return tbl


def generate_cv_docx(data: CVData) -> bytes:
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
    
    # 1. HEADER
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    cell_left = header_table.cell(0, 0)
    cell_left.width = Inches(4.5)
    
    p_name = cell_left.paragraphs[0]
    run_name = p_name.add_run(data.nama)
    run_name.font.size = Pt(18) # Nama tetap besar untuk hierarki
    run_name.font.bold = True
    run_name.font.color.rgb = COLOR_PRIMARY
    run_name.font.name = FONT_MAIN
    
    p_jab = cell_left.add_paragraph()
    p_jab.space_before = Pt(2)
    run_jab = p_jab.add_run(data.jabatan)
    run_jab.font.size = Pt(12) # Minimal 12pt
    run_jab.font.color.rgb = COLOR_SECONDARY
    run_jab.font.name = FONT_MAIN
    
    p_role = cell_left.add_paragraph()
    p_role.space_before = Pt(4)
    run_role = p_role.add_run(f"🎯 {data.peran_tim}")
    run_role.font.size = Pt(12) # Minimal 12pt
    run_role.font.bold = True
    run_role.font.color.rgb = COLOR_TEXT
    run_role.font.name = FONT_MAIN
    
    run_role2 = p_role.add_run(f" — {data.peran_teknis}")
    run_role2.font.size = Pt(12) # Minimal 12pt
    run_role2.font.color.rgb = COLOR_LIGHT
    run_role2.font.name = FONT_MAIN
    
    cell_right = header_table.cell(0, 1)
    cell_right.width = Inches(1.5)
    p_photo = cell_right.paragraphs[0]
    p_photo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    if data.foto_b64:
        try:
            img_data = base64.b64decode(data.foto_b64)
            run_img = p_photo.add_run()
            run_img.add_picture(io.BytesIO(img_data), width=Inches(1.0), height=Inches(1.33))
        except Exception:
            run_ph = p_photo.add_run("[FOTO 3x4]")
            run_ph.font.size = Pt(12)
            run_ph.font.color.rgb = COLOR_LIGHT
    else:
        run_ph = p_photo.add_run("[FOTO 3x4]")
        run_ph.font.size = Pt(12)
        run_ph.font.color.rgb = COLOR_LIGHT
    
    for row in header_table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={"val": "none"}, bottom={"val": "none"}, left={"val": "none"}, right={"val": "none"})
    
    p_sep = doc.add_paragraph()
    p_sep.space_before = Pt(4)
    p_sep.space_after = Pt(4)
    pPr = p_sep._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "0077B6")
    pBdr.append(bottom)
    pPr.append(pBdr)
    
    # 2. IDENTITAS
    add_styled_heading(doc, "A. Identitas Diri", level=1)
    
    info_table = doc.add_table(rows=7, cols=2)
    info_table.style = "Table Grid"
    info_data = [
        ("Nama Lengkap", data.nama),
        ("NIP / NIDN", f"{data.nip} / {data.nidn}"),
        ("Tempat, Tanggal Lahir", data.ttl),
        ("Jabatan Fungsional", data.jabatan),
        ("Alamat", data.alamat),
        ("Email", data.email),
        ("Telepon / HP", data.telepon),
    ]
    
    for i, (label, val) in enumerate(info_data):
        cell_l = info_table.cell(i, 0)
        cell_r = info_table.cell(i, 1)
        cell_l.width = Inches(1.2)
        cell_r.width = Inches(4.8)
        
        set_cell_shading(cell_l, COLOR_BG_LIGHT)
        
        r_l = cell_l.paragraphs[0].add_run(label)
        r_l.bold = True
        r_l.font.size = Pt(12) # Minimal 12pt
        r_l.font.name = FONT_MAIN
        r_l.font.color.rgb = COLOR_PRIMARY
        
        r_r = cell_r.paragraphs[0].add_run(val)
        r_r.font.size = Pt(12) # Minimal 12pt
        r_r.font.name = FONT_MAIN
        r_r.font.color.rgb = COLOR_TEXT
    
    # 3. RESTORASI
    if data.restorasi and data.restorasi.strip():
        add_styled_heading(doc, "B. Kepakaran Khusus: Restorasi Terumbu Karang", level=1)
        
        items = []
        for line in data.restorasi.split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                lokasi = parts[0] if len(parts) > 0 else ""
                metode = parts[1] if len(parts) > 1 else ""
                periode = parts[2] if len(parts) > 2 else ""
                mitra = parts[3] if len(parts) > 3 else ""
                items.append(f"{lokasi} — {metode} ({periode}) | Mitra: {mitra}")
            else:
                items.append(line)
        
        add_highlight_box(doc, "Lokasi & Kegiatan Restorasi", items)
    
    # 4. PENDIDIKAN
    if data.pendidikan and data.pendidikan.strip():
        add_styled_heading(doc, "C. Riwayat Pendidikan", level=1)
        
        for line in data.pendidikan.split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                jenjang = parts[0] if len(parts) > 0 else ""
                institusi = parts[1] if len(parts) > 1 else ""
                bidang = parts[2] if len(parts) > 2 else ""
                tahun = parts[3] if len(parts) > 3 else ""
                
                p = doc.add_paragraph()
                p.space_after = Pt(3)
                
                r1 = p.add_run(f"{jenjang} — ")
                r1.bold = True
                r1.font.size = Pt(12) # Minimal 12pt
                r1.font.name = FONT_MAIN
                r1.font.color.rgb = COLOR_SECONDARY
                
                r2 = p.add_run(f"{institusi}")
                r2.font.size = Pt(12) # Minimal 12pt
                r2.font.name = FONT_MAIN
                
                r3 = p.add_run(f" ({bidang}, {tahun})")
                r3.font.size = Pt(12) # Minimal 12pt
                r3.font.name = FONT_MAIN
                r3.font.color.rgb = COLOR_LIGHT
            else:
                add_bullet(doc, line)
    
    # 5. PENELITIAN
    if data.penelitian and data.penelitian.strip():
        add_styled_heading(doc, "D. Pengalaman Penelitian (5 Tahun Terakhir)", level=1)
        
        for line in data.penelitian.split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                tahun = parts[0] if len(parts) > 0 else ""
                judul = parts[1] if len(parts) > 1 else ""
                dana = parts[2] if len(parts) > 2 else ""
                
                p = doc.add_paragraph()
                p.space_after = Pt(3)
                
                r1 = p.add_run(f"[{tahun}] ")
                r1.bold = True
                r1.font.size = Pt(12) # Minimal 12pt
                r1.font.color.rgb = COLOR_SECONDARY
                r1.font.name = FONT_MAIN
                
                r2 = p.add_run(judul)
                r2.font.size = Pt(12) # Minimal 12pt
                r2.font.name = FONT_MAIN
                
                r3 = p.add_run(f" — {dana}")
                r3.italic = True
                r3.font.size = Pt(12) # Minimal 12pt
                r3.font.color.rgb = COLOR_LIGHT
                r3.font.name = FONT_MAIN
    
    # 6. PUBLIKASI
    if data.publikasi and data.publikasi.strip():
        add_styled_heading(doc, "E. Publikasi Ilmiah Terpilih", level=1)
        
        for line in data.publikasi.split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                tahun = parts[0] if len(parts) > 0 else ""
                judul = parts[1] if len(parts) > 1 else ""
                jurnal = parts[2] if len(parts) > 2 else ""
                indeks = parts[3] if len(parts) > 3 else ""
                
                badge = f"[{indeks}] " if indeks else ""
                add_bullet(doc, f"{judul} ({jurnal}, {tahun})", bold_prefix=f"📄 {badge}")
    
    # 7. PENGABDIAN & KEBIJAKAN
    if data.pengabdian and data.pengabdian.strip():
        add_styled_heading(doc, "F. Pengabdian kepada Masyarakat", level=1)
        for line in data.pengabdian.split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                tahun = parts[0] if len(parts) > 0 else ""
                kegiatan = parts[1] if len(parts) > 1 else ""
                lokasi = parts[2] if len(parts) > 2 else ""
                add_bullet(doc, f"{kegiatan} ({lokasi})", bold_prefix=f"🤝 [{tahun}] ")
    
    if data.kebijakan and data.kebijakan.strip():
        add_styled_heading(doc, "G. Perumusan Kebijakan Publik", level=1)
        for line in data.kebijakan.split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                tahun = parts[0] if len(parts) > 0 else ""
                kegiatan = parts[1] if len(parts) > 1 else ""
                lingkup = parts[2] if len(parts) > 2 else ""
                add_bullet(doc, f"{kegiatan} ({lingkup})", bold_prefix=f"📋 [{tahun}] ")
    
    # 8. SERTIFIKASI
    if data.sertifikasi and data.sertifikasi.strip():
        add_styled_heading(doc, "H. Sertifikasi & Pengalaman Selam", level=1)
        for line in data.sertifikasi.split("\n"):
            line = line.strip()
            if line: add_bullet(doc, line, bold_prefix="🤿 ")
    
    # 9. KEAHLIAN
    if data.keahlian and data.keahlian.strip():
        add_styled_heading(doc, "I. Keahlian Inti", level=1)
        p = doc.add_paragraph()
        p.space_before = Pt(4)
        keahlian_list = [k.strip() for k in data.keahlian.split(",") if k.strip()]
        for i, k in enumerate(keahlian_list):
            run = p.add_run(k)
            run.font.size = Pt(12) # Minimal 12pt
            run.font.name = FONT_MAIN
            run.font.color.rgb = COLOR_TEXT
            if i < len(keahlian_list) - 1:
                run_sep = p.add_run("  •  ")
                run_sep.font.size = Pt(12)
                run_sep.font.color.rgb = COLOR_ACCENT
    
    # FOOTER
    p_footer = doc.add_paragraph()
    p_footer.space_before = Pt(20)
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_f = p_footer.add_run("— Dokumen ini dibuat secara otomatis oleh CV Generator —")
    run_f.font.size = Pt(12) # Minimal 12pt
    run_f.font.italic = True
    run_f.font.color.rgb = COLOR_LIGHT
    run_f.font.name = FONT_MAIN
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()