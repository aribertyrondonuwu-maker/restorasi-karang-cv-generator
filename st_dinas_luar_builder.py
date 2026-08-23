    # 2. KOP SURAT DENGAN LOGO & ALAMAT LENGKAP
    logo_path = "logo_unsrat.png"
    has_logo = os.path.exists(logo_path)
    
    if has_logo:
        # Layout Logo Kiri + Teks Kanan menggunakan Tabel Invisible
        kop_tbl = doc.add_table(rows=1, cols=2)
        kop_tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        cell_logo = kop_tbl.cell(0, 0)
        cell_logo.width = Inches(1.0)
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_img = p_logo.add_run()
        run_img.add_picture(logo_path, width=Inches(0.9))
        
        cell_text = kop_tbl.cell(0, 1)
        cell_text.width = Inches(5.0)
        p_kop = cell_text.paragraphs[0]
        p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Isi Teks Kop Surat Sesuai Standar FPIK
        r1 = p_kop.add_run("KEMENTERIAN PENDIDIKAN TINGGI, SAINS, DAN TEKNOLOGI\n")
        r1.font.name = FONT_OFFICIAL; r1.font.size = Pt(12); r1.bold = True
        
        r2 = p_kop.add_run("UNIVERSITAS SAM RATULANGI\n")
        r2.font.name = FONT_OFFICIAL; r2.font.size = Pt(12); r2.bold = True
        
        r3 = p_kop.add_run("FAKULTAS PERIKANAN DAN ILMU KELAUTAN\n")
        r3.font.name = FONT_OFFICIAL; r3.font.size = Pt(12); r3.bold = True
        
        # Baris Alamat Lengkap (Sesuai Gambar Referensi)
        r4 = p_kop.add_run("Alamat : Kampus UNSRAT Manado 95115  Laman : http://fpik.unsrat.ac.id ; Email: fpik@unsrat.ac.id")
        r4.font.name = FONT_OFFICIAL; r4.font.size = Pt(12)
        
        # Hapus border tabel kop agar invisible
        for row in kop_tbl.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = OxmlElement('w:tcBorders')
                for edge in ['top', 'left', 'bottom', 'right']:
                    element = OxmlElement(f'w:{edge}')
                    element.set(qn('w:val'), 'none')
                    tcBorders.append(element)
                tcPr.append(tcBorders)
                
    else:
        # Fallback jika logo tidak ditemukan (Teks tetap di tengah)
        p_kop = doc.add_paragraph()
        p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        r1 = p_kop.add_run("KEMENTERIAN PENDIDIKAN TINGGI, SAINS, DAN TEKNOLOGI\n")
        r1.font.name = FONT_OFFICIAL; r1.font.size = Pt(12); r1.bold = True
        
        r2 = p_kop.add_run("UNIVERSITAS SAM RATULANGI\n")
        r2.font.name = FONT_OFFICIAL; r2.font.size = Pt(12); r2.bold = True
        
        r3 = p_kop.add_run("FAKULTAS PERIKANAN DAN ILMU KELAUTAN\n")
        r3.font.name = FONT_OFFICIAL; r3.font.size = Pt(12); r3.bold = True
        
        # Baris Alamat Lengkap (Sesuai Gambar Referensi)
        r4 = p_kop.add_run("Alamat : Kampus UNSRAT Manado 95115  Laman : http://fpik.unsrat.ac.id ; Email: fpik@unsrat.ac.id")
        r4.font.name = FONT_OFFICIAL; r4.font.size = Pt(12)
