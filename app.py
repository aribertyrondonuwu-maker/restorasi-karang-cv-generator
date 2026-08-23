"""app.py — Web Generator Terintegrasi: CV & Surat Tugas Dinas Luar"""
import streamlit as st
import base64
from datetime import date
from cv_builder import CVData, generate_cv_docx
from st_dinas_luar_builder import DataSuratTugas, generate_st_dinas_luar

st.set_page_config(page_title="Generator Tim Pelaksana PPS Bitung", page_icon="🌊", layout="wide")

st.title("🌊 Generator Dokumen Tim Pelaksana Swakelola")
st.subheader("Relokasi & Restorasi Terumbu Karang PPS Bitung TA 2026")
st.markdown("---")

tab1, tab2 = st.tabs(["📄 Generator CV Modern", "✉️ Generator Surat Tugas Dinas Luar"])

# ==========================================
# TAB 1: GENERATOR CV
# ==========================================
with tab1:
    st.subheader("👤 Identitas Diri")
    col1, col2 = st.columns(2)
    with col1:
        nama = st.text_input("Nama Lengkap (dengan gelar)", value="Dr. Ir. Ari Berty Rondonuwu, M.Si.")
        nip = st.text_input("NIP", value="196801291993031001")
        ttl = st.text_input("Tempat, Tanggal Lahir", value="Tareran, 29 Januari 1968")
    with col2:
        jabatan = st.text_input("Jabatan Fungsional", value="Lektor (Penata)")
        nidn = st.text_input("NIDN", value="0029016804")
        afiliasi = st.text_input("Afiliasi Institusi", value="FPIK Universitas Sam Ratulangi")
    
    col3, col4 = st.columns(2)
    with col3:
        alamat = st.text_area("Alamat Rumah", value="Perumahan Duta Graha Blok B No. 14A, Manado-95136", height=60)
        email = st.text_input("Email Aktif", value="arirondonuwu@unsrat.ac.id")
    with col4:
        telepon = st.text_input("No. Telepon/HP", value="081356033368")
        foto = st.file_uploader("Unggah Pas Foto (3x4)", type=['jpg', 'png'])

    st.subheader("🎯 Peran dalam Tim Swakelola")
    col5, col6 = st.columns(2)
    with col5:
        peran_tim = st.selectbox("Kedudukan dalam Tim", ["Ketua Tim Pelaksana", "Anggota Tim Pelaksana", "Pembantu Peneliti", "Penyelam Bersertifikat"])
    with col6:
        peran_teknis = st.text_input("Peran Teknis", value="Ahli Rehabilitasi Terumbu Karang (Team Leader)")

    st.subheader("🎓 Riwayat Pendidikan")
    pendidikan = st.text_area("Format: Jenjang | Institusi | Bidang | Tahun", height=100, value="S1 | Universitas Sam Ratulangi | MSP | 1986-1991\nS2 | Università Delle Marche, Italia | Biodiversity & Management of Coral Reef | 2003-2004\nS2 | Universitas Sam Ratulangi | Ilmu Perairan | 2015-2017\nS3 | Universitas Sam Ratulangi | Ilmu Kelautan | 2017-2020")

    st.subheader("🌊 Kepakaran: Restorasi Terumbu Karang")
    restorasi = st.text_area("Format: Lokasi | Metode | Periode | Mitra", height=120, value="Pantai Malalayang, Manado | Transplantasi & CSC | 2009-kini | PT. TJ Silfanus\nTN Bunaken | Terumbu buatan + transplantasi | 2018-2020 | Balai TN Bunaken\nPulau Lembeh, Bitung | Terumbu buatan | 2020-2022 | DKP Bitung")

    st.subheader("🔬 Pengalaman Penelitian (5 Tahun Terakhir)")
    penelitian = st.text_area("Format: Tahun | Judul Penelitian | Sumber Dana", height=120, value="2022 | Pemetaan Habitat Perairan Dangkal dengan UAV di Likupang | LPPM Unsrat\n2021 | Pemetaan Ekosistem Terumbu Karang Pulau Serena dengan UAV | LPPM Unsrat\n2019-2020 | Optimalisasi Artificial Reef dan Restorasi Terumbu Karang di Poopoh | DRPM DIKTI")

    st.subheader("📚 Publikasi Ilmiah Terpilih")
    publikasi = st.text_area("Format: Tahun | Judul | Jurnal | Indeks", height=100, value="2021 | Shallow water habitat mapping with UAV in Serena Island | AACL Bioflux 14(6) | Scopus Q3\n2020 | Mitochondrial CO1 sequences of Banggai Cardinalfish | AACL Bioflux 13(2) | Scopus Q3")

    st.subheader("🤝 Pengabdian & Kebijakan Publik")
    pengabdian = st.text_area("Format: Tahun | Kegiatan | Lokasi", height=80, value="2020 | Narasumber PEN Restorasi Terumbu Karang (ICRG) | Bali\n2021 | Transplantasi Karang Pantai Malalayang | Manado")
    kebijakan = st.text_area("Format: Tahun | Kegiatan | Lingkup", height=80, value="2021 | Tim Integrasi Perda RZWP3K ke Revisi Perda RTRW Sulut | Provinsi Sulut\n2020 | Tim Penyusun RPJMD Provinsi Sulawesi Utara 2021-2026 | Provinsi Sulut")

    st.subheader("🛠️ Keahlian Inti")
    keahlian = st.text_input("Pisahkan dengan koma", value="Koralogi, Ekologi Pesisir, Restorasi Terumbu Karang, UAV/Drone Mapping, GIS, DNA Barcoding, Scientific Diving")

    if st.button("⚡ Generate CV Modern", type="primary", use_container_width=True):
        foto_b64 = base64.b64encode(foto.getvalue()).decode('utf-8') if foto else None
        data = CVData(nama=nama, jabatan=jabatan, nip=nip, nidn=nidn, ttl=ttl, alamat=alamat, email=email, telepon=telepon, afiliasi=afiliasi, pendidikan=pendidikan, peran_tim=peran_tim, peran_teknis=peran_teknis, penelitian=penelitian, publikasi=publikasi, restorasi=restorasi, pengabdian=pengabdian, kebijakan=kebijakan, keahlian=keahlian, foto_b64=foto_b64)
        
        with st.spinner("Menyusun dokumen CV modern..."):
            docx_bytes = generate_cv_docx(data)
        st.success("✅ CV berhasil disusun!")
        st.download_button(label="⬇️ Unduh CV (.docx)", data=docx_bytes, file_name=f"CV_{nama.replace(' ', '_').replace('.', '')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

# ==========================================
# TAB 2: GENERATOR SURAT TUGAS
# ==========================================
with tab2:
    st.info("ℹ️ Surat ini otomatis menyertakan klausul pembebasan tugas mengajar dan larangan rangkap bayar sesuai Diktum KETUJUH SK Dekan.")
    col1, col2 = st.columns(2)
    with col1:
        st_nama = st.text_input("Nama Personil", value="Dr. Ir. Ari Berty Rondonuwu, M.Si.")
        st_nip = st.text_input("NIP", value="196801291993031001")
        st_jabatan = st.text_input("Jabatan", value="Lektor (Penata) / Dosen FPIK")
    with col2:
        st_peran = st.selectbox("Peran dalam Tim", ["Ketua Tim Pelaksana", "Anggota Tim Pelaksana", "Pembantu Peneliti"])
        st_jenis = st.text_input("Jenis Tugas", value="Survei Awal Ekologi dan Pemetaan Dasar")
        st_lokasi = st.text_input("Lokasi Tugas", value="Pelabuhan Perikanan Samudera (PPS) Bitung, Sulawesi Utara")
    
    col3, col4, col5 = st.columns(3)
    with col3: tgl_mulai = st.date_input("Tanggal Mulai", value=date(2026, 9, 15))
    with col4: tgl_selesai = st.date_input("Tanggal Selesai", value=date(2026, 9, 16))
    with col5: jml_hari = st.number_input("Jumlah Hari", min_value=1, value=2)

    klaim_8_oj = st.checkbox("Aktifkan Klausul 8 OJ/hari (Khusus Pembantu Peneliti)", value=(st_peran == "Pembantu Peneliti"))

    if st.button("✉️ Generate Surat Tugas Dinas Luar", type="primary", use_container_width=True):
        data_st = DataSuratTugas(
            nama=st_nama, nip=st_nip, jabatan=st_jabatan, peran_tim=st_peran,
            jenis_tugas=st_jenis, lokasi=st_lokasi, tanggal_mulai=tgl_mulai,
            tanggal_selesai=tgl_selesai, jumlah_hari=jml_hari, klaim_8_oj=klaim_8_oj
        )
        with st.spinner("Menyusun Surat Tugas..."):
            st_bytes = generate_st_dinas_luar(data_st)
        st.success("✅ Surat Tugas berhasil disusun!")
        st.download_button(label="⬇️ Unduh Surat Tugas (.docx)", data=st_bytes, file_name=f"ST_DL_{st_nama.replace(' ', '_').replace('.', '')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)