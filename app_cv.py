"""
app_cv.py — Web Generator CV & Surat Tugas Dinas Luar
"""
import streamlit as st
import base64
from datetime import date
from cv_builder import CVData, generate_cv_docx
from st_dinas_luar_builder import DataSuratTugas, generate_st_dinas_luar

st.set_page_config(page_title="Generator Tim PPS Bitung", page_icon="🌊", layout="wide")

st.title("Generator Dokumen Tim Pelaksana Swakelola")
st.subheader("Relokasi & Restorasi Terumbu Karang PPS Bitung TA 2026")
st.markdown("---")

tab1, tab2 = st.tabs(["Generator CV", "Generator Surat Tugas Dinas Luar"])

# ==========================================
# TAB 1: CV
# ==========================================
with tab1:
    st.info("CV ini dirancang untuk 15 orang wajib: 7 Tim Pelaksana ASN + 8 Penyelam Bersertifikat")

    with st.form("cv_form", clear_on_submit=False):
        st.subheader("Identitas Diri")
        cv_col1, cv_col2 = st.columns(2)
        with cv_col1:
            cv_nama = st.text_input("Nama Lengkap (dengan gelar)", value="Dr. Ir. Ari Berty Rondonuwu, M.Si.", key="cv_nama_input")
            cv_nip  = st.text_input("NIP", value="196801291993031001", key="cv_nip_input")
            cv_ttl  = st.text_input("Tempat, Tanggal Lahir", value="Tareran, 29 Januari 1968", key="cv_ttl_input")
        with cv_col2:
            cv_jabatan = st.text_input("Jabatan Fungsional", value="Lektor (Penata)", key="cv_jabatan_input")
            cv_nidn    = st.text_input("NIDN", value="0029016804", key="cv_nidn_input")
            cv_afiliasi = st.text_input("Afiliasi Institusi", value="FPIK Universitas Sam Ratulangi", key="cv_afiliasi_input")

        cv_col3, cv_col4 = st.columns(2)
        with cv_col3:
            cv_alamat = st.text_area("Alamat Rumah", value="Perumahan Duta Graha Blok B No. 14A, Manado-95136", height=60, key="cv_alamat_input")
            cv_email  = st.text_input("Email Aktif", value="arirondonuwu@unsrat.ac.id", key="cv_email_input")
        with cv_col4:
            cv_telepon = st.text_input("No. Telepon/HP", value="081356033368", key="cv_telepon_input")
            cv_foto    = st.file_uploader("Unggah Pas Foto (3x4)", type=["jpg", "png", "jpeg"], key="cv_foto_input")

        st.subheader("Peran dalam Tim Swakelola")
        cv_col5, cv_col6 = st.columns(2)
        with cv_col5:
            cv_peran = st.selectbox("Kedudukan dalam Tim", ["Ketua Tim Pelaksana", "Anggota Tim Pelaksana", "Pembantu Peneliti", "Penyelam Bersertifikat"], key="cv_peran_input")
        with cv_col6:
            cv_peran_teknis = st.text_input("Peran Teknis", value="Ahli Rehabilitasi Terumbu Karang (Team Leader)", key="cv_peran_teknis_input")

        st.subheader("Riwayat Pendidikan")
        cv_pendidikan = st.text_area(
            "Format: Jenjang | Institusi | Bidang | Tahun", height=100,
            value="S1 | Universitas Sam Ratulangi | MSP | 1986-1991\nS2 | Università Delle Marche, Italia | Biodiversity & Management of Coral Reef | 2003-2004\nS2 | Universitas Sam Ratulangi | Ilmu Perairan | 2015-2017\nS3 | Universitas Sam Ratulangi | Ilmu Kelautan | 2017-2020",
            key="cv_pendidikan_input"
        )

        st.subheader("Kepakaran: Restorasi Terumbu Karang")
        cv_restorasi = st.text_area(
            "Format: Lokasi | Metode | Periode | Mitra", height=120,
            value="Pantai Malalayang, Manado | Transplantasi & CSC | 2009-kini | PT. TJ Silfanus\nTN Bunaken | Terumbu buatan + transplantasi | 2018-2020 | Balai TN Bunaken\nPulau Lembeh, Bitung | Terumbu buatan | 2020-2022 | DKP Bitung\nPoopoh, Minahasa | Transplantasi partisipatif | 2015-2020 | Masyarakat lokal",
            key="cv_restorasi_input"
        )

        st.subheader("Pengalaman Penelitian (5 Tahun Terakhir)")
        cv_penelitian = st.text_area(
            "Format: Tahun | Judul Penelitian | Sumber Dana", height=120,
            value="2022 | Pemetaan Habitat Perairan Dangkal dengan UAV di Likupang | LPPM Unsrat\n2021 | Pemetaan Ekosistem Terumbu Karang Pulau Serena dengan UAV | LPPM Unsrat\n2019-2020 | Optimalisasi Artificial Reef dan Restorasi Terumbu Karang di Poopoh | DRPM DIKTI\n2019-2020 | Monitoring Kesehatan Terumbu Karang di Raja Ampat | COREMAP-CTI LIPI",
            key="cv_penelitian_input"
        )

        st.subheader("Publikasi Ilmiah Terpilih")
        cv_publikasi = st.text_area(
            "Format: Tahun | Judul | Jurnal | Indeks", height=100,
            value="2021 | Shallow water habitat mapping with UAV in Serena Island | AACL Bioflux 14(6) | Scopus Q3\n2020 | Mitochondrial CO1 sequences of Banggai Cardinalfish | AACL Bioflux 13(2) | Scopus Q3\n2023 | Biometrik Otolit Ikan Kardinal Banggai | Jurnal Ilmiah PLATAX 11(1) | SINTA",
            key="cv_publikasi_input"
        )

        st.subheader("Pengabdian & Kebijakan Publik")
        cv_col7, cv_col8 = st.columns(2)
        with cv_col7:
            cv_pengabdian = st.text_area(
                "Format: Tahun | Kegiatan | Lokasi", height=100,
                value="2020 | Narasumber PEN Restorasi Terumbu Karang (ICRG) | Bali\n2021 | Transplantasi Karang Pantai Malalayang | Manado\n2016 | Modul Pelatihan Pemantauan Terumbu Karang (CCDP-IFAD) | Sulawesi Utara",
                key="cv_pengabdian_input"
            )
        with cv_col8:
            cv_kebijakan = st.text_area(
                "Format: Tahun | Kegiatan | Lingkup", height=100,
                value="2021 | Tim Integrasi Perda RZWP3K ke Revisi Perda RTRW Sulut | Provinsi Sulut\n2020 | Tim Penyusun RPJMD Provinsi Sulawesi Utara 2021-2026 | Provinsi Sulut",
                key="cv_kebijakan_input"
            )

        st.subheader("Sertifikasi & Pengalaman Selam (Opsional)")
        cv_sertifikasi = st.text_area(
            "Format: Sertifikat / Pengalaman", height=80,
            value="Sertifikasi: CMAS 3-Star / PADI Divemaster\nPengalaman selam ilmiah: >500 jam selam\nSpesialisasi: Survei terumbu karang, transplantasi, monitoring",
            key="cv_sertifikasi_input"
        )

        st.subheader("Keahlian Inti")
        cv_keahlian = st.text_input(
            "Pisahkan dengan koma",
            value="Koralogi, Ekologi Pesisir, Restorasi Terumbu Karang, UAV/Drone Mapping, GIS, DNA Barcoding, Scientific Diving",
            key="cv_keahlian_input"
        )

        submit_cv = st.form_submit_button("Generate CV", type="primary", use_container_width=True)

    if submit_cv:
        foto_b64 = None
        if cv_foto:
            foto_b64 = base64.b64encode(cv_foto.getvalue()).decode("utf-8")

        data_cv = CVData(
            nama=cv_nama, jabatan=cv_jabatan, nip=cv_nip, nidn=cv_nidn, ttl=cv_ttl,
            alamat=cv_alamat, email=cv_email, telepon=cv_telepon, afiliasi=cv_afiliasi,
            pendidikan=cv_pendidikan, peran_tim=cv_peran, peran_teknis=cv_peran_teknis,
            penelitian=cv_penelitian, publikasi=cv_publikasi, restorasi=cv_restorasi,
            pengabdian=cv_pengabdian, kebijakan=cv_kebijakan, keahlian=cv_keahlian,
            sertifikasi=cv_sertifikasi, foto_b64=foto_b64
        )

        with st.spinner("Menyusun dokumen CV..."):
            docx_bytes = generate_cv_docx(data_cv)

        st.success("CV berhasil disusun!")
        st.download_button(
            label="Unduh CV (.docx)",
            data=docx_bytes,
            file_name=f"CV_{cv_nama.replace(' ', '_').replace('.', '')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="cv_download_button"
        )

# ==========================================
# TAB 2: SURAT TUGAS
# ==========================================
with tab2:
    st.info("Surat ini otomatis menyertakan klausul pembebasan tugas mengajar dan larangan rangkap bayar sesuai Diktum KETUJUH SK Dekan.")

    with st.form("st_form", clear_on_submit=False):
        st.subheader("Data Personil")
        st_col1, st_col2 = st.columns(2)
        with st_col1:
            st_nama    = st.text_input("Nama Personil", value="Dr. Ir. Ari Berty Rondonuwu, M.Si.", key="st_nama_input")
            st_nip     = st.text_input("NIP", value="196801291993031001", key="st_nip_input")
            st_jabatan = st.text_input("Jabatan", value="Lektor (Penata) / Dosen FPIK", key="st_jabatan_input")
        with st_col2:
            st_peran  = st.selectbox("Peran dalam Tim", ["Ketua Tim Pelaksana", "Anggota Tim Pelaksana", "Pembantu Peneliti"], key="st_peran_input")
            st_jenis  = st.text_input("Jenis Tugas", value="Survei Awal Ekologi dan Pemetaan Dasar", key="st_jenis_input")
            st_lokasi = st.text_input("Lokasi Tugas", value="Pelabuhan Perikanan Samudera (PPS) Bitung, Sulawesi Utara", key="st_lokasi_input")

        st.subheader("Periode Penugasan")
        st_col3, st_col4, st_col5 = st.columns(3)
        with st_col3:
            tgl_mulai   = st.date_input("Tanggal Mulai",   value=date(2026, 9, 15), key="st_tgl_mulai_input")
        with st_col4:
            tgl_selesai = st.date_input("Tanggal Selesai", value=date(2026, 9, 16), key="st_tgl_selesai_input")
        with st_col5:
            jml_hari = st.number_input("Jumlah Hari", min_value=1, max_value=30, value=2, key="st_jml_hari_input")

        st.subheader("Nomor Surat")
        st_nomor = st.text_input("Nomor Surat", value="800/.../FPIK-UNSRAT/2026", key="st_nomor_input")

        st.subheader("Klausul Khusus")
        klaim_8_oj = st.checkbox(
            "Aktifkan Klausul 8 OJ/hari (Khusus Pembantu Peneliti)",
            value=(st_peran == "Pembantu Peneliti"),
            key="st_klaim_8oj_input"
        )

        submit_st = st.form_submit_button("Generate Surat Tugas Dinas Luar", type="primary", use_container_width=True)

    if submit_st:
        data_st = DataSuratTugas(
            nomor_surat=st_nomor,
            nama=st_nama, nip=st_nip, jabatan=st_jabatan,
            peran_tim=st_peran, jenis_tugas=st_jenis, lokasi=st_lokasi,
            tanggal_mulai=tgl_mulai, tanggal_selesai=tgl_selesai,
            jumlah_hari=jml_hari, klaim_8_oj=klaim_8_oj
        )

        with st.spinner("Menyusun Surat Tugas..."):
            st_bytes = generate_st_dinas_luar(data_st)

        st.success("Surat Tugas berhasil disusun!")
        st.download_button(
            label="Unduh Surat Tugas (.docx)",
            data=st_bytes,
            file_name=f"ST_DL_{st_nama.replace(' ', '_').replace('.', '')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="st_download_button"
        )

st.markdown("---")
st.caption("Generator Dokumen Tim Pelaksana Swakelola — FPIK UNSRAT 2026")
