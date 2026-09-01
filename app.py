"""
app.py — Titik masuk utama aplikasi web.

Web Generator CV & Surat Tugas Dinas Luar untuk Tim Pelaksana kegiatan
Relokasi dan Replanting/Restorasi Terumbu Karang PPS Kota Bitung TA 2026,
Fakultas Perikanan dan Ilmu Kelautan, Universitas Sam Ratulangi.

Jalankan dengan perintah:  streamlit run app.py
"""

from datetime import date

import pandas as pd
import streamlit as st

import utils
from utils import (
    NAMA_KEGIATAN_PENDEK, LOGO_PATH, logo_tersedia, BATAS_MB,
    LABEL_DOKUMEN, MAKS_SERTIFIKAT, validasi_ukuran, adalah_pdf,
    nama_berkas_aman,
)
from cv_builder import (
    CVData, Lampiran, generate_cv_pdf, generate_cv_docx,
    JENIS_TENAGA_AHLI, JENIS_PENYELAM,
)
from st_dinas_luar_builder import (
    DataSuratTugas, generate_st_dinas_luar, generate_st_dinas_luar_pdf,
)

# =====================================================================
# KONFIGURASI HALAMAN DAN TEMA
# =====================================================================

st.set_page_config(
    page_title="Generator Dokumen Tim Pelaksana — FPIK UNSRAT",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BIRU_LAUT = "#1B3F6B"
HIJAU_KARANG = "#2D7D46"

GAYA_APLIKASI = f"""
<style>
:root {{
    --biru-laut: {BIRU_LAUT};
    --hijau-karang: {HIJAU_KARANG};
}}

/* Kepala halaman */
.kepala-aplikasi {{
    background: {BIRU_LAUT};
    padding: 18px 24px;
    border-radius: 10px;
    margin-bottom: 8px;
}}
.kepala-judul {{
    color: #FFFFFF;
    font-size: 24px;
    font-weight: 700;
    margin: 0;
    line-height: 1.3;
}}
.kepala-subjudul {{
    color: #C8DCEC;
    font-size: 15px;
    margin: 4px 0 0 0;
}}
.kepala-instansi {{
    color: {HIJAU_KARANG};
    background: #E7F3EB;
    display: inline-block;
    padding: 2px 10px;
    border-radius: 4px;
    font-size: 13px;
    font-weight: 600;
    margin-top: 8px;
}}

/* Judul bagian form */
h3 {{
    color: {BIRU_LAUT} !important;
    border-bottom: 2px solid {HIJAU_KARANG};
    padding-bottom: 6px;
    margin-top: 22px !important;
}}

/* Tombol utama */
div.stButton > button[kind="primary"],
div.stDownloadButton > button {{
    background-color: {BIRU_LAUT};
    color: #FFFFFF;
    border: none;
    font-weight: 600;
}}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button:hover {{
    background-color: {HIJAU_KARANG};
    color: #FFFFFF;
}}

/* Tab */
.stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {{
    color: {BIRU_LAUT};
    font-weight: 700;
}}

/* Kartu status unggahan */
.kartu-berkas {{
    border: 1px solid #D6E0EA;
    border-left: 4px solid {HIJAU_KARANG};
    background: #F6FAFD;
    border-radius: 6px;
    padding: 8px 12px;
    font-size: 13px;
    margin-top: 6px;
}}
.kartu-berkas-kosong {{
    border-left-color: #C0392B;
    background: #FDF3F2;
}}
</style>
"""
st.markdown(GAYA_APLIKASI, unsafe_allow_html=True)


# =====================================================================
# GERBANG KATA SANDI
# =====================================================================

def periksa_kata_sandi() -> bool:
    """Menampilkan formulir kata sandi dan menahan akses sampai lolos.

    Kata sandi diambil dari Streamlit Secrets (kunci APP_PASSWORD) agar
    tidak tertulis langsung di kode sumber yang bersifat publik. Apabila
    secrets belum diatur, gerbang ini dilewati secara otomatis supaya
    aplikasi tetap bisa dijalankan saat pengembangan lokal.
    """
    kata_sandi_baku = st.secrets.get("APP_PASSWORD", None) if hasattr(
        st, "secrets") else None

    # Tidak ada kata sandi yang diatur -> lewati gerbang (mode pengembangan)
    if not kata_sandi_baku:
        return True

    if st.session_state.get("sudah_login", False):
        return True

    st.markdown(
        "<div style='max-width:420px;margin:80px auto 0;text-align:center;'>"
        "<h3 style='color:#1B3F6B;'>🌊 Akses Terbatas</h3>"
        "<p style='color:#666;font-size:14px;'>Masukkan kata sandi untuk "
        "membuka Generator Dokumen Tim Pelaksana Swakelola.</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    kol_kiri, kol_tengah, kol_kanan = st.columns([1, 1.2, 1])
    with kol_tengah:
        masukan = st.text_input(
            "Kata Sandi", type="password", key="masukan_kata_sandi",
            label_visibility="collapsed", placeholder="Kata sandi...",
        )
        tekan = st.button("Masuk", type="primary", width="stretch")

        if tekan:
            if masukan == kata_sandi_baku:
                st.session_state.sudah_login = True
                st.rerun()
            else:
                st.error("Kata sandi salah. Silakan coba lagi.")

    return False


# =====================================================================
# INISIALISASI SESSION STATE
# =====================================================================

TABEL_KOSONG = {
    "pendidikan": pd.DataFrame([
        {"Jenjang": "", "Jurusan": "", "Universitas": "", "Kota": "",
         "Tahun Lulus": ""}
    ]),
    "sertifikasi": pd.DataFrame([
        {"Nama Sertifikat": "", "Nomor": "", "Lembaga Penerbit": "",
         "Tahun": "", "Masa Berlaku": ""}
    ]),
    "pengalaman_kerja": pd.DataFrame([
        {"Nama Pekerjaan/Proyek": "", "Tahun": "", "Jabatan/Posisi": "",
         "Instansi Pemberi Kerja": "", "Lokasi": ""} for _ in range(5)
    ]),
    "publikasi": pd.DataFrame([
        {"Judul": "", "Jurnal/Prosiding": "", "Tahun": ""}
    ]),
    "pengalaman_selam": pd.DataFrame([
        {"Lokasi Penyelaman": "", "Tahun": "", "Jenis Kegiatan": "",
         "Kedalaman Maks (m)": "", "Lama Kegiatan": "", "Pemberi Kerja": ""}
        for _ in range(5)
    ]),
    "personil_st": pd.DataFrame([
        {"Nama": "", "NIP": "", "Jabatan": "", "Peran dalam Tim": ""}
    ]),
}

BIDANG_KEAHLIAN_PILIHAN = [
    "Ekologi Laut", "Terumbu Karang", "Oseanografi", "Biologi Perairan",
    "Pengelolaan Pesisir", "Mangrove", "Lingkungan Hidup",
]

KEAHLIAN_SELAM_PILIHAN = [
    "Transplantasi Karang", "Reef Survey", "Underwater Photography",
    "Videografi Bawah Air", "ROV Assist",
    "Penandaan dan Pemetaan Terumbu Karang", "Pengambilan Sampel Biologi",
]

PERAN_TIM_PILIHAN = [
    "Ketua Tim Pelaksana", "Anggota Tim Pelaksana", "Pembantu Peneliti",
    "Penyelam Bersertifikat",
]


def siapkan_session_state():
    """Menyiapkan seluruh kunci session_state agar data form tidak hilang.

    Streamlit menjalankan ulang skrip setiap kali pengguna berinteraksi
    dengan antarmuka. Tanpa session_state, isian form akan hilang.
    """
    for nama, kerangka in TABEL_KOSONG.items():
        if nama not in st.session_state:
            st.session_state[nama] = kerangka.copy()

    berkas_awal = {
        "berkas_foto": None,
        "berkas_ktp": None,
        "berkas_lisensi": None,
        "berkas_sertifikat": [],
        "hasil_cv_pdf": None,
        "hasil_cv_docx": None,
        "hasil_st_pdf": None,
        "hasil_st_docx": None,
        "data_cv_terakhir": None,
        "cv_ringkasan_afiliasi": "",
    }
    for kunci, nilai in berkas_awal.items():
        if kunci not in st.session_state:
            st.session_state[kunci] = nilai


siapkan_session_state()


# =====================================================================
# KOMPONEN ANTARMUKA BERSAMA
# =====================================================================

def tampilkan_kepala():
    """Menampilkan kepala aplikasi berisi logo dan nama kegiatan."""
    kolom_logo, kolom_teks = st.columns([1, 9])

    with kolom_logo:
        if logo_tersedia():
            st.image(LOGO_PATH, width=78)
        else:
            st.markdown(
                "<div style='width:78px;height:78px;border:2px dashed #B0C4D8;"
                "border-radius:8px;display:flex;align-items:center;"
                "justify-content:center;color:#8FA6BC;font-size:11px;"
                "text-align:center;'>LOGO<br>INSTANSI</div>",
                unsafe_allow_html=True,
            )

    with kolom_teks:
        st.markdown(
            "<div class='kepala-aplikasi'>"
            "<p class='kepala-judul'>Generator Dokumen Tim Pelaksana "
            "Swakelola</p>"
            f"<p class='kepala-subjudul'>{NAMA_KEGIATAN_PENDEK}</p>"
            "<span class='kepala-instansi'>Fakultas Perikanan dan Ilmu "
            "Kelautan — Universitas Sam Ratulangi</span>"
            "</div>",
            unsafe_allow_html=True,
        )


def pratinjau_berkas(berkas, label: str, wajib: bool = True):
    """Menampilkan pratinjau berkas unggahan beserta status validasinya.

    Gambar ditampilkan sebagai gambar kecil, sedangkan berkas PDF
    ditampilkan sebagai kartu keterangan dengan nama berkas.
    """
    if berkas is None:
        kelas = "kartu-berkas kartu-berkas-kosong" if wajib else "kartu-berkas"
        keterangan = "Wajib diunggah" if wajib else "Opsional"
        st.markdown(
            f"<div class='{kelas}'>Belum ada berkas — {keterangan}</div>",
            unsafe_allow_html=True,
        )
        return

    if adalah_pdf(berkas.name):
        ukuran_kb = len(berkas.getvalue()) / 1024
        st.markdown(
            f"<div class='kartu-berkas'>📄 <b>{berkas.name}</b><br>"
            f"Dokumen PDF · {ukuran_kb:,.0f} KB</div>",
            unsafe_allow_html=True,
        )
    else:
        st.image(berkas.getvalue(), width=140, caption=berkas.name)


def bagian_unggah_dokumen(kunci_awalan: str = "cv"):
    """Menampilkan empat widget unggahan dokumen wajib beserta validasinya.

    Berkas yang lolos validasi ukuran disimpan ke session_state agar tetap
    tersedia meskipun antarmuka dijalankan ulang.
    """
    st.markdown("### 📎 Upload Dokumen")
    st.caption(
        "Unggah seluruh dokumen persyaratan terlebih dahulu sebelum "
        "melanjutkan pengisian data. Berkas akan disisipkan sebagai halaman "
        "lampiran pada dokumen CV."
    )

    kol1, kol2 = st.columns(2)

    # --- 1. Pas foto 3x4 ---
    with kol1:
        st.markdown(f"**1. {LABEL_DOKUMEN['foto']}** · maks {BATAS_MB['foto']} MB")
        foto = st.file_uploader(
            "Pas Foto 3x4 (JPG/PNG)",
            type=["jpg", "jpeg", "png"],
            key=f"{kunci_awalan}_unggah_foto",
            label_visibility="collapsed",
        )
        galat = validasi_ukuran(foto, "foto")
        if galat:
            st.error(galat)
        else:
            st.session_state.berkas_foto = foto
        pratinjau_berkas(st.session_state.berkas_foto, LABEL_DOKUMEN["foto"])

    # --- 2. KTP ---
    with kol2:
        st.markdown(f"**2. {LABEL_DOKUMEN['ktp']}** · maks {BATAS_MB['ktp']} MB")
        ktp = st.file_uploader(
            "KTP (JPG/PNG/PDF)",
            type=["jpg", "jpeg", "png", "pdf"],
            key=f"{kunci_awalan}_unggah_ktp",
            label_visibility="collapsed",
        )
        galat = validasi_ukuran(ktp, "ktp")
        if galat:
            st.error(galat)
        else:
            st.session_state.berkas_ktp = ktp
        pratinjau_berkas(st.session_state.berkas_ktp, LABEL_DOKUMEN["ktp"])

    kol3, kol4 = st.columns(2)

    # --- 3. Lisensi menyelam SCUBA ---
    with kol3:
        st.markdown(
            f"**3. {LABEL_DOKUMEN['lisensi']}** · maks {BATAS_MB['lisensi']} MB"
        )
        st.caption("Berlaku untuk PADI, SSI, CMAS, TNI-AL, atau lisensi lain.")
        lisensi = st.file_uploader(
            "License Menyelam SCUBA (JPG/PNG/PDF)",
            type=["jpg", "jpeg", "png", "pdf"],
            key=f"{kunci_awalan}_unggah_lisensi",
            label_visibility="collapsed",
        )
        galat = validasi_ukuran(lisensi, "lisensi")
        if galat:
            st.error(galat)
        else:
            st.session_state.berkas_lisensi = lisensi
        pratinjau_berkas(
            st.session_state.berkas_lisensi, LABEL_DOKUMEN["lisensi"]
        )

    # --- 4. Sertifikat lainnya (opsional, banyak berkas) ---
    with kol4:
        st.markdown(
            f"**4. {LABEL_DOKUMEN['sertifikat']}** (opsional) · maks "
            f"{BATAS_MB['sertifikat']} MB per berkas"
        )
        st.caption(
            f"Contoh: sertifikat BNSP, pelatihan, medis selam (fit to dive). "
            f"Maksimal {MAKS_SERTIFIKAT} berkas."
        )
        sertifikat = st.file_uploader(
            "Sertifikat Lainnya (JPG/PNG/PDF)",
            type=["jpg", "jpeg", "png", "pdf"],
            accept_multiple_files=True,
            key=f"{kunci_awalan}_unggah_sertifikat",
            label_visibility="collapsed",
        )

        if sertifikat:
            if len(sertifikat) > MAKS_SERTIFIKAT:
                st.error(
                    f"Jumlah berkas melebihi batas. Maksimal "
                    f"{MAKS_SERTIFIKAT} berkas sekaligus."
                )
                sertifikat = sertifikat[:MAKS_SERTIFIKAT]

            lolos = []
            for berkas in sertifikat:
                galat = validasi_ukuran(berkas, "sertifikat")
                if galat:
                    st.error(galat)
                else:
                    lolos.append(berkas)
            st.session_state.berkas_sertifikat = lolos

        daftar = st.session_state.berkas_sertifikat
        if daftar:
            st.markdown(
                f"<div class='kartu-berkas'>{len(daftar)} berkas siap "
                f"dilampirkan</div>", unsafe_allow_html=True,
            )
            for berkas in daftar:
                pratinjau_berkas(berkas, "Sertifikat", wajib=False)
        else:
            pratinjau_berkas(None, "Sertifikat", wajib=False)


def susun_lampiran() -> list:
    """Menyusun daftar lampiran berurutan dari berkas yang telah diunggah.

    Urutan: KTP, License Menyelam SCUBA, kemudian sertifikat lainnya.
    """
    daftar = []

    if st.session_state.berkas_ktp:
        b = st.session_state.berkas_ktp
        daftar.append(Lampiran(
            judul="KTP", nama_berkas=b.name, data=b.getvalue(),
            adalah_pdf=adalah_pdf(b.name),
        ))

    if st.session_state.berkas_lisensi:
        b = st.session_state.berkas_lisensi
        daftar.append(Lampiran(
            judul="License Menyelam SCUBA", nama_berkas=b.name,
            data=b.getvalue(), adalah_pdf=adalah_pdf(b.name),
        ))

    for b in st.session_state.berkas_sertifikat:
        daftar.append(Lampiran(
            judul=f"Sertifikat — {b.name}", nama_berkas=b.name,
            data=b.getvalue(), adalah_pdf=adalah_pdf(b.name),
        ))

    return daftar


def bersihkan_tabel(kerangka: pd.DataFrame) -> list:
    """Membuang baris kosong dari tabel dan mengubahnya menjadi daftar dict."""
    if kerangka is None or kerangka.empty:
        return []
    bersih = kerangka.fillna("").astype(str)
    hasil = []
    for _, baris in bersih.iterrows():
        nilai = {k: str(v).strip() for k, v in baris.items()}
        if any(nilai.values()):
            hasil.append(nilai)
    return hasil


# =====================================================================
# HALAMAN 1 — GENERATOR CV
# =====================================================================

def halaman_cv():
    """Menampilkan halaman form dan pembuatan dokumen Curriculum Vitae."""
    st.markdown("## 📄 Generator Curriculum Vitae")

    jenis_cv = st.radio(
        "Jenis CV yang akan dibuat",
        [JENIS_TENAGA_AHLI, JENIS_PENYELAM],
        horizontal=True,
        key="cv_jenis",
        help=(
            "Tenaga Ahli menekankan pendidikan, publikasi, dan pengalaman "
            "kerja. Tenaga Spesialis Penyelaman menekankan lisensi selam, "
            "jam selam, dan pengalaman penyelaman."
        ),
    )
    penyelam = jenis_cv == JENIS_PENYELAM

    # Tempat menampilkan blok galat validasi di bagian atas form
    wadah_galat = st.container()

    # ---------------------------------------------------------------
    # 1. DATA PRIBADI
    # ---------------------------------------------------------------
    st.markdown("### 👤 Data Pribadi")
    kol1, kol2, kol3 = st.columns(3)

    with kol1:
        nama = st.text_input("Nama Lengkap (dengan gelar) *", key="cv_nama")
        tempat_lahir = st.text_input("Tempat Lahir *", key="cv_tempat_lahir")
        tanggal_lahir = st.date_input(
            "Tanggal Lahir *", value=date(1990, 1, 1),
            min_value=date(1940, 1, 1), max_value=date.today(),
            format="DD/MM/YYYY", key="cv_tanggal_lahir",
        )
        jenis_kelamin = st.selectbox(
            "Jenis Kelamin *", ["", "Laki-laki", "Perempuan"],
            key="cv_jenis_kelamin",
        )

    with kol2:
        agama = st.selectbox(
            "Agama *",
            ["", "Islam", "Kristen Protestan", "Katolik", "Hindu", "Buddha",
             "Khonghucu", "Lainnya"],
            key="cv_agama",
        )
        kewarganegaraan = st.text_input(
            "Kewarganegaraan *", value="Indonesia", key="cv_kewarganegaraan"
        )
        nomor_ktp = st.text_input(
            "Nomor KTP (NIK) *", max_chars=16, key="cv_nomor_ktp",
            help="16 digit angka, harus sama dengan nomor pada scan KTP.",
        )
        npwp = st.text_input("NPWP (opsional)", key="cv_npwp")

    with kol3:
        telepon = st.text_input("Nomor Telepon / HP *", key="cv_telepon")
        email = st.text_input("Email *", key="cv_email")
        alamat = st.text_area("Alamat Lengkap *", height=100, key="cv_alamat")

    with st.expander("Data kepegawaian dan peran dalam tim (opsional untuk non-ASN)"):
        kolp1, kolp2 = st.columns(2)
        with kolp1:
            nip = st.text_input("NIP", key="cv_nip")
            nidn = st.text_input("NIDN", key="cv_nidn")
            jabatan = st.text_input("Jabatan Fungsional", key="cv_jabatan")
        with kolp2:
            afiliasi = st.text_input(
                "Afiliasi Institusi",
                value="Fakultas Perikanan dan Ilmu Kelautan, "
                      "Universitas Sam Ratulangi",
                key="cv_afiliasi",
            )
            peran_tim = st.selectbox(
                "Kedudukan dalam Tim", [""] + PERAN_TIM_PILIHAN,
                key="cv_peran_tim",
            )
            peran_teknis = st.text_input("Peran Teknis", key="cv_peran_teknis")

    # ---------------------------------------------------------------
    # 2. UPLOAD DOKUMEN
    # ---------------------------------------------------------------
    bagian_unggah_dokumen("cv")

    # ---------------------------------------------------------------
    # 3. PENDIDIKAN
    # ---------------------------------------------------------------
    judul_pendidikan = (
        "🎓 Pendidikan Formal" if penyelam else "🎓 Riwayat Pendidikan"
    )
    st.markdown(f"### {judul_pendidikan}")
    st.caption("Klik baris terakhir untuk menambah data. Minimal 1 baris wajib diisi.")
    st.session_state.pendidikan = st.data_editor(
        st.session_state.pendidikan,
        num_rows="dynamic", width="stretch", key="ed_pendidikan",
        column_config={
            "Jenjang": st.column_config.SelectboxColumn(
                "Jenjang", options=["SMA/SMK", "D3", "D4", "S1", "S2", "S3"],
                width="small",
            ),
            "Tahun Lulus": st.column_config.TextColumn("Tahun Lulus", width="small"),
        },
    )

    # ---------------------------------------------------------------
    # 4. BAGIAN KHUSUS SESUAI JENIS CV
    # ---------------------------------------------------------------
    if penyelam:
        st.markdown("### 🤿 Sertifikat Selam Utama")
        kols1, kols2, kols3 = st.columns(3)
        with kols1:
            lisensi_jenis = st.selectbox(
                "Jenis Lisensi *",
                ["", "PADI", "SSI", "CMAS", "TNI-AL", "Lainnya"],
                key="cv_lisensi_jenis",
            )
            lisensi_nomor = st.text_input(
                "Nomor Sertifikat *", key="cv_lisensi_nomor"
            )
        with kols2:
            lisensi_level = st.selectbox(
                "Level Sertifikasi *",
                ["", "Open Water", "Advanced", "Rescue Diver", "Divemaster",
                 "Instructor", "Scientific Diver"],
                key="cv_lisensi_level",
            )
            lisensi_terbit = st.date_input(
                "Tanggal Terbit *", value=date(2020, 1, 1),
                min_value=date(1970, 1, 1), format="DD/MM/YYYY",
                key="cv_lisensi_terbit",
            )
        with kols3:
            seumur_hidup = st.checkbox(
                "Berlaku seumur hidup", value=True, key="cv_lisensi_seumur"
            )
            lisensi_berlaku = None
            if not seumur_hidup:
                lisensi_berlaku = st.date_input(
                    "Masa Berlaku", value=date(2030, 1, 1),
                    format="DD/MM/YYYY", key="cv_lisensi_berlaku",
                )

        st.markdown("### 🩺 Sertifikat Medis Selam (Fit to Dive)")
        kolm1, kolm2 = st.columns(2)
        with kolm1:
            medis_nomor = st.text_input("Nomor Sertifikat *", key="cv_medis_nomor")
            medis_penerbit = st.text_input(
                "Nama Dokter / Klinik Penerbit *", key="cv_medis_penerbit"
            )
        with kolm2:
            medis_terbit = st.date_input(
                "Tanggal Terbit *", value=date.today(),
                format="DD/MM/YYYY", key="cv_medis_terbit",
            )
            medis_berlaku = st.date_input(
                "Berlaku Sampai *", value=date(date.today().year + 1, 1, 1),
                format="DD/MM/YYYY", key="cv_medis_berlaku",
            )

        st.markdown("### 🌊 Kompetensi Penyelaman")
        total_jam_selam = st.number_input(
            "Total Jam Selam Terverifikasi (jam) *",
            min_value=0, max_value=100000, step=10, key="cv_total_jam",
        )
        st.markdown("**Keahlian Khusus Penyelaman** (pilih yang sesuai)")
        kolk = st.columns(2)
        keahlian_selam = []
        for i, keahlian in enumerate(KEAHLIAN_SELAM_PILIHAN):
            with kolk[i % 2]:
                if st.checkbox(keahlian, key=f"cv_keahlian_{i}"):
                    keahlian_selam.append(keahlian)

        st.markdown("### 🐠 Pengalaman Penyelaman")
        st.caption("Minimal 1 baris terisi. Tersedia 5 baris kosong.")
        st.session_state.pengalaman_selam = st.data_editor(
            st.session_state.pengalaman_selam,
            num_rows="dynamic", width="stretch",
            key="ed_pengalaman_selam",
            column_config={
                "Tahun": st.column_config.TextColumn("Tahun", width="small"),
                "Kedalaman Maks (m)": st.column_config.TextColumn(
                    "Kedalaman Maks (m)", width="small"
                ),
            },
        )

        bidang_keahlian = []
        publikasi_bersih = []

    else:
        st.markdown("### 🔬 Bidang Keahlian")
        bidang_pilih = st.multiselect(
            "Pilih bidang keahlian *", BIDANG_KEAHLIAN_PILIHAN,
            key="cv_bidang_pilih",
        )
        bidang_manual = st.text_input(
            "Bidang keahlian lainnya (pisahkan dengan koma)",
            key="cv_bidang_manual",
        )
        bidang_keahlian = list(bidang_pilih)
        if bidang_manual.strip():
            bidang_keahlian += [
                b.strip() for b in bidang_manual.split(",") if b.strip()
            ]

        lisensi_jenis = lisensi_nomor = lisensi_level = ""
        lisensi_terbit = lisensi_berlaku = None
        medis_nomor = medis_penerbit = ""
        medis_terbit = medis_berlaku = None
        total_jam_selam = 0
        keahlian_selam = []

    # ---------------------------------------------------------------
    # 5. SERTIFIKASI KOMPETENSI
    # ---------------------------------------------------------------
    judul_sertifikasi = (
        "📜 Sertifikasi Kompetensi Lainnya" if penyelam
        else "📜 Sertifikasi Kompetensi"
    )
    st.markdown(f"### {judul_sertifikasi}")
    st.session_state.sertifikasi = st.data_editor(
        st.session_state.sertifikasi,
        num_rows="dynamic", width="stretch", key="ed_sertifikasi",
        column_config={
            "Tahun": st.column_config.TextColumn("Tahun", width="small"),
        },
    )

    # ---------------------------------------------------------------
    # 6. PENGALAMAN KERJA DAN PUBLIKASI (khusus Tenaga Ahli)
    # ---------------------------------------------------------------
    if not penyelam:
        st.markdown("### 💼 Pengalaman Kerja")
        st.caption("Minimal 1 baris terisi. Tersedia 5 baris kosong.")
        st.session_state.pengalaman_kerja = st.data_editor(
            st.session_state.pengalaman_kerja,
            num_rows="dynamic", width="stretch",
            key="ed_pengalaman_kerja",
            column_config={
                "Tahun": st.column_config.TextColumn("Tahun", width="small"),
            },
        )

        st.markdown("### 📚 Publikasi Ilmiah (opsional)")
        st.session_state.publikasi = st.data_editor(
            st.session_state.publikasi,
            num_rows="dynamic", width="stretch", key="ed_publikasi",
            column_config={
                "Tahun": st.column_config.TextColumn("Tahun", width="small"),
            },
        )
        publikasi_bersih = bersihkan_tabel(st.session_state.publikasi)

    # ---------------------------------------------------------------
    # 6b. RINGKASAN KETERKAITAN DENGAN KEGIATAN
    # ---------------------------------------------------------------
    st.markdown("### 🔗 Ringkasan Keterkaitan dengan Kegiatan")
    st.caption(
        "Ringkasan berikut menjelaskan keterkaitan riwayat pengalaman, "
        "penelitian, dan keahlian di atas dengan kegiatan restorasi terumbu "
        "karang. Disusun otomatis dari data yang sudah diisi, dan dapat "
        "disunting atau ditambahkan secara manual sebelum dokumen dibuat."
    )

    pengalaman_pratinjau = bersihkan_tabel(
        st.session_state.pengalaman_selam if penyelam
        else st.session_state.pengalaman_kerja
    )
    publikasi_pratinjau = (
        [] if penyelam else bersihkan_tabel(st.session_state.publikasi)
    )

    if st.button("🔄 Buat / Perbarui Draf Otomatis", key="cv_tombol_draf_afiliasi"):
        st.session_state.cv_ringkasan_afiliasi = utils.buat_draf_afiliasi(
            nama=nama, jenis_cv=jenis_cv, bidang_keahlian=bidang_keahlian,
            pengalaman=pengalaman_pratinjau, publikasi=publikasi_pratinjau,
            keahlian_selam=keahlian_selam,
            total_jam_selam=int(total_jam_selam) if penyelam else 0,
        )

    ringkasan_afiliasi = st.text_area(
        "Ringkasan Keterkaitan dengan Kegiatan",
        key="cv_ringkasan_afiliasi", height=140,
        label_visibility="collapsed",
        placeholder=(
            "Klik tombol di atas untuk membuat draf otomatis, lalu sunting "
            "atau tambahkan sesuai kebutuhan. Boleh juga ditulis manual "
            "sepenuhnya tanpa memakai draf otomatis."
        ),
    )

    # ---------------------------------------------------------------
    # 7. PERNYATAAN DAN TANDA TANGAN
    # ---------------------------------------------------------------
    st.markdown("### ✍️ Pernyataan dan Tanda Tangan")
    st.info(
        "Demikian daftar riwayat hidup ini saya buat dengan sebenar-benarnya "
        "untuk dapat dipergunakan sebagaimana mestinya."
    )
    kolt1, kolt2, kolt3 = st.columns(3)
    with kolt1:
        tempat_ttd = st.text_input("Tempat", value="Bitung", key="cv_tempat_ttd")
    with kolt2:
        tanggal_ttd = st.date_input(
            "Tanggal", value=date.today(), format="DD/MM/YYYY",
            key="cv_tanggal_ttd",
        )
    with kolt3:
        nama_terang = st.text_input(
            "Nama Terang", key="cv_nama_terang",
            help="Kosongkan untuk memakai nama lengkap di atas.",
        )

    # ---------------------------------------------------------------
    # 8. TOMBOL GENERATE DAN VALIDASI
    # ---------------------------------------------------------------
    st.markdown("---")
    tekan_generate = st.button(
        "⚡ Generate CV", type="primary", width="stretch",
        key="tombol_generate_cv",
    )

    if tekan_generate:
        pendidikan_bersih = bersihkan_tabel(st.session_state.pendidikan)
        sertifikasi_bersih = bersihkan_tabel(st.session_state.sertifikasi)
        kerja_bersih = bersihkan_tabel(st.session_state.pengalaman_kerja)
        selam_bersih = bersihkan_tabel(st.session_state.pengalaman_selam)
        publikasi_bersih = bersihkan_tabel(st.session_state.publikasi)

        galat = validasi_form_cv(
            penyelam=penyelam, nama=nama, tempat_lahir=tempat_lahir,
            jenis_kelamin=jenis_kelamin, agama=agama,
            kewarganegaraan=kewarganegaraan, alamat=alamat, telepon=telepon,
            email=email, nomor_ktp=nomor_ktp,
            pendidikan=pendidikan_bersih, bidang_keahlian=bidang_keahlian,
            pengalaman_kerja=kerja_bersih, pengalaman_selam=selam_bersih,
            lisensi_jenis=lisensi_jenis, lisensi_nomor=lisensi_nomor,
            lisensi_level=lisensi_level, medis_nomor=medis_nomor,
            medis_penerbit=medis_penerbit, keahlian_selam=keahlian_selam,
        )

        if galat:
            with wadah_galat:
                pesan = "\n".join(f"- {g}" for g in galat)
                st.error(
                    f"**Dokumen belum dapat dibuat. Lengkapi "
                    f"{len(galat)} hal berikut:**\n\n{pesan}"
                )
            st.session_state.hasil_cv_pdf = None
            st.session_state.hasil_cv_docx = None
        else:
            data = CVData(
                jenis_cv=jenis_cv,
                nama=nama, tempat_lahir=tempat_lahir,
                tanggal_lahir=tanggal_lahir, jenis_kelamin=jenis_kelamin,
                agama=agama, kewarganegaraan=kewarganegaraan, alamat=alamat,
                telepon=telepon, email=email, nomor_ktp=nomor_ktp, npwp=npwp,
                nip=nip, nidn=nidn, jabatan=jabatan, afiliasi=afiliasi,
                peran_tim=peran_tim, peran_teknis=peran_teknis,
                pendidikan=pendidikan_bersih,
                sertifikasi=sertifikasi_bersih,
                pengalaman_kerja=kerja_bersih,
                publikasi=publikasi_bersih,
                bidang_keahlian=bidang_keahlian,
                lisensi_jenis=lisensi_jenis, lisensi_nomor=lisensi_nomor,
                lisensi_level=lisensi_level, lisensi_terbit=lisensi_terbit,
                lisensi_berlaku=lisensi_berlaku,
                medis_nomor=medis_nomor, medis_penerbit=medis_penerbit,
                medis_terbit=medis_terbit, medis_berlaku=medis_berlaku,
                total_jam_selam=int(total_jam_selam),
                keahlian_selam=keahlian_selam,
                pengalaman_selam=selam_bersih,
                tempat_ttd=tempat_ttd, tanggal_ttd=tanggal_ttd,
                nama_terang=nama_terang,
                ringkasan_afiliasi=ringkasan_afiliasi,
                foto=(st.session_state.berkas_foto.getvalue()
                      if st.session_state.berkas_foto else None),
                lampiran=susun_lampiran(),
            )

            with st.spinner("Menyusun dokumen CV beserta lampiran..."):
                try:
                    st.session_state.hasil_cv_pdf = generate_cv_pdf(data)
                    st.session_state.hasil_cv_docx = generate_cv_docx(data)
                    st.session_state.data_cv_terakhir = data
                except Exception as e:
                    st.session_state.hasil_cv_pdf = None
                    st.session_state.hasil_cv_docx = None
                    st.error(f"Terjadi kesalahan saat menyusun dokumen: {e}")

    # ---------------------------------------------------------------
    # 9. PRATINJAU DAN UNDUHAN
    # ---------------------------------------------------------------
    if st.session_state.hasil_cv_pdf:
        st.success("Dokumen CV berhasil disusun.")
        data = st.session_state.data_cv_terakhir
        tampilkan_pratinjau_cv(data)

        nama_unduh = nama_berkas_aman(data.nama, "CV")
        kolu1, kolu2 = st.columns(2)
        with kolu1:
            st.download_button(
                "⬇️ Download CV (PDF)",
                data=st.session_state.hasil_cv_pdf,
                file_name=f"{nama_unduh}.pdf",
                mime="application/pdf",
                width="stretch",
            )
        with kolu2:
            st.download_button(
                "⬇️ Download CV (Word)",
                data=st.session_state.hasil_cv_docx,
                file_name=f"{nama_unduh}.docx",
                mime=("application/vnd.openxmlformats-officedocument."
                      "wordprocessingml.document"),
                width="stretch",
            )


def validasi_form_cv(**k) -> list:
    """Memeriksa kelengkapan seluruh isian wajib pada form CV.

    Mengembalikan daftar pesan galat. Daftar kosong berarti form lolos
    validasi dan dokumen boleh dibuat.
    """
    galat = []

    # --- Data pribadi ---
    if not k["nama"].strip():
        galat.append("Nama Lengkap belum diisi.")
    if not k["tempat_lahir"].strip():
        galat.append("Tempat Lahir belum diisi.")
    if not k["jenis_kelamin"]:
        galat.append("Jenis Kelamin belum dipilih.")
    if not k["agama"]:
        galat.append("Agama belum dipilih.")
    if not k["kewarganegaraan"].strip():
        galat.append("Kewarganegaraan belum diisi.")
    if not k["alamat"].strip():
        galat.append("Alamat Lengkap belum diisi.")
    if not k["telepon"].strip():
        galat.append("Nomor Telepon belum diisi.")

    email = k["email"].strip()
    if not email:
        galat.append("Email belum diisi.")
    elif "@" not in email or "." not in email.split("@")[-1]:
        galat.append("Format Email tidak valid.")

    nik = k["nomor_ktp"].strip()
    if not nik:
        galat.append("Nomor KTP (NIK) belum diisi.")
    elif not (nik.isdigit() and len(nik) == 16):
        galat.append("Nomor KTP (NIK) harus berupa 16 digit angka.")

    # --- Dokumen unggahan wajib ---
    if not st.session_state.berkas_foto:
        galat.append("Pas Foto 3x4 belum diunggah.")
    if not st.session_state.berkas_ktp:
        galat.append("Scan KTP belum diunggah.")
    if not st.session_state.berkas_lisensi:
        galat.append("License Menyelam SCUBA belum diunggah.")

    # --- Pendidikan ---
    if not k["pendidikan"]:
        galat.append("Riwayat Pendidikan minimal 1 baris harus diisi.")

    # --- Isian khusus per jenis CV ---
    if k["penyelam"]:
        if not k["lisensi_jenis"]:
            galat.append("Jenis Lisensi selam belum dipilih.")
        if not k["lisensi_nomor"].strip():
            galat.append("Nomor Sertifikat selam belum diisi.")
        if not k["lisensi_level"]:
            galat.append("Level Sertifikasi selam belum dipilih.")
        if not k["medis_nomor"].strip():
            galat.append("Nomor Sertifikat Medis Selam belum diisi.")
        if not k["medis_penerbit"].strip():
            galat.append("Dokter/Klinik penerbit sertifikat medis belum diisi.")
        if not k["keahlian_selam"]:
            galat.append("Keahlian Khusus Penyelaman minimal pilih satu.")
        if not k["pengalaman_selam"]:
            galat.append("Pengalaman Penyelaman minimal 1 baris harus diisi.")
    else:
        if not k["bidang_keahlian"]:
            galat.append("Bidang Keahlian belum dipilih atau diisi.")
        if not k["pengalaman_kerja"]:
            galat.append("Pengalaman Kerja minimal 1 baris harus diisi.")

    return galat


def tampilkan_pratinjau_cv(data: CVData):
    """Menampilkan pratinjau ringkas isi dokumen sebelum diunduh."""
    with st.expander("🔍 Pratinjau isi dokumen", expanded=True):
        kol1, kol2 = st.columns([1, 3])

        with kol1:
            if data.foto:
                st.image(utils.potong_pas_foto(data.foto), width=110)

        with kol2:
            st.markdown(f"**{data.nama}**")
            st.caption(f"{data.jenis_cv} · {data.ttl()}")
            if data.peran_tim:
                st.caption(f"Peran: {data.peran_tim}")

        ringkas = {
            "Nomor KTP (NIK)": data.nomor_ktp,
            "Alamat": data.alamat,
            "Telepon": data.telepon,
            "Email": data.email,
        }
        st.table(pd.DataFrame(
            [{"Keterangan": k, "Isi": v} for k, v in ringkas.items()]
        ).set_index("Keterangan"))

        kolr1, kolr2, kolr3 = st.columns(3)
        kolr1.metric("Baris Pendidikan", len(data.pendidikan))
        if data.adalah_penyelam():
            kolr2.metric("Pengalaman Selam", len(data.pengalaman_selam))
            kolr3.metric("Total Jam Selam", f"{data.total_jam_selam} jam")
        else:
            kolr2.metric("Pengalaman Kerja", len(data.pengalaman_kerja))
            kolr3.metric("Publikasi", len(data.publikasi))

        st.markdown("**Halaman lampiran yang akan disisipkan:**")
        if data.lampiran:
            for i, lam in enumerate(data.lampiran, start=1):
                st.markdown(f"- Lampiran {i} — {lam.judul} (`{lam.nama_berkas}`)")
        else:
            st.caption("Tidak ada lampiran.")

        if data.ringkasan_afiliasi.strip():
            st.markdown("**Ringkasan Keterkaitan dengan Kegiatan:**")
            st.caption(data.ringkasan_afiliasi)


# =====================================================================
# HALAMAN 2 — GENERATOR SURAT TUGAS DINAS LUAR
# =====================================================================

def halaman_surat_tugas():
    """Menampilkan halaman form dan pembuatan Surat Tugas Dinas Luar."""
    st.markdown("## ✉️ Generator Surat Tugas Dinas Luar")
    st.info(
        "Surat otomatis menyertakan klausul pembebasan tugas mengajar dan "
        "larangan rangkap bayar sesuai Diktum KETUJUH SK Dekan. Surat dapat "
        "diterbitkan untuk satu orang maupun satu tim sekaligus."
    )

    wadah_galat = st.container()

    st.markdown("### 📋 Identitas Surat")
    kol1, kol2 = st.columns(2)
    with kol1:
        nomor_surat = st.text_input(
            "Nomor Surat *", value="800/     /UN12.6/TU.00.00/2026",
            key="st_nomor",
            help="Contoh format: 800/1234/UN12.6/TU.00.00/2026",
        )
    with kol2:
        unit_kerja = st.text_input(
            "Unit Kerja",
            value="Fakultas Perikanan dan Ilmu Kelautan, "
                  "Universitas Sam Ratulangi",
            key="st_unit_kerja",
        )

    st.markdown("### 👥 Personil yang Ditugaskan")
    st.caption(
        "Isi satu baris untuk penugasan perorangan, atau beberapa baris "
        "untuk penugasan tim."
    )
    st.session_state.personil_st = st.data_editor(
        st.session_state.personil_st,
        num_rows="dynamic", width="stretch", key="ed_personil_st",
        column_config={
            "Peran dalam Tim": st.column_config.SelectboxColumn(
                "Peran dalam Tim", options=PERAN_TIM_PILIHAN,
            ),
        },
    )

    st.markdown("### 🗺️ Rincian Penugasan")
    kolr1, kolr2 = st.columns(2)
    with kolr1:
        jenis_tugas = st.text_input(
            "Jenis Tugas *", value="Survei Awal Ekologi dan Pemetaan Dasar",
            key="st_jenis_tugas",
        )
    with kolr2:
        lokasi = st.text_input(
            "Lokasi Tugas *",
            value="Pelabuhan Perikanan Samudera (PPS) Bitung, Sulawesi Utara",
            key="st_lokasi",
        )

    kolw1, kolw2, kolw3 = st.columns(3)
    with kolw1:
        tanggal_mulai = st.date_input(
            "Tanggal Mulai *", value=date(2026, 9, 15),
            format="DD/MM/YYYY", key="st_tanggal_mulai",
        )
    with kolw2:
        tanggal_selesai = st.date_input(
            "Tanggal Selesai *", value=date(2026, 9, 16),
            format="DD/MM/YYYY", key="st_tanggal_selesai",
        )
    with kolw3:
        jumlah_hari = st.number_input(
            "Jumlah Hari Kerja *", min_value=1, max_value=60, value=2,
            key="st_jumlah_hari",
        )

    st.markdown("### ⚙️ Klausul dan Penanda Tangan")
    kolk1, kolk2, kolk3 = st.columns(3)
    with kolk1:
        klaim_8_oj = st.checkbox(
            "Aktifkan klausul 8 OJ/hari",
            key="st_klaim_8oj",
            help="Hanya berlaku bila terdapat personil berperan Pembantu Peneliti.",
        )
    with kolk2:
        tempat_ttd = st.text_input("Tempat Tanda Tangan", value="Manado",
                                   key="st_tempat_ttd")
    with kolk3:
        tanggal_ttd = st.date_input(
            "Tanggal Surat", value=date.today(), format="DD/MM/YYYY",
            key="st_tanggal_ttd",
        )

    st.markdown("---")
    tekan = st.button(
        "✉️ Generate Surat Tugas", type="primary", width="stretch",
        key="tombol_generate_st",
    )

    if tekan:
        personil = bersihkan_tabel(st.session_state.personil_st)
        galat = []

        if not nomor_surat.strip() or "     " in nomor_surat:
            galat.append("Nomor Surat belum diisi lengkap.")
        if not personil:
            galat.append("Daftar personil belum diisi (minimal 1 orang).")
        else:
            for i, orang in enumerate(personil, start=1):
                if not orang.get("Nama", "").strip():
                    galat.append(f"Nama personil pada baris {i} belum diisi.")
                if not orang.get("Peran dalam Tim", "").strip():
                    galat.append(f"Peran personil pada baris {i} belum dipilih.")
        if not jenis_tugas.strip():
            galat.append("Jenis Tugas belum diisi.")
        if not lokasi.strip():
            galat.append("Lokasi Tugas belum diisi.")
        if tanggal_selesai < tanggal_mulai:
            galat.append("Tanggal Selesai tidak boleh lebih awal dari Tanggal Mulai.")

        if galat:
            with wadah_galat:
                pesan = "\n".join(f"- {g}" for g in galat)
                st.error(
                    f"**Surat belum dapat dibuat. Lengkapi "
                    f"{len(galat)} hal berikut:**\n\n{pesan}"
                )
            st.session_state.hasil_st_pdf = None
            st.session_state.hasil_st_docx = None
        else:
            data_st = DataSuratTugas(
                nomor_surat=nomor_surat, personil=personil,
                unit_kerja=unit_kerja, jenis_tugas=jenis_tugas, lokasi=lokasi,
                tanggal_mulai=tanggal_mulai, tanggal_selesai=tanggal_selesai,
                jumlah_hari=int(jumlah_hari), klaim_8_oj=klaim_8_oj,
                tempat_ttd=tempat_ttd, tanggal_ttd=tanggal_ttd,
            )
            with st.spinner("Menyusun Surat Tugas..."):
                try:
                    st.session_state.hasil_st_pdf = generate_st_dinas_luar_pdf(data_st)
                    st.session_state.hasil_st_docx = generate_st_dinas_luar(data_st)
                except Exception as e:
                    st.session_state.hasil_st_pdf = None
                    st.session_state.hasil_st_docx = None
                    st.error(f"Terjadi kesalahan saat menyusun surat: {e}")

    if st.session_state.hasil_st_pdf:
        st.success("Surat Tugas berhasil disusun.")
        personil = bersihkan_tabel(st.session_state.personil_st)
        nama_utama = personil[0].get("Nama", "Tim") if personil else "Tim"
        akhiran = "Tim" if len(personil) > 1 else nama_utama
        nama_unduh = nama_berkas_aman(akhiran, "ST_Dinas_Luar")

        kolu1, kolu2 = st.columns(2)
        with kolu1:
            st.download_button(
                "⬇️ Download Surat Tugas (PDF)",
                data=st.session_state.hasil_st_pdf,
                file_name=f"{nama_unduh}.pdf",
                mime="application/pdf",
                width="stretch",
            )
        with kolu2:
            st.download_button(
                "⬇️ Download Surat Tugas (Word)",
                data=st.session_state.hasil_st_docx,
                file_name=f"{nama_unduh}.docx",
                mime=("application/vnd.openxmlformats-officedocument."
                      "wordprocessingml.document"),
                width="stretch",
            )


# =====================================================================
# HALAMAN 3 — PANDUAN PENGGUNAAN
# =====================================================================

def halaman_panduan():
    """Menampilkan panduan singkat penggunaan aplikasi."""
    st.markdown("## ℹ️ Panduan Penggunaan")

    st.markdown("### Dokumen yang wajib disiapkan")
    st.table(pd.DataFrame([
        {"Dokumen": "Pas Foto 3x4", "Format": "JPG, PNG",
         "Maks": "2 MB", "Wajib": "Ya"},
        {"Dokumen": "KTP", "Format": "JPG, PNG, PDF",
         "Maks": "2 MB", "Wajib": "Ya"},
        {"Dokumen": "License Menyelam SCUBA", "Format": "JPG, PNG, PDF",
         "Maks": "5 MB", "Wajib": "Ya"},
        {"Dokumen": "Sertifikat Lainnya", "Format": "JPG, PNG, PDF",
         "Maks": "5 MB per berkas (maks 10 berkas)", "Wajib": "Tidak"},
    ]).set_index("Dokumen"))

    st.markdown("### Langkah pembuatan CV")
    st.markdown(
        "1. Pilih jenis CV: **Tenaga Ahli** atau **Tenaga Spesialis Penyelaman**.\n"
        "2. Isi seluruh data pribadi bertanda bintang (*).\n"
        "3. Unggah keempat dokumen persyaratan pada bagian Upload Dokumen.\n"
        "4. Lengkapi tabel pendidikan, sertifikasi, dan pengalaman.\n"
        "5. Periksa blok pernyataan dan tanda tangan.\n"
        "6. Tekan **Generate CV**, periksa pratinjau, lalu unduh PDF atau Word."
    )

    st.markdown("### Catatan penting")
    st.warning(
        "Nomor KTP yang diisi pada form harus sama persis dengan nomor yang "
        "tertera pada scan KTP yang diunggah. Ketidaksesuaian akan menyebabkan "
        "dokumen ditolak pada tahap verifikasi administrasi."
    )
    st.info(
        "Lampiran berformat PDF hanya dapat ditampilkan utuh pada berkas "
        "keluaran PDF. Pada berkas Word, lampiran PDF hanya dicatat sebagai "
        "keterangan rujukan, sedangkan lampiran gambar tetap disisipkan."
    )

    st.markdown("### Spesifikasi dokumen keluaran")
    st.markdown(
        "- Ukuran kertas A4 potret\n"
        "- Margin: atas 2,5 cm · bawah 2,5 cm · kiri 3 cm · kanan 2 cm\n"
        "- Huruf Times New Roman 12 pt untuk isi, 14 pt tebal untuk judul\n"
        "- Lampiran disisipkan sebagai halaman terpisah dengan judul jelas"
    )


# =====================================================================
# NAVIGASI UTAMA
# =====================================================================

def main():
    """Menjalankan aplikasi: menampilkan kepala, sidebar, dan halaman aktif."""
    if not periksa_kata_sandi():
        st.stop()

    tampilkan_kepala()

    with st.sidebar:
        if logo_tersedia():
            st.image(LOGO_PATH, width=90)
        st.markdown("### Menu Navigasi")
        halaman = st.radio(
            "Pilih halaman",
            ["Generator CV", "Generator Surat Tugas", "Panduan Penggunaan"],
            label_visibility="collapsed",
            key="menu_navigasi",
        )

        st.markdown("---")
        st.markdown("### Status Dokumen")
        status = [
            ("Pas Foto 3x4", st.session_state.berkas_foto is not None),
            ("KTP", st.session_state.berkas_ktp is not None),
            ("License Selam", st.session_state.berkas_lisensi is not None),
            ("Sertifikat Lainnya",
             len(st.session_state.berkas_sertifikat) > 0),
        ]
        for label, ada in status:
            st.markdown(f"{'✅' if ada else '⬜'} {label}")

        st.markdown("---")
        if st.button("🔄 Reset Seluruh Form", width="stretch"):
            for kunci in list(st.session_state.keys()):
                del st.session_state[kunci]
            st.rerun()

        if hasattr(st, "secrets") and st.secrets.get("APP_PASSWORD", None):
            st.markdown("---")
            if st.button("🔒 Keluar", width="stretch"):
                st.session_state.sudah_login = False
                st.rerun()

        st.markdown("---")
        st.caption(
            "Generator Dokumen Tim Pelaksana Swakelola\n\n"
            "FPIK UNSRAT — Tahun Anggaran 2026"
        )

    if halaman == "Generator CV":
        halaman_cv()
    elif halaman == "Generator Surat Tugas":
        halaman_surat_tugas()
    else:
        halaman_panduan()

    st.markdown("---")
    st.caption(
        "Generator Dokumen Tim Pelaksana Swakelola — "
        "Fakultas Perikanan dan Ilmu Kelautan, Universitas Sam Ratulangi, 2026"
    )


if __name__ == "__main__":
    main()
