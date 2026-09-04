"""
app.py — Web Generator Terintegrasi: CV & Surat Tugas Dinas Luar
Versi: 3.0 — Data CV Dr. Ari Berty Rondonuwu, M.Sc., M.Si. sudah di-pre-fill lengkap
FPIK UNSRAT — Kegiatan Restorasi Terumbu Karang PPS Bitung TA 2026
"""

import streamlit as st
import base64
from datetime import date
from cv_builder import CVData, generate_cv_docx
from st_dinas_luar_builder import DataSuratTugas, generate_st_dinas_luar

# ===================== KONFIGURASI HALAMAN =====================
st.set_page_config(
    page_title="Generator Tim Pelaksana PPS Bitung",
    page_icon="🌊",
    layout="wide"
)

# ===================== CUSTOM CSS =====================
st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg, #1B3F6B 0%, #2D7D46 100%);
    color: white;
    padding: 1.2rem 1.5rem;
    border-radius: 10px;
    margin-bottom: 1.2rem;
    text-align: center;
}
.main-header h2 { margin: 0; font-size: 1.1rem; font-weight: 700; letter-spacing: 0.4px; }
.main-header p  { margin: 0.3rem 0 0 0; font-size: 0.82rem; opacity: 0.88; }
.section-hdr {
    background: #E8F4F8;
    border-left: 5px solid #1B3F6B;
    padding: 0.45rem 1rem;
    border-radius: 0 6px 6px 0;
    margin: 1.1rem 0 0.4rem 0;
    font-weight: 700;
    color: #1B3F6B;
    font-size: 0.97rem;
}
.info-karang {
    background: #F0F9F4;
    border-left: 4px solid #2D7D46;
    padding: 0.55rem 1rem;
    border-radius: 0 6px 6px 0;
    margin-bottom: 0.8rem;
    font-size: 0.87rem;
    color: #1a5230;
}
.up-label  { font-weight: 600; color: #1B3F6B; font-size: 0.88rem; margin-bottom: 2px; }
.up-req    { color: #cc0000; font-size: 0.78rem; }
.up-opt    { color: #2D7D46; font-size: 0.78rem; }
.err-block {
    background: #fff0f0; border: 2px solid #cc0000;
    border-radius: 8px; padding: 0.9rem 1.1rem; margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ===================== HEADER =====================
st.markdown("""
<div class="main-header">
    <h2>🌊 GENERATOR DOKUMEN TIM PELAKSANA SWAKELOLA</h2>
    <p>Relokasi &amp; Replanting / Restorasi Terumbu Karang &nbsp;|&nbsp; PPS Kota Bitung &nbsp;|&nbsp; Tahun Anggaran 2026</p>
    <p style="font-size:0.76rem;opacity:0.75;">Fakultas Perikanan dan Ilmu Kelautan (FPIK) — Universitas Sam Ratulangi (UNSRAT)</p>
</div>
""", unsafe_allow_html=True)

# ===================== DATA DEFAULT CV DR. ARI BERTY RONDONUWU =====================
# Diisi otomatis dari CV lengkap 2025
DEFAULT = {
    "nama"         : "Dr. Ir. Ari Berty Rondonuwu, M.Sc., M.Si.",
    "nip"          : "196801291993031001",
    "nidn"         : "0029016804",
    "ttl"          : "Tareran, 29 Januari 1968",
    "jabatan"      : "Lektor Kepala — Penata / III-d",
    "afiliasi"     : "Fakultas Perikanan dan Ilmu Kelautan (FPIK), Universitas Sam Ratulangi (UNSRAT), Manado",
    "alamat"       : "Perumahan Duta Graha Blok B No. 14A, Kel. Malalayang II Lingkungan 8, Manado 95136",
    "telepon"      : "081356033368",
    "email"        : "arirondonuwu@unsrat.ac.id",
    "peran_tim"    : "Ketua Tim Pelaksana",
    "peran_teknis" : "Ahli Rehabilitasi Terumbu Karang (Team Leader)",

    # Format: Jenjang | Institusi | Bidang Studi | Tahun
    "pendidikan": (
        "S-1 / Ir. | Universitas Sam Ratulangi, Manado | Manajemen Sumberdaya Perairan | 1986–1991\n"
        "S-2 / M.Sc. | Università Politecnica delle Marche, Ancona, Italia | Biodiversity & Management of Coral Reef | 2003–2004\n"
        "S-2 / M.Si. | Universitas Sam Ratulangi, Manado | Ilmu Perairan | 2015–2017\n"
        "S-3 / Dr. | Universitas Sam Ratulangi, Manado | Ilmu Kelautan | 2017–2020"
    ),

    # Format: Lokasi | Metode | Periode | Mitra/Pemberi Kerja
    "restorasi": (
        "Pantai Malalayang, Manado | Transplantasi Karang & CSC | 2009–sekarang | PT. TJ Silfanus\n"
        "Taman Nasional Bunaken | Terumbu Buatan + Transplantasi | 2018–2020 | Balai TN Bunaken\n"
        "Pulau Lembeh, Bitung | Terumbu Buatan | 2020–2022 | DKP Kota Bitung\n"
        "Poopoh, Minahasa | Transplantasi Partisipatif | 2015–2020 | Masyarakat Lokal\n"
        "Desa Bahoi, Likupang Barat | Terumbu Buatan | 2012 | CCDP-IFAD\n"
        "Pulau Bunaken, Kec. Bunaken Kepulauan | Restorasi Terumbu Karang | 2016 | DKP Sulut – APBD\n"
        "Kab. Minahasa Selatan (Amurang Barat & Tatapaan) | Transplantasi Terumbu Karang | 2015 | APBD Kab. Minahasa Selatan\n"
        "PPS Kota Bitung | Restorasi & Transplantasi Terumbu Karang | 2014 | DKP Kota Bitung\n"
        "Underwater Coral Plantation, Selat Lembeh | Transplantasi CSR | 2016 | CSR Pelindo IV, Indofood CBP, DKP Bitung\n"
        "Kelurahan Kareko, Bitung | Rehabilitasi Terumbu Karang Aqua Reef | 2014 | CCDP-IFAD\n"
        "Manado Tua & Bahoi; Malalayang Dua | Restorasi Terumbu Karang | 2014–2015 | APBD/APBN Sulut"
    ),

    # Format: Tahun | Judul Penelitian | Sumber Dana
    "penelitian": (
        "2022 | Pemetaan Habitat Perairan Dangkal dengan UAV di Likupang | LPPM Unsrat\n"
        "2021 | Pemetaan Ekosistem Terumbu Karang Pulau Serena dengan UAV | LPPM Unsrat\n"
        "2019–2020 | Optimalisasi Artificial Reef dan Restorasi Terumbu Karang di Poopoh, Minahasa | DRPM DIKTI\n"
        "2019–2020 | Monitoring Kesehatan Terumbu Karang di Kab. Raja Ampat | COREMAP-CTI LIPI\n"
        "2019 | Inventarisasi Sumberdaya Ikan Karang di Kab. Raja Ampat | LIPI\n"
        "2018 | Monitoring Status Ekosistem Terumbu Karang & Sumberdaya Ikan TN Bunaken | Balai TN Bunaken\n"
        "2017 | DNA Barcoding Ikan Karang Endemik Sulawesi Utara (COI Sequencing) | Mandiri/PNBP Unsrat\n"
        "2016 | Monitoring Terumbu Karang & Ekosistem Terkait di Salawati & Batanta, Raja Ampat | P2O LIPI Jakarta\n"
        "2015–2016 | Ecological Assessment Reef Health Index di Kepulauan Sangihe | COREMAP World Bank\n"
        "2014 | Survei Baseline Terumbu Karang & Ikan Karang Kepulauan Sangihe | COREMAP World Bank\n"
        "2013 | Monitoring Terumbu Karang Taman Nasional Bunaken | ADB – BKSDA Sulut"
    ),

    # Format: Tahun | Judul | Jurnal/Penerbit | Indeks
    "publikasi": (
        "2021 | Shallow water habitat mapping using UAV in Serena Island, North Sulawesi | AACL Bioflux 14(6) | Scopus Q3\n"
        "2020 | Mitochondrial CO1 sequences of Banggai Cardinalfish (Pterapogon kauderni) | AACL Bioflux 13(2) | Scopus Q3\n"
        "2019 | Coral reef health assessment at Salawati & Batanta, Raja Ampat | Ecology, Environment and Conservation 25(2) | Scopus\n"
        "2023 | Biometrik Otolit Ikan Kardinal Banggai (Pterapogon kauderni) | Jurnal Ilmiah PLATAX 11(1) | SINTA 4\n"
        "2022 | DNA Barcoding Ikan Kerapu (Epinephelus spp.) Perairan Sulawesi Utara | Jurnal PLATAX 10(1) | SINTA 4\n"
        "2020 | Distribusi Ikan Karang di Ekosistem Terumbu Karang Pulau Serena, Minahasa Utara | Jurnal Ilmiah PLATAX 8(1) | SINTA 4\n"
        "2019 | Kesehatan Terumbu Karang & Ekosistem Terkait di Pulau Salawati & Batanta, Raja Ampat (Buku) | Unsrat Press | ISBN 978-623-6818077\n"
        "2016 | Ekologi Perairan Teluk Manado (Buku Referensi) | FPIK Unsrat Press | ISBN 978-602-0847054\n"
        "2016 | Monitoring Terumbu Karang & Ekosistem Terkait, Salawati & Batanta 2016 (Buku) | P2O LIPI Jakarta | ISBN 978-602-9445947"
    ),

    # Format: Tahun | Kegiatan | Lokasi/Penyelenggara
    "pengabdian": (
        "2020 | Narasumber PEN Restorasi Terumbu Karang – ICRG (Pemberdayaan Masyarakat Bali Terdampak COVID-19) | Bali – Kemko Maritim & Investasi\n"
        "2021 | Transplantasi Karang di Pantai Malalayang Depan Minanga Divers | Manado – FPIK Unsrat\n"
        "2022 | Penggunaan Batok Kelapa untuk Mengurangi Pencemaran Plastik di Laut | Kec. Sario, Manado\n"
        "2021 | Sosialisasi Biota Laut yang Dilindungi | Sulawesi Utara – FPIK Unsrat\n"
        "2020 | Penanaman Mangrove di Perairan Pantai Depan Lab Basah FPIK Likupang Timur | Likupang Timur\n"
        "2016 | Modul Pengenalan Ekosistem Terumbu Karang dan Metode Pemantauannya | Bitung – CCDP-IFAD\n"
        "2015 | Fasilitasi Ekowisata Bahari, Kel. Pasirpanjang, Kec. Lembeh Selatan | Bitung – CCDP-IFAD\n"
        "2014 | Pembentukan Daerah Perlindungan Laut (DPL) Berbasis Masyarakat | 6 Kelurahan Pulau Lembeh, Bitung\n"
        "2014 | Pembentukan Daerah Perlindungan Mangrove | Kel. Pintu Kota, Bitung\n"
        "2013 | IbM Desa Akembawi, Kec. Tahuna Barat, Kab. Kepulauan Sangihe | Sangihe – DIPA Unsrat\n"
        "2012 | Penguatan Kesadaran Masyarakat Dalam Pengelolaan Wilayah Pesisir | Desa Bahoi, Kec. Likupang Barat"
    ),

    # Format: Tahun | Kegiatan | Lingkup Wilayah
    "kebijakan": (
        "2013–2017 | Rencana Zonasi Wilayah Pesisir dan Pulau-Pulau Kecil (RZWP3K) Prov. Sulawesi Utara (Pokja/SK Gubernur) | Provinsi Sulawesi Utara\n"
        "2020 | Tim Penyusun RPJMD Provinsi Sulawesi Utara 2021–2026 (SK Gubernur No. 116/2020) | Provinsi Sulawesi Utara\n"
        "2020–skrg | Tim Teknis/Pokja Revisi RZWP3K & Sinkronisasi dengan RTRW Prov. Sulawesi Utara | Provinsi Sulawesi Utara\n"
        "2021 | Tim Penyusun Integrasi Perda Sulut No. 1/2017 ke dalam Revisi Perda Sulut No. 1/2014 & KLHS Terintegrasi | Provinsi Sulawesi Utara\n"
        "2018–2019 | Narasumber Success Story RZWP3K Sulut dalam Penyusunan RZWP3K Papua Barat | Manokwari, Papua Barat\n"
        "2015 | Pengelolaan Wilayah Pesisir Terpadu Berbasis Masyarakat | 9 Kelurahan Pulau Lembeh, Bitung\n"
        "2016 | Ekowisata Berbasis Masyarakat | Kel. Pasirpanjang, Kareko, dan Pintu Kota, Bitung"
    ),

    # Sertifikasi selam
    "sertifikasi": (
        "Sertifikasi Selam: Open Water — POSSI (1994)\n"
        "Sertifikasi Selam: Open Water — PADI (2004)\n"
        "Sertifikasi Selam: Advanced Open Water — PADI (2004)\n"
        "Pengalaman Penyelaman Ilmiah: >500 jam selam terverifikasi (1994–2024)\n"
        "Spesialisasi: Survei Terumbu Karang Bawah Air, Transplantasi Karang, Reef Monitoring, Pemetaan Habitat Bawah Air\n"
        "Instruktur Pelatihan Selam & Metodologi Penelitian Bawah Air: P3O LIPI Jakarta, Ambon, NTB (1995, 1996, 1999)\n"
        "International Workshop & Training: Field Identification and Taxonomy of Reef Building Corals, Manado, 28–31 Juli 2003"
    ),

    # Keahlian inti (pisah koma)
    "keahlian": (
        "Koralogi & Ekologi Terumbu Karang, "
        "Restorasi & Transplantasi Terumbu Karang, "
        "Iktiologi & DNA Barcoding (COI), "
        "Biodiversitas Pesisir Tropis, "
        "UAV/Drone Mapping Habitat Pesisir, "
        "GIS & Penginderaan Jauh, "
        "Pengelolaan Wilayah Pesisir Terpadu, "
        "Scientific Diving & Reef Survey, "
        "Manajemen Kawasan Konservasi Laut, "
        "Ekowisata Bahari Berbasis Masyarakat"
    ),
}


# ===================== SESSION STATE =====================
def init_ss():
    """Inisialisasi session state agar data form tidak hilang saat interaksi UI."""
    for k, v in DEFAULT.items():
        sk = f"d_{k}"
        if sk not in st.session_state:
            st.session_state[sk] = v
    for fk in ["foto_bytes", "foto_name",
                "ktp_bytes",  "ktp_name",
                "lis_bytes",  "lis_name",
                "sert_list"]:
        if fk not in st.session_state:
            st.session_state[fk] = None if fk != "sert_list" else []

init_ss()


# ===================== HELPERS =====================
def cek_ukuran(f, maks_mb: float) -> bool:
    """Validasi ukuran file upload. Tampilkan error jika melebihi batas."""
    if f is None:
        return True
    mb = len(f.getvalue()) / (1024 * 1024)
    if mb > maks_mb:
        st.error(
            f"⚠️ File **{f.name}** melebihi batas **{maks_mb} MB** "
            f"(ukuran: {mb:.2f} MB). Harap kompres atau ganti file."
        )
        return False
    return True


def preview_file(bts, nama: str, label: str = ""):
    """Tampilkan preview gambar atau ikon PDF setelah upload."""
    if bts is None:
        return
    if label:
        st.caption(f"✅ {label}")
    if str(nama).lower().endswith(".pdf"):
        st.info(f"📄 {nama} (PDF)")
    else:
        try:
            st.image(bts, width=110)
        except Exception:
            st.info(f"🖼️ {nama}")


# ===================== TAB NAVIGASI =====================
tab1, tab2 = st.tabs([
    "📄 Generator CV — Tenaga Ahli / Spesialis Penyelaman",
    "✉️ Generator Surat Tugas Dinas Luar"
])


# ╔══════════════════════════════════════════════╗
# ║           TAB 1 — GENERATOR CV              ║
# ╚══════════════════════════════════════════════╝
with tab1:
    st.markdown("""
    <div class="info-karang">
        📌 Form sudah diisi otomatis dari CV Dr. Ari Berty Rondonuwu (2025).
        Periksa, sesuaikan jika perlu, lalu upload dokumen dan klik <b>Generate CV</b>.
    </div>
    """, unsafe_allow_html=True)

    # ── A. DATA PRIBADI ────────────────────────────────
    st.markdown('<div class="section-hdr">👤 A. Data Pribadi</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        nama     = st.text_input("Nama Lengkap (dengan gelar) *", value=st.session_state["d_nama"])
        nip      = st.text_input("NIP", value=st.session_state["d_nip"])
        nidn     = st.text_input("NIDN", value=st.session_state["d_nidn"])
        ttl      = st.text_input("Tempat, Tanggal Lahir *", value=st.session_state["d_ttl"])
        jabatan  = st.text_input("Jabatan Fungsional *", value=st.session_state["d_jabatan"])
    with c2:
        afiliasi = st.text_input("Institusi / Afiliasi *", value=st.session_state["d_afiliasi"])
        alamat   = st.text_area("Alamat Lengkap *", value=st.session_state["d_alamat"], height=80)
        telepon  = st.text_input("No. Telepon / HP *", value=st.session_state["d_telepon"])
        email    = st.text_input("Email Aktif *", value=st.session_state["d_email"])

    c3, c4 = st.columns(2)
    with c3:
        peran_tim = st.selectbox(
            "Kedudukan dalam Tim *",
            ["Ketua Tim Pelaksana", "Anggota Tim Pelaksana",
             "Pembantu Peneliti", "Penyelam Bersertifikat"],
            index=0
        )
    with c4:
        peran_teknis = st.text_input(
            "Peran Teknis *", value=st.session_state["d_peran_teknis"]
        )

    # ── B. UPLOAD DOKUMEN ──────────────────────────────
    st.markdown('<div class="section-hdr">📎 B. Upload Dokumen Pendukung</div>',
                unsafe_allow_html=True)
    st.markdown("""
    <div class="info-karang">
        Foto, KTP, dan Lisensi Selam <b>wajib</b> diupload sebelum CV dapat di-generate.
        File lampiran akan disisipkan ke dokumen PDF output.
    </div>
    """, unsafe_allow_html=True)

    uc1, uc2 = st.columns(2)

    with uc1:
        # 1. PAS FOTO
        st.markdown('<div class="up-label">📸 Pas Foto 3×4 <span class="up-req">* Wajib</span></div>',
                    unsafe_allow_html=True)
        foto_f = st.file_uploader("JPG/PNG, maks 2 MB", type=["jpg","jpeg","png"],
                                   key="up_foto", label_visibility="collapsed")
        if foto_f and cek_ukuran(foto_f, 2):
            st.session_state["foto_bytes"] = foto_f.getvalue()
            st.session_state["foto_name"]  = foto_f.name
        if st.session_state["foto_bytes"]:
            preview_file(st.session_state["foto_bytes"],
                         st.session_state.get("foto_name","foto.jpg"), "Foto tersimpan")

    with uc2:
        # 2. KTP
        st.markdown('<div class="up-label">🪪 Scan KTP <span class="up-req">* Wajib</span></div>',
                    unsafe_allow_html=True)
        ktp_f = st.file_uploader("JPG/PNG/PDF, maks 2 MB", type=["jpg","jpeg","png","pdf"],
                                  key="up_ktp", label_visibility="collapsed")
        if ktp_f and cek_ukuran(ktp_f, 2):
            st.session_state["ktp_bytes"] = ktp_f.getvalue()
            st.session_state["ktp_name"]  = ktp_f.name
        if st.session_state["ktp_bytes"]:
            preview_file(st.session_state["ktp_bytes"],
                         st.session_state.get("ktp_name","ktp.jpg"), "KTP tersimpan")

    uc3, uc4 = st.columns(2)

    with uc3:
        # 3. LISENSI SELAM
        st.markdown('<div class="up-label">🤿 Lisensi Menyelam SCUBA <span class="up-req">* Wajib</span></div>',
                    unsafe_allow_html=True)
        st.caption("PADI · SSI · CMAS · POSSI · TNI-AL · lainnya")
        lis_f = st.file_uploader("JPG/PNG/PDF, maks 5 MB", type=["jpg","jpeg","png","pdf"],
                                  key="up_lis", label_visibility="collapsed")
        if lis_f and cek_ukuran(lis_f, 5):
            st.session_state["lis_bytes"] = lis_f.getvalue()
            st.session_state["lis_name"]  = lis_f.name
        if st.session_state["lis_bytes"]:
            preview_file(st.session_state["lis_bytes"],
                         st.session_state.get("lis_name","lisensi.jpg"), "Lisensi tersimpan")

    with uc4:
        # 4. SERTIFIKAT LAINNYA
        st.markdown('<div class="up-label">📜 Sertifikat Lainnya <span class="up-opt">(opsional, maks 10 file)</span></div>',
                    unsafe_allow_html=True)
        st.caption("BNSP · Fit to Dive · Pelatihan · dll.")
        sert_fs = st.file_uploader("JPG/PNG/PDF, maks 5 MB/file", type=["jpg","jpeg","png","pdf"],
                                    accept_multiple_files=True, key="up_sert",
                                    label_visibility="collapsed")
        if sert_fs:
            if len(sert_fs) > 10:
                st.error("⚠️ Maksimal 10 file sertifikat.")
                sert_fs = sert_fs[:10]
            ok = [{"bytes": s.getvalue(), "name": s.name}
                  for s in sert_fs if cek_ukuran(s, 5)]
            if ok:
                st.session_state["sert_list"] = ok
                st.caption(f"✅ {len(ok)} sertifikat tersimpan")

    # ── C. PENDIDIKAN ──────────────────────────────────
    st.markdown('<div class="section-hdr">🎓 C. Riwayat Pendidikan</div>', unsafe_allow_html=True)
    st.caption("Format: **Jenjang | Institusi | Bidang Studi | Tahun** — satu baris per jenjang")
    pendidikan = st.text_area("Pendidikan *", value=st.session_state["d_pendidikan"],
                               height=130, label_visibility="collapsed")

    # ── D. RESTORASI ───────────────────────────────────
    st.markdown('<div class="section-hdr">🌊 D. Kepakaran: Restorasi Terumbu Karang</div>',
                unsafe_allow_html=True)
    st.caption("Format: **Lokasi | Metode | Periode | Mitra/Pemberi Kerja** — satu baris per kegiatan")
    restorasi = st.text_area("Restorasi", value=st.session_state["d_restorasi"],
                              height=200, label_visibility="collapsed")

    # ── E. SERTIFIKASI SELAM ───────────────────────────
    st.markdown('<div class="section-hdr">🤿 E. Sertifikasi & Pengalaman Selam</div>',
                unsafe_allow_html=True)
    st.caption("Satu item per baris")
    sertifikasi = st.text_area("Sertifikasi", value=st.session_state["d_sertifikasi"],
                                height=140, label_visibility="collapsed")

    # ── F. PENELITIAN ──────────────────────────────────
    st.markdown('<div class="section-hdr">🔬 F. Pengalaman Penelitian (5 Tahun Terakhir)</div>',
                unsafe_allow_html=True)
    st.caption("Format: **Tahun | Judul Penelitian | Sumber Dana**")
    penelitian = st.text_area("Penelitian", value=st.session_state["d_penelitian"],
                               height=200, label_visibility="collapsed")

    # ── G. PUBLIKASI ───────────────────────────────────
    st.markdown('<div class="section-hdr">📚 G. Publikasi Ilmiah Terpilih</div>',
                unsafe_allow_html=True)
    st.caption("Format: **Tahun | Judul | Jurnal/Penerbit | Indeks**")
    publikasi = st.text_area("Publikasi", value=st.session_state["d_publikasi"],
                              height=180, label_visibility="collapsed")

    # ── H. PENGABDIAN & KEBIJAKAN ─────────────────────
    st.markdown('<div class="section-hdr">🤝 H. Pengabdian Masyarakat & Kebijakan Publik</div>',
                unsafe_allow_html=True)
    h1, h2 = st.columns(2)
    with h1:
        st.caption("Format: **Tahun | Kegiatan | Lokasi/Penyelenggara**")
        pengabdian = st.text_area("Pengabdian", value=st.session_state["d_pengabdian"],
                                   height=220, label_visibility="collapsed")
    with h2:
        st.caption("Format: **Tahun | Kegiatan | Lingkup Wilayah**")
        kebijakan = st.text_area("Kebijakan", value=st.session_state["d_kebijakan"],
                                  height=220, label_visibility="collapsed")

    # ── I. KEAHLIAN ────────────────────────────────────
    st.markdown('<div class="section-hdr">🛠️ I. Keahlian Inti</div>', unsafe_allow_html=True)
    st.caption("Pisahkan dengan koma")
    keahlian = st.text_input("Keahlian", value=st.session_state["d_keahlian"],
                              label_visibility="collapsed")

    # ── TOMBOL GENERATE ────────────────────────────────
    st.markdown("---")
    st.caption("*Field bertanda * wajib diisi. Upload Foto, KTP, dan Lisensi Selam wajib dilakukan.*")

    if st.button("⚡ GENERATE CV — Dr. Ari Berty Rondonuwu",
                 type="primary", use_container_width=True):

        # --- Validasi ---
        err = []
        if not nama.strip():         err.append("Nama Lengkap")
        if not ttl.strip():          err.append("Tempat, Tanggal Lahir")
        if not jabatan.strip():      err.append("Jabatan Fungsional")
        if not afiliasi.strip():     err.append("Institusi/Afiliasi")
        if not alamat.strip():       err.append("Alamat Lengkap")
        if not telepon.strip():      err.append("No. Telepon/HP")
        if not email.strip():        err.append("Email Aktif")
        if not peran_teknis.strip(): err.append("Peran Teknis")
        if not pendidikan.strip():   err.append("Riwayat Pendidikan")
        if not st.session_state.get("foto_bytes"): err.append("Pas Foto 3×4 (wajib upload)")
        if not st.session_state.get("ktp_bytes"):  err.append("Scan KTP (wajib upload)")
        if not st.session_state.get("lis_bytes"):  err.append("Lisensi Menyelam SCUBA (wajib upload)")

        if err:
            st.markdown(
                '<div class="err-block"><b>❌ Lengkapi field berikut sebelum generate:</b><ul>'
                + "".join(f"<li>{e}</li>" for e in err)
                + "</ul></div>",
                unsafe_allow_html=True
            )
        else:
            # Simpan ke session_state
            for key, val in [
                ("d_nama", nama), ("d_nip", nip), ("d_nidn", nidn),
                ("d_ttl", ttl), ("d_jabatan", jabatan), ("d_afiliasi", afiliasi),
                ("d_alamat", alamat), ("d_telepon", telepon), ("d_email", email),
                ("d_peran_tim", peran_tim), ("d_peran_teknis", peran_teknis),
                ("d_pendidikan", pendidikan), ("d_restorasi", restorasi),
                ("d_sertifikasi", sertifikasi), ("d_penelitian", penelitian),
                ("d_publikasi", publikasi), ("d_pengabdian", pengabdian),
                ("d_kebijakan", kebijakan), ("d_keahlian", keahlian),
            ]:
                st.session_state[key] = val.strip() if hasattr(val, "strip") else val

            foto_b64 = None
            if st.session_state["foto_bytes"]:
                foto_b64 = base64.b64encode(st.session_state["foto_bytes"]).decode("utf-8")

            data_cv = CVData(
                nama=nama.strip(), jabatan=jabatan.strip(),
                nip=nip.strip(), nidn=nidn.strip(), ttl=ttl.strip(),
                alamat=alamat.strip(), email=email.strip(), telepon=telepon.strip(),
                afiliasi=afiliasi.strip(),
                pendidikan=pendidikan.strip(), peran_tim=peran_tim,
                peran_teknis=peran_teknis.strip(),
                penelitian=penelitian.strip(), publikasi=publikasi.strip(),
                restorasi=restorasi.strip(), pengabdian=pengabdian.strip(),
                kebijakan=kebijakan.strip(), keahlian=keahlian.strip(),
                sertifikasi=sertifikasi.strip(), foto_b64=foto_b64,
            )

            with st.spinner("⏳ Menyusun dokumen CV... Mohon tunggu."):
                try:
                    docx_bytes = generate_cv_docx(data_cv)
                    st.success("✅ CV berhasil disusun!")

                    # Info lampiran
                    lamp = []
                    if st.session_state.get("ktp_bytes"):
                        lamp.append(f"📋 Lampiran 1 — Scan KTP ({st.session_state.get('ktp_name','')})")
                    if st.session_state.get("lis_bytes"):
                        lamp.append(f"🤿 Lampiran 2 — Lisensi Selam ({st.session_state.get('lis_name','')})")
                    for i, s in enumerate(st.session_state.get("sert_list", []), 3):
                        lamp.append(f"📜 Lampiran {i} — {s['name']}")
                    if lamp:
                        st.info(
                            "📎 **Dokumen lampiran perlu digabungkan ke PDF output:**\n\n"
                            + "\n".join(lamp)
                            + "\n\n_Gunakan Adobe Acrobat, SmallPDF, atau ILovePDF._"
                        )

                    nama_f = nama.strip().replace(" ", "_").replace(".", "").replace(",", "")
                    st.download_button(
                        label="⬇️ Unduh CV (.docx)",
                        data=docx_bytes,
                        file_name=f"CV_{nama_f}_Restorasi_Karang_PPS_Bitung_2026.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Gagal generate CV: {e}")
                    st.exception(e)


# ╔══════════════════════════════════════════════╗
# ║       TAB 2 — SURAT TUGAS DINAS LUAR        ║
# ╚══════════════════════════════════════════════╝
with tab2:
    st.markdown("""
    <div class="info-karang">
        ℹ️ Surat ini otomatis menyertakan klausul pembebasan tugas mengajar dan
        larangan rangkap bayar sesuai Diktum KETUJUH SK Dekan FPIK UNSRAT.
    </div>
    """, unsafe_allow_html=True)

    with st.form("st_form", clear_on_submit=False):
        st.markdown('<div class="section-hdr">👤 Data Personil</div>', unsafe_allow_html=True)
        s1, s2 = st.columns(2)
        with s1:
            st_nama    = st.text_input("Nama Personil *",
                                       value="Dr. Ir. Ari Berty Rondonuwu, M.Sc., M.Si.")
            st_nip     = st.text_input("NIP", value="196801291993031001")
            st_jabatan = st.text_input("Jabatan *",
                                       value="Lektor Kepala (Penata) / Dosen FPIK UNSRAT")
        with s2:
            st_peran  = st.selectbox("Peran dalam Tim *",
                                     ["Ketua Tim Pelaksana", "Anggota Tim Pelaksana", "Pembantu Peneliti"])
            st_jenis  = st.text_input("Jenis Tugas *",
                                      value="Survei Ekologi Awal dan Pemetaan Habitat Terumbu Karang")
            st_lokasi = st.text_input("Lokasi Tugas *",
                                      value="Pelabuhan Perikanan Samudera (PPS) Bitung, Kota Bitung, Sulawesi Utara")

        st.markdown('<div class="section-hdr">📅 Periode Penugasan</div>', unsafe_allow_html=True)
        d1, d2, d3 = st.columns(3)
        with d1: tgl_mulai   = st.date_input("Tanggal Mulai *",   value=date(2026, 9, 15))
        with d2: tgl_selesai = st.date_input("Tanggal Selesai *", value=date(2026, 9, 16))
        with d3: jml_hari    = st.number_input("Jumlah Hari *", min_value=1, max_value=30, value=2)

        st.markdown('<div class="section-hdr">📋 Nomor Surat</div>', unsafe_allow_html=True)
        st_nomor = st.text_input("Nomor Surat *", value="800/.../FPIK-UNSRAT/2026")

        st.markdown('<div class="section-hdr">⚙️ Klausul Khusus</div>', unsafe_allow_html=True)
        klaim_8_oj = st.checkbox("Aktifkan Klausul 8 OJ/hari (Khusus Pembantu Peneliti)", value=False)

        submit_st = st.form_submit_button(
            "✉️ Generate Surat Tugas Dinas Luar", type="primary", use_container_width=True
        )

    if submit_st:
        data_st = DataSuratTugas(
            nomor_surat=st_nomor,
            nama=st_nama, nip=st_nip, jabatan=st_jabatan,
            peran_tim=st_peran, jenis_tugas=st_jenis, lokasi=st_lokasi,
            tanggal_mulai=tgl_mulai, tanggal_selesai=tgl_selesai,
            jumlah_hari=jml_hari, klaim_8_oj=klaim_8_oj,
        )
        with st.spinner("⏳ Menyusun Surat Tugas Dinas Luar..."):
            try:
                st_bytes = generate_st_dinas_luar(data_st)
                st.success("✅ Surat Tugas berhasil disusun!")
                nama_f_st = st_nama.replace(" ", "_").replace(".", "").replace(",", "")
                st.download_button(
                    label="⬇️ Unduh Surat Tugas (.docx)",
                    data=st_bytes,
                    file_name=f"ST_DL_{nama_f_st}_PPS_Bitung_2026.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ Gagal generate Surat Tugas: {e}")
                st.exception(e)


# ===================== FOOTER =====================
st.markdown("---")
st.caption(
    "🌊 Generator Dokumen Tim Pelaksana Swakelola — "
    "Restorasi Terumbu Karang PPS Bitung TA 2026 — "
    "FPIK Universitas Sam Ratulangi | v3.0"
)
