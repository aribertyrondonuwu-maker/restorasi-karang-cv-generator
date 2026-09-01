"""
utils.py — Kumpulan fungsi bantu yang dipakai bersama oleh seluruh modul.

Berisi: format tanggal Bahasa Indonesia, validasi ukuran berkas unggahan,
normalisasi gambar, dan konstanta identitas kegiatan.
"""

import io
import os
from datetime import date
from typing import Optional, Tuple

from PIL import Image, ImageOps

# =====================================================================
# KONSTANTA IDENTITAS KEGIATAN
# =====================================================================

NAMA_KEGIATAN = (
    "RELOKASI DAN REPLANTING/RESTORASI TERUMBU KARANG "
    "KOTA BITUNG TAHUN 2026"
)
NAMA_KEGIATAN_PENDEK = "Relokasi & Restorasi Terumbu Karang PPS Bitung TA 2026"

INSTANSI_KEMENTERIAN = "KEMENTERIAN PENDIDIKAN TINGGI, SAINS, DAN TEKNOLOGI"
INSTANSI_UNIVERSITAS = "UNIVERSITAS SAM RATULANGI"
INSTANSI_FAKULTAS = "FAKULTAS PERIKANAN DAN ILMU KELAUTAN"
INSTANSI_ALAMAT = (
    "Kampus UNSRAT Manado 95115  |  Laman: http://fpik.unsrat.ac.id  |  "
    "Surel: fpik@unsrat.ac.id"
)

DEKAN_NAMA = "Dr. Ir. Ockstan Jurike Kalesaran, M.Sc."
DEKAN_NIP = "196910241994032014"

# --- Lokasi berkas logo ---
_HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(_HERE, "logo_unsrat.png")
if not os.path.exists(LOGO_PATH):
    LOGO_PATH = "logo_unsrat.png"


def logo_tersedia() -> bool:
    """Mengembalikan True apabila berkas logo instansi ditemukan."""
    return os.path.exists(LOGO_PATH)


# =====================================================================
# BATAS UKURAN BERKAS UNGGAHAN (dalam MB, sesuai ketentuan teknis)
# =====================================================================

BATAS_MB = {
    "foto": 2,
    "ktp": 2,
    "lisensi": 5,
    "sertifikat": 5,
}

LABEL_DOKUMEN = {
    "foto": "Pas Foto 3x4",
    "ktp": "KTP",
    "lisensi": "License Menyelam SCUBA",
    "sertifikat": "Sertifikat Lainnya",
}

MAKS_SERTIFIKAT = 10


# =====================================================================
# FORMAT TANGGAL
# =====================================================================

_BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def format_tanggal_indonesia(tgl: Optional[date]) -> str:
    """Mengubah objek date menjadi teks tanggal Bahasa Indonesia.

    Contoh: date(2026, 9, 15) -> "15 September 2026".
    Mengembalikan string kosong bila tanggal tidak diisi.
    """
    if not tgl:
        return ""
    return f"{tgl.day} {_BULAN_ID[tgl.month - 1]} {tgl.year}"


# =====================================================================
# VALIDASI BERKAS UNGGAHAN
# =====================================================================

def validasi_ukuran(berkas, jenis: str) -> Optional[str]:
    """Memeriksa ukuran berkas terhadap batas yang ditetapkan.

    Mengembalikan pesan galat bila melebihi batas, atau None bila lolos.
    Parameter `jenis` harus salah satu kunci pada konstanta BATAS_MB.
    """
    if berkas is None:
        return None

    batas_mb = BATAS_MB.get(jenis, 5)
    batas_byte = batas_mb * 1024 * 1024
    ukuran = getattr(berkas, "size", None)

    if ukuran is None:
        try:
            ukuran = len(berkas.getvalue())
        except Exception:
            return None

    if ukuran > batas_byte:
        return (
            f"File {berkas.name} melebihi batas ukuran {batas_mb} MB. "
            f"Harap kompres atau gunakan file lain."
        )
    return None


def ekstensi(nama_berkas: str) -> str:
    """Mengembalikan ekstensi berkas dalam huruf kecil tanpa titik."""
    return os.path.splitext(nama_berkas or "")[1].lower().lstrip(".")


def adalah_pdf(nama_berkas: str) -> bool:
    """Mengembalikan True apabila berkas berekstensi .pdf."""
    return ekstensi(nama_berkas) == "pdf"


def adalah_gambar(nama_berkas: str) -> bool:
    """Mengembalikan True apabila berkas merupakan gambar JPG/PNG."""
    return ekstensi(nama_berkas) in ("jpg", "jpeg", "png")


# =====================================================================
# NORMALISASI GAMBAR
# =====================================================================

def normalisasi_gambar(data_gambar: bytes, maks_sisi: int = 1600) -> bytes:
    """Menormalkan gambar agar aman disisipkan ke PDF maupun DOCX.

    Langkah: koreksi orientasi EXIF, konversi ke mode RGB (membuang alpha),
    perkecil bila sisi terpanjang melebihi `maks_sisi`, lalu simpan ulang
    sebagai JPEG. Bila gagal, data asli dikembalikan apa adanya.
    """
    try:
        img = Image.open(io.BytesIO(data_gambar))
        img = ImageOps.exif_transpose(img)

        if img.mode in ("RGBA", "LA", "P"):
            latar = Image.new("RGB", img.size, (255, 255, 255))
            img = img.convert("RGBA")
            latar.paste(img, mask=img.split()[-1])
            img = latar
        else:
            img = img.convert("RGB")

        if max(img.size) > maks_sisi:
            rasio = maks_sisi / max(img.size)
            ukuran_baru = (int(img.width * rasio), int(img.height * rasio))
            img = img.resize(ukuran_baru, Image.LANCZOS)

        keluaran = io.BytesIO()
        img.save(keluaran, format="JPEG", quality=88)
        return keluaran.getvalue()
    except Exception:
        return data_gambar


def dimensi_gambar(data_gambar: bytes) -> Tuple[int, int]:
    """Mengembalikan (lebar, tinggi) gambar dalam piksel."""
    try:
        with Image.open(io.BytesIO(data_gambar)) as img:
            return img.size
    except Exception:
        return (0, 0)


def potong_pas_foto(data_gambar: bytes) -> bytes:
    """Memotong gambar menjadi rasio 3:4 (pas foto) dengan pemotongan tengah.

    Digunakan agar pas foto pada CV selalu proporsional meskipun berkas
    yang diunggah memiliki rasio berbeda.
    """
    try:
        img = Image.open(io.BytesIO(data_gambar))
        img = ImageOps.exif_transpose(img)
        img = img.convert("RGB")

        rasio_target = 3 / 4
        lebar, tinggi = img.size
        rasio_kini = lebar / tinggi

        if rasio_kini > rasio_target:
            lebar_baru = int(tinggi * rasio_target)
            kiri = (lebar - lebar_baru) // 2
            img = img.crop((kiri, 0, kiri + lebar_baru, tinggi))
        else:
            tinggi_baru = int(lebar / rasio_target)
            atas = (tinggi - tinggi_baru) // 2
            img = img.crop((0, atas, lebar, atas + tinggi_baru))

        keluaran = io.BytesIO()
        img.save(keluaran, format="JPEG", quality=90)
        return keluaran.getvalue()
    except Exception:
        return normalisasi_gambar(data_gambar)


# =====================================================================
# BANTUAN TEKS
# =====================================================================

def aman(teks) -> str:
    """Mengembalikan teks yang sudah dibersihkan, atau tanda strip bila kosong."""
    if teks is None:
        return "-"
    teks = str(teks).strip()
    return teks if teks else "-"


def nama_berkas_aman(nama: str, awalan: str = "Dokumen") -> str:
    """Menyusun nama berkas unduhan yang aman dari karakter khusus."""
    bersih = "".join(c for c in (nama or "") if c.isalnum() or c in " -_")
    bersih = "_".join(bersih.split())
    return f"{awalan}_{bersih}" if bersih else awalan


# =====================================================================
# DRAF RINGKASAN KETERKAITAN DENGAN KEGIATAN
# =====================================================================

_KATA_KUNCI_RELEVAN = [
    "karang", "terumbu", "reef", "restorasi", "transplantasi", "coral",
    "konservasi laut", "ekosistem laut", "selam", "diving", "pesisir",
]


def _mengandung_kata_kunci(teks: str) -> bool:
    """Memeriksa apakah suatu teks memuat kata kunci yang relevan dengan
    kegiatan restorasi terumbu karang."""
    teks = (teks or "").lower()
    return any(kata in teks for kata in _KATA_KUNCI_RELEVAN)


def buat_draf_afiliasi(nama: str, jenis_cv: str,
                       bidang_keahlian: Optional[list] = None,
                       pengalaman: Optional[list] = None,
                       publikasi: Optional[list] = None,
                       keahlian_selam: Optional[list] = None,
                       total_jam_selam: int = 0) -> str:
    """Menyusun draf ringkasan keterkaitan seseorang dengan kegiatan restorasi
    terumbu karang, berdasarkan data pengalaman kerja/penyelaman, publikasi,
    dan keahlian yang telah diisi pada form.

    Hasilnya adalah draf awal yang ditujukan untuk disunting lebih lanjut,
    bukan teks final yang siap pakai tanpa pemeriksaan.
    """
    bidang_keahlian = bidang_keahlian or []
    pengalaman = pengalaman or []
    publikasi = publikasi or []
    keahlian_selam = keahlian_selam or []
    nama_singkat = (nama or "").strip() or "Yang bersangkutan"

    if not (pengalaman or publikasi or bidang_keahlian or keahlian_selam):
        return (
            "Draf belum dapat disusun karena data pengalaman, keahlian, "
            "atau penelitian pada form di atas belum diisi. Lengkapi data "
            "tersebut terlebih dahulu, lalu tekan tombol \"Buat / Perbarui "
            "Draf Otomatis\"."
        )

    kalimat = []

    if jenis_cv == "Tenaga Spesialis Penyelaman":
        lokasi_relevan = []
        for b in pengalaman:
            gabungan = " ".join(str(v) for v in b.values())
            if _mengandung_kata_kunci(gabungan):
                lokasi = str(b.get("Lokasi Penyelaman", "")).strip()
                if lokasi and lokasi not in lokasi_relevan:
                    lokasi_relevan.append(lokasi)

        if total_jam_selam:
            teks_jam = (
                f"{nama_singkat} memiliki total {total_jam_selam} jam "
                f"selam terverifikasi"
            )
            if keahlian_selam:
                teks_jam += f", dengan keahlian khusus pada {', '.join(keahlian_selam)}"
            kalimat.append(teks_jam + ".")
        elif keahlian_selam:
            kalimat.append(
                f"{nama_singkat} memiliki keahlian khusus penyelaman pada "
                f"{', '.join(keahlian_selam)}."
            )

        if lokasi_relevan:
            kalimat.append(
                "Pengalaman penyelaman yang relevan dengan kegiatan "
                "restorasi terumbu karang telah dilaksanakan di "
                f"{', '.join(lokasi_relevan)}."
            )
        elif pengalaman:
            kalimat.append(
                f"Tercatat {len(pengalaman)} pengalaman penyelaman pada "
                "berbagai lokasi dan jenis kegiatan."
            )

        kalimat.append(
            "Dengan latar belakang tersebut, yang bersangkutan memiliki "
            "kompetensi teknis yang relevan untuk mendukung pelaksanaan "
            f"kegiatan {NAMA_KEGIATAN.title()}."
        )
    else:
        lokasi_relevan = []
        for b in pengalaman:
            gabungan = " ".join(str(v) for v in b.values())
            if _mengandung_kata_kunci(gabungan):
                lokasi = str(b.get("Lokasi", "")).strip()
                if lokasi and lokasi not in lokasi_relevan:
                    lokasi_relevan.append(lokasi)

        publikasi_relevan = [
            p.get("Judul", "").strip() for p in publikasi
            if _mengandung_kata_kunci(p.get("Judul", ""))
            and str(p.get("Judul", "")).strip()
        ]

        if bidang_keahlian:
            kalimat.append(
                f"{nama_singkat} memiliki latar belakang keahlian pada "
                f"bidang {', '.join(bidang_keahlian)}."
            )

        if lokasi_relevan:
            kalimat.append(
                "Pengalaman kerja yang relevan dengan kegiatan restorasi "
                f"terumbu karang telah dilaksanakan di "
                f"{', '.join(lokasi_relevan)}."
            )
        elif pengalaman:
            kalimat.append(
                f"Tercatat {len(pengalaman)} pengalaman kerja pada "
                "berbagai proyek dan instansi."
            )

        if publikasi_relevan:
            kalimat.append(
                f"Yang bersangkutan turut mempublikasikan "
                f"{len(publikasi_relevan)} karya ilmiah yang relevan "
                "dengan topik terumbu karang dan ekosistem pesisir."
            )

        kalimat.append(
            "Dengan latar belakang tersebut, yang bersangkutan memiliki "
            "keterkaitan langsung dengan pelaksanaan kegiatan "
            f"{NAMA_KEGIATAN.title()}."
        )

    return " ".join(kalimat)
