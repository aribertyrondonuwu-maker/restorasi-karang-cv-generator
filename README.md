# Web Generator CV & Surat Tugas Dinas Luar

Aplikasi web untuk tim administrasi FPIK UNSRAT dalam menghasilkan dokumen
resmi kegiatan **Relokasi dan Replanting/Restorasi Terumbu Karang
PPS Kota Bitung Tahun Anggaran 2026**.

Dokumen yang dihasilkan:

1. **Curriculum Vitae** — untuk Tenaga Ahli dan Tenaga Spesialis Penyelaman
2. **Surat Tugas Dinas Luar** — untuk penugasan perorangan maupun satu tim

---

## Persyaratan sistem

- Python 3.10 atau lebih baru
- Koneksi internet saat pemasangan dependensi

## Pemasangan

```bash
git clone https://github.com/aribertyrondonuwu-maker/restorasi-karang-cv-generator.git
cd restorasi-karang-cv-generator

python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt
```

## Menjalankan aplikasi

```bash
streamlit run app.py
```

Aplikasi akan terbuka di `http://localhost:8501`.

---

## Struktur berkas

| Berkas | Keterangan |
|---|---|
| `app.py` | Titik masuk utama, navigasi sidebar, seluruh form dan validasi |
| `cv_builder.py` | Mesin pembuat CV berformat PDF dan DOCX beserta lampiran |
| `st_dinas_luar_builder.py` | Mesin pembuat Surat Tugas Dinas Luar (PDF dan DOCX) |
| `utils.py` | Fungsi bantu bersama: format tanggal, validasi berkas, olah gambar |
| `requirements.txt` | Daftar dependensi Python |
| `logo_unsrat.png` | Logo instansi untuk kop dokumen |

> Berkas `app_cv.py` pada versi lama **sudah tidak digunakan** dan dapat
> dihapus. Seluruh fungsinya telah dipindahkan ke `app.py`.

---

## Dokumen yang wajib diunggah

| Dokumen | Format | Ukuran maksimal | Wajib |
|---|---|---|---|
| Pas Foto 3x4 | JPG, PNG | 2 MB | Ya |
| KTP | JPG, PNG, PDF | 2 MB | Ya |
| License Menyelam SCUBA | JPG, PNG, PDF | 5 MB | Ya |
| Sertifikat Lainnya | JPG, PNG, PDF | 5 MB per berkas, maks 10 berkas | Tidak |

Lisensi selam berlaku untuk semua penerbit: PADI, SSI, CMAS, TNI-AL,
atau lembaga lainnya.

Berkas KTP, lisensi selam, dan sertifikat lainnya disisipkan otomatis
sebagai halaman lampiran di bagian akhir dokumen CV, masing-masing dengan
judul halaman yang jelas (`Lampiran 1 — KTP`, `Lampiran 2 — License
Menyelam SCUBA`, dan seterusnya).

---

## Spesifikasi dokumen keluaran

- Ukuran kertas **A4 potret**
- Margin: atas 2,5 cm · bawah 2,5 cm · kiri 3 cm · kanan 2 cm
- Huruf **Times New Roman** 12 pt untuk isi, 14 pt tebal untuk judul
- Kop resmi instansi dengan logo pada setiap dokumen
- Pas foto 3x4 di pojok kanan atas halaman pertama CV
- Nomor halaman pada setiap lembar

Format utama adalah **PDF**. Format **DOCX** disediakan sebagai keluaran
tambahan apabila dokumen masih perlu disunting.

### Catatan mengenai lampiran pada berkas Word

Lampiran berformat PDF hanya dapat ditampilkan utuh pada berkas keluaran
PDF. Pada berkas Word, lampiran PDF dicatat sebagai keterangan rujukan,
sedangkan lampiran berupa gambar tetap disisipkan sebagai halaman penuh.

---

## Alur penggunaan

### Membuat CV

1. Buka menu **Generator CV** pada sidebar.
2. Pilih jenis CV: Tenaga Ahli atau Tenaga Spesialis Penyelaman.
3. Isi seluruh data pribadi bertanda bintang (`*`).
4. Unggah keempat dokumen persyaratan pada bagian **Upload Dokumen**.
5. Lengkapi tabel pendidikan, sertifikasi, dan pengalaman.
6. Periksa blok pernyataan dan tanda tangan.
7. Tekan **Generate CV**, periksa pratinjau, lalu unduh PDF atau Word.

Validasi berjalan saat tombol ditekan. Apabila ada isian yang belum
lengkap, seluruh kekurangan ditampilkan dalam satu blok merah di bagian
atas form.

### Membuat Surat Tugas Dinas Luar

1. Buka menu **Generator Surat Tugas**.
2. Isi nomor surat sesuai format penomoran instansi.
3. Isi tabel personil — satu baris untuk perorangan, beberapa baris untuk tim.
4. Lengkapi rincian penugasan dan periode pelaksanaan.
5. Aktifkan klausul 8 OJ/hari bila terdapat personil Pembantu Peneliti.
6. Tekan **Generate Surat Tugas**, lalu unduh PDF atau Word.

---

## Catatan penting

Nomor KTP yang diisi pada form harus **sama persis** dengan nomor yang
tertera pada scan KTP yang diunggah. Ketidaksesuaian akan menyebabkan
dokumen ditolak pada tahap verifikasi administrasi.

Data form disimpan pada `st.session_state` selama sesi berlangsung,
sehingga isian tidak hilang saat berpindah menu. Gunakan tombol
**Reset Seluruh Form** pada sidebar untuk mengosongkan seluruh isian.

---

## Lisensi dan kontak

Fakultas Perikanan dan Ilmu Kelautan
Universitas Sam Ratulangi, Manado
Laman: http://fpik.unsrat.ac.id · Surel: fpik@unsrat.ac.id
