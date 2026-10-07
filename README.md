# ZAIN.NET — Perapih Dokumen Akademik V4

Versi **GitHub Pages upload langsung** untuk merapikan file Microsoft Word `.docx` secara lokal di browser. Tidak membutuhkan npm, Vite, backend, maupun API Gemini.

## Mode dokumen V4

### 1. Skripsi
- Struktur BAB I–V.
- Bagian awal dan isi memakai mesin penomoran terpisah.
- Sampul pertama tanpa nomor tampak.
- Cover 2/bagian awal dapat memakai Romawi kecil; isi dimulai dari angka 1.
- Halaman awal BAB: nomor bawah-tengah; halaman isi lain: kanan-atas.
- TOC native Microsoft Word Level 1–3.

### 2. Proposal
- Auto Deteksi **Dengan BAB / Tanpa BAB**.
- Proposal tanpa BAB mengenali bagian A/B/C dan subbagian angka/huruf.
- Mendukung struktur proposal penelitian dan heading umum proposal kegiatan.
- Sampul tanpa nomor tampak; halaman awal yang tersedia Romawi; isi utama mulai angka 1.

### 3. Makalah
- Auto Deteksi **Dengan BAB / Tanpa BAB**.
- Sampul tanpa nomor tampak.
- Kata Pengantar dan Daftar Isi memakai Romawi jika tersedia.
- Isi utama dimulai dari angka 1.

## Pengaman V4
- **Semua (Aman)** memaksa Mode Aman.
- Setiap tahap perbaikan memiliki pemeriksaan konten dan rollback otomatis bila teks/tabel/gambar berpotensi hilang.
- Validasi akhir memeriksa teks utama, tabel, drawing, media, footnote, endnote, komentar, dan embedding.
- Header/footer lama dikloning per section agar perubahan nomor halaman tidak merusak section lain.
- Field PAGE lama dibersihkan utuh (termasuk cached result), kemudian dibuat ulang sesuai profil dokumen.
- Section landscape dipertahankan pada Mode Aman.
- Cover dan bagian awal tidak ikut terkena pembersihan Enter massal.
- Enter/jarak ekstrem pada BAB, judul, subjudul dan sub-subjudul dirapikan secara terarah.
- TOC memakai field Microsoft Word `TOC` Level 1–3, bukan teks daftar isi palsu.

## File yang harus di-upload
Upload semua file berikut langsung ke **root repository**:

- `index.html`
- `app.js`
- `processor.js`
- `styles.css`
- `jszip.min.js`
- `README.md`
- `CARA_UPLOAD.txt`

## GitHub Pages
Cara paling sederhana:

1. Upload semua file di atas ke root repository.
2. `Settings → Pages`.
3. `Source: Deploy from a branch`.
4. Branch `main` dan folder `/(root)`.
5. Save dan tunggu deployment selesai.

Jika repository lama Anda sudah memakai workflow **Static HTML / GitHub Actions**, file V4 tetap bisa diletakkan di root lalu tunggu workflow deploy selesai.

## Setelah DOCX dibuka di Microsoft Word
Tekan **Ctrl+A → F9** agar field PAGE, TOC, Daftar Tabel, dan Daftar Gambar diperbarui menggunakan mesin layout Microsoft Word.

> Selalu simpan file asli. Aplikasi mengedit struktur XML DOCX secara selektif, tetapi dokumen Word dari kampus berbeda dapat memiliki struktur khusus yang perlu diperiksa secara visual.
