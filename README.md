# 📘 Aplikasi Rapor SD Kelas 1-6 (Kurikulum Merdeka)

Aplikasi web sederhana berbasis **Streamlit** untuk membuat Rapor Peserta Didik Sekolah Dasar (Kelas 1 sampai 6) dengan cepat, akurat, dan langsung bisa diunduh dalam format PDF profesional.

## ✨ Fitur Utama

- **Input nilai super mudah** untuk semua mata pelajaran SD (mode Single Student)
- **Batch Processing dari Excel** — Buat puluhan atau ratusan rapor PDF sekaligus hanya dengan mengisi **1 file Excel**! (fitur baru & sangat powerful untuk sekolah)
- **Deskripsi Capaian Kompetensi otomatis** yang positif, memotivasi, dan sesuai kaidah Kurikulum Merdeka
- **Predikat otomatis** (A/B/C/D) berdasarkan nilai
- **PDF rapi & profesional** siap cetak (ukuran A4)
- Mendukung perbedaan mata pelajaran antar kelas:
  - Kelas 1-2 (Fase A): Agama, Pancasila, BI, Matematika, PJOK, Seni Budaya, Muatan Lokal
  - Kelas 3-4 (Fase B): + IPAS
  - Kelas 5-6 (Fase C): + Bahasa Inggris
- Format mengikuti **Panduan Pembelajaran dan Asesmen** Kemendikdasmen

**Update v1.1 (Juni 2026):** 
- Perbaikan layout PDF: proporsi kolom tabel lebih baik (lebih lega untuk deskripsi), wrapping teks dinamis berdasarkan lebar aktual (tidak overflow lagi), garis vertikal pemisah kolom tabel, spasi lebih rapi & profesional.
- Perbaikan penomoran section (A/B/C/D) yang konsisten & benar, termasuk CP/TP opsional.
- Fitur Batch Excel kini menggunakan identitas sekolah dari UI (sebelumnya hardcoded).
- Deskripsi & keterangan di PDF lebih rapi, tidak terpotong, siap cetak.

**Update v1.2 (perbaikan identitas):** 
- Penulisan bagian **IDENTITAS PESERTA DIDIK** dirapikan total: semua tanda titik dua (:) sekarang disejajarkan secara vertikal dengan posisi horizontal yang konsisten di setiap kolom (kiri & kanan). 
- Label dan nilai dipisahkan sehingga tidak lagi bergantung pada padding spasi (yang tidak akurat di font proporsional seperti Helvetica). 
- Hasil: tampilan lebih profesional, rapi, dan mudah dibaca — persis seperti formulir rapor resmi.

**Update v1.3 (Muatan Lokal fleksibel + PDF Lebih Pintar - Juni 2026):** 
- Menambahkan **pilihan nama Muatan Lokal** yang bisa disesuaikan langsung di UI (Single Student & Batch Excel). 
- Sekarang guru bisa mengganti label "Muatan Lokal" menjadi nama spesifik sesuai daerah/sekolah, misalnya: **Bahasa Jawa**, **Bahasa Sunda**, **Budaya Betawi**, **Seni Tradisi Lokal**, **Kearifan Lokal**, dll.
- Nilai tetap diinput di kolom yang sama (atau field nilai), deskripsi otomatis tetap relevan (fokus pelestarian budaya & kearifan daerah).
- Fitur ini memudahkan sekolah yang memiliki Muatan Lokal khusus tanpa perlu menambah mapel manual setiap kali.
- Batch Excel tetap kompatibel (gunakan kolom Nilai_Muatan_Lokal untuk nilainya).

**Perbaikan PDF Penting (v1.3):**
- **Tabel Nilai sekarang 100% anti-potong**: Ketika Anda menambahkan mata pelajaran ekstra via fitur "➕ Tambah Mapel", tinggi setiap baris dihitung **otomatis & dinamis** berdasarkan jumlah baris teks aktual setelah wrapping. Tidak ada lagi teks deskripsi atau nama mapel yang terpotong/hilang, border tabel & garis vertikal menyesuaikan sempurna.
- **Pilihan Ukuran Kertas sebelum Download**: Sebelum klik tombol Generate, Anda bisa pilih **A4 (standar)** atau **F4/Folio (lebih tinggi)**. F4 memberikan ruang ekstra ~33 mm sehingga rapor tetap rapi dan lega meskipun ada 5–10 mapel tambahan.
- Cocok untuk sekolah yang sering menambahkan mapel unggulan atau Muatan Lokal khusus.

## 📥 Cara Install & Menjalankan (5 Menit)

### 1. Pastikan Python sudah terinstall
Buka terminal / Command Prompt dan ketik:
```bash
python --version
```
Harus muncul versi Python 3.9 atau lebih baru.

### 2. Download / Salin Folder `rapor_sd_app`

### 3. Buka terminal di folder tersebut dan install dependencies:
```bash
pip install -r requirements.txt
```

### 4. Jalankan aplikasinya:
```bash
streamlit run app.py
```

Aplikasi akan otomatis terbuka di browser (biasanya http://localhost:8501)

## 🖥️ Cara Menggunakan Aplikasi

1. **Isi Identitas Sekolah** (nama SD, NPSN, alamat)
2. **Isi Identitas Siswa** (nama, NISN, pilih Kelas & Semester & Tahun Ajaran)
3. **Pilih Agama** siswa (penting untuk nama mata pelajaran Agama)
4. **Input Nilai Akhir** (0-100) untuk setiap mata pelajaran
   - Predikat (A/B/C/D) akan muncul otomatis
5. **Edit Deskripsi** (opsional) — deskripsi sudah dibuat otomatis sesuai nilai dan mapel. Anda bisa sesuaikan dengan observasi nyata.
6. Isi **Kehadiran** dan **Catatan Wali Kelas**
7. Isi nama & NIP **Wali Kelas** dan **Kepala Sekolah**
8. Klik tombol besar **GENERATE & DOWNLOAD RAPOR PDF**
9. PDF langsung terunduh — siap dicetak atau diarsipkan!

## 📋 Contoh Output PDF

PDF berisi:
- Header resmi + nama sekolah
- Identitas lengkap siswa + Fase
- Tabel Nilai + Capaian Kompetensi (dengan deskripsi yang rapi)
- Ketidakhadiran
- Catatan Wali Kelas
- 3 kolom tanda tangan (Wali Kelas, Kepala Sekolah, Orang Tua)
- Tanggal & tempat

## ⚠️ Catatan Penting

- Aplikasi ini **gratis** dan dibuat untuk memudahkan guru.
- Deskripsi otomatis sudah menggunakan bahasa positif & memotivasi sesuai contoh resmi Kemendikdasmen.
- **Anda tetap harus menyesuaikan deskripsi** dengan kondisi nyata siswa.

## 📱 Cara Install sebagai Aplikasi di Android (PWA)

Aplikasi ini sudah didukung sebagai **Progressive Web App (PWA)**, sehingga bisa di-install di HP Android seperti aplikasi biasa.

### Langkah-langkah:

1. **Deploy aplikasi ke internet** (paling mudah pakai Streamlit Community Cloud - gratis):
   - Buat akun di [share.streamlit.io](https://share.streamlit.io)
   - Upload folder `rapor_sd_app` ke GitHub (private atau public)
   - Deploy dari Streamlit Cloud

2. **Buka di HP Android** menggunakan browser **Chrome**

3. Ketuk menu **⋮** (tiga titik) di pojok kanan atas

4. Pilih **"Add to Home screen"** atau **"Install Rapor SD"**

5. Tekan **Install** / **Tambahkan**

6. Ikon aplikasi akan muncul di layar utama HP Anda seperti aplikasi native.

Setelah di-install, aplikasi akan terbuka dalam mode fullscreen dan terasa seperti aplikasi Android asli.

> Catatan: Untuk pengalaman terbaik, aplikasi sebaiknya diakses melalui HTTPS (saat di-deploy).
- Untuk sekolah dengan banyak siswa, Anda bisa menjalankan aplikasi berulang kali atau mengembangkan versi batch (Excel → banyak PDF).
- Logo sekolah belum disertakan di versi ini (bisa ditambahkan kemudian).

## 🛠️ Pengembangan Lanjutan (Opsional)

Jika ingin lebih advanced:
- Tambah upload logo sekolah ke PDF
- Support multiple siswa sekaligus (import dari Excel)
- Simpan data ke database
- Fitur tanda tangan digital

## 📞 Dukungan

Jika ada error atau ingin request fitur tambahan, silakan hubungi pembuat aplikasi.

---

**Dibuat dengan ❤️ untuk para guru SD Indonesia**  
Format sesuai Panduan Pembelajaran dan Asesmen Kurikulum Merdeka (edisi terbaru)

Selamat menggunakan! Semoga mempermudah administrasi rapor Anda. 🙏