# Panduan Scraping & Pembentukan Dataset ABSA Penginapan Jawa Timur

Pipeline ini dirancang untuk mengumpulkan data profil penginapan dan ulasan Google Maps di wilayah **Jawa Timur, Indonesia** (mencakup 38 Kabupaten dan Kota: Surabaya, Malang, Batu, Banyuwangi, Jember, Kediri, Sidoarjo, Pasuruan, Probolinggo, Blitar, Madura, dll.), mengekstrak fasilitas dan balasan resmi pemilik hotel, melakukan pembersihan NLP, mendeteksi aspek halal dan umum, serta menghasilkan file CSV per tahap dan buku kerja Excel komprehensif (`Dataset ABSA Penginapan Jawa Timur.xlsx`) dengan 13 sheet lengkap (`ulasan`, `kalimat_aspek`, `daftar_penginapan`, `olah_fasilitas`, `stat_fasilitas`, `kamus_aspek`, serta sheet statistik ringkasan).

---

## 🌟 Fitur Unggulan Pipeline

1. **Cakupan Representatif se-Jawa Timur (38 Kabupaten & Kota)**:
   - Pengambilan sampel seimbang antara kelompok **Halal / Syariah** dan kelompok **Umum** di seluruh kawasan strategis Jawa Timur (Metropolitan Surabaya Raya, Malang Raya & Kota Batu, Tapal Kuda, Mataraman, Pantura, dan Madura).
2. **Scraping Ulasan Penuh (Full Infinite Scroll)**:
   - Mengambil puluhan hingga ratusan ulasan per penginapan (default: 100 ulasan, dapat disesuaikan hingga seluruh ulasan).
   - Otomatis memperluas ulasan panjang yang terpotong (`... Lainnya` / `Lihat lainnya`).
3. **Ekstraksi Fasilitas Hotel Lengkap (`fasilitas`)**:
   - Mengekstrak seluruh fitur pada bagian *Detail hotel* (contoh: *Wi-Fi gratis, Sarapan gratis, Parkir gratis, Kolam renang luar, AC, Spa, Restoran, Pusat kebugaran, Bebas asap rokok, Antar-jemput bandara, dll.*).
   - Menyediakan sheet khusus `olah_fasilitas` (matriks indikator fasilitas per hotel) dan `stat_fasilitas` (analisis sebaran fasilitas hotel di Jawa Timur).
4. **Ekstraksi Balasan / Tanggapan Pemilik Hotel (`balasan_pemilik`)**:
   - Mendeteksi dan mengekstrak tanggapan resmi pengelola hotel (`Tanggapan dari pemilik`), termasuk teks panjang yang terlipat.
5. **Autentikasi Google Anti-Captcha (`login_google.py`)**:
   - Membuka Chrome native dengan direktori profil persisten `data/browser_profile/` untuk login akun Google 1x tanpa terjebak loop *"Konfirmasi bahwa Anda bukan robot"*.
6. **Ekspor Excel Multi-Sheet Bebas Error Formula**:
   - Menghasilkan file Excel 13 sheet dengan sanitasi otomatis teks ulasan berawalan `=` atau `@` sehingga file dapat langsung dibuka di Microsoft Excel tanpa peringatan perbaikan (*repair warning*).
7. **Pembersihan Baris Tunggal CSV (`clean_csv_cell`)**:
   - Menjamin setiap teks ulasan, tanggal, dan balasan tersimpan tepat dalam 1 baris CSV, bebas newline acak (`\n`, `\r`) yang merusak struktur baris.

---

## 📁 Struktur Direktori & File

```
scraping_jatim_wawa/
│
├── data/                                         # Direktori penyimpanan dataset & sesi
│   ├── browser_profile/                         # Profil sesi login browser (diabaikan oleh Git)
│   ├── daftar_penginapan_jatim_mentah.csv       # Hasil Tahap 1: Profil & koordinat penginapan di Jawa Timur
│   ├── raw_ulasan_checkpoint.csv                 # Hasil Tahap 2: Database ulasan mentah (checkpoint)
│   ├── dataset_absa_penginapan_halal_jatim.csv   # Hasil Tahap 3: Dataset ABSA siap anotasi
│   └── Dataset ABSA Penginapan Jawa Timur.xlsx   # Hasil Akhir: File Excel multi-sheet komprehensif
│
├── 01_scrape_daftar_penginapan.py                # Tahap 1: Scrape daftar penginapan se-Jawa Timur
├── 02_scrape_ulasan_penginapan.py                # Tahap 2: Scrape ulasan, fasilitas, & balasan pemilik
├── 03_proses_dataset_absa.py                     # Tahap 3: Preprocessing NLP & Export Excel multi-sheet
├── enrich_detail_penginapan.py                   # Utilitas: Melengkapi telepon, website, harga (USD), alamat
├── login_google.py                               # Utilitas: Login akun Google tanpa loop captcha
├── run_pipeline.py                               # Runner menu interaktif (All-in-one)
├── requirements.txt                              # Daftar dependensi library Python
└── README.md                                     # Dokumentasi teknis & panduan eksekusi
```

---

## 🛠️ Instalasi & Persiapan Awal

### 1. Install Dependensi Python & Playwright
Buka terminal / PowerShell di root folder proyek dan jalankan:
```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Login Akun Google (Direkomendasikan 1x di Awal)
Untuk membuka akses penuh ke seluruh ulasan di Google Maps tanpa batasan dan tanpa captcha:
```bash
python login_google.py
```
> Jendela Google Chrome asli akan terbuka. Silakan login ke akun Google Anda. Setelah berhasil masuk ke akun Google, kembali ke terminal dan tekan **[ENTER]**. Sesi login Anda tersimpan permanen di folder `data/browser_profile/`.

---

## 🚀 Cara Menjalankan Pipeline

Anda dapat menjalankan seluruh alur melalui **Runner Interaktif**:
```bash
python run_pipeline.py
```
Menu interaktif yang tersedia:
```text
=============================================================
   PIPELINE SCRAPING & DATASET ABSA PENGINAPAN JAWA TIMUR
=============================================================
1. Jalankan Tahap 1: Scraping Daftar Penginapan
2. Jalankan Tahap 2: Scraping Ulasan Penginapan
3. Jalankan Tahap 3: Pemrosesan ABSA & Export Excel Lengkap
4. Jalankan Semua Tahap (1 -> 2 -> 3)
5. Lengkapi Metadata Penginapan (Alamat, Telepon, Website, Harga)
6. Login Akun Google (1x Saja untuk Buka Semua Ulasan Tanpa Captcha)
0. Keluar
=============================================================
```

---

## 📖 Panduan Eksekusi per Tahap

### **Tahap 1: Scraping Daftar Penginapan Jawa Timur**
```bash
python 01_scrape_daftar_penginapan.py
```
- Menjalankan pencarian representatif untuk hotel, vila, resort, homestay, dan guest house di seluruh 38 Kabupaten & Kota di Jawa Timur.
- Kueri mencakup kelompok **Halal/Syariah** (Surabaya, Malang, Batu, Banyuwangi, Jember, Kediri, Bromo, Madura, dll.) dan kelompok **Umum** se-Jawa Timur.
- Output: `data/daftar_penginapan_jatim_mentah.csv`.
- Kolom: `data_id`, `place_id`, `nama_penginapan`, `tipe`, `alamat`, `lat`, `lon`, `rating_rata2`, `jumlah_ulasan_total`, `kelompok_sampel`, `area_query`, `kab_query`, `query_sumber`, `kabupaten`, `url`, `harga`, `telepon`, `website`.

---

### **Tahap 2: Scraping Ulasan, Fasilitas, & Balasan Pemilik**
```bash
# Menjalankan scraping dengan target default (100 ulasan per hotel)
python 02_scrape_ulasan_penginapan.py

# Menjalankan dengan kuota ulasan tertentu (contoh: 300 ulasan per hotel)
python 02_scrape_ulasan_penginapan.py --max 300

# Mulai langsung dari nomor urut penginapan tertentu (misal: penginapan ke-50)
python 02_scrape_ulasan_penginapan.py --max 100 --start 50

# Mode Sweep: Periksa kembali hotel-hotel lama yang ulasannya belum mencapai target kuota
python 02_scrape_ulasan_penginapan.py --max 100 --retry-incomplete

# Menguji scraping hanya untuk satu hotel tertentu
python 02_scrape_ulasan_penginapan.py --hotel "Hotel Majapahit Surabaya" --max 100

# Menjalankan di background (headless mode)
python 02_scrape_ulasan_penginapan.py --max 100 --headless
```

**Fitur Cerdas Tahap 2:**
- **Auto-Resume Sekuensial**: Otomatis mendeteksi baris hotel terakhir pada `data/raw_ulasan_checkpoint.csv` dan melanjutkan ke hotel berikutnya tanpa mengulang dari awal.
- **Fast Append Streaming**: Ulasan baru langsung di-append ke CSV seketika per hotel (`mode="a"`).
- **Ekstraksi Fasilitas & Balasan Pemilik**: Mengambil daftar fasilitas dari bagian *Detail hotel* dan respon tanggapan resmi dari pihak hotel.

---

### **(Opsional) Melengkapi Metadata Penginapan**
```bash
python enrich_detail_penginapan.py
```
- Membuka halaman penginapan untuk melengkapi nomor telepon, website resmi, alamat detail, dan harga standar internasional (USD).
- Menyimpan pembaruan secara berkala (setiap 5 penginapan) dan aman terhadap file lock di Windows.

---

### **Tahap 3: Pemrosesan ABSA & Export ke Excel Lengkap**
```bash
python 03_proses_dataset_absa.py
```
- Melakukan pembersihan teks NLP (`teks_bersih`) dan deteksi bahasa (`bahasa`).
- Mengonversi format tanggal relatif ke standar `iso_date` dan `tanggal_iso`.
- Menentukan `status_halal` dan `bukti_halal` berbasis aturan (*rule-based*).
- Menandai flag aspek halal (`flag_makanan_halal`, `flag_fasilitas_ibadah`, `flag_privasi_kamar`, `flag_kolam_renang`, `flag_citra_syariah`).
- Memecah ulasan per kalimat untuk diekstraksi ke sheet `kalimat_aspek` berdasarkan kamus 8 aspek (makanan halal, fasilitas ibadah, kesesuaian syariah, kebersihan, pelayanan, kamar & fasilitas, lokasi, harga).
- Membentuk matriks analisis fasilitas pada sheet `olah_fasilitas` dan statistik sebaran fasilitas pada sheet `stat_fasilitas`.
- Menghasilkan 2 file luaran:
  1. `data/dataset_absa_penginapan_halal_jatim.csv` (Dataset ulasan lengkap dengan kolom label aspek siap anotasi).
  2. `data/Dataset ABSA Penginapan Jawa Timur.xlsx` (File Excel 13-sheet komprehensif).

---

## 📊 Skema & Struktur Data

### 1. File Checkpoint Ulasan (`data/raw_ulasan_checkpoint.csv`)
Terdiri dari 17 kolom standar:
| No | Nama Kolom | Keterangan |
|---|---|---|
| 1 | `review_id` | Hash unik ulasan (kombinasi reviewer, tanggal, teks) |
| 2 | `data_id` | ID unik Google Maps untuk penginapan |
| 3 | `nama_penginapan` | Nama hotel / penginapan |
| 4 | `kabupaten` | Kabupaten/Kota di Jawa Timur |
| 5 | `kelompok_sampel` | Kelompok sampel (`halal` / `umum`) |
| 6 | `kata_kunci` | Kata kunci pencarian awal |
| 7 | `fasilitas` | Fasilitas hotel (*Wi-Fi gratis, Sarapan, Kolam renang luar, AC, dll.*) |
| 8 | `teks_ulasan` | Teks ulasan asli pengunjung (sudah diperluas penuh & dibersihkan enter-nya) |
| 9 | `teks_terjemahan_google` | Terjemahan ulasan jika dalam bahasa asing |
| 10 | `rating_penginapan` | Rating rata-rata keseluruhan penginapan dari ringkasan Google Maps berdesimal (*contoh: 4.4, 4.9*) |
| 11 | `rating_bintang` | Rating bintang yang diberikan reviewer per ulasan berdesimal (*contoh: 5.0, 4.0, 1.0*) |
| 12 | `tanggal` | Teks tanggal ulasan asli di Google Maps (*misal: "3 bulan lalu"*) |
| 13 | `jumlah_like` | Jumlah like pada ulasan |
| 14 | `ada_balasan_pemilik` | Boolean (`True` jika pemilik hotel memberikan tanggapan resmi) |
| 15 | `balasan_pemilik` | Teks lengkap respon / tanggapan dari pihak hotel |
| 16 | `sumber` | Sumber data (`Google Maps`) |
| 17 | `reviewer_total_ulasan` | Jumlah total kontribusi ulasan yang ditulis oleh reviewer |

---

### 2. File Excel Akhir (`data/Dataset ABSA Penginapan Jawa Timur.xlsx`)
File Excel ini terdiri dari **13 lembar kerja (*sheets*)**:

1. **`ulasan` (36 Kolom Lengkap)**:
   - Metadata Tempat: `place_id`, `data_id`, `nama_tempat`, `alamat`, `kategori`, `rating_tempat`, `total_ulasan_tempat`, `harga`, `telepon`, `website`, `latitude`, `longitude`, `kabupaten`.
   - Klasifikasi Syariah: `status_halal`, `bukti_halal`, `kelompok_sampel`, `kata_kunci`.
   - Fasilitas: `fasilitas`.
   - Data Ulasan: `review_id`, `reviewer_total_ulasan`, `rating_ulasan`, `rating_bintang`, `tanggal`, `iso_date`, `tanggal_iso`, `ulasan`, `teks_ulasan`, `teks_terjemahan_google`, `likes`, `ada_balasan_pemilik`, `balasan_pemilik`, `sumber`, `teks_bersih`, `bahasa`, `label_halal_nama`.
2. **`kalimat_aspek`**:
   Potongan per kalimat ulasan yang cocok dengan kamus aspek, dilengkapi kolom `kalimat`, `aspek`, `aspek_inti_halal`, `rating_ulasan`, dan `bahasa`.
3. **`daftar_penginapan` (23 Kolom)**:
   Daftar seluruh profil penginapan unik di Jawa Timur beserta fasilitas, status halal, kontak (telepon, website), koordinat, rating rata-rata, dan total ulasan.
4. **`olah_fasilitas`**:
   Matriks analisis fasilitas per penginapan, memuat `total_fasilitas`, `daftar_fasilitas`, dan kolom biner (*Ya* / *Tidak*) untuk 20 amenitas utama (Wi-Fi, Kolam Renang, AC, Restoran, Sarapan, Parkir, Antar-Jemput, Spa, Gym, dll.).
5. **`stat_fasilitas`**:
   Tabel ranking dan rekapitulasi seluruh jenis fasilitas unik di Jawa Timur, mencakup frekuensi hotel, persentase ketersediaan, perbandingan hotel halal vs umum, serta rata-rata rating penginapan.
6. **`kamus_aspek`**:
   Kamus kata kunci dan pola regex untuk 8 aspek analisis sentimen (makanan halal, fasilitas ibadah, kesesuaian syariah, kebersihan, pelayanan, kamar & fasilitas, lokasi, harga).
7. **Sheet Statistik Ringkasan**:
   - `stat_ringkas`: Ringkasan umum total hotel, total ulasan, kalimat aspek, dan rentang tanggal.
   - `stat_status_halal`: Distribusi penginapan berdasarkan status halal.
   - `stat_aspek`: Sebaran frekuensi kemunculan masing-masing aspek.
   - `stat_bahasa`: Sebaran bahasa ulasan (Indonesia, Inggris, dll.).
   - `stat_rating`: Distribusi rating bintang 1 hingga 5.
   - `stat_tahun`: Distribusi tahun ulasan.
   - `stat_penginapan`: Rekap jumlah ulasan dan rata-rata rating per penginapan.

---

## 💡 Tips & Pemecahan Masalah (Troubleshooting)

- **Mencegah Crash / Editor Freeze (Antigravity / VS Code)**:
  - **TUTUP TAB FILE CSV DI EDITOR SAAT SCRAPING AKTIF**: Jangan biarkan tab file `data/raw_ulasan_checkpoint.csv` terbuka di editor saat scraping sedang berjalan. File yang memuat puluhan ribu baris akan memicu *file watcher* merender ulang dokumen jutaan karakter, yang dapat membuat editor kehabisan memori (*OOM*).
  - Menjalankan scraping langsung melalui terminal Command Prompt / PowerShell murni menghemat memori sistem dan menjamin kestabilan proses.
- **Melanjutkan Scraping Tanpa Mengulang dari Awal**:
  - Cukup jalankan kembali `python 02_scrape_ulasan_penginapan.py`. Skrip otomatis mendeteksi hotel terakhir di file checkpoint dan langsung melanjutkan ke hotel berikutnya.
- **Looping Captcha ("Konfirmasi bahwa Anda bukan robot")**:
  - Jalankan `python login_google.py` untuk login via browser Google Chrome asli Anda satu kali saja. Sesi login akan disimpan permanen di `data/browser_profile/`.
- **Mengecek Progres Data Tanpa Membuka File di Editor**:
  - Jalankan perintah berikut di terminal PowerShell terpisah:
    ```powershell
    python -c "import pandas as pd; df = pd.read_csv('data/raw_ulasan_checkpoint.csv'); print(f'Total ulasan: {len(df)} | Total hotel: {df["data_id"].nunique()}')"
    ```
