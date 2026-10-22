# Bot Auto-Fill & Komentar Kuesioner FASIH-SM BPS (SE2026)

Bot otomasi pembersihan data (*auto-cleaning*) dan pengisian komentar kuesioner pada aplikasi **FASIH-SM BPS** (`https://fasih-sm.bps.go.id`) berbasis **Playwright**.

---

## 🌟 Fitur Utama

1. **Cukup 1 Command (100% Standalone Engine)**:
   - Tidak perlu membuka Google Chrome atau login manual terlebih dahulu.
   - Skrip otomatis meluncurkan Google Chrome dengan profil terisolasi (`./chrome_data`).
   - Otomatis melakukan login SSO BPS jika belum login atau sesi telah berakhir.
2. **Auto-Resume & Checkpoint (Aman & Efisien)**:
   - Menyimpan progres secara real-time ke file JSON (`source/*_progress.json`).
   - Baris yang sudah berstatus **`SUCCESS` otomatis dilewati**, sehingga Anda dapat menghentikan dan melanjutkan proses kapan saja tanpa menduplikasi pekerjaan.
   - Baris yang pernah `FAILED` otomatis dicoba ulang saat skrip dijalankan kembali.
3. **Rekap Hasil Otomatis ke Excel**:
   - Seluruh status baris, waktu eksekusi, dan detail pesan otomatis diekspor ke file Excel hasil (`source/*_result.xlsx`).
4. **Opsi Headless & Visual**:
   - Dapat dijalankan secara visual (jendela Chrome terlihat di layar) atau di latar belakang (`--headless`).

---

## 📁 Struktur Direktori

```text
auto-cleaning/
├── batch_cleaning.py       # Script batch runner otomasi kirim anomali & nomor KK (cleaning1.xlsx)
├── cleaning_nik.py         # Script batch runner otomasi komentar NIK duplikat #0 (cleaning_0_nik_duplikat.xlsx)
├── update_penjualan.py     # Script batch runner otomasi update nilai penjualan & komentar (penjualan.xlsx)
├── login_fasih.py          # Script login otomatis SSO BPS (menyimpan sesi ke ./chrome_data)
├── config.py               # Konfigurasi URL, kredensial SSO, dan path bersama
├── source/                 # Folder file sumber Excel & rekap hasil
│   ├── cleaning1.xlsx
│   ├── cleaning1_progress.json
│   ├── cleaning1_result.xlsx
│   ├── cleaning_0_nik_duplikat.xlsx
│   ├── cleaning_nik_progress.json
│   ├── cleaning_nik_result.xlsx
│   ├── penjualan.xlsx
│   ├── penjualan_progress.json
│   └── penjualan_result.xlsx
├── chrome_data/            # Direktori profil browser Chrome terisolasi (di-ignore git)
├── pyproject.toml          # Konfigurasi dependensi project
├── requirements.txt        # Daftar library Python
└── README.md               # Dokumentasi lengkap
```

---

## ⚙️ Konfigurasi Akun SSO BPS (`config.py`)

Kredensial SSO BPS untuk auto-login disimpan di [`config.py`](config.py):

```python
# SSO BPS Credentials (dapat disesuaikan oleh pengguna lain)
SSO_USERNAME = os.getenv("SSO_USERNAME")
SSO_PASSWORD = os.getenv("SSO_PASSWORD")
AUTO_LOGIN_SSO = os.getenv("AUTO_LOGIN_SSO", "true").lower() in ("true", "1", "yes")
```

Pengguna lain dapat langsung mengganti nilai di file `config.py` atau menyetel environment variable:
```bash
export SSO_USERNAME="username_anda"
export SSO_PASSWORD="password_anda"
python cleaning_nik.py
```

---

## 1️⃣ Batch Cleaning NIK Duplikat Komentar `#0` (`cleaning_nik.py`)

Gunakan skrip `cleaning_nik.py` untuk otomatis memberikan komentar `#0` pada pertanyaan **"Nomor Induk Kependudukan (NIK)"** untuk seluruh ART dari file Excel `source/cleaning_0_nik_duplikat.xlsx` (Total 9.635 baris).

> [!TIP]
> **Optimasi Pengelompokan Assignment ID**:
> Skrip secara otomatis mengelompokkan 9.635 baris Excel ke dalam **5.112 unique Assignment ID**. Setiap assignment **hanya dibuka dan disubmit 1 kali**, di mana seluruh ART yang terdaftar pada assignment tersebut dikomentari secara berurutan dalam satu sesi.

### Alur Kerja per Assignment:
1. **Otomatis Meluncurkan Google Chrome** dengan profil terisolasi (`./chrome_data`).
2. **Otomatis Login SSO BPS** jika sesi habis (kredensial dari `config.py`).
3. **Membuka link kuesioner** dalam mode edit (`.../edit`) sekali per assignment.
4. **Masuk ke section sidebar** `SE2026 - L BLOK I`.
5. **Loop untuk Seluruh ART**:
   - Mencari sub-section ART yang sesuai dengan kolom **`NAMA ART SE`** dari Excel dan mengkliknya.
   - Mencari baris pertanyaan `Nomor Induk Kependudukan (NIK)` untuk ART tersebut.
   - Mengklik tombol Catatan, mengisi komentar `#0` (jika belum ada), dan menekan tombol **Simpan**.
6. **Masuk ke section CATATAN**, memastikan switch **`Tampilkan Anomali Usaha dan Keluarga`** aktif (tanpa mencentang switch admin dan tanpa mengisi catatan penjelasan).
7. **Klik Kirim utama**, memastikan **0 Galat** pada modal Ringkasan Validasi, lalu mengklik tombol **Kirim** di modal.
8. **Konfirmasi Final**: Mengklik tombol **Konfirmasi** dan memverifikasi notifikasi sukses (*"Assignment Berhasil Disubmit"*).
9. **Checkpointing & Resume**: Otomatis menyimpan progres per baris ke `source/cleaning_nik_progress.json` dan mengekspor rekapitulasi ke `source/cleaning_nik_result.xlsx`.

### Contoh Perintah:

```bash
# 1. Jalankan langsung untuk seluruh assignment (Otomatis skip yang sudah SUCCESS)
python cleaning_nik.py

# 2. Jalankan di latar belakang tanpa jendela browser (Headless mode)
python cleaning_nik.py --headless

# 3. Preview daftar assignment & ART yang akan diproses tanpa mengeksekusi (Dry-run)
python cleaning_nik.py --dry-run --limit 10

# 4. Uji coba proses 5 assignment pertama / berikutnya
python cleaning_nik.py --limit 5

# 5. Memproses rentang baris tertentu (contoh: baris 100 sampai 200)
python cleaning_nik.py --start 100 --end 200

# 6. Memaksa proses ulang baris yang sudah pernah sukses
python cleaning_nik.py --force
```

---

## 2️⃣ Batch Cleaning Anomali & Nomor KK (`batch_cleaning.py`)

Gunakan skrip `batch_cleaning.py` untuk menjalankan pembersihan dan pengiriman otomatis dari file Excel `source/cleaning1.xlsx`:

### Alur Kerja per Baris:
1. **Otomatis Meluncurkan Google Chrome** dengan profil terisolasi (`./chrome_data`).
2. **Otomatis Login SSO BPS** jika sesi habis.
3. **Membuka mode edit** (`.../edit`).
4. **Masuk blok `SE2026 - P`** $\rightarrow$ Cari **`Nomor KK`** $\rightarrow$ Isi komentar `#0` & simpan.
5. **Masuk blok `CATATAN`** $\rightarrow$ Memastikan switch **`Tampilkan Anomali Usaha dan Keluarga`** aktif.
6. **Validasi 0 Galat** $\rightarrow$ Klik **Kirim** $\rightarrow$ Konfirmasi $\rightarrow$ Verifikasi submit sukses.
7. **Rekap Hasil**: Hasil eksekusi otomatis direkap dan diekspor ke file Excel `source/cleaning1_result.xlsx`.

### Contoh Perintah:

```bash
# 1. Jalankan langsung (Cukup 1 Perintah: Chrome buka otomatis & login otomatis!)
python batch_cleaning.py

# 2. Jalankan di latar belakang (Headless mode)
python batch_cleaning.py --headless

# 3. Preview daftar baris yang akan diproses tanpa mengeksekusi (Dry-run)
python batch_cleaning.py --dry-run --limit 10

# 4. Uji coba proses 5 baris berikutnya
python batch_cleaning.py --limit 5

# 5. Memproses rentang baris tertentu (contoh: baris 50 sampai 100)
python batch_cleaning.py --start 50 --end 100

# 6. Memaksa proses ulang SEMUA baris
python batch_cleaning.py --force
```

---

## 3️⃣ Batch Update Nilai Penjualan & Komentar (`update_penjualan.py`)

Gunakan skrip `update_penjualan.py` untuk mengubah nilai penjualan/pendapatan pada section **`KETERANGAN USAHA/PERUSAHAAN`** berdasarkan file Excel `source/penjualan.xlsx`.

### Format File Sumber (`source/penjualan.xlsx`)

| Kolom | Keterangan |
|-------|------------|
| `link` | Link assignment FASIH (boleh tanpa `/edit`) |
| `nama_usaha` | Nama sub-blok usaha, contoh: `BENGKEL MOTOR [SUPARDI]` |
| `nilai_penjualan` | Nilai penjualan/pendapatan yang akan diisi, contoh: `52834000` |
| `komentar` | Komentar catatan pada field nilai penjualan, contoh: `#1` |

> [!TIP]
> Buat file contoh otomatis dengan `python update_penjualan.py --init-sample`, lalu isi barisnya.

### Alur Kerja per Baris:
1. **Otomatis Meluncurkan Google Chrome** dengan profil terisolasi (`./chrome_data`).
2. **Otomatis Login SSO BPS** jika sesi habis (kredensial dari `config.py`).
3. **Membuka mode edit** (`.../edit`).
4. **Masuk section `KETERANGAN USAHA/PERUSAHAAN`** → Memilih sub-blok usaha sesuai `nama_usaha`.
5. **Mengisi textfield nilai penjualan** sesuai kolom `nilai_penjualan`.
6. **Membuka tombol Catatan** pada field nilai penjualan → isi komentar (mis. `#1`) → simpan.
7. **Masuk section `CATATAN`** → aktifkan switch **`Tampilkan Anomali Usaha dan Keluarga`**, aktifkan switch **`Anomali diselesaikan oleh admin`**, isi penjelasan `-`.
8. **Validasi 0 Galat** → Klik **Kirim** → Konfirmasi → Verifikasi submit sukses.
9. **Rekap Hasil**: Hasil eksekusi otomatis direkap dan diekspor ke `source/penjualan_result.xlsx`.

### Contoh Perintah:

```bash
# 0. Buat file Excel sumber contoh lalu isi datanya
python update_penjualan.py --init-sample

# 1. Jalankan langsung (skip yang sudah SUCCESS)
python update_penjualan.py

# 2. Jalankan di latar belakang (Headless mode)
python update_penjualan.py --headless

# 3. Preview daftar baris yang akan diproses tanpa mengeksekusi (Dry-run)
python update_penjualan.py --dry-run --limit 10

# 4. Uji coba proses 5 baris berikutnya
python update_penjualan.py --limit 5

# 5. Memproses rentang baris tertentu (contoh: baris 1 sampai 50)
python update_penjualan.py --start 1 --end 50

# 6. Memaksa proses ulang SEMUA baris
python update_penjualan.py --force
```
