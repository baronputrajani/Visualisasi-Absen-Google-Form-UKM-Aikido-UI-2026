# Automation Sheet Absensi & Generasi Grafik Google Apps Script

Script otomatisasi Google Apps Script untuk merekap data kehadiran dari **Google Forms** secara otomatis ke dalam **spreadsheet** bulanan. Script ini membuat matriks presensi (menggunakan tanda centang `✓`), menghitung total kehadiran per hari latihan (Selasa, Kamis, Sabtu), menyusun tabel rangkuman, dan membuat grafik batang berkelompok secara dinamis.

---

## Fitur Utama

- **Otomatisasi Berbasis Trigger**: Dapat dijalankan secara manual atau otomatis melalui event `On Form Submit`.
- **Penanganan Kolom Respon Dinamis (Kategori Member)**: Membaca nama responden secara langsung dari Kolom D (Warga UKM) atau Kolom E (Member Tenkei)
- **Dynamic Date Matrix**: Membentuk kolom tanggal presensi bulanan secara otomatis khusus untuk hari latihan rutin (**Selasa, Kamis, Sabtu**).
- **Generation system**: Rewrite/replace semua data pada bulan ini, sedangkan bulan sebelum-sebelumnya tidak disentuh.
- **Format & Freeze Otomatis**: 
  - Kolom A ("Nama Responden") diatur pada ukuran *fixed* **175px**.
  - Kolom header di-*merge* secara vertikal untuk keterbacaan yang rapi.
  - Pembekuan baris (Freeze Row 2) & kolom (Freeze Column 1).
- **Tabel Rangkuman (Helper Table)**: Menyusun rincian kehadiran per hari latihan dan total kehadiran bulanan.
- **Grouped Column Chart**: Menghasilkan grafik batang vertikal otomatis tepat di bawah tabel rangkuman untuk visualisasi distribusi kehadiran.

---

# Alur
> *Catatan: untuk mempermudah penjelasan, tab respon google form dinamakan sebagai `"Form Responses 1"` (dapat disesuaikan di program)*

## ALUR USER

```text
[ User Mengisi GForm ]
          │
          ▼
[ Respon Masuk ke Tab "Form Responses 1" ]
          │
          ▼
[ Trigger / Script Dieksekusi ]
          │
          ▼
[ Script Membaca Data, Menghasilkan Sheet Bulanan dan Rangkuman ]
```
---

## Alur Script
```text
[ Sheet "Form Responses 1" terupdate  ]
          │
          ▼
[ Trigger / Script Dieksekusi ]
          │
          ▼
[ Script Membaca:
Kolom A yang berisi timestamp dan 
kolom D dan E yang berisi nama ]
          │
          ▼
[ Script Meng-generate Tab Bulan ini ]  
(tab akan selalu di-create dan recreate jika nama bulan tersebut sudah ada di tab lain. bukan update, tapi recreate)
          │
          ▼
[ Output Tab Bulan ini: ]
  - Nama, hari dan tanggal (Selasa, Kamis, Sabtu)
  - Matriks Kehadiran (✓) & Total Kehadiran
  - Tabel Rangkuman Per Hari
  - Grafik Batang Vertikal Berkelompok
          │
          ▼
[ Eksekusi Rekap Tahunan: ]
Script Memindai Seluruh Tab Bulanan & Menghasilkan Tab "Rangkuman Kehadiran"
```

---
 
# Google Form
Pertanyaan pada google form:  
1. Multiple choice: **kamu dari UKM / Tenkei?**  
2. Redirect bercabang:  
  - Jika memilih UKM ➔ (Redirect) Multiple choice: **Nama Mahasiswa UKM**  
  - Jika memilih Tenkei ➔ (Redirect) Multiple choice: **Nama Warga Tenkei**

```text
Kamu dari UKM / Tenkei? 
├── UKM:  
│   └── Multiple choice Nama Mahasiswa UKM  
├── Tenkei:  
│   └── Multiple choice Nama Warga Tenkei  
└── Submit
```


## Struktur Data Response Form (`Form Responses 1`)
```text
KOLOM       HEADER                  KETERANGAN
A           Timestamp               Acuan tanggal, bulan, dan tahun absensi
B           Nama (obsolete)         tidak digunakan lagi
C           UKM / Tenkei            jawaban untuk pertanyaan pertama dari UKM atau Tenkei
D           Nama Warga UKM          nama warga ukm yang absen
E           Nama Member Tenkei      nama warga tenkei yang absen
```

## Struktur Data Tab Bulanan
```text
KOLOM       HEADER              KETERANGAN                            Formatting
A           Nama Responden      nama yang absen                       Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
B-n         Hari, Tanggal       matriks kehadiran                     Baris 1: Hari, baris 2: Tanggal; bold, Arial, 10 pt, background "light green 3" (#d9ead3)
n+1         Total kehadiran                                           Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light cornflower blue 2" (#a4c2f4)
n+2         Nama                nama yang absen +                     Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
                                di bawah ada grafik batang vertikal
n+3         Selasa              total kehadiran pada hari selasa      Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
n+4         Kamis               total kehadiran pada hari kamis       Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
n+5         Sabtu               total kehadiran pada hari sabtu       Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
n+6         Total kehadiran                                           Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light cornflower blue 2" (#a4c2f4)
```

## Struktur Data Tab Bulanan
```text
KOLOM       HEADER              KETERANGAN                          Formatting
A           Nama Responden      nama yang absen                     Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
B-n         Tahun, Bulan        jumlah kehadiran di bulan tersebut  Baris 1: Tahun, baris 2: Bulan; bold, Arial, 10 pt, background baris 1 "light cornflower blue 2" (#a4c2f4), background baris 2 "light green 3" (#d9ead3)
```