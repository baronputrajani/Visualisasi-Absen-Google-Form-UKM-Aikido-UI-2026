# Automation Sheet Absensi & Generasi Grafik Google Apps Script

Script otomatisasi Google Apps Script untuk merekap data kehadiran dari **Google Forms** secara otomatis ke dalam **spreadsheet** bulanan. Script ini membuat matriks presensi (menggunakan tanda centang `✓`), menghitung total kehadiran per hari latihan (Selasa, Kamis, Sabtu), menyusun tabel rangkuman, dan membuat grafik batang berkelompok secara dinamis.

---

## Fitur Utama

- **Otomatisasi Berbasis Trigger**: Dapat dijalankan secara manual atau otomatis melalui event `On Form Submit`.
- **Penanganan Kolom Respon Dinamis (Kategori Member)**: Membaca nama responden secara langsung dari Kolom B (Nama)
- **Dynamic Date Matrix**: Membentuk kolom tanggal presensi bulanan secara otomatis khusus untuk hari latihan rutin (**Selasa, Kamis, Sabtu**).
- **Format & Freeze Otomatis**: 
  - Kolom A ("Nama Responden") diatur pada ukuran *fixed* **175px**.
  - Kolom header di-*merge* secara vertikal untuk keterbacaan yang rapi.
  - Pembekuan baris (Freeze Row 2) & kolom (Freeze Column 1).
- **Tabel Rangkuman (Helper Table)**: Menyusun rincian kehadiran per hari latihan dan total kehadiran bulanan.
- **Grouped Column Chart**: Menghasilkan grafik batang vertikal otomatis tepat di bawah tabel rangkuman untuk visualisasi distribusi kehadiran.

---

# Alur
untuk mempermudah penjelasan, tab respon google form dinamakan sebagai "Form Responses 1"
dapat diubah diprogram

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
[ Script Membaca Data & Menghasilkan Sheet Bulanan ]
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
kolom B yang berisi nama ]
          │
          ▼
[ Script Meng-generate Tab Bulan ini ] #tab akan selalu di-create dan recreate apabila ada nama bulan tersebut sudah ada di tab lain. bukan update, tapi recreate
          │
          ▼
[ Tab Bulan ini ]
          │
          ▼
[ Nama, hari dan tanggal (selasa, kamis, sabtu), matriks kehadiran, total kehadiran, rangkuman, grafik batang vertikal ]
```

---
 
# Google Form
Pertanyaan pada google form hanyalah dropdown list **NAMA**


## Struktur Data Response Form
```text
KOLOM       HEADER         KETERANGAN
A           Timestamp      sebagai acuan tanggal absensi
B           Nama           nama yang absen
```

## Struktur Data Tab Bulanan
```text
KOLOM       HEADER              KETERANGAN
A           Nama Responden      nama yang absen
B-n         Hari, Tanggal       matriks kehadiran
n+1         Total kehadiran
n+2         Nama                nama yang absen + di paling bawah nama ada grafik batang vertikal
n+3         Selasa              total kehadiran pada hari selasa
n+4         Kamis               total kehadiran pada hari kamis
n+5         Sabtu               total kehadiran pada hari sabtu
n+6         Total kehadiran
```