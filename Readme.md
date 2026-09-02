# Automation Sheet Absensi & Generasi Grafik Google Apps Script

Script otomatisasi Google Apps Script untuk merekap data kehadiran dari **Google Forms** secara otomatis ke dalam **spreadsheet** bulanan. Script ini membuat matriks presensi (menggunakan tanda centang `✓`), menghitung total kehadiran per hari latihan (Selasa, Kamis, Sabtu), menyusun tabel rangkuman, dan membuat grafik batang berkelompok secara dinamis.

---

## Fitur Utama

- **Otomatisasi Berbasis Trigger**: Dapat dijalankan secara manual atau otomatis melalui event `On Form Submit`.
- **Penanganan Kolom Respon Dinamis (Kategori Member)**: Membaca nama responden secara langsung dari Kolom D (Warga UKM) atau Kolom E (Member Tenkei)
- **Dynamic Date Matrix**: Membentuk kolom tanggal presensi bulanan secara otomatis khusus untuk hari latihan rutin (**Selasa, Kamis, Sabtu**).
- **Generation system**: Rewrite/replace semua data pada bulan ini, sedangkan bulan sebelum-sebelumnya tidak disentuh.
- **Rangkuman Kehadiran Tahunan**: Fitur kompilasi otomatis (`buatSheetRangkumanTahun()`) yang memindai semua tab bulanan dan menyusun matriks kehadiran tahunan dengan pengelompokan berdasarkan tahun secara dinamis.
- **Pengaturan Posisi Tab Otomatis (`aturUrutanTab`)**: Menyusun tata letak tab spreadsheet secara otomatis dari kiri ke kanan dengan urutan terstruktur.
- **Format & Freeze Otomatis**: 
  - Kolom A ("Nama Responden") diatur pada ukuran *fixed* **175px**.
  - Kolom header di-*merge* secara vertikal untuk keterbacaan yang rapi.
  - Pembekuan baris (Freeze Row 2) & kolom (Freeze Column 1).
- **Tabel Rangkuman (Helper Table)**: Menyusun rincian kehadiran per hari latihan dan total kehadiran bulanan.
- **Grouped Column Chart**: Menghasilkan grafik batang vertikal otomatis tepat di bawah tabel rangkuman untuk visualisasi distribusi kehadiran.

---

## TATA LETAK & URUTAN TAB SPREADSHEET

Fungsi `aturUrutanTab()` secara otomatis menyusun ulang urutan tab dari kiri ke kanan sebagai berikut:

```text
[ JAWABAN FORM ] ➔ [ Rangkuman Kehadiran ] ➔ [ Bulan Terbaru ] ➔ [ Bulan Sebelumnya... ]
```

---

# Alur
> *Catatan: untuk mempermudah penjelasan, tab respon google form dinamakan sebagai `"Form Responses 1"` (dapat disesuaikan di program).  
Pada script di atas, tab respon google form dinamakan sebagai `"JAWABAN FORM"`. dapat dilihat di kode line 12*

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
          │
          ▼
[ Script Menata Posisi Tab (aturUrutanTab()) ]
```
### Ringkasan Function Utama program
| Nama Function | Deskripsi |
| -------- | -------- |
| `buatSheetBulanIniManual()`   | desc 2   |
| `generateSheetAbsensi(targetBulan, targetTahun)`   | desc 5   |
| `buatTabelRangkumanDanGrafik(...)` | desc |
| `buatSheetRangkumanTahun()` | desc |
| `aturUrutanTab()` | desc |

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

---

# Struktur Data
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

## Struktur Data Tab RANGKUMAN KEHADIRAN
```text
KOLOM       HEADER              KETERANGAN                          Formatting
A           Nama Responden      nama yang absen                     Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light green 3" (#d9ead3)
B-n         Tahun, Bulan        jumlah kehadiran di bulan tersebut  Baris 1: Tahun, baris 2: Bulan; bold, Arial, 10 pt, background baris 1 "light cornflower blue 2" (#a4c2f4), background baris 2 "light green 3" (#d9ead3)
n+1         Total Kehadiran     total kehadiran akumulatif anggota  Baris 1 dan 2 merged header, bold, Arial, 10 pt, background "light cornflower blue 2" (#a4c2f4)
```

---

# Version History
###### (biar gausah scroll git lagi hehehe)
### Visualisasi Absen Gform **V1** | 28 Juli 2026
* First deploy 
 
### Visualisasi Absen Gform **V2**    | 7 Agustus 2026  
* Perubahan nama tab sumber data
yang semula bernama `"Form Responses 1"` di-rename menjadi `"JAWABAN FORM"`. 
* Perubahan pertanyaan pada Google Form  
Previous version:  
`1. Multiple-choice: Nama Anggota`  
Current version:  
`1. Multiple choice: Asal Instansi dari UKM / Tenkei?`  
`2. Redirect bercabang multiple choice Nama Anggota.`
```text
Kamu dari UKM / Tenkei? 
├── UKM:  
│   └── Multiple choice Nama Mahasiswa UKM  
├── Tenkei:  
│   └── Multiple choice Nama Warga Tenkei  
└── Submit
```
* Perubahan struktur sumber data di tab `"JAWABAN FORM"`.  
Previous version:
```text
KOLOM       HEADER         KETERANGAN
A           Timestamp      sebagai acuan tanggal absensi
B           Nama           nama yang absen
```
Current version: 
```text
KOLOM       HEADER                  KETERANGAN
A           Timestamp               Acuan tanggal dan bulan absensi
B           Nama (obsolete)         tidak digunakan lagi
C           UKM / Tenkei            jawaban untuk pertanyaan pertama dari UKM atau Tenkei
D           Nama Warga UKM          nama warga ukm yang absen
E           Nama Member Tenkei      nama warga tenkei yang absen
```

### Visualisasi Absen Gform **V3**    | 2 September 2026  
* New Feature: `buatSheetRangkumanTahun()`  
Meng-generate tab `"Rangkuman Kehadiran"` yang berisi rangkuman data kehadiran latihan tiap bulannya.  
Struktur data tab-nya:
```text
KOLOM       HEADER              KETERANGAN
A           Nama Responden      nama yang absen
B-n         Tahun, Bulan        jumlah kehadiran di bulan tersebut.
                                Baris 1: Tahun, baris 2: Bulan.
```

### Visualisasi Absen Gform **V3.1**  | 2 September 2026  
* New Feature: `aturUrutanTab()`  
  Memindahkan dan mengurutkan posisi tab di Spreadsheet sesuai urutan:
   1. "JAWABAN FORM" (Paling Kiri)
   2. "Rangkuman Kehadiran"
   3. Tab Bulanan secara Terbalik (Bulan Terbaru -> Bulan Terlama)
* Perbaikan Fungsi: `buatSheetRangkumanTahun()`
  * Mengurutkan data kehadiran berdasarkan frekuensi kehadiran.
  * Menambahkan kolom `"Total Kehadiran"` di paling kanan header (n+1), sehingga strukturnya menjadi seperti berikut:
```text
KOLOM       HEADER              KETERANGAN
A           Nama Responden      nama yang absen
B-n         Tahun, Bulan        jumlah kehadiran di bulan tersebut.
                                Baris 1: Tahun, baris 2: Bulan.
n+1         Total Kehadiran     total kehadiran akumulatif anggota
```

* Minor Patch:  
  - Perubahan nama tab `"Rangkuman Kehadiran"` menjadi `"RANGKUMAN KEHADIRAN"`  
  - Penambahan section `Ringkasan Function Utama Program` di `Readme.md`
  - Penambahan section `Version History` di `Readme.md`
  - Penambahan comment versi script di baris paling atas kode
  - Perubahan penulisan Readme.md yang semula menggunakan format plain text menjadi menggunakan tabel di bagian tertentu




  TODO:
  - benerin ringkasan fungsi utama program
  - commit tanggal 3
  - ubah plain text format jadi tabel