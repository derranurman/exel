# Sistem Keluar Masuk Barang (Scan Barcode)

File Excel untuk mencatat keluar / masuk barang pakai scanner barcode.
Dibuat tanpa VBA supaya bisa dibuka di Excel versi apa saja (termasuk Excel di
HP / LibreOffice / Google Sheets).

## File utama

- `Sistem_Keluar_Masuk_Barang.xlsx` -> file yang dipakai sehari-hari
- `build_xlsx.py` -> script Python untuk regenerate file xlsx dari nol

## Struktur Sheet

| Sheet | Isi |
| --- | --- |
| `Input Barang` | Scan kode barang masuk. Kolom: No, Kode Barang, Tanggal Masuk |
| `Output Agus` | Scan kode barang keluar oleh Agus. Kolom: No, Kode Barang, Tanggal Keluar |
| `Output Rexa` | Scan kode barang keluar oleh Rexa. Kolom: No, Kode Barang, Tanggal Keluar |
| `Rekap` | Ringkasan total + rekap per-kode (Input, Agus, Rexa, Sisa) |

## Fitur

1. **Scan -> auto tanggal**. Formula `=IF(B2="","",IF(C2="",NOW(),C2))` mengisi
   kolom C otomatis waktu kolom B terisi. Butuh **Iterative Calculation** aktif
   (sudah di-set di file).
2. **Auto nomor urut** di kolom A.
3. **Anti duplikat pada sheet `Input Barang`**. Kalau kode yang sudah ada
   discan lagi, muncul popup "Kode Duplikat". Sel yang duplikat juga
   di-highlight merah (conditional formatting).
4. **Validasi keras pada `Output Agus` dan `Output Rexa`**. Kalau kode yang
   discan TIDAK ADA di `Input Barang`, Excel menolak input dan menampilkan
   popup "Kode Tidak Terdaftar".
5. **Freeze header** + cursor aktif di `B2` pas buka sheet, jadi begitu file
   dibuka tinggal scan.
6. **Rekap** otomatis pakai `COUNTA` + `COUNTIF` -> tidak perlu refresh.

## Langkah Pemakaian (urutan penting)

1. Buka file di Excel.
2. (Pertama kali saja) aktifkan Iterative Calculation:
   - `File -> Options -> Formulas -> centang Enable iterative calculation`
   - Maximum Iterations = `100`, Maximum Change = `0.001`
   - Klik OK, lalu Save.
3. Ke sheet **Input Barang**, scan semua barang yang masuk.
4. Ke sheet **Output Agus** / **Output Rexa**, scan barang yang keluar.
5. Lihat sheet **Rekap** untuk total dan saldo per kode.

## Kenapa tanpa macro / VBA?

Karena VBA cuma jalan di Microsoft Excel desktop (Windows/Mac) dan bikin
file-nya jadi `.xlsm` yang sering diblok email. Versi ini murni `.xlsx`, tetap
dapat fitur yang diminta: auto tanggal, notif duplikat, notif kode tidak
terdaftar, dan auto rekap.

## Regenerate file

Kalau mau bongkar / tambah kolom, ubah `build_xlsx.py` lalu jalankan:

```bash
python3 build_xlsx.py
```

File `Sistem_Keluar_Masuk_Barang.xlsx` akan dibuat ulang.
