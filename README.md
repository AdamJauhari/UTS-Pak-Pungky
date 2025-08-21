# Sistem Manajemen Zakat (Python + MySQL)

Program ini adalah aplikasi berbasis Python untuk mengelola data zakat, master beras, dan transaksi zakat beras menggunakan database MySQL. Data dapat diekspor ke file Excel untuk keperluan laporan.

## Fitur Utama
- Menambah, mengedit, dan menghapus data pembayar zakat
- Menambah dan melihat data master beras
- Mencatat transaksi zakat beras
- Melihat daftar transaksi zakat
- Ekspor data zakat dan transaksi ke file Excel

## Kebutuhan Sistem
- Python 3.x
- MySQL Server
- Modul Python: `mysql-connector-python`, `pandas`

## Instalasi Modul
Jalankan perintah berikut di terminal untuk menginstal modul yang dibutuhkan:

```
pip install mysql-connector-python pandas
```

## Cara Menjalankan
1. Pastikan MySQL Server sudah berjalan dan dapat diakses dengan user `root` tanpa password (atau sesuaikan di kode).
2. Jalankan program:
   ```
   python "uts mysql.py"
   ```
3. Ikuti menu interaktif di terminal untuk mengelola data zakat.

## Struktur Database
- **zakat_data**: Data pembayar zakat
- **master_beras**: Data jenis beras dan harga
- **transaksi_zakat**: Data transaksi zakat beras

## Ekspor Data
- Data pembayar zakat: `data_zakat.xlsx`
- Data transaksi zakat: `data_transaksi_zakat.xlsx`

## Catatan
- Pastikan koneksi ke MySQL sesuai dengan konfigurasi di kode.
- Jika ada error modul, pastikan sudah menginstal semua kebutuhan Python.

## Lisensi
Program ini dibuat untuk keperluan pembelajaran dan tugas UTS.
