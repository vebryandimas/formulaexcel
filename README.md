# Formula A-Z

Kamus formula dan template **Excel / WPS Office** siap pakai untuk perkantoran, sekolah, dan keperluan sehari-hari. Gratis, tanpa makro, dan bebas dimodifikasi.

## Isi

Semua ada dalam satu file: [`Formula-A-Z.xlsx`](Formula-A-Z.xlsx) (12 sheet, sekitar 1.700 rumus).

| Sheet | Kegunaan |
|---|---|
| `Beranda` | Daftar isi, cara pakai, kode warna, catatan penting |
| `Kamus_Formula` | 136 fungsi berurutan A-Z: sintaks, kegunaan, contoh, dan **hasil langsung**. Bisa difilter per huruf, kategori, dan tingkat |
| `Latihan_Formula` | Data contoh + 25 rumus latihan (SUMIFS, VLOOKUP, INDEX-MATCH, DATEDIF, PMT, IRR, dll.) |
| `Kantor_Gaji` | Penggajian: tunjangan, lembur, potongan absen, iuran, gaji bersih |
| `Kantor_Faktur` | Faktur/invoice dengan diskon, PPN, total, dan jatuh tempo |
| `Kantor_Anggaran` | Anggaran vs realisasi dengan status otomatis |
| `Kantor_Stok` | Stok barang, peringatan restock, nilai persediaan |
| `Kantor_Cicilan` | Kalkulator cicilan anuitas dan flat, jadwal angsuran sampai 120 bulan |
| `Kantor_Penjualan` | Rekap penjualan per sales, produk, dan bulan, lengkap dengan grafik |
| `Sekolah_Nilai` | Nilai akhir berbobot, predikat, ketuntasan, peringkat kelas |
| `Sekolah_Absensi` | Absensi siswa bulanan (H/S/I/A) dengan rekap dan persentase hadir |
| `Umum_Kalkulator` | 11 kalkulator: BMI, umur, diskon, PPN, persentase, hari kerja, suhu, jam kerja, tabungan, patungan, angka Romawi |

Versi teks kamus untuk dibaca langsung di GitHub: [`KAMUS_FORMULA.md`](KAMUS_FORMULA.md).

## Cara pakai

1. Unduh `Formula-A-Z.xlsx` (klik berkas, lalu tombol **Download raw file**).
2. Buka dengan Microsoft Excel 2010 ke atas atau WPS Spreadsheet.
3. Pilih sheet yang dibutuhkan, isi **sel kuning**, dan hapus data contoh.
4. Jangan menimpa sel putih karena berisi rumus otomatis.

**Kode warna:** kuning + huruf biru = isian, putih + huruf hitam = rumus, hijau muda = total/hasil utama.

## Catatan penting

- **Pemisah rumus.** Excel/WPS dengan pengaturan regional Indonesia memakai titik koma (`;`) sebagai pemisah argumen, bukan koma (`,`). Rumus di dalam file otomatis menyesuaikan. Namun saat menyalin contoh dari `KAMUS_FORMULA.md`, ganti `,` dengan `;` bila perlu.
- **Fungsi `TEXT`.** Kode format di dalamnya dapat berbeda antar bahasa/regional.
- **Angka contoh.** Seluruh angka, tarif (misalnya PPN dan iuran), dan nama pada template hanyalah ilustrasi. Sesuaikan dengan kebijakan dan peraturan yang berlaku (pajak, ketenagakerjaan, perbankan). Template ini bukan nasihat pajak atau hukum.
- **Kompatibilitas.** Hanya memakai fungsi klasik (tanpa `XLOOKUP`, `FILTER`, `LET`, dsb.) dan tanpa makro, sehingga aman di Excel lama, WPS Office, dan LibreOffice.
- **Google Sheets.** File dapat diimpor ke Google Sheets, tetapi tampilan dan validasi data bisa sedikit berbeda.

## Berkontribusi

Ide template baru, perbaikan, atau tambahan fungsi sangat diterima. Lihat [CONTRIBUTING.md](CONTRIBUTING.md).

## Lisensi

[MIT](LICENSE): bebas dipakai, diubah, dan dibagikan, termasuk untuk keperluan komersial, dengan mencantumkan hak cipta asli.
