# Panduan Kontribusi

Terima kasih sudah ingin membantu! Cara termudah:

- **Menemukan kesalahan atau punya ide?** Buka *Issue* dan jelaskan sheet, sel, dan hasil yang diharapkan.
- **Ingin menambah template atau fungsi?** Buat *Pull Request* dengan file `Formula-A-Z.xlsx` yang sudah diubah.

## Aturan agar file tetap konsisten

1. **Rumus, bukan angka mati.** Hasil perhitungan harus berupa rumus supaya ikut berubah saat data diubah.
2. **Fungsi klasik saja.** Hindari `XLOOKUP`, `FILTER`, `SORT`, `UNIQUE`, `LET`, `LAMBDA`, dan makro agar tetap jalan di Excel lama dan WPS.
3. **Kode warna.** Sel isian: latar kuning pucat + huruf biru. Sel rumus: huruf hitam. Total/hasil utama: hijau muda.
4. **Font Arial** di seluruh sheet.
5. **Tanpa error.** Pastikan tidak ada `#DIV/0!`, `#N/A`, `#NAME?`, `#REF!`, atau `#VALUE!` pada kondisi normal. Lindungi pembagian dengan `IF` atau `IFERROR`.
6. **Data contoh fiktif.** Jangan memasukkan data pribadi nyata.
7. **Angka aturan (pajak, iuran, dsb.)** ditaruh di sel parameter yang diberi catatan, bukan ditulis langsung di dalam rumus.
8. **Fungsi baru di kamus:** isi sintaks, kegunaan, contoh yang berdiri sendiri, dan pastikan kolom hasil menampilkan nilai yang benar.

## Ide pengembangan

- Terbilang rupiah untuk kwitansi
- Jadwal piket/shift, kalender kerja, daftar hadir rapat
- Rapor semester lengkap dengan deskripsi capaian
- Pembukuan sederhana UMKM (buku kas, laba rugi)
- Kartu piutang/utang dan pengingat jatuh tempo

Dengan berkontribusi, Anda setuju karyanya dirilis di bawah lisensi MIT yang sama.
