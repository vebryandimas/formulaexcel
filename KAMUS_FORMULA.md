# Kamus Formula A-Z

Daftar 136 fungsi Excel/WPS berurutan abjad. Versi interaktif (dengan hasil langsung) ada di sheet `Kamus_Formula` pada `Formula-A-Z.xlsx`.

> Pemisah argumen pada contoh memakai koma (`,`). Pada Excel/WPS berbahasa Indonesia gunakan titik koma (`;`).

| Fungsi | Kategori | Sintaks | Kegunaan | Contoh |
|---|---|---|---|---|
| **ABS** | Matematika | `ABS(angka)` | Nilai mutlak: membuang tanda minus. | `=ABS(-15)` |
| **ADDRESS** | Pencarian & Referensi | `ADDRESS(baris, kolom)` | Membuat alamat sel dalam bentuk teks. | `=ADDRESS(2,3)` |
| **AND** | Logika | `AND(syarat1, syarat2, ...)` | BENAR jika semua syarat terpenuhi. | `=AND(5>3,2>1)` |
| **AVERAGE** | Statistik | `AVERAGE(angka1, angka2, ...)` | Rata-rata. | `=AVERAGE(D6:D15)` |
| **AVERAGEIF** | Statistik | `AVERAGEIF(rentang, kriteria, [rentang_rata])` | Rata-rata dengan satu syarat. | `=AVERAGEIF(C6:C15,"IT",D6:D15)` |
| **AVERAGEIFS** | Statistik | `AVERAGEIFS(rentang_rata, rentang1, kriteria1, ...)` | Rata-rata dengan banyak syarat. | `=AVERAGEIFS(D6:D15,C6:C15,"IT",F6:F15,">80")` |
| **BIN2DEC** | Konversi | `BIN2DEC(biner)` | Biner ke desimal. | `=BIN2DEC("1010")` |
| **CEILING** | Matematika | `CEILING(angka, kelipatan)` | Membulatkan ke atas ke kelipatan tertentu. | `=CEILING(23,5)` |
| **CHAR** | Teks | `CHAR(kode)` | Karakter dari kode angka. | `=CHAR(65)` |
| **CHOOSE** | Pencarian & Referensi | `CHOOSE(nomor, nilai1, nilai2, ...)` | Memilih nilai berdasarkan nomor urut. | `=CHOOSE(2,"Senin","Selasa","Rabu")` |
| **CLEAN** | Teks | `CLEAN(teks)` | Membuang karakter tak tercetak. | `=CLEAN("Halo"&CHAR(7)&"Dunia")` |
| **CODE** | Teks | `CODE(teks)` | Kode angka karakter pertama. | `=CODE("A")` |
| **COLUMN** | Pencarian & Referensi | `COLUMN([sel])` | Nomor kolom sebuah sel. | `=COLUMN(C5)` |
| **COLUMNS** | Pencarian & Referensi | `COLUMNS(rentang)` | Jumlah kolom dalam rentang. | `=COLUMNS(B6:F15)` |
| **CONCATENATE** | Teks | `CONCATENATE(teks1, teks2, ...)` | Menggabungkan teks (alternatif: operator &). | `=CONCATENATE("Halo ","Dunia")` |
| **CONVERT** | Konversi | `CONVERT(angka, dari, ke)` | Konversi satuan (mis. "C" ke "F", "km" ke "mi"). | `=CONVERT(100,"C","F")` |
| **CORREL** | Statistik | `CORREL(larik1, larik2)` | Koefisien korelasi dua data. | `=CORREL(D6:D15,F6:F15)` |
| **COUNT** | Statistik | `COUNT(nilai1, nilai2, ...)` | Menghitung sel berisi angka. | `=COUNT(D6:D15)` |
| **COUNTA** | Statistik | `COUNTA(nilai1, nilai2, ...)` | Menghitung sel yang tidak kosong. | `=COUNTA(B6:B15)` |
| **COUNTBLANK** | Statistik | `COUNTBLANK(rentang)` | Menghitung sel kosong. | `=COUNTBLANK(B6:B20)` |
| **COUNTIF** | Statistik | `COUNTIF(rentang, kriteria)` | Menghitung sel yang memenuhi satu syarat. | `=COUNTIF(C6:C15,"IT")` |
| **COUNTIFS** | Statistik | `COUNTIFS(rentang1, kriteria1, ...)` | Menghitung dengan banyak syarat. | `=COUNTIFS(C6:C15,"IT",F6:F15,">=90")` |
| **DATE** | Tanggal & Waktu | `DATE(tahun, bulan, hari)` | Membuat tanggal dari tahun, bulan, hari. | `=DATE(2026,12,31)` |
| **DATEDIF** | Tanggal & Waktu | `DATEDIF(awal, akhir, satuan)` | Selisih dua tanggal. Satuan: "Y" tahun, "M" bulan, "D" hari, "YM", "MD". | `=DATEDIF(DATE(2000,1,15),DATE(2026,11,26),"Y")` |
| **DATEVALUE** | Tanggal & Waktu | `DATEVALUE(teks_tanggal)` | Mengubah teks tanggal menjadi tanggal. | `=DATEVALUE("2026-11-26")` |
| **DAY** | Tanggal & Waktu | `DAY(tanggal)` | Mengambil hari dari tanggal. | `=DAY(DATE(2026,11,26))` |
| **DAYS360** | Tanggal & Waktu | `DAYS360(awal, akhir)` | Selisih hari dengan 1 tahun = 360 hari (perbankan). | `=DAYS360(DATE(2026,1,1),DATE(2026,12,31))` |
| **DB** | Keuangan | `DB(biaya, sisa, umur, periode)` | Penyusutan saldo menurun tetap. | `=DB(1000000,100000,6,1)` |
| **DDB** | Keuangan | `DDB(biaya, sisa, umur, periode)` | Penyusutan saldo menurun ganda. | `=DDB(1000000,100000,6,1)` |
| **DEC2BIN** | Konversi | `DEC2BIN(angka)` | Desimal ke biner. | `=DEC2BIN(10)` |
| **DEC2HEX** | Konversi | `DEC2HEX(angka)` | Desimal ke heksadesimal. | `=DEC2HEX(255)` |
| **EDATE** | Tanggal & Waktu | `EDATE(tanggal, bulan)` | Tanggal sekian bulan sebelum/sesudahnya. | `=EDATE(DATE(2026,1,31),1)` |
| **EOMONTH** | Tanggal & Waktu | `EOMONTH(tanggal, bulan)` | Tanggal akhir bulan. | `=EOMONTH(DATE(2026,2,10),0)` |
| **EVEN** | Matematika | `EVEN(angka)` | Membulatkan ke atas ke bilangan genap terdekat. | `=EVEN(3)` |
| **EXACT** | Teks | `EXACT(teks1, teks2)` | Membandingkan dua teks persis (peka huruf besar/kecil). | `=EXACT("abc","ABC")` |
| **EXP** | Matematika | `EXP(angka)` | Bilangan e (2,718...) dipangkatkan angka. | `=EXP(1)` |
| **FACT** | Matematika | `FACT(angka)` | Faktorial suatu bilangan. | `=FACT(5)` |
| **FALSE** | Logika | `FALSE()` | Menghasilkan nilai logika SALAH. | `=FALSE()` |
| **FIND** | Teks | `FIND(cari, teks, [mulai])` | Posisi teks (peka huruf besar/kecil). | `=FIND("a","Banana")` |
| **FLOOR** | Matematika | `FLOOR(angka, kelipatan)` | Membulatkan ke bawah ke kelipatan tertentu. | `=FLOOR(23,5)` |
| **FV** | Keuangan | `FV(bunga, periode, angsuran, [pv])` | Nilai masa depan tabungan/investasi. | `=FV(0.06/12,60,-500000)` |
| **GCD** | Matematika | `GCD(angka1, angka2, ...)` | FPB (faktor persekutuan terbesar). | `=GCD(12,18)` |
| **HEX2DEC** | Konversi | `HEX2DEC(heksa)` | Heksadesimal ke desimal. | `=HEX2DEC("FF")` |
| **HLOOKUP** | Pencarian & Referensi | `HLOOKUP(cari, tabel, nomor_baris, [tepat])` | Mencari ke samping pada baris judul tabel. | `=HLOOKUP("Gaji",B5:F15,2,FALSE)` |
| **HOUR** | Tanggal & Waktu | `HOUR(waktu)` | Mengambil jam dari waktu. | `=HOUR(TIME(14,30,0))` |
| **HYPERLINK** | Pencarian & Referensi | `HYPERLINK(alamat, [teks])` | Membuat tautan yang bisa diklik. | `=HYPERLINK("https://github.com","GitHub")` |
| **IF** | Logika | `IF(syarat, jika_benar, jika_salah)` | Percabangan: hasil berbeda tergantung syarat. | `=IF(F6>=80,"Baik","Cukup")` |
| **IFERROR** | Logika | `IFERROR(nilai, jika_error)` | Mengganti hasil error dengan nilai lain. | `=IFERROR(VLOOKUP("Zzz",B6:F15,2,FALSE),"Tidak ada")` |
| **INDEX** | Pencarian & Referensi | `INDEX(rentang, baris, [kolom])` | Mengambil nilai pada posisi tertentu dalam rentang. | `=INDEX(D6:D15,4)` |
| **INDIRECT** | Pencarian & Referensi | `INDIRECT(teks_alamat)` | Mengubah teks alamat menjadi referensi sel. | `=INDIRECT("Latihan_Formula!B6")` |
| **INT** | Matematika | `INT(angka)` | Membulatkan ke bawah menjadi bilangan bulat. | `=INT(8.9)` |
| **IPMT** | Keuangan | `IPMT(bunga, periode_ke, jumlah_periode, pv)` | Porsi bunga pada angsuran ke-n. | `=IPMT(0.12/12,1,12,-10000000)` |
| **IRR** | Keuangan | `IRR(arus_kas)` | Tingkat pengembalian internal suatu proyek. | `=IRR(H6:H10)` |
| **ISBLANK** | Informasi | `ISBLANK(sel)` | BENAR jika sel kosong. | `=ISBLANK(Z100)` |
| **ISERROR** | Informasi | `ISERROR(nilai)` | BENAR jika nilai berupa error. | `=ISERROR(1/0)` |
| **ISEVEN** | Informasi | `ISEVEN(angka)` | BENAR jika angka genap. | `=ISEVEN(4)` |
| **ISNA** | Informasi | `ISNA(nilai)` | BENAR jika nilai berupa kesalahan N/A (data tidak tersedia). | `=ISNA(NA())` |
| **ISNUMBER** | Informasi | `ISNUMBER(nilai)` | BENAR jika nilai berupa angka. | `=ISNUMBER(123)` |
| **ISODD** | Informasi | `ISODD(angka)` | BENAR jika angka ganjil. | `=ISODD(7)` |
| **ISTEXT** | Informasi | `ISTEXT(nilai)` | BENAR jika nilai berupa teks. | `=ISTEXT("abc")` |
| **LARGE** | Statistik | `LARGE(rentang, k)` | Nilai terbesar ke-k. | `=LARGE(D6:D15,2)` |
| **LCM** | Matematika | `LCM(angka1, angka2, ...)` | KPK (kelipatan persekutuan terkecil). | `=LCM(4,6)` |
| **LEFT** | Teks | `LEFT(teks, [jumlah])` | Mengambil karakter dari kiri. | `=LEFT("Indonesia",3)` |
| **LEN** | Teks | `LEN(teks)` | Jumlah karakter. | `=LEN("Excel")` |
| **LN** | Matematika | `LN(angka)` | Logaritma natural. | `=LN(10)` |
| **LOG** | Matematika | `LOG(angka, [basis])` | Logaritma dengan basis tertentu (default 10). | `=LOG(100,10)` |
| **LOOKUP** | Pencarian & Referensi | `LOOKUP(cari, rentang_cari, rentang_hasil)` | Pencarian cocok-terdekat pada data terurut naik (cocok untuk predikat nilai). | `=LOOKUP(85,Sekolah_Nilai!F5:F8,Sekolah_Nilai!G5:G8)` |
| **LOWER** | Teks | `LOWER(teks)` | Mengubah menjadi huruf kecil. | `=LOWER("HALO")` |
| **MATCH** | Pencarian & Referensi | `MATCH(cari, rentang, [tipe])` | Posisi sebuah nilai dalam rentang (tipe 0 = persis). | `=MATCH("Dina",B6:B15,0)` |
| **MAX** | Statistik | `MAX(angka1, angka2, ...)` | Nilai tertinggi. | `=MAX(D6:D15)` |
| **MEDIAN** | Statistik | `MEDIAN(angka1, angka2, ...)` | Nilai tengah. | `=MEDIAN(D6:D15)` |
| **MID** | Teks | `MID(teks, mulai, jumlah)` | Mengambil karakter dari tengah. | `=MID("Indonesia",4,3)` |
| **MIN** | Statistik | `MIN(angka1, angka2, ...)` | Nilai terendah. | `=MIN(D6:D15)` |
| **MINUTE** | Tanggal & Waktu | `MINUTE(waktu)` | Mengambil menit dari waktu. | `=MINUTE(TIME(14,30,0))` |
| **MOD** | Matematika | `MOD(angka, pembagi)` | Sisa hasil bagi. Berguna untuk cek genap/ganjil. | `=MOD(17,5)` |
| **MODE** | Statistik | `MODE(angka1, angka2, ...)` | Nilai yang paling sering muncul. | `=MODE(1,2,2,3,3,3,4)` |
| **MONTH** | Tanggal & Waktu | `MONTH(tanggal)` | Mengambil bulan dari tanggal. | `=MONTH(DATE(2026,11,26))` |
| **MROUND** | Matematika | `MROUND(angka, kelipatan)` | Membulatkan ke kelipatan terdekat. | `=MROUND(1234,100)` |
| **NETWORKDAYS** | Tanggal & Waktu | `NETWORKDAYS(awal, akhir, [libur])` | Jumlah hari kerja (Senin-Jumat) di antara dua tanggal. | `=NETWORKDAYS(DATE(2026,11,1),DATE(2026,11,30))` |
| **NOT** | Logika | `NOT(syarat)` | Membalik BENAR menjadi SALAH dan sebaliknya. | `=NOT(5>3)` |
| **NOW** | Tanggal & Waktu | `NOW()` | Tanggal dan jam saat ini (berubah setiap hitung ulang). | `=NOW()` |
| **NPER** | Keuangan | `NPER(bunga, angsuran, pv)` | Jumlah periode pelunasan. | `=NPER(0.01,-1000000,10000000)` |
| **NPV** | Keuangan | `NPV(bunga, nilai1, nilai2, ...)` | Nilai sekarang bersih dari arus kas masa depan. | `=NPV(0.1,H7:H10)+H6` |
| **ODD** | Matematika | `ODD(angka)` | Membulatkan ke atas ke bilangan ganjil terdekat. | `=ODD(4)` |
| **OFFSET** | Pencarian & Referensi | `OFFSET(awal, baris, kolom)` | Referensi yang digeser dari sel awal. | `=OFFSET(Latihan_Formula!B5,2,1)` |
| **OR** | Logika | `OR(syarat1, syarat2, ...)` | BENAR jika salah satu syarat terpenuhi. | `=OR(5<3,2>1)` |
| **PERCENTILE** | Statistik | `PERCENTILE(rentang, k)` | Persentil ke-k (0 sampai 1). | `=PERCENTILE(F6:F15,0.9)` |
| **PI** | Matematika | `PI()` | Nilai konstanta pi. | `=PI()` |
| **PMT** | Keuangan | `PMT(bunga, periode, pv)` | Angsuran tetap per periode (anuitas). | `=PMT(0.12/12,12,-10000000)` |
| **POWER** | Matematika | `POWER(angka, pangkat)` | Perpangkatan (sama dengan operator ^). | `=POWER(2,10)` |
| **PPMT** | Keuangan | `PPMT(bunga, periode_ke, jumlah_periode, pv)` | Porsi pokok pada angsuran ke-n. | `=PPMT(0.12/12,1,12,-10000000)` |
| **PRODUCT** | Matematika | `PRODUCT(angka1, angka2, ...)` | Mengalikan semua angka. | `=PRODUCT(2,3,4)` |
| **PROPER** | Teks | `PROPER(teks)` | Huruf awal tiap kata menjadi kapital. | `=PROPER("budi santoso")` |
| **PV** | Keuangan | `PV(bunga, periode, angsuran)` | Nilai sekarang dari serangkaian angsuran. | `=PV(0.01,12,-1000000)` |
| **QUOTIENT** | Matematika | `QUOTIENT(angka, pembagi)` | Hasil bagi bulat (tanpa sisa). | `=QUOTIENT(17,5)` |
| **RAND** | Matematika | `RAND()` | Bilangan acak 0 sampai 1 (berubah setiap hitung ulang). | `=RAND()` |
| **RANDBETWEEN** | Matematika | `RANDBETWEEN(bawah, atas)` | Bilangan bulat acak dalam rentang (berubah setiap hitung ulang). | `=RANDBETWEEN(1,100)` |
| **RANK** | Statistik | `RANK(angka, rentang, [urutan])` | Peringkat angka dalam daftar (urutan 0 = terbesar dulu). | `=RANK(F11,F6:F15)` |
| **RATE** | Keuangan | `RATE(periode, angsuran, pv)` | Suku bunga per periode. | `=RATE(12,-1000000,10000000)` |
| **REPLACE** | Teks | `REPLACE(teks, mulai, jumlah, teks_baru)` | Mengganti karakter berdasarkan posisi. | `=REPLACE("2026-11",6,2,"12")` |
| **REPT** | Teks | `REPT(teks, jumlah)` | Mengulang teks. | `=REPT("*",5)` |
| **RIGHT** | Teks | `RIGHT(teks, [jumlah])` | Mengambil karakter dari kanan. | `=RIGHT("Indonesia",4)` |
| **ROMAN** | Konversi | `ROMAN(angka)` | Angka biasa ke angka Romawi. | `=ROMAN(2026)` |
| **ROUND** | Matematika | `ROUND(angka, digit)` | Pembulatan biasa ke jumlah digit tertentu. | `=ROUND(3.14159,2)` |
| **ROUNDDOWN** | Matematika | `ROUNDDOWN(angka, digit)` | Pembulatan ke bawah. | `=ROUNDDOWN(3.789,1)` |
| **ROUNDUP** | Matematika | `ROUNDUP(angka, digit)` | Pembulatan ke atas. | `=ROUNDUP(3.141,2)` |
| **ROW** | Pencarian & Referensi | `ROW([sel])` | Nomor baris sebuah sel. | `=ROW(C5)` |
| **ROWS** | Pencarian & Referensi | `ROWS(rentang)` | Jumlah baris dalam rentang. | `=ROWS(B6:F15)` |
| **SEARCH** | Teks | `SEARCH(cari, teks, [mulai])` | Posisi teks (tidak peka huruf besar/kecil). | `=SEARCH("N","banana")` |
| **SIGN** | Matematika | `SIGN(angka)` | Tanda angka: 1 positif, 0 nol, -1 negatif. | `=SIGN(-7)` |
| **SLN** | Keuangan | `SLN(biaya, sisa, umur)` | Penyusutan garis lurus per periode. | `=SLN(100000000,10000000,5)` |
| **SMALL** | Statistik | `SMALL(rentang, k)` | Nilai terkecil ke-k. | `=SMALL(D6:D15,2)` |
| **SQRT** | Matematika | `SQRT(angka)` | Akar kuadrat. | `=SQRT(144)` |
| **STDEV** | Statistik | `STDEV(angka1, angka2, ...)` | Simpangan baku sampel. | `=STDEV(F6:F15)` |
| **SUBSTITUTE** | Teks | `SUBSTITUTE(teks, lama, baru, [ke])` | Mengganti teks tertentu di dalam teks. | `=SUBSTITUTE("0812-3456","-","")` |
| **SUBTOTAL** | Matematika | `SUBTOTAL(kode, rentang)` | Subtotal yang mengabaikan baris tersembunyi hasil filter (kode 9 = jumlah). | `=SUBTOTAL(9,D6:D15)` |
| **SUM** | Matematika | `SUM(angka1, angka2, ...)` | Menjumlahkan angka atau rentang sel. | `=SUM(D6:D15)` |
| **SUMIF** | Matematika | `SUMIF(rentang, kriteria, [rentang_jumlah])` | Menjumlahkan dengan satu syarat. | `=SUMIF(C6:C15,"Keuangan",D6:D15)` |
| **SUMIFS** | Matematika | `SUMIFS(rentang_jumlah, rentang1, kriteria1, ...)` | Menjumlahkan dengan banyak syarat. | `=SUMIFS(D6:D15,C6:C15,"IT",F6:F15,">80")` |
| **SUMPRODUCT** | Matematika | `SUMPRODUCT(larik1, larik2, ...)` | Mengalikan lalu menjumlahkan; juga dipakai untuk hitung bersyarat. | `=SUMPRODUCT((C6:C15="IT")*D6:D15)` |
| **TEXT** | Teks | `TEXT(nilai, format)` | Mengubah angka menjadi teks berformat. Kode format dapat berbeda antar bahasa Excel. | `=TEXT(7,"000")` |
| **TIME** | Tanggal & Waktu | `TIME(jam, menit, detik)` | Membuat waktu. | `=TIME(8,30,0)` |
| **TODAY** | Tanggal & Waktu | `TODAY()` | Tanggal hari ini (berubah setiap hari). | `=TODAY()` |
| **TRIM** | Teks | `TRIM(teks)` | Membuang spasi berlebih. | `=TRIM("  halo   dunia  ")` |
| **TRUE** | Logika | `TRUE()` | Menghasilkan nilai logika BENAR. | `=TRUE()` |
| **TRUNC** | Matematika | `TRUNC(angka, [digit])` | Memotong desimal tanpa membulatkan. | `=TRUNC(8.97)` |
| **TYPE** | Informasi | `TYPE(nilai)` | Jenis nilai: 1 angka, 2 teks, 4 logika, 16 error. | `=TYPE("abc")` |
| **UPPER** | Teks | `UPPER(teks)` | Mengubah menjadi huruf kapital. | `=UPPER("halo")` |
| **VALUE** | Teks | `VALUE(teks)` | Mengubah teks angka menjadi angka. | `=VALUE("123")+1` |
| **VAR** | Statistik | `VAR(angka1, angka2, ...)` | Varians sampel. | `=VAR(F6:F15)` |
| **VLOOKUP** | Pencarian & Referensi | `VLOOKUP(cari, tabel, nomor_kolom, [tepat])` | Mencari ke bawah pada kolom pertama tabel (FALSE = persis). | `=VLOOKUP("Dina",B6:F15,3,FALSE)` |
| **WEEKDAY** | Tanggal & Waktu | `WEEKDAY(tanggal, [tipe])` | Nomor hari dalam seminggu (tipe 2: Senin = 1). | `=WEEKDAY(DATE(2026,11,26),2)` |
| **WEEKNUM** | Tanggal & Waktu | `WEEKNUM(tanggal, [tipe])` | Nomor minggu dalam setahun. | `=WEEKNUM(DATE(2026,11,26))` |
| **WORKDAY** | Tanggal & Waktu | `WORKDAY(awal, hari, [libur])` | Tanggal setelah sekian hari kerja. | `=WORKDAY(DATE(2026,11,26),5)` |
| **YEAR** | Tanggal & Waktu | `YEAR(tanggal)` | Mengambil tahun dari tanggal. | `=YEAR(DATE(2026,11,26))` |
| **YEARFRAC** | Tanggal & Waktu | `YEARFRAC(awal, akhir, [basis])` | Selisih dua tanggal dalam pecahan tahun. | `=YEARFRAC(DATE(2020,1,1),DATE(2026,7,1))` |
