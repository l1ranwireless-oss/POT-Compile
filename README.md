# POT Compile

## Menjalankan deployment

Pastikan Docker Engine dan Docker Compose plugin tersedia, lalu jalankan:

```sh
sh deploy.sh
```

Script membangun image Nginx langsung dari `index.html` terbaru dan menjalankan layanan. Buka `http://localhost:8080`. Port dapat diubah, misalnya `PORT=9000 sh deploy.sh`.

Saat halaman dibuka, menu awal menyediakan tautan **Yuk Bikin POT**, **POT Automated**, dan **Excel Compiler**. Tema default menggunakan palet warna sejuk; preferensi mode terang/gelap dapat diubah melalui tombol tema.

## Excel Compiler

Pilih **Excel Compiler** pada menu awal untuk menggabungkan 2 sampai 20 file `.xlsx`, `.xls`, atau `.csv` menjadi satu sheet. Sertakan setidaknya satu `.xlsx` sebagai template agar workbook hasil mempertahankan sheet dan formatnya; baris dari file lain dipetakan berdasarkan header (tidak peka kapitalisasi dan spasi) dan urutan kolom template. Header yang berbeda, duplikat, atau kosong ditolak agar nilai tidak bergeser. Batas total upload adalah 500 MiB. Upload dan workbook dibaca/ditulis bertahap melalui file sementara agar penggunaan memori tetap terkendali.

Fitur ini menggunakan Python API terbuka pada deployment Docker. GitHub Pages bersifat statis dan tidak menjalankan API Python, sehingga Excel Compiler memerlukan deployment Docker.

## Database CMDB

API membuat tabel SQLite `cmdb_files` secara otomatis saat pertama kali dijalankan dan menyimpan satu file sumber CMDB pada named volume Docker `cmdb_data`. File CMDB yang telah tersimpan pada format server sebelumnya akan dimigrasikan ke SQLite saat startup. Tidak diperlukan service atau proses database terpisah. Gunakan **Simpan CMDB** untuk penyimpanan pertama, **Update Server** untuk mengganti file, dan **Muat Server** untuk memulihkan salinan server. Cache browser di IndexedDB tetap tersedia untuk pemulihan lokal.

Jalankan `sh deploy.sh` untuk memulai layanan. Layanan web di-bind ke `127.0.0.1` secara default dan port dapat diubah, misalnya `PORT=9000 sh deploy.sh`. Endpoint API tidak memiliki autentikasi, jadi jangan paparkan ke jaringan publik tanpa proteksi akses pada reverse proxy. Untuk akses jaringan, gunakan HTTPS dan batasi akses ke pengguna/jaringan yang berwenang. GitHub Pages bersifat statis dan tidak menyediakan API kompilasi atau penyimpanan CMDB.

## Pencocokan site

Sheet `bcp_segmentation_export_sheet` menggunakan daftar Site ID dari POT sebagai dasar. Kolom `Impact Category` yang kosong diisi `Fully Site Down`. Site yang tidak ditemukan di file Segmentation tetap disertakan dengan status `UP`; data Segmentation dan CMDB dicocokkan menggunakan Site ID yang dinormalisasi. Nilai POT/Segmentation yang sudah terisi dipertahankan, sedangkan kolom kosong yang tersedia di CMDB dilengkapi dari CMDB.

GDOC mengambil setiap Site ID nonkosong dari daftar POT, tanpa terpengaruh filter tampilan, dan mengisi Site ID serta Key dari baris POT bila baris segmentation pasangannya kosong. Setiap Site ID unik dipertahankan (menggunakan baris POT terbaru untuk ID yang berulang); detail yang kosong dilengkapi dari POT, Segmentation, atau kolom terkait di CMDB. Status GDOC dipasangkan kembali menggunakan Site ID yang sama dan mengikuti nilai `UP`/`DOWN` eksplisit pada `Remark` output segmentation; jika tidak ada, status output segmentation dan hasil pencocokan Segmentation menjadi fallback. Sebelum workbook dibuat, validasi memastikan semua Site ID sumber masuk ke GDOC dan setiap baris GDOC memiliki status `UP` atau `DOWN`; ketidaklengkapan akan menghasilkan error, bukan sheet yang tampak berhasil tetapi kehilangan data. Kolom `Super Sites` GDOC menggunakan nilai sumber yang tersedia terlebih dahulu, lalu mengambil `Program` dari CMDB sebagai fallback sebelum `Is Hub`.

Perhitungan site down, PG, dan ambang MC >25% hanya menghitung status `DOWN`. `INCIDENT SUMMARY` mengelompokkan site `DOWN` berdasarkan `MC Cluster` di CMDB, lalu mengakumulasi jumlah site dan denominator CMDB dari seluruh MC terdampak; jumlah kategori Big/Medium dan alarm juga ditampilkan lintas MC terdampak. Ringkasan mengambil semua Site ID GDOC, termasuk status `UP`, serta menampilkan daftar khusus site yang tidak ada di Segmentation Info sehingga site pulih tidak hilang dari report. Untuk Site Name, MC Cluster, Hub Type, Traffic Class, P0, dan VIP, nilai CMDB per Site ID diprioritaskan; output GDOC menjadi fallback jika nilai CMDB kosong. Detail `Transport Outage` diprioritaskan dari kolom yang sama pada output `bcp_segmentation_export_sheet` berdasarkan Site ID, kemudian CMDB dan input Segmentation menjadi fallback; detail ditampilkan di antara BTS Name dan status genset, misalnya `08TIS0017_SELUMA_TIMUR_MT (TRANSPORT OUTAGE) (NO DG)`. Kategori `Super Site` dibandingkan dengan kolom `Program` CMDB per Site ID; nilai kosong, `Istimewa`, atau penanda negatif diabaikan.