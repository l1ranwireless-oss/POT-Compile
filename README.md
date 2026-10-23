# POT Compile

## Menjalankan deployment

Pastikan Docker Engine dan Docker Compose plugin tersedia, lalu jalankan:

```sh
sh deploy.sh
```

Script membuat folder `dist/`, menyalin aplikasi, membangun image Nginx, dan menjalankan layanan. Buka `http://localhost:8080`. Port dapat diubah, misalnya `PORT=9000 sh deploy.sh`.

## Database CMDB

Saat file CMDB diunggah, aplikasi memvalidasi sheet dan kolom `Site Id`, lalu membuat cache chunk untuk kalkulasi. File sumber disimpan ke IndexedDB browser agar dapat dipulihkan setelah refresh, dan dapat dihapus melalui tombol **Clear**. Jika IndexedDB tidak tersedia atau penyimpanan gagal, CMDB tetap dapat digunakan selama sesi browser saat ini, tetapi perlu diunggah kembali setelah refresh. Data tetap berada di browser/perangkat pengguna; deployment statis ini tidak mengirim CMDB ke server. Akses melalui `localhost` atau HTTPS untuk mengaktifkan penyimpanan IndexedDB.

## Pencocokan site

Sheet `bcp_segmentation_export_sheet` menggunakan daftar Site ID dari POT sebagai dasar. Kolom `Impact Category` yang kosong diisi `Fully Site Down`. Site yang tidak ditemukan di file Segmentation tetap disertakan dengan status `UP`; data Segmentation dan CMDB dicocokkan menggunakan Site ID yang dinormalisasi. Nilai POT/Segmentation yang sudah terisi dipertahankan, sedangkan kolom kosong yang tersedia di CMDB dilengkapi dari CMDB. GDOC mencakup seluruh Site ID unik dari POT, memilih data terbaru saat ID berulang, mencocokkan status ke Segmentation, dan melengkapi data kosong dari CMDB. Perhitungan site down, PG, dan ambang MC >25% hanya menghitung status `DOWN`.