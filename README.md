# POT Compile

## Menjalankan deployment

Pastikan Docker Engine dan Docker Compose plugin tersedia, lalu jalankan:

```sh
sh deploy.sh
```

Script membangun image Nginx langsung dari `index.html` terbaru dan menjalankan layanan. Buka `http://localhost:8080`. Port dapat diubah, misalnya `PORT=9000 sh deploy.sh`.

Saat halaman dibuka, menu awal menyediakan tautan **Yuk Bikin POT** dan pilihan **Compiler Data POT**. Tema default menggunakan palet warna sejuk; preferensi mode terang/gelap dapat diubah melalui tombol tema.

## Database CMDB

Deployment Docker menyediakan API CMDB terautentikasi dan menyimpan file pada named volume persisten `cmdb_data` di `/data/cmdb`. File tetap divalidasi dan diproses di browser; salinan sumber hanya dikirim ke server ketika tombol **Simpan CMDB** atau **Update CMDB** digunakan. Tombol **Muat Server** mengambil salinan tersimpan. IndexedDB tetap digunakan sebagai cache browser lokal.

Atur token rahasia sebelum menjalankan deployment:

```sh
export CMDB_API_TOKEN="$(openssl rand -hex 32)"
sh deploy.sh
```

Masukkan token yang sama pada kontrol server di kartu CMDB. Token hanya disimpan di memori halaman dan tidak dimasukkan ke repository. **Simpan CMDB** membuat salinan server pertama; **Update CMDB** memvalidasi file yang dipilih di browser lalu mengganti salinan server; **Muat Server** memulihkan file tersebut ke aplikasi. Batas upload adalah 100 MB. Data server bertahan selama volume `cmdb_data` dipertahankan.

Layanan web di-bind ke `127.0.0.1` secara default dan port dapat diubah, misalnya `PORT=9000 sh deploy.sh`. Untuk akses jaringan, tempatkan di belakang reverse proxy HTTPS dan batasi akses ke pengguna/jaringan yang berwenang. Jangan kirim token melalui koneksi HTTP publik. GitHub Pages bersifat static dan tidak menyediakan API/penyimpanan CMDB server; fitur server hanya tersedia pada deployment Docker ini.

## Pencocokan site

Sheet `bcp_segmentation_export_sheet` menggunakan daftar Site ID dari POT sebagai dasar. Kolom `Impact Category` yang kosong diisi `Fully Site Down`. Site yang tidak ditemukan di file Segmentation tetap disertakan dengan status `UP`; data Segmentation dan CMDB dicocokkan menggunakan Site ID yang dinormalisasi. Nilai POT/Segmentation yang sudah terisi dipertahankan, sedangkan kolom kosong yang tersedia di CMDB dilengkapi dari CMDB.

GDOC mengambil Site ID langsung dari kolom F (`Site Id`) seluruh baris output segmentation, tanpa terpengaruh filter tampilan. Setiap Site ID unik dipertahankan (menggunakan baris POT terbaru untuk ID yang berulang); detail yang kosong dilengkapi dari POT, Segmentation, atau kolom terkait di CMDB. Status GDOC dipasangkan kembali menggunakan Site ID yang sama dan mengikuti nilai `UP`/`DOWN` eksplisit pada `Remark` output segmentation; jika tidak ada, status output segmentation dan hasil pencocokan Segmentation menjadi fallback. Kolom `Super Sites` GDOC menggunakan nilai sumber yang tersedia terlebih dahulu, lalu mengambil `Program` dari CMDB sebagai fallback sebelum `Is Hub`.

Perhitungan site down, PG, dan ambang MC >25% hanya menghitung status `DOWN`. `INCIDENT SUMMARY` mengelompokkan site `DOWN` berdasarkan `MC Cluster` di CMDB, lalu mengakumulasi jumlah site dan denominator CMDB dari seluruh MC terdampak; jumlah kategori Big/Medium dan alarm juga ditampilkan lintas MC terdampak. Untuk Site Name, MC Cluster, Hub Type, Traffic Class, P0, dan VIP, nilai CMDB per Site ID diprioritaskan; output GDOC menjadi fallback jika nilai CMDB kosong. Detail `Transport Outage` diambil dari CMDB berdasarkan Site ID, dengan Segmentation sebagai fallback, dan ditampilkan di antara BTS Name dan status genset, misalnya `08TIS0017_SELUMA_TIMUR_MT (TRANSPORT OUTAGE) (NO DG)`. Kategori `Super Site` dibandingkan dengan kolom `Program` CMDB per Site ID; nilai kosong, `Istimewa`, atau penanda negatif diabaikan.