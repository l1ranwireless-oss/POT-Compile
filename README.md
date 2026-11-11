# POT Compile

## Menjalankan deployment

Pastikan Docker Engine dan Docker Compose plugin tersedia, lalu jalankan:

```sh
sh deploy.sh
```

Script membangun image Nginx dan Python API dari source terbaru, lalu menjalankan layanan web dan API. Buka `http://localhost:8080`. Port dapat diubah, misalnya `PORT=9000 sh deploy.sh`.

Saat halaman dibuka, menu awal menyediakan tautan **Yuk Bikin POT**, **POT Automated**, dan **Excel Compiler**. Tema default menggunakan palet warna sejuk; preferensi mode terang/gelap dapat diubah melalui tombol tema. Inter dan Space Grotesk dimuat dari Google Fonts CDN untuk tipografi; jika CDN tidak tersedia, font sistem digunakan sebagai fallback.

Dashboard POT Automated menampilkan sapaan sesuai waktu setempat, cuaca terkini Jakarta, dan headline telekomunikasi, teknologi, serta finansial Indonesia. Tombol **Lokasi saya** meminta izin browser secara eksplisit; tanpa izin cuaca tetap memakai Jakarta. Cuaca diambil dari Open-Meteo dan berita dari Google News RSS melalui API backend, dilengkapi atribusi dan tautan ke penerbit. Kedua panel diperbarui berkala; bila layanan online tidak tersedia, UI menampilkan status kegagalan dan tools utama tetap dapat digunakan.

## Excel Compiler

Pilih **Excel Compiler** pada menu awal untuk menggabungkan file `.xlsx`, `.xlsm`, `.xls`, atau `.csv` menjadi satu sheet. Pemrosesan berjalan lokal di browser; file tidak diunggah ke server dan fitur dapat digunakan pada GitHub Pages. Anda dapat memilih semua sheet atau nama sheet tertentu, mengatur baris header (atau memakai nama kolom otomatis), menambahkan kolom sumber file/nama sheet, melewati sheet kosong, menambah file dari folder, serta mengubah urutan input. Header kolom yang sama digabungkan, sedangkan kolom yang tidak ada pada suatu sheet akan dibiarkan kosong. Hasil diunduh sebagai file XLSX dengan lebar kolom otomatis.

## Database CMDB

Jika `CMDB_DATABASE_URL` atau `DATABASE_URL` tersedia, PostgreSQL menjadi penyimpanan utama satu file sumber CMDB; API membuat tabel `cmdb_source_files` dan memindahkan CMDB SQLite lama ke PostgreSQL saat startup bila tabel PostgreSQL masih kosong. Simpan dan update berikutnya menulis langsung ke PostgreSQL, sementara **Muat Server** membaca dari database tersebut. Tanpa URL PostgreSQL, deployment lokal tetap memakai SQLite pada named volume Docker `cmdb_data`. Cache browser di IndexedDB tetap tersedia untuk pemulihan lokal. Endpoint `/api/health` memeriksa koneksi PostgreSQL saat PostgreSQL digunakan.

Setelah CMDB di-chunk untuk kompilasi, API juga menyimpan cache site sementara pada PostgreSQL bila `CMDB_STAGING_DATABASE_URL` atau `DATABASE_URL` tersedia. API membuat tabel `cmdb_staging_sessions` dan `cmdb_staging_chunks` saat startup; tiap sesi dibatasi 256 MiB. Data sesi dihapus setelah proses kompilasi berhasil maupun gagal; jika proses terputus, sesi kedaluwarsa setelah 24 jam secara default dan dibersihkan setiap 15 menit. Atur `CMDB_STAGING_TTL_SECONDS` ke nilai 60–2592000 untuk mengubah masa kedaluwarsa. Staging bersifat sementara dan terpisah dari file CMDB permanen.

Jalankan `sh deploy.sh` untuk memulai layanan. Layanan web di-bind ke `127.0.0.1` secara default dan port dapat diubah, misalnya `PORT=9000 sh deploy.sh`. Endpoint API tidak memiliki autentikasi, jadi jangan paparkan ke jaringan publik tanpa proteksi akses pada reverse proxy. Untuk akses jaringan, gunakan HTTPS dan batasi akses ke pengguna/jaringan yang berwenang. Excel Compiler tetap dapat digunakan pada GitHub Pages karena file diproses di browser; GitHub Pages tidak menyediakan API untuk POT Automated atau penyimpanan CMDB server.

## Deploy dan auto-wake di Railway

Railway tidak menjalankan `sh deploy.sh` dari browser. Sebagai gantinya, deploy web dan API sebagai dua service dari repository yang sama:

1. Buat service PostgreSQL dan service API dengan Dockerfile `Dockerfile.api`. Atur `PORT=8000`, `RAILWAY_RUN_UID=0`, dan `DATABASE_URL` sebagai reference ke URL koneksi PostgreSQL (contoh Railway: `${{Postgres.DATABASE_URL}}`, sesuaikan nama service PostgreSQL). `DATABASE_URL` mengaktifkan penyimpanan CMDB permanen sekaligus staging; API membuat tabel dan memigrasikan CMDB SQLite lama bila tabel Postgres masih kosong. Jangan menempelkan URL atau kredensial database ke source code.
2. Pada service API, aktifkan **Serverless** di pengaturan deploy Railway dan atur healthcheck path ke `/api/health`. Bila menggunakan `CMDB_DATABASE_URL` atau `CMDB_STAGING_DATABASE_URL` yang berbeda, atur keduanya ke URL service PostgreSQL yang sesuai.
3. Buat service web dengan Dockerfile `Dockerfile`. Tambahkan public domain untuk service ini, atur target `PORT=80`, lalu atur `API_HOST` ke private domain API, misalnya `excel-api.railway.internal`, dan `API_PORT=8000`.
4. Pastikan kedua service berada di project dan environment Railway yang sama. Nginx meneruskan `/api/` melalui private network; halaman POT Automated menggunakan API untuk memuat cuaca dan berita.

Railway menyebut request pertama setelah tidur dapat mengalami cold start atau sesekali merespons 502; halaman yang memakai API dapat memerlukan percobaan ulang. Healthcheck Railway hanya digunakan untuk validasi deployment, bukan pemantauan berkelanjutan. API tidak memiliki autentikasi, jadi jangan publikasikan domain API secara langsung; aksesnya lewat proxy web tetap terbuka dan perlu dilindungi di lapisan jaringan bila aplikasi dipakai untuk data sensitif.

## Pencocokan site

POT Automated hanya meminta file POT dan CMDB; tidak membaca atau memakai file Segmentation terpisah. Sheet `bcp_segmentation_export_sheet` dibentuk dari Site Down List di POT, lalu kolom yang kosong dilengkapi dari CMDB berdasarkan Site ID. Header dicocokkan tanpa membedakan kapitalisasi atau tanda baca, dengan alias kolom umum untuk nama Site ID, nama site, MC, hub, status, VIP, genset, dan atribut terkait. Nilai POT diprioritaskan; CMDB menjadi fallback. Kolom `Impact Category` yang kosong diisi `Fully Site Down`.

Status `UP` atau `DOWN` dipakai jika tersedia pada baris POT. Karena Site Down List secara default memuat site yang sedang down, baris tanpa kolom status eksplisit diperlakukan sebagai `DOWN`. Untuk GDOC, snapshot `Key` pada sheet `bcp_segmentation_export_sheet` dibaca dari tanggal dan waktu di akhir BCP index, termasuk jam, menit, dan detik (contoh `BCP-20261110-0126_2026-11-11 00:49:12`). Status ditentukan dengan menggabungkan daftar site di seluruh Key dan membandingkannya dengan snapshot terbaru: jika Site ID hanya ada di Key yang lebih lama dan tidak ada pada snapshot terbaru, status akhirnya menjadi `UP`; site yang ada pada snapshot terbaru menggunakan status baris POT. GDOC mengambil setiap Site ID nonkosong dari POT tanpa terpengaruh filter tampilan dan menggunakan baris POT terbaru untuk Site ID berulang. Sebelum workbook dibuat, validasi memastikan semua Site ID sumber masuk ke GDOC dan setiap status valid; ketidaklengkapan menghasilkan error, bukan laporan yang tampak berhasil tetapi kehilangan data. Kolom `Super Sites` GDOC menggunakan nilai POT yang tersedia, lalu mengambil `Program` dari CMDB sebelum `Is Hub`.

`INCIDENT SUMMARY` menggunakan status unik per Site ID dari GDOC sebagai sumber hitungan DOWN/UP, kategori, dan daftar site; angka total-down dari Summary Report hanya ditampilkan sebagai cross-check agar angka ringkasan tidak bertentangan dengan daftar site. Snapshot Key di seluruh data segmentation menentukan waktu mulai/update dan aging sampai detik. PG dihitung dari semua baris POT, bukan hanya baris yang sedang ditampilkan oleh filter segmentation. Ringkasan juga menampilkan kolom Alarm, Alarm Identifier, Power Alarm, Mains Failure, Router, dan DWDM dari baris segmentation terbaru per site, ditambah temuan ZIP; hitungan power alarm hanya memasukkan alarm listrik atau kecocokan ZIP, bukan alarm umum. Site ID dan MC ditampilkan pada daftar prioritas agar tiap angka dapat ditelusuri ke data. Perhitungan site down, PG, dan ambang MC >25% hanya menghitung status `DOWN`. `INCIDENT SUMMARY` mengelompokkan site `DOWN` berdasarkan `MC Cluster` di CMDB, lalu mengakumulasi jumlah site dan denominator CMDB dari seluruh MC terdampak; jumlah kategori Big/Medium dan alarm juga ditampilkan lintas MC terdampak. Status GDOC mengikuti snapshot Key terbaru hingga presisi detik: site yang tidak lagi tercantum pada snapshot terbaru dianggap `UP`, sementara site yang tercantum memakai status eksplisit POT atau default `DOWN`. Untuk Site Name, MC Cluster, Hub Type, Traffic Class, P0, dan VIP pada ringkasan, nilai CMDB per Site ID menjadi prioritas; output GDOC menjadi fallback jika nilai CMDB kosong. Detail `Transport Outage` diambil dari output `bcp_segmentation_export_sheet` berdasarkan Site ID, lalu CMDB menjadi fallback. Kategori `Super Site` dibandingkan dengan kolom `Program` CMDB per Site ID; nilai kosong, `Istimewa`, atau penanda negatif diabaikan.