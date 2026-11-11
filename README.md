# POT Compile

## Menjalankan deployment

Pastikan Docker Engine dan Docker Compose plugin tersedia, lalu jalankan:

```sh
sh deploy.sh
```

Script membangun image Nginx dan Python API dari source terbaru, lalu menjalankan layanan web dan API. Buka `http://localhost:8080`. Port dapat diubah, misalnya `PORT=9000 sh deploy.sh`.

Saat halaman dibuka, menu awal menyediakan tautan **Yuk Bikin POT**, **POT Automated**, dan **Excel Compiler**. Tema default menggunakan palet warna sejuk; preferensi mode terang/gelap dapat diubah melalui tombol tema.

## Excel Compiler

Pilih **Excel Compiler** pada menu awal untuk menggabungkan 2 sampai 20 file `.xlsx`, `.xls`, atau `.csv` menjadi satu sheet. Sertakan setidaknya satu `.xlsx` sebagai template agar workbook hasil mempertahankan sheet dan formatnya; baris dari file lain dipetakan berdasarkan header (tidak peka kapitalisasi dan spasi) dan urutan kolom template. Header yang berbeda, duplikat, atau kosong ditolak agar nilai tidak bergeser. Batas total upload adalah 500 MiB. Upload dan workbook dibaca/ditulis bertahap melalui file sementara agar penggunaan memori tetap terkendali.

Fitur ini menggunakan Python API terbuka pada deployment Docker. Saat halaman Excel Compiler dibuka, pengecekan kesehatan API otomatis membangunkan service Python yang sedang tidur, lalu mencoba ulang selama proses startup. GitHub Pages bersifat statis dan tidak menjalankan API Python, sehingga Excel Compiler memerlukan deployment server.

## Database CMDB

API membuat tabel SQLite `cmdb_files` secara otomatis saat pertama kali dijalankan dan menyimpan satu file sumber CMDB pada named volume Docker `cmdb_data`. File CMDB yang telah tersimpan pada format server sebelumnya akan dimigrasikan ke SQLite saat startup. Tidak diperlukan service atau proses database terpisah. Gunakan **Simpan CMDB** untuk penyimpanan pertama, **Update Server** untuk mengganti file, dan **Muat Server** untuk memulihkan salinan server. Cache browser di IndexedDB tetap tersedia untuk pemulihan lokal.

Setelah CMDB di-chunk untuk kompilasi, API dapat menyimpan cache site sementara pada PostgreSQL. Fitur ini opsional untuk deployment lokal dan otomatis aktif bila `DATABASE_URL` tersedia. API membuat tabel `cmdb_staging_sessions` dan `cmdb_staging_chunks` saat startup; tiap sesi dibatasi 256 MiB. Data sesi dihapus setelah proses kompilasi berhasil maupun gagal; jika proses terputus, sesi kedaluwarsa setelah 24 jam secara default dan dibersihkan setiap 15 menit. Atur `CMDB_STAGING_TTL_SECONDS` ke nilai 60–2592000 untuk mengubah masa kedaluwarsa. Penyimpanan staging bukan pengganti cache kompilasi browser atau SQLite untuk file CMDB.

Jalankan `sh deploy.sh` untuk memulai layanan. Layanan web di-bind ke `127.0.0.1` secara default dan port dapat diubah, misalnya `PORT=9000 sh deploy.sh`. Endpoint API tidak memiliki autentikasi, jadi jangan paparkan ke jaringan publik tanpa proteksi akses pada reverse proxy. Untuk akses jaringan, gunakan HTTPS dan batasi akses ke pengguna/jaringan yang berwenang. GitHub Pages bersifat statis dan tidak menyediakan API kompilasi atau penyimpanan CMDB.

## Deploy dan auto-wake di Railway

Railway tidak menjalankan `sh deploy.sh` dari browser. Sebagai gantinya, deploy web dan API sebagai dua service dari repository yang sama:

1. Buat service PostgreSQL dan service API dengan Dockerfile `Dockerfile.api`. Tambahkan volume persisten pada API di mount path `/data`, lalu atur `CMDB_DATABASE_PATH=/data/cmdb.sqlite3`, `PORT=8000`, `RAILWAY_RUN_UID=0`, dan `DATABASE_URL` sebagai reference ke URL koneksi PostgreSQL (contoh Railway: `${{Postgres.DATABASE_URL}}`, sesuaikan nama service PostgreSQL). API akan membuat tabel staging otomatis; jangan menempelkan URL atau kredensial database ke source code.
2. Pada service API, aktifkan **Serverless** di pengaturan deploy Railway dan atur healthcheck path ke `/api/health`. Gunakan satu replica untuk service API yang memakai volume SQLite; volume mempertahankan database ketika service tidur.
3. Buat service web dengan Dockerfile `Dockerfile`. Tambahkan public domain untuk service ini, atur target `PORT=80`, lalu atur `API_HOST` ke private domain API, misalnya `excel-api.railway.internal`, dan `API_PORT=8000`.
4. Pastikan kedua service berada di project dan environment Railway yang sama. Nginx meneruskan `/api/` melalui private network; membuka menu **Excel Compiler** memanggil healthcheck yang membangunkan API bila sedang tidur.

Railway menyebut request pertama setelah tidur dapat mengalami cold start atau sesekali merespons 502; halaman Excel Compiler otomatis mencoba ulang. Healthcheck Railway hanya digunakan untuk validasi deployment, bukan pemantauan berkelanjutan. API tidak memiliki autentikasi, jadi jangan publikasikan domain API secara langsung; aksesnya lewat proxy web tetap terbuka dan perlu dilindungi di lapisan jaringan bila aplikasi dipakai untuk data sensitif.

## Pencocokan site

POT Automated hanya meminta file POT dan CMDB; tidak membaca atau memakai file Segmentation terpisah. Sheet `bcp_segmentation_export_sheet` dibentuk dari Site Down List di POT, lalu kolom yang kosong dilengkapi dari CMDB berdasarkan Site ID. Header dicocokkan tanpa membedakan kapitalisasi atau tanda baca, dengan alias kolom umum untuk nama Site ID, nama site, MC, hub, status, VIP, genset, dan atribut terkait. Nilai POT diprioritaskan; CMDB menjadi fallback. Kolom `Impact Category` yang kosong diisi `Fully Site Down`.

Status `UP` atau `DOWN` dipakai jika tersedia pada baris POT. Karena Site Down List secara default memuat site yang sedang down, baris tanpa kolom status eksplisit diperlakukan sebagai `DOWN`. Untuk GDOC, snapshot `Key` pada sheet `bcp_segmentation_export_sheet` dibaca dari tanggal dan waktu di akhir BCP index, termasuk jam, menit, dan detik (contoh `BCP-20261110-0126_2026-11-11 00:49:12`). Status ditentukan dengan menggabungkan daftar site di seluruh Key dan membandingkannya dengan snapshot terbaru: jika Site ID hanya ada di Key yang lebih lama dan tidak ada pada snapshot terbaru, status akhirnya menjadi `UP`; site yang ada pada snapshot terbaru menggunakan status baris POT. GDOC mengambil setiap Site ID nonkosong dari POT tanpa terpengaruh filter tampilan dan menggunakan baris POT terbaru untuk Site ID berulang. Sebelum workbook dibuat, validasi memastikan semua Site ID sumber masuk ke GDOC dan setiap status valid; ketidaklengkapan menghasilkan error, bukan laporan yang tampak berhasil tetapi kehilangan data. Kolom `Super Sites` GDOC menggunakan nilai POT yang tersedia, lalu mengambil `Program` dari CMDB sebelum `Is Hub`.

Perhitungan site down, PG, dan ambang MC >25% hanya menghitung status `DOWN`. `INCIDENT SUMMARY` mengelompokkan site `DOWN` berdasarkan `MC Cluster` di CMDB, lalu mengakumulasi jumlah site dan denominator CMDB dari seluruh MC terdampak; jumlah kategori Big/Medium dan alarm juga ditampilkan lintas MC terdampak. Status GDOC mengikuti snapshot Key terbaru hingga presisi detik: site yang tidak lagi tercantum pada snapshot terbaru dianggap `UP`, sementara site yang tercantum memakai status eksplisit POT atau default `DOWN`. Untuk Site Name, MC Cluster, Hub Type, Traffic Class, P0, dan VIP pada ringkasan, nilai CMDB per Site ID menjadi prioritas; output GDOC menjadi fallback jika nilai CMDB kosong. Detail `Transport Outage` diambil dari output `bcp_segmentation_export_sheet` berdasarkan Site ID, lalu CMDB menjadi fallback; detail ditampilkan di antara BTS Name dan status genset, misalnya `08TIS0017_SELUMA_TIMUR_MT (TRANSPORT OUTAGE) (NO DG)`. Kategori `Super Site` dibandingkan dengan kolom `Program` CMDB per Site ID; nilai kosong, `Istimewa`, atau penanda negatif diabaikan.