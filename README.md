# POT Compile

## Menjalankan deployment

Pastikan Docker Engine dan Docker Compose plugin tersedia, lalu jalankan:

```sh
sh deploy.sh
```

Script membuat folder `dist/`, menyalin aplikasi, membangun image Nginx, dan menjalankan layanan. Buka `http://localhost:8080`. Port dapat diubah, misalnya `PORT=9000 sh deploy.sh`.

## Database CMDB

Saat file CMDB diunggah, aplikasi memvalidasi sheet dan kolom `Site Id`, membuat cache chunk untuk kalkulasi, lalu menyimpan file sumber ke IndexedDB browser. Sumber akan dipulihkan setelah refresh dan dapat dihapus melalui tombol **Clear**. Data tetap berada di browser/perangkat pengguna; deployment statis ini tidak mengirim CMDB ke server. Akses melalui `localhost` atau HTTPS agar IndexedDB tersedia.