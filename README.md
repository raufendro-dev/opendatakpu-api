# OPEN DATA KPU API

Open Data KPU API adalah API unofficial untuk mendapatkan data seputar pemilu berbentuk json. Open Data KPU API mengambil data yang terdapat pada halaman http://opendata.kpu.go.id

### Run
#### Rekapitulasi Daftar Pemilih Tetap (DPT) Dalam Negeri Pemilu Tahun 2024
Live link sample : http://103.193.178.139:6969/opendatakpu/dalamnegeri?kota=

Catatan : 
- Jika ingin menampilkan seluruh kota cukup kosongkan parameter kota seperti sample diatas
- Jika ingin menampilkan kota tertentu bisa diisi seperti berikut http://103.193.178.139:6969/opendatakpu/dalamnegeri?kota=sleman


## Cara penggunaan

### Install Library Python

Install terlebih dahulu library python berikut
- Flask
- openpyxl

Atau dapat menggunakan venv yang terdapat pada repository ini

### Run WebServer
Jalankan file web
- python3 web.py
- buka pada browser, contoh : http://localhost:6969/opendatakpu/dalamnegeri?kota=

### Catatan
- Link pada browser bisa menggunakan ip atau domain anda
- Port bisa diubah di web.py
- File logpengunjung.log adalah log dari pencarian user. Anda bisa menggunakannya jika dibutuhkan