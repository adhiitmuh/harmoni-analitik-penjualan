# Harmoni Analitik Penjualan

Analitik penjualan Olsera + Rekomendasi Stok. Satu file `index.html` (compat SDK, Chart.js). `generate.py` tidak dipakai lagi — edit `index.html` langsung.
Firebase project `harmoni-indonesia` (auth Portal, akses `apps.harmoni_analitik`; role admin = Owner/Admin, selain itu operator).
Aturan kerja lintas app: lihat `harmoni-indonesia/CLAUDE.md`.

## Data
- `penjualan` (Owner/Admin, per bulan) → ringkasan `penjualan_operator` (dibaca operator & Stok Manager).
- Upload stok Olsera (Toko/Gudang) → `settings/stok_fisik` + ikut menulis Stok Manager (`sinkronKeStokManager`, project harmoni-stok; tandai `modal_baru_olsera` bila modal naik >10%).
- `settings/modal_produk` (dari Stok Manager): modal (`MODAL_SM`, dipakai lebih dulu di `skuInfo`), supplier utama (`SUP_SKU`), lama kirim (`SUP_INFO`), harga per supplier (`h`).
- `settings/final_order` — Final Order (cloud; localStorage hanya cadangan).

## Rekomendasi Stok
- `rek_qty` = rata-rata N bulan (3/6/12) × faktor tren / musiman. `targetStok(s)` = rek_qty + laku/hari × lama kirim supplier → dipakai Rek. Stok & Perlu Beli.
- Filter: perlu beli / cukup / kosong & laku / hampir habis & laku / kosong / toko kosong gudang ada; filter kategori, tren, supplier.
- Excel/CSV mengikuti kolom yang tampil (panel Kolom).

## Uji
Di sandbox CDN Chart.js tidak termuat → skrip berhenti di bagian grafik; uji fungsi dengan mengekstrak potongan kode (lihat riwayat sesi).

## Barang Custom (`js/custom.js`, tab 🧵 Owner/Admin)
- Upload laporan Olsera **Item Penjualan berdasarkan Tanggal** (Excel; PDF juga dibaca lewat pdf.js — posisi kolom tetap). Item custom = Unit kosong & Modal 0 (nama diketik kasir). Jenis = teks dalam kurung di akhir nama.
- Simpan `penjualan_custom/{YYYY-MM}` {rows:[{t,no,k,p,i,j,q,r}], semua:{omzet,baris}} (Owner/Admin) + `penjualan_custom_ringkas/{YYYY-MM}` {j:[{k,q}]} untuk Stok Manager (kebutuhan bahan custom di Daftar Belanja). Upload ulang bulan sama = ganti.
- Upload boleh ZIP (JSZip). Dokumen bulan juga menyimpan `kanal` {Toko|Shopee|Tokopedia|TikTok: {omzet,qty,struk}} & `topKanal` (200 barang teratas per marketplace) dari prefiks no. pesanan (SP… Shopee, INV… Tokopedia) → tab 🛍️ Toko vs Shopee. Baris "Custom" di Top Produk bisa diklik → tab 🧵 periode sama.
- Top Produk punya kolom Stok Toko/Gudang (jumlah SKU produk+varian dari `settings/stok_fisik`).

## Multi-outlet (`js/outlet.js`) — Tahap 1
- Tiap cabang/titik punya akun Olsera sendiri. `?outlet=ID` = mode outlet: upload "Penjualan berdasarkan SKU" disimpan `outlet_penjualan/{outlet}__{bulan}` {outlet, bulan, rows} (bukan `penjualan`), tanpa localStorage, tanpa ringkasan operator; stok HQ tidak dipakai; tab Inventori/Rekomendasi/Custom/Shopee disembunyikan.
- Staf outlet = Portal `apps.harmoni_analitik.outlet` (dipaksa ke `?outlet=` miliknya, jadi admin untuk outletnya). Owner/Admin HQ (tanpa outlet) bisa pilih lokasi di period bar + tab 🏬 Outlet (perbandingan & kelola `outlet/{id}` {nama, tipe cabang|titik, jenis keluarga|mitra, pemilik, aktif} — tulis Owner).
- Ringkasan `outlet_ringkas/{outlet|hq}` {b:{'Jun-26':{omzet,profit,qty,baris}}} ditulis saat upload/data cloud berubah → perbandingan antar outlet tanpa memuat data mentah (siap ribuan outlet).
- Rules: `adminHQ()` = adminAnalitik tanpa outlet (penjualan, penjualan_custom*, penjualan_operator write); `stafOutlet()` hanya dokumen outletnya.
- Tahap berikutnya: stok per outlet + katalog SKU HQ (Tahap 2), pesanan outlet → HQ & belanja HQ gabungan (Tahap 3), tarik otomatis via API Olsera (Tahap 4).
