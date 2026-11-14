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
