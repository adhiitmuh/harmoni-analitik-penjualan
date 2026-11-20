// Barang Custom — dari laporan Olsera "Item Penjualan berdasarkan Tanggal" (Excel atau PDF).
// Item custom = baris tanpa satuan (Unit Pengukuran kosong) dan modal 0 → nama yang diketik kasir.
// Simpan (Owner/Admin): penjualan_custom/{YYYY-MM} {rows, semua:{omzet,baris}} + ringkasan qty per jenis
// penjualan_custom_ringkas/{YYYY-MM} {j:[[jenis,qty]]} yang dibaca Stok Manager (kebutuhan bahan custom → Daftar Belanja).
// Jenis = teks dalam kurung di akhir nama, mis. "DEWI (US KECIL LIS MAGNET)" → "US KECIL LIS MAGNET"; tanpa kurung = nama item.
(function () {
  const BLN = { jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6, jul: 7, agu: 8, agt: 8, ags: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12 }
  const NAMA_BLN = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
  const pad = n => String(n).padStart(2, '0')
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
  const rp = n => 'Rp ' + Math.round(n || 0).toLocaleString('id-ID')
  const angka = v => { if (typeof v === 'number') return v; const s = String(v ?? '').replace(/IDR|Rp/gi, '').trim(); const neg = /^-|\(.*\)/.test(s); const d = s.replace(/[^\d,]/g, '').replace(',', '.'); return (neg ? -1 : 1) * (parseFloat(d) || 0) }
  // Isi kurung yang hanya ukuran (42, 15, XL, S-M…) bukan jenis → pakai nama di depan kurung
  const UKURAN = /^(\d{1,3}([.,]\d)?|[SML]|XS|XL|XXL|XXXL|\d?XL|ALL ?SIZE|[SML]\s*-\s*[SML]|NO\.?\s*\d+)$/
  const jenisDari = nama => { const n = String(nama || '').toUpperCase().replace(/\s+/g, ' ').trim(); const m = n.match(/\(([^()]+)\)\s*$/)
    if (m && !UKURAN.test(m[1].trim())) return m[1].trim()
    return (m ? n.slice(0, m.index) : n).trim() || n }
  // Kelompok besar (kata kunci) untuk ringkasan
  const KELOMPOK = [
    ['Papan nama', /\bUS\b|\bKL\b|KACA LAPIS|P\/HH|PRA\/HH|H\/|PAPAN|NAMA DADA|NAME ?TAG/],
    ['Logo / lokasi / lambang', /LOGO|LOKASI|LAMBANG|BADGE|EMBLEM/],
    ['Bordir langsung', /BORDIR/],
    ['Pin / medali', /\bPIN\b|MEDALI|FIBER/],
    ['Atribut / tanda', /ATRIBUT|TANDA|PANGKAT|TALI|SELEMPANG|KOKARDE|MONOGRAM/],
    ['Topi / dasi / sabuk', /TOPI|DASI|SABUK|KOPEL|PECI/],
    ['Pakaian', /BAJU|KAOS|KEMEJA|CELANA|ROK|JAKET|ROMPI|SERAGAM|PDH|PDL|SAFARI/],
  ]
  const kelompokDari = (nama, jenis) => { const t = (jenis + ' ' + nama).toUpperCase(); return (KELOMPOK.find(([, re]) => re.test(t)) || ['Lainnya'])[0] }
  const tglIso = v => {
    if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`
    if (typeof v === 'number' && v > 20000) { const d = new Date(Math.round((v - 25569) * 864e5)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}` }
    const s = String(v || '')
    let m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`
    m = s.match(/(\d{1,2})[-\s\/]([A-Za-z]{3})[A-Za-z]*[-\s\/]*(\d{4})/); if (m && BLN[m[2].toLowerCase()]) return `${m[3]}-${pad(BLN[m[2].toLowerCase()])}-${pad(m[1])}`
    m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`
    return ''
  }

  // ── Baca Excel ────────────────────────────────────────────────────
  function bacaExcel(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true })
    for (const n of wb.SheetNames) {
      const a = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: true })
      const hi = a.findIndex(r => r.some(c => /^item$|nama item|item name/i.test(String(c).trim())) && r.some(c => /qty|jumlah/i.test(String(c))))
      if (hi < 0) continue
      const H = a[hi].map(c => String(c).replace(/\s+/g, ' ').trim().toLowerCase())
      // Nama kolom persis dulu (Excel Olsera: item name, item uom, amount, total cost, item group…), baru pola longgar
      const kol = (persis, ...re) => { for (const n of persis) { const i = H.indexOf(n); if (i >= 0) return i } for (const r of re) { const i = H.findIndex(h => r.test(h)); if (i >= 0) return i } return -1 }
      const K = { no: kol(['order no', 'no. pesanan', 'no pesanan'], /no\.? ?pesanan|order no|no\.? ?struk/), tgl: kol(['order date', 'tanggal jual', 'tanggal'], /tanggal/, /date/),
        plg: kol(['customer name', 'pelanggan'], /pelanggan/), oleh: kol(['sales name', 'serve by', 'penjualan oleh'], /penjualan oleh/),
        item: kol(['item name', 'item', 'nama item'], /^item$/), qty: kol(['qty', 'jumlah'], /^qty$/),
        unit: kol(['item uom', 'uom', 'unit pengukuran', 'satuan'], /^unit pengukuran|^satuan/), total: kol(['amount', 'total pesanan', 'subtotal', 'total'], /^total pesanan/),
        modal: kol(['total cost', 'modal produk', 'modal', 'cost perunit'], /^modal/), grup: kol(['item group', 'kategori'], /group/), sumber: kol(['order source'], /order source/) }
      if (K.item < 0 || K.qty < 0) continue
      const rows = []
      for (const r of a.slice(hi + 1)) {
        const item = String(r[K.item] || '').replace(/\s+/g, ' ').trim(); if (!item) continue
        rows.push({ no: String(r[K.no] ?? '').trim(), tgl: tglIso(r[K.tgl]), plg: String(r[K.plg] ?? '').trim(), oleh: String(r[K.oleh] ?? '').trim(), item,
          qty: angka(r[K.qty]), unit: K.unit >= 0 ? String(r[K.unit] ?? '').trim() : '', total: angka(r[K.total]), modal: K.modal >= 0 ? angka(r[K.modal]) : 0,
          grup: K.grup >= 0 ? String(r[K.grup] ?? '').trim().toUpperCase() : '', sumber: K.sumber >= 0 ? String(r[K.sumber] ?? '').trim().toUpperCase() : '' })
      }
      return rows
    }
    throw new Error('Kolom "Item" & "Qty" tidak ditemukan — pastikan ini laporan Olsera "Item Penjualan berdasarkan Tanggal"')
  }

  // ── Baca PDF (posisi teks; kolom laporan Olsera tetap) ─────────────────
  async function muatPdfJs() {
    if (window.pdfjsLib) return window.pdfjsLib
    const B = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/'
    await new Promise((ok, gagal) => { const s = document.createElement('script'); s.src = B + 'pdf.min.js'; s.onload = ok; s.onerror = () => gagal(new Error('Gagal memuat pembaca PDF — cek internet')); document.head.appendChild(s) })
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = B + 'pdf.worker.min.js'
    return window.pdfjsLib
  }
  async function bacaPdf(buf, st) {
    const lib = await muatPdfJs(), pdf = await lib.getDocument({ data: new Uint8Array(buf) }).promise, rows = []
    for (let pg = 1; pg <= pdf.numPages; pg++) {
      if (st && pg % 20 === 0) st.textContent = `⏳ Membaca PDF halaman ${pg}/${pdf.numPages}...`
      const page = await pdf.getPage(pg), H = page.getViewport({ scale: 1 }).height, tc = await page.getTextContent()
      const ws = []
      tc.items.forEach(it => { const s = it.str || ''; if (!s.trim()) return
        const x0 = it.transform[4], y = H - it.transform[5], w = it.width || 0, per = s.length ? w / s.length : 0
        let i = 0; s.split(/(\s+)/).forEach(t => { if (t.trim()) ws.push({ x: x0 + i * per, r: x0 + (i + t.length) * per, y, t }); i += t.length }) })
      const anchors = ws.filter(w => w.x < 40 && /^[0-9A-Z]{12,}$/.test(w.t)).sort((a, b) => a.y - b.y)
      anchors.forEach((a, i) => {
        const y0 = a.y - 3, y1 = i + 1 < anchors.length ? anchors[i + 1].y - 3 : 1e9
        const isi = ws.filter(w => w.y >= y0 && w.y < y1).sort((p, q) => p.y - q.y || p.x - q.x)
        const kol = (x0, x1) => isi.filter(w => w.x >= x0 && w.x < x1).map(w => w.t).join(' ').trim()
        const qtyW = isi.find(w => Math.abs(w.y - a.y) < 2 && /^[\d.]+$/.test(w.t) && w.r > 426 && w.r < 432 && w.x > 395)
        const num = isi.filter(w => w.x >= 480 && /^-?[\d.,]+$/.test(w.t))
        const uang = isi.filter(w => w.x >= 480 && w.t === 'IDR').sort((p, q) => p.x - q.x).map(w => {
          const kanan = num.find(n => Math.abs(n.y - w.y) < 2 && n.x > w.x && n.x - w.x < 25)
          const bawah = num.filter(n => n.y > w.y + 2 && n.y < w.y + 14 && n.x >= w.x - 30 && n.x < w.x + 45).sort((p, q) => p.y - q.y)[0]
          return angka((kanan || bawah || {}).t || 0) })
        rows.push({ no: a.t, tgl: tglIso(kol(120, 165)), plg: kol(165, 240), oleh: kol(240, 286), item: isi.filter(w => w.x >= 300 && w.x < 436 && w !== qtyW).map(w => w.t).join(' ').replace(/\s+/g, ' ').trim(),
          qty: qtyW ? angka(qtyW.t) : 0, unit: kol(436, 480), total: uang[0] || 0, modal: uang[2] || 0 })
      })
    }
    if (!rows.length) throw new Error('Tidak ada baris terbaca — pastikan ini PDF laporan Olsera "Item Penjualan berdasarkan Tanggal"')
    return rows
  }

  // Kelompokkan per bulan → dokumen yang disimpan
  const kanalDari = no => /^SP/i.test(String(no || '').trim()) ? 'Shopee' : /^INV/i.test(String(no || '').trim()) ? 'Tokopedia' : /^(TT|TIKTOK)/i.test(String(no || '').trim()) ? 'TikTok' : 'Toko'
  function perBulan(rows) {
    const B = {}
    rows.forEach(r => { const b = (r.tgl || '').slice(0, 7); if (!/^\d{4}-\d{2}$/.test(b)) return
      const x = B[b] = B[b] || { rows: [], omzet: 0, baris: 0, kanal: {}, itemKanal: {} }
      x.omzet += r.total || 0; x.baris++
      // Sumber pesanan dari nomor order Olsera (SP… = Shopee, INV… = Tokopedia, lainnya = kasir toko)
      const kn = r.sumber ? (/SHOPEE/.test(r.sumber) ? 'Shopee' : /TOKOPEDIA/.test(r.sumber) ? 'Tokopedia' : /TIKTOK/.test(r.sumber) ? 'TikTok' : /POINT OF SALE|POS/.test(r.sumber) ? 'Toko' : r.sumber.charAt(0) + r.sumber.slice(1).toLowerCase()) : kanalDari(r.no), k = x.kanal[kn] = x.kanal[kn] || { omzet: 0, qty: 0, struk: new Set() }
      k.omzet += r.total || 0; k.qty += r.qty || 0; if (r.no) k.struk.add(r.no)
      if (kn !== 'Toko') { const ik = x.itemKanal[kn] = x.itemKanal[kn] || {}, nm = String(r.item || '').toUpperCase().replace(/\s+/g, ' ').trim(), e = ik[nm] = ik[nm] || { i: nm, q: 0, r: 0 }; e.q += r.qty || 0; e.r += r.total || 0 }
      if (r.item && (r.grup ? r.grup === 'CUSTOM' : (!r.unit && !(r.modal > 0)))) x.rows.push({ t: r.tgl, no: r.no, k: r.oleh, p: r.plg, i: r.item, j: jenisDari(r.item), q: r.qty || 0, r: r.total || 0 }) })
    return B
  }

  async function simpan(B, nama) {
    const db = _app.firestore(), now = firebase.firestore.FieldValue.serverTimestamp()
    for (const [b, x] of Object.entries(B)) {
      let rows = x.rows
      if (JSON.stringify(rows).length > 900000) rows = rows.map(({ p, ...s }) => s)   // batas 1 MB dokumen: buang nama pelanggan
      const ring = {}; rows.forEach(r => { ring[r.j] = (ring[r.j] || 0) + (r.q || 0) })
      const kanal = {}; Object.entries(x.kanal).forEach(([k, v]) => { kanal[k] = { omzet: v.omzet, qty: v.qty, struk: v.struk.size } })
      const topKanal = {}; Object.entries(x.itemKanal).forEach(([k, m]) => { topKanal[k] = Object.values(m).sort((a, c) => c.r - a.r).slice(0, 200) })
      const isi = { rows, semua: { omzet: x.omzet, baris: x.baris }, kanal, topKanal, sumber: nama, updated_at: now, updated_by: _currentUserName || '' }
      if (_OUTLET) { await db.collection('outlet_custom').doc(_OUTLET + '__' + b).set({ ...isi, outlet: _OUTLET, bulan: b }); continue }   // custom outlet (laporan akun Olsera outlet)
      await db.collection('penjualan_custom').doc(b).set(isi)
      await db.collection('penjualan_custom_ringkas').doc(b).set({ j: Object.entries(ring).map(([k, q]) => ({ k, q })), updated_at: now })
    }
  }

  // ── Data & tampilan ───────────────────────────────────────────────
  let DATA = null   // { 'YYYY-MM': {rows, semua} }
  let KATA_STOK = []   // kata kunci barang custom yang distok/disiapkan (settings/custom_stok, diatur Owner)
  window._cuData = () => DATA; window._cuKataStok = () => KATA_STOK
  let DARI_TOP = null   // daftar 'YYYY-MM' periode Top Produk saat baris "Custom" diklik
  // Label bulan Analitik ('Jan-25', 'Agt-26', 'Mei-26') → 'YYYY-MM'
  const keIso = l => { const m = String(l).match(/^([A-Za-z]+)[-\s]?(\d{2,4})$/); if (!m) return ''; const b = BLN[m[1].slice(0, 3).toLowerCase()]; const y = +m[2] < 100 ? 2000 + +m[2] : +m[2]; return b ? `${y}-${pad(b)}` : '' }
  async function muat(paksa) {
    if (DATA && !paksa) return DATA
    DATA = {}
    try { await _fbReady; const db = _app.firestore()
      if (_OUTLET) (await db.collection('outlet_custom').where('outlet', '==', _OUTLET).get()).forEach(d => { DATA[d.data().bulan || d.id.split('__').pop()] = d.data() })
      else (await db.collection('penjualan_custom').get()).forEach(d => { DATA[d.id] = d.data() })
      // Jenis dihitung ulang saat dimuat (aturan jenis bisa diperbaiki tanpa upload ulang)
      Object.values(DATA).forEach(d => (d.rows || []).forEach(r => { r.j = jenisDari(r.i) }))
      try { const k = await db.collection('settings').doc('custom_stok').get(); KATA_STOK = k.exists ? (k.data().kata || []) : [] } catch (e) { KATA_STOK = [] } }
    catch (e) { console.warn('penjualan_custom', e); document.getElementById('cuStatus').textContent = '⚠️ Data custom gagal dimuat: ' + (e.code === 'permission-denied' ? 'akun ini tidak punya akses' : e.message) }
    return DATA
  }
  const labelBln = b => NAMA_BLN[+b.slice(5)] + ' ' + b.slice(0, 4)

  function pilihanPeriode() {
    const sel = document.getElementById('cuPeriode'), v = sel.value, bln = Object.keys(DATA).sort()
    const thn = [...new Set(bln.map(b => b.slice(0, 4)))]
    sel.innerHTML = (DARI_TOP ? `<option value="__top">Sesuai periode Top Produk (${DARI_TOP.length} bulan)</option>` : '') + '<option value="">Semua bulan</option>' + thn.map(t => `<option value="${t}">Tahun ${t}</option>`).join('') + bln.slice().reverse().map(b => `<option value="${b}">${labelBln(b)}</option>`).join('')
    if ([...sel.options].some(o => o.value === v)) sel.value = v
  }
  function terpilih() {
    const per = document.getElementById('cuPeriode').value, q = document.getElementById('cuCari').value.trim().toUpperCase(), kel = document.getElementById('cuKelompok').value
    const bln = Object.keys(DATA).sort().filter(b => per === '__top' ? DARI_TOP.includes(b) : !per || b.startsWith(per))
    const rows = []; bln.forEach(b => (DATA[b].rows || []).forEach(r => { const g = kelompokDari(r.i, r.j); if ((!q || (r.i + ' ' + r.j).toUpperCase().includes(q)) && (!kel || g === kel)) rows.push({ ...r, b, g }) }))
    return { bln, rows }
  }
  function render() {
    if (!DATA) return
    pilihanPeriode()
    const kosong = !Object.keys(DATA).length
    document.getElementById('cuIsi').style.display = kosong ? 'none' : ''
    document.getElementById('cuKosong').style.display = kosong ? '' : 'none'
    renderStokCustom()
    if (kosong) return
    const { bln, rows } = terpilih(), sum = (a, k) => a.reduce((t, x) => t + (x[k] || 0), 0)
    const omzetToko = bln.reduce((t, b) => t + (DATA[b].semua?.omzet || 0), 0)
    const kartu = (l, v, s = '') => `<div class="kpi-card"><div class="kpi-label">${l}</div><div class="kpi-value" style="font-size:1.15rem">${v}</div>${s ? `<div class="kpi-sub">${s}</div>` : ''}</div>`
    document.getElementById('cuKartu').innerHTML = kartu('Omzet custom', rp(sum(rows, 'r')), omzetToko ? `${(100 * sum(rows, 'r') / omzetToko).toFixed(1)}% dari omzet toko` : '')
      + kartu('Qty custom', sum(rows, 'q').toLocaleString('id-ID'), `${rows.length.toLocaleString('id-ID')} baris item`)
      + kartu('Struk berisi custom', new Set(rows.map(r => r.no)).size.toLocaleString('id-ID'), `${bln.length} bulan data`)
      + kartu('Rata-rata / bulan', rp(bln.length ? sum(rows, 'r') / bln.length : 0))
    // Per bulan
    document.getElementById('cuBulan').innerHTML = '<tr><th>Bulan</th><th class="r">Qty</th><th class="r">Omzet custom</th><th class="r">Struk</th><th class="r">% omzet toko</th></tr>' + bln.map(b => {
      const a = rows.filter(r => r.b === b), om = sum(a, 'r'), tot = DATA[b].semua?.omzet || 0
      return `<tr><td>${labelBln(b)}</td><td class="r">${sum(a, 'q').toLocaleString('id-ID')}</td><td class="r">${rp(om)}</td><td class="r">${new Set(a.map(r => r.no)).size}</td><td class="r">${tot ? (100 * om / tot).toFixed(1) + '%' : '—'}</td></tr>` }).join('')
    // Per kelompok
    const G = {}; rows.forEach(r => { const g = G[r.g] = G[r.g] || { q: 0, r: 0, n: 0 }; g.q += r.q; g.r += r.r; g.n++ })
    document.getElementById('cuKel').innerHTML = '<tr><th>Kelompok</th><th class="r">Qty</th><th class="r">Omzet</th><th class="r">Baris</th></tr>' + Object.entries(G).sort((a, b) => b[1].r - a[1].r)
      .map(([k, g]) => `<tr><td><a href="#" onclick="cuPilihKelompok(${esc(JSON.stringify(k))});return false" style="color:var(--g);font-weight:600">${esc(k)}</a></td><td class="r">${g.q.toLocaleString('id-ID')}</td><td class="r">${rp(g.r)}</td><td class="r">${g.n}</td></tr>`).join('')
    // Per jenis
    const J = {}; rows.forEach(r => { const j = J[r.j] = J[r.j] || { q: 0, r: 0, n: 0, g: r.g, bln: {}, contoh: r.i }; j.q += r.q; j.r += r.r; j.n++; j.bln[r.b] = (j.bln[r.b] || 0) + r.q })
    const urut = document.getElementById('cuUrut').value
    const jl = Object.entries(J).sort((a, b) => urut === 'qty' ? b[1].q - a[1].q : urut === 'n' ? b[1].n - a[1].n : b[1].r - a[1].r)
    window._cuJenis = jl
    document.getElementById('cuJenisN').textContent = `${jl.length.toLocaleString('id-ID')} jenis`
    document.getElementById('cuJenis').innerHTML = '<tr><th>#</th><th>Jenis</th><th>Kelompok</th><th class="r">Qty</th><th class="r">Omzet</th><th class="r">Baris</th><th>Bulan ramai</th></tr>' + jl.slice(0, 200).map(([k, j], i) => {
      const ramai = Object.entries(j.bln).sort((a, b) => b[1] - a[1]).slice(0, 2).map(([b]) => labelBln(b)).join(', ')
      return `<tr><td>${i + 1}</td><td style="white-space:normal;min-width:220px"><b>${esc(k)}</b>${k !== j.contoh.toUpperCase() ? `<div style="font-size:.7rem;color:#64748b">cth: ${esc(j.contoh)}</div>` : ''}</td><td>${esc(j.g)}</td><td class="r">${j.q.toLocaleString('id-ID')}</td><td class="r">${rp(j.r)}</td><td class="r">${j.n}</td><td>${esc(ramai)}</td></tr>` }).join('')
      + (jl.length > 200 ? `<tr><td colspan="7" style="color:#64748b">… ${jl.length - 200} jenis lain — persempit dengan cari / kelompok, atau unduh Excel</td></tr>` : '')
    renderTahun()
    // Detail
    const det = rows.slice().sort((a, b) => String(b.t).localeCompare(String(a.t)))
    document.getElementById('cuDetailN').textContent = `${det.length.toLocaleString('id-ID')} baris`
    document.getElementById('cuDetail').innerHTML = '<tr><th>Tanggal</th><th>Item (diketik kasir)</th><th>Jenis</th><th class="r">Qty</th><th class="r">Omzet</th><th>Kasir</th></tr>' + det.slice(0, 300).map(r =>
      `<tr><td>${esc(r.t)}</td><td style="white-space:normal;min-width:220px">${esc(r.i)}</td><td style="white-space:normal">${esc(r.j)}</td><td class="r">${(r.q || 0).toLocaleString('id-ID')}</td><td class="r">${rp(r.r)}</td><td>${esc(r.k)}</td></tr>`).join('')
      + (det.length > 300 ? `<tr><td colspan="6" style="color:#64748b">Menampilkan 300 dari ${det.length.toLocaleString('id-ID')} baris — unduh Excel untuk semuanya</td></tr>` : '')
  }
  // Dari Top Produk: buka tab ini dengan periode yang sama
  window.cuBukaDariTop = async () => {
    try { const idx = (typeof FILTERED_IDX !== 'undefined' && FILTERED_IDX.length) ? FILTERED_IDX : BULAN.map((_, i) => i); DARI_TOP = idx.map(i => keIso(BULAN[i])).filter(Boolean) } catch (e) { DARI_TOP = null }
    showTab('custom', document.getElementById('nav-custom'))
    await muat(); pilihanPeriode(); const sel = document.getElementById('cuPeriode'); if (DARI_TOP && [...sel.options].some(o => o.value === '__top')) sel.value = '__top'
    document.getElementById('cuKelompok').value = ''; document.getElementById('cuCari').value = ''; render()
    const ada = DARI_TOP ? DARI_TOP.filter(b => DATA[b]).length : 0
    const st = document.getElementById('cuStatus'); if (DARI_TOP && ada < DARI_TOP.length) st.textContent = `ℹ️ Rincian custom tersedia untuk ${ada} dari ${DARI_TOP.length} bulan periode ini — upload laporan Item Penjualan bulan lainnya.`
  }
  // ── Toko vs Shopee / marketplace ──
  function renderKanal() {
    if (!DATA) return
    const sel = document.getElementById('knPeriode'), v = sel.value, semuaBln = Object.keys(DATA).sort(), adaKanal = semuaBln.filter(b => DATA[b].kanal)
    const thn = [...new Set(adaKanal.map(b => b.slice(0, 4)))]
    sel.innerHTML = '<option value="">Semua bulan</option>' + thn.map(t => `<option value="${t}">Tahun ${t}</option>`).join('') + adaKanal.slice().reverse().map(b => `<option value="${b}">${labelBln(b)}</option>`).join('')
    if ([...sel.options].some(o => o.value === v)) sel.value = v
    const lama = semuaBln.filter(b => !DATA[b].kanal)
    document.getElementById('knInfo').innerHTML = !adaKanal.length ? 'Belum ada data. Upload laporan Olsera <b>Item Penjualan berdasarkan Tanggal</b> di tab 🧵 Barang Custom (sekali upload untuk dua tab ini).'
      : `Sumber dari nomor pesanan Olsera (SP… = Shopee, INV… = Tokopedia). Nilai = harga barang di Olsera (sebelum potongan marketplace).${lama.length ? ` <b style="color:#b45309">Bulan ${lama.map(labelBln).join(', ')} di-upload sebelum fitur ini — upload ulang supaya ikut terpisah.</b>` : ''}`
    document.getElementById('knIsi').style.display = adaKanal.length ? '' : 'none'; if (!adaKanal.length) return
    const bln = adaKanal.filter(b => !sel.value || b.startsWith(sel.value))
    const KAN = [...new Set(bln.flatMap(b => Object.keys(DATA[b].kanal)))].sort((a, b) => a === 'Toko' ? -1 : b === 'Toko' ? 1 : a.localeCompare(b))
    const tot = {}; KAN.forEach(k => tot[k] = { omzet: 0, qty: 0, struk: 0 }); bln.forEach(b => KAN.forEach(k => { const x = DATA[b].kanal[k]; if (x) { tot[k].omzet += x.omzet; tot[k].qty += x.qty; tot[k].struk += x.struk } }))
    const semua = KAN.reduce((t, k) => t + tot[k].omzet, 0), ikon = k => k === 'Toko' ? '🏪' : k === 'Shopee' ? '🛍️' : '🛒'
    document.getElementById('knKartu').innerHTML = KAN.map(k => `<div class="kpi-card"><div class="kpi-label">${ikon(k)} ${esc(k)}</div><div class="kpi-value" style="font-size:1.1rem">${rp(tot[k].omzet)}</div><div class="kpi-sub">${semua ? (100 * tot[k].omzet / semua).toFixed(1) + '% · ' : ''}${tot[k].struk.toLocaleString('id-ID')} pesanan · ${tot[k].qty.toLocaleString('id-ID')} pcs</div></div>`).join('')
    document.getElementById('knBulan').innerHTML = `<tr><th>Bulan</th>${KAN.map(k => `<th class="r">${esc(k)}</th>`).join('')}<th class="r">% marketplace</th></tr>` + bln.map(b => { const kn = DATA[b].kanal, t = KAN.reduce((s2, k) => s2 + (kn[k]?.omzet || 0), 0), mp = t - (kn.Toko?.omzet || 0)
      return `<tr><td>${labelBln(b)}</td>${KAN.map(k => `<td class="r">${kn[k] ? rp(kn[k].omzet) + `<div style="font-size:.66rem;color:#64748b">${kn[k].struk} pesanan</div>` : '—'}</td>`).join('')}<td class="r">${t ? (100 * mp / t).toFixed(1) + '%' : '—'}</td></tr>` }).join('')
    const mk = KAN.filter(k => k !== 'Toko'), pilih = document.getElementById('knKanal'), pv = pilih.value
    pilih.innerHTML = mk.map(k => `<option>${esc(k)}</option>`).join(''); if (mk.includes(pv)) pilih.value = pv
    const kk = pilih.value, M = {}; bln.forEach(b => (DATA[b].topKanal?.[kk] || []).forEach(e => { const x = M[e.i] = M[e.i] || { i: e.i, q: 0, r: 0 }; x.q += e.q; x.r += e.r }))
    const urut = document.getElementById('knUrut').value, top = Object.values(M).sort((a, c) => urut === 'qty' ? c.q - a.q : c.r - a.r)
    document.getElementById('knTop').innerHTML = '<tr><th>#</th><th>Barang</th><th class="r">Qty</th><th class="r">Nilai</th></tr>' + (top.length ? top.slice(0, 100).map((x, i) => `<tr><td>${i + 1}</td><td style="white-space:normal;min-width:200px">${esc(x.i)}</td><td class="r">${x.q.toLocaleString('id-ID')}</td><td class="r">${rp(x.r)}</td></tr>`).join('') : '<tr><td colspan="4" style="color:#64748b">Tidak ada penjualan marketplace di periode ini</td></tr>')
  }
  window.knRender = () => renderKanal()
  // ── Custom yang distok (topi/dasi sekolah, dll.): rata-rata 3 bulan per jenis yang cocok kata kunci → usulan siapkan ──
  const cocokStok = (j, i) => { const t = (j + ' ' + i).toUpperCase(); return KATA_STOK.find(k => t.includes(k)) }
  function hitungCustomStok(n = 3) {
    const bln = Object.keys(DATA || {}).sort().slice(-n), J = {}
    bln.forEach(b => (DATA[b].rows || []).forEach(r => { const k = cocokStok(r.j, r.i); if (!k) return; const x = J[r.j] = J[r.j] || { j: r.j, kata: k, q: 0, contoh: r.i }; x.q += r.q || 0 }))
    return { bln, list: Object.values(J).map(x => ({ ...x, avg: bln.length ? +(x.q / bln.length).toFixed(1) : 0 })).sort((a, b) => b.avg - a.avg) }
  }
  window._cuHitungStok = hitungCustomStok
  window._cuMuat = paksa => muat(paksa)
  function renderStokCustom() {
    const box = document.getElementById('cuStok'); if (!box) return
    const owner = !!window._isOwnerAn
    document.getElementById('cuKata').value = KATA_STOK.join(', '); document.getElementById('cuKata').disabled = !owner
    document.getElementById('cuKataBtn').style.display = owner ? '' : 'none'
    if (!KATA_STOK.length) { box.innerHTML = '<p style="font-size:.78rem;color:#64748b">Belum ada kata kunci. Owner isi di atas, mis. <b>TOPI, DASI</b>.</p>'; return }
    const { bln, list } = hitungCustomStok(3), bulan = +document.getElementById('cuStokBln').value || 1
    box.innerHTML = list.length ? `<p style="font-size:.74rem;color:#475569;margin-bottom:6px">Rata-rata ${bln.map(labelBln).join(', ')}. Usulan = laku/bln × ${bulan} bulan.</p><div class="cu-scroll"><table><tr><th>Jenis custom</th><th>Kata</th><th class="r">Laku/bln</th><th class="r">Usulan siapkan</th></tr>${list.slice(0, 200).map(x => `<tr><td style="white-space:normal;min-width:200px"><b style="font-weight:600">${esc(x.j)}</b>${x.j !== x.contoh.toUpperCase() ? `<div style="font-size:.66rem;color:#64748b">cth: ${esc(x.contoh)}</div>` : ''}</td><td>${esc(x.kata)}</td><td class="r">${x.avg}</td><td class="r" style="font-weight:700">${Math.ceil(x.avg * bulan)}</td></tr>`).join('')}</table></div>`
      : '<p style="font-size:.78rem;color:#64748b">Tidak ada penjualan custom yang cocok dengan kata kunci di 3 bulan terakhir.</p>'
  }
  window.cuRenderStok = () => renderStokCustom()
  window.cuSimpanKata = async () => {
    const kata = document.getElementById('cuKata').value.split(',').map(t => t.trim().toUpperCase()).filter(Boolean)
    try { await _app.firestore().collection('settings').doc('custom_stok').set({ kata, updated_at: firebase.firestore.FieldValue.serverTimestamp(), updated_by: _currentUserName || '' }); KATA_STOK = kata; renderStokCustom(); alert('✅ Kata kunci custom yang distok disimpan') }
    catch (e) { alert('Gagal: ' + e.message) }
  }
  // ── Perbandingan per tahun (abaikan pilihan periode; ikut cari & kelompok) ──
  function renderTahun() {
    const box = document.getElementById('cuThn'); if (!box || !DATA) return
    const q = document.getElementById('cuCari').value.trim().toUpperCase(), kel = document.getElementById('cuKelompok').value
    const ukur = document.getElementById('cuThnUkur').value, ytd = document.getElementById('cuThnYtd').checked, urut = document.getElementById('cuThnUrut').value
    const semuaBln = Object.keys(DATA).sort(), tahun = [...new Set(semuaBln.map(b => b.slice(0, 4)))]
    if (tahun.length < 2) { box.innerHTML = '<p style="font-size:.78rem;color:#64748b">Perlu data minimal 2 tahun untuk dibandingkan.</p>'; return }
    const thnAkhir = tahun[tahun.length - 1], blnAkhir = semuaBln.filter(b => b.startsWith(thnAkhir)).map(b => +b.slice(5))
    const batas = ytd ? Math.max(...blnAkhir) : 12
    const nilai = { omzet: r => r.r || 0, qty: r => r.q || 0, struk: () => 0 }[ukur]
    // agregasi: [tahun][bulan], [tahun][kelompok], [tahun][jenis]; struk dihitung unik
    const M = {}, G = {}, J = {}, SM = {}, SG = {}, SJ = {}
    semuaBln.forEach(b => { const t = b.slice(0, 4), m = +b.slice(5); if (m > batas) return
      ;(DATA[b].rows || []).forEach(r => { const g = kelompokDari(r.i, r.j); if ((q && !(r.i + ' ' + r.j).toUpperCase().includes(q)) || (kel && g !== kel)) return
        const v = nilai(r); const add = (O, a, k) => { (O[a] = O[a] || {})[k] = (O[a][k] || 0) + v }; add(M, t, m); add(G, t, g); add(J, t, r.j)
        if (ukur === 'struk') { const s2 = (O, a, k) => { ((O[a] = O[a] || {})[k] = O[a][k] || new Set()).add(r.no) }; s2(SM, t, m); s2(SG, t, g); s2(SJ, t, r.j) } }) })
    const ambil = (O, S, t, k) => ukur === 'struk' ? (S[t]?.[k]?.size || 0) : (O[t]?.[k] || 0)
    const fmt = v => ukur === 'omzet' ? rpJtC(v) : Math.round(v).toLocaleString('id-ID')
    const t1 = tahun[tahun.length - 2], t2 = thnAkhir
    const delta = (a, b) => a ? `<span style="font-weight:700;color:${b >= a ? '#166534' : '#b91c1c'}">${b >= a ? '▲' : '▼'} ${Math.abs(Math.round((b - a) / a * 100))}%</span>` : (b ? '<span style="color:#166534;font-weight:700">baru</span>' : '—')
    const kolTh = tahun.map(t => `<th class="r">${t}${t === thnAkhir && ytd ? '' : ''}</th>`).join('')
    const total = t => { let s3 = 0; for (let m = 1; m <= batas; m++) s3 += ambil(M, SM, t, m); return s3 }
    const NB = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
    let html = `<p style="font-size:.74rem;color:#475569;margin-bottom:6px">${ytd ? `Dibandingkan sampai bulan yang sama: <b>Jan–${NB[batas]}</b> tiap tahun.` : 'Semua bulan tiap tahun.'} Perubahan = ${t2} dibanding ${t1}.</p>
      <div class="cu-grid"><div><div class="cu-scroll"><table><tr><th>Bulan</th>${kolTh}<th class="r">${t2} vs ${t1}</th></tr>`
    for (let m = 1; m <= batas; m++) html += `<tr><td>${NB[m]}</td>${tahun.map(t => `<td class="r">${fmt(ambil(M, SM, t, m))}</td>`).join('')}<td class="r">${delta(ambil(M, SM, t1, m), ambil(M, SM, t2, m))}</td></tr>`
    html += `<tr style="font-weight:700;background:#f8fafc"><td>Total</td>${tahun.map(t => `<td class="r">${fmt(total(t))}</td>`).join('')}<td class="r">${delta(total(t1), total(t2))}</td></tr></table></div></div>`
    const kels = [...new Set(tahun.flatMap(t => Object.keys(G[t] || {})))].sort((a, b) => ambil(G, SG, t2, b) - ambil(G, SG, t2, a))
    html += `<div><div class="cu-scroll"><table><tr><th>Kelompok</th>${kolTh}<th class="r">${t2} vs ${t1}</th></tr>${kels.map(g => `<tr><td>${esc(g)}</td>${tahun.map(t => `<td class="r">${fmt(ambil(G, SG, t, g))}</td>`).join('')}<td class="r">${delta(ambil(G, SG, t1, g), ambil(G, SG, t2, g))}</td></tr>`).join('')}</table></div></div></div>`
    const js = [...new Set([...Object.keys(J[t1] || {}), ...Object.keys(J[t2] || {})])].map(j => ({ j, a: ambil(J, SJ, t1, j), b: ambil(J, SJ, t2, j) })).map(x => ({ ...x, d: x.b - x.a }))
    const list = js.sort((x, y) => urut === 'naik' ? y.d - x.d : urut === 'turun' ? x.d - y.d : y.b - x.b).slice(0, 50)
    html += `<div style="margin-top:12px;font-size:.78rem;font-weight:700">Jenis custom — ${urut === 'naik' ? 'paling naik' : urut === 'turun' ? 'paling turun' : 'terbesar ' + t2} (${t1} vs ${t2})</div>
      <div class="cu-scroll"><table><tr><th>#</th><th>Jenis</th><th class="r">${t1}</th><th class="r">${t2}</th><th class="r">Selisih</th><th class="r">Perubahan</th></tr>${list.map((x, i) => `<tr><td>${i + 1}</td><td style="white-space:normal;min-width:200px">${esc(x.j)}</td><td class="r">${fmt(x.a)}</td><td class="r">${fmt(x.b)}</td><td class="r" style="color:${x.d >= 0 ? '#166534' : '#b91c1c'}">${x.d >= 0 ? '+' : '−'}${fmt(Math.abs(x.d))}</td><td class="r">${delta(x.a, x.b)}</td></tr>`).join('')}</table></div>`
    box.innerHTML = html
  }
  const rpJtC = n => Math.abs(n) >= 1e6 ? 'Rp ' + (n / 1e6).toLocaleString('id-ID', { maximumFractionDigits: 1 }) + ' jt' : rp(n)
  window.cuRenderTahun = () => renderTahun()
  window.cuPilihKelompok = k => { document.getElementById('cuKelompok').value = k; render() }
  window.cuRender = () => render()
  window.cuExcel = () => {
    const { rows } = terpilih(), wb = XLSX.utils.book_new()
    const tambah = (data, n, lebar) => { const ws = XLSX.utils.json_to_sheet(data); ws['!cols'] = lebar.map(w => ({ wch: w })); XLSX.utils.book_append_sheet(wb, ws, n) }
    const J = {}; rows.forEach(r => { const j = J[r.j] = J[r.j] || { Jenis: r.j, Kelompok: r.g, Qty: 0, Omzet: 0, Baris: 0, Contoh: r.i }; j.Qty += r.q; j.Omzet += r.r; j.Baris++ })
    const Bl = {}; rows.forEach(r => { const b = Bl[r.b] = Bl[r.b] || { Bulan: labelBln(r.b), Qty: 0, Omzet: 0, Struk: new Set() }; b.Qty += r.q; b.Omzet += r.r; b.Struk.add(r.no) })
    tambah(Object.keys(Bl).sort().map(b => ({ ...Bl[b], Struk: Bl[b].Struk.size, '% omzet toko': DATA[b]?.semua?.omzet ? +(100 * Bl[b].Omzet / DATA[b].semua.omzet).toFixed(1) : '' })), 'Per Bulan', [12, 10, 14, 8, 12])
    tambah(Object.values(J).sort((a, b) => b.Omzet - a.Omzet), 'Per Jenis', [45, 22, 10, 14, 8, 55])
    tambah(rows.map(r => ({ Tanggal: r.t, 'No. Struk': r.no, Kasir: r.k, Pelanggan: r.p || '', 'Item (diketik kasir)': r.i, Jenis: r.j, Kelompok: r.g, Qty: r.q, Omzet: r.r })), 'Semua Item', [11, 20, 10, 20, 55, 30, 18, 7, 12])
    XLSX.writeFile(wb, 'Barang_Custom_Harmoni.xlsx')
  }
  // Folder yang di-compress (ZIP) → file .xlsx/.pdf di dalamnya
  async function bukaZip(list, st) {
    if (!list.some(f => /\.zip$/i.test(f.name))) return list
    if (typeof JSZip === 'undefined') { st.textContent = '⏳ Memuat pembuka ZIP...'
      await new Promise((ok, gagal) => { const sc = document.createElement('script'); sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'; sc.onload = ok; sc.onerror = () => gagal(new Error('Gagal memuat pembuka ZIP — cek internet')); document.head.appendChild(sc) }) }
    const hasil = list.filter(f => !/\.zip$/i.test(f.name))
    for (const z of list.filter(f => /\.zip$/i.test(f.name))) {
      st.textContent = '⏳ Membuka ' + z.name + '...'
      const arsip = await JSZip.loadAsync(await z.arrayBuffer())
      for (const e of Object.values(arsip.files)) { const nama = e.name.split('/').pop()
        if (e.dir || e.name.includes('__MACOSX') || nama.startsWith('.') || !/\.(xlsx|xls|pdf)$/i.test(nama)) continue
        hasil.push(new File([await e.async('blob')], nama)) }
    }
    if (!hasil.length) throw new Error('Tidak ada file Excel/PDF di dalam ZIP')
    return hasil
  }
  window.cuUpload = async files => {
    const st = document.getElementById('cuStatus'); if (!files || !files.length) return
    try {
      if (_peranAnalitik !== 'admin') throw new Error('Upload hanya untuk Owner/Admin')
      if (typeof XLSX === 'undefined') throw new Error('Pembaca Excel belum termuat — muat ulang halaman')
      const semua = []
      files = await bukaZip([...files], st)
      for (const f of files) {
        st.textContent = `⏳ Membaca ${f.name}...`
        const buf = await f.arrayBuffer()
        semua.push(...(/\.pdf$/i.test(f.name) ? await bacaPdf(buf, st) : bacaExcel(buf)))
      }
      const B = perBulan(semua), bl = Object.keys(B).sort()
      if (!bl.length) throw new Error('Tanggal tidak terbaca di file ini')
      const nCu = bl.reduce((t, b) => t + B[b].rows.length, 0)
      if (!confirm(`Terbaca ${semua.length.toLocaleString('id-ID')} baris item, ${nCu.toLocaleString('id-ID')} item custom, bulan: ${bl.map(labelBln).join(', ')}.\n\nData custom bulan tersebut akan diganti dengan file ini. Lanjut simpan?`)) { st.textContent = 'Dibatalkan'; return }
      st.textContent = '☁️ Menyimpan...'
      await simpan(B, [...files].map(f => f.name).join(', '))
      await muat(true); render()
      st.textContent = `✅ Tersimpan: ${bl.map(labelBln).join(', ')} · ${nCu.toLocaleString('id-ID')} item custom${_OUTLET ? ' (outlet ini)' : ' (Stok Manager ikut terbarui)'}`
    } catch (e) { console.error(e); st.textContent = '⚠️ ' + e.message }
    finally { const inp = document.getElementById('cuFile'); if (inp) inp.value = '' }
  }

  // ── Pasang tab ────────────────────────────────────────────────────
  function pasang() {
    const nav = document.querySelector('.nav'), main = document.querySelector('.main'); if (!nav || !main || document.getElementById('nav-custom')) return
    const btn = document.createElement('button'); btn.id = 'nav-custom'; btn.className = 'nav-btn'; btn.style.display = 'none'; btn.textContent = '🧵 Barang Custom'
    btn.onclick = function () { showTab('custom', this) }
    nav.appendChild(btn)
    const tab = document.createElement('div'); tab.id = 'tab-custom'; tab.className = 'tab-content'
    tab.innerHTML = `<style>
      #tab-custom table{width:100%;border-collapse:collapse;font-size:.78rem} #tab-custom th{text-align:left;font-size:.68rem;color:#64748b;padding:6px 8px;border-bottom:2px solid rgba(3,69,67,.1);white-space:nowrap;position:sticky;top:0;background:#fff}
      #tab-custom td{padding:6px 8px;border-bottom:1px solid rgba(3,69,67,.06);white-space:nowrap} #tab-custom .r{text-align:right} #tab-custom .cu-scroll{max-height:60vh;overflow:auto}
      #tab-custom .cu-filter{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:12px} #tab-custom .cu-filter select,#tab-custom .cu-filter input{padding:6px 10px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2);font-size:.8rem;min-width:0}
      #tab-custom .cu-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px} @media(max-width:640px){#tab-custom .cu-grid{grid-template-columns:1fr}}
    </style>
    <div class="card">
      <div class="card-title">🧵 Barang Custom (nama diketik kasir di Olsera)</div>
      <p style="font-size:.78rem;color:#475569;margin-bottom:10px">Upload laporan Olsera <b>Laporan → Item Penjualan berdasarkan Tanggal</b> (Excel atau PDF), boleh beberapa bulan sekaligus atau satu file ZIP berisi banyak file. Item tanpa satuan & tanpa modal dianggap custom. Upload ulang bulan yang sama akan mengganti datanya. Ringkasan qty per jenis ikut terkirim ke <b>Stok Manager</b> untuk menghitung kebutuhan bahan di Daftar Belanja.</p>
      <label class="pbtn active" style="cursor:pointer;display:inline-block">📤 Upload Item Penjualan<input type="file" id="cuFile" accept=".xlsx,.xls,.pdf,.zip" multiple style="display:none" onchange="cuUpload(this.files)"></label>
      <span id="cuStatus" style="font-size:.75rem;color:#64748b;margin-left:8px"></span>
    </div>
    <div class="card"><div class="card-title">🧢 Custom yang distok / disiapkan <select id="cuStokBln" onchange="cuRenderStok()" style="font-size:.75rem;padding:4px 8px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2)"><option value="1">untuk 1 bulan</option><option value="2">2 bulan</option><option value="3">3 bulan</option></select></div>
      <p style="font-size:.76rem;color:#475569;margin-bottom:6px">Barang custom yang laku rutin (mis. topi & dasi sekolah) bisa disiapkan lebih dulu. Kata kunci berlaku untuk HQ & semua outlet.</p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px"><input id="cuKata" placeholder="Kata kunci, pisah koma — mis. TOPI, DASI" style="flex:1;min-width:200px;padding:6px 10px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2)"><button id="cuKataBtn" class="pbtn active" onclick="cuSimpanKata()">💾 Simpan</button></div>
      <div id="cuStok"></div></div>
    <div id="cuKosong" class="card" style="display:none;color:#64748b;font-size:.85rem">Belum ada data barang custom. Upload laporan Item Penjualan di atas.</div>
    <div id="cuIsi" style="display:none">
      <div class="cu-filter">
        <select id="cuPeriode" onchange="cuRender()"></select>
        <select id="cuKelompok" onchange="cuRender()"><option value="">Semua kelompok</option>${KELOMPOK.map(([k]) => `<option>${k}</option>`).join('')}<option>Lainnya</option></select>
        <input id="cuCari" type="search" placeholder="Cari nama item / jenis…" oninput="clearTimeout(window._cuT);window._cuT=setTimeout(cuRender,250)" style="flex:1;min-width:160px">
        <button class="pbtn" onclick="cuExcel()">⬇️ Excel</button>
      </div>
      <div class="kpi-grid" id="cuKartu"></div>
      <div class="cu-grid">
        <div class="card"><div class="card-title">📅 Per bulan</div><div class="cu-scroll"><table id="cuBulan"></table></div></div>
        <div class="card"><div class="card-title">🗂️ Per kelompok</div><div class="cu-scroll"><table id="cuKel"></table></div></div>
      </div>
      <div class="card"><div class="card-title">📅 Perbandingan per tahun
        <span style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;font-weight:500">
          <select id="cuThnUkur" onchange="cuRenderTahun()" style="font-size:.75rem;padding:4px 8px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2)"><option value="omzet">Omzet</option><option value="qty">Qty</option><option value="struk">Jumlah struk</option></select>
          <select id="cuThnUrut" onchange="cuRenderTahun()" style="font-size:.75rem;padding:4px 8px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2)"><option value="naik">Jenis paling naik</option><option value="turun">Jenis paling turun</option><option value="besar">Jenis terbesar</option></select>
          <label style="font-size:.72rem;display:flex;gap:4px;align-items:center;color:#475569"><input type="checkbox" id="cuThnYtd" checked onchange="cuRenderTahun()" style="width:auto"> sampai bulan yang sama</label></span></div>
        <p style="font-size:.72rem;color:#64748b;margin-bottom:6px">Ikut filter kelompok & pencarian di atas (periode diabaikan).</p>
        <div id="cuThn"></div></div>
      <div class="card"><div class="card-title">🏆 Jenis custom <span id="cuJenisN" style="font-size:.72rem;color:#64748b;font-weight:500"></span>
        <select id="cuUrut" onchange="cuRender()" style="font-size:.75rem;padding:4px 8px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2)"><option value="omzet">Omzet terbesar</option><option value="qty">Qty terbanyak</option><option value="n">Paling sering</option></select></div>
        <div class="cu-scroll"><table id="cuJenis"></table></div></div>
      <div class="card"><div class="card-title">📋 Rincian item <span id="cuDetailN" style="font-size:.72rem;color:#64748b;font-weight:500"></span></div><div class="cu-scroll"><table id="cuDetail"></table></div></div>
    </div>`
    main.appendChild(tab)
    const btn2 = document.createElement('button'); btn2.id = 'nav-kanal'; btn2.className = 'nav-btn'; btn2.style.display = 'none'; btn2.textContent = '🛍️ Toko vs Shopee'
    btn2.onclick = function () { showTab('kanal', this) }
    nav.appendChild(btn2)
    const tab2 = document.createElement('div'); tab2.id = 'tab-kanal'; tab2.className = 'tab-content'
    tab2.innerHTML = `<style>#tab-kanal table{width:100%;border-collapse:collapse;font-size:.78rem} #tab-kanal th{text-align:left;font-size:.68rem;color:#64748b;padding:6px 8px;border-bottom:2px solid rgba(3,69,67,.1);white-space:nowrap;position:sticky;top:0;background:#fff}
      #tab-kanal td{padding:6px 8px;border-bottom:1px solid rgba(3,69,67,.06);white-space:nowrap;vertical-align:top} #tab-kanal .r{text-align:right} #tab-kanal select{padding:6px 10px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2);font-size:.8rem}</style>
    <div class="card"><div class="card-title">🛍️ Penjualan Toko vs Shopee / marketplace</div><p id="knInfo" style="font-size:.78rem;color:#475569"></p>
      <div style="margin-top:8px"><select id="knPeriode" onchange="knRender()"></select></div></div>
    <div id="knIsi" style="display:none">
      <div class="kpi-grid" id="knKartu"></div>
      <div class="card"><div class="card-title">📅 Per bulan</div><div style="max-height:60vh;overflow:auto"><table id="knBulan"></table></div></div>
      <div class="card"><div class="card-title">🏆 Barang terlaris di <select id="knKanal" onchange="knRender()"></select> <select id="knUrut" onchange="knRender()"><option value="omzet">Nilai terbesar</option><option value="qty">Qty terbanyak</option></select></div>
        <div style="max-height:60vh;overflow:auto"><table id="knTop"></table></div></div>
    </div>`
    main.appendChild(tab2)
    const asli = window.showTab
    window.showTab = function (id, b) { asli(id, b); if (id === 'custom') muat().then(render); if (id === 'kanal') muat().then(renderKanal) }
    _fbReady.then(() => { if (_peranAnalitik === 'admin' && (_OUTLET || !_outletSaya)) { btn.style.display = ''; btn2.style.display = '' } }).catch(() => {})
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', pasang); else pasang()
  window._cuTest = { bacaExcel, perBulan, jenisDari, kelompokDari, tglIso }
})()
