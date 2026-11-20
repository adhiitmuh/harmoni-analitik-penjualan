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
  const jenisDari = nama => { const n = String(nama || '').toUpperCase().replace(/\s+/g, ' ').trim(); const m = n.match(/\(([^()]+)\)\s*$/); return (m ? m[1] : n).trim() }
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
      const kol = (...re) => { for (const r of re) { const i = H.findIndex(h => r.test(h)); if (i >= 0) return i } return -1 }
      const K = { no: kol(/no\.? ?pesanan|order no|no\.? ?order|no\.? ?struk/), tgl: kol(/tanggal/, /date/), plg: kol(/pelanggan/, /customer/), oleh: kol(/penjualan oleh|oleh|sales|served/),
        item: kol(/^item$/, /nama item|item name/), qty: kol(/^qty$/, /qty/, /jumlah/), unit: kol(/unit|satuan/), total: kol(/total pesanan/, /^total/, /subtotal/), modal: kol(/modal/, /cost/) }
      if (K.item < 0 || K.qty < 0) continue
      const rows = []
      for (const r of a.slice(hi + 1)) {
        const item = String(r[K.item] || '').replace(/\s+/g, ' ').trim(); if (!item) continue
        rows.push({ no: String(r[K.no] ?? '').trim(), tgl: tglIso(r[K.tgl]), plg: String(r[K.plg] ?? '').trim(), oleh: String(r[K.oleh] ?? '').trim(), item,
          qty: angka(r[K.qty]), unit: K.unit >= 0 ? String(r[K.unit] ?? '').trim() : '', total: angka(r[K.total]), modal: K.modal >= 0 ? angka(r[K.modal]) : 0 })
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
  function perBulan(rows) {
    const B = {}
    rows.forEach(r => { const b = (r.tgl || '').slice(0, 7); if (!/^\d{4}-\d{2}$/.test(b)) return
      const x = B[b] = B[b] || { rows: [], omzet: 0, baris: 0 }
      x.omzet += r.total || 0; x.baris++
      if (!r.unit && !(r.modal > 0) && r.item) x.rows.push({ t: r.tgl, no: r.no, k: r.oleh, p: r.plg, i: r.item, j: jenisDari(r.item), q: r.qty || 0, r: r.total || 0 }) })
    return B
  }

  async function simpan(B, nama) {
    const db = _app.firestore(), now = firebase.firestore.FieldValue.serverTimestamp()
    for (const [b, x] of Object.entries(B)) {
      let rows = x.rows
      if (JSON.stringify(rows).length > 900000) rows = rows.map(({ p, ...s }) => s)   // batas 1 MB dokumen: buang nama pelanggan
      const ring = {}; rows.forEach(r => { ring[r.j] = (ring[r.j] || 0) + (r.q || 0) })
      await db.collection('penjualan_custom').doc(b).set({ rows, semua: { omzet: x.omzet, baris: x.baris }, sumber: nama, updated_at: now, updated_by: _currentUserName || '' })
      await db.collection('penjualan_custom_ringkas').doc(b).set({ j: Object.entries(ring).map(([k, q]) => ({ k, q })), updated_at: now })
    }
  }

  // ── Data & tampilan ───────────────────────────────────────────────
  let DATA = null   // { 'YYYY-MM': {rows, semua} }
  let DARI_TOP = null   // daftar 'YYYY-MM' periode Top Produk saat baris "Custom" diklik
  // Label bulan Analitik ('Jan-25', 'Agt-26', 'Mei-26') → 'YYYY-MM'
  const keIso = l => { const m = String(l).match(/^([A-Za-z]+)[-\s]?(\d{2,4})$/); if (!m) return ''; const b = BLN[m[1].slice(0, 3).toLowerCase()]; const y = +m[2] < 100 ? 2000 + +m[2] : +m[2]; return b ? `${y}-${pad(b)}` : '' }
  async function muat(paksa) {
    if (DATA && !paksa) return DATA
    DATA = {}
    try { await _fbReady; (await _app.firestore().collection('penjualan_custom').get()).forEach(d => { DATA[d.id] = d.data() }) }
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
  window.cuUpload = async files => {
    const st = document.getElementById('cuStatus'); if (!files || !files.length) return
    try {
      if (_peranAnalitik !== 'admin') throw new Error('Upload hanya untuk Owner/Admin')
      if (typeof XLSX === 'undefined') throw new Error('Pembaca Excel belum termuat — muat ulang halaman')
      const semua = []
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
      st.textContent = `✅ Tersimpan: ${bl.map(labelBln).join(', ')} · ${nCu.toLocaleString('id-ID')} item custom (Stok Manager ikut terbarui)`
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
      <p style="font-size:.78rem;color:#475569;margin-bottom:10px">Upload laporan Olsera <b>Laporan → Item Penjualan berdasarkan Tanggal</b> (Excel, atau PDF), boleh beberapa bulan sekaligus. Item tanpa satuan & tanpa modal dianggap custom. Upload ulang bulan yang sama akan mengganti datanya. Ringkasan qty per jenis ikut terkirim ke <b>Stok Manager</b> untuk menghitung kebutuhan bahan di Daftar Belanja.</p>
      <label class="pbtn active" style="cursor:pointer;display:inline-block">📤 Upload Item Penjualan<input type="file" id="cuFile" accept=".xlsx,.xls,.pdf" multiple style="display:none" onchange="cuUpload(this.files)"></label>
      <span id="cuStatus" style="font-size:.75rem;color:#64748b;margin-left:8px"></span>
    </div>
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
      <div class="card"><div class="card-title">🏆 Jenis custom <span id="cuJenisN" style="font-size:.72rem;color:#64748b;font-weight:500"></span>
        <select id="cuUrut" onchange="cuRender()" style="font-size:.75rem;padding:4px 8px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2)"><option value="omzet">Omzet terbesar</option><option value="qty">Qty terbanyak</option><option value="n">Paling sering</option></select></div>
        <div class="cu-scroll"><table id="cuJenis"></table></div></div>
      <div class="card"><div class="card-title">📋 Rincian item <span id="cuDetailN" style="font-size:.72rem;color:#64748b;font-weight:500"></span></div><div class="cu-scroll"><table id="cuDetail"></table></div></div>
    </div>`
    main.appendChild(tab)
    const asli = window.showTab
    window.showTab = function (id, b) { asli(id, b); if (id === 'custom') muat().then(render) }
    _fbReady.then(() => { if (_peranAnalitik === 'admin') btn.style.display = '' }).catch(() => {})
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', pasang); else pasang()
  window._cuTest = { bacaExcel, perBulan, jenisDari, kelompokDari, tglIso }
})()
