// Multi-outlet Tahap 2–3 (mode ?outlet=): stok outlet + pesanan barang ke HQ.
// • Stok: upload export produk Olsera akun outlet (kolom sku, stock_qty, name, variant_names, category) →
//   outlet_stok/{outlet} {stok:{SKU:qty}, info:{SKU:[nama,varian,kategori]}, sendiri:{SKU:true}}.
// • Laku/bln outlet = rata-rata 3 bulan terakhir dari outlet_penjualan (laporan yang di-upload outlet).
// • Sumber barang: outlet Mitra → semua dari HQ; outlet Keluarga boleh menandai barang "produksi/beli sendiri" (tidak dipesan ke HQ).
// • Usulan pesan = laku × lama stok − stok − yang sedang dipesan. Pesanan → outlet_pesanan (diajukan → dikirim HQ → diterima outlet; stok outlet bertambah).
(function () {
  if (!_OUTLET) return
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
  const BLN = { jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6, jul: 7, agu: 8, agt: 8, ags: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12 }
  const urutBln = l => { const m = String(l).match(/^([A-Za-z]+)[-\s]?(\d{2,4})$/); if (!m) return 0; const y = +m[2] < 100 ? 2000 + +m[2] : +m[2]; return y * 100 + (BLN[m[1].slice(0, 3).toLowerCase()] || 0) }
  const STATUS = { diajukan: ['⏳ Menunggu HQ', '#92400e', '#fef3c7'], dikirim: ['🚚 Dikirim HQ', '#1d4ed8', '#dbeafe'], diterima: ['✅ Diterima', '#166534', '#dcfce7'], ditolak: ['✕ Ditolak HQ', '#991b1b', '#fee2e2'], batal: ['Dibatalkan', '#475569', '#f1f5f9'] }
  let S = null   // { outlet, stok, info, sendiri, laku, nama, pesanan, bulan }
  const FV = () => firebase.firestore.FieldValue

  async function muat() {
    const db = _app.firestore()
    const [o, st, pj, ps] = await Promise.all([db.collection('outlet').doc(_OUTLET).get(), db.collection('outlet_stok').doc(_OUTLET).get(),
      db.collection('outlet_penjualan').where('outlet', '==', _OUTLET).get(), db.collection('outlet_pesanan').where('outlet', '==', _OUTLET).get()])
    const sd = st.exists ? st.data() : {}
    S = { outlet: o.exists ? o.data() : { nama: _OUTLET }, stok: sd.stok || {}, info: sd.info || {}, sendiri: sd.sendiri || {}, stokAt: sd.updated_at, laku: {}, nama: {}, pesanan: [], bulan: [] }
    // Laku per SKU: rata-rata 3 bulan terakhir
    const per = {}; pj.forEach(d => (d.data().rows || []).forEach(r => { const k = String(r.sku || '').trim(); if (!k) return; (per[k] = per[k] || {})[r.bulan] = (per[k][r.bulan] || 0) + (+r.qty || 0); if (!S.nama[k]) S.nama[k] = [r.produk, r.varian, r.kategori] }))
    S.bulan = [...new Set(Object.values(per).flatMap(x => Object.keys(x)))].sort((a, b) => urutBln(a) - urutBln(b)).slice(-3)
    Object.entries(per).forEach(([k, m]) => { S.laku[k] = S.bulan.length ? +(S.bulan.reduce((t, b) => t + (m[b] || 0), 0) / S.bulan.length).toFixed(1) : 0 })
    ps.forEach(d => S.pesanan.push({ id: d.id, ...d.data() }))
    S.pesanan.sort((a, b) => (b.dibuat_at?.seconds || 0) - (a.dibuat_at?.seconds || 0))
    return S
  }
  const mitra = () => S?.outlet?.jenis === 'mitra'
  // Barang "stok sendiri" ditetapkan Owner per outlet (kata kunci nama/kategori) — tidak dipesan ke HQ
  const kataSendiri = () => mitra() ? [] : (S?.outlet?.stokSendiri || [])
  const isSendiri = (nama, kat = '') => { const t = (nama + ' ' + kat).toUpperCase(); return kataSendiri().find(k => t.includes(k)) || '' }
  const namaSku = k => { const i = S.info[k] || S.nama[k]; return i ? [i[0], i[1]].filter(x => x && x !== '-').join(' · ') : k }
  const sedangDipesan = () => { const m = {}; S.pesanan.filter(p => p.status === 'diajukan' || p.status === 'dikirim').forEach(p => (p.items || []).forEach(it => { m[it.sku] = (m[it.sku] || 0) + (+(p.status === 'dikirim' ? it.kirim ?? it.qty : it.qty) || 0) })); return m }

  function render() {
    if (!S) return
    const bln = +document.getElementById('osBulan').value || 1, q = document.getElementById('osCari').value.trim().toUpperCase(), f = document.getElementById('osFilter').value
    const dp = sedangDipesan()
    // Barang custom (diketik kasir / kategori CUSTOM) tidak dipesan dari stok HQ
    const custom = k => { const i = S.info[k] || S.nama[k] || []; return /^custom$/i.test(String(k).trim()) || /^custom$/i.test(String(i[0] || '').trim()) || /^custom$/i.test(String(i[2] || '').trim()) }
    const sku = [...new Set([...Object.keys(S.stok), ...Object.keys(S.laku)])].filter(k => !custom(k))
    const rows = sku.map(k => { const stok = +S.stok[k] || 0, laku = S.laku[k] || 0, i = S.info[k] || S.nama[k] || [], sendiri = isSendiri(namaSku(k), i[2] || ''), dipesan = dp[k] || 0
      const usul = sendiri ? 0 : Math.max(0, Math.ceil(laku * bln) - stok - dipesan)
      return { k, nama: namaSku(k), stok, laku, sendiri, dipesan, usul } })
      .filter(r => (!q || (r.k + ' ' + r.nama).toUpperCase().includes(q)) && (f === 'semua' || (f === 'pesan' && r.usul > 0) || (f === 'kosong' && r.stok <= 0 && r.laku > 0) || (f === 'sendiri' && r.sendiri)))
      .sort((a, b) => b.usul - a.usul || b.laku - a.laku)
    window._osRows = rows
    const n = rows.filter(r => r.usul > 0).length
    document.getElementById('osInfo').innerHTML = `${S.stokAt?.toDate ? 'Stok di-upload ' + S.stokAt.toDate().toLocaleString('id-ID', { dateStyle: 'medium', timeStyle: 'short' }) : '<b style="color:#b45309">Stok outlet belum di-upload</b>'} · laku/bln = rata-rata ${S.bulan.length ? S.bulan.join(', ') : '(belum ada data penjualan outlet)'} · <b>${n}</b> barang perlu dipesan${mitra() ? ' · outlet Mitra: semua barang dari HQ' : ''}`
    document.getElementById('osTabel').innerHTML = `<tr><th>Barang</th><th class="r">Stok</th><th class="r">Laku/bln</th><th class="r">Dipesan</th>${mitra() ? '' : '<th>Sumber</th>'}<th class="r">Pesan ke HQ</th></tr>` + rows.slice(0, 400).map((r, i) => `<tr>
      <td style="white-space:normal;min-width:180px"><b style="font-weight:600">${esc(r.nama)}</b><div style="font-size:.66rem;color:#94a3b8">${esc(r.k)}</div></td>
      <td class="r" style="${r.stok <= 0 ? 'color:#b91c1c;font-weight:700' : ''}">${r.stok.toLocaleString('id-ID')}</td><td class="r">${r.laku || '—'}</td><td class="r">${r.dipesan || ''}</td>
      ${mitra() ? '' : `<td style="font-size:.7rem;color:${r.sendiri ? '#5b21b6' : '#94a3b8'}">${r.sendiri ? '🏠 stok sendiri (' + esc(r.sendiri) + ')' : 'HQ'}</td>`}
      <td class="r"><input type="number" min="0" inputmode="numeric" data-sku="${esc(r.k)}" value="${r.usul || ''}" placeholder="0" ${r.sendiri ? 'disabled' : ''} style="width:72px;text-align:right"></td></tr>`).join('')
      + (rows.length > 400 ? `<tr><td colspan="6" style="color:#64748b">… ${rows.length - 400} barang lain — pakai pencarian</td></tr>` : '')
      + (!rows.length ? '<tr><td colspan="6" style="color:#64748b">Tidak ada barang. Upload stok outlet dan laporan penjualan outlet dulu.</td></tr>' : '')
    // Riwayat pesanan
    document.getElementById('osPesanan').innerHTML = S.pesanan.length ? S.pesanan.slice(0, 50).map(p => { const st = STATUS[p.status] || [p.status, '#475569', '#f1f5f9']
      const tgl = p.dibuat_at?.toDate ? p.dibuat_at.toDate().toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }) : ''
      return `<div style="border:1.5px solid #e5e7eb;border-radius:12px;padding:10px 12px;margin-bottom:8px">
        <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap"><div><b>${tgl}</b> · ${(p.items || []).length} barang · ${(p.items || []).reduce((t, i) => t + (+i.qty || 0), 0)} pcs ${p.catatan ? `<span style="color:#64748b">· ${esc(p.catatan)}</span>` : ''}</div>
        <span style="font-size:.72rem;font-weight:700;color:${st[1]};background:${st[2]};border-radius:999px;padding:2px 10px">${st[0]}</span></div>
        ${p.alasan ? `<div style="font-size:.74rem;color:#991b1b">Alasan HQ: ${esc(p.alasan)}</div>` : ''}
        <details style="margin-top:4px"><summary style="cursor:pointer;font-size:.74rem;color:#475569">Rincian</summary><table style="margin-top:4px">${(p.items || []).map((it, j) => `<tr><td style="white-space:normal">${esc(it.nama || it.sku)}</td><td class="r">pesan ${it.qty}</td>${p.status !== 'diajukan' && it.kirim != null ? `<td class="r">kirim ${it.kirim}</td>` : ''}${p.status === 'dikirim' ? `<td class="r">terima <input type="number" min="0" id="osT-${p.id}-${j}" value="${it.kirim ?? it.qty}" style="width:64px;text-align:right"></td>` : it.terima != null ? `<td class="r">terima ${it.terima}</td>` : ''}</tr>`).join('')}</table></details>
        ${p.status === 'diajukan' ? `<button class="pbtn" style="margin-top:6px" onclick="osBatal('${p.id}')">Batalkan</button>` : ''}
        ${p.status === 'dikirim' ? `<button class="pbtn active" style="margin-top:6px" onclick="osTerima('${p.id}')">✅ Barang sudah diterima</button>` : ''}</div>` }).join('')
      : '<p style="font-size:.8rem;color:#64748b">Belum ada pesanan ke HQ.</p>'
  }

  window.osRender = () => render()
  window.renderCustomOs = () => renderCustom().catch(() => {})
  async function renderCustom() {
    const tb = document.getElementById('osCustom'); if (!tb || typeof _cuMuat !== 'function') return
    await _cuMuat(); const { bln, list } = _cuHitungStok(3), bulan = +document.getElementById('osBulan').value || 1, dp = sedangDipesan()
    window._osCu = list.map(x => { const sendiri = isSendiri(x.j + ' ' + x.contoh), k = 'CUSTOM:' + x.j, dipesan = dp[k] || 0; return { ...x, k, sendiri, dipesan, usul: sendiri ? 0 : Math.max(0, Math.ceil(x.avg * bulan) - dipesan) } })
    tb.innerHTML = !_cuKataStok().length ? '<tr><td style="color:#64748b">Owner belum mengisi kata kunci custom yang distok (Analitik HQ → 🧵 Barang Custom).</td></tr>'
      : !list.length ? `<tr><td style="color:#64748b">Belum ada penjualan custom yang cocok${bln.length ? '' : ' — upload laporan Item Penjualan outlet di tab 🧵'}.</td></tr>`
      : '<tr><th>Jenis custom</th><th class="r">Laku/bln</th><th class="r">Dipesan</th><th>Sumber</th><th class="r">Pesan ke HQ</th></tr>' + window._osCu.map(x => `<tr><td style="white-space:normal;min-width:180px"><b style="font-weight:600">${esc(x.j)}</b></td><td class="r">${x.avg}</td><td class="r">${x.dipesan || ''}</td><td style="font-size:.7rem;color:${x.sendiri ? '#5b21b6' : '#94a3b8'}">${x.sendiri ? '🏠 stok sendiri' : 'HQ'}</td>
        <td class="r"><input type="number" min="0" data-cu="${esc(x.k)}" data-nama="${esc(x.j)}" value="${x.usul || ''}" placeholder="0" ${x.sendiri ? 'disabled' : ''} style="width:72px;text-align:right"></td></tr>`).join('')
  }
  window.osKirimCustom = async () => {
    const items = [...document.querySelectorAll('#osCustom input[data-cu]')].map(el => ({ sku: el.dataset.cu, nama: el.dataset.nama + ' (custom)', qty: Math.max(0, parseInt(el.value) || 0), custom: true })).filter(x => x.qty > 0)
    if (!items.length) return alert('Isi qty custom yang mau dipesan ke HQ')
    if (!confirm(`Pesan custom ke HQ: ${items.length} jenis, ${items.reduce((t, i) => t + i.qty, 0)} pcs?`)) return
    try { await _app.firestore().collection('outlet_pesanan').add({ outlet: _OUTLET, outletNama: S.outlet.nama || _OUTLET, jenis: S.outlet.jenis || '', status: 'diajukan', items, catatan: 'Custom yang distok', dibuat_at: FV().serverTimestamp(), dibuat_oleh: _currentUserName || '' })
      await muat(); render(); renderCustom(); alert('✅ Pesanan custom terkirim ke HQ') } catch (e) { alert('Gagal: ' + e.message) }
  }
  window.osKirim = async () => {
    const items = [...document.querySelectorAll('#osTabel input[data-sku]')].map(el => ({ sku: el.dataset.sku, qty: Math.max(0, parseInt(el.value) || 0) })).filter(x => x.qty > 0).map(x => ({ ...x, nama: namaSku(x.sku) }))
    if (!items.length) return alert('Isi qty barang yang mau dipesan ke HQ')
    const catatan = document.getElementById('osCatatan').value.trim()
    if (!confirm(`Kirim pesanan ke HQ: ${items.length} barang, ${items.reduce((t, i) => t + i.qty, 0)} pcs?`)) return
    try {
      await _app.firestore().collection('outlet_pesanan').add({ outlet: _OUTLET, outletNama: S.outlet.nama || _OUTLET, jenis: S.outlet.jenis || '', status: 'diajukan', items, catatan,
        dibuat_at: FV().serverTimestamp(), dibuat_oleh: _currentUserName || '' })
      document.getElementById('osCatatan').value = ''; await muat(); render(); alert('✅ Pesanan terkirim ke HQ. Pantau statusnya di Riwayat pesanan.')
    } catch (e) { alert('Gagal mengirim pesanan: ' + e.message) }
  }
  window.osBatal = async id => { if (!confirm('Batalkan pesanan ini?')) return
    try { await _app.firestore().collection('outlet_pesanan').doc(id).update({ status: 'batal', batal_at: FV().serverTimestamp(), batal_oleh: _currentUserName || '' }); await muat(); render() } catch (e) { alert('Gagal: ' + e.message) } }
  window.osTerima = async id => {
    const p = S.pesanan.find(x => x.id === id); if (!p) return
    const items = (p.items || []).map((it, j) => ({ ...it, terima: Math.max(0, parseInt(document.getElementById(`osT-${id}-${j}`)?.value) || 0) }))
    if (!confirm(`Konfirmasi barang diterima (${items.reduce((t, i) => t + i.terima, 0)} pcs)? Stok outlet akan bertambah.`)) return
    try {
      const db = _app.firestore(), tambah = {}; items.forEach(it => { if (it.terima > 0) tambah[it.sku] = FV().increment(it.terima) })
      await db.collection('outlet_pesanan').doc(id).update({ status: 'diterima', items, diterima_at: FV().serverTimestamp(), diterima_oleh: _currentUserName || '' })
      if (Object.keys(tambah).length) await db.collection('outlet_stok').doc(_OUTLET).set({ stok: tambah }, { merge: true })
      await muat(); render()
    } catch (e) { alert('Gagal: ' + e.message) }
  }
  window.osUploadStok = async files => {
    const st = document.getElementById('osStatus'); if (!files?.length) return
    try {
      const stok = {}, info = {}
      for (const f of files) {
        const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' }), ws = wb.Sheets[wb.SheetNames[0]], rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
        const hi = rows.findIndex(r => (r || []).map(x => String(x).trim().toLowerCase()).includes('sku')); if (hi < 0) throw new Error(f.name + ': kolom "sku" tidak ditemukan — pakai file export produk Olsera')
        const h = rows[hi].map(x => String(x).trim().toLowerCase()), iS = h.indexOf('sku'), iQ = h.indexOf('stock_qty'), iN = h.indexOf('name'), iV = h.indexOf('variant_names'), iK = h.indexOf('category')
        if (iQ < 0) throw new Error(f.name + ': kolom "stock_qty" tidak ditemukan')
        rows.slice(hi + 1).forEach(r => { const k = String(r[iS] || '').trim(); if (!k) return; stok[k] = (stok[k] || 0) + (Number(r[iQ]) || 0); if (iN >= 0 && r[iN]) info[k] = [String(r[iN]).trim(), iV >= 0 ? String(r[iV] || '').trim() : '', iK >= 0 ? String(r[iK] || '').trim().toUpperCase() : ''] })
      }
      if (!Object.keys(stok).length) throw new Error('Tidak ada stok terbaca')
      st.textContent = '☁️ Menyimpan...'
      await _app.firestore().collection('outlet_stok').doc(_OUTLET).set({ outlet: _OUTLET, stok, info, updated_at: FV().serverTimestamp(), updated_by: _currentUserName || '' }, { merge: false })
      // set tanpa merge mengganti peta stok; tanda "produksi sendiri" dipertahankan
      if (Object.keys(S?.sendiri || {}).length) await _app.firestore().collection('outlet_stok').doc(_OUTLET).set({ sendiri: S.sendiri }, { merge: true })
      st.textContent = `✅ Stok outlet tersimpan: ${Object.keys(stok).length} SKU`; await muat(); render()
    } catch (e) { st.textContent = '⚠️ ' + e.message }
    finally { const el = document.getElementById('osFile'); if (el) el.value = '' }
  }

  function pasang() {
    const nav = document.querySelector('.nav'), main = document.querySelector('.main'); if (!nav || !main || document.getElementById('nav-ostok')) return
    const btn = document.createElement('button'); btn.id = 'nav-ostok'; btn.className = 'nav-btn'; btn.textContent = '📦 Stok & Pesan ke HQ'; btn.onclick = function () { showTab('ostok', this) }
    nav.appendChild(btn)
    const tab = document.createElement('div'); tab.id = 'tab-ostok'; tab.className = 'tab-content'
    tab.innerHTML = `<style>#tab-ostok table{width:100%;border-collapse:collapse;font-size:.78rem} #tab-ostok th{text-align:left;font-size:.68rem;color:#64748b;padding:6px 8px;border-bottom:2px solid rgba(3,69,67,.1);white-space:nowrap;position:sticky;top:0;background:#fff;z-index:1}
      #tab-ostok td{padding:6px 8px;border-bottom:1px solid rgba(3,69,67,.06);white-space:nowrap} #tab-ostok .r{text-align:right} #tab-ostok input,#tab-ostok select{padding:6px 8px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2);font-size:.8rem;min-width:0}</style>
    <div class="card"><div class="card-title">📦 Stok outlet</div>
      <p style="font-size:.76rem;color:#475569;margin-bottom:8px">Upload <b>export produk</b> dari akun Olsera outlet ini (kolom sku & stock_qty), sama seperti export stok di HQ.</p>
      <label class="pbtn active" style="cursor:pointer;display:inline-block">📤 Upload stok outlet<input type="file" id="osFile" accept=".xlsx,.xls" multiple style="display:none" onchange="osUploadStok(this.files)"></label>
      <span id="osStatus" style="font-size:.74rem;color:#64748b;margin-left:8px"></span></div>
    <div class="card"><div class="card-title">🛒 Pesan barang ke HQ</div>
      <p id="osInfo" style="font-size:.74rem;color:#475569;margin-bottom:8px"></p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px">
        <input id="osCari" type="search" placeholder="Cari barang / SKU…" oninput="clearTimeout(window._osT);window._osT=setTimeout(osRender,250)" style="flex:1;min-width:160px">
        <select id="osFilter" onchange="osRender()"><option value="pesan">Perlu dipesan</option><option value="kosong">Kosong & laku</option><option value="semua">Semua barang</option><option value="sendiri">Stok sendiri</option></select>
        <select id="osBulan" onchange="osRender();renderCustomOs()"><option value="1">Stok untuk 1 bulan</option><option value="2">2 bulan</option><option value="0.5">2 minggu</option></select>
      </div>
      <div style="max-height:60vh;overflow:auto"><table id="osTabel"></table></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:10px"><input id="osCatatan" placeholder="Catatan untuk HQ (opsional)" style="flex:1;min-width:180px"><button class="pbtn active" onclick="osKirim()">📨 Kirim pesanan ke HQ</button></div></div>
    <div class="card"><div class="card-title">🧢 Custom yang distok</div><p style="font-size:.74rem;color:#475569;margin-bottom:6px">Dari laporan <b>Item Penjualan</b> outlet ini (upload di tab 🧵 Barang Custom) yang cocok kata kunci Owner (mis. TOPI, DASI). Yang termasuk "stok sendiri" outlet ini tidak dipesan ke HQ.</p>
      <div style="max-height:50vh;overflow:auto"><table id="osCustom"></table></div>
      <div style="margin-top:8px"><button class="pbtn active" onclick="osKirimCustom()">📨 Pesan custom ini ke HQ</button></div></div>
    <div class="card"><div class="card-title">📋 Riwayat pesanan ke HQ</div><div id="osPesanan"></div></div>`
    main.appendChild(tab)
    const asli = window.showTab
    window.showTab = function (id, b) { asli(id, b); if (id === 'ostok') (S ? Promise.resolve() : muat()).then(() => { render(); renderCustom().catch(e => console.warn(e)) }).catch(e => { document.getElementById('osInfo').textContent = '⚠️ ' + e.message }) }
  }
  function mulai() { _fbReady.then(() => { if (_peranAnalitik === 'admin') pasang() }).catch(() => {}) }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mulai); else mulai()
  window._osMulai = mulai; window._osMuat = async () => { await muat(); render() }   // untuk uji
})()
