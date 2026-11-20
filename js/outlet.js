// Multi-outlet (cabang/titik) — Tahap 1.
// • Tiap outlet punya akun Olsera sendiri → staf outlet (Portal: apps.harmoni_analitik.outlet) upload laporan
//   "Penjualan berdasarkan SKU" outletnya; disimpan outlet_penjualan/{outlet}__{bulan}. Halaman memakai ?outlet=ID.
// • Ringkasan kecil per outlet outlet_ringkas/{outlet|'hq'} {b:{'Jun-26':{omzet,profit,qty,baris}}} → perbandingan
//   antar outlet tanpa memuat data mentah (tetap ringan untuk ribuan outlet).
// • Daftar outlet outlet/{id} {nama, jenis keluarga|mitra, tipe cabang|titik, aktif} dikelola Owner.
(function () {
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
  const rp = n => 'Rp ' + Math.round(n || 0).toLocaleString('id-ID')
  const rpJt = n => Math.abs(n) >= 1e6 ? 'Rp ' + (n / 1e6).toLocaleString('id-ID', { maximumFractionDigits: 1 }) + ' jt' : rp(n)
  const BLN = { jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6, jul: 7, agu: 8, agt: 8, ags: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12 }
  const urutBln = l => { const m = String(l).match(/^([A-Za-z]+)[-\s]?(\d{2,4})$/); if (!m) return 0; const y = +m[2] < 100 ? 2000 + +m[2] : +m[2]; return y * 100 + (BLN[m[1].slice(0, 3).toLowerCase()] || 0) }
  const idOutlet = n => String(n || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40)
  const JENIS = { keluarga: 'Keluarga (boleh produksi & belanja sendiri)', mitra: 'Mitra (semua barang dari HQ)' }
  let OUTLETS = []   // [{id, nama, jenis, tipe, aktif}]
  const adminHQ = () => _peranAnalitik === 'admin' && !_outletSaya && !_OUTLET

  // Ringkasan per bulan outlet aktif (HQ = 'hq') — dipanggil setelah upload / saat data cloud berubah
  window.simpanRingkasOutlet = async rows => {
    try {
      if (!_app || _peranAnalitik !== 'admin') return
      const b = {}; (rows || []).forEach(r => { const x = b[r.bulan] = b[r.bulan] || { omzet: 0, profit: 0, qty: 0, baris: 0 }; x.omzet += +r.omzet || 0; x.profit += +r.profit || 0; x.qty += +r.qty || 0; x.baris++ })
      await _app.firestore().collection('outlet_ringkas').doc(_OUTLET || 'hq').set({ outlet: _OUTLET || 'hq', b, updated_at: firebase.firestore.FieldValue.serverTimestamp(), updated_by: _currentUserName || '' }, { merge: true })
    } catch (e) { console.warn('outlet_ringkas', e) }
  }

  async function muatOutlet() {
    try { const s = await _app.firestore().collection('outlet').get(); OUTLETS = []; s.forEach(d => OUTLETS.push({ id: d.id, ...d.data() })); OUTLETS.sort((a, c) => String(a.nama).localeCompare(String(c.nama))) }
    catch (e) { console.warn('outlet', e); OUTLETS = [] }
    return OUTLETS
  }
  const namaOutlet = id => id === 'hq' ? 'HQ' : (OUTLETS.find(o => o.id === id)?.nama || id)

  // ── Mode outlet: label + sembunyikan fitur yang khusus HQ (stok HQ, rekomendasi, custom, Shopee) ──
  function modeOutlet() {
    const nama = namaOutlet(_OUTLET)
    const sub = document.getElementById('headerSub'); if (sub) sub.insertAdjacentHTML('beforebegin', `<div style="display:inline-block;background:#fef3c7;color:#92400e;border-radius:999px;padding:2px 10px;font-size:.72rem;font-weight:700;margin:4px 0">🏬 Outlet: ${esc(nama)}</div>`)
    ;['inventori', 'stok', 'custom', 'kanal'].forEach(id => { const b = document.getElementById('nav-' + id); if (b) b.style.display = 'none' })
    document.title = 'Analitik ' + nama
    const lbl = document.getElementById('globalUploadStatus'); if (lbl && !lbl.textContent) lbl.textContent = `Upload laporan Olsera "Penjualan berdasarkan SKU" akun outlet ${nama}`
  }

  // ── HQ: pilih lokasi + tab Outlet (kelola & bandingkan) ──
  function pasangPilihLokasi() {
    const bar = document.querySelector('.period-bar'); if (!bar || document.getElementById('pilihOutlet')) return
    const sel = document.createElement('select'); sel.id = 'pilihOutlet'; sel.title = 'Lihat analisa outlet lain'
    sel.style.cssText = 'padding:4px 10px;border-radius:20px;border:1.5px solid #f59e0b;font-size:.75rem;font-weight:700;color:#92400e;background:#fffbeb;cursor:pointer'
    sel.innerHTML = `<option value="">🏢 HQ</option>` + OUTLETS.filter(o => o.aktif !== false).map(o => `<option value="${esc(o.id)}">🏬 ${esc(o.nama)}</option>`).join('')
    sel.value = _OUTLET
    sel.onchange = () => { const u = new URL(location.href); if (sel.value) u.searchParams.set('outlet', sel.value); else u.searchParams.delete('outlet'); location.href = u.toString() }
    bar.insertBefore(sel, bar.firstChild)
  }
  function pasangTab() {
    const nav = document.querySelector('.nav'), main = document.querySelector('.main'); if (!nav || !main || document.getElementById('nav-outlet')) return
    const btn = document.createElement('button'); btn.id = 'nav-outlet'; btn.className = 'nav-btn'; btn.textContent = '🏬 Outlet'; btn.onclick = function () { showTab('outlet', this) }
    nav.appendChild(btn)
    const tab = document.createElement('div'); tab.id = 'tab-outlet'; tab.className = 'tab-content'
    const owner = !!window._isOwnerAn
    tab.innerHTML = `<style>#tab-outlet table{width:100%;border-collapse:collapse;font-size:.78rem} #tab-outlet th{text-align:left;font-size:.68rem;color:#64748b;padding:6px 8px;border-bottom:2px solid rgba(3,69,67,.1);white-space:nowrap;position:sticky;top:0;background:#fff}
      #tab-outlet td{padding:6px 8px;border-bottom:1px solid rgba(3,69,67,.06);white-space:nowrap} #tab-outlet .r{text-align:right} #tab-outlet input,#tab-outlet select{padding:6px 10px;border-radius:8px;border:1.5px solid rgba(3,69,67,.2);font-size:.8rem;min-width:0}</style>
    <div class="card"><div class="card-title">📊 Perbandingan outlet <select id="olPeriode" onchange="olRender()"><option value="3">3 bulan terakhir</option><option value="6">6 bulan terakhir</option><option value="12" selected>12 bulan terakhir</option><option value="0">Semua bulan</option></select></div>
      <p style="font-size:.76rem;color:#475569;margin-bottom:8px">Omzet tiap outlet dari laporan Olsera yang di-upload outlet masing-masing (akun outletnya sendiri). Klik nama outlet untuk melihat analisa lengkapnya.</p>
      <input id="olCari" type="search" placeholder="Cari outlet…" oninput="olRender()" style="margin-bottom:8px;width:100%;max-width:320px">
      <div style="max-height:65vh;overflow:auto"><table id="olTabel"></table></div></div>
    <div class="card"><div class="card-title">🏬 Daftar outlet <span id="olN" style="font-size:.72rem;color:#64748b;font-weight:500"></span></div>
      <div id="olDaftar"></div>
      ${owner ? `<details id="olForm" style="margin-top:10px"><summary style="cursor:pointer;font-weight:700;font-size:.82rem">+ Tambah / ubah outlet</summary>
        <input type="hidden" id="olId">
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:8px;margin-top:8px">
          <label style="font-size:.72rem;color:#475569">Nama outlet *<input id="olNama" placeholder="cth: Cabang Sudiang" style="width:100%"></label>
          <label style="font-size:.72rem;color:#475569">Tipe<select id="olTipe" style="width:100%"><option value="cabang">Cabang</option><option value="titik">Titik</option></select></label>
          <label style="font-size:.72rem;color:#475569">Jenis<select id="olJenis" style="width:100%"><option value="keluarga">${JENIS.keluarga}</option><option value="mitra">${JENIS.mitra}</option></select></label>
          <label style="font-size:.72rem;color:#475569">Pemilik / PJ<input id="olPemilik" placeholder="opsional" style="width:100%"></label>
        </div>
        <label style="font-size:.78rem;display:flex;gap:6px;align-items:center;margin-top:8px"><input type="checkbox" id="olAktif" checked style="width:auto"> Aktif</label>
        <div style="display:flex;gap:8px;margin-top:10px"><button class="pbtn active" onclick="olSimpan()">💾 Simpan outlet</button><button class="pbtn" onclick="olBatal()">Batal</button></div>
        <p style="font-size:.72rem;color:#64748b;margin-top:8px">Setelah outlet dibuat: di <b>Portal → Kelola Pengguna</b>, beri staf outlet akses Analitik dan pilih outletnya. Staf itu hanya melihat & meng-upload data outletnya.</p>
      </details>` : ''}</div>`
    main.appendChild(tab)
    const asli = window.showTab
    window.showTab = function (id, b) { asli(id, b); if (id === 'outlet') olRender() }
  }

  let RINGKAS = null
  async function muatRingkas(paksa) {
    if (RINGKAS && !paksa) return RINGKAS
    RINGKAS = {}
    try { (await _app.firestore().collection('outlet_ringkas').get()).forEach(d => { RINGKAS[d.id] = d.data().b || {} }) } catch (e) { console.warn('outlet_ringkas', e) }
    return RINGKAS
  }
  window.olRender = async () => {
    await muatRingkas()
    const n = +document.getElementById('olPeriode').value, q = document.getElementById('olCari').value.trim().toLowerCase()
    const semuaBln = [...new Set(Object.values(RINGKAS).flatMap(b => Object.keys(b)))].sort((a, c) => urutBln(a) - urutBln(c))
    const bln = n ? semuaBln.slice(-n) : semuaBln, terakhir = semuaBln[semuaBln.length - 1]
    const ids = ['hq', ...OUTLETS.map(o => o.id)].filter((id, i, a) => a.indexOf(id) === i)
    const baris = ids.map(id => { const b = RINGKAS[id] || {}, sum = k => bln.reduce((t, m) => t + (b[m]?.[k] || 0), 0)
      const o = OUTLETS.find(x => x.id === id), ada = bln.filter(m => b[m]).length
      return { id, nama: namaOutlet(id), o, omzet: sum('omzet'), profit: sum('profit'), qty: sum('qty'), ada, akhir: b[terakhir]?.omzet || 0 } })
      .filter(x => !q || x.nama.toLowerCase().includes(q)).sort((a, c) => c.omzet - a.omzet)
    const tot = baris.reduce((t, x) => t + x.omzet, 0)
    const enam = bln.slice(-6)
    document.getElementById('olTabel').innerHTML = `<tr><th>Outlet</th><th>Jenis</th>${enam.map(m => `<th class="r">${esc(m)}</th>`).join('')}<th class="r">Total ${bln.length} bln</th><th class="r">Porsi</th><th class="r">Profit</th><th class="r">Margin</th></tr>`
      + baris.slice(0, 300).map(x => { const b = RINGKAS[x.id] || {}
        return `<tr><td><a href="?${x.id === 'hq' ? '' : 'outlet=' + encodeURIComponent(x.id)}" style="color:var(--g);font-weight:700">${x.id === 'hq' ? '🏢' : '🏬'} ${esc(x.nama)}</a>${x.ada < bln.length ? `<div style="font-size:.66rem;color:#b45309">data ${x.ada}/${bln.length} bulan</div>` : ''}</td>
        <td style="font-size:.72rem">${x.id === 'hq' ? 'Pusat' : esc((x.o?.tipe === 'titik' ? 'Titik' : 'Cabang') + ' · ' + (x.o?.jenis === 'mitra' ? 'Mitra' : 'Keluarga'))}</td>
        ${enam.map(m => `<td class="r">${b[m] ? rpJt(b[m].omzet) : '<span style="color:#cbd5e1">—</span>'}</td>`).join('')}
        <td class="r" style="font-weight:700">${rpJt(x.omzet)}</td><td class="r">${tot ? (100 * x.omzet / tot).toFixed(1) + '%' : '—'}</td><td class="r">${rpJt(x.profit)}</td><td class="r">${x.omzet ? (100 * x.profit / x.omzet).toFixed(1) + '%' : '—'}</td></tr>` }).join('')
      + (baris.length > 300 ? `<tr><td colspan="${enam.length + 6}" style="color:#64748b">… ${baris.length - 300} outlet lain — pakai pencarian</td></tr>` : '')
      + (Object.keys(RINGKAS).length <= 1 ? `<tr><td colspan="${enam.length + 6}" style="color:#64748b;white-space:normal">Belum ada data outlet lain. Outlet mulai muncul setelah staf outlet meng-upload laporan Olsera-nya.</td></tr>` : '')
    document.getElementById('olN').textContent = `· ${OUTLETS.length} outlet`
    document.getElementById('olDaftar').innerHTML = OUTLETS.length ? `<table><tr><th>Nama</th><th>Tipe</th><th>Jenis</th><th>Pemilik / PJ</th><th>Status</th><th></th></tr>${OUTLETS.map(o => `<tr><td><b>${esc(o.nama)}</b> <span style="font-size:.66rem;color:#94a3b8">${esc(o.id)}</span></td><td>${o.tipe === 'titik' ? 'Titik' : 'Cabang'}</td><td>${o.jenis === 'mitra' ? 'Mitra' : 'Keluarga'}</td><td>${esc(o.pemilik || '')}</td><td>${o.aktif === false ? '<span style="color:#b91c1c">nonaktif</span>' : 'aktif'}</td><td>${window._isOwnerAn ? `<button class="pbtn" onclick="olUbah('${esc(o.id)}')">✏️</button>` : ''}</td></tr>`).join('')}</table>`
      : `<p style="font-size:.8rem;color:#64748b">Belum ada outlet.${window._isOwnerAn ? ' Tambahkan di bawah (cth: Cabang Sudiang, Harmonis Sentral).' : ''}</p>`
  }
  window.olUbah = id => { const o = OUTLETS.find(x => x.id === id); if (!o) return
    document.getElementById('olId').value = o.id; document.getElementById('olNama').value = o.nama || ''; document.getElementById('olTipe').value = o.tipe || 'cabang'
    document.getElementById('olJenis').value = o.jenis || 'keluarga'; document.getElementById('olPemilik').value = o.pemilik || ''; document.getElementById('olAktif').checked = o.aktif !== false
    const f = document.getElementById('olForm'); f.open = true; f.scrollIntoView({ behavior: 'smooth' }) }
  window.olBatal = () => { ['olId', 'olNama', 'olPemilik'].forEach(i => document.getElementById(i).value = ''); document.getElementById('olAktif').checked = true; document.getElementById('olForm').open = false }
  window.olSimpan = async () => {
    const nama = document.getElementById('olNama').value.trim(); if (!nama) return alert('Nama outlet wajib diisi')
    const id = document.getElementById('olId').value || idOutlet(nama)
    if (!id || id === 'hq') return alert('Nama outlet tidak valid')
    if (!document.getElementById('olId').value && OUTLETS.some(o => o.id === id)) return alert('Outlet dengan nama ini sudah ada')
    try {
      await _app.firestore().collection('outlet').doc(id).set({ nama, tipe: document.getElementById('olTipe').value, jenis: document.getElementById('olJenis').value,
        pemilik: document.getElementById('olPemilik').value.trim(), aktif: document.getElementById('olAktif').checked, updated_at: firebase.firestore.FieldValue.serverTimestamp(), updated_by: _currentUserName || '' }, { merge: true })
      olBatal(); await muatOutlet(); const s = document.getElementById('pilihOutlet'); if (s) s.remove(); pasangPilihLokasi(); olRender()
    } catch (e) { alert('Gagal menyimpan outlet: ' + (e.code === 'permission-denied' ? 'hanya Owner yang bisa mengelola outlet' : e.message)) }
  }

  function mulai() {
    _fbReady.then(async () => {
      await muatOutlet()
      if (_OUTLET) { modeOutlet(); if (!_outletSaya) pasangPilihLokasi(); return }
      if (adminHQ()) { pasangPilihLokasi(); pasangTab() }
    }).catch(() => {})
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mulai); else mulai()
  window._olMulai = mulai   // untuk uji
})()
