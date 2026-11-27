// Foto barang (thumbnail) di tabel produk — supaya jelas barang yang mana.
// Disimpan di Firestore harmoni-indonesia: foto_produk/{kunci} { foto: dataURL JPEG kecil, produk, varian, oleh, at }
// Kunci: nama produk (+ "__" + varian bila fotonya khusus varian). Tampil: foto varian dulu, kalau tidak ada foto produk.
(function () {
  const cache = new Map()            // kunci → dataURL | null
  let antre = new Set(), timer = null
  const norm = s => String(s || '').trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 300)
  const kunciProduk = p => norm(p)
  const kunciVarian = (p, v) => norm(p) + '__' + norm(v)
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const bolehUbah = () => typeof _peranAnalitik !== 'undefined' && _peranAnalitik === 'admin'
  const col = () => _app.firestore().collection('foto_produk')

  // Sel tabel: kotak kecil, diisi setelah foto dimuat
  window.fotoSel = function (produk, varian) {
    return `<td class="foto-sel" style="width:46px;padding:4px 6px"><div class="fp-box" data-p="${esc(produk)}" data-v="${esc(varian || '')}" onclick="fotoBuka(this)" title="Foto barang"
      style="width:38px;height:38px;border-radius:8px;background:#f1f5f9;border:1px dashed #cbd5e1;display:flex;align-items:center;justify-content:center;font-size:14px;color:#94a3b8;cursor:pointer;overflow:hidden;background-size:cover;background-position:center">📷</div></td>`
  }
  // Versi kecil di dalam sel nama produk (tabel dengan kolom sticky)
  window.fotoMini = function (produk, varian) {
    return `<div class="fp-box" data-p="${esc(produk)}" data-v="${esc(varian || '')}" onclick="fotoBuka(this)" title="Foto barang"
      style="display:inline-flex;vertical-align:middle;margin-right:6px;width:28px;height:28px;border-radius:6px;background:#f1f5f9;border:1px dashed #cbd5e1;align-items:center;justify-content:center;font-size:11px;color:#94a3b8;cursor:pointer;overflow:hidden;background-size:cover;background-position:center">📷</div>`
  }
  function pasang(el) {
    const kv = kunciVarian(el.dataset.p, el.dataset.v), kp = kunciProduk(el.dataset.p)
    const f = (el.dataset.v && cache.get(kv)) || cache.get(kp)
    if (f) { el.style.backgroundImage = `url("${f}")`; el.style.border = '1px solid #e2e8f0'; el.textContent = '' }
    else { el.style.backgroundImage = ''; el.style.border = '1px dashed #cbd5e1'; el.textContent = bolehUbah() ? '📷' : '' }
  }
  async function muat() {
    timer = null
    const ids = [...antre].filter(k => !cache.has(k)); antre = new Set()
    for (let i = 0; i < ids.length; i += 10) {
      const part = ids.slice(i, i + 10)
      try {
        const snap = await col().where(firebase.firestore.FieldPath.documentId(), 'in', part).get()
        part.forEach(k => cache.set(k, null)); snap.forEach(d => cache.set(d.id, d.data().foto || null))
      } catch (e) { part.forEach(k => cache.set(k, null)) }
    }
    document.querySelectorAll('.fp-box').forEach(pasang)
  }
  // Panggil setelah tabel digambar
  window.fotoIsi = function (root) {
    if (typeof _app === 'undefined' || !_app) return
    (root || document).querySelectorAll('.fp-box').forEach(el => {
      const kv = kunciVarian(el.dataset.p, el.dataset.v), kp = kunciProduk(el.dataset.p)
      if (el.dataset.v && !cache.has(kv)) antre.add(kv)
      if (!cache.has(kp)) antre.add(kp)
      pasang(el)
    })
    if (antre.size && !timer) timer = setTimeout(muat, 150)
  }

  // Kecilkan foto: sisi terpanjang 480px, JPEG — cukup jelas, ukuran kecil (±30–60 KB)
  function kecilkan(file) {
    return new Promise((ok, gagal) => {
      const r = new FileReader()
      r.onload = () => { const img = new Image(); img.onload = () => {
        const s = Math.min(1, 480 / Math.max(img.width, img.height)), c = document.createElement('canvas')
        c.width = Math.round(img.width * s); c.height = Math.round(img.height * s)
        const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height)
        ok(c.toDataURL('image/jpeg', 0.78)) }; img.onerror = gagal; img.src = r.result }
      r.onerror = gagal; r.readAsDataURL(file)
    })
  }

  // Jendela foto: lihat besar; Owner/Admin bisa tambah, ganti, hapus
  window.fotoBuka = function (el) {
    const produk = el.dataset.p, varian = el.dataset.v
    const kv = kunciVarian(produk, varian), kp = kunciProduk(produk)
    const fV = varian ? cache.get(kv) : null, fP = cache.get(kp), f = fV || fP
    if (!f && !bolehUbah()) return
    const bg = document.createElement('div')
    bg.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.6);z-index:9999;display:flex;align-items:center;justify-content:center;padding:14px'
    bg.innerHTML = `<div style="background:#fff;border-radius:14px;max-width:420px;width:100%;max-height:92vh;overflow:auto;padding:14px;font-size:.85rem">
      <div style="display:flex;justify-content:space-between;gap:8px;align-items:flex-start;margin-bottom:8px"><div><b>${esc(produk)}</b>${varian ? ` <span style="color:#7c3aed">· ${esc(varian)}</span>` : ''}
        ${f ? `<div style="font-size:.72rem;color:#64748b">${fV ? 'Foto khusus varian ini' : 'Foto produk (semua varian)'}</div>` : ''}</div>
        <button type="button" data-x style="border:1.5px solid #94a3b8;background:#f1f5f9;border-radius:8px;padding:6px 12px;font-weight:700;cursor:pointer">✕ Tutup</button></div>
      ${f ? `<img src="${f}" style="width:100%;border-radius:10px;display:block;margin-bottom:10px">` : '<div style="padding:30px;text-align:center;color:#94a3b8;background:#f8fafc;border-radius:10px;margin-bottom:10px">Belum ada foto</div>'}
      ${bolehUbah() ? `<div style="display:grid;gap:8px">
        ${varian ? `<label style="display:flex;gap:6px;align-items:center"><input type="radio" name="fpScope" value="v" ${fV ? 'checked' : ''}> Untuk varian <b>${esc(varian)}</b> saja</label>
        <label style="display:flex;gap:6px;align-items:center"><input type="radio" name="fpScope" value="p" ${fV ? '' : 'checked'}> Untuk semua varian ${esc(produk)}</label>` : ''}
        <div style="display:flex;gap:8px;flex-wrap:wrap"><label style="flex:1;min-width:140px;text-align:center;background:#034543;color:#fff;border-radius:8px;padding:9px;font-weight:700;cursor:pointer">📷 ${f ? 'Ganti foto' : 'Pilih / ambil foto'}<input type="file" accept="image/*" data-file style="display:none"></label>
          <button type="button" data-tempel style="flex:1;min-width:120px;border:1.5px solid #034543;background:#fff;color:#034543;border-radius:8px;padding:9px;font-weight:700;cursor:pointer">📋 Tempel foto</button></div>
        ${f ? `<button type="button" data-hapus style="border:none;background:#fee2e2;color:#b91c1c;border-radius:8px;padding:8px;font-weight:700;cursor:pointer">Hapus ${fV ? 'foto varian ini' : 'foto produk'}</button>` : ''}
        <div data-status style="font-size:.75rem;color:#64748b"></div></div>` : ''}</div>`
    document.body.appendChild(bg)
    const tutup = () => bg.remove()
    bg.onclick = e => { if (e.target === bg) tutup() }
    bg.querySelector('[data-x]').onclick = tutup
    if (!bolehUbah()) return
    const st = bg.querySelector('[data-status]')
    const tujuan = () => (varian && bg.querySelector('input[name=fpScope]:checked')?.value === 'v') ? kv : kp
    const simpan = async file => {
      try {
        st.textContent = '⏳ Menyimpan…'
        const foto = await kecilkan(file), k = tujuan()
        await col().doc(k).set({ foto, produk, varian: k === kv ? varian : '', oleh: _currentUserName || '', at: firebase.firestore.FieldValue.serverTimestamp() })
        cache.set(k, foto); document.querySelectorAll('.fp-box').forEach(pasang); tutup()
      } catch (e) { st.textContent = '❌ Gagal menyimpan: ' + e.message }
    }
    bg.querySelector('[data-file]').onchange = e => { const f = e.target.files[0]; if (f) simpan(f) }
    bg.querySelector('[data-tempel]').onclick = async () => {
      try { for (const it of await navigator.clipboard.read()) { const t = it.types.find(x => x.startsWith('image/')); if (t) return simpan(new File([await it.getType(t)], 'foto.png', { type: t })) } st.textContent = 'Tidak ada gambar di clipboard' }
      catch (e) { st.textContent = 'Tidak bisa menempel — pakai Pilih foto' }
    }
    const h = bg.querySelector('[data-hapus]')
    if (h) h.onclick = async () => {
      const k = fV ? kv : kp
      if (!confirm(`Hapus ${fV ? 'foto varian ini' : 'foto produk ' + produk}?`)) return
      try { await col().doc(k).delete(); cache.set(k, null); document.querySelectorAll('.fp-box').forEach(pasang); tutup() }
      catch (e) { st.textContent = '❌ Gagal menghapus: ' + e.message }
    }
  }
})()
