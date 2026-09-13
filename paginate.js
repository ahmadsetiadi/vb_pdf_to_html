/* =====================================================================
   paginate.js — dijalankan di WebView2 oleh PageAssembler.vb

   Input (template di dokumen kerja):
     #tplHeader : isi header.html   → diambil elemen .hdr
     #tplFooter : isi footer.html   → diambil elemen .ftr
     #tplBody   : isi Page1.html    → diambil elemen .bdy (body panjang, belum dipecah)

   Cara pecah: blok diukur satu per satu di halaman yang sedang aktif.
   - muat                → lanjut
   - tidak muat, splittable (flow / ol / li / .b / table / tbody)
                         → dipecah ke anak-anaknya; wrapper ("shell") dibuat ulang di
                           halaman berikutnya (marker li kosong, colgroup + header tabel diulang)
   - tidak muat, .cols   → tiap kolom dipecah sendiri-sendiri; kotak kolom mengisi sisa
                           ruang halaman dan dibuat lagi di halaman berikutnya
   - tidak muat, atomik  → pindah ke halaman berikutnya
   Shell yang masih kosong di halaman lama dibuang (marker tidak yatim); judul (p.sub / .bar)
   ikut pindah bersama blok di bawahnya.

   data (opsional): objek data dari VB, mis. { riders: ["FEC"] } → diterapkan ke template
   lewat applyData() (bind.js) sebelum dipecah; blok [data-if] yang kondisinya salah menjadi
   hidden dan dilewati. Kalau tidak diberikan → pakai window.riplayData (data.js) kalau ada;
   kalau tidak ada juga → mode template (semua blok tampil).
   ===================================================================== */
function paginate(gapTopMm, gapBottomMm, data) {
  gapTopMm = gapTopMm || 0; gapBottomMm = gapBottomMm || 0;   // jarak kotak header→body, body→kotak footer
  const mm = v => parseFloat(v) || 0;
  const PX_PER_MM = 96 / 25.4;
  const tpl = id => document.getElementById(id).content;

  const hdrSrc = tpl('tplHeader').querySelector('.hdr');
  const ftrSrc = tpl('tplFooter').querySelector('.ftr');
  const bdySrc = tpl('tplBody').querySelector('.bdy');
  if (!hdrSrc || !ftrSrc || !bdySrc) throw new Error('Elemen .hdr / .ftr / .bdy tidak ditemukan');

  if (data === undefined) data = (window.riplayData === undefined) ? null : window.riplayData;
  if (typeof applyData === 'function') [hdrSrc, ftrSrc, bdySrc].forEach(r => applyData(data, r));

  // area body: diukur dari border kotak header (.hdr .box) dan kotak footer (.ftr .box)
  const probePg = document.createElement('div');
  probePg.className = 'page';
  probePg.appendChild(hdrSrc.cloneNode(true));
  probePg.appendChild(ftrSrc.cloneNode(true));
  document.body.appendChild(probePg);
  const pgTop = probePg.getBoundingClientRect().top;
  const hBox = probePg.querySelector('.hdr .box') || probePg.querySelector('.hdr');
  const fBox = probePg.querySelector('.ftr .box') || probePg.querySelector('.ftr');
  const hdrBottomMm = (hBox.getBoundingClientRect().bottom - pgTop) / PX_PER_MM;
  const ftrTopMm    = (fBox.getBoundingClientRect().top - pgTop) / PX_PER_MM;
  probePg.remove();

  const bodyTop = hdrBottomMm + gapTopMm;
  const bodyH   = (ftrTopMm - gapBottomMm) - bodyTop;

  const root = document.getElementById('pages');
  root.innerHTML = '';
  const pages = [];          // elemen .bdy tiap halaman
  let cur = null, curIdx = -1;
  const path = [];           // ancestor sumber yang sedang "terbuka"
  let shells = [];           // shell untuk tiap ancestor di halaman aktif
  const shellMap = new Map();// src → { pageIdx: shell }  (untuk cari shell di halaman tertentu)

  const SPLITTABLE = 'div.flow, div.cond, ol.lst, li, div.b, table.tbl, tbody';
  const KEEP_WITH_NEXT = 'p.sub, .bar';
  const isSplittable = n => n.matches(SPLITTABLE) && n.children.length > 0;

  // ---------- halaman ----------
  function newPage() {
    const pg = document.createElement('div');
    pg.className = 'page';
    pg.appendChild(hdrSrc.cloneNode(true));
    const b = document.createElement('div');
    b.className = 'bdy';
    b.style.top = bodyTop + 'mm';
    b.style.height = bodyH + 'mm';
    pg.appendChild(b);
    pg.appendChild(ftrSrc.cloneNode(true));
    root.appendChild(pg);
    pages.push(b);
    gotoPage(pages.length - 1);
  }
  function gotoPage(i) { cur = pages[i]; curIdx = i; if (!pagePlaced[i]) pagePlaced[i] = []; }
  function nextPage() { if (curIdx + 1 < pages.length) gotoPage(curIdx + 1); else newPage(); }

  const limit = () => cur.getBoundingClientRect().bottom + 0.5;          // batas bawah area body (px)
  const placedCount = () => cur.querySelectorAll('[data-placed]').length;

  // bawah terjauh dari elemen & semua turunannya (kolom berisi fixed height, dsb.)
  function bottomOf(el) {
    let b = el.getBoundingClientRect().bottom;
    for (const d of el.querySelectorAll('*')) {
      const r = d.getBoundingClientRect();
      if (r.height > 0 && r.bottom > b) b = r.bottom;
    }
    return b;
  }
  const fits = el => bottomOf(el) <= limit();
  const parentShell = () => shells.length ? shells[shells.length - 1] : cur;
  const isEmptyShell = sh => !sh.querySelector('[data-placed]');

  function remember(src, sh) {
    if (!shellMap.has(src)) shellMap.set(src, {});
    shellMap.get(src)[curIdx] = sh;
  }

  // ---------- shell: clone tanpa anak untuk elemen sumber di halaman aktif ----------
  function openShell(src, first) {
    let sh;
    if (src.matches('.cols') && shellMap.has(src) && shellMap.get(src)[curIdx]) {
      sh = shellMap.get(src)[curIdx];                   // cols yang sama sudah ada di halaman ini (kolom lain)
      shells.push(sh);
      return sh;
    }
    sh = src.cloneNode(false);
    if (!first) sh.style.marginTop = '0';
    if (src.tagName === 'LI') {                         // marker: asli di awal item, kosong di lanjutan (indentasi tetap)
      const m = src.querySelector(':scope > .m');
      if (m) {
        const mc = m.cloneNode(first);
        if (!first) mc.innerHTML = '&nbsp;';
        sh.appendChild(mc);
      }
    }
    if (src.tagName === 'TABLE') {                      // colgroup selalu ikut, header row diulang
      const cg = src.querySelector(':scope > colgroup');
      if (cg) sh.appendChild(cg.cloneNode(true));
      if (!first) {                                     // semua baris judul di awal tabel (bisa > 1 baris: rowspan/colspan)
        const tb = document.createElement('tbody');
        for (const tr of Array.from(src.rows)) {
          if (!tr.children.length || !Array.from(tr.children).every(c => c.tagName === 'TH')) break;
          tb.appendChild(tr.cloneNode(true));
        }
        if (tb.children.length) sh.appendChild(tb);
      }
    }
    if (src.matches('.col')) sh.style.height = '100%';  // kolom setinggi blok cols
    parentShell().appendChild(sh);
    if (src.matches('.cols')) {                         // cols mengisi sisa ruang halaman
      const top = sh.getBoundingClientRect().top;
      sh.style.height = ((limit() - 0.5 - top) / PX_PER_MM).toFixed(2) + 'mm';
    }
    shells.push(sh);
    remember(src, sh);
    return sh;
  }

  function breakPage() {
    // shell yang belum berisi apa pun di halaman lama → buang, nanti dibuat ulang utuh (marker ikut)
    const firstFlags = [];
    for (let i = shells.length - 1; i >= 0; i--) {
      const sh = shells[i];
      if (!sh.matches('.cols') && isEmptyShell(sh)) { sh.remove(); firstFlags[i] = true; }
      else firstFlags[i] = false;
    }
    nextPage();
    const srcs = path.slice();
    shells = [];
    srcs.forEach((s, i) => openShell(s, firstFlags[i] === true));
  }

  function placed(c, src) { c.dataset.placed = '1'; pagePlaced[curIdx].push({ src, clone: c }); }

  // ---------- judul yatim (keep-with-next) ----------
  // Urutan blok yang sudah ditaruh per halaman: {src, clone}. Saat harus pindah halaman, deretan judul
  // di dasar halaman (p.sub, .bar, baris <th> tabel) ikut dipindah ke atas halaman baru bersama blok berikutnya.
  const pagePlaced = [];
  const isHeadRow = n => n.tagName === 'TR' && n.children.length > 0 && Array.from(n.children).every(c => c.tagName === 'TH');
  // judul: p.sub / .bar / baris th, atau blok utuh yang seluruh teksnya judul (mis. <ol><li>I. <p class="sub">…) 
  function isHeading(n) {
    if (n.matches(KEEP_WITH_NEXT) || isHeadRow(n)) return true;
    if (n.matches('table, .cols, .page-break') || !n.querySelector(KEEP_WITH_NEXT)) return false;
    const w = document.createTreeWalker(n, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      if (!t.nodeValue.trim()) continue;
      const p = t.parentElement;
      if (!p.closest(KEEP_WITH_NEXT) && !p.closest('.m')) return false;
    }
    return true;
  }
  function trailingHeadings() {
    const list = pagePlaced[curIdx] || [];
    let i = list.length;
    while (i > 0 && isHeading(list[i - 1].src)) i--;
    if (i === list.length || i === 0) return [];          // tidak ada judul di dasar / halaman hanya berisi judul
    return list.slice(i);
  }
  // shell untuk src (tanpa mendaftarkannya ke rantai aktif)
  function makeShell(src, first) {
    const sh = src.cloneNode(false);
    if (!first) sh.style.marginTop = '0';
    if (src.tagName === 'LI') {
      const m = src.querySelector(':scope > .m');
      if (m) { const mc = m.cloneNode(first); if (!first) mc.innerHTML = '&nbsp;'; sh.appendChild(mc); }
    }
    if (src.tagName === 'TABLE') {
      const cg = src.querySelector(':scope > colgroup');
      if (cg) sh.appendChild(cg.cloneNode(true));
    }
    if (src.matches('.col')) sh.style.height = '100%';
    return sh;
  }
  // buang judul dari halaman lama; shell yang jadi kosong ikut dibuang (kecuali .cols, dipakai kolom lain)
  function removeCarried(items) {
    const list = pagePlaced[curIdx];
    for (const it of items) {
      list.splice(list.indexOf(it), 1);
      let el = it.clone.parentElement;
      it.clone.remove();
      while (el && el !== cur && !el.matches('.cols') && !el.querySelector('[data-placed]')) {
        const p = el.parentElement; el.remove(); el = p;
      }
    }
  }
  // taruh kembali judul di halaman baru, di bawah shell ancestor yang benar (dibuat ulang bila perlu)
  function placeCarried(items) {
    const made = new Map();                                 // src ancestor → shell baru di halaman ini
    for (const it of items) {
      const chain = [];                                     // ancestor antara src dan shell yang sudah ada
      let a = it.src.parentElement, parentSh = null, nextSh = null;
      while (a) {
        if (a === bdySrc) { parentSh = cur; break; }
        if (made.has(a)) { parentSh = made.get(a); break; }
        const pi = path.indexOf(a);
        if (pi >= 0) { parentSh = shells[pi]; nextSh = shells[pi + 1] || null; break; }
        chain.unshift(a); a = a.parentElement;
      }
      if (!parentSh) continue;
      for (const anc of chain) {
        const sh = makeShell(anc, true);
        if (anc === chain[0]) sh.style.marginTop = '0';
        parentSh.insertBefore(sh, nextSh); made.set(anc, sh);
        parentSh = sh; nextSh = null;
      }
      const k = it.src.cloneNode(true);
      if (!chain.length) k.style.marginTop = '0';
      parentSh.insertBefore(k, nextSh);
      placed(k, it.src);
    }
  }
  // pindah halaman; judul di dasar halaman lama dibawa serta
  function breakPageCarry() {
    const carry = trailingHeadings();
    if (carry.length) removeCarried(carry);
    breakPage();
    if (carry.length) placeCarried(carry);
  }

  // ---------- page break manual ----------
  //   <div class="page-break"></div>                → pindah halaman di titik ini (elemen tidak ikut dicetak)
  //   <p style="page-break-before:always">…</p>     → blok ini mulai di halaman baru
  //   <p style="break-before:page">…</p>            → sama
  const isBreakMarker = n => n.matches('.page-break, hr.page-break');
  const breaksBefore  = n => /(page-)?break-before\s*:\s*(always|page)/i.test(n.getAttribute('style') || '') ||
                             n.classList.contains('break-before');
  function forceBreak() { if (placedCount() > 0) breakPage(); }
  // blok yang mengandung page break di dalamnya tidak boleh ditaruh utuh → harus dipecah
  const hasBreakInside = n => !!n.querySelector('.page-break, .break-before, [style*="break-before"]');

  // ---------- blok biasa ----------
  function layout(node) {
    if (node.hidden) return;                                   // blok kondisional yang tidak aktif
    if (isBreakMarker(node)) { forceBreak(); return; }
    if (breaksBefore(node)) forceBreak();
    if (node.matches('.cols')) { layoutCols(node); return; }

    let c = node.cloneNode(true);
    const mustSplit = isSplittable(node) && hasBreakInside(node);
    if (!mustSplit) {
      parentShell().appendChild(c);
      if (fits(c)) { placed(c, node); return; }
      parentShell().removeChild(c);
    }

    if (isSplittable(node)) {
      path.push(node);
      openShell(node, true);
      for (const k of Array.from(node.children)) {
        if (node.tagName === 'LI' && k.matches('.m')) continue;                // sudah di shell
        if (node.tagName === 'TABLE' && k.tagName === 'COLGROUP') continue;   // sudah di shell
        layout(k);
      }
      path.pop();
      shells.pop();
      return;
    }

    // atomik: halaman masih kosong → taruh saja (blok lebih tinggi dari halaman)
    if (placedCount() === 0) { parentShell().appendChild(c); placed(c, node); return; }

    // judul di dasar halaman (p.sub / .bar / baris th) ikut pindah bersama blok ini
    breakPageCarry();
    c = node.cloneNode(true);
    parentShell().appendChild(c);
    placed(c, node);
  }

  // ---------- blok kolom: tiap kolom dipecah sendiri, mulai dari halaman yang sama ----------
  function layoutCols(node) {
    const cols = Array.from(node.children).filter(c => c.matches('.col'));

    // coba utuh dulu (kolom dengan height tetap dibiarkan apa adanya) — kecuali ada page break di dalamnya
    if (!hasBreakInside(node)) {
      const c = node.cloneNode(true);
      parentShell().appendChild(c);
      if (fits(c)) { placed(c, node); return; }
      parentShell().removeChild(c);
    }

    // sisa ruang terlalu kecil (< 25mm) → mulai di halaman baru
    const probe = document.createElement('div');
    parentShell().appendChild(probe);
    const colsTop = probe.getBoundingClientRect().top;
    probe.remove();
    if (placedCount() > 0 && (limit() - colsTop) < 25 * PX_PER_MM) breakPage();

    let startIdx = curIdx;
    path.push(node);
    openShell(node, true);
    const base = shells.slice();
    let maxIdx = startIdx;

    for (const col of cols) {
      gotoPage(startIdx);
      shells = base.slice();
      path.push(col);
      openShell(col, true);
      for (const k of Array.from(col.children)) layout(k);
      path.pop(); shells.pop();
      if (curIdx > maxIdx) maxIdx = curIdx;
    }

    // semua kolom pindah dari halaman awal (judulnya yatim) → cols kosong di halaman awal dibuang
    if (maxIdx > startIdx && isEmptyShell(shellMap.get(node)[startIdx])) {
      shellMap.get(node)[startIdx].remove();
      delete shellMap.get(node)[startIdx];
      startIdx++;
    }

    // kolom yang isinya sudah habis / dipindah tetap diberi kotak kosong (border lanjut, posisi kolom lain tetap)
    for (let idx = startIdx; idx <= maxIdx; idx++) {
      const colsSh = shellMap.get(node)[idx];
      cols.forEach((col, ci) => {
        if (shellMap.has(col) && shellMap.get(col)[idx] && shellMap.get(col)[idx].isConnected) return;
        const sh = col.cloneNode(false);
        sh.style.height = '100%';
        if (idx > startIdx) sh.style.marginTop = '0';
        colsSh.insertBefore(sh, colsSh.children[ci] || null);
        if (!shellMap.has(col)) shellMap.set(col, {});
        shellMap.get(col)[idx] = sh;
      });
    }

    path.pop();                                          // cols
    gotoPage(maxIdx);
    shells = path.map(s => shellMap.get(s)[maxIdx]);     // rantai shell ancestor di halaman terakhir
    const last = shellMap.get(node)[maxIdx];
    last.dataset.placed = '1';

    // kalau masih ada blok setelah cols → cols terakhir dipendekkan sampai isi terbawah
    if (node.nextElementSibling) {
      let b = last.getBoundingClientRect().top;
      for (const d of last.querySelectorAll('[data-placed]')) {
        const r = d.getBoundingClientRect();
        if (r.bottom > b) b = r.bottom;
      }
      last.style.height = ((b - last.getBoundingClientRect().top) / PX_PER_MM + 3).toFixed(2) + 'mm';
    }
  }

  newPage();
  for (const k of Array.from(bdySrc.children)) layout(k);

  root.querySelectorAll('[data-placed]').forEach(e => delete e.dataset.placed);
  root.querySelectorAll('[hidden]').forEach(e => e.remove());

  // nomor halaman: <<page>> / <<totalpages>> (template PDF), {{PAGE}} / {{TOTAL}}, atau "dari <b>N</b>"
  const total = pages.length;
  pages.forEach((b, i) => {
    const f = b.parentElement.querySelector('.ftr');
    let done = 0;
    if (typeof replacePlaceholders === 'function')
      done = replacePlaceholders(f, path => {
        const k = path.toLowerCase();
        if (k === 'page' || k === 'pageno') return String(i + 1);
        if (k === 'totalpages' || k === 'total' || k === 'pages') return String(total);
        return undefined;
      });
    let h = f.innerHTML.replace(/\{\{PAGE\}\}/g, String(i + 1));
    if (h.indexOf('{{TOTAL}}') >= 0) h = h.replace(/\{\{TOTAL\}\}/g, String(total));
    else if (!done) h = h.replace(/(dari\s*<b>)\d+(<\/b>)/, '$1' + total + '$2');
    if (h !== f.innerHTML) f.innerHTML = h;
  });

  return JSON.stringify({ pages: total, html: root.innerHTML });
}
