/* =====================================================================
   bind.js — isi data dari VB ke template HTML
   Dipakai oleh: PageN.html / AllBody.html (preview di browser, data dari data.js)
                 dan paginate.js (di WebView2, data dari PageAssembler.Data).

   applyData(data, root)
     data : objek data, mis. { riders: ["FEC"], funds: [{fundname:"…", funddesc:"…", fundrisk:"…"}, …] }
            null / undefined = MODE TEMPLATE → semua blok tampil, baris berulang tampil apa adanya
     root : elemen / DocumentFragment tempat mencari (default: document)
     hasil: jumlah elemen [data-each] + [data-if] yang diproses

   1. Baris berulang (dibuat DirectiveProcessor.vb dari <<funds.fundname>> di sel tabel PDF):
        <tr data-each="funds"><td>{{fundname}}</td><td>{{funddesc}}</td>…</tr>
      Baris template di-clone satu per item array data["funds"]; {{field}} diganti nilai item
      (di-escape HTML; {{a.b}} = field bersarang; {{$no}} = nomor urut 1.., {{$index}} = 0..);
      hasil clone diberi atribut data-each-clone, baris template diberi hidden.
      applyData boleh dipanggil ulang: clone lama dibuang dulu.

   2. Blok kondisional (dari <<if …>> … <<endif>>):
        <div class="cond" data-if="riders.contains('FEC')"> … </div>
      Ekspresi dievaluasi sebagai JavaScript. Nama variabel = key di data.
        .contains(x)   → .includes(x)   (array atau string)
        kutip keriting → kutip lurus
        variabel tidak ada / error → false (dicatat di console)

   3. Variabel biasa di mana saja (header, footer, body): <<nama>> atau <<obj.field>>
      → diganti nilai dari data (resolvePath). Contoh: data.footer = {agentname:"Budi"} →
      <<footer.agentname>> menjadi "Budi". Placeholder yang tidak ada di data dibiarkan apa adanya
      (tetap merah = belum ada datanya). Kalau span pembungkusnya berwarna merah penanda (#E0241B),
      warna itu dibuang supaya nilainya tampil normal. <<page>> / <<totalpages>> diisi paginate.js.
      Semua otomatis: tambah key baru di data (RiplayData.Build di VB) → langsung dikenali di HTML.

   Elemen yang tidak aktif diberi atribut hidden (CSS: [hidden]{display:none});
   paginate.js melewati elemen hidden sehingga tidak masuk ke AllPages.html.
   ===================================================================== */
const PLACEHOLDER_RX = /<<\s*([A-Za-z_$][\w$]*(?:\.[\w$]+)*)\s*>>/g;
const PLACEHOLDER_RED = /^(#e0241b|rgb\(224,\s*36,\s*27\))$/i;

/* Ganti <<path>> di semua text node di bawah root. resolve(path) → string, atau undefined = biarkan.
   Mengembalikan jumlah placeholder yang diganti. Dipakai applyData (data) dan paginate.js (nomor halaman). */
function replacePlaceholders(root, resolve) {
  let n = 0;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode: t => (t.nodeValue.indexOf('<<') >= 0 && !/^(SCRIPT|STYLE)$/.test(t.parentNode.nodeName))
                     ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP });
  const nodes = [];
  for (let t = walker.nextNode(); t; t = walker.nextNode()) nodes.push(t);
  for (const t of nodes) {
    let hit = false;
    const out = t.nodeValue.replace(PLACEHOLDER_RX, (m, path) => {
      const v = resolve(path);
      if (v === undefined) return m;
      hit = true; n++;
      return v === null ? '' : String(v);
    });
    if (!hit) continue;
    t.nodeValue = out;
    const el = t.parentElement;
    if (el && PLACEHOLDER_RED.test(el.style.color)) el.style.color = '';
  }
  return n;
}

function applyData(data, root) {
  root = root || document;
  const templateMode = (data === null || data === undefined);
  let n = 0;

  // ---------- 1. baris berulang ----------
  for (const tpl of Array.from(root.querySelectorAll('[data-each]'))) {
    n++;
    // buang hasil clone dari applyData sebelumnya
    let sib = tpl.nextElementSibling;
    while (sib && sib.hasAttribute('data-each-clone')) { const nx = sib.nextElementSibling; sib.remove(); sib = nx; }
    if (templateMode) { tpl.hidden = false; continue; }

    const name = tpl.getAttribute('data-each');
    let items = resolvePath(data, name);
    if (!Array.isArray(items)) {
      console.warn('data-each "' + name + '": bukan array di data → 0 baris');
      items = [];
    }
    let after = tpl;
    items.forEach((item, i) => {
      const c = tpl.cloneNode(true);
      c.removeAttribute('data-each');
      c.setAttribute('data-each-clone', name);
      c.hidden = false;
      c.innerHTML = fillPlaceholders(c.innerHTML, item, i);
      after.after(c);
      after = c;
    });
    tpl.hidden = true;
  }

  // ---------- 2. blok kondisional ----------
  const els = Array.from(root.querySelectorAll('[data-if]'));
  if (root.nodeType === 1 && root.hasAttribute('data-if')) els.unshift(root);
  for (const el of els) {
    n++;
    const ok = templateMode ? true : evalCondition(el.getAttribute('data-if'), data);
    el.hidden = !ok;
  }

  // ---------- 3. variabel biasa <<nama>> / <<obj.field>> ----------
  if (!templateMode) {
    n += replacePlaceholders(root, path => {
      const v = resolvePath(data, path);
      if (v === undefined || v === null || typeof v === 'object') return undefined;   // array/objek bukan teks
      return String(v);
    });
  }
  return n;
}

// {{field}} / {{a.b}} / {{$no}} / {{$index}} → nilai item (di-escape HTML)
function fillPlaceholders(html, item, index) {
  return html.replace(/\{\{\s*([\w$.]+)\s*\}\}/g, (m, key) => {
    if (key === '$no') return String(index + 1);
    if (key === '$index') return String(index);
    const v = resolvePath(item, key);
    if (v === undefined) console.warn('{{' + key + '}}: field tidak ada di item');
    return escapeHtml(v === undefined || v === null ? '' : String(v));
  });
}

function resolvePath(obj, path) {
  return String(path).split('.').reduce((o, k) => (o === undefined || o === null) ? undefined : o[k], obj);
}

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function evalCondition(expr, data) {
  const src = String(expr)
    .replace(/[‘’‚‛]/g, "'")
    .replace(/[“”„‟]/g, '"')
    .replace(/\.contains\s*\(/gi, '.includes(');
  const keys = Object.keys(data).filter(k => /^[A-Za-z_$][\w$]*$/.test(k));
  try {
    const fn = new Function(...keys, '"use strict"; return (' + src + ');');
    return !!fn(...keys.map(k => data[k]));
  } catch (e) {
    console.warn('data-if "' + expr + '": ' + e.message + ' → false');
    return false;
  }
}
