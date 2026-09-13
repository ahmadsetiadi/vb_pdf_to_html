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

   Elemen yang tidak aktif diberi atribut hidden (CSS: [hidden]{display:none});
   paginate.js melewati elemen hidden sehingga tidak masuk ke AllPages.html.
   ===================================================================== */
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
