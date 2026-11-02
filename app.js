'use strict';

// Personel Rehberi
// Excel dosyası telefonda seçilir, okunur ve yalnızca bu cihazın tarayıcı deposunda (IndexedDB) saklanır.
// Uygulamanın kendisinde hiçbir kişisel veri yoktur; bu yüzden kodu internete koymak güvenlidir.

const LIBS = [
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
];
const AUTO_LOCK_MS = 60 * 1000;
// PIN için gereken şifreleme yalnızca https:// veya localhost'ta vardır (ör. http://192.168... adresinde yoktur).
const CAN_PIN = !!(window.isSecureContext && window.crypto?.subtle);

// ---------- Depolama (IndexedDB) ----------

const DB = (() => {
  let opening;
  const open = () => opening ??= new Promise((ok, fail) => {
    const r = indexedDB.open('personel-rehberi', 1);
    r.onupgradeneeded = () => {
      r.result.createObjectStore('kv');
      r.result.createObjectStore('images');
    };
    r.onsuccess = () => ok(r.result);
    r.onerror = () => fail(r.error);
  });
  const req = (r) => new Promise((ok, fail) => { r.onsuccess = () => ok(r.result); r.onerror = () => fail(r.error); });
  const done = (tx) => new Promise((ok, fail) => { tx.oncomplete = () => ok(); tx.onerror = tx.onabort = () => fail(tx.error); });

  return {
    async get(store, key) {
      const db = await open();
      return req(db.transaction(store).objectStore(store).get(key));
    },
    async set(store, key, value) {
      const db = await open();
      const tx = db.transaction(store, 'readwrite');
      tx.objectStore(store).put(value, key);
      return done(tx);
    },
    async del(store, key) {
      const db = await open();
      const tx = db.transaction(store, 'readwrite');
      tx.objectStore(store).delete(key);
      return done(tx);
    },
    async entries(store) {
      const db = await open();
      const os = db.transaction(store).objectStore(store);
      const [keys, values] = await Promise.all([req(os.getAllKeys()), req(os.getAll())]);
      return keys.map((k, i) => [k, values[i]]);
    },
    async replaceImages(map) {
      const db = await open();
      const tx = db.transaction('images', 'readwrite');
      const os = tx.objectStore('images');
      os.clear();
      for (const [k, v] of map) os.put(v, k);
      return done(tx);
    },
    async wipe() {
      const db = await open();
      const tx = db.transaction(['kv', 'images'], 'readwrite');
      tx.objectStore('kv').clear();
      tx.objectStore('images').clear();
      return done(tx);
    },
  };
})();

// ---------- Durum ----------

const state = {
  people: [],
  meta: null,          // { fileName, importedAt, attachments: { sayfaAdı: [resimAnahtarı] } }
  images: new Map(),   // anahtar -> { type, buf }
  notes: new Map(),    // kişi id -> kullanıcının notu
  pinHash: null,
  locked: false,
  query: '',
  filter: 'all',
  listScroll: 0,
  hiddenAt: 0,
};

const app = document.getElementById('app');
const urlCache = new Map();

function imgUrl(key) {
  if (!key) return null;
  if (urlCache.has(key)) return urlCache.get(key);
  const img = state.images.get(key);
  if (!img) return null;
  const url = URL.createObjectURL(new Blob([img.buf], { type: img.type }));
  urlCache.set(key, url);
  return url;
}

function clearUrlCache() {
  for (const u of urlCache.values()) URL.revokeObjectURL(u);
  urlCache.clear();
}

// ---------- Yardımcılar ----------

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

const TR_MAP = { ç: 'c', ğ: 'g', ı: 'i', ö: 'o', ş: 's', ü: 'u', â: 'a', î: 'i', û: 'u' };
const norm = (s) => String(s ?? '').toLocaleLowerCase('tr').replace(/[çğıöşüâîû]/g, (c) => TR_MAP[c]);

const digits = (s) => String(s ?? '').replace(/\D/g, '');

function initials(name) {
  const w = String(name).trim().split(/\s+/);
  return ((w[0]?.[0] ?? '') + (w.length > 1 ? w[w.length - 1][0] : '')).toLocaleUpperCase('tr');
}

function maskTc(tc) {
  const d = digits(tc);
  return d.length === 11 ? d.slice(0, 3) + '•••••' + d.slice(8) : tc;
}

function durumKey(d) {
  const n = norm(d).trim().replace(/\s+/g, ' ');
  return n === 'silasiz' ? 'silahsiz' : n;
}

function durumClass(d) {
  const k = durumKey(d);
  if (k === 'silahsiz') return 'ok';
  if (k === 'silahli') return 'warn';
  return '';
}

function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove('show'), 2400);
}

async function sha256(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}
const pinHash = (pin) => sha256('personel-rehberi:' + pin);

function loadScript(src) {
  return new Promise((ok, fail) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = ok;
    s.onerror = () => fail(new Error('Gerekli kütüphane indirilemedi. İnternet bağlantınızı kontrol edin.'));
    document.head.appendChild(s);
  });
}
async function loadLibs() {
  if (!window.XLSX) await loadScript(LIBS[0]);
  if (!window.JSZip) await loadScript(LIBS[1]);
}

const ICON = {
  back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 18l-6-6 6-6"/></svg>',
  chev: '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg>',
  search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>',
  gear: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
  lock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>',
  phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1.9.4 1.8.7 2.7a2 2 0 0 1-.5 2.1L8 9.8a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.7.7a2 2 0 0 1 1.7 2z"/></svg>',
  people: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8"/></svg>',
  map: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 10c0 7-9 13-9 13S3 17 3 10a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>',
  close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6L6 18M6 6l12 12"/></svg>',
};

function avatar(p, cls = 'avatar') {
  const u = imgUrl(p.photo);
  return u
    ? `<img class="${cls}" src="${u}" alt="" loading="lazy">`
    : `<div class="${cls}">${esc(initials(p.name))}</div>`;
}

// ---------- Excel okuma ----------

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Kişi kartındaki etiketler. Sıra önemli: "YAKININ TELEFONU", "YAKINI"dan önce aranmalı.
const LABELS = [
  ['tc', 'TC K'],
  ['dogum', 'DOĞUM'],
  ['tel', 'KENDİ TELEFONU'],
  ['kan', 'KAN GRUBU'],
  ['durum', 'DURUM'],
  ['gorev', 'GÖREV YERİ'],
  ['yakinTel', 'YAKININ TELEFONU'],
  ['yakin', 'YAKINI'],
  ['adres', 'ADRES'],
  ['kayitlar', 'TUTANAK'],
  ['notlar', 'NOTLAR'],
];

function resolvePath(fromPart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = fromPart.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function cellText(c) {
  if (!c || c.v == null) return '';
  if (c.t === 'n') {
    if (c.z && XLSX.SSF.is_date(c.z)) {
      const d = XLSX.SSF.parse_date_code(c.v);
      return [d.d, d.m].map((x) => String(x).padStart(2, '0')).join('.') + '.' + d.y;
    }
    return String(c.v);
  }
  if (c.t === 'b') return c.v ? 'Evet' : 'Hayır';
  return String(c.v).trim();
}

function linkedSheet(c) {
  if (!c) return null;
  const target = c.l?.Target || (c.f && (c.f.match(/HYPERLINK\(\s*"([^"]+)"/i) || [])[1]);
  if (!target || !target.startsWith('#')) return null;
  const m = target.slice(1).match(/^'?(.+?)'?!/);
  return m ? m[1].replace(/''/g, "'") : null;
}

function parsePerson(ws, sheetName) {
  const enc = XLSX.utils.encode_cell;
  const at = (r, c) => ws[enc({ r, c })];
  const cells = Object.keys(ws)
    .filter((a) => a[0] !== '!')
    .map((a) => ({ ...XLSX.utils.decode_cell(a), cell: ws[a] }))
    .sort((x, y) => x.r - y.r || x.c - y.c);

  const pos = {};
  let no = null;
  for (const { r, c, cell } of cells) {
    const text = cellText(cell);
    if (!text) continue;
    const up = text.toLocaleUpperCase('tr');
    const m = text.match(/Kayıt No:\s*(\d+)/i);
    if (m) no = Number(m[1]);
    for (const [key, label] of LABELS) {
      if (up.includes(label)) {
        if (!pos[key]) pos[key] = { r, c };
        break;
      }
    }
  }
  if (!pos.tc) return null; // Kişi kartı değil (ör. tutanak sayfası)

  const below = (key) => (pos[key] ? cellText(at(pos[key].r + 1, pos[key].c)) : '');
  const p = {
    id: sheetName,
    no,
    name: cellText(at(3, 1)) || sheetName,
    photo: null,
    kayitlar: [],
    notlar: '',
  };
  for (const [key] of LABELS) if (key !== 'kayitlar' && key !== 'notlar') p[key] = below(key);

  // Tutanak / izin tablosu: başlık satırındaki TARİH / TÜR / AÇIKLAMA sütunlarını bul.
  if (pos.kayitlar) {
    const hr = pos.kayitlar.r + 1;
    const end = pos.notlar ? pos.notlar.r : hr + 30;
    const col = { tarih: 1, tur: 3, aciklama: 5 };
    for (const { r, c, cell } of cells) {
      if (r !== hr) continue;
      const up = cellText(cell).toLocaleUpperCase('tr');
      if (up.includes('TARİH')) col.tarih = c;
      else if (up.includes('TÜR')) col.tur = c;
      else if (up.includes('AÇIKLAMA')) col.aciklama = c;
    }
    for (let r = hr + 1; r < end; r++) {
      const ac = at(r, col.aciklama);
      const link = linkedSheet(ac);
      const rec = {
        tarih: cellText(at(r, col.tarih)),
        tur: cellText(at(r, col.tur)),
        aciklama: link ? '' : cellText(ac),
        link,
      };
      if (rec.tarih || rec.tur || rec.aciklama || rec.link) p.kayitlar.push(rec);
    }
  }

  if (pos.notlar) {
    p.notlar = cells
      .filter(({ r }) => r > pos.notlar.r)
      .map(({ cell }) => cellText(cell))
      .filter(Boolean)
      .join('\n');
  }
  return p;
}

async function readWorkbook(file, progress) {
  progress(5, 'Kütüphaneler hazırlanıyor…');
  await loadLibs();
  progress(15, 'Dosya okunuyor…');
  const buf = await file.arrayBuffer();
  await new Promise((r) => setTimeout(r, 30));
  const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true, cellHTML: false });
  progress(45, 'Fotoğraflar ayıklanıyor…');
  const zip = await JSZip.loadAsync(buf);

  const xml = async (path) => {
    const f = zip.file(path);
    return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null;
  };
  const relsOf = async (part) => {
    const i = part.lastIndexOf('/');
    const doc = await xml(part.slice(0, i) + '/_rels/' + part.slice(i + 1) + '.rels');
    const out = {};
    if (doc) {
      for (const r of doc.getElementsByTagName('Relationship')) {
        out[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: resolvePath(part, r.getAttribute('Target')) };
      }
    }
    return out;
  };

  const wbRels = await relsOf('xl/workbook.xml');
  const sheetPath = {};
  for (const s of (await xml('xl/workbook.xml')).getElementsByTagName('sheet')) {
    const rid = s.getAttributeNS(NS_R, 'id') || s.getAttribute('r:id');
    if (wbRels[rid]) sheetPath[s.getAttribute('name')] = wbRels[rid].target;
  }

  const imagesOf = async (sheetName) => {
    const path = sheetPath[sheetName];
    if (!path) return [];
    const out = [];
    for (const rel of Object.values(await relsOf(path))) {
      if (!rel.type.endsWith('/drawing')) continue;
      const drawRels = await relsOf(rel.target);
      const doc = await xml(rel.target);
      if (!doc) continue;
      for (const blip of doc.getElementsByTagNameNS(NS_A, 'blip')) {
        const target = drawRels[blip.getAttributeNS(NS_R, 'embed') || blip.getAttribute('r:embed')]?.target;
        if (target && !out.includes(target)) out.push(target);
      }
    }
    return out;
  };

  const people = [];
  const sheetImages = {};
  const names = wb.SheetNames;
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    sheetImages[name] = await imagesOf(name);
    const p = parsePerson(wb.Sheets[name], name);
    if (p) {
      p.photo = sheetImages[name][0] || null;
      people.push(p);
    }
    if (i % 10 === 0) progress(45 + Math.round((i / names.length) * 40), `Kartlar okunuyor… (${people.length})`);
  }
  if (!people.length) throw new Error('Bu dosyada kişi kartı bulunamadı. Doğru Excel dosyasını seçtiğinizden emin olun.');

  // Yalnızca kullanılan resimleri sakla: kişi fotoğrafları ve bağlantılı tutanak sayfalarının resimleri.
  const attachments = {};
  for (const p of people) for (const k of p.kayitlar) if (k.link) attachments[k.link] = sheetImages[k.link] || [];
  const needed = new Set([...people.map((p) => p.photo).filter(Boolean), ...Object.values(attachments).flat()]);

  progress(88, 'Kaydediliyor…');
  const images = new Map();
  for (const path of needed) {
    const f = zip.file(path);
    if (!f) continue;
    const ext = path.split('.').pop().toLowerCase();
    const type = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[ext] || 'application/octet-stream';
    images.set(path, { type, buf: await f.async('arraybuffer') });
  }

  people.sort((a, b) => (a.no ?? 1e9) - (b.no ?? 1e9) || a.name.localeCompare(b.name, 'tr'));
  return { people, images, meta: { fileName: file.name, importedAt: Date.now(), attachments } };
}

async function saveImport({ people, images, meta }) {
  await DB.replaceImages(images);
  await DB.set('kv', 'people', people);
  await DB.set('kv', 'meta', meta);
  clearUrlCache();
  state.people = people;
  state.images = images;
  state.meta = meta;
  try { await navigator.storage?.persist?.(); } catch { /* isteğe bağlı */ }
}

// ---------- Ekranlar ----------

function renderSetup(isReimport = false) {
  app.innerHTML = `
    <div class="center">
      <img class="logo" src="icons/icon-192.png" alt="">
      <h1>${isReimport ? 'Excel’i yeniden yükle' : 'Personel Rehberi'}</h1>
      <p>Personel Excel dosyanızı seçin. Dosya bu telefonda okunur ve <b>yalnızca bu cihazda</b> saklanır; hiçbir sunucuya gönderilmez.</p>
      <input id="file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>
      <button class="btn" id="pick">Excel dosyası seç</button>
      ${isReimport ? '<button class="btn secondary" id="cancel">Vazgeç</button>' : ''}
      <div id="prog" hidden>
        <div class="progress"><div></div></div>
        <p class="hint" id="progText"></p>
      </div>
      <div class="note"><b>İpucu:</b> Excel dosyasını telefona WhatsApp, e-posta veya Google Drive ile gönderip "Dosyalar / İndirilenler" içinden seçebilirsiniz. Yükledikten sonra telefondaki kopyayı silebilirsiniz.</div>
    </div>`;
  const input = app.querySelector('#file');
  const pick = app.querySelector('#pick');
  pick.onclick = () => input.click();
  app.querySelector('#cancel')?.addEventListener('click', () => history.back());
  input.onchange = async () => {
    const file = input.files[0];
    if (!file) return;
    pick.disabled = true;
    const prog = app.querySelector('#prog');
    prog.hidden = false;
    const setProg = (pct, text) => {
      prog.querySelector('.progress > div').style.width = pct + '%';
      app.querySelector('#progText').textContent = text;
    };
    try {
      const data = await readWorkbook(file, setProg);
      await saveImport(data);
      setProg(100, `${data.people.length} kişi yüklendi.`);
      toast(`${data.people.length} kişi yüklendi`);
      if (!state.pinHash && !isReimport && CAN_PIN) location.hash = '#/pin';
      else location.hash = '#/';
      route();
    } catch (err) {
      console.error(err);
      setProg(0, '');
      toast(err.message || 'Dosya okunamadı');
      app.querySelector('#progText').textContent = 'Hata: ' + (err.message || err);
      pick.disabled = false;
      input.value = '';
    }
  };
}

function renderPinPad({ title, text, onDone, onSkip, skipLabel }) {
  let pin = '';
  app.innerHTML = `
    <div class="center" style="align-items:center;text-align:center">
      <div class="icon-btn" style="color:var(--accent)">${ICON.lock}</div>
      <h1 style="font-size:22px">${esc(title)}</h1>
      <p>${esc(text)}</p>
      <div class="pin-dots">${'<span></span>'.repeat(4)}</div>
      <div class="keypad">
        ${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => `<button data-k="${n}">${n}</button>`).join('')}
        ${onSkip ? `<button class="ghost" data-k="skip">${esc(skipLabel || 'Atla')}</button>` : '<span></span>'}
        <button data-k="0">0</button>
        <button class="ghost" data-k="del" aria-label="Sil">⌫</button>
      </div>
    </div>`;
  const dots = app.querySelector('.pin-dots');
  const paint = () => dots.querySelectorAll('span').forEach((s, i) => s.classList.toggle('on', i < pin.length));
  const shake = () => { dots.classList.remove('shake'); void dots.offsetWidth; dots.classList.add('shake'); };
  app.querySelector('.keypad').onclick = async (e) => {
    const k = e.target.closest('button')?.dataset.k;
    if (!k) return;
    if (k === 'skip') return onSkip();
    if (k === 'del') pin = pin.slice(0, -1);
    else if (pin.length < 4) pin += k;
    paint();
    if (pin.length === 4) {
      const entered = pin;
      setTimeout(async () => {
        const ok = await onDone(entered);
        if (ok === false) { pin = ''; paint(); shake(); }
      }, 120);
    }
  };
}

function renderLock() {
  renderPinPad({
    title: 'Kilitli',
    text: 'Devam etmek için PIN kodunuzu girin.',
    onDone: async (pin) => {
      if ((await pinHash(pin)) !== state.pinHash) return false;
      state.locked = false;
      route();
    },
  });
}

function renderSetPin() {
  const first = () => renderPinPad({
    title: 'PIN belirleyin',
    text: 'Uygulamayı açarken sorulacak 4 haneli bir kod seçin. Telefon başkasının eline geçerse bilgileri korur.',
    onSkip: () => { location.hash = '#/'; },
    skipLabel: state.pinHash ? 'Vazgeç' : 'Atla',
    onDone: (pin) => second(pin),
  });
  const second = (pin1) => renderPinPad({
    title: 'PIN’i tekrar girin',
    text: 'Onaylamak için aynı kodu tekrar girin.',
    onSkip: first,
    skipLabel: 'Geri',
    onDone: async (pin2) => {
      if (pin2 !== pin1) { toast('Kodlar eşleşmedi, tekrar deneyin'); first(); return; }
      state.pinHash = await pinHash(pin1);
      await DB.set('kv', 'pinHash', state.pinHash);
      toast('PIN kaydedildi');
      location.hash = '#/';
    },
  });
  first();
}

function filteredPeople() {
  const q = norm(state.query.trim());
  const qd = digits(state.query);
  return state.people.filter((p) => {
    if (state.filter === 'kayit' && !p.kayitlar.length) return false;
    if (state.filter.startsWith('durum:') && durumKey(p.durum) !== state.filter.slice(6)) return false;
    if (!q) return true;
    if (norm([p.name, p.gorev, p.yakin, p.kan, p.adres].join(' ')).includes(q)) return true;
    return qd.length >= 3 && (digits(p.tel).includes(qd) || digits(p.yakinTel).includes(qd) || digits(p.tc).startsWith(qd));
  });
}

function listHtml() {
  const list = filteredPeople();
  if (!list.length) return '<div class="empty">Sonuç bulunamadı.</div>';
  const groups = new Map();
  for (const p of list) {
    const L = (p.name[0] || '#').toLocaleUpperCase('tr');
    if (!groups.has(L)) groups.set(L, []);
    groups.get(L).push(p);
  }
  return [...groups].map(([L, ps]) => `
    <div class="letter">${esc(L)}</div>
    <div class="group">
      ${ps.map((p) => `
        <button class="row" data-id="${esc(p.id)}">
          ${avatar(p)}
          <div class="grow">
            <div class="name">${esc(p.name)}</div>
            <div class="meta">${[p.durum, p.gorev, p.kayitlar.length ? `${p.kayitlar.length} kayıt` : ''].filter(Boolean).map(esc).join(' · ') || '&nbsp;'}</div>
          </div>
          ${ICON.chev}
        </button>`).join('')}
    </div>`).join('');
}

function renderList() {
  // Excel'de aynı durum farklı yazılmış olabilir (Silahlı / SİLAHLI); bunları tek filtrede topla.
  const variants = new Map();
  for (const p of state.people) {
    const k = durumKey(p.durum);
    if (!k) continue;
    if (!variants.has(k)) variants.set(k, new Map());
    const v = variants.get(k);
    v.set(p.durum, (v.get(p.durum) || 0) + 1);
  }
  const durumlar = [...variants].map(([k, v]) => {
    const label = [...v].sort((a, b) => b[1] - a[1])[0][0];
    const n = [...v.values()].reduce((a, b) => a + b, 0);
    return [k, label, n];
  }).sort((a, b) => b[2] - a[2]);
  const withRec = state.people.filter((p) => p.kayitlar.length).length;
  const chips = [
    ['all', 'Tümü'],
    ...durumlar.map(([k, label, n]) => ['durum:' + k, `${label} (${n})`]),
    ...(withRec ? [['kayit', `Kaydı olanlar (${withRec})`]] : []),
  ];
  app.innerHTML = `
    <div class="bar" style="flex-wrap:wrap;padding-bottom:0">
      <h1>Personel <span class="sub" id="count"></span></h1>
      <button class="icon-btn" id="settings" aria-label="Ayarlar">${ICON.gear}</button>
      <div style="flex-basis:100%"></div>
      <div class="search-wrap" style="padding:4px 0 0;flex:1">
        <label class="search">${ICON.search}<input id="q" type="search" placeholder="İsim, telefon, görev yeri…" autocomplete="off" value="${esc(state.query)}"></label>
      </div>
      <div style="flex-basis:100%"></div>
      <div class="chips" style="padding-left:0;padding-right:0">
        ${chips.map(([k, l]) => `<button class="chip" data-f="${esc(k)}" aria-pressed="${state.filter === k}">${esc(l)}</button>`).join('')}
      </div>
    </div>
    <div class="list" id="list"></div>`;

  const listEl = app.querySelector('#list');
  const update = () => {
    listEl.innerHTML = listHtml();
    app.querySelector('#count').textContent = `(${filteredPeople().length})`;
  };
  update();
  app.querySelector('#q').oninput = (e) => { state.query = e.target.value; update(); };
  app.querySelector('.chips').onclick = (e) => {
    const b = e.target.closest('.chip');
    if (!b) return;
    state.filter = b.dataset.f;
    app.querySelectorAll('.chip').forEach((c) => c.setAttribute('aria-pressed', c === b));
    update();
  };
  listEl.onclick = (e) => {
    const row = e.target.closest('.row');
    if (!row) return;
    state.listScroll = window.scrollY;
    location.hash = '#/k/' + encodeURIComponent(row.dataset.id);
  };
  app.querySelector('#settings').onclick = () => { state.listScroll = window.scrollY; location.hash = '#/ayarlar'; };
  requestAnimationFrame(() => window.scrollTo(0, state.listScroll));
}

function field(k, v, extra = '') {
  return `<div class="field"><span class="k">${esc(k)}</span>${v ? `<span class="v">${extra || esc(v)}</span>` : '<span class="v blank">—</span>'}</div>`;
}
function telLink(v) {
  const d = digits(v);
  return d ? `<a class="linkish" href="tel:${d}">${esc(v)}</a>` : '';
}

async function renderDetail(id) {
  const p = state.people.find((x) => x.id === id);
  if (!p) { location.hash = '#/'; return; }
  const note = state.notes.has(p.id) ? state.notes.get(p.id) : p.notlar;
  const photo = imgUrl(p.photo);
  const mapUrl = p.adres ? 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(p.adres) : '';
  const total = state.people.length;

  app.innerHTML = `
    <div class="bar">
      <button class="icon-btn" id="back" aria-label="Geri">${ICON.back}</button>
      <h1 style="font-size:17px">Kişi Kartı</h1>
      <span class="sub">${p.no ? `${String(p.no).padStart(3, '0')} / ${total}` : ''}</span>
    </div>
    <div class="detail">
      <div class="hero">
        ${photo ? `<img class="photo" src="${photo}" alt="${esc(p.name)} fotoğrafı">` : `<div class="photo">${esc(initials(p.name))}</div>`}
        <h2>${esc(p.name)}</h2>
        <div>${p.durum ? `<span class="badge ${durumClass(p.durum)}">${esc(p.durum)}</span>` : ''}</div>
      </div>

      <div class="actions">
        <a class="action" href="tel:${digits(p.tel)}" aria-disabled="${!digits(p.tel)}">${ICON.phone}Ara</a>
        <a class="action" href="tel:${digits(p.yakinTel)}" aria-disabled="${!digits(p.yakinTel)}">${ICON.people}Yakınını ara</a>
        <a class="action" href="${mapUrl}" target="_blank" rel="noopener" aria-disabled="${!mapUrl}">${ICON.map}Adres</a>
      </div>

      <div class="section">
        <h3>Kimlik</h3>
        <div class="fields">
          <div class="field"><span class="k">TC Kimlik No</span>${p.tc ? `<button class="linkish" id="tc" aria-label="TC numarasını göster">${esc(maskTc(p.tc))}</button>` : '<span class="v blank">—</span>'}</div>
          ${field('Doğum Tarihi', p.dogum)}
          ${field('Kan Grubu', p.kan)}
        </div>
      </div>

      <div class="section">
        <h3>Görev</h3>
        <div class="fields">
          ${field('Durum', p.durum)}
          ${field('Görev Yeri', p.gorev)}
        </div>
      </div>

      <div class="section">
        <h3>İletişim</h3>
        <div class="fields">
          ${field('Kendi Telefonu', p.tel, telLink(p.tel))}
          ${field('Yakını', p.yakin)}
          ${field('Yakınının Telefonu', p.yakinTel, telLink(p.yakinTel))}
          <div class="field block"><span class="k">Adres</span><div class="v ${p.adres ? '' : 'blank'}">${esc(p.adres || '—')}</div></div>
        </div>
      </div>

      <div class="section">
        <h3>Tutanak / İzin Kayıtları</h3>
        <div class="fields">
          ${p.kayitlar.length ? p.kayitlar.map((k, i) => `
            <div class="rec">
              <div class="top"><span>${esc(k.tarih || 'Tarih yok')}</span></div>
              <div class="what">${esc(k.tur || '—')}</div>
              ${k.aciklama ? `<div>${esc(k.aciklama)}</div>` : ''}
              ${k.link && state.meta?.attachments?.[k.link]?.length ? `<button class="btn secondary" style="height:40px" data-rec="${i}">📷 Tutanağı görüntüle</button>` : ''}
            </div>`).join('') : '<div class="field"><span class="k">Kayıt yok</span></div>'}
        </div>
      </div>

      <div class="section">
        <h3>Notlar</h3>
        <textarea id="note" placeholder="Bu kişiyle ilgili not ekleyin…">${esc(note)}</textarea>
        <p class="hint" id="noteHint">Notlar otomatik olarak bu telefona kaydedilir.</p>
      </div>
    </div>`;

  window.scrollTo(0, 0);
  app.querySelector('#back').onclick = () => (history.length > 1 ? history.back() : (location.hash = '#/'));
  const tc = app.querySelector('#tc');
  if (tc) {
    let shown = false;
    tc.onclick = () => { shown = !shown; tc.textContent = shown ? p.tc : maskTc(p.tc); };
  }
  app.querySelectorAll('[data-rec]').forEach((b) => {
    b.onclick = () => { location.hash = `#/k/${encodeURIComponent(p.id)}/t/${b.dataset.rec}`; };
  });
  let t;
  app.querySelector('#note').oninput = (e) => {
    const v = e.target.value;
    clearTimeout(t);
    t = setTimeout(async () => {
      state.notes.set(p.id, v);
      await DB.set('kv', 'note:' + p.id, v);
      app.querySelector('#noteHint').textContent = 'Kaydedildi ✓';
    }, 500);
  };
}

function renderViewer(id, idx) {
  const p = state.people.find((x) => x.id === id);
  const rec = p?.kayitlar[idx];
  const keys = rec?.link ? state.meta?.attachments?.[rec.link] || [] : [];
  if (!keys.length) { location.hash = '#/k/' + encodeURIComponent(id); return; }
  app.innerHTML = `
    <div class="viewer">
      <div class="bar">
        <button class="icon-btn" id="close" aria-label="Kapat">${ICON.close}</button>
        <h1 style="font-size:16px">${esc(rec.tur || 'Tutanak')}</h1>
        <span class="sub" style="color:#bbb">${esc(rec.tarih)}</span>
      </div>
      <div class="imgs">${keys.map((k) => `<img src="${imgUrl(k)}" alt="Tutanak">`).join('')}</div>
    </div>`;
  app.querySelector('#close').onclick = () => history.back();
}

function renderSettings() {
  const m = state.meta;
  app.innerHTML = `
    <div class="bar">
      <button class="icon-btn" id="back" aria-label="Geri">${ICON.back}</button>
      <h1 style="font-size:17px">Ayarlar</h1>
    </div>
    <div class="settings">
      <div class="note">
        <b>${state.people.length} kişi</b> yüklü<br>
        ${m ? `Dosya: ${esc(m.fileName)}<br>Yüklenme: ${esc(new Date(m.importedAt).toLocaleString('tr-TR'))}` : ''}
      </div>
      <button class="btn" id="reimport">Güncel Excel’i yükle</button>
      <p class="hint">Excel’de değişiklik yaptığınızda dosyayı tekrar yükleyin. Uygulamada yazdığınız notlar korunur.</p>
      ${CAN_PIN
        ? `<button class="btn secondary" id="pin">${state.pinHash ? 'PIN’i değiştir' : 'PIN belirle'}</button>`
        : '<p class="hint">PIN kilidi yalnızca uygulama https:// adresinden (ör. GitHub Pages) açıldığında kullanılabilir.</p>'}
      ${state.pinHash ? '<button class="btn secondary" id="nopin">PIN’i kaldır</button>' : ''}
      ${state.pinHash ? '<button class="btn secondary" id="lock">Şimdi kilitle</button>' : ''}
      <div style="height:10px"></div>
      <button class="btn danger" id="wipe">Tüm verileri bu telefondan sil</button>
      <p class="hint">Veriler yalnızca bu cihazda tutulur. Uygulamayı silmek veya tarayıcı verilerini temizlemek de verileri siler.</p>
    </div>`;
  app.querySelector('#back').onclick = () => (location.hash = '#/');
  app.querySelector('#reimport').onclick = () => (location.hash = '#/yukle');
  app.querySelector('#pin')?.addEventListener('click', () => (location.hash = '#/pin'));
  app.querySelector('#nopin')?.addEventListener('click', async () => {
    if (!confirm('PIN kilidi kaldırılsın mı?')) return;
    state.pinHash = null;
    await DB.del('kv', 'pinHash');
    toast('PIN kaldırıldı');
    renderSettings();
  });
  app.querySelector('#lock')?.addEventListener('click', () => { state.locked = true; location.hash = '#/'; route(); });
  app.querySelector('#wipe').onclick = async () => {
    if (!confirm('Bu telefondaki tüm personel bilgileri, fotoğraflar ve notlar silinecek. Emin misiniz?')) return;
    await DB.wipe();
    clearUrlCache();
    Object.assign(state, { people: [], meta: null, images: new Map(), notes: new Map(), pinHash: null, locked: false });
    toast('Tüm veriler silindi');
    location.hash = '#/';
    route();
  };
}

// ---------- Yönlendirme ----------

function route() {
  if (state.locked) return renderLock();
  const h = decodeURIComponent(location.hash.slice(1) || '/');
  if (!state.people.length) return renderSetup();
  if (h === '/pin' && CAN_PIN) return renderSetPin();
  if (h === '/yukle') return renderSetup(true);
  if (h === '/ayarlar') return renderSettings();
  let m = h.match(/^\/k\/(.+)\/t\/(\d+)$/);
  if (m) return renderViewer(m[1], Number(m[2]));
  m = h.match(/^\/k\/(.+)$/);
  if (m) return renderDetail(m[1]);
  renderList();
}

window.addEventListener('hashchange', route);

document.addEventListener('visibilitychange', () => {
  if (document.hidden) state.hiddenAt = Date.now();
  else if (state.pinHash && !state.locked && Date.now() - state.hiddenAt > AUTO_LOCK_MS) {
    state.locked = true;
    route();
  }
});

async function start() {
  try {
    const [people, meta, pin, kv, imgs] = await Promise.all([
      DB.get('kv', 'people'),
      DB.get('kv', 'meta'),
      DB.get('kv', 'pinHash'),
      DB.entries('kv'),
      DB.entries('images'),
    ]);
    state.people = people || [];
    state.meta = meta || null;
    state.pinHash = pin || null;
    state.locked = !!pin;
    state.images = new Map(imgs);
    for (const [k, v] of kv) if (typeof k === 'string' && k.startsWith('note:')) state.notes.set(k.slice(5), v);
  } catch (err) {
    console.error(err);
    toast('Depolama açılamadı. Gizli sekmede değilseniz sayfayı yenileyin.');
  }
  route();
}

if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

start();
