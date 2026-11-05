'use strict';

/**
 * Fuentes para los PDF generados (Word, PowerPoint, Excel):
 *  - Equivalentes libres de las fuentes de Office con las mismas medidas
 *    (Carlito = Calibri, Caladea = Cambria, Arimo = Arial, Tinos = Times New Roman,
 *    Cousine = Courier New, Gelasio = Georgia…).
 *  - Cualquier fuente de Google Fonts por su nombre (Roboto, Montserrat, Open Sans…),
 *    con sus grosores (Light, SemiBold, Black…).
 *  - Recorte propio de las fuentes al guardar: solo se guardan los glifos usados, pero
 *    con sus números originales (el recorte de pdf-lib renumera glifos y pierde letras).
 */
const FontBank = (() => {
  const GF = 'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl';
  const EXPO = 'https://cdn.jsdelivr.net/npm/@expo-google-fonts';
  const ghFamily = (dir, name) => ({ files: ['Regular', 'Bold', 'Italic', 'BoldItalic'].map((v) => `${GF}/${dir}/${name}-${v}.ttf`) });
  const expo = (slug) => ({ expo: slug });

  /** Fuentes de Office → equivalente libre (mismas medidas cuando existe). */
  const OFFICE = [
    [/^(calibri|aptos|carlito)/, ghFamily('carlito', 'Carlito')],
    [/^(cambria|caladea)/, ghFamily('caladea', 'Caladea')],
    [/^(arial|helvetica|liberation sans|arimo|nimbus sans)/, expo('arimo')],
    [/^(times|liberation serif|tinos|nimbus roman)/, expo('tinos')],
    [/^(courier|liberation mono|cousine|nimbus mono)/, ghFamily('cousine', 'Cousine')],
    [/^georgia/, expo('gelasio')],
    [/^(segoe|selawik|verdana|tahoma|frutiger|myriad)/, expo('open-sans')],
    [/^(consolas|lucida console|menlo|monaco|sf mono)/, expo('inconsolata')],
    [/^(trebuchet|gill sans)/, expo('fira-sans')],
    [/^(garamond|adobe garamond)/, expo('eb-garamond')],
    [/^(book antiqua|palatino)/, expo('gelasio')],
    [/^comic sans/, expo('comic-neue')],
    [/^impact/, expo('anton')],
    [/^franklin gothic/, expo('libre-franklin')],
    [/^century gothic/, expo('didact-gothic')],
    [/^(candara|corbel|constantia)/, ghFamily('carlito', 'Carlito')],
  ];
  const WEIGHTS = [
    [/\b(thin|hairline)\b/, 100], [/\b(extra ?light|ultra ?light)\b/, 200], [/\blight\b/, 300], [/\b(medium)\b/, 500],
    [/\b(semi ?bold|demi ?bold)\b/, 600], [/\b(extra ?bold|ultra ?bold)\b/, 800], [/\b(black|heavy)\b/, 900],
  ];

  const bytesCache = {};
  const listCache = {};
  let fontkitReady = null;

  /** Almacén permanente en el navegador (IndexedDB): cada fuente se descarga una sola vez. */
  const store = (() => {
    let dbp = null;
    const open = () => (dbp = dbp || new Promise((res, rej) => {
      const r = indexedDB.open('converter-fuentes', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('f');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }));
    return {
      async get(k) {
        try {
          const db = await open();
          return await new Promise((res) => { const q = db.transaction('f').objectStore('f').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(undefined); });
        } catch (_) { return undefined; }
      },
      async set(k, v) {
        try { const db = await open(); db.transaction('f', 'readwrite').objectStore('f').put(v, k); } catch (_) { /* sin almacén: no pasa nada */ }
      },
    };
  })();

  /** Descompone "Open Sans SemiBold" en familia base y grosor. */
  function parseName(name) {
    let n = String(name || '').toLowerCase().replace(/[_-]+/g, ' ').trim();
    let weight = 400;
    for (const [re, w] of WEIGHTS) if (re.test(n)) { weight = w; n = n.replace(re, ' '); }
    n = n.replace(/\b(regular|normal|book|roman|condensed|narrow|display|text|ms|mt|std|pro|nova)\b/g, ' ').replace(/\s+/g, ' ').trim();
    return { base: n, weight };
  }

  /** Clave de la familia libre que usaremos para un nombre de fuente. */
  function familyFor(name) {
    const { base } = parseName(name);
    for (const [re] of OFFICE) if (re.test(base)) return re.source;
    return base ? 'google:' + base.replace(/ /g, '-') : OFFICE[0][0].source;
  }

  function sourceFor(name) {
    const { base, weight } = parseName(name);
    for (const [re, src] of OFFICE) if (re.test(base)) return { ...src, weight, key: re.source };
    if (base) return { expo: base.replace(/ /g, '-'), weight, key: 'google:' + base.replace(/ /g, '-'), guessed: true };
    return { ...OFFICE[0][1], weight, key: OFFICE[0][0].source };
  }

  async function load(url) {
    if (!bytesCache[url]) {
      bytesCache[url] = (async () => {
        const saved = await store.get('ttf:' + url);
        if (saved) return saved;
        const r = await fetch(url);
        if (!r.ok) throw new Error('fuente no disponible');
        const buf = await r.arrayBuffer();
        store.set('ttf:' + url, buf);
        return buf;
      })().catch((e) => { delete bytesCache[url]; throw e; });
    }
    return bytesCache[url];
  }

  /** Lista de archivos .ttf de un paquete @expo-google-fonts (con su grosor y estilo). */
  async function expoList(slug) {
    if (!listCache[slug]) {
      listCache[slug] = (async () => {
        const saved = await store.get('lista:' + slug);
        if (saved) return saved;
        const v = await fetch(`https://data.jsdelivr.com/v1/package/resolve/npm/@expo-google-fonts/${slug}@latest`).then((r) => (r.ok ? r.json() : null));
        if (!v || !v.version) return null;
        const d = await fetch(`https://data.jsdelivr.com/v1/package/npm/@expo-google-fonts/${slug}@${v.version}/flat`).then((r) => (r.ok ? r.json() : null));
        const files = (d?.files || []).map((f) => f.name).filter((n) => /\.ttf$/.test(n));
        const list = files.map((f) => {
          const m = /\/(\d{3})[A-Za-z]+(_Italic)?\//.exec(f);
          return m ? { url: `${EXPO}/${slug}@${v.version}${f}`, weight: +m[1], italic: !!m[2] } : null;
        }).filter(Boolean);
        if (list.length) store.set('lista:' + slug, list);
        return list;
      })().catch(() => null);
    }
    return listCache[slug];
  }

  /** URL del archivo de fuente que mejor encaja (o null). */
  async function fileFor(src, bold, italic) {
    const want = bold ? Math.max(700, src.weight + 300) : src.weight;
    if (src.files) return src.files[(bold ? 1 : 0) + (italic ? 2 : 0)];
    const list = await expoList(src.expo);
    if (!list || !list.length) return null;
    const pool = list.filter((f) => f.italic === italic).length ? list.filter((f) => f.italic === italic) : list;
    return pool.slice().sort((a, b) => Math.abs(a.weight - want) - Math.abs(b.weight - want) || b.weight - a.weight)[0].url;
  }

  /* ── recorte de fuentes TrueType conservando los números de glifo ── */
  function stripTTF(input, used) {
    const src = new Uint8Array(input);
    const dv = new DataView(src.buffer, src.byteOffset, src.byteLength);
    const tag = (o) => String.fromCharCode(src[o], src[o + 1], src[o + 2], src[o + 3]);
    const n = dv.getUint16(4);
    const T = {};
    for (let i = 0; i < n; i++) {
      const o = 12 + i * 16;
      T[tag(o)] = { off: dv.getUint32(o + 8), len: dv.getUint32(o + 12) };
    }
    if (!T.glyf || !T.loca || !T.head || !T.maxp) return null;
    const longLoca = dv.getInt16(T.head.off + 50) === 1;
    const numGlyphs = dv.getUint16(T.maxp.off + 4);
    const locaAt = (g) => (longLoca ? dv.getUint32(T.loca.off + g * 4) : dv.getUint16(T.loca.off + g * 2) * 2);
    // glifos a conservar (+ componentes de los compuestos)
    const keep = new Set([0]);
    const queue = [0, ...used];
    while (queue.length) {
      const g = queue.pop();
      if (g >= numGlyphs) continue;
      keep.add(g);
      const a = locaAt(g), b = locaAt(g + 1);
      if (b <= a || dv.getInt16(T.glyf.off + a) >= 0) continue;
      let p = T.glyf.off + a + 10, flags;
      do {
        flags = dv.getUint16(p);
        const gi = dv.getUint16(p + 2);
        if (!keep.has(gi)) queue.push(gi);
        p += 4 + (flags & 1 ? 4 : 2) + (flags & 8 ? 2 : flags & 0x40 ? 4 : flags & 0x80 ? 8 : 0);
      } while (flags & 0x20);
    }
    // nuevo glyf + loca (formato largo)
    const loca = new DataView(new ArrayBuffer((numGlyphs + 1) * 4));
    let size = 0;
    for (const g of keep) { const l = locaAt(g + 1) - locaAt(g); size += (l + 3) & ~3; }
    const glyf = new Uint8Array(size);
    let pos = 0;
    for (let g = 0; g < numGlyphs; g++) {
      loca.setUint32(g * 4, pos);
      if (!keep.has(g)) continue;
      const a = locaAt(g), l = locaAt(g + 1) - a;
      if (l > 0) { glyf.set(src.subarray(T.glyf.off + a, T.glyf.off + a + l), pos); pos += (l + 3) & ~3; }
    }
    loca.setUint32(numGlyphs * 4, pos);
    // tablas finales
    const KEEP = ['OS/2', 'cmap', 'cvt ', 'fpgm', 'gasp', 'head', 'hhea', 'hmtx', 'maxp', 'name', 'prep'];
    const out = {};
    for (const t of KEEP) if (T[t]) out[t] = src.slice(T[t].off, T[t].off + T[t].len);
    out.glyf = glyf;
    out.loca = new Uint8Array(loca.buffer);
    if (T.post) { const post = src.slice(T.post.off, T.post.off + 32); new DataView(post.buffer).setUint32(0, 0x00030000); out.post = post; }
    new DataView(out.head.buffer).setInt16(50, 1); // indexToLocFormat = largo
    new DataView(out.head.buffer).setUint32(8, 0); // checkSumAdjustment (se recalcula)
    const tags = Object.keys(out).sort();
    const sum = (u8) => { let s = 0; const p = new Uint8Array((u8.length + 3) & ~3); p.set(u8); const v = new DataView(p.buffer); for (let i = 0; i < p.length; i += 4) s = (s + v.getUint32(i)) >>> 0; return s; };
    let total = 12 + tags.length * 16;
    for (const t of tags) total += (out[t].length + 3) & ~3;
    const font = new Uint8Array(total);
    const fv = new DataView(font.buffer);
    fv.setUint32(0, dv.getUint32(0));
    fv.setUint16(4, tags.length);
    const es = Math.floor(Math.log2(tags.length));
    fv.setUint16(6, 16 * 2 ** es); fv.setUint16(8, es); fv.setUint16(10, tags.length * 16 - 16 * 2 ** es);
    let off = 12 + tags.length * 16;
    let headOff = 0;
    tags.forEach((t, i) => {
      const d = out[t];
      for (let k = 0; k < 4; k++) font[12 + i * 16 + k] = t.charCodeAt(k);
      fv.setUint32(12 + i * 16 + 4, sum(d));
      fv.setUint32(12 + i * 16 + 8, off);
      fv.setUint32(12 + i * 16 + 12, d.length);
      font.set(d, off);
      if (t === 'head') headOff = off;
      off += (d.length + 3) & ~3;
    });
    fv.setUint32(headOff + 8, (0xb1b0afba - sum(font)) >>> 0);
    return font;
  }

  /** Banco de fuentes para un documento pdf-lib concreto. */
  function forDoc(doc) {
    const embedded = {};
    const tracked = [];
    let fallback = null;
    let ok = true;
    const FEATURES = { liga: false, clig: false, dlig: false, rlig: false, calt: false };

    // Al guardar: recorta cada fuente a los glifos usados
    const save = doc.save.bind(doc);
    doc.save = async (opts) => {
      for (const t of tracked) {
        try {
          const emb = t.font.embedder;
          // también las fuentes cargadas solo para medir texto: se quedan sin glifos
          const stripped = stripTTF(t.bytes, t.used);
          if (!stripped) continue;
          emb.fontData = stripped;
          const glyphs = [...t.used].sort((a, b) => a - b).map((id) => emb.font.getGlyph(id));
          emb.glyphCache = { access: () => glyphs, invalidate() {} };
        } catch (e) {
          console.warn('No se pudo recortar una fuente; se incrusta entera', e);
        }
      }
      return save(opts);
    };

    async function embedFile(url) {
      const bytes = await load(url);
      const f = await doc.embedFont(bytes, { subset: false, features: FEATURES });
      // comprobación: algunas fuentes no las entiende el lector de fuentes de pdf-lib
      f.widthOfTextAtSize('Aá ñ€ 0', 10);
      f.encodeText('Aá ñ€ 0');
      const t = { font: f, bytes, used: new Set() };
      const emb = f.embedder;
      const enc = emb.encodeText.bind(emb);
      emb.encodeText = (text) => {
        for (const g of emb.font.layout(text, emb.fontFeatures).glyphs) t.used.add(g.id);
        return enc(text);
      };
      tracked.push(t);
      f.__uni = true;
      return f;
    }

    const api = {
      /** Precarga fontkit y las fuentes indicadas (nombres de Office o de Google Fonts). */
      async prepare(names = ['Calibri']) {
        try {
          if (!fontkitReady) fontkitReady = U.loadScript('vendor/fontkit.umd.min.js');
          await fontkitReady;
          doc.registerFontkit(window.fontkit);
          const srcs = [...new Set(names)].map(sourceFor);
          await Promise.all(srcs.map(async (s) => { const u = await fileFor(s, false, false); if (u) await load(u).catch(() => {}); }));
        } catch (_) {
          ok = false;
        }
        fallback = await stdFonts(doc);
      },
      /** Fuente para (nombre, negrita, cursiva). Si no se encuentra, usa la equivalente más parecida. */
      async get(name, bold, italic) {
        if (!fallback) fallback = await stdFonts(doc);
        const std = fallback[(bold ? 'b' : '') + (italic ? 'i' : '')];
        if (!ok) return std;
        const src = sourceFor(name);
        const key = `${src.key}|${src.weight}|${bold ? 1 : 0}${italic ? 1 : 0}`;
        if (!embedded[key]) {
          embedded[key] = (async () => {
            try {
              const url = await fileFor(src, bold, italic);
              if (!url) throw new Error('sin archivo');
              return await embedFile(url);
            } catch (_) {
              // familia no encontrada: Carlito (o Tinos si parece una fuente con serifa)
              if (src.key !== OFFICE[0][0].source) return api.get(/serif|roman|times|book/i.test(name || '') && !/sans/i.test(name || '') ? 'Times New Roman' : 'Calibri', bold, italic);
              return std;
            }
          })();
        }
        return embedded[key];
      },
      familyFor,
    };
    return api;
  }

  return { forDoc, familyFor, stripTTF };
})();
