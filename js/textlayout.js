'use strict';

/* ════════════════════════════════════════════════════════════════
   Motor de maquetación de texto para pdf-lib (Word, PowerPoint, Excel)
   ════════════════════════════════════════════════════════════════ */

async function stdFonts(doc) {
  const S = PDFLib.StandardFonts;
  return {
    '': await doc.embedFont(S.Helvetica),
    b: await doc.embedFont(S.HelveticaBold),
    i: await doc.embedFont(S.HelveticaOblique),
    bi: await doc.embedFont(S.HelveticaBoldOblique),
  };
}

/** Limpia el texto según lo que admita la fuente (Unicode o solo WinAnsi). */
const cleanText = (t, font) =>
  font && font.__uni ? String(t).replace(/[\u0000-\u0008\u000B-\u001F\u007F​-‍﻿]/g, '') : winAnsi(String(t));

/** Métricas verticales de la fuente (en em), como las usa Word para el interlineado. */
function fontMetrics(font) {
  if (font.__m) return font.__m;
  let asc = 0.77, desc = 0.23;
  try {
    const fk = font.embedder && font.embedder.font;
    if (fk) {
      const upm = fk.unitsPerEm;
      const os2 = fk['OS/2'];
      if (os2 && os2.winAscent) { asc = os2.winAscent / upm; desc = os2.winDescent / upm; }
      else { asc = fk.ascent / upm; desc = -fk.descent / upm; }
    }
  } catch (_) { /* métricas por defecto */ }
  return (font.__m = { asc, desc });
}

/**
 * Parte runs en líneas.
 * run: {text, font, size, color, underline, strike, highlight, rise} | {box:{w,h,draw(page,x,baseline)}}
 * opts: {maxW, firstOffset, tabs:[{pos,align,leader}], tabOrigin, defFont, defSize,
 *        lineBox(i, lines) → {offset, width}: hueco disponible para la línea i (texto que rodea imágenes)}
 * Devuelve [{items, width, x0, limit, asc, desc, last}]
 */
function layoutRuns(runs, opts) {
  const { maxW, firstOffset = 0, tabs = [], tabOrigin = 0, defFont, defSize = 11 } = opts;
  const lines = [];
  const newLine = () => {
    const first = lines.length === 0 ? firstOffset : 0;
    let box = opts.lineBox ? opts.lineBox(lines.length, lines) : null;
    // hueco demasiado estrecho junto a una imagen: línea vacía y seguimos más abajo
    for (let guard = 0; box && box.width - first < 30 && guard < 200; guard++) {
      const m = fontMetrics(defFont);
      lines.push({ items: [], width: 0, x0: 0, limit: 0, asc: m.asc * defSize, desc: m.desc * defSize, last: false, spacer: true });
      box = opts.lineBox(lines.length, lines);
    }
    if (box) return { items: [], width: 0, x0: first + box.offset, limit: box.offset + box.width };
    return { items: [], width: 0, x0: first, limit: maxW };
  };
  let cur = newLine();
  const avail = () => cur.limit - cur.x0;
  const finish = (last = false) => {
    while (cur.items.length && cur.items[cur.items.length - 1].kind === 'space') cur.width -= cur.items.pop().w;
    resolveTabs(cur, tabs, tabOrigin);
    let asc = 0, desc = 0;
    for (const it of cur.items) {
      if (it.kind === 'box') asc = Math.max(asc, it.h);
      else if (it.font) {
        const m = fontMetrics(it.font);
        asc = Math.max(asc, m.asc * it.size + Math.max(0, it.rise || 0));
        desc = Math.max(desc, m.desc * it.size);
      }
    }
    if (!asc) {
      const m = fontMetrics(cur.font || defFont);
      const s = cur.size || defSize;
      asc = m.asc * s; desc = m.desc * s;
    }
    Object.assign(cur, { asc, desc, last });
    lines.push(cur);
    cur = newLine();
  };
  const add = (it) => { cur.items.push(it); cur.width += it.w; };

  for (const r of runs) {
    if (r.box) {
      if (cur.width + r.box.w > avail() && cur.items.length) finish();
      add({ kind: 'box', w: r.box.w, h: r.box.h, draw: r.box.draw });
      continue;
    }
    const font = r.font || defFont;
    const size = r.size || defSize;
    cur.font = cur.font || font; cur.size = cur.size || size;
    const style = { font, size, color: r.color, underline: r.underline, strike: r.strike, highlight: r.highlight, rise: r.rise, fn: r.fn };
    String(r.text).split('\n').forEach((seg, si) => {
      if (si > 0) finish();
      cur.font = cur.font || font; cur.size = cur.size || size;
      for (const tok of seg.split(/( +|\t)/)) {
        if (!tok) continue;
        if (tok === '\t') {
          const x = tabOrigin + cur.x0 + cur.width;
          const stop = tabs.find((t) => t.pos > x + 0.5);
          const pos = stop ? stop.pos : (Math.floor(x / 36) + 1) * 36;
          add({ kind: 'tab', w: Math.max(2, pos - x), stop: stop || { pos, align: 'left' }, ...style });
          continue;
        }
        if (/^ +$/.test(tok)) {
          if (cur.items.length) add({ kind: 'space', text: ' ', w: font.widthOfTextAtSize(' ', size) * tok.length, ...style });
          continue;
        }
        const text = cleanText(tok, font);
        if (!text) continue;
        const w = font.widthOfTextAtSize(text, size);
        if (cur.width + w > avail() && cur.items.length) finish();
        if (w <= avail()) { add({ kind: 'text', text, w, ...style }); continue; }
        let chunk = '';
        for (const ch of text) {
          if (chunk && cur.width + font.widthOfTextAtSize(chunk + ch, size) > avail()) {
            add({ kind: 'text', text: chunk, w: font.widthOfTextAtSize(chunk, size), ...style });
            finish();
            chunk = ch;
          } else chunk += ch;
        }
        if (chunk) add({ kind: 'text', text: chunk, w: font.widthOfTextAtSize(chunk, size), ...style });
      }
    });
  }
  finish(true);
  return lines;
}

/** Ajusta los tabuladores alineados a la derecha / centrados una vez montada la línea. */
function resolveTabs(line, tabs, origin) {
  if (!line.items.some((i) => i.kind === 'tab')) return;
  let x = origin + line.x0;
  const items = line.items;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.kind !== 'tab') { x += it.w; continue; }
    const stop = tabs.find((t) => t.pos > x + 0.5) || { pos: (Math.floor(x / 36) + 1) * 36, align: 'left' };
    let seg = 0;
    for (let j = i + 1; j < items.length && items[j].kind !== 'tab'; j++) seg += items[j].w;
    let w = stop.pos - x;
    if (stop.align === 'right') w -= seg;
    else if (stop.align === 'center') w -= seg / 2;
    else if (stop.align === 'decimal') w -= seg;
    it.w = Math.max(stop.align === 'left' ? 2 : 0, w);
    it.leader = stop.leader;
    x += it.w;
  }
  line.width = items.reduce((a, b) => a + b.w, 0);
}

/** Dibuja una línea con la línea base en `base`. */
function drawTextLine(page, line, x, base, maxW, align = 'left', defColor = PDFLib.rgb(0, 0, 0)) {
  const free = (line.limit ?? maxW) - line.x0 - line.width;
  let cx = x + line.x0 + (align === 'center' ? free / 2 : align === 'right' ? free : 0);
  let extra = 0;
  if (align === 'both' && !line.last && free > 0) {
    const spaces = line.items.filter((i) => i.kind === 'space').length;
    if (spaces) extra = free / spaces;
  }
  for (const it of line.items) {
    const w = it.w + (it.kind === 'space' ? extra : 0);
    if (it.kind === 'box') {
      it.draw(page, cx, base);
    } else if (it.kind === 'text') {
      const color = it.color || defColor;
      if (it.highlight) {
        const m = fontMetrics(it.font);
        page.drawRectangle({ x: cx, y: base - m.desc * it.size, width: w, height: (m.asc + m.desc) * it.size, color: it.highlight });
      }
      page.drawText(it.text, { x: cx, y: base + (it.rise || 0), size: it.size, font: it.font, color });
      const th = Math.max(0.5, it.size / 16);
      if (it.underline) page.drawLine({ start: { x: cx, y: base - it.size * 0.12 }, end: { x: cx + w, y: base - it.size * 0.12 }, thickness: th, color });
      if (it.strike) page.drawLine({ start: { x: cx, y: base + it.size * 0.28 }, end: { x: cx + w, y: base + it.size * 0.28 }, thickness: th, color });
    } else if (it.kind === 'tab' && it.leader && it.leader !== 'none' && w > 6) {
      const dot = it.leader === 'hyphen' ? '-' : it.leader === 'underscore' ? '_' : '.';
      const dw = it.font.widthOfTextAtSize(dot, it.size);
      const n = Math.floor((w - 4) / (dw + 1));
      if (n > 0) page.drawText(Array(n).fill(dot).join(''), { x: cx + w - n * dw - 2, y: base, size: it.size, font: it.font, color: it.color || defColor });
    } else if (it.kind === 'space' && it.underline) {
      page.drawLine({ start: { x: cx, y: base - it.size * 0.12 }, end: { x: cx + w, y: base - it.size * 0.12 }, thickness: Math.max(0.5, it.size / 16), color: it.color || defColor });
    }
    cx += w;
  }
}

/** Incrusta una imagen de cualquier formato que el navegador sepa leer. */
async function embedAnyImage(doc, bytes, mime) {
  try {
    if (/png/.test(mime)) return { img: await doc.embedPng(bytes) };
    if (/jpe?g/.test(mime)) return { img: await doc.embedJpg(bytes) };
  } catch (_) { /* probamos con el navegador */ }
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  try {
    const im = new Image();
    im.src = url;
    await im.decode();
    const c = document.createElement('canvas');
    c.width = im.naturalWidth || 300;
    c.height = im.naturalHeight || 150;
    c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
    const png = await U.toBlob(c, 'image/png');
    return { img: await doc.embedPng(await png.arrayBuffer()) };
  } catch (_) {
    return null; // EMF/WMF y similares: el navegador no los entiende
  } finally {
    URL.revokeObjectURL(url);
  }
}

const mimeOf = (name) => {
  const e = String(name).split('.').pop().toLowerCase();
  return { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' }[e] || 'application/octet-stream';
};

/* Utilidades XML (los documentos de Office usan muchos espacios de nombres: buscamos por nombre local) */
const X = {
  parse: (s) => new DOMParser().parseFromString(s, 'application/xml'),
  kids: (el, name) => (el ? [...el.children].filter((c) => !name || c.localName === name) : []),
  kid: (el, name) => (el ? [...el.children].find((c) => c.localName === name) || null : null),
  path: (el, ...names) => names.reduce((e, n) => (e ? [...e.children].find((c) => c.localName === n) || null : null), el),
  desc: (el, name) => (el ? el.getElementsByTagNameNS('*', name)[0] || null : null),
  all: (el, name) => (el ? [...el.getElementsByTagNameNS('*', name)] : []),
  /** atributo por nombre local (w:val, r:id…) */
  attr: (el, name) => {
    if (!el) return null;
    for (const a of el.attributes) if (a.localName === name) return a.value;
    return null;
  },
  num: (el, name, d = 0) => {
    const v = X.attr(el, name);
    return v == null || v === '' || isNaN(+v) ? d : +v;
  },
};

/** Lee las relaciones de una parte de un paquete OOXML. */
async function ooxmlRels(zip, partPath) {
  const i = partPath.lastIndexOf('/');
  const dir = partPath.slice(0, i);
  const f = zip.file(`${dir}/_rels/${partPath.slice(i + 1)}.rels`);
  const map = {};
  if (!f) return map;
  const d = X.parse(await f.async('string'));
  const resolve = (t) => {
    if (t.startsWith('/')) return t.slice(1);
    const parts = dir ? dir.split('/') : [];
    for (const s of t.split('/')) {
      if (s === '..') parts.pop();
      else if (s !== '.') parts.push(s);
    }
    return parts.join('/');
  };
  for (const r of X.all(d, 'Relationship')) {
    const ext = r.getAttribute('TargetMode') === 'External';
    map[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', target: ext ? null : resolve(r.getAttribute('Target')), url: ext ? r.getAttribute('Target') : null };
  }
  return map;
}

/** Colores del tema de Office (theme1.xml) y fuentes mayor/menor. */
function parseTheme(themeDoc) {
  const scheme = {};
  const cs = X.desc(themeDoc, 'clrScheme');
  if (cs) for (const c of cs.children) {
    const v = X.kid(c, 'srgbClr')?.getAttribute('val') || X.kid(c, 'sysClr')?.getAttribute('lastClr');
    if (v) scheme[c.localName] = v;
  }
  Object.assign(scheme, {
    tx1: scheme.dk1 || '000000', bg1: scheme.lt1 || 'FFFFFF', tx2: scheme.dk2 || '44546A', bg2: scheme.lt2 || 'E7E6E6',
    dark1: scheme.dk1 || '000000', light1: scheme.lt1 || 'FFFFFF', dark2: scheme.dk2 || '44546A', light2: scheme.lt2 || 'E7E6E6',
    text1: scheme.dk1 || '000000', background1: scheme.lt1 || 'FFFFFF', text2: scheme.dk2 || '44546A', background2: scheme.lt2 || 'E7E6E6',
    hyperlink: scheme.hlink || '0563C1',
  });
  for (let i = 1; i <= 6; i++) scheme['accent' + i] = scheme['accent' + i] || ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'][i - 1];
  const fs = X.desc(themeDoc, 'fontScheme');
  const major = X.desc(X.desc(fs, 'majorFont'), 'latin')?.getAttribute('typeface') || 'Calibri Light';
  const minor = X.desc(X.desc(fs, 'minorFont'), 'latin')?.getAttribute('typeface') || 'Calibri';
  return { scheme, major, minor };
}

const hexRgb = (hex) => {
  let h = String(hex || '000000').replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(h)) h = /^[0-9a-f]{3}$/i.test(h) ? h.replace(/./g, '$&$&') : '000000';
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
};
const toRgb = (hex) => PDFLib.rgb(...hexRgb(hex));

/**
 * Color de un elemento de relleno OOXML (solidFill, gs, fillRef…) con sus
 * modificadores (lumMod, lumOff, shade, tint, alpha). Devuelve {hex, alpha} o null.
 */
function ooxmlColor(el, scheme, phClr) {
  if (!el) return null;
  const c = X.kid(el, 'srgbClr') || X.kid(el, 'schemeClr') || X.kid(el, 'sysClr') || X.kid(el, 'prstClr') || X.kid(el, 'scrgbClr');
  if (!c) return null;
  let hex;
  if (c.localName === 'srgbClr') hex = c.getAttribute('val');
  else if (c.localName === 'schemeClr') hex = c.getAttribute('val') === 'phClr' && phClr ? phClr : scheme[c.getAttribute('val')] || scheme.tx1;
  else if (c.localName === 'sysClr') hex = c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000');
  else if (c.localName === 'scrgbClr') hex = ['r', 'g', 'b'].map((k) => Math.round(Math.min(1, X.num(c, k) / 100000) * 255).toString(16).padStart(2, '0')).join('');
  else hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' }[c.getAttribute('val')] || '000000';
  if (!hex) return null;
  let [r, g, b] = hexRgb(hex);
  const f = (v) => Math.max(0, Math.min(1, v));
  const mod = X.kid(c, 'lumMod'), off = X.kid(c, 'lumOff'), shade = X.kid(c, 'shade'), tint = X.kid(c, 'tint');
  if (mod || off) {
    const mm = X.num(mod, 'val', 100000) / 100000, oo = X.num(off, 'val', 0) / 100000;
    [r, g, b] = [r, g, b].map((v) => f(v * mm + oo));
  }
  if (shade) { const s = X.num(shade, 'val') / 100000; [r, g, b] = [r, g, b].map((v) => v * s); }
  if (tint) { const t = X.num(tint, 'val') / 100000; [r, g, b] = [r, g, b].map((v) => v * t + (1 - t)); }
  const alpha = X.kid(c, 'alpha');
  const out = [r, g, b].map((v) => Math.round(f(v) * 255).toString(16).padStart(2, '0')).join('');
  return { hex: out, alpha: alpha ? X.num(alpha, 'val') / 100000 : 1 };
}
