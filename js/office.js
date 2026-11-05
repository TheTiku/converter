'use strict';

/* ════════════════════════════════════════════════════════════════
   Conversiones de Office (básicas), todo en el navegador:
   Word/Excel/PowerPoint → PDF y PDF → Word/Excel/PowerPoint.
   Word → docx.js · PowerPoint → pptx.js · PDF → Word/PowerPoint → pdfconv.js
   ════════════════════════════════════════════════════════════════ */

const LIBS = {
  xlsx: 'vendor/xlsx.full.min.js',
  pptx: 'vendor/pptxgen.bundle.js',
};

Object.assign(ICONS, {
  word: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="m8 12 1.3 6L12 13l2.7 5L16 12"/>',
  excel: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="m9 12 6 6M15 12l-6 6"/>',
  ppt: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M10 18v-6h2.5a2 2 0 0 1 0 4H10"/>',
});

const ACCEPT = {
  word: { exts: ['docx'], label: 'documento Word (.docx)', badge: 'DOCX', color: '#2b579a' },
  excel: { exts: ['xlsx', 'xls', 'xlsm', 'ods', 'csv'], label: 'hoja de cálculo (Excel, ODS o CSV)', badge: 'XLS', color: '#217346' },
  ppt: { exts: ['pptx'], label: 'presentación PowerPoint (.pptx)', badge: 'PPTX', color: '#c43e1c' },
};

/* ───────────── Excel → PDF ───────────── */

async function excelToPdf(bytes, opts, progress) {
  progress(0.05, 'Leyendo la hoja de cálculo…');
  await U.loadScript(LIBS.xlsx);
  let wb;
  try {
    wb = XLSX.read(bytes, { type: 'array', cellDates: true });
  } catch (_) {
    throw new Error('No se pudo leer la hoja de cálculo.');
  }
  const doc = await PDFLib.PDFDocument.create();
  const bank = FontBank.forDoc(doc);
  await bank.prepare(['Calibri']);
  const fonts = { '': await bank.get('Calibri', false, false), b: await bank.get('Calibri', true, false) };
  const A4 = [595.28, 841.89];
  const m = 36;
  const isNum = (s) => /^[-+(]?[€$£]?\s?[\d.,\s]+%?\)?\s?[€$£]?$/.test(s) && /\d/.test(s);
  let sheetsDone = 0;

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    let rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true });
    rows = rows.map((r) => r.map((c) => String(c ?? '')));
    while (rows.length && rows[rows.length - 1].every((c) => !c.trim())) rows.pop();
    if (!rows.length) continue;
    let nCols = Math.max(...rows.map((r) => r.length));
    while (nCols > 0 && rows.every((r) => !(r[nCols - 1] || '').trim())) nCols--;
    if (!nCols) continue;
    sheetsDone++;

    const size = 9;
    const widths = Array.from({ length: nCols }, (_, j) => {
      const w = Math.max(...rows.map((r, i) => (() => { const f = fonts[i === 0 && opts.header ? 'b' : '']; return f.widthOfTextAtSize(cleanText(r[j] || '', f), size); })()));
      return Math.min(Math.max(w + 10, 28), 220);
    });
    const total = widths.reduce((a, b) => a + b, 0);
    const landscape = opts.orient === 'landscape' || (opts.orient === 'auto' && total > A4[0] - m * 2);
    const [PW, PH] = landscape ? [A4[1], A4[0]] : A4;
    const cw = PW - m * 2;
    const k = Math.min(1, cw / total);
    const fs = Math.max(4, size * k);
    const ws2 = widths.map((w) => w * k);
    const rh = fs * 1.7;

    let page, y;
    const newPage = (first) => {
      page = doc.addPage([PW, PH]);
      y = PH - m;
      if (first && wb.SheetNames.length > 1) {
        page.drawText(cleanText(name, fonts.b), { x: m, y: y - 12, size: 12, font: fonts.b });
        y -= 24;
      }
      if (!first && opts.header) drawRow(rows[0], true);
    };
    const fit = (t, font, w) => {
      t = cleanText(t, font);
      if (font.widthOfTextAtSize(t, fs) <= w) return t;
      while (t && font.widthOfTextAtSize(t + '…', fs) > w) t = t.slice(0, -1);
      return t + '…';
    };
    const drawRow = (r, head) => {
      let x = m;
      const font = head ? fonts.b : fonts[''];
      for (let j = 0; j < nCols; j++) {
        const w = ws2[j];
        if (opts.grid || head) {
          page.drawRectangle({
            x, y: y - rh, width: w, height: rh,
            borderWidth: opts.grid ? 0.4 : 0, borderColor: PDFLib.rgb(0.75, 0.75, 0.8),
            color: head ? PDFLib.rgb(0.93, 0.95, 0.93) : undefined,
          });
        }
        const t = fit(r[j] || '', font, w - 6);
        if (t) {
          const tw = font.widthOfTextAtSize(t, fs);
          const tx = !head && isNum(t) ? x + w - 3 - tw : x + 3;
          page.drawText(t, { x: tx, y: y - rh / 2 - fs * 0.35, size: fs, font, color: PDFLib.rgb(0.1, 0.1, 0.12) });
        }
        x += w;
      }
      y -= rh;
    };
    newPage(true);
    rows.forEach((r, i) => {
      if (y - rh < m) newPage(false);
      drawRow(r, i === 0 && opts.header);
    });
    progress(sheetsDone / wb.SheetNames.length, `Hoja «${name}» lista…`);
  }
  if (!sheetsDone) throw new Error('La hoja de cálculo está vacía.');
  return doc.save();
}

/* ───────────── PDF → estructura (títulos, párrafos, listas) ───────────── */

async function pdfBlocks(pdf, progress) {
  const lines = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    progress && progress((i - 1) / pdf.numPages, `Leyendo página ${i} de ${pdf.numPages}…`);
    const tc = await (await pdf.getPage(i)).getTextContent();
    let cur = null;
    for (const it of tc.items) {
      if (!it.str && !it.hasEOL) continue;
      const size = Math.round(Math.hypot(it.transform[2], it.transform[3]) * 10) / 10;
      const y = it.transform[5];
      if (!cur || Math.abs(cur.y - y) > Math.max(2, size * 0.5)) {
        if (cur) lines.push(cur);
        cur = { page: i, y, size, text: '', chars: {} };
      }
      cur.text += it.str;
      if (it.str.trim()) cur.chars[size] = (cur.chars[size] || 0) + it.str.length;
      if (it.hasEOL) { lines.push(cur); cur = null; }
    }
    if (cur) lines.push(cur);
  }
  const real = lines.filter((l) => l.text.trim());
  if (!real.length) throw new Error('Este PDF no contiene texto (¿es escaneado? Prueba antes con OCR).');
  const freq = {};
  for (const l of real) {
    const best = Object.entries(l.chars).sort((a, b) => b[1] - a[1])[0];
    if (best) l.size = +best[0];
    freq[l.size] = (freq[l.size] || 0) + l.text.length;
  }
  const body = +Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
  const blocks = [];
  let para = null;
  let prev = null;
  const clean = (s) => s.replace(/\s+/g, ' ').replace(/(\w)- (\w)/g, '$1$2').trim();
  const flush = () => { if (para) { blocks.push({ type: 'p', text: clean(para.join(' ')) }); para = null; } };
  for (const l of real) {
    const text = l.text.trim();
    const r = l.size / body;
    const samePage = prev && prev.page === l.page;
    const gap = samePage ? prev.y - l.y : 0;
    if (prev && !samePage) { flush(); blocks.push({ type: 'break' }); }
    const bullet = text.match(/^([•●▪◦‣∙·\-–*]|\d+[.)])\s+(.*)$/);
    const last = blocks[blocks.length - 1];
    if (r >= 1.2 && text.length < 160) {
      flush();
      blocks.push({ type: 'h', level: r >= 1.7 ? 1 : r >= 1.35 ? 2 : 3, text });
    } else if (bullet) {
      flush();
      const ordered = /^\d/.test(bullet[1]);
      blocks.push({ type: 'li', ordered, marker: ordered ? bullet[1].replace(')', '.') : '-', text: bullet[2] });
    } else if (!para && last && last.type === 'li' && samePage && gap < l.size * 1.6) {
      last.text += ' ' + text; // continuación de la viñeta
    } else {
      if (para && gap > l.size * 1.9) flush();
      (para = para || []).push(text);
    }
    prev = l;
  }
  flush();
  return blocks;
}

function blocksToMarkdown(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.type === 'break') continue;
    const prevLi = out.length && out[out.length - 1].li;
    if (b.type === 'li') {
      if (out.length && !prevLi) out.push({ t: '' });
      out.push({ t: `${b.marker} ${b.text}`, li: true });
    } else {
      if (out.length) out.push({ t: '' });
      out.push({ t: b.type === 'h' ? `${'#'.repeat(b.level)} ${b.text}` : b.text });
    }
  }
  return out.map((o) => o.t).join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

/* ───────────── PDF → Excel ───────────── */

/** Interpreta el texto de una celda: número, moneda, porcentaje o fecha (con su formato de Excel). */
function parseCell(text) {
  const t = String(text).trim();
  if (!t) return null;
  // fechas: 05/11/2026, 5-11-26, 2026-11-05
  let m = t.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$/);
  if (m) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    const d = new Date(Date.UTC(y, +m[2] - 1, +m[1]));
    if (d.getUTCDate() === +m[1] && d.getUTCMonth() === +m[2] - 1) return { v: d, z: 'dd/mm/yyyy', text: t };
  }
  m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if (!isNaN(d)) return { v: d, z: 'dd/mm/yyyy', text: t };
  }
  const cur = /€|\bEUR\b/.test(t) ? '€' : /\$|\bUSD\b/.test(t) ? '$' : /£|\bGBP\b/.test(t) ? '£' : null;
  const pct = /%\s*$/.test(t);
  // códigos con ceros a la izquierda, teléfonos, IBAN… se quedan como texto
  if (!cur && !pct && ((/^0\d/.test(t) && !/[.,]/.test(t)) || /\d\s+\d/.test(t))) return { v: t, text: t };
  let n = t.replace(/€|\bEUR\b|\$|\bUSD\b|£|\bGBP\b|%/g, '').replace(/[\s  ]/g, '');
  let neg = false;
  if (/^\(.*\)$/.test(n)) { neg = true; n = n.slice(1, -1); }
  if (/^[-−–]/.test(n)) { neg = true; n = n.slice(1); } else if (/^\+/.test(n)) n = n.slice(1);
  if (/[-−]$/.test(n)) { neg = true; n = n.slice(0, -1); }
  if (!/^\d[\d.,]*$/.test(n) || n.replace(/\D/g, '').length > 15) return { v: t, text: t };
  let dec = '', thousands = false;
  if (n.includes('.') && n.includes(',')) {
    dec = n.lastIndexOf(',') > n.lastIndexOf('.') ? ',' : '.';
    thousands = true;
  } else if (n.includes(',')) {
    if (/^\d{1,3}(,\d{3}){2,}$/.test(n)) thousands = true; // 1,234,567
    else if (/^\d+,\d+$/.test(n)) dec = ',';
    else return { v: t, text: t };
  } else if (n.includes('.')) {
    if (/^\d{1,3}(\.\d{3})+$/.test(n)) thousands = true; // 1.234 / 1.234.567
    else if (/^\d+\.\d+$/.test(n)) dec = '.';
    else return { v: t, text: t };
  }
  const tsep = dec === ',' ? '.' : dec === '.' ? ',' : thousands ? (n.includes('.') ? '.' : ',') : '';
  const clean = n.split(tsep || '\u0000').join('').replace(dec || '\u0000', '.');
  let v = parseFloat(clean);
  if (isNaN(v)) return { v: t, text: t };
  if (neg) v = -v;
  const decimals = dec ? (n.split(dec).pop().length) : 0;
  const numFmt = `#,##0${decimals ? '.' + '0'.repeat(Math.min(decimals, 4)) : ''}`;
  if (pct) return { v: v / 100, z: decimals ? `0.${'0'.repeat(Math.min(decimals, 4))}%` : '0%', text: t };
  if (cur === '€') return { v, z: `${numFmt} "€"`, text: t };
  if (cur) return { v, z: `"${cur}"${numFmt}`, text: t };
  return { v, z: decimals || thousands ? numFmt : undefined, text: t };
}

/** Extrae las tablas de cada página. Devuelve [{name, rows: [[{v, z, text}]]}] */
async function pdfTables(pdf, progress) {
  const sheets = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    progress((i - 1) / pdf.numPages, `Buscando tablas en la página ${i}…`);
    const tc = await (await pdf.getPage(i)).getTextContent();
    const items = tc.items.filter((it) => it.str && it.str.trim()).map((it) => ({
      x: it.transform[4], y: it.transform[5], w: it.width, s: it.str, size: Math.hypot(it.transform[2], it.transform[3]) || 10,
    }));
    // 1) filas por línea base
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    for (const it of items) {
      const r = lines[lines.length - 1];
      if (r && Math.abs(r.y - it.y) <= Math.max(2, Math.min(r.size, it.size) * 0.4)) { r.items.push(it); r.size = Math.max(r.size, it.size); }
      else lines.push({ y: it.y, size: it.size, items: [it] });
    }
    // 2) celdas: fragmentos muy próximos forman una celda
    for (const r of lines) {
      r.items.sort((a, b) => a.x - b.x);
      const cells = [];
      for (const it of r.items) {
        const last = cells[cells.length - 1];
        const gap = last ? it.x - last.x1 : 0;
        if (last && gap < it.size * 0.8) {
          last.text += (gap > it.size * 0.15 && !/\s$/.test(last.text) ? ' ' : '') + it.s;
          last.x1 = Math.max(last.x1, it.x + it.w);
        } else cells.push({ x0: it.x, x1: it.x + it.w, text: it.s });
      }
      r.cells = cells;
    }
    // 3) columnas: zonas horizontales ocupadas por muchas filas con varias celdas
    const multi = lines.filter((r) => r.cells.length >= 2);
    const counts = {};
    multi.forEach((r) => { counts[r.cells.length] = (counts[r.cells.length] || 0) + 1; });
    const mode = +(Object.entries(counts).sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]?.[0] || 0);
    const tableRows = multi.filter((r) => r.cells.length >= Math.max(2, mode - 1));
    let cols = [];
    if (tableRows.length >= 2) {
      const minX = Math.floor(Math.min(...tableRows.flatMap((r) => r.cells.map((c) => c.x0))));
      const maxX = Math.ceil(Math.max(...tableRows.flatMap((r) => r.cells.map((c) => c.x1))));
      const cover = new Uint16Array(maxX - minX + 2);
      for (const r of tableRows) for (const c of r.cells) for (let x = Math.floor(c.x0); x <= Math.ceil(c.x1); x++) cover[x - minX]++;
      const need = Math.max(1, Math.ceil(tableRows.length * 0.2));
      let start = -1;
      for (let x = 0; x <= cover.length; x++) {
        const on = x < cover.length && cover[x] >= need;
        if (on && start < 0) start = x;
        if (!on && start >= 0) { cols.push({ x0: start + minX, x1: x - 1 + minX }); start = -1; }
      }
      // huecos muy pequeños no separan columnas
      cols = cols.reduce((acc, c) => { const l = acc[acc.length - 1]; if (l && c.x0 - l.x1 < 4) l.x1 = c.x1; else acc.push(c); return acc; }, []);
    }
    if (cols.length < 2) {
      // sin estructura de tabla: una columna por cada inicio de texto distinto
      const xs = lines.flatMap((r) => r.cells.map((c) => c.x0)).sort((a, b) => a - b);
      for (const x of xs) { const l = cols[cols.length - 1]; if (l && x - l.x0 < 12) l.x1 = Math.max(l.x1, x); else cols.push({ x0: x, x1: x }); }
    }
    const colOf = (c) => {
      let best = -1, bestOv = 0;
      cols.forEach((k, j) => { const ov = Math.min(c.x1, k.x1) - Math.max(c.x0, k.x0); if (ov > bestOv) { bestOv = ov; best = j; } });
      if (best >= 0) return best;
      let d = Infinity;
      cols.forEach((k, j) => { const dist = Math.abs((c.x0 + c.x1) / 2 - (k.x0 + k.x1) / 2); if (dist < d) { d = dist; best = j; } });
      return best;
    };
    // 4) filas de la hoja + unión de líneas de continuación (descripciones en varias líneas)
    // texto fuera de las tablas (títulos, párrafos): empieza en la columna donde empieza el texto
    const colOfStart = (c) => {
      let k = 0;
      cols.forEach((col, j) => { if (col.x0 <= c.x0 + 4) k = j; });
      return k;
    };
    const rows = [];
    let prev = null;
    for (const r of lines) {
      const arr = Array(cols.length).fill('');
      const isTable = r.cells.length >= Math.max(2, mode - 1);
      for (const c of r.cells) { const k = isTable ? colOf(c) : colOfStart(c); arr[k] = arr[k] ? `${arr[k]} ${c.text}` : c.text; }
      const filled = arr.filter(Boolean).length;
      const numeric = arr.some((v) => v && typeof parseCell(v).v !== 'string');
      if (prev && prev.isTable && !isTable && filled <= Math.max(1, cols.length - 2) && !numeric && prev.y - r.y <= prev.size * 1.75) {
        arr.forEach((v, k) => { if (v) prev.arr[k] = prev.arr[k] ? `${prev.arr[k]} ${v}` : v; });
        prev.y = r.y;
        continue;
      }
      prev = { arr, isTable, y: r.y, size: r.size };
      rows.push(prev);
    }
    const parsed = rows.map((r) => r.arr.map((v) => (v ? parseCell(v) : null)));
    // columnas de códigos (códigos postales, referencias…): si alguno empieza por 0, toda la columna es texto
    for (let k = 0; k < cols.length; k++) {
      if (!parsed.some((r) => r[k] && /^0\d+$/.test(r[k].text))) continue;
      parsed.forEach((r) => { if (r[k] && /^\d+$/.test(r[k].text)) r[k] = { v: r[k].text, text: r[k].text }; });
    }
    sheets.push({ name: `Página ${i}`, rows: parsed });
  }
  return sheets;
}

/* ───────────── Herramientas ───────────── */

TOOLS.push(
  {
    id: 'word-a-pdf', name: 'Word a PDF', cat: 'convertir', icon: 'word', color: '#2b579a',
    desc: 'Convierte documentos Word (.docx) a PDF manteniendo el diseño.',
    accept: ACCEPT.word, multiple: false, mode: 'files', action: 'Convertir a PDF',
    options: () => `<small class="note">Se conservan estilos, fuentes (con equivalentes libres de las de Office), listas, tablas, imágenes, gráficos, columnas, encabezados y pies de página. Las notas al pie se colocan al final del documento.</small>`,
    async run({ files, progress }) {
      const f = files[0].file;
      const { bytes, skipped, pages } = await Docx.toPdf(await U.bytes(f), progress);
      return {
        files: [{ name: `${U.baseName(f.name)}.pdf`, blob: U.pdfBlob(bytes) }],
        message: `${pages} página${pages === 1 ? '' : 's'}.` + (skipped ? ` ${skipped} imagen(es) en formato EMF/WMF no se han podido incluir.` : ''),
      };
    },
  },
  {
    id: 'excel-a-pdf', name: 'Excel a PDF', cat: 'convertir', icon: 'excel', color: '#217346',
    desc: 'Convierte hojas de cálculo (Excel, ODS, CSV) a PDF ajustadas al ancho de la página.',
    accept: ACCEPT.excel, multiple: false, mode: 'files', action: 'Convertir a PDF',
    options: () => `
      ${select('orient', 'Orientación', [['auto', 'Automática'], ['portrait', 'Vertical'], ['landscape', 'Horizontal']], 'auto')}
      <label class="check"><input type="checkbox" name="header" checked> La primera fila es un encabezado (se repite en cada página)</label>
      <label class="check"><input type="checkbox" name="grid" checked> Mostrar cuadrícula</label>
      <small class="note">Cada hoja empieza en una página nueva. No se conservan colores, celdas combinadas ni gráficos.</small>`,
    async run({ files, opts, progress }) {
      const f = files[0].file;
      const bytes = await excelToPdf(await U.bytes(f), opts, progress);
      return [{ name: `${U.baseName(f.name)}.pdf`, blob: U.pdfBlob(bytes) }];
    },
  },
  {
    id: 'powerpoint-a-pdf', name: 'PowerPoint a PDF', cat: 'convertir', icon: 'ppt', color: '#c43e1c',
    desc: 'Convierte presentaciones PowerPoint (.pptx) a PDF, una diapositiva por página.',
    accept: ACCEPT.ppt, multiple: false, mode: 'files', action: 'Convertir a PDF',
    options: () => `<small class="note">Se conservan fondos, formas, degradados, textos con sus fuentes, viñetas, tablas, imágenes, SmartArt y gráficos. Las animaciones y efectos 3D no se reproducen.</small>`,
    async run({ files, progress }) {
      const f = files[0].file;
      const { bytes, skipped, count } = await Pptx.toPdf(await U.bytes(f), progress);
      return {
        files: [{ name: `${U.baseName(f.name)}.pdf`, blob: U.pdfBlob(bytes) }],
        message: `${count} diapositiva${count > 1 ? 's' : ''} convertida${count > 1 ? 's' : ''}.` + (skipped ? ` ${skipped} imagen(es) en formato no compatible se han omitido.` : ''),
      };
    },
  },
  {
    id: 'pdf-a-word', name: 'PDF a Word', cat: 'convertir', icon: 'word', color: '#1e4e8c',
    desc: 'Convierte un PDF en un documento Word (.docx) editable.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Convertir a Word',
    options: () => `<small class="note">Se recuperan párrafos, títulos, negritas, cursivas, fuentes, colores, tablas e imágenes. Los diseños muy complejos (varias columnas, cuadros superpuestos) pueden necesitar retoques.</small>`,
    async run({ files, progress }) {
      const f = files[0].file;
      const pdf = await U.pdfJs(await U.bytes(f));
      try {
        return [{ name: `${U.baseName(f.name)}.docx`, blob: await PdfConv.toDocx(pdf, progress) }];
      } finally {
        pdf.destroy();
      }
    },
  },
  {
    id: 'pdf-a-excel', name: 'PDF a Excel', cat: 'convertir', icon: 'excel', color: '#185c37',
    desc: 'Extrae las tablas y datos de un PDF a una hoja de Excel.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Convertir a Excel',
    options: () => `
      <div class="field"><label class="lbl">Organización</label>
        ${choice('layout', 'pages', 'Una hoja por página', '', true)}
        ${choice('layout', 'one', 'Todo en una sola hoja', '')}
      </div>
      <small class="note">Detecta las columnas aunque estén alineadas a la derecha y une las descripciones de varias líneas. Los números, importes, porcentajes y fechas se convierten en valores de Excel con su formato. Para PDF escaneados, usa antes OCR.</small>`,
    async run({ files, opts, progress }) {
      const f = files[0].file;
      const pdf = await U.pdfJs(await U.bytes(f));
      let sheets;
      try {
        sheets = await pdfTables(pdf, progress);
      } finally {
        pdf.destroy();
      }
      if (!sheets.some((s) => s.rows.length)) throw new Error('Este PDF no contiene texto (¿es escaneado? Prueba antes con OCR).');
      progress(1, 'Creando el archivo Excel…');
      await U.loadScript(LIBS.xlsx);
      const wb = XLSX.utils.book_new();
      const addSheet = (name, rows) => {
        const ws = XLSX.utils.aoa_to_sheet(rows.map((r) => r.map((c) => (c ? c.v : null))), { cellDates: true });
        // formato de número, moneda, porcentaje o fecha de cada celda
        rows.forEach((r, R) => r.forEach((c, C) => {
          const ref = XLSX.utils.encode_cell({ r: R, c: C });
          if (c && c.z && ws[ref]) ws[ref].z = c.z;
        }));
        const n = Math.max(0, ...rows.map((r) => r.length));
        ws['!cols'] = Array.from({ length: n }, (_, k) => ({ wch: Math.min(60, Math.max(8, ...rows.map((r) => (r[k] ? r[k].text.length : 0) + 2))) }));
        XLSX.utils.book_append_sheet(wb, ws, name);
      };
      if (opts.layout === 'one') addSheet('PDF', sheets.flatMap((s, i) => (i ? [[], ...s.rows] : s.rows)));
      else sheets.forEach((s) => s.rows.length && addSheet(s.name, s.rows));
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      return [{ name: `${U.baseName(f.name)}.xlsx`, blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) }];
    },
  },
  {
    id: 'pdf-a-powerpoint', name: 'PDF a PowerPoint', cat: 'convertir', icon: 'ppt', color: '#a5361a',
    desc: 'Convierte cada página del PDF en una diapositiva de PowerPoint (.pptx) con texto editable.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Convertir a PowerPoint',
    options: () => `
      <div class="field"><label class="lbl">Modo</label>
        ${choice('mode', 'edit', 'Texto editable', 'El fondo de la página queda como imagen y el texto, en cuadros editables', true)}
        ${choice('mode', 'image', 'Imagen exacta', 'Cada página es una imagen: idéntica, pero no editable')}
      </div>`,
    async run({ files, opts, progress }) {
      const f = files[0].file;
      const bytes = await U.bytes(f);
      const pdf = await U.pdfJs(bytes);
      try {
        const { blob, note } = await PdfConv.toPptx(pdf, bytes, opts.mode, progress);
        return { files: [{ name: `${U.baseName(f.name)}.pptx`, blob }], message: note };
      } finally {
        pdf.destroy();
      }
    },
  }
);
