'use strict';

/**
 * Conversión de PDF a formatos editables de Office:
 *  - PDF → Word: párrafos, títulos, negritas/cursivas, fuentes, colores, tablas e imágenes.
 *  - PDF → PowerPoint: fondo de la página sin texto + cuadros de texto editables en su posición.
 */
const PdfConv = (() => {
  const { Util, OPS } = pdfjsLib;

  /* ── nombres de fuente del PDF → familia de Office ── */
  function familyOf(raw, generic) {
    const n = String(raw || '').replace(/^[A-Z]{6}\+/, '');
    const known = [
      [/times/i, 'Times New Roman'], [/arial|helvetica|liberationsans|arimo/i, 'Arial'], [/calibri|carlito/i, 'Calibri'],
      [/cambria|caladea/i, 'Cambria'], [/courier|cousine/i, 'Courier New'], [/georgia/i, 'Georgia'], [/verdana/i, 'Verdana'],
      [/garamond/i, 'Garamond'], [/tahoma/i, 'Tahoma'], [/trebuchet/i, 'Trebuchet MS'], [/segoe/i, 'Segoe UI'], [/aptos/i, 'Aptos'],
      [/century ?gothic/i, 'Century Gothic'], [/book ?antiqua|palatino/i, 'Palatino Linotype'], [/consolas/i, 'Consolas'],
    ];
    for (const [re, name] of known) if (re.test(n)) return name;
    const base = n.replace(/[-,_]?(Bold|Italic|Oblique|Regular|Semibold|SemiBold|Light|Medium|Black|Heavy|Book|Roman|MT|PS|PSMT|Std|Pro)+$/gi, '').replace(/(MT|PSMT|PS)$/, '');
    const spaced = base.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[-_]+/g, ' ').trim();
    if (spaced && !/^(g_|f\d|font)/i.test(spaced)) return spaced;
    return generic === 'serif' ? 'Times New Roman' : generic === 'monospace' ? 'Courier New' : 'Calibri';
  }

  /** Analiza una página: líneas (con segmentos y estilos) e imágenes, en coordenadas de vista (pt, Y hacia abajo). */
  async function analyze(page, { images = false } = {}) {
    const vp = page.getViewport({ scale: 1 });
    const ops = await page.getOperatorList();
    const tc = await page.getTextContent();
    const fonts = {};
    const fontOf = (name) => {
      if (fonts[name]) return fonts[name];
      let f = null;
      try { f = page.commonObjs.get(name); } catch (_) { /* no cargada */ }
      const raw = (f && f.name) || '';
      const generic = tc.styles[name]?.fontFamily;
      return (fonts[name] = {
        family: familyOf(raw, generic),
        bold: /bold|black|heavy|semibold|demi/i.test(raw) || !!(f && (f.bold || f.black)),
        italic: /italic|oblique/i.test(raw) || !!(f && f.italic),
      });
    };
    const items = [];
    for (const it of tc.items) {
      // pdf.js añade "espacios" artificiales que tapan los huecos entre columnas: los ignoramos
      if (!it.str || !it.str.trim()) continue;
      const tx = Util.transform(vp.transform, it.transform);
      const size = Math.hypot(tx[2], tx[3]);
      if (size < 1) continue;
      const rotated = Math.abs(tx[1]) > 0.1 * size || Math.abs(tx[2]) > 0.1 * size;
      items.push({ str: it.str, x: tx[4], base: tx[5], size, w: it.width * vp.scale, rotated, ...fontOf(it.fontName) });
    }
    // líneas
    items.sort((a, b) => a.base - b.base || a.x - b.x);
    const lines = [];
    for (const it of items) {
      if (it.rotated) continue;
      const tol = (L) => {
        const lo = Math.min(L.size, it.size), hi = Math.max(L.size, it.size);
        return Math.max(lo * 0.35, hi / lo > 1.25 ? hi * 0.5 : 0); // admite superíndices y subíndices
      };
      const l = lines.find((L) => Math.abs(L.base - it.base) < tol(L));
      if (l) {
        l.items.push(it);
        if (it.size > l.size) { l.size = it.size; l.base = it.base; }
      }
      else lines.push({ base: it.base, size: it.size, items: [it], minX: it.x });
    }
    for (const l of lines) {
      l.items.sort((a, b) => a.x - b.x);
      const segs = [];
      for (const it of l.items) {
        let seg = segs[segs.length - 1];
        const end = seg ? seg.x + seg.w : 0;
        const gap = seg ? it.x - end : 0;
        if (!seg || gap > Math.max(it.size * 1.6, 10)) {
          seg = { x: it.x, w: 0, runs: [], size: it.size };
          segs.push(seg);
        } else if (gap > it.size * 0.15 && !/\s$/.test(seg.runs[seg.runs.length - 1].text) && !/^\s/.test(it.str)) {
          seg.runs[seg.runs.length - 1].text += ' ';
        }
        const last = seg.runs[seg.runs.length - 1];
        if (last && last.bold === it.bold && last.italic === it.italic && last.family === it.family && Math.abs(last.size - it.size) < 0.5) {
          last.text += it.str;
          last.w = it.x + it.w - last.x;
        } else seg.runs.push({ text: it.str, bold: it.bold, italic: it.italic, family: it.family, size: it.size, x: it.x, w: it.w, base: it.base });
        seg.w = Math.max(seg.w, it.x + it.w - seg.x);
        seg.size = Math.max(seg.size, it.size);
      }
      for (const sg of segs) for (const r of sg.runs) {
        if (r.size < l.size * 0.8 && Math.abs(r.base - l.base) > l.size * 0.15) r.vert = r.base < l.base ? 'superscript' : 'subscript';
      }
      l.segs = segs;
      l.x = segs[0].x;
      l.right = segs[segs.length - 1].x + segs[segs.length - 1].w;
      l.top = l.base - l.size * 0.9;
    }
    lines.sort((a, b) => a.base - b.base);

    // imágenes con su posición (siguiendo la matriz de transformación)
    const imgs = [];
    if (images) {
      let ctm = [1, 0, 0, 1, 0, 0];
      const stack = [];
      const seen = new Set();
      for (let k = 0; k < ops.fnArray.length; k++) {
        const fn = ops.fnArray[k], args = ops.argsArray[k];
        if (fn === OPS.save) stack.push(ctm);
        else if (fn === OPS.restore) ctm = stack.pop() || ctm;
        else if (fn === OPS.transform) ctm = Util.transform(ctm, args);
        else if (fn === OPS.paintFormXObjectBegin) { stack.push(ctm); if (args[0]) ctm = Util.transform(ctm, args[0]); }
        else if (fn === OPS.paintFormXObjectEnd) ctm = stack.pop() || ctm;
        else if (fn === OPS.paintImageXObject || fn === OPS.paintInlineImageXObject || fn === OPS.paintJpegXObject) {
          const m = Util.transform(vp.transform, ctm);
          const pts = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([u, v]) => [m[0] * u + m[2] * v + m[4], m[1] * u + m[3] * v + m[5]]);
          const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
          const box = { x: Math.min(...xs), top: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
          if (box.w < 12 || box.h < 12) continue;
          const key = typeof args[0] === 'string' ? args[0] + JSON.stringify(box) : null;
          if (key && seen.has(key)) continue;
          if (key) seen.add(key);
          const obj = typeof args[0] === 'string' ? await U.pdfObj(page, args[0]) : args[0];
          const canvas = U.pdfImageToCanvas(obj);
          if (canvas) imgs.push({ ...box, canvas });
        }
      }
    }
    return { w: vp.width, h: vp.height, lines, images: imgs };
  }

  /* ── color del texto: muestreo de píxeles ── */
  function sampleColor(orig, bg, scale, x, top, w, h) {
    const X0 = Math.max(0, Math.floor(x * scale)), Y0 = Math.max(0, Math.floor(top * scale));
    const W0 = Math.min(orig.width - X0, Math.ceil(w * scale)), H0 = Math.min(orig.height - Y0, Math.ceil(h * scale));
    if (W0 < 1 || H0 < 1) return '000000';
    const o = orig.getContext('2d').getImageData(X0, Y0, W0, H0).data;
    const b = bg ? bg.getContext('2d').getImageData(X0, Y0, W0, H0).data : null;
    let ref = [255, 255, 255];
    if (!b) {
      // color de fondo = el más frecuente (cuantizado)
      const counts = new Map();
      for (let i = 0; i < o.length; i += 16) { const k = ((o[i] >> 4) << 8) | ((o[i + 1] >> 4) << 4) | (o[i + 2] >> 4); counts.set(k, (counts.get(k) || 0) + 1); }
      const top1 = [...counts.entries()].sort((p, q) => q[1] - p[1])[0][0];
      ref = [((top1 >> 8) & 15) * 17, ((top1 >> 4) & 15) * 17, (top1 & 15) * 17];
    }
    let maxD = 0;
    const d = new Float32Array(o.length / 4);
    for (let i = 0, p = 0; i < o.length; i += 4, p++) {
      const r = b ? [b[i], b[i + 1], b[i + 2]] : ref;
      d[p] = Math.abs(o[i] - r[0]) + Math.abs(o[i + 1] - r[1]) + Math.abs(o[i + 2] - r[2]);
      if (d[p] > maxD) maxD = d[p];
    }
    if (maxD < 40) return '000000';
    let R = 0, G = 0, B = 0, n = 0;
    for (let i = 0, p = 0; i < o.length; i += 4, p++) {
      if (d[p] >= maxD * 0.7) { R += o[i]; G += o[i + 1]; B += o[i + 2]; n++; }
    }
    return [R, G, B].map((v) => Math.round(v / n).toString(16).padStart(2, '0')).join('');
  }

  /** Color de fondo dominante de una zona (para el sombreado de celdas). */
  function dominantColor(canvas, scale, x, top, w, h) {
    const X0 = Math.max(0, Math.floor(x * scale)), Y0 = Math.max(0, Math.floor(top * scale));
    const W0 = Math.min(canvas.width - X0, Math.ceil(w * scale)), H0 = Math.min(canvas.height - Y0, Math.ceil(h * scale));
    if (W0 < 1 || H0 < 1) return null;
    const o = canvas.getContext('2d').getImageData(X0, Y0, W0, H0).data;
    const buckets = new Map();
    for (let i = 0; i < o.length; i += 8) {
      const k = ((o[i] >> 4) << 8) | ((o[i + 1] >> 4) << 4) | (o[i + 2] >> 4);
      const b = buckets.get(k) || [0, 0, 0, 0];
      b[0] += o[i]; b[1] += o[i + 1]; b[2] += o[i + 2]; b[3]++;
      buckets.set(k, b);
    }
    const best = [...buckets.values()].sort((a, b) => b[3] - a[3])[0];
    return best ? best.slice(0, 3).map((v) => Math.round(v / best[3]).toString(16).padStart(2, '0')).join('') : null;
  }

  async function renderCanvas(pdf, n, scale) {
    return (await U.render(pdf, n, { scale })).canvas;
  }

  /* ════════════ PDF → Word ════════════ */
  const xe = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');

  async function toDocx(pdf, progress) {
    const pages = [];
    const SC = 1.5;
    for (let i = 1; i <= pdf.numPages; i++) {
      progress(((i - 1) / pdf.numPages) * 0.8, `Analizando página ${i} de ${pdf.numPages}…`);
      const page = await pdf.getPage(i);
      const an = await analyze(page, { images: true });
      const canvas = await renderCanvas(pdf, i, SC);
      for (const l of an.lines) for (const s of l.segs) {
        for (const r of s.runs) r.color = sampleColor(canvas, null, SC, r.x, l.base - r.size * 0.75, r.w, r.size * 0.8);
        s.bg = dominantColor(canvas, SC, s.x - 1, l.base - s.size * 1.05, s.w + 2, s.size * 1.4);
      }
      canvas.width = canvas.height = 0;
      pages.push(an);
      page.cleanup();
    }
    const allLines = pages.flatMap((p) => p.lines);
    if (!allLines.length && !pages.some((p) => p.images.length)) throw new Error('Este PDF no contiene texto (¿es escaneado? Prueba antes con OCR).');
    // tamaño y fuente dominantes
    const freq = {}, fam = {};
    for (const l of allLines) for (const s of l.segs) for (const r of s.runs) {
      const k = Math.round(r.size * 2) / 2;
      freq[k] = (freq[k] || 0) + r.text.length;
      fam[r.family] = (fam[r.family] || 0) + r.text.length;
    }
    const body = +(Object.entries(freq).sort((a, b) => b[1] - a[1])[0]?.[0] || 11);
    const mainFont = Object.entries(fam).sort((a, b) => b[1] - a[1])[0]?.[0] || 'Calibri';
    const P0 = pages[0];
    const minXs = pages.map((p) => Math.min(...p.lines.map((l) => l.x), ...p.images.map((m) => m.x), p.w)).filter((v) => v < P0.w);
    const maxXs = pages.map((p) => Math.max(...p.lines.map((l) => l.right), ...p.images.map((m) => m.x + m.w), 0)).filter((v) => v > 0);
    const median = (a) => (a.length ? a.slice().sort((x, y) => x - y)[Math.floor(a.length / 2)] : 72);
    const L = Math.max(28, Math.min(110, median(minXs)));
    const R = Math.max(L + 100, Math.min(P0.w - 28, median(maxXs)));
    const topM = Math.max(28, Math.min(110, median(pages.map((p) => Math.min(...p.lines.map((l) => l.top), ...p.images.map((m) => m.top), 72)))));

    const media = [];
    const bodyXml = [];
    const runXml = (r) => {
      const props = [
        r.family && r.family !== mainFont ? `<w:rFonts w:ascii="${xe(r.family)}" w:hAnsi="${xe(r.family)}" w:cs="${xe(r.family)}"/>` : '',
        r.bold ? '<w:b/>' : '', r.italic ? '<w:i/>' : '', r.vert ? `<w:vertAlign w:val="${r.vert}"/>` : '',
        r.color && r.color !== '000000' && !/^(0|1|2)[0-9a-f](0|1|2)[0-9a-f](0|1|2)[0-9a-f]$/i.test(r.color) ? `<w:color w:val="${r.color}"/>` : '',
        ((sz) => `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`)(Math.max(2, Math.round((r.vert ? r.size / 0.65 : r.size) * 2))),
      ].join('');
      return r.text.split('\t').map((t, i) => `${i ? `<w:r><w:rPr>${props}</w:rPr><w:tab/></w:r>` : ''}${t ? `<w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${xe(t)}</w:t></w:r>` : ''}`).join('');
    };
    const imageXml = (canvasEl, w, h, jc, indent) => {
      const id = media.length + 1;
      media.push({ id, canvas: canvasEl });
      const k = Math.min(1, (R - L) / w);
      const cx = Math.round(w * k * 12700), cy = Math.round(h * k * 12700);
      return `<w:p><w:pPr>${jc ? `<w:jc w:val="${jc}"/>` : ''}${indent > 4 ? `<w:ind w:left="${Math.round(indent * 20)}"/>` : ''}<w:spacing w:before="0" w:after="120"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Imagen ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="imagen${id}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    };

    pages.forEach((pg, pi) => {
      progress(0.8 + (0.15 * pi) / pages.length, 'Creando el documento Word…');
      // detección de tablas: filas consecutivas con 2+ segmentos alineados
      const used = new Set();
      const tables = [];
      for (let i = 0; i < pg.lines.length; i++) {
        if (used.has(i) || pg.lines[i].segs.length < 2) continue;
        const group = [i];
        for (let j = i + 1; j < pg.lines.length; j++) {
          const prev = pg.lines[group[group.length - 1]], cur = pg.lines[j];
          if (cur.segs.length < 2 || cur.base - prev.base > Math.max(prev.size, cur.size) * 3) break;
          group.push(j);
        }
        if (group.length < 2) continue;
        const xs = group.flatMap((g) => pg.lines[g].segs.map((s) => s.x)).sort((a, b) => a - b);
        const anchors = [];
        for (const x of xs) { const a = anchors[anchors.length - 1]; if (a && x - a.max < 14) { a.max = x; a.n++; } else anchors.push({ min: x, max: x, n: 1 }); }
        const strong = anchors.filter((a) => a.n >= Math.max(2, group.length * 0.5));
        if (strong.length < 2) continue;
        group.forEach((g) => used.add(g));
        // ¿tabla o texto a varias columnas? Las columnas tienen celdas con mucho texto corrido
        const cellChars = group.flatMap((g) => pg.lines[g].segs.map((s) => s.runs.reduce((a, r) => a + r.text.length, 0)));
        const avgChars = cellChars.reduce((a, b) => a + b, 0) / cellChars.length;
        tables.push({ rows: group, anchors: strong, top: pg.lines[group[0]].top, cols: avgChars > 28 });
      }
      // elementos ordenados verticalmente
      const els = [
        ...pg.lines.map((l, i) => (used.has(i) ? null : { kind: 'line', top: l.top, l })).filter(Boolean),
        ...tables.map((t) => ({ kind: t.cols ? 'cols' : 'table', top: t.top, t })),
        ...pg.images.map((m) => ({ kind: 'img', top: m.top, m })),
      ].sort((a, b) => a.top - b.top);

      let para = null;
      const flush = (gapAfter) => {
        if (!para) return;
        const { lines, align } = para;
        const L = para.bounds.L;
        const first = lines[0];
        const size = Math.max(...lines.map((l) => l.size));
        const ratio = size / body;
        const words = lines.reduce((a, l) => a + l.segs.reduce((b, s) => b + s.runs.reduce((c, r) => c + r.text.split(/\s+/).length, 0), 0), 0);
        const heading = ratio >= 1.2 && words < 30 ? (ratio >= 1.7 ? 1 : ratio >= 1.35 ? 2 : 3) : 0;
        const indent = align === 'left' || align === 'both' ? Math.max(0, Math.min(...lines.slice(lines.length > 1 ? 1 : 0).map((l) => l.x)) - L) : 0;
        const firstInd = lines.length > 1 && align !== 'center' ? first.x - L - indent : 0;
        // runs: une líneas con espacio (o quita el guion de corte)
        const runs = [];
        lines.forEach((l, li) => {
          l.segs.forEach((s, si) => {
            s.runs.forEach((r, ri) => {
              const rr = { ...r };
              if (si > 0 && ri === 0) rr.text = '\t' + rr.text;
              runs.push(rr);
            });
          });
          if (li < lines.length - 1) {
            const last = runs[runs.length - 1];
            if (/[A-Za-zÀ-ÿ]-$/.test(last.text)) last.text = last.text.slice(0, -1);
            else if (!/\s$/.test(last.text)) last.text += ' ';
          }
        });
        const lineGap = lines.length > 1 ? (lines[lines.length - 1].base - first.base) / (lines.length - 1) : size * 1.2;
        const lineVal = Math.round(Math.max(240, Math.min(480, (lineGap / (size * 1.22)) * 240)));
        const after = Math.max(0, Math.min(36, (gapAfter ?? size * 0.6) - size * 0.3));
        bodyXml.push(`<w:p><w:pPr>${heading ? `<w:pStyle w:val="Heading${heading}"/>` : ''}${align !== 'left' ? `<w:jc w:val="${align}"/>` : ''}${indent > 4 || Math.abs(firstInd) > 4 ? `<w:ind w:left="${Math.round(indent * 20)}"${firstInd > 4 ? ` w:firstLine="${Math.round(firstInd * 20)}"` : firstInd < -4 ? ` w:hanging="${Math.round(-firstInd * 20)}"` : ''}/>` : ''}<w:spacing w:before="0" w:after="${Math.round(after * 20)}" w:line="${lineVal}" w:lineRule="auto"/></w:pPr>${runs.map(runXml).join('')}</w:p>`);
        para = null;
      };
      const alignOf = (l, { L, R }) => {
        const width = l.right - l.x;
        const center = (l.x + l.right) / 2;
        if (Math.abs(center - (L + R) / 2) < (R - L) * 0.05 && width < (R - L) * 0.85 && l.x > L + l.size * 2) return 'center';
        if (Math.abs(l.right - R) < l.size && l.x > L + (R - L) * 0.35) return 'right';
        return 'left';
      };
      let prevBottom = null;
      const takeLine = (l, bounds) => {
        const { R } = bounds;
        const align = alignOf(l, bounds);
        if (para) {
          const prev = para.lines[para.lines.length - 1];
          const gap = l.base - prev.base;
          const sizeChange = Math.abs(l.size - prev.size) > prev.size * 0.15;
          const shortPrev = prev.right < R - prev.size * 4 && para.align !== 'center';
          const indentJump = Math.abs(l.x - prev.x) > l.size * 1.5 && para.lines.length > 1;
          const bullet = /^([•●▪◦‣∙·\-–*]|\d+[.)])\s/.test(l.segs[0].runs[0].text);
          if (gap < 0 || gap > Math.max(l.size, prev.size) * 1.65 || sizeChange || shortPrev || indentJump || align !== para.align || bullet || l.segs.length > 1 || prev.segs.length > 1) {
            flush(Math.max(0, gap - prev.size * 1.2 + prev.size * 0.6));
          }
        }
        if (!para) para = { lines: [], align, bounds };
        para.lines.push(l);
        prevBottom = l.base + l.size * 0.25;
      };
      for (const e of els) {
        if (e.kind === 'img') {
          flush(prevBottom != null ? e.top - prevBottom : null);
          const m = e.m;
          const center = Math.abs(m.x + m.w / 2 - (L + R) / 2) < (R - L) * 0.08;
          bodyXml.push(imageXml(m.canvas, m.w, m.h, center ? 'center' : null, center ? 0 : m.x - L));
          prevBottom = m.top + m.h;
          continue;
        }
        if (e.kind === 'cols') {
          flush(prevBottom != null ? e.t.top - prevBottom : null);
          const { anchors, rows } = e.t;
          const colOf = (x) => { let k = 0; anchors.forEach((a, j) => { if (x >= a.min - 8) k = j; }); return k; };
          let bottom = prevBottom;
          anchors.forEach((a, k) => {
            const pls = [];
            for (const ri of rows) {
              const ln = pg.lines[ri];
              const segs = ln.segs.filter((sg) => colOf(sg.x) === k);
              if (!segs.length) continue;
              const right = Math.max(...segs.map((sg) => sg.x + sg.w));
              pls.push({ ...ln, segs: [{ ...segs[0], runs: segs.flatMap((sg) => sg.runs), w: right - segs[0].x }], x: segs[0].x, right });
            }
            if (!pls.length) return;
            const bounds = { L: a.min, R: Math.max(...pls.map((p) => p.right)) };
            for (const pl of pls) takeLine(pl, bounds);
            flush(null);
            bottom = Math.max(bottom ?? 0, prevBottom ?? 0);
          });
          prevBottom = bottom;
          continue;
        }
        if (e.kind === 'table') {
          flush(prevBottom != null ? e.top - prevBottom : null);
          const { anchors, rows } = e.t;
          const edges = anchors.map((a) => a.min - 3).concat([Math.max(R, ...rows.map((r) => pg.lines[r].right)) + 3]);
          const widths = edges.slice(0, -1).map((x, k) => Math.max(20, edges[k + 1] - x));
          const colOf = (x) => { let k = 0; anchors.forEach((a, j) => { if (x >= a.min - 8) k = j; }); return k; };
          const border = (s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>`;
          let x = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="${Math.round(Math.max(0, anchors[0].min - L - 3) * 20)}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${Math.round(w * 20)}"/>`).join('')}</w:tblGrid>`;
          for (const ri of rows) {
            const cells = widths.map(() => []);
            const bgs = widths.map(() => null);
            for (const s of pg.lines[ri].segs) {
              const k = colOf(s.x);
              if (s.bg && !bgs[k]) bgs[k] = s.bg;
              const c = cells[k];
              if (c.length) c[c.length - 1] = { ...c[c.length - 1], text: c[c.length - 1].text + ' ' };
              c.push(...s.runs);
            }
            const rowBg = bgs.find((b) => b && !/^f[4-9a-f]f[4-9a-f]f[4-9a-f]$/i.test(b));
            const shd = (k) => { const b = bgs[k] || rowBg; return b && !/^f[4-9a-f]f[4-9a-f]f[4-9a-f]$/i.test(b) ? `<w:shd w:val="clear" w:color="auto" w:fill="${b}"/>` : ''; };
            x += `<w:tr>${cells.map((runs, k) => `<w:tc><w:tcPr><w:tcW w:w="${Math.round(widths[k] * 20)}" w:type="dxa"/>${shd(k)}</w:tcPr><w:p><w:pPr><w:spacing w:before="20" w:after="20"/></w:pPr>${runs.map(runXml).join('')}</w:p></w:tc>`).join('')}</w:tr>`;
          }
          bodyXml.push(x + '</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>');
          prevBottom = pg.lines[rows[rows.length - 1]].base + pg.lines[rows[rows.length - 1]].size * 0.3;
          continue;
        }
        takeLine(e.l, { L, R });
      }
      flush(null);
      if (pi < pages.length - 1) bodyXml.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    });

    // paquete .docx
    progress(0.96, 'Guardando…');
    const zip = new JSZip();
    const hd = (n, sz) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    for (const m of media) {
      const jpg = m.canvas.width * m.canvas.height > 640000;
      const blob = await U.toBlob(m.canvas, jpg ? 'image/jpeg' : 'image/png', 0.9);
      m.file = `media/imagen${m.id}.${jpg ? 'jpeg' : 'png'}`;
      zip.file(`word/${m.file}`, blob);
      m.canvas.width = m.canvas.height = 0;
    }
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${media.map((m) => `<Relationship Id="rIdImg${m.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${m.file}"/>`).join('')}</Relationships>`);
    zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${xe(mainFont)}" w:hAnsi="${xe(mainFont)}" w:cs="${xe(mainFont)}" w:eastAsia="${xe(mainFont)}"/><w:sz w:val="${Math.round(body * 2)}"/><w:lang w:val="es-ES"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${hd(1, Math.round(body * 3.4))}${hd(2, Math.round(body * 2.8))}${hd(3, Math.round(body * 2.4))}<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${bodyXml.join('')}<w:sectPr><w:pgSz w:w="${Math.round(P0.w * 20)}" w:h="${Math.round(P0.h * 20)}"${P0.w > P0.h ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${Math.round(topM * 20)}" w:right="${Math.round(Math.max(28, P0.w - R) * 20)}" w:bottom="${Math.round(topM * 20)}" w:left="${Math.round(L * 20)}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  /* ════════════ Quitar el texto de un PDF (para el fondo de las diapositivas) ════════════ */
  function stripTextOps(str) {
    const n = str.length;
    const out = [];
    let operands = [];
    let i = 0;
    const ws = (c) => c === 32 || c === 10 || c === 13 || c === 9 || c === 12 || c === 0;
    const delim = (c) => c === 40 || c === 41 || c === 60 || c === 62 || c === 91 || c === 93 || c === 123 || c === 125 || c === 47 || c === 37;
    while (i < n) {
      const c = str.charCodeAt(i);
      if (ws(c)) { i++; continue; }
      if (c === 37) { while (i < n && str[i] !== '\n' && str[i] !== '\r') i++; continue; }
      const start = i;
      let isOp = false;
      if (c === 40) {
        let depth = 0;
        do {
          const ch = str[i];
          if (ch === '\\') { i += 2; continue; }
          if (ch === '(') depth++;
          else if (ch === ')') depth--;
          i++;
        } while (i < n && depth > 0);
      } else if (c === 60 && str[i + 1] === '<') i += 2;
      else if (c === 62 && str[i + 1] === '>') i += 2;
      else if (c === 60) { while (i < n && str[i] !== '>') i++; i++; }
      else if (c === 91 || c === 93 || c === 123 || c === 125) i++;
      else if (c === 47) { i++; while (i < n && !ws(str.charCodeAt(i)) && !delim(str.charCodeAt(i))) i++; }
      else {
        while (i < n && !ws(str.charCodeAt(i)) && !delim(str.charCodeAt(i))) i++;
        const tok = str.slice(start, i);
        if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(tok) && tok !== 'true' && tok !== 'false' && tok !== 'null') isOp = true;
      }
      if (i === start) i++;
      const text = str.slice(start, i);
      if (!isOp) { operands.push(text); continue; }
      if (text === 'BI') {
        const idPos = str.indexOf('ID', i);
        const re = /\sEI(?=[\s]|$)/g;
        re.lastIndex = idPos + 3;
        const m = idPos >= 0 ? re.exec(str) : null;
        const end = m ? m.index + m[0].length : n;
        out.push(...operands, str.slice(start, end));
        operands = [];
        i = end;
        continue;
      }
      if (text === 'Tj' || text === 'TJ') { operands = []; continue; }
      if (text === "'" || text === '"') { operands = []; out.push('T*'); continue; }
      out.push(...operands, text);
      operands = [];
    }
    out.push(...operands);
    return out.join(' ');
  }

  async function stripText(bytes) {
    const { PDFDocument, PDFName, PDFArray, PDFRawStream, PDFDict, PDFRef, decodePDFRawStream } = PDFLib;
    const doc = await PDFDocument.load(bytes);
    const ctx = doc.context;
    const latin1 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return s; };
    const toU8 = (s) => { const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i) & 255; return u; };
    const decode = (st) => (st instanceof PDFRawStream ? decodePDFRawStream(st).decode() : st.getContents());
    const seen = new Set();
    const SKIP = new Set(['/Filter', '/DecodeParms', '/Length']);
    const doResources = (res) => {
      const xo = res instanceof PDFDict ? res.lookup(PDFName.of('XObject')) : null;
      if (!(xo instanceof PDFDict)) return;
      for (const [, ref] of xo.entries()) {
        if (!(ref instanceof PDFRef) || seen.has(ref.toString())) continue;
        seen.add(ref.toString());
        const obj = ctx.lookup(ref);
        if (!obj || !obj.dict || obj.dict.lookup(PDFName.of('Subtype')) !== PDFName.of('Form')) continue;
        try {
          const ns = ctx.flateStream(toU8(stripTextOps(latin1(decode(obj)))));
          for (const [k, v] of obj.dict.entries()) if (!SKIP.has(k.toString())) ns.dict.set(k, v);
          ctx.assign(ref, ns);
          doResources(ns.dict.lookup(PDFName.of('Resources')));
        } catch (_) { /* dejamos el formulario como está */ }
      }
    };
    for (const page of doc.getPages()) {
      const c = page.node.Contents();
      if (!c) continue;
      const streams = c instanceof PDFArray ? c.asArray().map((r) => ctx.lookup(r)) : [c];
      const data = streams.map((s) => latin1(decode(s))).join('\n');
      const ns = ctx.flateStream(toU8(stripTextOps(data)));
      page.node.set(PDFName.of('Contents'), ctx.register(ns));
      doResources(page.node.Resources());
    }
    return doc.save();
  }

  /* ════════════ PDF → PowerPoint ════════════ */
  async function toPptx(pdf, bytes, mode, progress) {
    const ownZip = window.JSZip;
    await U.loadScript('vendor/pptxgen.bundle.js');
    window.JSZip = ownZip;
    const first = (await pdf.getPage(1)).getViewport({ scale: 1 });
    const pptx = new PptxGenJS();
    const SW = first.width / 72, SH = first.height / 72;
    pptx.defineLayout({ name: 'PDF', width: +SW.toFixed(3), height: +SH.toFixed(3) });
    pptx.layout = 'PDF';

    let bgPdf = null;
    let note = '';
    if (mode !== 'image') {
      try {
        progress(0.02, 'Separando el texto del fondo…');
        bgPdf = await pdfjsLib.getDocument({ data: await stripText(bytes) }).promise;
      } catch (e) {
        console.warn(e);
        note = 'No se pudo separar el texto del fondo en este PDF: las diapositivas se han creado como imágenes.';
      }
    }
    const SC = 2;
    for (let i = 1; i <= pdf.numPages; i++) {
      progress(0.05 + (0.9 * (i - 1)) / pdf.numPages, `Creando diapositiva ${i} de ${pdf.numPages}…`);
      const page = await pdf.getPage(i);
      const vp = page.getViewport({ scale: 1 });
      const k = Math.min(SW / (vp.width / 72), SH / (vp.height / 72));
      const ox = (SW - (vp.width / 72) * k) / 2, oy = (SH - (vp.height / 72) * k) / 2;
      const slide = pptx.addSlide();
      const orig = await renderCanvas(pdf, i, SC);
      if (!bgPdf) {
        slide.addImage({ data: orig.toDataURL('image/jpeg', 0.88), x: ox, y: oy, w: (vp.width / 72) * k, h: (vp.height / 72) * k });
        orig.width = orig.height = 0;
        continue;
      }
      const bgc = await renderCanvas(bgPdf, i, SC);
      slide.addImage({ data: bgc.toDataURL('image/jpeg', 0.9), x: ox, y: oy, w: (vp.width / 72) * k, h: (vp.height / 72) * k });
      const an = await analyze(page);
      for (const l of an.lines) {
        for (const s of l.segs) {
          const runs = s.runs.map((r) => ({
            text: r.text,
            options: {
              bold: r.bold, italic: r.italic, fontFace: r.family, fontSize: Math.max(1, +(r.size * k).toFixed(1)),
              color: sampleColor(orig, bgc, SC, r.x, l.base - r.size * 0.8, r.w, r.size * 0.95),
            },
          }));
          const top = (l.base - s.size * 0.93) / 72;
          slide.addText(runs, {
            x: ox + (s.x / 72) * k, y: oy + top * k, w: ((s.w + s.size * 0.6) / 72) * k, h: ((s.size * 1.25) / 72) * k,
            margin: 0, valign: 'top', wrap: false, fit: 'none', paraSpaceBefore: 0, paraSpaceAfter: 0, lineSpacingMultiple: 1,
          });
        }
      }
      orig.width = orig.height = bgc.width = bgc.height = 0;
      page.cleanup();
    }
    if (bgPdf) bgPdf.destroy();
    progress(0.97, 'Guardando la presentación…');
    return { blob: await pptx.write({ outputType: 'blob' }), note };
  }

  return { toDocx, toPptx, stripText, analyze };
})();
