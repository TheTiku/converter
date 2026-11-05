'use strict';

/**
 * Conversor propio de Word (.docx) a PDF.
 * Interpreta estilos, listas, tablas, secciones (tamaño de página, márgenes, columnas),
 * encabezados y pies de página con campos PÁGINA/NUMPÁGINAS, imágenes en línea y flotantes,
 * cuadros de texto, gráficos y notas al pie.
 */
const Docx = (() => {
  const TW = 20;
  const EMU = 12700;
  const HIGHLIGHT = {
    yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080',
    darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080',
    lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
  };
  const { kid, kids, attr, num, desc, all } = X;

  /* ── utilidades de propiedades ── */
  const merge = (...objs) => {
    const o = {};
    for (const src of objs) if (src) for (const k in src) if (src[k] !== undefined) o[k] = src[k];
    return o;
  };
  const toggle = (el, n) => {
    const e = kid(el, n);
    if (!e) return undefined;
    const v = attr(e, 'val');
    return !(v === '0' || v === 'false' || v === 'off');
  };
  const symChar = (ch) => {
    const c = ch.charCodeAt(0);
    if (c >= 0xf000 && c <= 0xf0ff) {
      const m = { 0xf0b7: '•', 0xf0a7: '▪', 0xf0d8: '►', 0xf0fc: '✓', 0xf076: '◆', 0xf0a8: '○', 0xf06f: '○', 0xf0e0: '→', 0xf02d: '-' };
      return m[c] || '•';
    }
    return ch;
  };

  async function toPdf(bytes, progress) {
    let zip;
    try {
      zip = await JSZip.loadAsync(bytes);
    } catch (_) {
      throw new Error('No se pudo leer el documento. Asegúrate de que es un .docx (los .doc antiguos no son compatibles).');
    }
    const read = async (p) => {
      const f = p && zip.file(p);
      return f ? X.parse(await f.async('string')) : null;
    };
    let docPath = 'word/document.xml';
    try {
      const rr = X.parse(await zip.file('_rels/.rels').async('string'));
      const main = all(rr, 'Relationship').find((r) => /officeDocument$/.test(r.getAttribute('Type')));
      if (main) docPath = main.getAttribute('Target').replace(/^\//, '');
    } catch (_) { /* ruta por defecto */ }
    const docXml = await read(docPath);
    if (!docXml) throw new Error('No se pudo leer el documento. Asegúrate de que es un .docx válido.');
    const rels = await ooxmlRels(zip, docPath);
    const relOf = (t) => Object.values(rels).find((r) => r.type.endsWith('/' + t))?.target;
    const stylesDoc = await read(relOf('styles'));
    const numDoc = await read(relOf('numbering'));
    const themeDoc = await read(relOf('theme'));
    const footDoc = await read(relOf('footnotes'));
    const theme = parseTheme(themeDoc);

    const pdf = await PDFLib.PDFDocument.create();
    const bank = FontBank.forDoc(pdf);
    const fontNames = new Set([theme.minor, theme.major]);
    for (const d of [stylesDoc, docXml]) for (const f of all(d, 'rFonts')) { const a = attr(f, 'ascii'); if (a) fontNames.add(a); }
    progress(0.03, 'Cargando las fuentes…');
    await bank.prepare([...fontNames]);

    /* ── estilos ── */
    const colorAttr = (el) => {
      if (!el) return undefined;
      const tc = attr(el, 'themeColor');
      if (tc) {
        let hex = theme.scheme[tc] || theme.scheme[tc.replace(/^(text|background)(\d)$/, (m, a, n) => (a === 'text' ? 'dk' : 'lt') + n)];
        const shade = attr(el, 'themeShade'), tint = attr(el, 'themeTint');
        if (hex && (shade || tint)) {
          let [r, g, b] = hexRgb(hex);
          if (shade) { const s = parseInt(shade, 16) / 255; [r, g, b] = [r, g, b].map((v) => v * s); }
          if (tint) { const t = parseInt(tint, 16) / 255; [r, g, b] = [r, g, b].map((v) => v * t + (1 - t)); }
          hex = [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
        }
        if (hex) return hex;
      }
      const v = attr(el, 'val') || attr(el, 'fill');
      return v && v !== 'auto' ? v : v === 'auto' ? 'auto' : undefined;
    };
    /** Color de relleno de un w:shd (fill / themeFill); null si es automático. */
    const shdFill = (s) => {
      const tf = attr(s, 'themeFill');
      if (tf && theme.scheme[tf]) return theme.scheme[tf];
      const f = attr(s, 'fill');
      return f && f !== 'auto' && /^[0-9a-f]{6}$/i.test(f) ? f : null;
    };
    const parseRPr = (el) => {
      if (!el) return {};
      const o = {
        bold: toggle(el, 'b'), italic: toggle(el, 'i'), strike: toggle(el, 'strike') || toggle(el, 'dstrike') || undefined,
        caps: toggle(el, 'caps'), smallCaps: toggle(el, 'smallCaps'), vanish: toggle(el, 'vanish'),
      };
      const u = kid(el, 'u');
      if (u) o.underline = attr(u, 'val') !== 'none';
      const sz = kid(el, 'sz');
      if (sz) o.size = num(sz, 'val', 22) / 2;
      const c = kid(el, 'color');
      if (c) o.color = colorAttr(c);
      const f = kid(el, 'rFonts');
      if (f) {
        const th = attr(f, 'asciiTheme') || attr(f, 'hAnsiTheme');
        o.font = attr(f, 'ascii') || attr(f, 'hAnsi') || (th ? (/major/.test(th) ? theme.major : theme.minor) : undefined);
      }
      const hl = kid(el, 'highlight');
      if (hl && HIGHLIGHT[attr(hl, 'val')]) o.highlight = HIGHLIGHT[attr(hl, 'val')];
      const shd = kid(el, 'shd');
      if (shd && !o.highlight) { const f2 = shdFill(shd); if (f2) o.highlight = f2; }
      const va = kid(el, 'vertAlign');
      if (va) o.vert = attr(va, 'val');
      const rs = kid(el, 'rStyle');
      if (rs) o.rStyle = attr(rs, 'val');
      return o;
    };
    const parsePPr = (el) => {
      if (!el) return {};
      const o = {
        keepNext: toggle(el, 'keepNext'), pageBreakBefore: toggle(el, 'pageBreakBefore'), contextual: toggle(el, 'contextualSpacing'),
      };
      const jc = attr(kid(el, 'jc'), 'val');
      if (jc) o.jc = { start: 'left', end: 'right', distribute: 'both', justify: 'both' }[jc] || jc;
      const sp = kid(el, 'spacing');
      if (sp) {
        if (attr(sp, 'before') != null) o.before = num(sp, 'before') / TW;
        if (attr(sp, 'after') != null) o.after = num(sp, 'after') / TW;
        if (attr(sp, 'beforeAutospacing') === '1') o.before = 14;
        if (attr(sp, 'afterAutospacing') === '1') o.after = 14;
        if (attr(sp, 'line') != null) { o.line = num(sp, 'line'); o.lineRule = attr(sp, 'lineRule') || 'auto'; }
      }
      const ind = kid(el, 'ind');
      if (ind) {
        const l = attr(ind, 'left') ?? attr(ind, 'start');
        const r = attr(ind, 'right') ?? attr(ind, 'end');
        if (l != null) o.left = +l / TW;
        if (r != null) o.right = +r / TW;
        if (attr(ind, 'firstLine') != null) { o.first = num(ind, 'firstLine') / TW; }
        if (attr(ind, 'hanging') != null) { o.first = -num(ind, 'hanging') / TW; }
      }
      const np = kid(el, 'numPr');
      if (np) {
        const id = attr(kid(np, 'numId'), 'val');
        if (id != null) o.numId = id;
        const lv = attr(kid(np, 'ilvl'), 'val');
        if (lv != null) o.ilvl = +lv;
      }
      const ps = kid(el, 'pStyle');
      if (ps) o.pStyle = attr(ps, 'val');
      const tabs = kid(el, 'tabs');
      if (tabs) o.tabs = kids(tabs, 'tab').map((t) => ({ pos: num(t, 'pos') / TW, align: attr(t, 'val'), leader: attr(t, 'leader') }));
      const bdr = kid(el, 'pBdr');
      if (bdr) {
        o.borders = {};
        for (const side of ['top', 'bottom', 'left', 'right']) {
          const b = kid(bdr, side);
          if (b && !/^(none|nil)$/.test(attr(b, 'val'))) o.borders[side] = { w: Math.max(0.4, num(b, 'sz', 4) / 8), color: colorAttr(kid(bdr, side)) || '000000', space: num(b, 'space') };
        }
      }
      const shd = kid(el, 'shd');
      if (shd) { const f = shdFill(shd); if (f) o.shd = f; }
      return o;
    };

    const styles = {};
    let defP = null, defT = null, defR = null;
    const sRoot = stylesDoc?.documentElement;
    const rPrDef = parseRPr(X.path(desc(sRoot, 'docDefaults'), 'rPrDefault', 'rPr'));
    const pPrDef = parsePPr(X.path(desc(sRoot, 'docDefaults'), 'pPrDefault', 'pPr'));
    for (const s of kids(sRoot, 'style')) {
      const id = attr(s, 'styleId');
      const type = attr(s, 'type');
      const cond = {};
      for (const c of kids(s, 'tblStylePr')) cond[attr(c, 'type')] = { rPr: parseRPr(kid(c, 'rPr')), pPr: parsePPr(kid(c, 'pPr')), tcPr: kid(c, 'tcPr'), tblPr: kid(c, 'tblPr') };
      styles[id] = { type, basedOn: attr(kid(s, 'basedOn'), 'val'), pPr: parsePPr(kid(s, 'pPr')), rPr: parseRPr(kid(s, 'rPr')), tblPr: kid(s, 'tblPr'), tcPr: kid(s, 'tcPr'), cond, name: attr(kid(s, 'name'), 'val') };
      if (attr(s, 'default') === '1') {
        if (type === 'paragraph') defP = id;
        if (type === 'table') defT = id;
        if (type === 'character') defR = id;
      }
    }
    const chain = (id) => {
      const out = [];
      const seen = new Set();
      while (id && styles[id] && !seen.has(id)) { seen.add(id); out.unshift(styles[id]); id = styles[id].basedOn; }
      return out;
    };
    const styleP = (id) => merge(...chain(id).map((s) => s.pPr));
    const styleR = (id) => merge(...chain(id).map((s) => s.rPr));

    /* ── numeración ── */
    const absNums = {}, nums = {}, counters = {};
    for (const a of kids(numDoc?.documentElement, 'abstractNum')) {
      const levels = {};
      for (const l of kids(a, 'lvl')) {
        levels[num(l, 'ilvl')] = {
          start: num(kid(l, 'start'), 'val', 1), fmt: attr(kid(l, 'numFmt'), 'val') || 'decimal', text: attr(kid(l, 'lvlText'), 'val') ?? '',
          pPr: parsePPr(kid(l, 'pPr')), rPr: parseRPr(kid(l, 'rPr')), suff: attr(kid(l, 'suff'), 'val') || 'tab',
        };
      }
      absNums[attr(a, 'abstractNumId')] = levels;
    }
    for (const n of kids(numDoc?.documentElement, 'num')) {
      const ov = {};
      for (const o of kids(n, 'lvlOverride')) { const st = kid(o, 'startOverride'); if (st) ov[num(o, 'ilvl')] = num(st, 'val', 1); }
      nums[attr(n, 'numId')] = { abs: attr(kid(n, 'abstractNumId'), 'val'), ov };
    }
    const roman = (n) => {
      const t = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
      let s = '';
      for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
      return s;
    };
    const fmtNum = (n, f) => {
      switch (f) {
        case 'lowerLetter': return String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
        case 'upperLetter': return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
        case 'lowerRoman': return roman(n);
        case 'upperRoman': return roman(n).toUpperCase();
        case 'decimalZero': return String(n).padStart(2, '0');
        case 'none': return '';
        default: return String(n);
      }
    };
    const numLevel = (numId, ilvl) => {
      const n = nums[numId];
      return n && absNums[n.abs] ? absNums[n.abs][ilvl] || null : null;
    };
    const nextLabel = (numId, ilvl) => {
      const n = nums[numId];
      const levels = n && absNums[n.abs];
      if (!levels || !levels[ilvl]) return null;
      const key = n.abs;
      const c = (counters[key] = counters[key] || {});
      for (let l = 0; l <= ilvl; l++) if (c[l] == null) c[l] = (n.ov[l] ?? levels[l]?.start ?? 1) - (l === ilvl ? 1 : 0);
      c[ilvl]++;
      for (const l in c) if (+l > ilvl) delete c[l];
      const lv = levels[ilvl];
      if (lv.fmt === 'bullet') return [...lv.text].map(symChar).join('') || '•';
      return lv.text.replace(/%(\d)/g, (m, d) => fmtNum(c[d - 1] ?? levels[d - 1]?.start ?? 1, levels[d - 1]?.fmt || 'decimal'));
    };

    /* ── imágenes y gráficos ── */
    const imgCache = {};
    let skipped = 0;
    const loadImg = async (target) => {
      if (!target) return null;
      if (!(target in imgCache)) {
        const f = zip.file(target);
        imgCache[target] = f ? embedAnyImage(pdf, await f.async('uint8array'), mimeOf(target)).then((e) => e && e.img) : Promise.resolve(null);
      }
      const img = await imgCache[target];
      if (!img) skipped++;
      return img;
    };
    const pending = [];

    /* ── fuentes de los runs ── */
    const resolveRun = async (text, rp) => {
      const symbolFont = /symbol|wingdings|webdings/i.test(rp.font || '');
      let t = symbolFont ? [...text].map((ch) => symChar(String.fromCharCode(0xf000 | (ch.charCodeAt(0) & 0xff)))).join('') : text;
      let size = rp.size || 10;
      if (rp.caps) t = t.toUpperCase();
      if (rp.smallCaps) { t = t.toUpperCase(); size *= 0.8; }
      let rise = 0;
      if (rp.vert === 'superscript') { rise = size * 0.33; size *= 0.65; }
      else if (rp.vert === 'subscript') { rise = -size * 0.14; size *= 0.65; }
      const font = await bank.get(symbolFont ? theme.minor : rp.font || theme.minor, !!rp.bold, !!rp.italic);
      return {
        text: t, font, size, rise,
        color: rp.color && rp.color !== 'auto' ? toRgb(rp.color) : null,
        underline: !!rp.underline, strike: !!rp.strike,
        highlight: rp.highlight ? toRgb(rp.highlight) : null,
      };
    };

    /* ── parseo de bloques ── */
    let fnCount = 0;
    const fnRefs = [];
    let fnCurrent = null; // número de la nota que se está leyendo

    async function parseDrawing(d, relsMap, partZipPath) {
      const inline = kid(d, 'inline');
      const anchor = kid(d, 'anchor');
      const el = inline || anchor;
      if (!el) return null;
      const ext = kid(el, 'extent');
      const w = num(ext, 'cx') / EMU, h = num(ext, 'cy') / EMU;
      const gd = desc(el, 'graphicData');
      const uri = gd?.getAttribute('uri') || '';
      const item = { w, h };
      if (/picture/.test(uri)) {
        const blip = desc(gd, 'blip');
        const rid = blip && (attr(blip, 'embed') || attr(blip, 'link'));
        item.img = await loadImg(relsMap[rid]?.target);
        if (!item.img) return null;
      } else if (/chart/.test(uri)) {
        const c = desc(gd, 'chart');
        const target = relsMap[attr(c, 'id')]?.target;
        item.chart = target ? await read(target) : null;
        if (!item.chart) return null;
      } else if (/wordprocessingShape/.test(uri)) {
        const wsp = desc(gd, 'wsp');
        const spPr = kid(wsp, 'spPr');
        item.fill = ooxmlColor(kid(spPr, 'solidFill'), theme.scheme)?.hex;
        item.line = ooxmlColor(kid(kid(spPr, 'ln'), 'solidFill'), theme.scheme)?.hex;
        const tb = desc(wsp, 'txbxContent');
        if (tb) item.blocks = await parseBlocks(tb, relsMap, {});
        const bp = kid(wsp, 'bodyPr');
        item.ins = [num(bp, 'lIns', 91440) / EMU, num(bp, 'tIns', 45720) / EMU, num(bp, 'rIns', 91440) / EMU, num(bp, 'bIns', 45720) / EMU];
        if (!item.fill && !item.line && !item.blocks) return null;
      } else {
        // grupos u otros objetos: intentamos sacar la primera imagen que contengan
        const blip = desc(gd, 'blip');
        if (!blip) return null;
        item.img = await loadImg(relsMap[attr(blip, 'embed')]?.target);
        if (!item.img) return null;
      }
      if (inline) return { inline: item };
      const pos = (n) => {
        const p = kid(anchor, n);
        return { rel: attr(p, 'relativeFrom') || 'column', off: kid(p, 'posOffset') ? +kid(p, 'posOffset').textContent / EMU : null, align: kid(p, 'align')?.textContent || null };
      };
      const wrapEl = kids(anchor).find((c) => /^wrap/.test(c.localName));
      const dist = (n) => num(anchor, n) / EMU;
      return { anchor: {
        ...item, posH: pos('positionH'), posV: pos('positionV'), wrap: wrapEl ? wrapEl.localName : 'wrapNone', wrapText: attr(wrapEl, 'wrapText') || 'bothSides',
        behind: attr(anchor, 'behindDoc') === '1', dist: { l: dist('distL'), r: dist('distR'), t: dist('distT'), b: dist('distB') },
      } };
    }

    async function parseRuns(container, baseR, relsMap, out, fld, linkCtx) {
      for (const c of container.children) {
        const n = c.localName;
        if (n === 'r') await parseRun(c, baseR, relsMap, out, fld, linkCtx);
        else if (n === 'hyperlink' || n === 'smartTag' || n === 'ins' || n === 'customXml' || n === 'moveTo' || n === 'dir' || n === 'bdo') await parseRuns(c, baseR, relsMap, out, fld, n === 'hyperlink' || linkCtx);
        else if (n === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) await parseRuns(sc, baseR, relsMap, out, fld, linkCtx); }
        else if (n === 'fldSimple') {
          const instr = (attr(c, 'instr') || '').trim().toUpperCase();
          const f = /^(PAGE|NUMPAGES|SECTIONPAGES)\b/.exec(instr);
          if (f) {
            const r = kid(c, 'r');
            const rp = merge(baseR, parseRPr(kid(r, 'rPr')));
            out.push({ ...(await resolveRun('0', rp)), field: f[1] });
          } else await parseRuns(c, baseR, relsMap, out, fld, linkCtx);
        } else if (n === 'oMath' || n === 'oMathPara') {
          const t = all(c, 't').map((x) => x.textContent).join('');
          if (t) out.push(await resolveRun(t, merge(baseR, { italic: true })));
        }
      }
    }

    async function parseRun(r, baseR, relsMap, out, fld, linkCtx) {
      const direct = parseRPr(kid(r, 'rPr'));
      const rp = merge(baseR, styleR(direct.rStyle || (linkCtx && !direct.rStyle && styles.Hyperlink ? 'Hyperlink' : null)), direct);
      if (rp.vanish) return;
      for (const k of r.children) {
        const n = k.localName;
        if (n === 'fldChar') {
          const t = attr(k, 'fldCharType');
          if (t === 'begin') { fld.state = 'instr'; fld.instr = ''; fld.dyn = null; }
          else if (t === 'separate') {
            fld.state = 'result';
            const f = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(fld.instr || '');
            if (f) { fld.dyn = f[1].toUpperCase(); out.push({ ...(await resolveRun('0', rp)), field: fld.dyn }); }
          } else if (t === 'end') {
            if (fld.state === 'instr') {
              const f = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(fld.instr || '');
              if (f) out.push({ ...(await resolveRun('0', rp)), field: f[1].toUpperCase() });
            }
            fld.state = null; fld.dyn = null;
          }
          continue;
        }
        if (n === 'instrText') { if (fld.state === 'instr') fld.instr += k.textContent; continue; }
        if (fld.state === 'instr' || (fld.state === 'result' && fld.dyn)) continue;
        if (n === 't' || n === 'delText') { if (n === 't' && k.textContent) out.push(await resolveRun(k.textContent, rp)); }
        else if (n === 'tab' || n === 'ptab') out.push(await resolveRun('\t', rp));
        else if (n === 'br') {
          const t = attr(k, 'type');
          if (t === 'page' || t === 'column') out.push({ brk: t });
          else out.push(await resolveRun('\n', rp));
        } else if (n === 'cr') out.push(await resolveRun('\n', rp));
        else if (n === 'noBreakHyphen') out.push(await resolveRun('-', rp));
        else if (n === 'sym') {
          const ch = String.fromCharCode(parseInt(attr(k, 'char') || '2022', 16));
          out.push(await resolveRun(symChar(ch.charCodeAt(0) < 0xf000 ? String.fromCharCode(0xf000 | ch.charCodeAt(0)) : ch), merge(rp, { font: theme.minor })));
        } else if (n === 'drawing') {
          const d = await parseDrawing(k, relsMap);
          if (d) out.push(d);
        } else if (n === 'pict' || n === 'object') {
          const im = desc(k, 'imagedata');
          const shape = desc(k, 'shape');
          if (im) {
            const img = await loadImg(relsMap[attr(im, 'id')]?.target);
            const st = (shape?.getAttribute('style') || '').toLowerCase();
            const dim = (p) => { const m = new RegExp(p + ':\\s*([\\d.]+)(pt|in|px|cm|mm)?').exec(st); if (!m) return null; const v = +m[1]; return { pt: v, in: v * 72, px: v * 0.75, cm: v * 28.35, mm: v * 2.835 }[m[2] || 'pt']; };
            if (img) out.push({ inline: { img, w: dim('width') || img.width * 0.75, h: dim('height') || img.height * 0.75 } });
          }
        } else if (n === 'footnoteReference') {
          const id = attr(k, 'id');
          fnRefs.push({ id, n: ++fnCount });
          out.push({ ...(await resolveRun(String(fnCount), merge(rp, { vert: 'superscript' }))), fn: id });
        } else if (n === 'footnoteRef' && fnCurrent != null) {
          out.push(await resolveRun(String(fnCurrent), merge(rp, { vert: 'superscript' })));
        }
      }
    }

    async function parsePara(p, relsMap, ts) {
      const direct = parsePPr(kid(p, 'pPr'));
      const sid = direct.pStyle || defP;
      let pp = merge(pPrDef, ts?.pPr, styleP(sid), direct);
      const baseR = merge(rPrDef, styleR(defR), ts?.rPr, styleR(sid), ts?.condR);
      const markR = merge(baseR, parseRPr(X.path(p, 'pPr', 'rPr')));
      let label = null;
      if (pp.numId && pp.numId !== '0') {
        const ilvl = pp.ilvl || 0;
        const lv = numLevel(pp.numId, ilvl);
        if (lv) {
          pp = merge(pPrDef, ts?.pPr, styleP(sid), lv.pPr, direct);
          const text = nextLabel(pp.numId, ilvl);
          if (text) label = { ...(await resolveRun(text, merge(markR, lv.rPr, { underline: false }))), suff: lv.suff };
        }
      }
      const runs = [];
      await parseRuns(p, baseR, relsMap, runs, {}, false);
      const markFont = await resolveRun(' ', markR);
      return { t: 'p', pp, runs, label, mark: markFont, style: sid };
    }

    async function parseTable(tbl, relsMap) {
      const tp = kid(tbl, 'tblPr');
      const sid = attr(kid(tp, 'tblStyle'), 'val') || defT;
      const sch = chain(sid);
      const look = kid(tp, 'tblLook');
      const lookVal = parseInt(attr(look, 'val') || '0', 16);
      const flag = (name, bit) => (attr(look, name) != null ? attr(look, name) === '1' : !!(lookVal & bit));
      const lk = { firstRow: flag('firstRow', 0x20), lastRow: flag('lastRow', 0x40), firstCol: flag('firstColumn', 0x80), lastCol: flag('lastColumn', 0x100), noHBand: flag('noHBand', 0x200), noVBand: flag('noVBand', 0x400) };
      if (!look) { lk.firstRow = true; lk.noHBand = false; }
      const cond = merge(...sch.map((s) => s.cond));
      const styleTblPr = sch.map((s) => s.tblPr).filter(Boolean);
      const styleTcPr = sch.map((s) => s.tcPr).filter(Boolean);
      const borderSet = (el) => {
        const o = {};
        if (!el) return o;
        for (const b of el.children) {
          const side = { start: 'left', end: 'right' }[b.localName] || b.localName;
          const v = attr(b, 'val');
          o[side] = !v || /^(none|nil)$/.test(v) ? null : { w: Math.max(0.25, num(b, 'sz', 4) / 8), color: colorAttr(b) || '000000' };
        }
        return o;
      };
      const tblBorders = merge(...styleTblPr.map((e) => borderSet(kid(e, 'tblBorders'))), borderSet(kid(tp, 'tblBorders')));
      const marEl = [...styleTblPr, tp].map((e) => kid(e, 'tblCellMar'));
      const mar = { left: 5.4, right: 5.4, top: 0, bottom: 0 };
      for (const m of marEl) if (m) for (const s of m.children) { const side = { start: 'left', end: 'right' }[s.localName] || s.localName; mar[side] = num(s, 'w') / TW; }
      const jc = attr(kid(tp, 'jc'), 'val') || attr(desc(styleTblPr[styleTblPr.length - 1], 'jc'), 'val');
      const ind = num(kid(tp, 'tblInd'), 'w') / TW;
      const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => num(g, 'w') / TW);
      const trs = kids(tbl, 'tr');
      const rows = [];
      for (let ri = 0; ri < trs.length; ri++) {
        const tr = trs[ri];
        const trPr = kid(tr, 'trPr');
        const isFirst = ri === 0 && lk.firstRow, isLast = ri === trs.length - 1 && lk.lastRow;
        const bandIdx = ri - (lk.firstRow ? 1 : 0);
        const condKeys = [];
        if (!lk.noHBand && !isFirst && !isLast) condKeys.push(bandIdx % 2 === 0 ? 'band1Horz' : 'band2Horz');
        if (isFirst) condKeys.push('firstRow');
        if (isLast) condKeys.push('lastRow');
        const cells = [];
        let gi = num(kid(trPr, 'gridBefore'), 'val', 0);
        const tcs = kids(tr).filter((c) => c.localName === 'tc' || c.localName === 'sdt');
        for (let ci = 0; ci < tcs.length; ci++) {
          const tc = tcs[ci].localName === 'sdt' ? desc(tcs[ci], 'tc') : tcs[ci];
          if (!tc) continue;
          const tcPr = kid(tc, 'tcPr');
          const span = num(kid(tcPr, 'gridSpan'), 'val', 1);
          const vm = kid(tcPr, 'vMerge');
          const cellKeys = condKeys.slice();
          if (lk.firstCol && ci === 0) cellKeys.push('firstCol');
          if (lk.lastCol && ci === tcs.length - 1) cellKeys.push('lastCol');
          const cc = cellKeys.map((k) => cond[k]).filter(Boolean);
          const condR = merge(...cc.map((c) => c.rPr));
          const condP = merge(...cc.map((c) => c.pPr));
          const shdOf = (e) => { const s = kid(e, 'shd'); return s ? shdFill(s) : undefined; };
          let shd = merge({ v: shdOf(styleTcPr[styleTcPr.length - 1]) }, ...cc.map((c) => ({ v: shdOf(c.tcPr) })), { v: shdOf(tcPr) }).v;
          const borders = merge(...cc.map((c) => borderSet(kid(c.tcPr, 'tcBorders'))), borderSet(kid(tcPr, 'tcBorders')));
          const ts = { pPr: merge(styleP(sid), condP), rPr: styleR(sid), condR };
          const blocks = await parseBlocks(tc, relsMap, ts);
          cells.push({
            gi, span, blocks, shd, borders,
            vMerge: vm ? (attr(vm, 'val') === 'restart' ? 'restart' : 'continue') : null,
            vAlign: attr(kid(tcPr, 'vAlign'), 'val') || 'top',
          });
          gi += span;
        }
        const trh = kid(trPr, 'trHeight');
        rows.push({ cells, header: toggle(trPr, 'tblHeader'), h: num(trh, 'val') / TW, hRule: attr(trh, 'hRule') || 'atLeast' });
      }
      return { t: 'tbl', grid, rows, borders: tblBorders, mar, jc, ind };
    }

    async function parseBlocks(container, relsMap, ts) {
      const out = [];
      for (const c of container.children) {
        const n = c.localName;
        if (n === 'p') out.push(await parsePara(c, relsMap, ts));
        else if (n === 'tbl') out.push(await parseTable(c, relsMap));
        else if (n === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) out.push(...(await parseBlocks(sc, relsMap, ts))); }
        else if (n === 'customXml' || n === 'ins') out.push(...(await parseBlocks(c, relsMap, ts)));
      }
      return out;
    }

    /* ── secciones ── */
    const parseSect = (s, prev) => {
      const pg = kid(s, 'pgSz'), mg = kid(s, 'pgMar'), cols = kid(s, 'cols');
      const refs = (name) => {
        const o = {};
        for (const r of kids(s, name)) o[attr(r, 'type') || 'default'] = attr(r, 'id');
        return o;
      };
      return {
        w: num(pg, 'w', 11906) / TW, h: num(pg, 'h', 16838) / TW,
        top: Math.abs(num(mg, 'top', 1440)) / TW, bottom: Math.abs(num(mg, 'bottom', 1440)) / TW,
        left: num(mg, 'left', 1440) / TW, right: num(mg, 'right', 1440) / TW,
        headerDist: num(mg, 'header', 720) / TW, footerDist: num(mg, 'footer', 720) / TW,
        cols: Math.max(1, num(cols, 'num', 1)), colSpace: num(cols, 'space', 720) / TW,
        type: attr(kid(s, 'type'), 'val') || 'nextPage',
        titlePg: toggle(s, 'titlePg') || false,
        headers: { ...(prev?.headers || {}), ...refs('headerReference') },
        footers: { ...(prev?.footers || {}), ...refs('footerReference') },
        pgStart: attr(kid(s, 'pgNumType'), 'start') != null ? num(kid(s, 'pgNumType'), 'start') : null,
      };
    };

    progress(0.08, 'Leyendo el documento…');
    const body = desc(docXml, 'body');
    // división en secciones
    const secEls = [];
    for (const c of body.children) {
      if (c.localName === 'p') { const s = X.path(c, 'pPr', 'sectPr'); if (s) secEls.push(s); }
    }
    const finalSect = kid(body, 'sectPr');
    if (finalSect) secEls.push(finalSect);
    const sections = [];
    for (const s of secEls) sections.push(parseSect(s, sections[sections.length - 1]));
    if (!sections.length) sections.push(parseSect(null, null));

    const blocks = [];
    let secIdx = 0;
    const kidsOfBody = [...body.children];
    for (let i = 0; i < kidsOfBody.length; i++) {
      const c = kidsOfBody[i];
      if (i % 20 === 0) progress(0.08 + (0.5 * i) / kidsOfBody.length, 'Leyendo el documento…');
      let parsed = [];
      try {
        if (c.localName === 'p') parsed = [await parsePara(c, rels, null)];
        else if (c.localName === 'tbl') parsed = [await parseTable(c, rels)];
        else if (c.localName === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) parsed = await parseBlocks(sc, rels, null); }
      } catch (e) {
        console.warn('Bloque omitido', e);
      }
      for (const b of parsed) { b.sec = Math.min(secIdx, sections.length - 1); blocks.push(b); }
      if (c.localName === 'p' && X.path(c, 'pPr', 'sectPr')) secIdx++;
    }
    // notas al pie: se leen aparte y se colocan al pie de la página donde aparece su llamada
    const footnotes = {};
    if (fnRefs.length && footDoc) {
      const fnRels = await ooxmlRels(zip, relOf('footnotes'));
      const notes = {};
      for (const f of kids(footDoc.documentElement, 'footnote')) notes[attr(f, 'id')] = f;
      for (const ref of fnRefs) {
        const el = notes[ref.id];
        if (!el) continue;
        fnCurrent = ref.n;
        const nb = await parseBlocks(el, fnRels, null);
        fnCurrent = null;
        const first = nb.find((b) => b.t === 'p');
        if (first && !first.runs.some((r) => r.text === String(ref.n))) first.runs.unshift(await resolveRun(`${ref.n} `, merge(rPrDef, { size: 8, vert: 'superscript' })));
        footnotes[ref.id] = nb;
      }
    }

    /* ── maquetación ── */
    const S = { page: null, sec: null, secIdx: -1, col: 0, y: 0, colTop: 0, pages: [] };
    const colW = (sec) => (sec.w - sec.left - sec.right - (sec.cols - 1) * sec.colSpace) / sec.cols;
    const colX = () => S.sec.left + S.col * (colW(S.sec) + S.sec.colSpace);
    const FN_SEP = 14; // separador de las notas al pie
    const bottomLimit = () => S.sec.bottom + (S.pg?.fnH || 0);
    const newPage = () => {
      const sec = S.sec;
      S.page = pdf.addPage([sec.w, sec.h]);
      S.pg = { page: S.page, sec, secIdx: S.secIdx, first: !sec.__pages, fns: [], fnH: 0, excl: [] };
      S.pages.push(S.pg);
      sec.__pages = (sec.__pages || 0) + 1;
      S.y = sec.h - sec.top;
      S.colTop = S.y;
      S.col = 0;
      S.atTop = true;
    };
    const nextCol = () => {
      if (S.col < S.sec.cols - 1) { S.col++; S.y = S.colTop; S.atTop = true; }
      else newPage();
    };
    const enterSection = (i) => {
      const sec = sections[i];
      const cont = sec.type === 'continuous' && S.page && S.sec && Math.abs(S.sec.w - sec.w) < 1 && Math.abs(S.sec.h - sec.h) < 1;
      S.secIdx = i;
      if (cont) {
        // balanceo sencillo: la nueva sección empieza bajo el punto más bajo usado
        S.y = Math.min(S.y, S.maxY ?? S.y);
        S.sec = sec;
        S.colTop = S.y;
        S.col = 0;
        S.maxY = S.y;
      } else {
        S.sec = sec;
        newPage();
      }
    };

    const lineH = (line, pp) => {
      const nat = line.asc + line.desc;
      const rule = pp.lineRule || 'auto';
      if (pp.line == null) return nat;
      if (rule === 'exact') return pp.line / TW;
      if (rule === 'atLeast') return Math.max(pp.line / TW, nat);
      return nat * (pp.line / 240);
    };
    const resolveFields = (runs, ctx) => runs.map((r) => (r.field ? { ...r, text: String(r.field === 'PAGE' ? ctx.page ?? 1 : ctx.total ?? '') } : r));

    /** Prepara las líneas de un párrafo para un ancho dado. */
    function layPara(b, width, ctx = {}, lineBox = null) {
      const pp = b.pp;
      const left = pp.left || 0, right = pp.right || 0, first = pp.first || 0;
      const tabs = (pp.tabs || []).filter((t) => t.align !== 'clear').sort((a, c) => a.pos - c.pos);
      const segments = [[]];
      const anchors = [];
      for (const r of resolveFields(b.runs, ctx)) {
        if (r.brk) { segments.push({ brk: r.brk }); segments.push([]); continue; }
        if (r.anchor) { anchors.push(r.anchor); continue; }
        const seg = segments[segments.length - 1];
        if (r.inline) {
          const it = r.inline;
          const k = Math.min(1, (width - left - right) / it.w || 1);
          const w = it.w * k, h = it.h * k;
          seg.push({ box: { w, h, draw: (page, x, base) => drawItem(page, it, x, base, w, h) } });
        } else seg.push(r);
      }
      const out = [];
      let firstSeg = true;
      for (const seg of segments) {
        if (!Array.isArray(seg)) { out.push(seg); continue; }
        let runs = seg;
        let tabsHere = tabs;
        if (firstSeg && b.label) {
          const sep = b.label.suff === 'space' ? ' ' : b.label.suff === 'nothing' ? '' : '\t';
          runs = [b.label, ...(sep ? [{ ...b.label, text: sep }] : []), ...seg];
          if (sep === '\t' && first < 0) tabsHere = [{ pos: left, align: 'left' }, ...tabs.filter((t) => t.pos > left)];
        }
        const lines = layoutRuns(runs, {
          maxW: Math.max(20, width - left - right), firstOffset: firstSeg ? first : 0, tabs: tabsHere, tabOrigin: left,
          defFont: b.mark.font, defSize: b.mark.size, lineBox: firstSeg ? lineBox : null,
        });
        out.push({ lines });
        firstSeg = false;
      }
      return { parts: out, anchors, left, right, align: pp.jc || 'left', before: pp.before || 0, after: pp.after || 0, pp };
    }

    function drawItem(page, it, x, base, w, h) {
      if (it.img) page.drawImage(it.img, { x, y: base, width: w, height: h });
      else if (it.chart) pending.push(Charts.draw(page, it.chart, { x, y: base, w, h }, { fonts: bank, scheme: theme.scheme, fontName: theme.minor }));
    }

    /** Dibuja bloques sin paginar (celdas, cuadros de texto, encabezados). Devuelve la altura. */
    function drawBlocks(page, list, x, top, width, ctx = {}, measure = false) {
      let y = top;
      let prev = null;
      for (const b of list) {
        if (b.t === 'p') {
          const L = layPara(b, width, ctx);
          const ctxSkip = prev && prev.t === 'p' && L.pp.contextual && prev.style === b.style;
          if (prev || L.before) y -= ctxSkip ? 0 : prev ? Math.max(0, L.before) : L.before;
          for (const part of L.parts) {
            if (!part.lines) continue;
            for (const line of part.lines) {
              const lh = lineH(line, L.pp);
              if (!measure) {
                if (L.pp.shd) page.drawRectangle({ x: x + L.left, y: y - lh, width: width - L.left - L.right, height: lh, color: toRgb(L.pp.shd) });
                drawTextLine(page, line, x + L.left, y - (lh - line.desc), width - L.left - L.right, L.align);
              }
              y -= lh;
            }
          }
          if (!measure) for (const a of L.anchors) drawAnchor(page, a, x, y, null);
          y -= ctxSkip ? 0 : L.after;
          prev = b;
        } else if (b.t === 'tbl') {
          const T = layTable(b, width, ctx);
          for (let ri = 0; ri < T.rows.length; ri++) {
            if (!measure) drawRow(page, T, ri, x + T.offset, y);
            y -= T.rows[ri].h;
          }
          prev = b;
        }
      }
      return top - y;
    }

    /** Calcula columnas y alturas de una tabla. */
    function layTable(t, avail, ctx) {
      let grid = t.grid.length ? t.grid.slice() : [];
      const nCols = Math.max(grid.length, ...t.rows.map((r) => r.cells.reduce((a, c) => Math.max(a, c.gi + c.span), 0)));
      while (grid.length < nCols) grid.push(avail / nCols);
      let total = grid.reduce((a, b) => a + b, 0);
      if (total > avail || total < 1) { const k = avail / (total || 1); grid = grid.map((g) => g * k); total = avail; }
      const offset = t.jc === 'center' ? (avail - total) / 2 : t.jc === 'right' || t.jc === 'end' ? avail - total : t.ind;
      const xs = [0];
      for (const g of grid) xs.push(xs[xs.length - 1] + g);
      const rows = t.rows.map((r) => {
        const cells = r.cells.map((c) => {
          const w = xs[Math.min(c.gi + c.span, nCols)] - xs[c.gi];
          const inner = Math.max(4, w - t.mar.left - t.mar.right);
          const contentH = c.vMerge === 'continue' ? 0 : drawBlocks(null, c.blocks, 0, 0, inner, ctx, true);
          return { ...c, x: xs[c.gi], w, inner, contentH };
        });
        let h = Math.max(0, ...cells.filter((c) => c.vMerge !== 'restart').map((c) => c.contentH + t.mar.top + t.mar.bottom));
        if (r.hRule === 'exact' && r.h) h = r.h;
        else h = Math.max(h, r.h || 0, 4);
        return { ...r, cells, h };
      });
      // celdas combinadas verticalmente: reparte la altura sobrante en la última fila
      rows.forEach((r, ri) => r.cells.forEach((c) => {
        if (c.vMerge !== 'restart') return;
        let end = ri;
        while (end + 1 < rows.length && rows[end + 1].cells.some((d) => d.gi === c.gi && d.vMerge === 'continue')) end++;
        c.spanRows = end - ri + 1;
        const need = c.contentH + t.mar.top + t.mar.bottom;
        const have = rows.slice(ri, end + 1).reduce((a, x) => a + x.h, 0);
        if (need > have) rows[end].h += need - have;
      }));
      return { t, rows, offset, nCols, ctx };
    }

    function drawRow(page, T, ri, x0, top) {
      const { t, rows } = T;
      const row = rows[ri];
      for (const c of row.cells) {
        let h = row.h;
        if (c.vMerge === 'restart' && c.spanRows > 1) h = rows.slice(ri, ri + c.spanRows).reduce((a, r) => a + r.h, 0);
        const x = x0 + c.x;
        const below = rows[ri + 1]?.cells.find((d) => d.gi === c.gi);
        const continuesBelow = below && below.vMerge === 'continue';
        if (c.shd && c.vMerge !== 'continue') page.drawRectangle({ x, y: top - h, width: c.w, height: h, color: toRgb(c.shd) });
        const last = c.gi + c.span >= T.nCols;
        const edge = (side, inside) => (c.borders[side] !== undefined ? c.borders[side] : inside ? t.borders[side === 'top' || side === 'bottom' ? 'insideH' : 'insideV'] : t.borders[side]);
        const line = (b, x1, y1, x2, y2) => b && page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: b.w, color: toRgb(b.color === 'auto' ? '000000' : b.color) });
        if (c.vMerge !== 'continue') line(edge('top', ri > 0), x, top, x + c.w, top);
        if (!continuesBelow) line(edge('bottom', ri < rows.length - 1), x, top - row.h, x + c.w, top - row.h);
        line(edge('left', c.gi > 0), x, top, x, top - row.h);
        line(edge('right', !last), x + c.w, top, x + c.w, top - row.h);
        if (c.vMerge === 'continue') continue;
        let ty = top - t.mar.top;
        const free = h - c.contentH - t.mar.top - t.mar.bottom;
        if (c.vAlign === 'center') ty -= free / 2;
        else if (c.vAlign === 'bottom') ty -= free;
        drawBlocks(page, c.blocks, x + t.mar.left, ty, c.inner, T.ctx);
      }
    }

    function drawAnchor(page, a, colLeft, paraTop, sec) {
      sec = sec || S.sec;
      const pw = sec.w, ph = sec.h;
      const frameH = { page: [0, pw], margin: [sec.left, pw - sec.right], column: [colLeft, colLeft + colW(sec)], character: [colLeft, colLeft + colW(sec)], leftMargin: [0, sec.left], rightMargin: [pw - sec.right, pw], insideMargin: [0, sec.left], outsideMargin: [pw - sec.right, pw] }[a.posH.rel] || [sec.left, pw - sec.right];
      let x = a.posH.off != null ? frameH[0] + a.posH.off : a.posH.align === 'center' ? (frameH[0] + frameH[1] - a.w) / 2 : a.posH.align === 'right' || a.posH.align === 'outside' ? frameH[1] - a.w : frameH[0];
      // en coordenadas desde arriba
      const pTopFromTop = ph - paraTop;
      const frameV = { page: [0, ph], margin: [sec.top, ph - sec.bottom], topMargin: [0, sec.top], bottomMargin: [ph - sec.bottom, ph], paragraph: [pTopFromTop, ph], line: [pTopFromTop, ph] }[a.posV.rel] || [pTopFromTop, ph];
      let yTop = a.posV.off != null ? frameV[0] + a.posV.off : a.posV.align === 'center' ? (frameV[0] + frameV[1] - a.h) / 2 : a.posV.align === 'bottom' ? frameV[1] - a.h : frameV[0];
      x = Math.max(-a.w / 2, Math.min(pw - a.w / 2, x));
      const yPdf = ph - yTop - a.h;
      if (a.img || a.chart) drawItem(page, a, x, yPdf, a.w, a.h);
      else {
        if (a.fill || a.line) page.drawRectangle({ x, y: yPdf, width: a.w, height: a.h, color: a.fill ? toRgb(a.fill) : undefined, borderColor: a.line ? toRgb(a.line) : undefined, borderWidth: a.line ? 0.75 : 0 });
        if (a.blocks) drawBlocks(page, a.blocks, x + a.ins[0], ph - yTop - a.ins[1], Math.max(10, a.w - a.ins[0] - a.ins[2]), { page: S.pages.length, total: '' });
      }
      return { x, w: a.w, top: ph - yTop, bottom: yPdf, wrap: a.behind ? 'wrapNone' : a.wrap, wrapText: a.wrapText, dist: a.dist || { l: 9, r: 9, t: 0, b: 0 }, rel: a.posV.rel };
    }

    /* ── recorrido del cuerpo con paginación ── */
    let prevBlock = null;
    for (let bi = 0; bi < blocks.length; bi++) {
      const b = blocks[bi];
      if (bi % 15 === 0) progress(0.6 + (0.3 * bi) / blocks.length, 'Maquetando el PDF…');
      if (b.sec !== S.secIdx) enterSection(b.sec);
      const width = colW(S.sec);
      try {
        if (b.t === 'tbl') {
          const T = layTable(b, width, { page: S.pages.length });
          const headerRows = [];
          for (let ri = 0; ri < T.rows.length && T.rows[ri].header; ri++) headerRows.push(ri);
          if (!S.atTop && prevBlock?.t === 'p') S.y -= 0;
          for (let ri = 0; ri < T.rows.length; ri++) {
            const rh = T.rows[ri].h;
            if (S.y - rh < bottomLimit() && !S.atTop) {
              nextCol();
              if (!headerRows.includes(ri)) for (const hr of headerRows) { drawRow(S.page, T, hr, colX() + T.offset, S.y); S.y -= T.rows[hr].h; }
            }
            drawRow(S.page, T, ri, colX() + T.offset, S.y);
            S.y -= rh;
            S.atTop = false;
          }
          S.maxY = Math.min(S.maxY ?? S.y, S.y);
          prevBlock = b;
          continue;
        }
        // párrafo
        const pp = b.pp;
        if (pp.pageBreakBefore && !S.atTop) newPage();
        const ctxSkip = prevBlock && prevBlock.t === 'p' && pp.contextual && prevBlock.style === b.style;
        if (!S.atTop) S.y -= ctxSkip ? 0 : pp.before || 0;
        let paraTop = S.y;
        // imágenes y cuadros flotantes: se dibujan antes y apartan el texto
        for (const r of b.runs) {
          if (!r.anchor) continue;
          const a = drawAnchor(S.page, r.anchor, colX(), paraTop);
          if (a.wrap === 'wrapTopAndBottom' && /paragraph|line/.test(a.rel)) { S.y = Math.min(S.y, a.bottom - a.dist.b - 4); paraTop = S.y; }
          else if (/wrapSquare|wrapTight|wrapThrough/.test(a.wrap)) S.pg.excl.push({ x1: a.x - a.dist.l, x2: a.x + a.w + a.dist.r, top: a.top + a.dist.t, bottom: a.bottom - a.dist.b, side: a.wrapText });
        }
        const est = (b.mark.size || 11) * 1.2;
        const lineBox = (left, right) => (S.pg.excl.length ? (i, prev) => {
          // simula la paginación: ¿en qué columna/página caerá esta línea?
          let col = S.col, y = paraTop, colTop = S.colTop, samePage = true;
          const place = (lh, first) => {
            if (y - lh < bottomLimit() - 0.5 && !(first && S.atTop)) {
              if (col < S.sec.cols - 1) { col++; y = colTop; }
              else { samePage = false; col = 0; y = S.sec.h - S.sec.top; colTop = y; }
            }
          };
          prev.forEach((l, k) => { const lh = lineH(l, pp); place(lh, k === 0); y -= lh; });
          place(est, prev.length === 0);
          // en una página nueva todavía no hay imágenes flotantes: ancho completo
          if (!samePage) return { offset: 0, width: width - left - right };
          const top = y;
          const cx = S.sec.left + col * (colW(S.sec) + S.sec.colSpace);
          const x0 = cx + left, x1 = cx + width - right;
          let free = [[x0, x1]];
          for (const e of S.pg.excl) {
            if (e.bottom >= top || e.top <= top - est) continue;
            const next = [];
            for (const [a0, a1] of free) {
              if (e.x2 <= a0 || e.x1 >= a1) { next.push([a0, a1]); continue; }
              if (e.x1 > a0 && e.side !== 'left') next.push([a0, e.x1]);
              if (e.x2 < a1 && e.side !== 'right') next.push([e.x2, a1]);
            }
            free = next;
          }
          const best = free.sort((p, q) => (q[1] - q[0]) - (p[1] - p[0]))[0] || [x0, x0];
          return { offset: best[0] - x0, width: best[1] - best[0] };
        } : null);
        const L = layPara(b, width, { page: S.pages.length }, lineBox(pp.left || 0, pp.right || 0));
        // mantener con el siguiente (títulos): que no queden solos al final de la página
        if (pp.keepNext && !S.atTop) {
          const h = L.parts.reduce((a, p) => a + (p.lines ? p.lines.reduce((s2, l) => s2 + lineH(l, pp), 0) : 0), 0);
          if (S.y - h - 30 < bottomLimit()) nextCol();
        }
        const textW = (sec) => sec.w - sec.left - sec.right;
        const fnHeight = (id) => (footnotes[id] ? drawBlocks(null, footnotes[id], 0, 0, textW(S.sec), {}, true) + 2 : 0);
        for (const part of L.parts) {
          if (part.brk) {
            if (part.brk === 'page') newPage(); else nextCol();
            continue;
          }
          for (const line of part.lines) {
            const lh = lineH(line, pp);
            // notas al pie de esta línea: reservan sitio al pie de la página
            const ids = [...new Set(line.items.filter((it) => it.fn).map((it) => it.fn))].filter((id) => footnotes[id]);
            const need = () => ids.reduce((acc, id) => acc + fnHeight(id), 0) + (ids.length && !S.pg.fns.length ? FN_SEP : 0);
            if (S.y - lh < bottomLimit() + need() - 0.5 && !S.atTop) nextCol();
            if (ids.length) { S.pg.fnH += need(); S.pg.fns.push(...ids); }
            if (!line.spacer) {
              const x = colX();
              if (pp.shd) S.page.drawRectangle({ x: x + L.left, y: S.y - lh, width: width - L.left - L.right, height: lh, color: toRgb(pp.shd) });
              drawTextLine(S.page, line, x + L.left, S.y - (lh - line.desc), width - L.left - L.right, L.align);
            }
            S.y -= lh;
            S.atTop = false;
          }
        }
        const bd = L.pp.borders;
        if (bd?.bottom) S.page.drawLine({ start: { x: colX() + L.left, y: S.y - 1 - bd.bottom.space }, end: { x: colX() + width - L.right, y: S.y - 1 - bd.bottom.space }, thickness: bd.bottom.w, color: toRgb(bd.bottom.color === 'auto' ? '000000' : bd.bottom.color) });
        if (bd?.top) S.page.drawLine({ start: { x: colX() + L.left, y: paraTop + 1 }, end: { x: colX() + width - L.right, y: paraTop + 1 }, thickness: bd.top.w, color: toRgb(bd.top.color === 'auto' ? '000000' : bd.top.color) });
        S.y -= ctxSkip ? 0 : L.after;
        S.maxY = Math.min(S.maxY ?? S.y, S.y);
        prevBlock = b;
      } catch (e) {
        console.warn('No se pudo maquetar un bloque', e);
      }
    }
    if (!S.page) { S.sec = sections[0]; newPage(); }

    /* ── notas al pie ── */
    for (const pg of S.pages) {
      if (!pg.fns.length) continue;
      const sec = pg.sec;
      const w = sec.w - sec.left - sec.right;
      let y = sec.bottom + pg.fnH;
      pg.page.drawLine({ start: { x: sec.left, y: y - 6 }, end: { x: sec.left + Math.min(144, w / 3), y: y - 6 }, thickness: 0.6, color: toRgb('808080') });
      y -= FN_SEP;
      for (const id of pg.fns) y -= drawBlocks(pg.page, footnotes[id], sec.left, y, w, {}) + 2;
    }

    /* ── encabezados y pies ── */
    progress(0.92, 'Añadiendo encabezados y pies de página…');
    const hfCache = {};
    const hfBlocks = async (rid) => {
      if (!rid || !rels[rid]) return null;
      if (!(rid in hfCache)) {
        hfCache[rid] = (async () => {
          const path = rels[rid].target;
          const d = await read(path);
          if (!d) return null;
          return parseBlocks(d.documentElement, await ooxmlRels(zip, path), null);
        })();
      }
      return hfCache[rid];
    };
    const total = S.pages.length;
    let pageNo = 0;
    let lastSec = -1;
    for (const pg of S.pages) {
      if (pg.secIdx !== lastSec) { if (pg.sec.pgStart != null) pageNo = pg.sec.pgStart - 1; lastSec = pg.secIdx; }
      pageNo++;
      const sec = pg.sec;
      const kind = sec.titlePg && pg.first ? 'first' : 'default';
      const hRef = sec.titlePg && pg.first ? sec.headers.first : sec.headers.default;
      const fRef = sec.titlePg && pg.first ? sec.footers.first : sec.footers.default;
      void kind;
      const ctx = { page: pageNo, total };
      const width = sec.w - sec.left - sec.right;
      try {
        const hb = await hfBlocks(hRef);
        if (hb && hb.length) drawBlocks(pg.page, hb, sec.left, sec.h - sec.headerDist, width, ctx);
        const fb = await hfBlocks(fRef);
        if (fb && fb.length) {
          const h = drawBlocks(null, fb, 0, 0, width, ctx, true);
          drawBlocks(pg.page, fb, sec.left, sec.footerDist + h, width, ctx);
        }
      } catch (e) {
        console.warn('Encabezado/pie omitido', e);
      }
    }
    await Promise.all(pending);
    progress(0.97, 'Guardando…');
    return { bytes: await pdf.save(), skipped, pages: total };
  }

  return { toPdf };
})();
