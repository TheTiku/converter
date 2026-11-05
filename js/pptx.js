'use strict';

/**
 * Conversor propio de PowerPoint (.pptx) a PDF.
 * Interpreta patrones/diseños/temas, fondos (sólidos, degradados e imagen), formas
 * con relleno y borde (incluidos degradados y giros), textos con su fuente y estilo
 * heredado, viñetas, tablas, imágenes, SmartArt y gráficos.
 */
const Pptx = (() => {
  const EMU = 12700;
  const { kid, kids, desc, all, attr, num } = X;

  /** Trazado SVG (Y hacia abajo) de formas predefinidas comunes. */
  function shapePath(geom, w, h, adj = {}) {
    const a = (k, d) => (adj[k] != null ? adj[k] / 100000 : d);
    const m = Math.min(w, h);
    switch (geom) {
      case 'roundRect': {
        const r = Math.min(m / 2, m * a('adj', 0.16667));
        return `M ${r} 0 L ${w - r} 0 Q ${w} 0 ${w} ${r} L ${w} ${h - r} Q ${w} ${h} ${w - r} ${h} L ${r} ${h} Q 0 ${h} 0 ${h - r} L 0 ${r} Q 0 0 ${r} 0 Z`;
      }
      case 'triangle': { const x = w * a('adj', 0.5); return `M ${x} 0 L ${w} ${h} L 0 ${h} Z`; }
      case 'rtTriangle': return `M 0 0 L ${w} ${h} L 0 ${h} Z`;
      case 'diamond': return `M ${w / 2} 0 L ${w} ${h / 2} L ${w / 2} ${h} L 0 ${h / 2} Z`;
      case 'parallelogram': { const o = m * a('adj', 0.25); return `M ${o} 0 L ${w} 0 L ${w - o} ${h} L 0 ${h} Z`; }
      case 'trapezoid': { const o = m * a('adj', 0.25); return `M ${o} 0 L ${w - o} 0 L ${w} ${h} L 0 ${h} Z`; }
      case 'pentagon': case 'homePlate': { const o = m * a('adj', 0.5); return `M 0 0 L ${w - o} 0 L ${w} ${h / 2} L ${w - o} ${h} L 0 ${h} Z`; }
      case 'chevron': { const o = m * a('adj', 0.5); return `M 0 0 L ${w - o} 0 L ${w} ${h / 2} L ${w - o} ${h} L 0 ${h} L ${o} ${h / 2} Z`; }
      case 'hexagon': { const o = m * a('adj', 0.25); return `M ${o} 0 L ${w - o} 0 L ${w} ${h / 2} L ${w - o} ${h} L ${o} ${h} L 0 ${h / 2} Z`; }
      case 'octagon': { const o = m * a('adj', 0.29289); return `M ${o} 0 L ${w - o} 0 L ${w} ${o} L ${w} ${h - o} L ${w - o} ${h} L ${o} ${h} L 0 ${h - o} L 0 ${o} Z`; }
      case 'rightArrow': { const t = h * a('adj1', 0.5), hw = m * a('adj2', 0.5); const y1 = (h - t) / 2; return `M 0 ${y1} L ${w - hw} ${y1} L ${w - hw} 0 L ${w} ${h / 2} L ${w - hw} ${h} L ${w - hw} ${h - y1} L 0 ${h - y1} Z`; }
      case 'leftArrow': { const t = h * a('adj1', 0.5), hw = m * a('adj2', 0.5); const y1 = (h - t) / 2; return `M ${w} ${y1} L ${hw} ${y1} L ${hw} 0 L 0 ${h / 2} L ${hw} ${h} L ${hw} ${h - y1} L ${w} ${h - y1} Z`; }
      case 'upArrow': { const t = w * a('adj1', 0.5), hh = m * a('adj2', 0.5); const x1 = (w - t) / 2; return `M ${x1} ${h} L ${x1} ${hh} L 0 ${hh} L ${w / 2} 0 L ${w} ${hh} L ${w - x1} ${hh} L ${w - x1} ${h} Z`; }
      case 'downArrow': { const t = w * a('adj1', 0.5), hh = m * a('adj2', 0.5); const x1 = (w - t) / 2; return `M ${x1} 0 L ${w - x1} 0 L ${w - x1} ${h - hh} L ${w} ${h - hh} L ${w / 2} ${h} L 0 ${h - hh} L ${x1} ${h - hh} Z`; }
      case 'flowChartTerminator': { const r = h / 2; return `M ${r} 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w - r} ${h} L ${r} ${h} A ${r} ${r} 0 0 1 ${r} 0 Z`; }
      case 'star5': {
        const pts = [];
        for (let i = 0; i < 10; i++) { const ang = -Math.PI / 2 + (i * Math.PI) / 5; const rr = i % 2 ? 0.38 : 0.5; pts.push(`${w / 2 + Math.cos(ang) * w * rr} ${h / 2 + Math.sin(ang) * h * rr}`); }
        return `M ${pts.join(' L ')} Z`;
      }
      default: return null;
    }
  }

  async function toPdf(bytes, progress) {
    let zip;
    try {
      zip = await JSZip.loadAsync(bytes);
    } catch (_) {
      throw new Error('No se pudo leer la presentación. Asegúrate de que es un .pptx (los .ppt antiguos no son compatibles).');
    }
    const parseCache = {};
    const xml = async (p) => {
      if (!p) return null;
      if (p in parseCache) return parseCache[p];
      const f = zip.file(p);
      return (parseCache[p] = f ? X.parse(await f.async('string')) : null);
    };
    const relsCache = {};
    const rels = async (p) => (relsCache[p] = relsCache[p] || ooxmlRels(zip, p));
    const relOf = (map, kind) => Object.values(map).find((r) => r.type.endsWith('/' + kind))?.target;

    const pres = await xml('ppt/presentation.xml');
    if (!pres) throw new Error('No se pudo leer la presentación (.pptx no válido).');
    const sz = desc(pres, 'sldSz');
    const W = num(sz, 'cx', 9144000) / EMU;
    const H = num(sz, 'cy', 6858000) / EMU;
    const presRels = await rels('ppt/presentation.xml');
    const slides = all(pres, 'sldId').map((s) => presRels[s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')]?.target).filter(Boolean);
    if (!slides.length) throw new Error('La presentación no tiene diapositivas.');

    const doc = await PDFLib.PDFDocument.create();
    const bank = FontBank.forDoc(doc);
    // fuentes usadas
    const faces = new Set(['Calibri']);
    const themeCache = {};
    const themeOf = async (masterPath) => {
      if (!themeCache[masterPath]) {
        const t = relOf(await rels(masterPath), 'theme');
        themeCache[masterPath] = parseTheme(await xml(t));
      }
      return themeCache[masterPath];
    };
    for (const f of Object.keys(zip.files)) {
      if (!/^ppt\/(slides|slideLayouts|slideMasters|theme)\/[^/]+\.xml$/.test(f)) continue;
      const s = await zip.file(f).async('string');
      for (const m of s.matchAll(/<a:latin typeface="([^"+][^"]*)"/g)) faces.add(m[1]);
    }
    progress(0.02, 'Cargando las fuentes…');
    await bank.prepare([...faces]);

    const imgCache = {};
    let skipped = 0;
    const image = async (target) => {
      if (!target) return null;
      if (!(target in imgCache)) {
        const f = zip.file(target);
        imgCache[target] = f ? embedAnyImage(doc, await f.async('uint8array'), mimeOf(target)).then((e) => e && e.img) : Promise.resolve(null);
      }
      const img = await imgCache[target];
      if (!img) skipped++;
      return img;
    };
    const pending = [];

    for (let si = 0; si < slides.length; si++) {
      progress(0.05 + (0.9 * si) / slides.length, `Convirtiendo diapositiva ${si + 1} de ${slides.length}…`);
      const sPath = slides[si];
      const sDoc = await xml(sPath);
      if (!sDoc) continue;
      if (attr(sDoc.documentElement, 'show') === '0') continue; // diapositiva oculta
      const sRels = await rels(sPath);
      const lPath = relOf(sRels, 'slideLayout');
      const lDoc = await xml(lPath);
      const lRels = lPath ? await rels(lPath) : {};
      const mPath = relOf(lRels, 'slideMaster');
      const mDoc = await xml(mPath);
      const mRels = mPath ? await rels(mPath) : {};
      const theme = mPath ? await themeOf(mPath) : parseTheme(null);
      const scheme = theme.scheme;
      const page = doc.addPage([W, H]);
      const col = (el, ph) => ooxmlColor(el, scheme, ph);
      const typeface = (tf) => (!tf ? null : tf === '+mj-lt' ? theme.major : tf === '+mn-lt' ? theme.minor : tf.startsWith('+') ? theme.minor : tf);

      /* ── rellenos ── */
      const fillOf = (spPr, style) => {
        if (!spPr && !style) return null;
        if (kid(spPr, 'noFill')) return null;
        const solid = kid(spPr, 'solidFill');
        if (solid) return { color: col(solid) };
        const grad = kid(spPr, 'gradFill');
        if (grad) return { grad: gradOf(grad) };
        const blip = kid(spPr, 'blipFill');
        if (blip) return { blip: desc(blip, 'blip') };
        const ref = kid(style, 'fillRef');
        if (ref && num(ref, 'idx') > 0) return { color: col(ref) };
        return null;
      };
      const gradOf = (g, ph) => {
        const stops = all(g, 'gs').map((s) => ({ pos: num(s, 'pos') / 100000, c: col(s, ph) })).filter((s) => s.c).sort((a, b) => a.pos - b.pos);
        const lin = kid(g, 'lin');
        return { stops, ang: lin ? num(lin, 'ang') / 60000 : 90, radial: !!kid(g, 'path') };
      };
      const colorAt = (stops, t) => {
        if (!stops.length) return [1, 1, 1];
        if (t <= stops[0].pos) return hexRgb(stops[0].c.hex);
        for (let i = 1; i < stops.length; i++) {
          if (t <= stops[i].pos) {
            const a = stops[i - 1], b = stops[i];
            const k = (t - a.pos) / (b.pos - a.pos || 1);
            const ca = hexRgb(a.c.hex), cb = hexRgb(b.c.hex);
            return ca.map((v, j) => v + (cb[j] - v) * k);
          }
        }
        return hexRgb(stops[stops.length - 1].c.hex);
      };
      /** Rellena un rectángulo (o forma) con un degradado aproximado por franjas. */
      const drawGrad = (g, x, y, w, h, path) => {
        if (path || g.radial) {
          const c = colorAt(g.stops, 0.5);
          if (path) page.drawSvgPath(path, { x, y: y + h, color: PDFLib.rgb(...c) });
          else page.drawRectangle({ x, y, width: w, height: h, color: PDFLib.rgb(...c) });
          return;
        }
        const N = 64;
        const ang = ((g.ang % 360) + 360) % 360;
        const vertical = (ang > 45 && ang < 135) || (ang > 225 && ang < 315);
        const reverse = vertical ? ang > 180 : ang > 90 && ang < 270;
        // cada franja cubre desde su inicio hasta el final: así no quedan rendijas entre franjas
        for (let i = 0; i < N; i++) {
          const t = (i + 0.5) / N;
          const c = PDFLib.rgb(...colorAt(g.stops, reverse ? 1 - t : t));
          if (vertical) page.drawRectangle({ x, y, width: w, height: h - (h * i) / N, color: c });
          else page.drawRectangle({ x: x + (w * i) / N, y, width: w - (w * i) / N, height: h, color: c });
        }
      };

      /* ── fondo ── */
      const bgOf = async (d, r) => {
        const bg = desc(d, 'bg');
        if (!bg) return null;
        const bgPr = kid(bg, 'bgPr');
        if (bgPr) {
          const f = fillOf(bgPr);
          if (f?.blip) return { img: await image(r[attr(f.blip, 'embed')]?.target) };
          return f;
        }
        const ref = kid(bg, 'bgRef');
        if (ref) {
          // índice >= 1001: relleno de fondo del tema
          const idx = num(ref, 'idx');
          const phClr = col(ref)?.hex;
          const lst = desc(await xml(relOf(mRels, 'theme')), 'bgFillStyleLst');
          const style = lst ? lst.children[idx - 1001] : null;
          if (style?.localName === 'gradFill') return { grad: gradOf(style, phClr) };
          if (style?.localName === 'solidFill') return { color: col(style, phClr) };
          return { color: col(ref) };
        }
        return null;
      };
      const bg = (await bgOf(sDoc, sRels)) || (await bgOf(lDoc, lRels)) || (await bgOf(mDoc, mRels));
      if (bg?.img) page.drawImage(bg.img, { x: 0, y: 0, width: W, height: H });
      else if (bg?.grad) drawGrad(bg.grad, 0, 0, W, H);
      else page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: bg?.color ? toRgb(bg.color.hex) : PDFLib.rgb(1, 1, 1) });

      /* ── herencia de marcadores de posición ── */
      const phOf = (sp) => {
        const nv = kid(sp, 'nvSpPr') || kid(sp, 'nvPicPr') || kid(sp, 'nvGraphicFramePr');
        const ph = desc(nv, 'ph');
        return ph ? { type: attr(ph, 'type') || 'body', idx: attr(ph, 'idx') } : null;
      };
      const findPh = (d, ph, byIdx) => {
        if (!d || !ph) return null;
        const list = all(d, 'sp').filter((s) => phOf(s));
        const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
        return (byIdx && ph.idx != null && list.find((s) => phOf(s).idx === ph.idx)) ||
          list.find((s) => phOf(s).type === ph.type) || list.find((s) => norm(phOf(s).type) === norm(ph.type)) || null;
      };
      const masterStyle = (ph) => {
        const tx = desc(mDoc, 'txStyles');
        if (!tx) return null;
        if (!ph) return kid(tx, 'otherStyle');
        return /title/i.test(ph.type) ? kid(tx, 'titleStyle') : ['body', 'subTitle', 'obj'].includes(ph.type) ? kid(tx, 'bodyStyle') : kid(tx, 'otherStyle');
      };

      /* ── texto ── */
      async function renderText(txBody, box, ph, layer, measure = false) {
        if (!txBody) return 0;
        const lSp = ph && layer === 'slide' ? findPh(lDoc, ph, true) : null;
        const mSp = ph && layer !== 'master' ? findPh(mDoc, ph, false) : null;
        const lists = [kid(txBody, 'lstStyle'), lSp && desc(lSp, 'lstStyle'), mSp && desc(mSp, 'lstStyle'), masterStyle(ph)].filter(Boolean);
        const bodies = [kid(txBody, 'bodyPr'), lSp && desc(lSp, 'bodyPr'), mSp && desc(mSp, 'bodyPr')].filter(Boolean);
        const bodyAttr = (a) => { for (const b of bodies) if (b.hasAttribute(a)) return b.getAttribute(a); return null; };
        const lvlProp = (lvl, fn) => {
          for (const l of lists) {
            const p = kid(l, `lvl${lvl + 1}pPr`);
            const v = p && fn(p);
            if (v != null) return v;
          }
          return null;
        };
        const ins = (a, d) => (bodyAttr(a) != null ? +bodyAttr(a) : d) / EMU;
        const [li, ti, ri, bi] = [ins('lIns', 91440), ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720)];
        const anchor = bodyAttr('anchor') || 't';
        const wrapNone = bodyAttr('wrap') === 'none';
        const autofit = desc(kid(txBody, 'bodyPr'), 'normAutofit');
        const fscale = autofit ? num(autofit, 'fontScale', 100000) / 100000 : 1;
        const lnRed = autofit ? num(autofit, 'lnSpcReduction', 0) / 100000 : 0;
        const maxW = wrapNone ? 5000 : Math.max(10, box.w - li - ri);

        const paras = [];
        let autoNum = 0;
        for (const p of kids(txBody, 'p')) {
          const pPr = kid(p, 'pPr');
          const lvl = num(pPr, 'lvl', 0);
          const algn = pPr?.getAttribute('algn') || lvlProp(lvl, (x) => x.getAttribute('algn'));
          const defR = (a) => lvlProp(lvl, (x) => { const d = kid(x, 'defRPr'); return d && d.hasAttribute(a) ? d.getAttribute(a) : null; });
          const defColor = lvlProp(lvl, (x) => col(kid(kid(x, 'defRPr'), 'solidFill')));
          const defFace = lvlProp(lvl, (x) => attr(kid(kid(x, 'defRPr'), 'latin'), 'typeface'));
          const isTitle = ph && /title/i.test(ph.type);
          const runs = [];
          for (const r of p.children) {
            if (r.localName === 'br') { runs.push({ text: '\n', font: await bank.get(theme.minor, false, false), size: 12 }); continue; }
            if (r.localName !== 'r' && r.localName !== 'fld') continue;
            const rPr = kid(r, 'rPr');
            let t = kid(r, 't')?.textContent || '';
            if (!t) continue;
            const szv = rPr?.getAttribute('sz') || defR('sz') || (isTitle ? 4400 : 1800);
            const c = col(kid(rPr, 'solidFill')) || col(kid(rPr, 'gradFill') && desc(kid(rPr, 'gradFill'), 'gs')) || defColor;
            const bold = (rPr?.getAttribute('b') ?? defR('b')) === '1';
            const italic = (rPr?.getAttribute('i') ?? defR('i')) === '1';
            const face = typeface(attr(kid(rPr, 'latin'), 'typeface') || defFace) || (isTitle ? theme.major : theme.minor);
            if ((rPr?.getAttribute('cap') ?? defR('cap')) === 'all') t = t.toUpperCase();
            const baseline = num(rPr, 'baseline', 0) / 100000;
            let size = (+szv / 100) * fscale;
            const rise = baseline ? size * baseline : 0;
            if (baseline) size *= 0.7;
            runs.push({
              text: t, size, rise,
              font: await bank.get(face, bold, italic),
              underline: !!(rPr?.getAttribute('u') && rPr.getAttribute('u') !== 'none'),
              strike: rPr?.getAttribute('strike') && rPr.getAttribute('strike') !== 'noStrike',
              color: c ? toRgb(c.hex) : toRgb(scheme.tx1),
            });
          }
          const endSz = (+(kid(p, 'endParaRPr')?.getAttribute('sz') || defR('sz') || (isTitle ? 4400 : 1800)) / 100) * fscale;
          const hasText = runs.some((r) => r.text.trim());
          const buNone = desc(pPr, 'buNone') || (!desc(pPr, 'buChar') && !desc(pPr, 'buAutoNum') && lvlProp(lvl, (x) => (kid(x, 'buNone') ? true : kid(x, 'buChar') || kid(x, 'buAutoNum') ? false : null)));
          const buChar = desc(pPr, 'buChar') || (!buNone && lvlProp(lvl, (x) => kid(x, 'buChar')));
          const buAuto = desc(pPr, 'buAutoNum') || (!buNone && lvlProp(lvl, (x) => kid(x, 'buAutoNum')));
          let bullet = null;
          if (hasText && !buNone) {
            if (buAuto) {
              const sch = attr(buAuto, 'type') || 'arabicPeriod';
              autoNum++;
              const n = /alphaLc/.test(sch) ? String.fromCharCode(96 + autoNum) : /alphaUc/.test(sch) ? String.fromCharCode(64 + autoNum) : /romanLc/.test(sch) ? ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'][autoNum - 1] || autoNum : String(autoNum);
              bullet = /Paren(R|Both)/.test(sch) ? `${n})` : `${n}.`;
            } else if (buChar) {
              const ch = attr(buChar, 'char') || '•';
              bullet = /[-]/.test(ch) || ch === '§' || ch === 'Ø' || ch === 'ü' || ch === 'q' || ch === 'n' ? '•' : ch;
            }
          }
          if (!buAuto) autoNum = 0;
          const marL = (+(pPr?.getAttribute('marL') ?? lvlProp(lvl, (x) => x.getAttribute('marL')) ?? 0)) / EMU;
          const indent = (+(pPr?.getAttribute('indent') ?? lvlProp(lvl, (x) => x.getAttribute('indent')) ?? 0)) / EMU;
          const lnPct = desc(kid(pPr, 'lnSpc'), 'spcPct') || lvlProp(lvl, (x) => desc(kid(x, 'lnSpc'), 'spcPct'));
          const lnPts = desc(kid(pPr, 'lnSpc'), 'spcPts') || lvlProp(lvl, (x) => desc(kid(x, 'lnSpc'), 'spcPts'));
          const spc = (n) => {
            const e = kid(pPr, n) || lvlProp(lvl, (x) => kid(x, n));
            if (!e) return 0;
            const pts = kid(e, 'spcPts');
            if (pts) return num(pts, 'val') / 100;
            const pct = kid(e, 'spcPct');
            return pct ? (num(pct, 'val') / 100000) * (runs[0]?.size || endSz) : 0;
          };
          const bulletColor = col(kid(pPr, 'buClr')) || lvlProp(lvl, (x) => col(kid(x, 'buClr')));
          const font0 = runs[0]?.font || (await bank.get(isTitle ? theme.major : theme.minor, false, false));
          const allRuns = runs.length ? runs : [{ text: '', font: font0, size: endSz }];
          const left = Math.min(marL, maxW * 0.7);
          const lines = layoutRuns(allRuns, { maxW: Math.max(10, maxW - left), defFont: font0, defSize: endSz });
          paras.push({
            lines, left, indent, bullet, font0, bSize: runs[0]?.size || endSz, bColor: bulletColor ? toRgb(bulletColor.hex) : runs[0]?.color,
            algn: { ctr: 'center', r: 'right', just: 'both', dist: 'both' }[algn] || 'left',
            lnMul: lnPct ? num(lnPct, 'val') / 100000 - lnRed : 1 - lnRed, lnPts: lnPts ? num(lnPts, 'val') / 100 : null,
            before: spc('spcBef'), after: spc('spcAft'),
          });
        }
        const lh = (p, l) => (p.lnPts ? p.lnPts : (l.asc + l.desc) * Math.max(0.5, p.lnMul));
        let total = 0;
        paras.forEach((p, i) => { total += (i ? p.before : 0) + p.lines.reduce((a, l) => a + lh(p, l), 0) + (i < paras.length - 1 ? p.after : 0); });
        if (measure) return total + ti + bi;
        const innerH = box.h - ti - bi;
        let top = H - box.y - ti;
        if (anchor === 'ctr') top -= (innerH - total) / 2;
        else if (anchor === 'b') top -= innerH - total;
        const bx = box.x + li;
        paras.forEach((p, i) => {
          if (i) top -= p.before;
          p.lines.forEach((line, j) => {
            const h = lh(p, line);
            const base = top - (h - line.desc);
            if (j === 0 && p.bullet) {
              const bl = layoutRuns([{ text: p.bullet, font: p.font0, size: p.bSize, color: p.bColor }], { maxW: 60, defFont: p.font0, defSize: p.bSize })[0];
              const bxPos = bx + p.left + p.indent;
              drawTextLine(page, bl, bxPos, base, 60, 'left');
            }
            drawTextLine(page, line, bx + p.left, base, (wrapNone ? box.w - li - ri : maxW) - p.left, p.algn);
            top -= h;
          });
          top -= p.after;
        });
        return total;
      }

      /* ── formas ── */
      const xfrmOf = (el, ph, layer) => {
        let x = desc(kid(el, 'spPr') || kid(el, 'grpSpPr'), 'xfrm') || kid(el, 'xfrm');
        if (!x && ph) {
          const src = (layer === 'slide' && findPh(lDoc, ph, true)) || (layer !== 'master' && findPh(mDoc, ph, false));
          x = src ? desc(kid(src, 'spPr'), 'xfrm') : null;
        }
        if (!x) return null;
        const off = kid(x, 'off'), ext = kid(x, 'ext');
        return { x: num(off, 'x'), y: num(off, 'y'), w: num(ext, 'cx'), h: num(ext, 'cy'), rot: num(x, 'rot') / 60000, flipH: attr(x, 'flipH') === '1', flipV: attr(x, 'flipV') === '1' };
      };
      /** Origen (esquina inferior izquierda) de un rectángulo girado alrededor de su centro. */
      const rotOrigin = (box, rotDeg) => {
        const a = (-rotDeg * Math.PI) / 180;
        const cx = box.x + box.w / 2, cy = H - box.y - box.h / 2;
        return { x: cx + (-box.w / 2) * Math.cos(a) - (-box.h / 2) * Math.sin(a), y: cy + (-box.w / 2) * Math.sin(a) + (-box.h / 2) * Math.cos(a), rotate: PDFLib.degrees(-rotDeg) };
      };

      const renderTree = async (tree, relsMap, layer, T) => {
        if (!tree) return;
        for (const el of tree.children) {
          const name = el.localName;
          const ph = ['sp', 'pic', 'graphicFrame'].includes(name) ? phOf(el) : null;
          if (ph && layer !== 'slide') continue;
          try {
            if (name === 'grpSp') {
              const g = desc(kid(el, 'grpSpPr'), 'xfrm');
              if (!g) { await renderTree(el, relsMap, layer, T); continue; }
              const off = kid(g, 'off'), ext = kid(g, 'ext'), co = kid(g, 'chOff'), ce = kid(g, 'chExt');
              const sx = num(ce, 'cx') ? num(ext, 'cx') / num(ce, 'cx') : 1;
              const sy = num(ce, 'cy') ? num(ext, 'cy') / num(ce, 'cy') : 1;
              const T2 = (b) => T({ ...b, x: num(off, 'x') + (b.x - num(co, 'x')) * sx, y: num(off, 'y') + (b.y - num(co, 'y')) * sy, w: b.w * sx, h: b.h * sy });
              await renderTree(el, relsMap, layer, T2);
            } else if (name === 'sp' || name === 'cxnSp') {
              await renderShape(el, ph, layer, T, relsMap);
            } else if (name === 'pic') {
              const xf = xfrmOf(el, ph, layer);
              const blip = desc(el, 'blip');
              if (!xf || !blip) continue;
              const box = T(xf);
              const img = await image(relsMap[attr(blip, 'embed')]?.target);
              if (!img) continue;
              if (xf.rot) page.drawImage(img, { ...rotOrigin(box, xf.rot), width: box.w, height: box.h });
              else page.drawImage(img, { x: box.x, y: H - box.y - box.h, width: box.w, height: box.h });
            } else if (name === 'graphicFrame') {
              await renderFrame(el, layer, T, relsMap);
            } else if (name === 'AlternateContent') {
              const choice = kid(el, 'Choice') || kid(el, 'Fallback');
              if (choice) await renderTree(choice, relsMap, layer, T);
            }
          } catch (e) {
            console.warn('Elemento omitido', e);
          }
        }
      };

      async function renderShape(el, ph, layer, T, relsMap) {
        const xf = xfrmOf(el, ph, layer);
        if (!xf) return;
        const box = T(xf);
        const spPr = kid(el, 'spPr');
        const style = kid(el, 'style');
        const pg = desc(spPr, 'prstGeom');
        const geom = attr(pg, 'prst') || (desc(spPr, 'custGeom') ? 'rect' : 'rect');
        const adj = {};
        for (const gd of all(pg, 'gd')) { const m = /val (-?\d+)/.exec(attr(gd, 'fmla') || ''); if (m) adj[attr(gd, 'name')] = +m[1]; }
        const fill = fillOf(spPr, style);
        const ln = kid(spPr, 'ln');
        let line = null;
        if (ln && kid(ln, 'solidFill')) line = col(kid(ln, 'solidFill'));
        else if (!(ln && kid(ln, 'noFill')) && style && num(kid(style, 'lnRef'), 'idx') > 0) line = col(kid(style, 'lnRef'));
        const lw = Math.max(0.5, num(ln, 'w', 12700) / EMU);
        const y = H - box.y - box.h;
        if (el.localName === 'cxnSp' || geom === 'line' || /Connector/.test(geom)) {
          if (line) {
            const x1 = xf.flipH ? box.x + box.w : box.x, x2 = xf.flipH ? box.x : box.x + box.w;
            const y1 = xf.flipV ? y : y + box.h, y2 = xf.flipV ? y + box.h : y;
            page.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: lw, color: toRgb(line.hex), opacity: line.alpha });
          }
        } else if (fill || line) {
          const path = geom === 'rect' || geom === 'ellipse' ? null : shapePath(geom, box.w, box.h, adj);
          const fc = fill?.color;
          const lineOpts = line ? { borderColor: toRgb(line.hex), borderWidth: lw, borderOpacity: line.alpha } : {};
          if (fill?.grad) drawGrad(fill.grad, box.x, y, box.w, box.h, path);
          if (fill?.blip) {
            const img = await image(relsMap[attr(fill.blip, 'embed')]?.target);
            if (img) page.drawImage(img, { x: box.x, y, width: box.w, height: box.h });
          }
          const colorOpts = fc ? { color: toRgb(fc.hex), opacity: fc.alpha } : {};
          if (geom === 'ellipse') page.drawEllipse({ x: box.x + box.w / 2, y: y + box.h / 2, xScale: box.w / 2, yScale: box.h / 2, ...(fill?.grad ? {} : colorOpts), ...lineOpts });
          else if (path) page.drawSvgPath(path, { x: box.x, y: y + box.h, ...(fill?.grad ? {} : colorOpts), ...lineOpts });
          else if (xf.rot && !fill?.grad) page.drawRectangle({ ...rotOrigin(box, xf.rot), width: box.w, height: box.h, ...colorOpts, ...lineOpts });
          else if (fc || line) page.drawRectangle({ x: box.x, y, width: box.w, height: box.h, ...(fill?.grad ? {} : colorOpts), ...lineOpts });
        }
        await renderText(kid(el, 'txBody'), box, ph, layer);
      }

      async function renderFrame(el, layer, T, relsMap) {
        const xf = kid(el, 'xfrm');
        if (!xf) return;
        const raw = { x: num(kid(xf, 'off'), 'x'), y: num(kid(xf, 'off'), 'y'), w: num(kid(xf, 'ext'), 'cx'), h: num(kid(xf, 'ext'), 'cy') };
        const box = T(raw);
        const gd = desc(el, 'graphicData');
        const uri = gd?.getAttribute('uri') || '';
        if (/chart/.test(uri)) {
          const c = desc(gd, 'chart');
          const target = relsMap[attr(c, 'id')]?.target;
          const cdoc = await xml(target);
          if (cdoc) pending.push(Charts.draw(page, cdoc, { x: box.x, y: H - box.y - box.h, w: box.w, h: box.h }, { fonts: bank, scheme, fontName: theme.minor }));
          return;
        }
        if (/diagram/.test(uri)) {
          // SmartArt: PowerPoint guarda una versión ya dibujada (diagramDrawing)
          let drawingPath = null;
          const ids = desc(gd, 'relIds');
          const dataPath = relsMap[attr(ids, 'dm')]?.target;
          const dataDoc = await xml(dataPath);
          const ext = desc(dataDoc, 'dataModelExt');
          if (ext && relsMap[attr(ext, 'relId')]) drawingPath = relsMap[attr(ext, 'relId')].target;
          if (!drawingPath) drawingPath = Object.values(relsMap).find((r) => /diagramDrawing$/.test(r.type))?.target;
          const dd = await xml(drawingPath);
          if (!dd) return;
          const dRels = await rels(drawingPath);
          const tree = desc(dd, 'spTree');
          await renderTree(tree, dRels, 'diagram', (b) => T({ ...b, x: raw.x + b.x, y: raw.y + b.y }));
          return;
        }
        const tbl = desc(gd, 'tbl');
        if (!tbl) return;
        const k = box.w / (raw.w / EMU || 1);
        const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => (num(g, 'w') / EMU) * k);
        let ty = box.y;
        for (const tr of kids(tbl, 'tr')) {
          let rh = (num(tr, 'h') / EMU) * k;
          let ci = 0;
          for (const tc of kids(tr, 'tc')) {
            const span = num(tc, 'gridSpan', 1);
            const w = cols.slice(ci, ci + span).reduce((a, b) => a + b, 0);
            if (attr(tc, 'vMerge') !== '1') rh = Math.max(rh, await renderText(kid(tc, 'txBody'), { x: 0, y: 0, w, h: 0 }, null, layer, true));
            ci += span;
          }
          let tx = box.x;
          ci = 0;
          for (const tc of kids(tr, 'tc')) {
            const span = num(tc, 'gridSpan', 1);
            const w = cols.slice(ci, ci + span).reduce((a, b) => a + b, 0);
            if (attr(tc, 'hMerge') !== '1' && attr(tc, 'vMerge') !== '1') {
              const tcPr = kid(tc, 'tcPr');
              const f = col(kid(tcPr, 'solidFill'));
              page.drawRectangle({ x: tx, y: H - ty - rh, width: w, height: rh, color: f ? toRgb(f.hex) : undefined, borderColor: PDFLib.rgb(0.7, 0.7, 0.72), borderWidth: 0.6 });
              await renderText(kid(tc, 'txBody'), { x: tx, y: ty, w, h: rh }, null, layer);
            }
            tx += w;
            ci += span;
          }
          ty += rh;
        }
      }

      const base = (b) => ({ ...b, x: b.x / EMU, y: b.y / EMU, w: b.w / EMU, h: b.h / EMU });
      const showMaster = attr(sDoc.documentElement, 'showMasterSp') !== '0' && attr(lDoc?.documentElement, 'showMasterSp') !== '0';
      if (showMaster && mDoc) await renderTree(desc(mDoc, 'spTree'), mRels, 'master', base);
      if (lDoc && attr(sDoc.documentElement, 'showMasterSp') !== '0') await renderTree(desc(lDoc, 'spTree'), lRels, 'layout', base);
      await renderTree(desc(sDoc, 'spTree'), sRels, 'slide', base);
    }
    await Promise.all(pending);
    return { bytes: await doc.save(), skipped, count: doc.getPageCount() };
  }

  return { toPdf };
})();
