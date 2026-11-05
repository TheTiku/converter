'use strict';

/**
 * Dibuja gráficos de Office (DrawingML Charts) con pdf-lib:
 * columnas/barras (agrupadas, apiladas, 100%), líneas, áreas, tarta, anillo y dispersión.
 */
const Charts = (() => {
  const text = (el) => X.all(el, 't').map((t) => t.textContent).join('');

  function cache(el) {
    if (!el) return [];
    const multi = X.desc(el, 'multiLvlStrCache');
    if (multi) {
      // categorías multinivel: usamos el nivel más detallado (el primero)
      const n = X.num(X.kid(multi, 'ptCount'), 'val', 0);
      const arr = Array(n).fill(null);
      const lvl = X.kid(multi, 'lvl');
      for (const pt of X.kids(lvl, 'pt')) arr[+pt.getAttribute('idx')] = X.kid(pt, 'v')?.textContent ?? null;
      return arr;
    }
    const c = X.desc(el, 'strCache') || X.desc(el, 'numCache') || X.desc(el, 'strLit') || X.desc(el, 'numLit');
    if (!c) {
      const v = X.desc(el, 'v');
      return v ? [v.textContent] : [];
    }
    const n = X.num(X.kid(c, 'ptCount'), 'val', 0);
    const arr = Array(n).fill(null);
    for (const pt of X.kids(c, 'pt')) arr[+pt.getAttribute('idx')] = X.kid(pt, 'v')?.textContent ?? null;
    return arr;
  }

  function niceScale(min, max, ticks = 5) {
    if (min === max) { max = min + 1; if (min > 0) min = 0; }
    const span = max - min;
    const raw = span / ticks;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => span / s <= ticks + 0.5) || 10 * mag;
    return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
  }

  const fmt = (v, step, pct) => {
    if (pct) return `${Math.round(v * 100)}%`;
    const dec = step < 1 ? Math.min(3, Math.ceil(-Math.log10(step))) : 0;
    return v.toLocaleString('es-ES', { minimumFractionDigits: dec, maximumFractionDigits: dec });
  };

  async function draw(page, chartDoc, box, { fonts, scheme, fontName = 'Calibri' }) {
    const font = await fonts.get(fontName, false, false);
    const bold = await fonts.get(fontName, true, false);
    const chart = X.desc(chartDoc, 'chart');
    const plot = X.desc(chart, 'plotArea');
    if (!plot) return;
    const palette = [1, 2, 3, 4, 5, 6].map((i) => scheme['accent' + i]);
    const pal = (i) => {
      const base = palette[i % 6];
      const round = Math.floor(i / 6);
      if (!round) return base;
      const [r, g, b] = hexRgb(base).map((v) => (round % 2 ? v * 0.6 : v + (1 - v) * 0.4));
      return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
    };
    const colorOf = (spPr, line = false) => {
      if (!spPr) return null;
      const el = line ? X.kid(X.kid(spPr, 'ln'), 'solidFill') : X.kid(spPr, 'solidFill') || X.desc(X.kid(spPr, 'gradFill'), 'gs');
      return ooxmlColor(el, scheme)?.hex || null;
    };

    // Fondo del gráfico
    const bgFill = colorOf(X.kid(X.desc(chartDoc, 'chartSpace') || chartDoc.documentElement, 'spPr'));
    if (bgFill) page.drawRectangle({ x: box.x, y: box.y, width: box.w, height: box.h, color: toRgb(bgFill) });

    const groups = X.kids(plot).filter((e) => /Chart$/.test(e.localName));
    if (!groups.length) return;
    let serIdx = 0;
    const sets = groups.map((g) => {
      const type = g.localName.replace('3D', '').replace('Chart', '');
      const series = X.kids(g, 'ser').map((s) => {
        const i = serIdx++;
        const dPt = {};
        for (const d of X.kids(s, 'dPt')) dPt[X.num(X.kid(d, 'idx'), 'val')] = colorOf(X.kid(d, 'spPr'));
        const spPr = X.kid(s, 'spPr');
        return {
          name: cache(X.kid(s, 'tx')).join(' ') || `Serie ${i + 1}`,
          cats: cache(X.kid(s, 'cat') || X.kid(s, 'xVal')),
          vals: cache(X.kid(s, 'val') || X.kid(s, 'yVal')).map((v) => (v == null || v === '' || isNaN(+v) ? null : +v)),
          color: (type === 'line' || type === 'scatter' ? colorOf(spPr, true) : null) || colorOf(spPr) || pal(i),
          dPt,
        };
      });
      return {
        type,
        series,
        barDir: X.attr(X.kid(g, 'barDir'), 'val') || 'col',
        grouping: X.attr(X.kid(g, 'grouping'), 'val') || 'clustered',
        gap: X.num(X.kid(g, 'gapWidth'), 'val', 150) / 100,
        hole: type === 'doughnut' ? X.num(X.kid(g, 'holeSize'), 'val', 50) / 100 : 0,
      };
    });

    const base = Math.max(6, Math.min(11, Math.min(box.w, box.h) / 28));
    let top = box.y + box.h - 6;
    const left = box.x + 6, right = box.x + box.w - 6, bottom = box.y + 6;

    // Título
    const titleEl = X.kid(chart, 'title');
    const autoDel = X.attr(X.kid(chart, 'autoTitleDeleted'), 'val') === '1';
    let title = titleEl ? text(titleEl) : '';
    const allSeries = sets.flatMap((s) => s.series);
    if (titleEl && !title && allSeries.length === 1) title = allSeries[0].name;
    if (!titleEl && !autoDel && allSeries.length === 1 && sets[0].type === 'pie') title = allSeries[0].name;
    if (title) {
      const ts = base * 1.4;
      const t = cleanText(title, bold);
      const tw = bold.widthOfTextAtSize(t, ts);
      page.drawText(t, { x: box.x + Math.max(4, (box.w - tw) / 2), y: top - ts, size: ts, font: bold, color: toRgb('404040') });
      top -= ts * 1.8;
    }

    // Leyenda
    const pie = sets.find((s) => s.type === 'pie' || s.type === 'doughnut');
    const legendEl = X.kid(chart, 'legend');
    const legendPos = legendEl ? X.attr(X.kid(legendEl, 'legendPos'), 'val') || 'r' : null;
    const entries = pie
      ? (pie.series[0]?.cats || []).map((c, i) => ({ name: c ?? `${i + 1}`, color: pie.series[0].dPt[i] || pal(i) }))
      : allSeries.map((s) => ({ name: s.name, color: s.color }));
    let plotL = left, plotR = right, plotB = bottom, plotT = top;
    if (legendPos && entries.length) {
      const ls = base * 0.95;
      if (legendPos === 'r' || legendPos === 'l') {
        const lw = Math.min(box.w * 0.35, Math.max(...entries.map((e) => font.widthOfTextAtSize(cleanText(String(e.name), font), ls))) + ls * 2.5);
        const lx = legendPos === 'r' ? right - lw : left;
        let ly = (plotT + plotB) / 2 + (entries.length * ls * 1.6) / 2;
        for (const e of entries) {
          page.drawRectangle({ x: lx, y: ly - ls * 0.85, width: ls * 0.8, height: ls * 0.8, color: toRgb(e.color) });
          page.drawText(fitText(String(e.name), font, ls, lw - ls * 1.4), { x: lx + ls * 1.2, y: ly - ls * 0.8, size: ls, font, color: toRgb('404040') });
          ly -= ls * 1.6;
        }
        if (legendPos === 'r') plotR = lx - 8; else plotL = lx + lw + 8;
      } else {
        const ly = legendPos === 't' ? plotT - ls : bottom;
        const widths = entries.map((e) => font.widthOfTextAtSize(cleanText(String(e.name), font), ls) + ls * 2.4);
        let lx = box.x + Math.max(4, (box.w - widths.reduce((a, b) => a + b, 0)) / 2);
        entries.forEach((e, i) => {
          if (lx + widths[i] > right) return;
          page.drawRectangle({ x: lx, y: ly, width: ls * 0.8, height: ls * 0.8, color: toRgb(e.color) });
          page.drawText(cleanText(String(e.name), font), { x: lx + ls * 1.2, y: ly + ls * 0.05, size: ls, font, color: toRgb('404040') });
          lx += widths[i];
        });
        if (legendPos === 't') plotT -= ls * 2; else plotB += ls * 2;
      }
    }

    if (pie) return drawPie(page, pie, { l: plotL, r: plotR, b: plotB, t: plotT }, pal);

    // Ejes (categorías / valores)
    const nCat = Math.max(...allSeries.map((s) => Math.max(s.vals.length, s.cats.length)), 1);
    const cats = (allSeries.find((s) => s.cats.length)?.cats || []).slice();
    const horizontal = sets.some((s) => s.type === 'bar' && s.barDir === 'bar');
    let vmin = 0, vmax = 0;
    const pct = sets.some((s) => s.grouping === 'percentStacked');
    for (const s of sets) {
      if (s.type === 'scatter') continue;
      const stacked = s.grouping === 'stacked' || s.grouping === 'percentStacked';
      for (let i = 0; i < nCat; i++) {
        if (stacked) {
          let pos = 0, neg = 0, tot = 0;
          for (const se of s.series) { const v = se.vals[i] || 0; if (v > 0) pos += v; else neg += v; tot += Math.abs(v); }
          if (s.grouping === 'percentStacked') { pos = tot ? pos / tot : 0; neg = tot ? neg / tot : 0; }
          vmax = Math.max(vmax, pos); vmin = Math.min(vmin, neg);
        } else for (const se of s.series) if (se.vals[i] != null) { vmax = Math.max(vmax, se.vals[i]); vmin = Math.min(vmin, se.vals[i]); }
      }
    }
    const scatter = sets.find((s) => s.type === 'scatter');
    if (scatter) for (const se of scatter.series) for (const v of se.vals) if (v != null) { vmax = Math.max(vmax, v); vmin = Math.min(vmin, v); }
    const sc = niceScale(vmin, vmax);
    const ls = base * 0.9;
    const valLabels = [];
    for (let v = sc.min; v <= sc.max + sc.step / 2; v += sc.step) valLabels.push(v);
    const labW = Math.max(...valLabels.map((v) => font.widthOfTextAtSize(fmt(v, sc.step, pct), ls)));
    const grid = toRgb('D9D9D9');
    const axisCol = toRgb('595959');
    const P = horizontal
      ? { l: plotL + Math.min(box.w * 0.3, Math.max(...cats.map((c) => font.widthOfTextAtSize(cleanText(String(c ?? ''), font), ls)), 10)) + 6, r: plotR - 4, b: plotB + ls * 1.8, t: plotT - 4 }
      : { l: plotL + labW + 6, r: plotR - 4, b: plotB + ls * 2, t: plotT - 4 };
    if (P.r - P.l < 20 || P.t - P.b < 20) return;
    const vPos = (v) => (horizontal ? P.l + ((v - sc.min) / (sc.max - sc.min)) * (P.r - P.l) : P.b + ((v - sc.min) / (sc.max - sc.min)) * (P.t - P.b));

    // Cuadrícula y etiquetas de valores
    for (const v of valLabels) {
      const p = vPos(v);
      const lab = fmt(v, sc.step, pct);
      if (horizontal) {
        page.drawLine({ start: { x: p, y: P.b }, end: { x: p, y: P.t }, thickness: 0.5, color: grid });
        page.drawText(lab, { x: p - font.widthOfTextAtSize(lab, ls) / 2, y: P.b - ls * 1.4, size: ls, font, color: axisCol });
      } else {
        page.drawLine({ start: { x: P.l, y: p }, end: { x: P.r, y: p }, thickness: 0.5, color: grid });
        page.drawText(lab, { x: P.l - 4 - font.widthOfTextAtSize(lab, ls), y: p - ls * 0.35, size: ls, font, color: axisCol });
      }
    }
    // Etiquetas de categorías
    const band = (horizontal ? P.t - P.b : P.r - P.l) / nCat;
    const every = Math.max(1, Math.ceil((horizontal ? ls * 1.3 : Math.max(...cats.map((c) => font.widthOfTextAtSize(cleanText(String(c ?? ''), font), ls)), 1) + 6) / band));
    for (let i = 0; i < nCat; i += every) {
      const lab = String(cats[i] ?? i + 1);
      if (horizontal) {
        const y = P.t - band * (i + 0.5);
        const t = fitText(lab, font, ls, P.l - plotL - 6);
        page.drawText(t, { x: P.l - 4 - font.widthOfTextAtSize(t, ls), y: y - ls * 0.35, size: ls, font, color: axisCol });
      } else {
        const x = P.l + band * (i + 0.5);
        const t = fitText(lab, font, ls, band * every - 2);
        page.drawText(t, { x: x - font.widthOfTextAtSize(t, ls) / 2, y: P.b - ls * 1.3, size: ls, font, color: axisCol });
      }
    }
    const zero = vPos(Math.max(sc.min, Math.min(sc.max, 0)));
    if (horizontal) page.drawLine({ start: { x: zero, y: P.b }, end: { x: zero, y: P.t }, thickness: 0.8, color: toRgb('BFBFBF') });
    else page.drawLine({ start: { x: P.l, y: zero }, end: { x: P.r, y: zero }, thickness: 0.8, color: toRgb('BFBFBF') });

    // Series
    for (const s of sets) {
      if (s.type === 'bar') {
        const stacked = s.grouping !== 'clustered' && s.grouping !== 'standard';
        const n = stacked ? 1 : s.series.length;
        const bw = band / (n + s.gap);
        for (let i = 0; i < nCat; i++) {
          let pos = 0, neg = 0;
          const tot = s.series.reduce((a, se) => a + Math.abs(se.vals[i] || 0), 0);
          s.series.forEach((se, k) => {
            let v = se.vals[i];
            if (v == null) return;
            if (s.grouping === 'percentStacked') v = tot ? v / tot : 0;
            let a, b;
            if (stacked) { if (v >= 0) { a = pos; pos += v; b = pos; } else { a = neg; neg += v; b = neg; } } else { a = 0; b = v; }
            const off = band * i + (band - bw * n) / 2 + (stacked ? 0 : k * bw);
            const color = toRgb(se.dPt[i] || se.color);
            const p1 = vPos(Math.max(sc.min, Math.min(a, b))), p2 = vPos(Math.min(sc.max, Math.max(a, b)));
            if (horizontal) page.drawRectangle({ x: p1, y: P.t - off - bw, width: Math.max(0.5, p2 - p1), height: bw * 0.98, color });
            else page.drawRectangle({ x: P.l + off, y: p1, width: bw * 0.98, height: Math.max(0.5, p2 - p1), color });
          });
        }
      } else if (s.type === 'line' || s.type === 'area' || s.type === 'radar') {
        const stacked = s.grouping === 'stacked' || s.grouping === 'percentStacked';
        const acc = Array(nCat).fill(0);
        const layers = s.series.map((se) => {
          const pts = [];
          for (let i = 0; i < nCat; i++) {
            let v = se.vals[i];
            if (v == null) { pts.push(null); continue; }
            if (stacked) { acc[i] += v; v = acc[i]; }
            pts.push({ x: P.l + band * (i + 0.5), y: vPos(v), base: stacked ? vPos(acc[i] - se.vals[i]) : zero });
          }
          return { se, pts };
        });
        for (const { se, pts } of (s.type === 'area' ? layers.slice().reverse() : layers)) {
          const color = toRgb(se.color);
          const ok = pts.filter(Boolean);
          if (s.type === 'area' && ok.length > 1) {
            const topY = box.y + box.h;
            const d = `M ${ok[0].x - box.x} ${topY - ok[0].base} ` + ok.map((p) => `L ${p.x - box.x} ${topY - p.y}`).join(' ') +
              ` L ${ok[ok.length - 1].x - box.x} ${topY - ok[ok.length - 1].base} ` + ok.slice().reverse().map((p) => `L ${p.x - box.x} ${topY - p.base}`).join(' ') + ' Z';
            page.drawSvgPath(d, { x: box.x, y: topY, color, opacity: 0.85 });
          } else {
            for (let i = 1; i < pts.length; i++) {
              if (pts[i - 1] && pts[i]) page.drawLine({ start: pts[i - 1], end: pts[i], thickness: Math.max(1.2, base / 5), color });
            }
            for (const p of ok) page.drawCircle({ x: p.x, y: p.y, size: Math.max(1.5, base / 4), color });
          }
        }
      } else if (s.type === 'scatter') {
        const xs = s.series.flatMap((se) => se.cats.map(Number)).filter((v) => !isNaN(v));
        const xsc = xs.length ? niceScale(Math.min(...xs), Math.max(...xs)) : { min: 0, max: 1, step: 1 };
        for (const se of s.series) {
          const color = toRgb(se.color);
          const pts = se.vals.map((v, i) => (v == null || isNaN(+se.cats[i]) ? null : { x: P.l + ((+se.cats[i] - xsc.min) / (xsc.max - xsc.min)) * (P.r - P.l), y: vPos(v) }));
          pts.forEach((p) => p && page.drawCircle({ x: p.x, y: p.y, size: Math.max(1.8, base / 3.5), color }));
        }
      }
    }
    // Eje
    page.drawLine({ start: { x: P.l, y: P.b }, end: horizontal ? { x: P.l, y: P.t } : { x: P.r, y: P.b }, thickness: 0.8, color: toRgb('BFBFBF') });
  }

  function drawPie(page, s, R, pal) {
    const se = s.series[0];
    if (!se) return;
    const vals = se.vals.map((v) => Math.max(0, v || 0));
    const tot = vals.reduce((a, b) => a + b, 0);
    if (!tot) return;
    const r = Math.max(4, Math.min(R.r - R.l, R.t - R.b) / 2 - 4);
    const cx = (R.l + R.r) / 2, cy = (R.b + R.t) / 2;
    let ang = -Math.PI / 2;
    vals.forEach((v, i) => {
      if (!v) return;
      const color = toRgb(se.dPt[i] || pal(i));
      if (v >= tot) { page.drawCircle({ x: cx, y: cy, size: r, color }); return; }
      const a2 = ang + (v / tot) * Math.PI * 2;
      // drawSvgPath usa coordenadas con Y hacia abajo desde (x, y)
      const p = (a) => `${(r * Math.cos(a)).toFixed(2)} ${(r * Math.sin(a)).toFixed(2)}`;
      const large = a2 - ang > Math.PI ? 1 : 0;
      page.drawSvgPath(`M 0 0 L ${p(ang)} A ${r} ${r} 0 ${large} 1 ${p(a2)} Z`, { x: cx, y: cy, color, borderColor: PDFLib.rgb(1, 1, 1), borderWidth: 1 });
      ang = a2;
    });
    if (s.hole) page.drawCircle({ x: cx, y: cy, size: r * s.hole, color: PDFLib.rgb(1, 1, 1) });
  }

  function fitText(t, font, size, w) {
    t = cleanText(t, font);
    if (font.widthOfTextAtSize(t, size) <= w) return t;
    while (t.length > 1 && font.widthOfTextAtSize(t + '…', size) > w) t = t.slice(0, -1);
    return cleanText(t + '…', font);
  }

  return { draw };
})();
