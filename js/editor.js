'use strict';

/**
 * Editor visual de páginas: muestra una página grande y permite colocar
 * texto, imágenes, rectángulos, trazos a mano y un área de recorte.
 * Las coordenadas de los objetos se guardan en unidades del "viewport" de pdf.js
 * a escala 1 (puntos, origen arriba a la izquierda), y al exportar se convierten
 * al espacio del PDF con viewport.convertToPdfPoint (respeta rotación y CropBox).
 */
const Editor = (() => {
  const FONTS = {
    Helvetica: { css: 'Helvetica, Arial, sans-serif', std: 'Helvetica' },
    Times: { css: '"Times New Roman", Times, serif', std: 'TimesRoman' },
    Courier: { css: '"Courier New", Courier, monospace', std: 'Courier' },
  };
  const TOOL_INFO = {
    select: ['Mover', 'M5 3l14 8-6 2-3 6z'],
    text: ['Texto', 'M4 7V5h16v2M9 19h6M12 5v14'],
    image: ['Imagen', 'M3 5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2zM8.5 10a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3M21 15l-5-5L5 21'],
    rect: ['Rectángulo', 'M3 5h18v14H3z'],
    draw: ['Dibujar', 'M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z'],
    ftext: ['Campo de texto', 'M3 7h18v10H3zM7 10v4'],
    fcheck: ['Casilla', 'M4 4h16v16H4zM8 12l3 3 5-6'],
    fselect: ['Lista', 'M3 7h18v10H3zM15 11l2 2 2-2'],
  };
  const LINE_H = 1.2; // interlineado del texto
  const ASCENT = 0.87; // posición de la línea base dentro de la primera línea (en em)

  let uid = 0;

  function create({ pdf, main, side, cfg }) {
    const st = {
      page: 1,
      n: pdf.numPages,
      objs: {},
      tool: (cfg.tools || ['select'])[0],
      style: { color: cfg.color || '#1c1b29', size: 18, width: 3, font: 'Helvetica', fill: cfg.fill !== false },
      zoom: 1,
      vp: null,
      crop: cfg.crop ? { fx: 0.06, fy: 0.06, fw: 0.88, fh: 0.88 } : null,
      sel: null,
    };
    let task = null;
    let renderSeq = 0;

    const tools = cfg.tools || [];
    main.innerHTML = `
      <div class="ed">
        <div class="ed-bar">
          <button type="button" class="ed-btn" data-nav="-1" aria-label="Página anterior">‹</button>
          <span class="ed-pg"></span>
          <button type="button" class="ed-btn" data-nav="1" aria-label="Página siguiente">›</button>
          ${cfg.crop ? '' : '<button type="button" class="ed-btn ed-undo">↶ Deshacer</button>'}
        </div>
        <div class="ed-wrap">
          <div class="ed-page">
            <canvas></canvas>
            <svg class="ed-ink"></svg>
            <div class="ed-layer ${cfg.crop ? 'cropping' : ''}"></div>
          </div>
        </div>
      </div>`;
    const wrap = main.querySelector('.ed-wrap');
    const pageEl = main.querySelector('.ed-page');
    const canvas = main.querySelector('canvas');
    const ink = main.querySelector('.ed-ink');
    const layer = main.querySelector('.ed-layer');

    /* ── Panel lateral: herramientas y estilo ── */
    const has = (t) => tools.includes(t);
    side.innerHTML = `
      ${tools.length > 1 ? `<div class="ed-tools">${tools
        .map((t) => `<button type="button" data-tool="${t}" title="${TOOL_INFO[t][0]}">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="${TOOL_INFO[t][1]}"/></svg>
            <span>${TOOL_INFO[t][0]}</span></button>`)
        .join('')}</div>` : ''}
      ${has('image') ? '<input type="file" accept="image/*" hidden class="ed-imgfile">' : ''}
      ${cfg.styles === false ? '' : `<div class="ed-style">
        <div class="row">
          <div class="field"><label class="lbl">Color</label><input type="color" data-s="color" value="${st.style.color}"></div>
          ${has('text') ? `<div class="field"><label class="lbl">Tamaño texto</label><input type="number" data-s="size" min="6" max="120" value="${st.style.size}"></div>` : ''}
        </div>
        ${has('text') ? `<div class="field"><label class="lbl">Fuente</label><select data-s="font">${Object.keys(FONTS)
          .map((f) => `<option>${f}</option>`).join('')}</select></div>` : ''}
        ${has('draw') ? `<div class="field"><label class="lbl">Grosor del trazo</label><input type="range" data-s="width" min="1" max="12" value="${st.style.width}"></div>` : ''}
        ${has('rect') ? `<label class="check"><input type="checkbox" data-s="fill" ${st.style.fill ? 'checked' : ''}> Rectángulo relleno</label>` : ''}
      </div>`}`;

    function setTool(t) {
      st.tool = t;
      side.querySelectorAll('[data-tool]').forEach((b) => b.classList.toggle('active', b.dataset.tool === t));
      layer.dataset.tool = t;
      if (t === 'image') side.querySelector('.ed-imgfile').click();
    }
    side.addEventListener('click', (e) => {
      const b = e.target.closest('[data-tool]');
      if (b) setTool(b.dataset.tool);
    });
    side.addEventListener('input', (e) => {
      const k = e.target.dataset.s;
      if (!k) return;
      st.style[k] = e.target.type === 'checkbox' ? e.target.checked : e.target.type === 'color' || k === 'font' ? e.target.value : +e.target.value;
      // Aplica el estilo al objeto seleccionado
      const o = st.sel;
      if (!o) return;
      if (k === 'color') o.color = st.style.color;
      if (o.type === 'text' && (k === 'size' || k === 'font')) o[k] = st.style[k];
      if (o.type === 'rect' && k === 'fill') o.fill = st.style.fill;
      refresh(o);
    });
    side.querySelector('.ed-imgfile')?.addEventListener('change', async (e) => {
      const f = e.target.files[0];
      e.target.value = '';
      setTool('select');
      if (!f) return;
      try {
        const { url, w, h } = await imageToPng(f);
        addImage(url, w, h, 0.4);
      } catch (_) {
        alert('No se pudo leer la imagen.');
      }
    });

    /* ── Renderizado ── */
    async function render() {
      const seq = ++renderSeq;
      const page = await pdf.getPage(st.page);
      if (seq !== renderSeq) return;
      const vp1 = page.getViewport({ scale: 1 });
      st.vp = vp1;
      const avail = Math.max(240, Math.min(wrap.clientWidth - 32, 980));
      st.zoom = avail / vp1.width;
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const vp = page.getViewport({ scale: st.zoom * dpr });
      const W = vp1.width * st.zoom;
      const H = vp1.height * st.zoom;
      if (task) task.cancel();
      canvas.width = Math.floor(vp.width);
      canvas.height = Math.floor(vp.height);
      Object.assign(pageEl.style, { width: W + 'px', height: H + 'px' });
      ink.setAttribute('viewBox', `0 0 ${vp1.width} ${vp1.height}`);
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      main.querySelector('.ed-pg').textContent = `Página ${st.page} de ${st.n}`;
      drawObjects();
      task = page.render({ canvasContext: ctx, viewport: vp });
      try { await task.promise; } catch (_) { /* cancelado */ }
    }

    const list = (p = st.page) => (st.objs[p] = st.objs[p] || []);
    const find = (id) => list().find((o) => o.id === +id) || (st.cropObj && st.cropObj.id === +id ? st.cropObj : null);

    function drawObjects() {
      layer.innerHTML = '';
      ink.innerHTML = '';
      for (const o of list()) o.type === 'draw' ? drawInk(o) : layer.appendChild(makeEl(o));
      if (st.crop) {
        const c = st.crop;
        st.cropObj = { id: -1, type: 'crop', x: c.fx * st.vp.width, y: c.fy * st.vp.height, w: c.fw * st.vp.width, h: c.fh * st.vp.height };
        layer.appendChild(makeEl(st.cropObj));
      }
      if (st.sel) select(list().includes(st.sel) ? st.sel : null);
    }

    function drawInk(o) {
      const p = document.createElementNS('http://www.w3.org/2000/svg', 'polyline');
      p.setAttribute('points', o.points.map((q) => q.join(',')).join(' '));
      Object.entries({ fill: 'none', stroke: o.color, 'stroke-width': o.width, 'stroke-linecap': 'round', 'stroke-linejoin': 'round' })
        .forEach(([k, v]) => p.setAttribute(k, v));
      ink.appendChild(p);
      return p;
    }

    function makeEl(o) {
      const el = document.createElement('div');
      el.className = `ed-obj ed-${o.type}`;
      el.dataset.id = o.id;
      if (o.type === 'text') {
        el.innerHTML = `<span class="ed-grip" title="Arrastra para mover">⠿</span><div class="ed-txt" contenteditable="true" spellcheck="false"></div>`;
        const t = el.querySelector('.ed-txt');
        t.textContent = o.text;
        t.addEventListener('input', () => {
          o.text = t.innerText.replace(/\n$/, '');
          measure(o, el);
        });
      } else if (o.type === 'image') {
        el.innerHTML = `<img src="${o.url}" alt="" draggable="false">`;
      }
      if (o.type !== 'crop') el.insertAdjacentHTML('beforeend', '<button type="button" class="ed-del" title="Eliminar">×</button>');
      if (o.type === 'field') el.insertAdjacentHTML('afterbegin', '<span class="ed-flabel"></span>');
      if (o.type === 'image' || o.type === 'rect' || o.type === 'crop' || o.type === 'field') el.insertAdjacentHTML('beforeend', '<span class="ed-rs se"></span>');
      if (o.type === 'crop') el.insertAdjacentHTML('beforeend', '<span class="ed-rs nw"></span>');
      o.el = el;
      refresh(o);
      return el;
    }

    function refresh(o) {
      const el = o.el;
      if (!el) return;
      const z = st.zoom;
      el.style.left = o.x * z + 'px';
      el.style.top = o.y * z + 'px';
      if (o.type === 'text') {
        const t = el.querySelector('.ed-txt');
        Object.assign(t.style, { fontSize: o.size * z + 'px', color: o.color, fontFamily: FONTS[o.font].css, lineHeight: LINE_H });
        requestAnimationFrame(() => measure(o, el));
      } else {
        el.style.width = o.w * z + 'px';
        el.style.height = o.h * z + 'px';
      }
      if (o.type === 'field') {
        const icon = { ftext: 'Aa', fcheck: '☑', fselect: '▾' }[o.ftype];
        el.querySelector('.ed-flabel').textContent = `${icon} ${o.name}`;
      }
      if (o.type === 'rect') {
        el.style.background = o.fill ? o.color : 'transparent';
        el.style.border = o.fill ? '0' : `${Math.max(1, 2 * z)}px solid ${o.color}`;
      }
    }

    function measure(o, el) {
      const t = el.querySelector('.ed-txt');
      o.w = t.offsetWidth / st.zoom;
      o.h = t.offsetHeight / st.zoom;
    }

    function select(o) {
      st.sel = o;
      if (cfg.onSelect) cfg.onSelect(o);
      layer.querySelectorAll('.ed-obj').forEach((el) => el.classList.toggle('selected', !!o && el === o.el));
    }

    function add(o) {
      o.id = ++uid;
      list().push(o);
      if (o.type === 'draw') drawInk(o);
      else layer.insertBefore(makeEl(o), st.cropObj ? st.cropObj.el : null);
      return o;
    }

    function remove(o) {
      st.objs[st.page] = list().filter((x) => x !== o);
      if (st.sel === o) st.sel = null;
      drawObjects();
    }

    function addImage(url, natW, natH, frac = 0.3) {
      const vp = st.vp;
      const w = Math.min(vp.width * frac, natW);
      const h = (w * natH) / natW;
      const o = add({ type: 'image', url, ratio: natH / natW, x: (vp.width - w) / 2, y: (vp.height - h) / 2, w, h });
      select(o);
      return o;
    }

    /* ── Interacción ── */
    const toPage = (e) => {
      const r = layer.getBoundingClientRect();
      return { x: (e.clientX - r.left) / st.zoom, y: (e.clientY - r.top) / st.zoom };
    };
    const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

    function drag(e, o, mode, start) {
      e.preventDefault();
      const p0 = start || toPage(e);
      const o0 = { x: o.x, y: o.y, w: o.w, h: o.h };
      const W = st.vp.width;
      const H = st.vp.height;
      const move = (ev) => {
        const p = toPage(ev);
        const dx = p.x - p0.x;
        const dy = p.y - p0.y;
        if (mode === 'move') {
          o.x = clamp(o0.x + dx, -o0.w * 0.8, W - o0.w * 0.2);
          o.y = clamp(o0.y + dy, -o0.h * 0.8, H - o0.h * 0.2);
        } else if (mode === 'create') {
          const q = { x: clamp(p.x, 0, W), y: clamp(p.y, 0, H) };
          o.x = Math.min(p0.x, q.x);
          o.y = Math.min(p0.y, q.y);
          o.w = Math.abs(q.x - p0.x);
          o.h = Math.abs(q.y - p0.y);
        } else if (mode === 'se') {
          o.w = clamp(o0.w + dx, 8, W - o.x);
          o.h = o.ratio ? o.w * o.ratio : clamp(o0.h + dy, 8, H - o.y);
        } else if (mode === 'nw') {
          const nx = clamp(o0.x + dx, 0, o0.x + o0.w - 8);
          const ny = clamp(o0.y + dy, 0, o0.y + o0.h - 8);
          o.w = o0.w + (o0.x - nx);
          o.h = o0.h + (o0.y - ny);
          o.x = nx;
          o.y = ny;
        }
        if (o.type === 'crop') {
          o.x = clamp(o.x, 0, W - o.w);
          o.y = clamp(o.y, 0, H - o.h);
          st.crop = { fx: o.x / W, fy: o.y / H, fw: o.w / W, fh: o.h / H };
        }
        refresh(o);
      };
      const up = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
        if (mode === 'create' && (o.w < 4 || o.h < 4)) {
          if (o.type !== 'field') remove(o);
          else {
            // clic sin arrastrar: tamaño por defecto
            Object.assign(o, o.ftype === 'fcheck' ? { w: 14, h: 14 } : { w: 160, h: 22 });
            refresh(o);
            select(o);
          }
        } else if (mode === 'create' && o.type === 'field') select(o);
      };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
    }

    layer.addEventListener('pointerdown', (e) => {
      if (e.button) return;
      const objEl = e.target.closest('.ed-obj');
      if (e.target.closest('.ed-del')) {
        const o = find(objEl.dataset.id);
        if (o) remove(o);
        return;
      }
      if (objEl) {
        const o = find(objEl.dataset.id);
        if (!o) return;
        if (o.type !== 'crop') select(o);
        const rs = e.target.closest('.ed-rs');
        if (rs) return drag(e, o, rs.classList.contains('nw') ? 'nw' : 'se');
        if (o.type === 'text' && !e.target.closest('.ed-grip')) return; // editar el texto
        return drag(e, o, 'move');
      }
      select(null);
      const p = toPage(e);
      if (st.tool === 'text') {
        e.preventDefault();
        const o = add({ type: 'text', text: 'Escribe aquí', x: p.x, y: p.y - st.style.size * 0.6, size: st.style.size, color: st.style.color, font: st.style.font, w: 0, h: 0 });
        select(o);
        const t = o.el.querySelector('.ed-txt');
        t.focus();
        document.getSelection().selectAllChildren(t);
      } else if (/^f(text|check|select)$/.test(st.tool)) {
        st.fieldN = (st.fieldN || 0) + 1;
        const prefix = { ftext: 'texto', fcheck: 'casilla', fselect: 'lista' }[st.tool];
        const o = add({ type: 'field', ftype: st.tool, name: `${prefix}${st.fieldN}`, options: st.tool === 'fselect' ? ['Opción 1', 'Opción 2'] : [], x: p.x, y: p.y, w: 0, h: 0 });
        drag(e, o, 'create', p);
      } else if (st.tool === 'rect') {
        const o = add({ type: 'rect', x: p.x, y: p.y, w: 0, h: 0, color: st.style.color, fill: st.style.fill });
        drag(e, o, 'create', p);
      } else if (st.tool === 'draw') {
        e.preventDefault();
        const o = { type: 'draw', points: [[p.x, p.y]], color: st.style.color, width: st.style.width };
        const line = drawInk(o);
        const move = (ev) => {
          const q = toPage(ev);
          o.points.push([+q.x.toFixed(1), +q.y.toFixed(1)]);
          line.setAttribute('points', o.points.map((x) => x.join(',')).join(' '));
        };
        const up = () => {
          window.removeEventListener('pointermove', move);
          window.removeEventListener('pointerup', up);
          line.remove();
          if (o.points.length > 1) add(o);
        };
        window.addEventListener('pointermove', move);
        window.addEventListener('pointerup', up);
      }
    });

    const onKey = (e) => {
      if (!st.sel || !layer.isConnected) return;
      if (document.activeElement && document.activeElement.isContentEditable) return;
      if (/^(INPUT|SELECT|TEXTAREA)$/.test(document.activeElement?.tagName)) return;
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        remove(st.sel);
      }
    };
    document.addEventListener('keydown', onKey);

    main.querySelector('.ed-bar').addEventListener('click', (e) => {
      const nav = e.target.closest('[data-nav]');
      if (nav) {
        const p = clamp(st.page + +nav.dataset.nav, 1, st.n);
        if (p !== st.page) {
          st.page = p;
          st.sel = null;
          render();
        }
      } else if (e.target.closest('.ed-undo')) {
        const l = list();
        if (l.length) remove(l[l.length - 1]);
      }
    });

    let rt;
    const onResize = () => {
      clearTimeout(rt);
      rt = setTimeout(() => layer.isConnected && render(), 150);
    };
    window.addEventListener('resize', onResize);

    setTool(st.tool);
    render();

    return {
      st,
      addImage,
      refresh,
      select,
      remove,
      add,
      render,
      /** Todos los objetos: [{page, objs}] */
      pages: () => Object.entries(st.objs).filter(([, l]) => l.length).map(([p, l]) => ({ page: +p, objs: l })),
      count: () => Object.values(st.objs).reduce((a, l) => a + l.length, 0),
      destroy() {
        document.removeEventListener('keydown', onKey);
        window.removeEventListener('resize', onResize);
        if (task) task.cancel();
      },
    };
  }

  /** Convierte una imagen cualquiera a PNG (dataURL) para poder incrustarla. */
  async function imageToPng(file) {
    const src = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.src = src;
      await img.decode();
      const c = document.createElement('canvas');
      c.width = img.naturalWidth;
      c.height = img.naturalHeight;
      c.getContext('2d').drawImage(img, 0, 0);
      return { url: c.toDataURL('image/png'), w: c.width, h: c.height };
    } finally {
      URL.revokeObjectURL(src);
    }
  }

  /** Dibuja en el documento pdf-lib todos los objetos del editor. */
  async function apply(doc, pdf, ed) {
    const fonts = {};
    const images = new Map();
    const font = async (name) => (fonts[name] = fonts[name] || (await doc.embedFont(PDFLib.StandardFonts[FONTS[name].std])));
    for (const { page: n, objs } of ed.pages()) {
      const page = doc.getPage(n - 1);
      const vp = (await pdf.getPage(n)).getViewport({ scale: 1 });
      const rot = PDFLib.degrees(page.getRotation().angle);
      const P = (x, y) => {
        const [px, py] = vp.convertToPdfPoint(x, y);
        return { x: px, y: py };
      };
      for (const o of objs) {
        if (o.type === 'text') {
          const f = await font(o.font);
          const lines = o.text.split('\n');
          try {
            lines.forEach((line, i) => {
              if (!line) return;
              const at = P(o.x, o.y + o.size * (ASCENT + LINE_H * i));
              page.drawText(line, { ...at, size: o.size, font: f, color: U.color(o.color), rotate: rot });
            });
          } catch (e) {
            throw new Error('El texto contiene caracteres no compatibles (por ejemplo emojis).');
          }
        } else if (o.type === 'image') {
          if (!images.has(o.url)) {
            const bytes = await (await fetch(o.url)).arrayBuffer();
            images.set(o.url, await doc.embedPng(bytes));
          }
          page.drawImage(images.get(o.url), { ...P(o.x, o.y + o.h), width: o.w, height: o.h, rotate: rot });
        } else if (o.type === 'rect') {
          const c = U.color(o.color);
          page.drawRectangle({
            ...P(o.x, o.y + o.h), width: o.w, height: o.h, rotate: rot,
            ...(o.fill ? { color: c } : { borderColor: c, borderWidth: 2 }),
          });
        } else if (o.type === 'draw') {
          const c = U.color(o.color);
          for (let i = 1; i < o.points.length; i++) {
            page.drawLine({
              start: P(...o.points[i - 1]), end: P(...o.points[i]),
              thickness: o.width, color: c, lineCap: PDFLib.LineCapStyle.Round,
            });
          }
        }
      }
    }
  }

  return { create, apply, imageToPng };
})();
