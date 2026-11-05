'use strict';

(() => {
  const app = document.getElementById('app');
  const $ = (sel, el = document) => el.querySelector(sel);
  const esc = (s) =>
    String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const svg = (name, size = 24) =>
    `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name]}</svg>`;

  const CATS = [
    ['todas', 'Todas'],
    ['organizar', 'Organizar'],
    ['optimizar', 'Optimizar'],
    ['convertir', 'Convertir'],
    ['editar', 'Editar'],
    ['seguridad', 'Seguridad'],
    ['flujos', 'Flujos de trabajo'],
  ];

  let S = null; // sesión de la herramienta abierta
  const sessions = new Map(); // id de herramienta → sesión guardada (con su pantalla)
  let uid = 0;
  let homeFilter = 'todas';

  /* ───────────── utilidades de UI ───────────── */
  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 4000);
  }

  // tool.accept: 'pdf' | 'image' | { exts: [...], label, badge, color }
  const extOf = (name) => (String(name).split('.').pop() || '').toLowerCase();
  function accepted(tool, f) {
    if (typeof tool.accept === 'object') return tool.accept.exts.includes(extOf(f.name));
    return tool.accept === 'pdf'
      ? f.type === 'application/pdf' || /\.pdf$/i.test(f.name)
      : f.type.startsWith('image/') || /\.(jpe?g|png|webp|gif|bmp|avif)$/i.test(f.name);
  }
  const acceptAttr = (tool) =>
    typeof tool.accept === 'object' ? tool.accept.exts.map((e) => '.' + e).join(',')
    : tool.accept === 'pdf' ? 'application/pdf,.pdf' : 'image/*';
  const acceptLabel = (tool, plural) =>
    typeof tool.accept === 'object' ? (plural ? 'archivos: ' : '') + tool.accept.label
    : tool.accept === 'pdf' ? (plural ? 'archivos PDF' : 'archivo PDF') : 'imágenes';

  function readForm(form) {
    const o = {};
    if (!form) return o;
    for (const el of form.elements) {
      if (!el.name) continue;
      if (el.type === 'checkbox') o[el.name] = el.checked;
      else if (el.type === 'file') o[el.name] = el.files[0] || null;
      else if (el.type === 'radio') { if (el.checked) o[el.name] = el.value; }
      else o[el.name] = el.value;
    }
    return o;
  }

  /** Sale de la herramienta sin perder nada: guarda su pantalla y su estado. */
  function leaveSession() {
    if (!S) return;
    const node = app.firstElementChild;
    if (node && node.classList.contains('tool')) {
      S.node = node;
      S.scrollY = window.scrollY;
      node.remove();
    }
    // la cámara se apaga al salir (las fotos se conservan)
    if (S.custom && S.custom.pause) try { S.custom.pause(); } catch (_) { /* nada */ }
    sessions.set(S.tool.id, S);
    S = null;
  }

  /** ¿La sesión tiene trabajo del usuario? */
  const hasWork = (sess) => sess && sess.view !== 'upload' && !(sess.tool.mode === 'panel' && sess.custom && sess.custom.isEmpty && sess.custom.isEmpty());

  /** Descarta del todo la sesión actual (Empezar de nuevo). */
  function closeSession() {
    if (!S) return;
    sessions.delete(S.tool.id);
    S.token = null;
    if (S.custom && S.custom.destroy) try { S.custom.destroy(); } catch (_) { /* nada */ }
    if (S.ed) S.ed.destroy();
    if (S.pdf) S.pdf.destroy();
    S = null;
  }

  /* ───────────── router ───────────── */
  function route() {
    const id = location.hash.replace(/^#\/?/, '');
    const tool = TOOLS.find((t) => t.id === id);
    if (S && tool && S.tool === tool) return;
    leaveSession();
    closeSearch();
    window.scrollTo(0, 0);
    if (tool) openTool(tool);
    else renderHome();
  }
  window.addEventListener('hashchange', route);

  /** Etiqueta "Abierta / Procesando / Lista" para herramientas con trabajo guardado. */
  function stateBadge(t, cls) {
    const sess = sessions.get(t.id);
    if (!hasWork(sess)) return '';
    if (sess.view === 'busy') return `<span class="${cls} busy">Procesando…</span>`;
    if (sess.view === 'done') return `<span class="${cls} done">Lista ✓</span>`;
    return `<span class="${cls}">Abierta</span>`;
  }

  /* ───────────── portada ───────────── */
  function renderHome() {
    document.title = 'Converter — Herramientas PDF gratis en tu navegador';
    app.innerHTML = `
      <section class="hero">
        <h1>Todas las herramientas PDF que necesitas, en un solo sitio</h1>
        <p>Une, divide, comprime, convierte, edita, firma y protege tus PDF gratis y en segundos.</p>
        <span class="badge">🔒 Privado: todo se procesa en tu navegador, nada se sube a internet</span>
      </section>
      ${launchFiles ? `<div class="launch-banner">📂 Has abierto <b>${esc(launchFiles.map((f) => f.name).join(', '))}</b>. Elige qué quieres hacer con ${launchFiles.length > 1 ? 'ellos' : 'él'}.
        <button type="button" class="link" id="launch-x">Descartar</button></div>` : ''}
      <div class="chips" role="tablist">${CATS.map(
        ([id, label]) => `<button class="chip" data-cat="${id}" role="tab">${label}</button>`
      ).join('')}</div>
      <section class="grid" id="grid"></section>`;

    const grid = $('#grid');
    const draw = () => {
      grid.innerHTML = TOOLS.filter((t) => homeFilter === 'todas' || t.cat === homeFilter)
        .map(
          (t) => `
          <a class="card" href="#/${t.id}" style="--c:${t.color}">
            ${stateBadge(t, 'card-state')}
            <span class="card-icon">${svg(t.icon, 26)}</span>
            <h3>${t.name}</h3>
            <p>${t.desc}</p>
          </a>`
        )
        .join('');
      app.querySelectorAll('.chip').forEach((c) => c.classList.toggle('active', c.dataset.cat === homeFilter));
    };
    $('#launch-x')?.addEventListener('click', () => { launchFiles = null; renderHome(); });
    app.querySelector('.chips').addEventListener('click', (e) => {
      const c = e.target.closest('.chip');
      if (!c) return;
      homeFilter = c.dataset.cat;
      draw();
    });
    draw();
  }

  /* ───────────── herramienta ───────────── */
  function openTool(tool) {
    document.title = `${tool.name} — Converter`;
    const saved = sessions.get(tool.id);
    if (saved && saved.node) {
      // volvemos a la herramienta tal como estaba
      S = saved;
      app.replaceChildren(saved.node);
      requestAnimationFrame(() => window.scrollTo(0, saved.scrollY || 0));
      if (S.ed && S.view === 'work') S.ed.render();
      if (S.pendingFile) { const f = S.pendingFile; S.pendingFile = null; loadSingle(f); }
      return;
    }
    S = { tool, files: [], pages: [], allPages: [], view: 'upload', token: {} };
    const launched = takeLaunchFiles(tool);
    if (launched) setTimeout(() => addFiles(launched), 0);
    app.innerHTML = `<section class="tool" style="--c:${tool.color}"><div id="stage"></div></section>`;
    if (tool.mode === 'panel') return openPanel();
    showUpload();
  }

  /** Herramientas que no empiezan con un archivo (escanear, HTML a PDF…). */
  async function openPanel() {
    const t = S.tool;
    showWork();
    try {
      S.custom = await t.mount({ main: $('#custom'), toast });
    } catch (e) {
      toast(e.message);
    }
    updateSummary();
  }

  function showUpload() {
    const t = S.tool;
    S.view = 'upload';
    const what = acceptLabel(t, t.multiple);
    $('#stage').innerHTML = `
      <div class="upload">
        <span class="tool-icon">${svg(t.icon, 34)}</span>
        <h1>${t.name}</h1>
        <p class="lead">${t.desc}</p>
        <label class="pick">
          Seleccionar ${what}
          <input type="file" hidden ${t.multiple ? 'multiple' : ''} accept="${acceptAttr(t)}">
        </label>
        <p class="hint">o arrastra y suelta ${t.multiple ? 'los archivos' : 'el archivo'} aquí</p>
      </div>`;
    $('#stage input[type=file]').addEventListener('change', (e) => addFiles(e.target.files));
  }

  async function addFiles(list) {
    if (!S || !['upload', 'work'].includes(S.view)) return;
    const t = S.tool;
    let files = [...list].filter((f) => accepted(t, f));
    if (!files.length) {
      toast(`Solo se admiten ${acceptLabel(t, true)}.`);
      return;
    }
    if (files.length < list.length) toast('Se han ignorado algunos archivos de otro tipo.');
    if (!t.multiple) files = files.slice(0, 1);

    if (t.mode === 'panel') { if (t.onFiles && S.custom) t.onFiles(S.custom, files); return; }
    if (t.mode !== 'files') return loadSingle(files[0]);

    if (!t.multiple) S.files = [];
    if (S.view !== 'work') showWork();
    else if (!t.multiple) $('#items').querySelectorAll('.item').forEach((el) => el.remove());
    const items = $('#items');
    const addCard = $('.add', items);
    for (const f of files) {
      const item = makeFileItem(f);
      S.files.push(item);
      items.insertBefore(item.el, addCard);
    }
    updateSummary();
  }

  function makeFileItem(file) {
    const t = S.tool;
    const item = { id: ++uid, file };
    const el = document.createElement('div');
    el.className = 'item';
    el.draggable = !!t.reorder;
    el.dataset.id = item.id;
    el.innerHTML = `
      <button class="item-x" title="Quitar" aria-label="Quitar">×</button>
      <div class="thumb"><div class="spinner sm"></div></div>
      <div class="item-name" title="${esc(file.name)}">${esc(file.name)}</div>
      <div class="item-meta">${U.fmtSize(file.size)}</div>`;
    $('.item-x', el).addEventListener('click', () => {
      S.files = S.files.filter((x) => x !== item);
      el.remove();
      if (!S.files.length) showUpload();
      else updateSummary();
    });
    item.el = el;

    const thumb = $('.thumb', el);
    const session = S;
    const token = S.token;
    (async () => {
      try {
        if (typeof t.accept === 'object') {
          thumb.innerHTML = `<span class="file-badge" style="--b:${t.accept.color}">${esc(extOf(file.name).toUpperCase())}</span>`;
        } else if (t.accept === 'image') {
          const img = new Image();
          img.alt = '';
          img.src = URL.createObjectURL(file);
          await img.decode();
          thumb.replaceChildren(img);
        } else {
          const pdf = await U.pdfJs(await U.bytes(file));
          if (session.token !== token) return pdf.destroy();
          const { canvas } = await U.render(pdf, 1, { box: 160 });
          thumb.replaceChildren(canvas);
          $('.item-meta', el).textContent = `${U.fmtSize(file.size)} · ${pdf.numPages} pág.`;
          pdf.destroy();
        }
      } catch (e) {
        if (/contraseña/.test(e.message)) {
          thumb.innerHTML = '<span class="thumb-lock">🔒</span>';
          $('.item-meta', el).textContent = `${U.fmtSize(file.size)} · protegido`;
        } else {
          thumb.innerHTML = '<span class="thumb-err">⚠</span>';
          el.classList.add('has-error');
          el.title = e.message;
        }
      }
    })();
    return item;
  }

  /* ─── herramientas de un solo PDF: páginas, editor o personalizadas ─── */
  async function loadSingle(file) {
    const session = S;
    S.view = 'loading';
    S.token = {};
    if (S.pdf) { S.pdf.destroy(); S.pdf = null; }
    $('#stage').innerHTML = `<div class="loading"><div class="spinner"></div><p>Cargando páginas…</p></div>`;
    try {
      S.srcFile = file;
      S.bytes = await U.bytes(file);
      session.pdf = await U.pdfJs(session.bytes);
    } catch (e) {
      if (S !== session) { session.view = 'upload'; return; }
      showUpload();
      toast(e.message);
      return;
    }
    // si el usuario cambió de herramienta mientras cargaba, se termina al volver
    if (S !== session) { session.pendingFile = file; session.view = 'upload'; return; }
    const t = S.tool;
    if (S.ed) { S.ed.destroy(); S.ed = null; }
    if (t.mode === 'editor') {
      showWork();
      S.ed = Editor.create({ pdf: S.pdf, main: $('#ed'), side: $('#ed-side'), cfg: t.editor });
      t.setup && t.setup({ side: $('#ed-extra'), ed: S.ed, pdf: S.pdf, toast });
      updateSummary();
      return;
    }
    if (t.mode === 'custom') {
      showWork();
      try {
        S.custom = await t.mount({ main: $('#custom'), bytes: S.bytes, pdf: S.pdf, file });
      } catch (e) {
        showUpload();
        toast(e.message);
        return;
      }
      updateSummary();
      return;
    }
    S.allPages = Array.from({ length: S.pdf.numPages }, (_, i) => ({ id: i, index: i, rotation: 0, selected: false }));
    S.pages = S.allPages.slice();
    showWork();

    const items = $('#items');
    S.pages.forEach((p) => {
      p.el = makePageEl(p);
      items.appendChild(p.el);
    });
    updateSummary();

    const token = S.token;
    for (const p of session.allPages) {
      if (session.token !== token) return;
      try {
        const { canvas } = await U.render(session.pdf, p.index + 1, { box: 150 });
        $('.rot', p.el).replaceChildren(canvas);
      } catch (_) {
        $('.rot', p.el).innerHTML = '<span class="thumb-err">⚠</span>';
      }
    }
  }

  function makePageEl(p) {
    const cfg = S.tool.pages;
    const el = document.createElement('div');
    el.className = 'item page';
    el.draggable = !!cfg.reorder;
    el.dataset.id = p.id;
    el.innerHTML = `
      <div class="thumb"><div class="rot"><div class="spinner sm"></div></div></div>
      <div class="page-num">${p.index + 1}</div>
      ${cfg.rotate || cfg.del ? `<div class="page-actions">
        ${cfg.rotate ? `<button data-a="rot" title="Girar 90°" aria-label="Girar">${svg('rotate', 16)}</button>` : ''}
        ${cfg.del ? `<button data-a="del" title="Eliminar página" aria-label="Eliminar">${svg('trash', 16)}</button>` : ''}
      </div>` : ''}
      ${cfg.select ? '<span class="tick"></span>' : ''}`;
    return el;
  }

  const applyRot = (p) => ($('.rot', p.el).style.transform = `rotate(${p.rotation}deg)`);

  function onPageClick(e) {
    const el = e.target.closest('.item.page');
    if (!el) return;
    const cfg = S.tool.pages;
    const p = S.pages.find((x) => String(x.id) === el.dataset.id);
    if (!p) return;
    const a = e.target.closest('[data-a]')?.dataset.a;
    if (a === 'rot') {
      p.rotation = (p.rotation + 90) % 360;
      applyRot(p);
    } else if (a === 'del') {
      S.pages = S.pages.filter((x) => x !== p);
      el.remove();
    } else if (cfg.select) {
      p.selected = !p.selected;
      el.classList.toggle('selected', p.selected);
    } else if (cfg.rotate && !cfg.reorder) {
      p.rotation = (p.rotation + 90) % 360;
      applyRot(p);
    }
    updateSummary();
  }

  function pageToolbar() {
    const cfg = S.tool.pages;
    let html = '';
    if (cfg.rotateAll) {
      html += `<div class="toolbar">
        <button type="button" data-t="left">${svg('rotateLeft', 18)} Todas a la izquierda</button>
        <button type="button" data-t="right">${svg('rotate', 18)} Todas a la derecha</button>
      </div>`;
    }
    if (cfg.select) {
      html += `<div class="toolbar">
        <button type="button" data-t="all">Seleccionar todas</button>
        <button type="button" data-t="none">Ninguna</button>
      </div>
      <div class="field"><label class="lbl" for="sel-range">Seleccionar por rango</label>
        <div class="inline"><input id="sel-range" placeholder="Ej.: 1-3, 7" autocomplete="off">
        <button type="button" data-t="range">Aplicar</button></div></div>`;
    }
    if (cfg.reset) html += `<div class="toolbar"><button type="button" data-t="reset">Restablecer todo</button></div>`;
    return html;
  }

  function onToolbar(e) {
    const b = e.target.closest('[data-t]');
    if (!b) return;
    const act = b.dataset.t;
    if (act === 'left' || act === 'right') {
      S.pages.forEach((p) => {
        p.rotation = (p.rotation + (act === 'left' ? 270 : 90)) % 360;
        applyRot(p);
      });
    } else if (act === 'all' || act === 'none') {
      S.pages.forEach((p) => {
        p.selected = act === 'all';
        p.el.classList.toggle('selected', p.selected);
      });
    } else if (act === 'range') {
      try {
        const ranges = U.parseRanges($('#sel-range').value, S.allPages.length);
        S.pages.forEach((p) => {
          p.selected = ranges.some(([a, z]) => p.index + 1 >= a && p.index + 1 <= z);
          p.el.classList.toggle('selected', p.selected);
        });
      } catch (err) {
        return toast(err.message);
      }
    } else if (act === 'reset') {
      const items = $('#items');
      S.pages = S.allPages.slice();
      S.pages.forEach((p) => {
        p.rotation = 0;
        p.selected = false;
        applyRot(p);
        items.appendChild(p.el);
      });
    }
    updateSummary();
  }

  /* ─── reordenar arrastrando ─── */
  function makeSortable(container, key) {
    let dragEl = null;
    container.addEventListener('dragstart', (e) => {
      const it = e.target.closest('.item');
      if (!it || !it.draggable) return;
      dragEl = it;
      it.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', '');
    });
    container.addEventListener('dragover', (e) => {
      if (!dragEl) return;
      e.preventDefault();
      const over = e.target.closest('.item');
      if (!over || over === dragEl || !over.draggable) return;
      const r = over.getBoundingClientRect();
      const after = e.clientX - r.left > r.width / 2;
      container.insertBefore(dragEl, after ? over.nextSibling : over);
    });
    container.addEventListener('dragend', () => {
      if (!dragEl) return;
      dragEl.classList.remove('dragging');
      dragEl = null;
      const ids = [...container.querySelectorAll('.item')].map((el) => el.dataset.id);
      S[key].sort((a, b) => ids.indexOf(String(a.id)) - ids.indexOf(String(b.id)));
    });
  }

  /* ─── vista de trabajo ─── */
  function showWork() {
    const t = S.tool;
    S.view = 'work';
    const pagesMode = t.mode === 'pages';
    const single = t.mode !== 'files' && t.mode !== 'panel';
    const mainHtml =
      t.mode === 'editor' ? '<div id="ed"></div>' :
      t.mode === 'custom' || t.mode === 'panel' ? '<div id="custom" class="custom"></div>' : null;
    $('#stage').innerHTML = `
      <div class="work">
        <div class="work-main">
          ${t.hint ? `<p class="work-hint">${t.hint}</p>` : ''}
          ${mainHtml || `<div class="items ${pagesMode ? `pages sel-${t.pages.select || 'none'}` : ''}" id="items">
            ${!pagesMode && t.multiple ? `
              <label class="item add" title="Añadir más">
                <span>+</span><small>Añadir más</small>
                <input type="file" hidden multiple accept="${acceptAttr(t)}">
              </label>` : ''}
          </div>`}
        </div>
        <aside class="side">
          <h2>${t.name}</h2>
          <div class="summary" id="summary"></div>
          ${pagesMode ? `<div id="ptools">${pageToolbar()}</div>` : ''}
          ${t.mode === 'editor' ? '<div id="ed-side"></div><div id="ed-extra"></div>' : ''}
          <form id="opts" class="opts" onsubmit="return false">${t.options ? t.options() : ''}</form>
          <div class="err" id="err" hidden></div>
          <button class="go" id="go" type="button">${t.action} <span aria-hidden="true">→</span></button>
          ${t.mode !== 'panel' && (single || !t.multiple) ? `<button class="link" id="change" type="button">Elegir otro archivo</button>` : ''}
        </aside>
      </div>`;

    const items = $('#items');
    if (items && ((pagesMode && t.pages.reorder) || t.reorder)) makeSortable(items, pagesMode ? 'pages' : 'files');
    if (pagesMode) {
      items.addEventListener('click', onPageClick);
      $('#ptools').addEventListener('click', onToolbar);
    }
    items && $('.add input', items)?.addEventListener('change', (e) => {
      addFiles(e.target.files);
      e.target.value = '';
    });
    $('#change')?.addEventListener('click', () => {
      S.token = {};
      if (S.ed) { S.ed.destroy(); S.ed = null; }
      if (S.pdf) { S.pdf.destroy(); S.pdf = null; }
      S.files = [];
      showUpload();
    });

    const form = $('#opts');
    const syncShow = () => {
      const o = readForm(form);
      form.querySelectorAll('[data-show]').forEach((el) => {
        const [k, v] = el.dataset.show.split('=');
        el.hidden = o[k] !== v;
      });
    };
    form.addEventListener('input', () => { syncShow(); hideErr(); });
    form.addEventListener('change', syncShow);
    if (t.setupOptions) t.setupOptions(form, { toast });
    syncShow();
    $('#go').addEventListener('click', run);
  }

  function updateSummary() {
    const el = $('#summary');
    if (!el) return;
    const t = S.tool;
    if (t.mode === 'panel') {
      el.innerHTML = S.custom && S.custom.summary ? S.custom.summary() : '';
    } else if (t.mode === 'editor' || t.mode === 'custom') {
      const n = S.pdf.numPages;
      el.innerHTML = `<b>${esc(S.srcFile.name)}</b><br>${n} página${n === 1 ? '' : 's'}`;
    } else if (t.mode === 'pages') {
      const n = S.pages.length;
      let txt = `<b>${esc(S.srcFile.name)}</b><br>${n} página${n === 1 ? '' : 's'}`;
      if (t.pages.select) {
        const k = S.pages.filter((p) => p.selected).length;
        txt += ` · <span class="hl">${k} seleccionada${k === 1 ? '' : 's'}</span>`;
      }
      el.innerHTML = txt;
    } else {
      const n = S.files.length;
      const size = S.files.reduce((a, f) => a + f.file.size, 0);
      el.innerHTML = `${n} archivo${n === 1 ? '' : 's'} · ${U.fmtSize(size)}`;
    }
    hideErr();
  }

  function showErr(msg, stage = document) {
    const e = $('#err', stage);
    e.textContent = msg;
    e.hidden = false;
  }
  function hideErr() {
    const e = $('#err');
    if (e) e.hidden = true;
  }

  /* ─── procesar ─── */
  async function run() {
    const t = S.tool;
    const session = S;
    const ctx = {
      files: S.files,
      file: S.srcFile,
      bytes: S.bytes,
      pages: S.pages,
      pdf: S.pdf,
      ed: S.ed,
      custom: S.custom,
      form: $('#opts'),
      opts: readForm($('#opts')),
      progress: () => {},
    };
    const msg = t.validate && t.validate(ctx);
    if (msg) return showErr(msg);
    hideErr();

    const stage = $('#stage');
    const work = stage.firstElementChild;
    const busy = document.createElement('div');
    busy.className = 'busy';
    busy.innerHTML = `<div class="busy-box"><div class="spinner"></div>
      <p id="busy-text">Procesando…</p><div class="bar"><span id="bar"></span></div></div>`;
    stage.appendChild(busy);
    session.view = 'busy';
    ctx.progress = (frac, text) => {
      $('#bar', busy).style.width = `${Math.round(Math.min(1, frac) * 100)}%`;
      if (text) $('#busy-text', busy).textContent = text;
    };
    // Deja que el navegador pinte el overlay antes de empezar el trabajo pesado.
    await new Promise((r) => setTimeout(r, 30));

    try {
      let res = await t.run(ctx);
      if (Array.isArray(res)) res = { files: res };
      let blob, name;
      if (res.files.length === 1) {
        ({ blob, name } = res.files[0]);
      } else {
        ctx.progress(1, 'Empaquetando en un ZIP…');
        blob = await U.zip(res.files);
        const base = U.baseName(session.srcFile?.name || session.files[0]?.file.name || 'Converter');
        name = `${base}_${t.id}.zip`;
      }
      busy.remove();
      showDone(session, stage, work, res, blob, name);
      if (S !== session) toast(`«${t.name}» ha terminado. Tienes el resultado al volver a la herramienta.`);
    } catch (e) {
      console.error(e);
      busy.remove();
      session.view = 'work';
      showErr(e.message || 'Ha ocurrido un error inesperado.', stage);
      if (S !== session) toast(`«${t.name}» no ha podido terminar. Vuelve a la herramienta para ver el motivo.`);
    }
  }

  function showDone(session, stage, work, res, blob, name) {
    const t = session.tool;
    session.view = 'done';
    work.remove();
    const done = document.createElement('div');
    done.className = 'done';
    const ext = name.split('.').pop().toUpperCase();
    done.innerHTML = `
      <div class="done-check">✓</div>
      <h1>${esc(res.title || '¡Listo! Tu archivo está preparado')}</h1>
      ${res.message ? `<p class="lead">${esc(res.message)}</p>` : ''}
      <button class="pick" id="dl">Descargar ${ext} <span aria-hidden="true">↓</span></button>
      <p class="hint">${esc(name)} · ${U.fmtSize(blob.size)}</p>
      ${res.html ? `<div class="preview">${res.html}</div>` : ''}
      <div class="done-actions">
        <button class="link" id="back">← Volver a editar</button>
        <button class="link" id="again">Empezar de nuevo</button>
        <a class="link" href="#/">Otras herramientas</a>
      </div>`;
    stage.appendChild(done);
    const dl = () => U.download(blob, name);
    $('#dl', done).addEventListener('click', dl);
    $('#back', done).addEventListener('click', () => {
      done.remove();
      stage.appendChild(work);
      session.view = 'work';
    });
    $('#again', done).addEventListener('click', () => {
      closeSession();
      openTool(t);
    });
    // descarga automática solo si el usuario sigue en la herramienta
    if (S === session) dl();
  }

  /* ─── arrastrar archivos a la página ─── */
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  let dragDepth = 0;
  document.addEventListener('dragenter', (e) => {
    if (!hasFiles(e) || !S || !['upload', 'work'].includes(S.view)) return;
    dragDepth++;
    document.body.classList.add('dragging');
  });
  document.addEventListener('dragleave', (e) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) document.body.classList.remove('dragging');
  });
  document.addEventListener('dragover', (e) => {
    if (hasFiles(e)) e.preventDefault();
  });
  document.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    document.body.classList.remove('dragging');
    if (S) addFiles(e.dataTransfer.files);
    else toast('Elige primero una herramienta.');
  });

  /* ───────────── buscador de herramientas ───────────── */
  const KEYWORDS = {
    'unir-pdf': 'juntar combinar fusionar agrupar', 'dividir-pdf': 'separar partir cortar trocear', 'comprimir-pdf': 'reducir tamano peso optimizar achicar',
    'eliminar-paginas': 'borrar quitar paginas', 'extraer-paginas': 'sacar separar paginas', 'organizar-pdf': 'ordenar reordenar mover paginas',
    'pdf-a-jpg': 'imagen png foto convertir', 'jpg-a-pdf': 'imagen foto png webp convertir', 'pdf-a-texto': 'txt texto extraer copiar',
    'rotar-pdf': 'girar voltear', 'rotar-varios-pdf': 'girar voltear lote muchos', 'marca-de-agua': 'sello logo confidencial borrador',
    'numeros-de-pagina': 'numerar paginacion', 'editar-pdf': 'escribir anadir texto dibujar anotar', 'recortar-pdf': 'margenes cortar',
    'formularios-pdf': 'rellenar completar formulario', 'crear-formulario': 'campos rellenables formulario', 'firmar-pdf': 'firma firmar rubrica',
    'censurar-pdf': 'tachar ocultar tapar privacidad', 'proteger-pdf': 'contrasena cifrar bloquear clave seguridad', 'desbloquear-pdf': 'quitar contrasena desproteger clave',
    'comparar-pdf': 'diferencias versiones cambios', 'ocr-pdf': 'escaneado reconocer texto buscar', 'pdf-a-markdown': 'md markdown texto ia',
    'word-a-pdf': 'docx doc word convertir', 'excel-a-pdf': 'xlsx xls hoja calculo csv convertir', 'powerpoint-a-pdf': 'pptx ppt presentacion diapositivas',
    'pdf-a-word': 'docx editable word', 'pdf-a-excel': 'xlsx tablas hoja calculo', 'pdf-a-powerpoint': 'pptx presentacion diapositivas',
    'pdf-a-pdfa': 'archivar archivo iso norma', 'reparar-pdf': 'arreglar danado corrupto roto recuperar', 'html-a-pdf': 'web pagina codigo',
    'escanear-pdf': 'camara foto movil escaner', 'flujos-de-trabajo': 'automatizar lote encadenar varios pasos',
  };
  const norm = (x) => String(x).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const qInput = $('#q');
  const qRes = $('#qres');
  let qSel = 0;
  let qList = [];
  function searchTools(q) {
    const words = norm(q).split(/\s+/).filter(Boolean);
    if (!words.length) return [];
    return TOOLS.map((t) => {
      const name = norm(t.name), hay = `${name} ${norm(t.desc)} ${t.cat} ${KEYWORDS[t.id] || ''}`;
      if (!words.every((w) => hay.includes(w))) return null;
      const score = words.reduce((a, w) => a + (name.startsWith(w) ? 6 : name.includes(w) ? 4 : (KEYWORDS[t.id] || '').includes(w) ? 2 : 1), 0);
      return { t, score };
    }).filter(Boolean).sort((a, b) => b.score - a.score).slice(0, 8).map((x) => x.t);
  }
  function drawSearch() {
    qList = searchTools(qInput.value);
    if (!qInput.value.trim()) { closeSearch(); return; }
    qSel = Math.min(qSel, Math.max(0, qList.length - 1));
    qRes.innerHTML = qList.length
      ? qList.map((t, i) => `<a class="qitem ${i === qSel ? 'active' : ''}" href="#/${t.id}" role="option" style="--c:${t.color}">
          <span class="qi">${svg(t.icon, 18)}</span><span><b>${t.name}</b><small>${t.desc}</small></span>${stateBadge(t, 'qstate')}</a>`).join('')
      : '<div class="qempty">No hay ninguna herramienta con ese nombre.</div>';
    qRes.hidden = false;
  }
  function closeSearch() {
    if (!qRes) return;
    qRes.hidden = true;
    qSel = 0;
  }
  qInput.addEventListener('input', () => { qSel = 0; drawSearch(); });
  qInput.addEventListener('focus', () => { if (qInput.value.trim()) drawSearch(); });
  qInput.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (!qList.length) return;
      qSel = (qSel + (e.key === 'ArrowDown' ? 1 : qList.length - 1)) % qList.length;
      drawSearch();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (qList[qSel]) { location.hash = '#/' + qList[qSel].id; qInput.value = ''; qInput.blur(); closeSearch(); }
    } else if (e.key === 'Escape') {
      qInput.value = '';
      closeSearch();
      qInput.blur();
    }
  });
  qRes.addEventListener('click', (e) => { if (e.target.closest('.qitem')) { qInput.value = ''; closeSearch(); } });
  document.addEventListener('pointerdown', (e) => { if (!e.target.closest('.search')) closeSearch(); });
  document.addEventListener('keydown', (e) => {
    if (e.key !== '/' || e.metaKey || e.ctrlKey) return;
    const a = document.activeElement;
    if (a && (/^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName) || a.isContentEditable)) return;
    e.preventDefault();
    qInput.focus();
  });

  /* ───────────── aplicación instalable (PWA) ───────────── */
  let launchFiles = null;
  /** Si hay archivos abiertos desde el sistema que esta herramienta acepta, se los entrega. */
  function takeLaunchFiles(tool) {
    if (!launchFiles) return null;
    const ok = launchFiles.filter((f) => accepted(tool, f));
    if (!ok.length) return null;
    launchFiles = null;
    return ok;
  }
  // Abrir archivos con Converter desde el explorador de archivos (app instalada)
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async (params) => {
      if (!params.files || !params.files.length) return;
      const files = await Promise.all(params.files.map((h) => h.getFile()));
      const ext = (f) => extOf(f.name);
      const direct = files.every((f) => /^(jpe?g|png|webp)$/.test(ext(f))) ? 'jpg-a-pdf'
        : files.length === 1 && ext(files[0]) === 'docx' ? 'word-a-pdf'
        : files.length === 1 && ext(files[0]) === 'xlsx' ? 'excel-a-pdf'
        : files.length === 1 && ext(files[0]) === 'pptx' ? 'powerpoint-a-pdf'
        : files.length === 1 && /^html?$/.test(ext(files[0])) ? 'html-a-pdf'
        : files.length > 1 && files.every((f) => ext(f) === 'pdf') ? 'unir-pdf' : null;
      launchFiles = files;
      if (direct) { sessions.delete(direct); if (S && S.tool.id === direct) closeSession(); location.hash = '#/' + direct; route(); }
      else { location.hash = '#/'; route(); }
    });
  }
  // Service worker: instalación y uso sin conexión (solo en https o localhost)
  const secure = location.protocol === 'https:' || ['localhost', '127.0.0.1', '[::1]'].includes(location.hostname);
  if ('serviceWorker' in navigator && secure) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Service worker no registrado', e)));
  }
  // Botón "Instalar app"
  let installPrompt = null;
  const installBtn = $('#install');
  const standalone = () => window.matchMedia('(display-mode: standalone)').matches || navigator.standalone;
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    installPrompt = e;
    if (!standalone()) installBtn.hidden = false;
  });
  installBtn.addEventListener('click', async () => {
    if (!installPrompt) return;
    installPrompt.prompt();
    const { outcome } = await installPrompt.userChoice;
    installPrompt = null;
    installBtn.hidden = true;
    if (outcome === 'accepted') toast('Converter se está instalando.');
  });
  window.addEventListener('appinstalled', () => { installBtn.hidden = true; toast('¡Converter instalada! Ya puedes abrirla como cualquier aplicación.'); });
  if (standalone()) document.documentElement.classList.add('standalone');

  window.ConverterUI = {
    refreshSummary: () => S && updateSummary(),
    toast,
    /** Para pruebas: descarta el trabajo guardado de una herramienta (o de todas). */
    discard(id) { if (id) sessions.delete(id); else sessions.clear(); },
    savedIds: () => [...sessions.keys()],
  };
  route();
})();
