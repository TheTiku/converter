'use strict';

/* Iconos de las herramientas nuevas */
Object.assign(ICONS, {
  sign: '<path d="M3 17c3-1 4-6 6-6s1 5 3 5 2-3 4-3 2 2 5 2"/><path d="M3 21h18"/>',
  edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',
  redact: '<rect x="3" y="5" width="18" height="4" rx="1"/><rect x="3" y="11" width="11" height="4" rx="1"/><path d="M3 19h14"/>',
  crop: '<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>',
  form: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h4"/><rect x="7" y="15" width="10" height="3" rx="1"/>',
  markdown: '<rect x="2" y="5" width="20" height="14" rx="2"/><path d="M6 15V9l3 3 3-3v6M17 9v6M15 13l2 2 2-2"/>',
  ocr: '<path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2"/><path d="M7 9h10M7 12h10M7 15h6"/>',
  compare: '<rect x="3" y="3" width="7" height="18" rx="1"/><rect x="14" y="3" width="7" height="18" rx="1"/><path d="M5.5 8h2M16.5 8h2M16.5 12h2M5.5 12h2"/>',
  lock: '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
  unlock: '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 7.5-2"/>',
});

const escHtml = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

/** Recorta los bordes transparentes de un canvas. */
function trimCanvas(c) {
  const ctx = c.getContext('2d');
  const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height);
  let x0 = width, y0 = height, x1 = -1, y1 = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] > 8) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  }
  if (x1 < 0) return null;
  const pad = 6;
  x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad);
  x1 = Math.min(width - 1, x1 + pad); y1 = Math.min(height - 1, y1 + pad);
  const out = document.createElement('canvas');
  out.width = x1 - x0 + 1;
  out.height = y1 - y0 + 1;
  out.getContext('2d').drawImage(c, x0, y0, out.width, out.height, 0, 0, out.width, out.height);
  return out;
}

/** Guarda el documento con los objetos del editor dibujados encima. */
async function saveEditor({ bytes, pdf, ed, file }, suffix) {
  const doc = await U.pdfLib(bytes);
  await Editor.apply(doc, pdf, ed);
  return [{ name: `${U.baseName(file.name)}_${suffix}.pdf`, blob: U.pdfBlob(await doc.save()) }];
}

/** Líneas de texto de un PDF: [{page, text}] */
async function pdfLines(pdf, progress, label) {
  const lines = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    progress && progress((i - 1) / pdf.numPages, `${label}: página ${i} de ${pdf.numPages}…`);
    const tc = await (await pdf.getPage(i)).getTextContent();
    const txt = tc.items.map((it) => (it.str || '') + (it.hasEOL ? '\n' : '')).join('');
    txt.split('\n').forEach((t) => {
      const norm = t.replace(/\s+/g, ' ').trim();
      if (norm) lines.push({ page: i, text: norm });
    });
  }
  return lines;
}

/** Diferencias entre dos listas de líneas (LCS). Devuelve [{op:'=', '-', '+', a, b}] */
function diffLines(A, B) {
  const n = A.length, m = B.length;
  const W = m + 1;
  const dp = (n + 1) * W > 2.5e7 ? null : new Uint16Array((n + 1) * W);
  if (!dp || n > 65000 || m > 65000) throw new Error('Los documentos son demasiado largos para compararlos en el navegador.');
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i * W + j] = A[i].text === B[j].text ? dp[(i + 1) * W + j + 1] + 1 : Math.max(dp[(i + 1) * W + j], dp[i * W + j + 1]);
    }
  }
  const ops = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (A[i].text === B[j].text) { ops.push({ op: '=', a: A[i], b: B[j] }); i++; j++; }
    else if (dp[(i + 1) * W + j] >= dp[i * W + j + 1]) ops.push({ op: '-', a: A[i++] });
    else ops.push({ op: '+', b: B[j++] });
  }
  while (i < n) ops.push({ op: '-', a: A[i++] });
  while (j < m) ops.push({ op: '+', b: B[j++] });
  return ops;
}

/* Caracteres que admiten las fuentes estándar del PDF (WinAnsi) */
const winAnsi = (s) => s.replace(/[^\x20-\x7E -ÿ€‘’“”•–—…]/g, '');

TOOLS.push(
  /* ───────────── EDITAR ───────────── */
  {
    id: 'editar-pdf', name: 'Editar PDF', cat: 'editar', icon: 'edit', color: '#e54666',
    desc: 'Añade texto, imágenes, rectángulos y dibujos a mano a tu PDF.',
    accept: 'pdf', multiple: false, mode: 'editor', action: 'Guardar PDF',
    hint: 'Elige una herramienta en el panel y haz clic sobre la página. Arrastra los elementos para moverlos.',
    editor: { tools: ['select', 'text', 'image', 'rect', 'draw'] },
    validate: ({ ed }) => (ed.count() ? null : 'Añade al menos un elemento al documento.'),
    run: (ctx) => saveEditor(ctx, 'editado'),
  },
  {
    id: 'recortar-pdf', name: 'Recortar PDF', cat: 'editar', icon: 'crop', color: '#00a2c7',
    desc: 'Elimina los márgenes o quédate solo con una zona de las páginas.',
    accept: 'pdf', multiple: false, mode: 'editor', action: 'Recortar PDF',
    hint: 'Arrastra el recuadro y sus esquinas para elegir la zona que quieres conservar.',
    editor: { tools: ['select'], crop: true, styles: false },
    options: () => `
      <div class="field"><label class="lbl">Aplicar a</label>
        ${choice('apply', 'all', 'Todas las páginas', '', true)}
        ${choice('apply', 'current', 'Solo la página actual', '')}
      </div>`,
    async run({ bytes, pdf, ed, file, opts }) {
      const doc = await U.pdfLib(bytes);
      const c = ed.st.crop;
      const nums = opts.apply === 'current' ? [ed.st.page] : Array.from({ length: pdf.numPages }, (_, i) => i + 1);
      for (const n of nums) {
        const vp = (await pdf.getPage(n)).getViewport({ scale: 1 });
        const [x1, y1] = vp.convertToPdfPoint(c.fx * vp.width, c.fy * vp.height);
        const [x2, y2] = vp.convertToPdfPoint((c.fx + c.fw) * vp.width, (c.fy + c.fh) * vp.height);
        doc.getPage(n - 1).setCropBox(Math.min(x1, x2), Math.min(y1, y2), Math.abs(x2 - x1), Math.abs(y2 - y1));
      }
      return [{ name: `${U.baseName(file.name)}_recortado.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'formularios-pdf', name: 'Rellenar formularios', cat: 'editar', icon: 'form', color: '#7c66dc',
    desc: 'Rellena los campos de un formulario PDF (textos, casillas y listas).',
    accept: 'pdf', multiple: false, mode: 'custom', action: 'Guardar formulario',
    options: () => `<label class="check"><input type="checkbox" name="flatten"> Aplanar: los campos dejarán de ser editables</label>`,
    async mount({ main, bytes }) {
      const doc = await U.pdfLib(bytes);
      const L = PDFLib;
      const fields = doc.getForm().getFields();
      const rows = [];
      for (const f of fields) {
        let kind = null, input = '';
        const id = `ff${rows.length}`;
        if (f instanceof L.PDFTextField) {
          kind = 'text';
          const v = escHtml(f.getText() || '');
          input = f.isMultiline() ? `<textarea id="${id}" rows="3">${v}</textarea>` : `<input id="${id}" value="${v}">`;
        } else if (f instanceof L.PDFCheckBox) {
          kind = 'check';
          input = `<label class="check"><input type="checkbox" id="${id}" ${f.isChecked() ? 'checked' : ''}> Marcado</label>`;
        } else if (f instanceof L.PDFDropdown || f instanceof L.PDFOptionList || f instanceof L.PDFRadioGroup) {
          kind = 'choice';
          const sel = f.getSelected();
          const cur = Array.isArray(sel) ? sel[0] : sel;
          input = `<select id="${id}"><option value="">—</option>${f.getOptions()
            .map((o) => `<option ${o === cur ? 'selected' : ''}>${escHtml(o)}</option>`).join('')}</select>`;
        }
        if (!kind) continue;
        const label = f.getName().split('.').pop().replace(/\[\d+\]$/, '').replace(/[_-]+/g, ' ');
        rows.push({ f, kind, id, html: `<div class="field ff"><label class="lbl" for="${id}">${escHtml(label)}</label>${input}</div>` });
      }
      if (!rows.length) throw new Error('Este PDF no tiene campos de formulario rellenables.');
      main.innerHTML = `<div class="form-card"><h3>${rows.length} campo${rows.length > 1 ? 's' : ''} encontrado${rows.length > 1 ? 's' : ''}</h3>${rows.map((r) => r.html).join('')}</div>`;
      rows.forEach((r) => (r.el = main.querySelector('#' + r.id)));
      return { doc, rows };
    },
    async run({ custom, file, opts }) {
      const { doc, rows } = custom;
      try {
        for (const r of rows) {
          if (r.kind === 'text') r.f.setText(r.el.value);
          else if (r.kind === 'check') r.el.checked ? r.f.check() : r.f.uncheck();
          else if (r.el.value) r.f.select(r.el.value);
          else if (r.f.clear) r.f.clear();
        }
        if (opts.flatten) doc.getForm().flatten();
        const bytes = await doc.save();
        return [{ name: `${U.baseName(file.name)}_relleno.pdf`, blob: U.pdfBlob(bytes) }];
      } catch (e) {
        console.error(e);
        throw new Error('No se pudo guardar el formulario. Revisa que no haya emojis u otros caracteres especiales.');
      }
    },
  },

  /* ───────────── SEGURIDAD ───────────── */
  {
    id: 'firmar-pdf', name: 'Firmar PDF', cat: 'seguridad', icon: 'sign', color: '#2a7e3b',
    desc: 'Dibuja, escribe o sube tu firma y colócala en el documento.',
    accept: 'pdf', multiple: false, mode: 'editor', action: 'Firmar PDF',
    hint: 'Crea tu firma en el panel, pulsa «Colocar firma» y arrástrala hasta su sitio.',
    editor: { tools: ['select'], styles: false },
    setup({ side, ed, toast }) {
      side.innerHTML = `
        <div class="field"><label class="lbl">Tu firma</label>
          <div class="tabs">
            <button type="button" data-m="draw" class="active">Dibujar</button>
            <button type="button" data-m="type">Escribir</button>
            <button type="button" data-m="upload">Subir imagen</button>
          </div>
          <div data-p="draw"><canvas class="sigpad" width="640" height="240"></canvas>
            <button type="button" class="link sig-clear">Borrar</button></div>
          <div data-p="type" hidden class="opts">
            <input class="sig-name" placeholder="Escribe tu nombre" maxlength="40">
            <select class="sig-font">
              <option value='"Brush Script MT", "Segoe Script", cursive'>Manuscrita</option>
              <option value='"Snell Roundhand", "Apple Chancery", "Lucida Handwriting", cursive'>Caligráfica</option>
              <option value='Georgia, "Times New Roman", serif'>Clásica</option>
            </select>
          </div>
          <div data-p="upload" hidden><input type="file" class="sig-file" accept="image/*"></div>
        </div>
        <div class="field"><label class="lbl">Color de la firma</label>
          <div class="swatches">${['#1c1b29', '#1d4ed8', '#c62a2f']
            .map((c, i) => `<button type="button" class="sw ${i ? '' : 'active'}" data-c="${c}" style="background:${c}" aria-label="Color"></button>`).join('')}</div></div>
        <button type="button" class="btn-sec sig-add">✍️ Colocar firma en esta página</button>
        <button type="button" class="link sig-all">Copiar la firma seleccionada a todas las páginas</button>`;

      let mode = 'draw';
      let color = '#1c1b29';
      const pad = side.querySelector('.sigpad');
      const pctx = pad.getContext('2d');
      const strokes = [];
      const redraw = () => {
        pctx.clearRect(0, 0, pad.width, pad.height);
        Object.assign(pctx, { lineWidth: 5, lineCap: 'round', lineJoin: 'round', strokeStyle: color });
        for (const s of strokes) {
          pctx.beginPath();
          s.forEach(([x, y], i) => (i ? pctx.lineTo(x, y) : pctx.moveTo(x, y)));
          if (s.length === 1) pctx.lineTo(s[0][0] + 0.1, s[0][1]);
          pctx.stroke();
        }
      };
      const pt = (e) => {
        const r = pad.getBoundingClientRect();
        return [((e.clientX - r.left) / r.width) * pad.width, ((e.clientY - r.top) / r.height) * pad.height];
      };
      pad.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        pad.setPointerCapture(e.pointerId);
        strokes.push([pt(e)]);
        redraw();
      });
      pad.addEventListener('pointermove', (e) => {
        if (!pad.hasPointerCapture(e.pointerId)) return;
        strokes[strokes.length - 1].push(pt(e));
        redraw();
      });
      side.querySelector('.sig-clear').addEventListener('click', () => { strokes.length = 0; redraw(); });

      side.querySelector('.tabs').addEventListener('click', (e) => {
        const b = e.target.closest('[data-m]');
        if (!b) return;
        mode = b.dataset.m;
        side.querySelectorAll('[data-m]').forEach((x) => x.classList.toggle('active', x === b));
        side.querySelectorAll('[data-p]').forEach((p) => (p.hidden = p.dataset.p !== mode));
      });
      side.querySelector('.swatches').addEventListener('click', (e) => {
        const b = e.target.closest('[data-c]');
        if (!b) return;
        color = b.dataset.c;
        side.querySelectorAll('.sw').forEach((x) => x.classList.toggle('active', x === b));
        redraw();
      });

      async function signatureImage() {
        if (mode === 'upload') {
          const f = side.querySelector('.sig-file').files[0];
          if (!f) throw new Error('Elige una imagen con tu firma.');
          return Editor.imageToPng(f);
        }
        let c = pad;
        if (mode === 'type') {
          const name = side.querySelector('.sig-name').value.trim();
          if (!name) throw new Error('Escribe tu nombre para crear la firma.');
          c = document.createElement('canvas');
          c.width = 1400; c.height = 300;
          const x = c.getContext('2d');
          x.font = `110px ${side.querySelector('.sig-font').value}`;
          x.fillStyle = color;
          x.textBaseline = 'middle';
          x.fillText(name, 30, 150);
        }
        const t = trimCanvas(c);
        if (!t) throw new Error('Dibuja tu firma en el recuadro.');
        return { url: t.toDataURL('image/png'), w: t.width, h: t.height };
      }

      side.querySelector('.sig-add').addEventListener('click', async () => {
        try {
          const { url, w, h } = await signatureImage();
          const o = ed.addImage(url, w, h, 0.28);
          // por defecto, abajo a la derecha
          o.x = ed.st.vp.width - o.w - ed.st.vp.width * 0.08;
          o.y = ed.st.vp.height - o.h - ed.st.vp.height * 0.08;
          ed.render();
        } catch (e) {
          toast(e.message);
        }
      });
      side.querySelector('.sig-all').addEventListener('click', () => {
        const o = ed.st.sel;
        if (!o || o.type !== 'image') return toast('Selecciona primero una firma colocada en la página.');
        let id = Date.now();
        for (let p = 1; p <= ed.st.n; p++) {
          if (p === ed.st.page) continue;
          const l = (ed.st.objs[p] = ed.st.objs[p] || []);
          l.push({ type: 'image', url: o.url, ratio: o.ratio, x: o.x, y: o.y, w: o.w, h: o.h, id: id++ });
        }
        toast(`Firma copiada a ${ed.st.n - 1} página${ed.st.n > 2 ? 's' : ''} más.`);
      });
    },
    validate: ({ ed }) => (ed.count() ? null : 'Coloca tu firma en el documento.'),
    run: (ctx) => saveEditor(ctx, 'firmado'),
  },
  {
    id: 'censurar-pdf', name: 'Censurar PDF', cat: 'seguridad', icon: 'redact', color: '#1c1b29',
    desc: 'Tapa de forma permanente texto o zonas con información sensible.',
    accept: 'pdf', multiple: false, mode: 'editor', action: 'Censurar PDF',
    hint: 'Dibuja recuadros sobre lo que quieras ocultar, o busca un texto para taparlo en todo el documento.',
    editor: { tools: ['rect'], styles: false, color: '#000000', fill: true },
    setup({ side, ed, pdf, toast }) {
      side.innerHTML = `
        <div class="field"><label class="lbl" for="rd-q">Buscar texto y censurarlo</label>
          <div class="inline"><input id="rd-q" placeholder="Ej.: DNI, un nombre…" autocomplete="off">
          <button type="button" class="rd-go">Buscar</button></div></div>
        <small class="note">Las páginas censuradas se convierten en imagen para que el texto tapado no se pueda recuperar.</small>`;
      side.querySelector('.rd-go').addEventListener('click', async () => {
        const q = side.querySelector('#rd-q').value.trim().toLowerCase();
        if (!q) return;
        let found = 0;
        let id = Date.now();
        for (let p = 1; p <= pdf.numPages; p++) {
          const page = await pdf.getPage(p);
          const vp = page.getViewport({ scale: 1 });
          const tc = await page.getTextContent();
          for (const it of tc.items) {
            const s = (it.str || '').toLowerCase();
            if (!s) continue;
            let idx = s.indexOf(q);
            while (idx !== -1) {
              const tx = pdfjsLib.Util.transform(vp.transform, it.transform);
              const fh = Math.hypot(tx[2], tx[3]);
              const w = it.width * vp.scale;
              const l = (ed.st.objs[p] = ed.st.objs[p] || []);
              l.push({
                type: 'rect', color: '#000000', fill: true, id: id++,
                x: tx[4] + (w * idx) / s.length - 1,
                y: tx[5] - fh * 0.92,
                w: (w * q.length) / s.length + 2,
                h: fh * 1.18,
              });
              found++;
              idx = s.indexOf(q, idx + q.length);
            }
          }
        }
        ed.render();
        toast(found ? `Se han marcado ${found} coincidencia${found > 1 ? 's' : ''}.` : 'No se ha encontrado ese texto.');
      });
    },
    validate: ({ ed }) => (ed.count() ? null : 'Marca al menos una zona para censurar.'),
    async run({ bytes, pdf, ed, file, progress }) {
      const doc = await U.pdfLib(bytes);
      const list = ed.pages();
      for (let k = 0; k < list.length; k++) {
        const { page: n, objs } = list[k];
        progress(k / list.length, `Censurando página ${n}…`);
        const scale = 2.5;
        const { canvas, width, height } = await U.render(pdf, n, { scale });
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#000';
        objs.forEach((o) => ctx.fillRect(o.x * scale, o.y * scale, o.w * scale, o.h * scale));
        const img = await doc.embedJpg(await (await U.toBlob(canvas, 'image/jpeg', 0.9)).arrayBuffer());
        canvas.width = canvas.height = 0;
        doc.removePage(n - 1);
        doc.insertPage(n - 1, [width, height]).drawImage(img, { x: 0, y: 0, width, height });
      }
      return [{ name: `${U.baseName(file.name)}_censurado.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'proteger-pdf', name: 'Proteger PDF', cat: 'seguridad', icon: 'lock', color: '#cf3c2e',
    desc: 'Pon una contraseña a tu PDF para que nadie más pueda abrirlo.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Proteger PDF',
    options: () => `
      <div class="field"><label class="lbl" for="o-pw">Contraseña</label>
        <input id="o-pw" type="password" name="pw" autocomplete="new-password"></div>
      <div class="field"><label class="lbl" for="o-pw2">Repite la contraseña</label>
        <input id="o-pw2" type="password" name="pw2" autocomplete="new-password"></div>
      <div class="field"><label class="lbl">Permisos al abrirlo</label>
        <label class="check"><input type="checkbox" name="print" checked> Permitir imprimir</label>
        <label class="check"><input type="checkbox" name="copy" checked> Permitir copiar texto</label>
        <label class="check"><input type="checkbox" name="modify"> Permitir modificar</label>
      </div>
      <small class="note">Cifrado AES de 256 bits. Si olvidas la contraseña no podrás recuperar el archivo.</small>`,
    validate: ({ opts }) => {
      if (!opts.pw) return 'Escribe una contraseña.';
      if (opts.pw !== opts.pw2) return 'Las contraseñas no coinciden.';
      return null;
    },
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const bytes = await U.bytes(file);
      progress(0.1, 'Preparando el cifrado…');
      const L = await U.pdfLibCrypto();
      let doc;
      try {
        doc = await L.PDFDocument.load(bytes);
      } catch (e) {
        throw new Error(/encrypt/i.test(String(e.message)) ? 'Este PDF ya está protegido. Desbloquéalo primero.' : 'No se pudo leer el PDF.');
      }
      const owner = Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, '0')).join('');
      doc.encrypt({
        userPassword: opts.pw,
        ownerPassword: owner,
        permissions: {
          printing: opts.print ? 'highResolution' : false,
          copying: !!opts.copy,
          modifying: !!opts.modify,
          annotating: !!opts.modify,
          fillingForms: true,
          contentAccessibility: true,
          documentAssembly: !!opts.modify,
        },
      });
      progress(0.6, 'Cifrando…');
      return [{ name: `${U.baseName(file.name)}_protegido.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'desbloquear-pdf', name: 'Desbloquear PDF', cat: 'seguridad', icon: 'unlock', color: '#e8890c',
    desc: 'Quita la contraseña y las restricciones de un PDF del que conoces la clave.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Desbloquear PDF',
    options: () => `
      <div class="field"><label class="lbl" for="o-upw">Contraseña del PDF</label>
        <input id="o-upw" type="password" name="pw" autocomplete="off">
        <small>Déjala vacía si el PDF se abre sin contraseña pero no deja imprimir o copiar.</small></div>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const bytes = await U.bytes(file);
      const L = await U.pdfLibCrypto();
      const probe = await L.PDFDocument.load(bytes, { ignoreEncryption: true }).catch(() => null);
      if (!probe) throw new Error('No se pudo leer el PDF. Puede que esté dañado.');
      if (!probe.isEncrypted) throw new Error('Este PDF no está protegido: no hace falta desbloquearlo.');
      await U.checkPassword(bytes, opts.pw || undefined);
      progress(0.4, 'Descifrando…');
      const src = await L.PDFDocument.load(bytes, { password: opts.pw || '' });
      const out = await L.PDFDocument.create();
      (await out.copyPages(src, src.getPageIndices())).forEach((p) => out.addPage(p));
      return [{ name: `${U.baseName(file.name)}_desbloqueado.pdf`, blob: U.pdfBlob(await out.save()) }];
    },
  },
  {
    id: 'comparar-pdf', name: 'Comparar PDF', cat: 'seguridad', icon: 'compare', color: '#0d74ce',
    desc: 'Compara el texto de dos versiones de un PDF y descubre qué ha cambiado.',
    accept: 'pdf', multiple: true, reorder: true, mode: 'files', action: 'Comparar',
    hint: 'El primer archivo es la versión original y el segundo, la nueva.',
    validate: ({ files }) => (files.length !== 2 ? 'Añade exactamente 2 archivos PDF para compararlos.' : null),
    async run({ files, progress }) {
      const [fa, fb] = files.map((f) => f.file);
      const pa = await U.pdfJs(await U.bytes(fa));
      const pb = await U.pdfJs(await U.bytes(fb));
      const A = await pdfLines(pa, (f, t) => progress(f / 2, t), 'Leyendo el original');
      const B = await pdfLines(pb, (f, t) => progress(0.5 + f / 2, t), 'Leyendo la versión nueva');
      pa.destroy(); pb.destroy();
      if (!A.length && !B.length) throw new Error('Ninguno de los dos PDF contiene texto (¿son escaneados? Prueba antes con OCR).');
      progress(1, 'Comparando…');
      const ops = diffLines(A, B);
      const del = ops.filter((o) => o.op === '-').length;
      const add = ops.filter((o) => o.op === '+').length;

      // Agrupa líneas iguales largas para que el informe sea legible
      const rows = [];
      for (let i = 0; i < ops.length; i++) {
        const o = ops[i];
        if (o.op !== '=') {
          rows.push(`<div class="d ${o.op === '-' ? 'del' : 'add'}"><span class="pg">p. ${(o.a || o.b).page}</span><span class="sg">${o.op === '-' ? '−' : '+'}</span>${escHtml((o.a || o.b).text)}</div>`);
          continue;
        }
        let j = i;
        while (j < ops.length && ops[j].op === '=') j++;
        const run = ops.slice(i, j);
        const show = (x) => `<div class="d"><span class="pg">p. ${x.a.page}</span><span class="sg"></span>${escHtml(x.a.text)}</div>`;
        if (run.length > 6) {
          if (i > 0) rows.push(...run.slice(0, 2).map(show));
          rows.push(`<div class="d gap">· · · ${run.length - (i > 0 ? 2 : 0) - (j < ops.length ? 2 : 0)} líneas iguales · · ·</div>`);
          if (j < ops.length) rows.push(...run.slice(-2).map(show));
        } else rows.push(...run.map(show));
        i = j - 1;
      }
      const css = `.cmp{font:14px/1.5 -apple-system,Segoe UI,Arial,sans-serif;color:#1c1b29}.cmp h2{font-size:18px;margin:0 0 4px}.cmp .stats{margin:0 0 14px;color:#6b6a7d}
        .cmp .stats b.r{color:#c62a2f}.cmp .stats b.g{color:#218358}.cmp .d{padding:3px 10px;border-left:3px solid transparent;white-space:pre-wrap;word-break:break-word}
        .cmp .d.del{background:#fff0f0;border-color:#e5484d;text-decoration:line-through;text-decoration-color:#e5484d88}.cmp .d.add{background:#ecfbf2;border-color:#30a46c}
        .cmp .gap{color:#8b8a9c;text-align:center;font-size:12px;background:#f5f5fa}.cmp .pg{display:inline-block;width:48px;color:#9a99ab;font-size:12px}.cmp .sg{display:inline-block;width:16px;font-weight:700}`;
      const body = `<div class="cmp"><h2>${escHtml(fa.name)} → ${escHtml(fb.name)}</h2>
        <p class="stats"><b class="r">${del} línea${del === 1 ? '' : 's'} eliminada${del === 1 ? '' : 's'}</b> · <b class="g">${add} línea${add === 1 ? '' : 's'} añadida${add === 1 ? '' : 's'}</b></p>
        ${rows.join('')}</div>`;
      const html = `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>Comparación de PDF</title><style>body{margin:24px}${css}</style></head><body>${body}</body></html>`;
      return {
        files: [{ name: 'comparacion.html', blob: new Blob([html], { type: 'text/html;charset=utf-8' }) }],
        title: del || add ? 'Hemos encontrado diferencias' : 'Los dos documentos tienen el mismo texto',
        message: `${del} eliminada${del === 1 ? '' : 's'}, ${add} añadida${add === 1 ? '' : 's'}. Descarga el informe para guardarlo.`,
        html: `<style>${css}</style>${body}`,
      };
    },
  },

  /* ───────────── OPTIMIZAR ───────────── */
  {
    id: 'ocr-pdf', name: 'OCR PDF', cat: 'optimizar', icon: 'ocr', color: '#3e9b4f',
    desc: 'Reconoce el texto de PDF escaneados para poder buscarlo y copiarlo.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Reconocer texto',
    options: () => `
      ${select('lang', 'Idioma del documento', [['spa', 'Español'], ['eng', 'Inglés'], ['spa+eng', 'Español + inglés'], ['cat', 'Catalán'], ['fra', 'Francés'], ['deu', 'Alemán'], ['ita', 'Italiano'], ['por', 'Portugués']], 'spa')}
      <div class="field"><label class="lbl">Resultado</label>
        ${choice('out', 'pdf', 'PDF con texto seleccionable', 'Mantiene el aspecto original', true)}
        ${choice('out', 'txt', 'Archivo de texto (.txt)', 'Solo el texto reconocido')}
      </div>
      <small class="note">La primera vez se descargan los datos del idioma (unos 10–20 MB). Cada página tarda unos segundos.</small>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const bytes = await U.bytes(file);
      const pdf = await U.pdfJs(bytes);
      const n = pdf.numPages;
      progress(0, 'Cargando el motor de reconocimiento…');
      await U.loadScript('vendor/tesseract.min.js');
      let cur = 0;
      const worker = await Tesseract.createWorker(opts.lang || 'spa', 1, {
        logger: (m) => {
          if (m.status === 'recognizing text') progress((cur + m.progress) / n, `Reconociendo texto: página ${cur + 1} de ${n}…`);
          else if (cur === 0 && m.progress < 1) progress(0, 'Descargando datos del idioma…');
        },
      });
      const doc = opts.out === 'pdf' ? await U.pdfLib(bytes) : null;
      const font = doc ? await doc.embedFont(StandardFonts.Helvetica) : null;
      const texts = [];
      let words = 0;
      try {
        for (cur = 0; cur < n; cur++) {
          const scale = 2.5;
          const { canvas } = await U.render(pdf, cur + 1, { scale });
          const { data } = await worker.recognize(canvas, {}, { text: true, blocks: !!doc });
          canvas.width = canvas.height = 0;
          texts.push(data.text.trim());
          if (!doc) continue;
          const page = doc.getPage(cur);
          const vp = (await pdf.getPage(cur + 1)).getViewport({ scale: 1 });
          const rot = degrees(page.getRotation().angle);
          for (const b of data.blocks || []) for (const par of b.paragraphs) for (const line of par.lines) for (const w of line.words) {
            const t = winAnsi(w.text || '').trim();
            if (!t) continue;
            const { x0, y0, x1, y1 } = w.bbox;
            const ww = (x1 - x0) / scale;
            const hh = (y1 - y0) / scale;
            const unit = font.widthOfTextAtSize(t, 1);
            const size = Math.max(1, Math.min(hh * 1.4, unit ? ww / unit : hh));
            const [px, py] = vp.convertToPdfPoint(x0 / scale, y1 / scale - hh * 0.15);
            page.drawText(t, { x: px, y: py, size, font, opacity: 0, rotate: rot });
            words++;
          }
        }
      } finally {
        await worker.terminate();
        pdf.destroy();
      }
      const base = U.baseName(file.name);
      if (!doc) {
        const txt = texts.map((t, i) => `──────── Página ${i + 1} ────────\n\n${t}`).join('\n\n');
        return [{ name: `${base}_ocr.txt`, blob: new Blob([txt], { type: 'text/plain;charset=utf-8' }) }];
      }
      return {
        files: [{ name: `${base}_ocr.pdf`, blob: U.pdfBlob(await doc.save()) }],
        message: `${words.toLocaleString('es')} palabras reconocidas. Ya puedes buscar y copiar el texto.`,
      };
    },
  },

  /* ───────────── CONVERTIR ───────────── */
  {
    id: 'pdf-a-markdown', name: 'PDF a Markdown', cat: 'convertir', icon: 'markdown', color: '#5d6ad1',
    desc: 'Convierte un PDF a Markdown con títulos, párrafos y listas. Ideal para notas e IA.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Convertir a Markdown',
    async run({ files, progress }) {
      const f = files[0].file;
      const pdf = await U.pdfJs(await U.bytes(f));
      try {
        const text = blocksToMarkdown(await pdfBlocks(pdf, progress)); // definidas en office.js
        return [{ name: `${U.baseName(f.name)}.md`, blob: new Blob([text], { type: 'text/markdown;charset=utf-8' }) }];
      } finally {
        pdf.destroy();
      }
    },
  }
);
