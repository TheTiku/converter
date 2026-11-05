'use strict';

/* Herramientas: PDF/A, Reparar, Rotar varios, Crear formularios, HTML a PDF, Escanear y Flujos de trabajo */

Object.assign(ICONS, {
  archive: '<rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8M10 12h4"/>',
  wrench: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.6 2.6-2.4-.6-.6-2.4z"/>',
  formNew: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h6M7 12h3M15 15h4M17 13v4"/>',
  code: '<path d="m8 8-5 4 5 4M16 8l5 4-5 4M14 4l-4 16"/>',
  camera: '<path d="M4 7h3l2-3h6l2 3h3a1 1 0 0 1 1 1v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V8a1 1 0 0 1 1-1z"/><circle cx="12" cy="13" r="4"/>',
  workflow: '<rect x="3" y="3" width="6" height="6" rx="1"/><rect x="15" y="15" width="6" height="6" rx="1"/><path d="M9 6h5a3 3 0 0 1 3 3v6"/><path d="m14.5 12.5 2.5 2.5 2.5-2.5"/>',
  rotateMany: '<path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><rect x="9" y="9" width="6" height="7" rx="1"/>',
});

const normFiles = (res) => (Array.isArray(res) ? res : res.files);
const asPdfFile = (f) => new File([f.blob], f.name, { type: f.blob.type || 'application/pdf' });

/** Cuenta las fuentes que el PDF usa sin incrustar (PDF/A exige que todas estén incrustadas). */
function nonEmbeddedFonts(doc, L) {
  const N = (n) => L.PDFName.of(n);
  let missing = 0;
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof L.PDFDict)) continue;
    if (obj.get(N('Type')) === N('FontDescriptor')) {
      if (!obj.get(N('FontFile')) && !obj.get(N('FontFile2')) && !obj.get(N('FontFile3'))) missing++;
    } else if (obj.get(N('Type')) === N('Font') && [N('Type1'), N('TrueType'), N('MMType1')].includes(obj.get(N('Subtype'))) && !obj.get(N('FontDescriptor'))) {
      missing++; // fuentes estándar (Helvetica, Times…) sin incrustar
    }
  }
  return missing;
}

/** Convierte las páginas de un PDF (pdf.js) en un PDF de imágenes con la librería indicada. */
async function rasterize(pdf, L, progress, scale = 2) {
  const out = await L.PDFDocument.create();
  for (let i = 1; i <= pdf.numPages; i++) {
    progress && progress((i - 1) / pdf.numPages, `Procesando página ${i} de ${pdf.numPages}…`);
    const { canvas, width, height } = await U.render(pdf, i, { scale });
    const jpg = await U.toBlob(canvas, 'image/jpeg', 0.9);
    canvas.width = canvas.height = 0;
    const img = await out.embedJpg(await jpg.arrayBuffer());
    out.addPage([width, height]).drawImage(img, { x: 0, y: 0, width, height });
  }
  return out;
}

/**
 * Reconstruye un PDF muy dañado: localiza todos sus objetos (también los que están dentro
 * de flujos de objetos comprimidos), busca el catálogo de páginas y crea un archivo nuevo
 * con su tabla de referencias. Devuelve los bytes o null.
 */
async function rebuildPdf(bytes) {
  const latin1 = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return s; };
  const src = latin1(bytes);
  const objs = new Map(); // número → texto "<< … >>" o "<< … >>stream…endstream"
  const re = /(\d+)\s+(\d+)\s+obj\b/g;
  let m;
  while ((m = re.exec(src))) {
    const start = re.lastIndex;
    const end = src.indexOf('endobj', start);
    if (end < 0) break; // objeto cortado: se descarta
    const next = src.slice(start, end);
    // si aparece otro "obj" antes del endobj, este objeto está roto
    if (/\n\d+\s+\d+\s+obj\b/.test(next.slice(0, next.indexOf('stream') > -1 ? next.indexOf('stream') : next.length))) continue;
    objs.set(+m[1], next.trim());
    re.lastIndex = end + 6;
  }
  // objetos dentro de flujos de objetos (ObjStm)
  const inflate = async (u8) => {
    try {
      const ds = new DecompressionStream('deflate');
      return new Uint8Array(await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer());
    } catch (_) { return null; }
  };
  for (const [num, body] of [...objs]) {
    if (!/\/Type\s*\/ObjStm/.test(body)) continue;
    const si = body.indexOf('stream');
    let data = body.slice(si + 6).replace(/^\r?\n/, '');
    data = data.slice(0, data.lastIndexOf('endstream'));
    const raw = Uint8Array.from(data, (c) => c.charCodeAt(0) & 255);
    const dec = /FlateDecode/.test(body.slice(0, si)) ? await inflate(raw) : raw;
    if (!dec) continue;
    const txt = latin1(dec);
    const first = +(/\/First\s+(\d+)/.exec(body) || [])[1];
    const n = +(/\/N\s+(\d+)/.exec(body) || [])[1];
    if (!first || !n) continue;
    const head = txt.slice(0, first).trim().split(/\s+/).map(Number);
    for (let k = 0; k < n; k++) {
      const on = head[k * 2], off = head[k * 2 + 1];
      const nextOff = k + 1 < n ? head[k * 2 + 3] : txt.length - first;
      if (!objs.has(on)) objs.set(on, txt.slice(first + off, first + nextOff).trim());
    }
    objs.delete(num);
  }
  for (const [num, body] of [...objs]) if (/\/Type\s*\/XRef/.test(body)) objs.delete(num);
  if (!objs.size) return null;
  // catálogo: el existente o uno nuevo que apunte al árbol de páginas raíz
  let root = [...objs].find(([, b]) => /\/Type\s*\/Catalog/.test(b))?.[0];
  if (root == null) {
    const pages = [...objs].find(([, b]) => /\/Type\s*\/Pages\b/.test(b) && !/\/Parent\s/.test(b));
    if (!pages) return null;
    root = Math.max(...objs.keys()) + 1;
    objs.set(root, `<< /Type /Catalog /Pages ${pages[0]} 0 R >>`);
  }
  const size = Math.max(...objs.keys()) + 1;
  let out = '%PDF-1.7\n%\xe2\xe3\xcf\xd3\n';
  const offsets = [];
  for (const [num, body] of [...objs].sort((a, b) => a[0] - b[0])) {
    offsets[num] = out.length;
    out += `${num} 0 obj\n${body}\nendobj\n`;
  }
  const xref = out.length;
  out += `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let i = 1; i < size; i++) out += offsets[i] != null ? `${String(offsets[i]).padStart(10, '0')} 00000 n \n` : '0000000000 65535 f \n';
  out += `trailer\n<< /Size ${size} /Root ${root} 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Uint8Array.from(out, (c) => c.charCodeAt(0) & 255);
}

/* ───────────── utilidades de imagen (escáner) ───────────── */
async function fileToCanvas(file, max = 2600) {
  const url = URL.createObjectURL(file);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.round(img.naturalWidth * k);
    c.height = Math.round(img.naturalHeight * k);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    return c;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function rotateCanvas(src, deg) {
  if (!deg) return src;
  const c = document.createElement('canvas');
  const swap = deg % 180 !== 0;
  c.width = swap ? src.height : src.width;
  c.height = swap ? src.width : src.height;
  const x = c.getContext('2d');
  x.translate(c.width / 2, c.height / 2);
  x.rotate((deg * Math.PI) / 180);
  x.drawImage(src, -src.width / 2, -src.height / 2);
  return c;
}

/** Filtros de "documento escaneado": mejorar, escala de grises o blanco y negro. */
function scanFilter(src, mode) {
  const c = document.createElement('canvas');
  c.width = src.width; c.height = src.height;
  const x = c.getContext('2d');
  x.drawImage(src, 0, 0);
  if (mode === 'color') return c;
  // iluminación del papel: la foto muy reducida y desenfocada (sombras, tono de la luz)
  const bgc = document.createElement('canvas');
  bgc.width = Math.max(8, Math.round(c.width / 24)); bgc.height = Math.max(8, Math.round(c.height / 24));
  const bx = bgc.getContext('2d');
  bx.filter = 'blur(3px)';
  bx.drawImage(src, 0, 0, bgc.width, bgc.height);
  const big = document.createElement('canvas');
  big.width = c.width; big.height = c.height;
  const gx = big.getContext('2d');
  gx.imageSmoothingQuality = 'high';
  gx.drawImage(bgc, 0, 0, c.width, c.height);
  const bg = gx.getImageData(0, 0, c.width, c.height).data;
  const im = x.getImageData(0, 0, c.width, c.height);
  const d = im.data;
  // divide por la iluminación: el papel queda blanco y la tinta conserva su color
  for (let i = 0; i < d.length; i += 4) {
    for (let k = 0; k < 3; k++) d[i + k] = Math.min(255, (d[i + k] * 245) / Math.max(1, bg[i + k]));
  }
  const n = d.length / 4;
  // niveles: percentiles 2% y 98% de la luminancia para estirar el contraste
  const hist = new Uint32Array(256);
  for (let i = 0; i < d.length; i += 4) hist[(d[i] * 77 + d[i + 1] * 150 + d[i + 2] * 29) >> 8]++;
  let lo = 0, hi = 255, acc = 0;
  for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc > n * 0.02) { lo = v; break; } }
  acc = 0;
  for (let v = 255; v >= 0; v--) { acc += hist[v]; if (acc > n * 0.02) { hi = v; break; } }
  const span = Math.max(1, hi - lo);
  let thr = 128;
  if (mode === 'bn') {
    // umbral de Otsu
    let sum = 0, sumB = 0, wB = 0, best = 0;
    for (let v = 0; v < 256; v++) sum += v * hist[v];
    for (let v = 0; v < 256; v++) {
      wB += hist[v]; if (!wB) continue;
      const wF = n - wB; if (!wF) break;
      sumB += v * hist[v];
      const mB = sumB / wB, mF = (sum - sumB) / wF;
      const between = wB * wF * (mB - mF) ** 2;
      if (between > best) { best = between; thr = v; }
    }
  }
  for (let i = 0; i < d.length; i += 4) {
    if (mode === 'mejorar') {
      for (let k = 0; k < 3; k++) d[i + k] = Math.max(0, Math.min(255, ((d[i + k] - lo) * 255) / span));
    } else {
      const l = (d[i] * 77 + d[i + 1] * 150 + d[i + 2] * 29) >> 8;
      const v = mode === 'bn' ? (l > thr ? 255 : 0) : Math.max(0, Math.min(255, ((l - lo) * 255) / span));
      d[i] = d[i + 1] = d[i + 2] = v;
    }
  }
  x.putImageData(im, 0, 0);
  return c;
}

/* ───────────── pasos disponibles en los flujos de trabajo ───────────── */
const WF_STEPS = (() => {
  const tool = (id) => TOOLS.find((t) => t.id === id);
  const fromTool = (id) => ({
    id,
    get name() { return tool(id).name; },
    options: () => (tool(id).options ? tool(id).options() : ''),
    validate: (opts, files) => (tool(id).validate ? tool(id).validate({ opts, files: files.map((f) => ({ file: f })) }) : null),
    run: async (files, opts, progress) => {
      const out = [];
      for (const f of files) out.push(...normFiles(await tool(id).run({ files: [{ file: f }], opts, progress })));
      return out;
    },
  });
  return [
    {
      id: 'unir-pdf', name: 'Unir todos los archivos', options: () => '<small class="note">Une en un solo PDF todos los archivos que lleguen a este paso.</small>',
      run: async (files, opts, progress) => (files.length < 2 ? files.map((f) => ({ name: f.name, blob: f })) : normFiles(await tool('unir-pdf').run({ files: files.map((f) => ({ file: f })), opts, progress }))),
    },
    fromTool('comprimir-pdf'),
    {
      id: 'rotar', name: 'Rotar todas las páginas',
      options: () => select('angle', 'Giro', [['90', '90° a la derecha'], ['180', '180°'], ['270', '90° a la izquierda']], '90'),
      run: async (files, opts) => {
        const out = [];
        for (const f of files) {
          const doc = await U.pdfLib(await U.bytes(f));
          doc.getPages().forEach((p) => p.setRotation(degrees((p.getRotation().angle + +opts.angle) % 360)));
          out.push({ name: f.name, blob: U.pdfBlob(await doc.save()) });
        }
        return out;
      },
    },
    {
      id: 'eliminar', name: 'Eliminar páginas',
      options: () => `<div class="field"><label class="lbl">Páginas a eliminar</label><input name="ranges" placeholder="Ej.: 1, 3-4"></div>`,
      validate: (opts) => (String(opts.ranges || '').trim() ? null : 'Indica qué páginas quieres eliminar.'),
      run: async (files, opts) => {
        const out = [];
        for (const f of files) {
          const doc = await U.pdfLib(await U.bytes(f));
          const total = doc.getPageCount();
          const del = new Set();
          for (const [a, b] of U.parseRanges(opts.ranges, total)) for (let i = a; i <= b; i++) del.add(i - 1);
          if (del.size >= total) throw new Error(`No se pueden eliminar todas las páginas de «${f.name}».`);
          [...del].sort((a, b) => b - a).forEach((i) => doc.removePage(i));
          out.push({ name: f.name, blob: U.pdfBlob(await doc.save()) });
        }
        return out;
      },
    },
    fromTool('numeros-de-pagina'),
    fromTool('marca-de-agua'),
    fromTool('ocr-pdf'),
    fromTool('reparar-pdf'),
    fromTool('pdf-a-pdfa'),
    fromTool('dividir-pdf'),
    { ...fromTool('proteger-pdf'), last: true },
  ];
})();

const WF_KEY = 'converter-flujos';
// los flujos guardados con el nombre anterior de la app se conservan
try { if (!localStorage.getItem(WF_KEY) && localStorage.getItem('pdfacil-flujos')) localStorage.setItem(WF_KEY, localStorage.getItem('pdfacil-flujos')); } catch (_) { /* sin almacenamiento */ }
const wfSaved = () => { try { return JSON.parse(localStorage.getItem(WF_KEY)) || []; } catch (_) { return []; } };

TOOLS.push(
  /* ───────────── PDF/A ───────────── */
  {
    id: 'pdf-a-pdfa', name: 'PDF a PDF/A', cat: 'convertir', icon: 'archive', color: '#4a5d7e',
    desc: 'Convierte tu PDF al estándar ISO PDF/A para archivarlo a largo plazo.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Convertir a PDF/A',
    options: () => `
      ${select('level', 'Nivel de conformidad', [['2B', 'PDF/A-2b (recomendado)'], ['1B', 'PDF/A-1b (máxima compatibilidad)'], ['3B', 'PDF/A-3b']], '2B')}
      <div class="field"><label class="lbl">Método</label>
        ${choice('pmode', 'keep', 'Conservar el contenido', 'Texto seleccionable. Requiere que las fuentes estén incrustadas', true)}
        ${choice('pmode', 'image', 'Páginas como imagen', 'Cumplimiento garantizado, pero el texto deja de ser seleccionable')}
      </div>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const bytes = await U.bytes(file);
      progress(0.05, 'Preparando…');
      const L = await U.pdfLibCrypto();
      let doc;
      let message = '';
      if (opts.pmode === 'image') {
        const pdf = await U.pdfJs(bytes);
        doc = await rasterize(pdf, L, progress);
        pdf.destroy();
      } else {
        try {
          doc = await L.PDFDocument.load(bytes);
        } catch (e) {
          throw new Error(/encrypt/i.test(String(e.message)) ? 'El PDF está protegido: PDF/A no admite cifrado. Desbloquéalo primero.' : 'No se pudo leer el PDF.');
        }
        const missing = nonEmbeddedFonts(doc, L);
        if (missing) message = `Atención: ${missing} fuente${missing > 1 ? 's no están incrustadas' : ' no está incrustada'}, así que el archivo podría no superar una validación estricta. Para un cumplimiento garantizado usa «Páginas como imagen».`;
      }
      progress(0.9, 'Aplicando el estándar PDF/A…');
      doc.convertToPDFA({ conformance: opts.level || '2B' });
      const out = await doc.save({ useObjectStreams: false });
      return { files: [{ name: `${U.baseName(file.name)}_PDFA.pdf`, blob: U.pdfBlob(out) }], title: `PDF/A-${(opts.level || '2B').toLowerCase()} listo`, message };
    },
  },

  /* ───────────── Reparar ───────────── */
  {
    id: 'reparar-pdf', name: 'Reparar PDF', cat: 'optimizar', icon: 'wrench', color: '#c2410c',
    desc: 'Recupera el contenido de PDF dañados o que no se abren bien.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Reparar PDF',
    options: () => `<small class="note">Se reconstruye el archivo página a página. Las páginas que no se puedan recuperar intactas se rescatan como imagen.</small>`,
    async run({ files, progress }) {
      const file = files[0].file;
      const bytes = await U.bytes(file);
      progress(0.05, 'Analizando el archivo…');
      const L = await U.pdfLibCrypto();
      const openLib = async (b) => {
        try {
          const d = await L.PDFDocument.load(b, { ignoreEncryption: true, throwOnInvalidObject: false });
          return d.isEncrypted || !d.getPageCount() ? null : d;
        } catch (_) { return null; }
      };
      const openJs = async (b) => { try { return await pdfjsLib.getDocument({ data: b.slice(), stopAtErrors: false }).promise; } catch (_) { return null; } };
      // fuentes candidatas: el archivo tal cual y una reconstrucción a partir de sus objetos
      progress(0.08, 'Reconstruyendo la estructura del archivo…');
      const rebuilt = await rebuildPdf(bytes).catch(() => null);
      const libs = [await openLib(bytes), rebuilt && (await openLib(rebuilt))].filter(Boolean);
      const jss = [await openJs(bytes), rebuilt && (await openJs(rebuilt))].filter(Boolean);
      const n = Math.max(0, ...libs.map((d) => d.getPageCount()), ...jss.map((d) => d.numPages));
      if (!n) throw new Error('No se ha podido recuperar ninguna página: el archivo está demasiado dañado o no es un PDF.');
      /** ¿La página se puede dibujar y tiene contenido? */
      const pageOk = async (page) => {
        try {
          const ops = await page.getOperatorList();
          return ops.fnArray.length > 2;
        } catch (_) { return false; }
      };
      const out = await L.PDFDocument.create();
      let ok = 0, raster = 0, lost = 0;
      for (let i = 0; i < n; i++) {
        progress(0.1 + (0.85 * i) / n, `Recuperando página ${i + 1} de ${n}…`);
        let done = false;
        // 1) página original (texto y vectores intactos), comprobada con pdf.js
        for (const src of libs) {
          if (i >= src.getPageCount()) continue;
          try {
            const tmp = await L.PDFDocument.create();
            const [tp] = await tmp.copyPages(src, [i]);
            tmp.addPage(tp);
            const check = await pdfjsLib.getDocument({ data: await tmp.save() }).promise;
            const good = await pageOk(await check.getPage(1));
            check.destroy();
            if (!good) continue;
            const [p] = await out.copyPages(src, [i]);
            out.addPage(p);
            ok++;
            done = true;
            break;
          } catch (_) { /* siguiente fuente */ }
        }
        if (done) continue;
        // 2) rescate como imagen
        for (const js of jss) {
          if (i >= js.numPages) continue;
          try {
            const pg = await js.getPage(i + 1);
            if (!(await pageOk(pg))) continue;
            const { canvas, width, height } = await U.render(js, i + 1, { scale: 2 });
            const img = await out.embedJpg(await (await U.toBlob(canvas, 'image/jpeg', 0.9)).arrayBuffer());
            out.addPage([width, height]).drawImage(img, { x: 0, y: 0, width, height });
            raster++;
            done = true;
            break;
          } catch (_) { /* siguiente fuente */ }
        }
        if (!done) lost++;
      }
      jss.forEach((d) => d.destroy());
      if (!out.getPageCount()) throw new Error('No se ha podido recuperar ninguna página.');
      const parts = [`${ok + raster} de ${n} páginas recuperadas`];
      if (raster) parts.push(`${raster} como imagen`);
      if (lost) parts.push(`${lost} irrecuperable${lost > 1 ? 's' : ''}`);
      return { files: [{ name: `${U.baseName(file.name)}_reparado.pdf`, blob: U.pdfBlob(await out.save()) }], title: 'PDF reparado', message: parts.join(' · ') + '.' };
    },
  },

  /* ───────────── Rotar varios ───────────── */
  {
    id: 'rotar-varios-pdf', name: 'Rotar varios PDF', cat: 'editar', icon: 'rotateMany', color: '#0b7ad1',
    desc: 'Gira todas las páginas de muchos PDF a la vez.',
    accept: 'pdf', multiple: true, mode: 'files', action: 'Rotar PDF',
    options: () => `<div class="field"><label class="lbl">Giro</label>
      ${choice('angle', '90', 'A la derecha', '90° en el sentido de las agujas del reloj', true)}
      ${choice('angle', '270', 'A la izquierda', '90° en sentido contrario')}
      ${choice('angle', '180', 'Darle la vuelta', '180°')}</div>`,
    async run({ files, opts, progress }) {
      const out = [];
      for (let i = 0; i < files.length; i++) {
        progress(i / files.length, `Rotando ${files[i].file.name}…`);
        const doc = await U.pdfLib(await U.bytes(files[i].file));
        doc.getPages().forEach((p) => p.setRotation(degrees((p.getRotation().angle + +opts.angle) % 360)));
        out.push({ name: `${U.baseName(files[i].file.name)}_rotado.pdf`, blob: U.pdfBlob(await doc.save()) });
      }
      return out;
    },
  },

  /* ───────────── Crear formularios ───────────── */
  {
    id: 'crear-formulario', name: 'Crear formulario PDF', cat: 'editar', icon: 'formNew', color: '#6d4ed6',
    desc: 'Añade campos rellenables a un PDF: textos, casillas y listas desplegables.',
    accept: 'pdf', multiple: false, mode: 'editor', action: 'Crear formulario',
    hint: 'Elige un tipo de campo y arrastra sobre la página para dibujarlo. Selecciona un campo para cambiar su nombre y opciones.',
    editor: { tools: ['select', 'ftext', 'fcheck', 'fselect'], styles: false },
    setup({ side, ed }) {
      side.innerHTML = '<div class="field-props"><small class="note">Selecciona un campo para ver sus propiedades.</small></div>';
      const box = side.querySelector('.field-props');
      const tool = TOOLS.find((t) => t.id === 'crear-formulario');
      tool.editor.onSelect = (o) => {
        if (!o || o.type !== 'field') { box.innerHTML = '<small class="note">Selecciona un campo para ver sus propiedades.</small>'; return; }
        box.innerHTML = `
          <div class="field"><label class="lbl" for="fp-name">Nombre del campo</label><input id="fp-name" value="${escHtml(o.name)}"></div>
          ${o.ftype === 'fselect' ? `<div class="field"><label class="lbl" for="fp-opts">Opciones (una por línea)</label><textarea id="fp-opts" rows="4">${escHtml(o.options.join('\n'))}</textarea></div>` : ''}
          ${o.ftype === 'ftext' ? `<label class="check"><input type="checkbox" id="fp-multi" ${o.multiline ? 'checked' : ''}> Varias líneas</label>` : ''}
          <label class="check"><input type="checkbox" id="fp-req" ${o.required ? 'checked' : ''}> Obligatorio</label>`;
        box.querySelector('#fp-name').addEventListener('input', (e) => { o.name = e.target.value.trim() || o.name; ed.refresh(o); });
        box.querySelector('#fp-opts')?.addEventListener('input', (e) => { o.options = e.target.value.split('\n').map((s) => s.trim()).filter(Boolean); });
        box.querySelector('#fp-multi')?.addEventListener('change', (e) => { o.multiline = e.target.checked; });
        box.querySelector('#fp-req').addEventListener('change', (e) => { o.required = e.target.checked; });
      };
    },
    validate: ({ ed }) => (ed.count() ? null : 'Dibuja al menos un campo en el documento.'),
    async run({ file, bytes, pdf, ed }) {
      const doc = await U.pdfLib(bytes);
      const form = doc.getForm();
      const used = new Set(form.getFields().map((f) => f.getName()));
      const unique = (name) => {
        let n = name.replace(/[.\s]+/g, '_') || 'campo';
        let k = 2;
        while (used.has(n)) n = `${name}_${k++}`;
        used.add(n);
        return n;
      };
      const grey = PDFLib.rgb(0.55, 0.55, 0.65);
      let count = 0;
      for (const { page: n, objs } of ed.pages()) {
        const page = doc.getPage(n - 1);
        const vp = (await pdf.getPage(n)).getViewport({ scale: 1 });
        for (const o of objs) {
          if (o.type !== 'field') continue;
          const [x1, y1] = vp.convertToPdfPoint(o.x, o.y);
          const [x2, y2] = vp.convertToPdfPoint(o.x + o.w, o.y + o.h);
          const rect = { x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1), borderColor: grey, borderWidth: 1 };
          const name = unique(o.name);
          let f;
          if (o.ftype === 'ftext') {
            f = form.createTextField(name);
            if (o.multiline) f.enableMultiline();
            f.addToPage(page, rect);
          } else if (o.ftype === 'fcheck') {
            f = form.createCheckBox(name);
            f.addToPage(page, rect);
          } else {
            f = form.createDropdown(name);
            const opts = o.options.length ? o.options : ['Opción 1'];
            try { f.addOptions(opts); } catch (_) { f.addOptions(opts.map(winAnsi)); }
            f.addToPage(page, rect);
          }
          if (o.required) f.enableRequired();
          count++;
        }
      }
      return { files: [{ name: `${U.baseName(file.name)}_formulario.pdf`, blob: U.pdfBlob(await doc.save()) }], message: `${count} campo${count === 1 ? '' : 's'} creado${count === 1 ? '' : 's'}.` };
    },
  },

  /* ───────────── HTML a PDF ───────────── */
  {
    id: 'html-a-pdf', name: 'HTML a PDF', cat: 'convertir', icon: 'code', color: '#e34c26',
    desc: 'Convierte una página o código HTML en un documento PDF.',
    accept: { exts: ['html', 'htm'], label: 'archivo HTML', badge: 'HTML', color: '#e34c26' }, multiple: false, mode: 'panel', action: 'Convertir a PDF',
    options: () => `
      ${select('size', 'Tamaño de página', [['a4', 'A4'], ['letter', 'Carta']], 'a4')}
      ${select('orient', 'Orientación', [['portrait', 'Vertical'], ['landscape', 'Horizontal']], 'portrait')}
      ${select('margin', 'Márgenes', [['36', 'Normales'], ['18', 'Pequeños'], ['0', 'Sin márgenes']], '36')}
      <small class="note">Las páginas web remotas no se pueden abrir directamente desde el navegador: abre el archivo .html o pega su código.</small>`,
    async mount({ main, toast }) {
      const sample = '<h1>Mi documento</h1>\n<p>Escribe o pega aquí tu <b>HTML</b>. También puedes abrir un archivo <code>.html</code>.</p>\n<ul><li>Admite estilos CSS</li><li>Imágenes</li><li>Tablas</li></ul>';
      main.innerHTML = `
        <div class="html-panel">
          <div class="field"><label class="lbl" for="html-src">Código HTML</label>
            <textarea id="html-src" rows="9" spellcheck="false">${escHtml(sample)}</textarea></div>
          <div class="toolbar">
            <label class="btn-sec html-open">Abrir archivo .html<input type="file" accept=".html,.htm,text/html" hidden></label>
            <button type="button" class="btn-sec html-print" title="Usa el diálogo de impresión del navegador: el texto queda seleccionable">Imprimir / Guardar como PDF</button>
          </div>
          <div class="lbl">Vista previa</div>
          <div class="html-frame-wrap"><iframe id="html-frame" title="Vista previa" sandbox="allow-same-origin allow-modals"></iframe></div>
        </div>`;
      const ta = main.querySelector('#html-src');
      const frame = main.querySelector('#html-frame');
      const state = { ta, frame, base: '', isEmpty: () => ta.value === sample };
      const render = () => { frame.srcdoc = `<!DOCTYPE html><html><head><meta charset="utf-8">${state.base}<style>body{font-family:-apple-system,Segoe UI,Arial,sans-serif;margin:0;padding:0;color:#111}img{max-width:100%}</style></head><body>${ta.value}</body></html>`; };
      let t;
      ta.addEventListener('input', () => { clearTimeout(t); t = setTimeout(render, 300); });
      main.querySelector('.html-open input').addEventListener('change', async (e) => {
        const f = e.target.files[0];
        if (f) { ta.value = await f.text(); render(); }
        e.target.value = '';
      });
      main.querySelector('.html-print').addEventListener('click', () => {
        try { frame.contentWindow.focus(); frame.contentWindow.print(); } catch (_) { toast('Tu navegador no permite imprimir la vista previa.'); }
      });
      render();
      return state;
    },
    onFiles: async (state, files) => { state.ta.value = await files[0].text(); state.ta.dispatchEvent(new Event('input')); },
    validate: ({ custom }) => (custom && custom.ta.value.trim() ? null : 'Escribe o abre algún contenido HTML.'),
    async run({ custom, opts, progress }) {
      progress(0.05, 'Preparando…');
      await U.loadScript('vendor/html2canvas.min.js');
      let [PW, PH] = opts.size === 'letter' ? [612, 792] : [595.28, 841.89];
      if (opts.orient === 'landscape') [PW, PH] = [PH, PW];
      const m = +opts.margin || 0;
      const cw = PW - m * 2, ch = PH - m * 2;
      const pxW = Math.round((cw * 96) / 72);
      const frame = custom.frame;
      const wrap = frame.parentElement;
      const prevW = frame.style.width;
      frame.style.width = pxW + 'px';
      wrap.classList.add('measuring');
      await new Promise((r) => setTimeout(r, 400));
      const d = frame.contentDocument;
      await Promise.all([...d.images].map((im) => (im.complete ? 0 : new Promise((r) => { im.onload = im.onerror = r; }))));
      if (d.fonts && d.fonts.ready) await d.fonts.ready;
      const fullH = d.documentElement.scrollHeight;
      progress(0.3, 'Dibujando la página…');
      const scale = 2;
      let canvas;
      try {
        canvas = await html2canvas(d.documentElement, { scale, useCORS: true, backgroundColor: '#ffffff', width: pxW, height: fullH, windowWidth: pxW, windowHeight: fullH, logging: false });
      } finally {
        frame.style.width = prevW;
        wrap.classList.remove('measuring');
      }
      const doc = await PDFLib.PDFDocument.create();
      const sliceH = Math.floor(((ch * 96) / 72) * scale);
      const pages = Math.max(1, Math.ceil(canvas.height / sliceH));
      for (let i = 0; i < pages; i++) {
        progress(0.6 + (0.35 * i) / pages, `Creando página ${i + 1} de ${pages}…`);
        const h = Math.min(sliceH, canvas.height - i * sliceH);
        const part = document.createElement('canvas');
        part.width = canvas.width; part.height = h;
        const px = part.getContext('2d');
        px.fillStyle = '#fff'; px.fillRect(0, 0, part.width, h);
        px.drawImage(canvas, 0, i * sliceH, canvas.width, h, 0, 0, canvas.width, h);
        const img = await doc.embedJpg(await (await U.toBlob(part, 'image/jpeg', 0.92)).arrayBuffer());
        const drawH = (h / sliceH) * ch;
        doc.addPage([PW, PH]).drawImage(img, { x: m, y: PH - m - drawH, width: cw, height: drawH });
      }
      const title = (d.title || d.querySelector('h1')?.textContent || 'pagina').trim().slice(0, 60).replace(/[\\/:*?"<>|]+/g, '');
      return { files: [{ name: `${title || 'pagina'}.pdf`, blob: U.pdfBlob(await doc.save()) }], message: `${pages} página${pages > 1 ? 's' : ''}. Si necesitas texto seleccionable, usa «Imprimir / Guardar como PDF».` };
    },
  },

  /* ───────────── Escanear a PDF ───────────── */
  {
    id: 'escanear-pdf', name: 'Escanear a PDF', cat: 'organizar', icon: 'camera', color: '#0f9488',
    desc: 'Usa la cámara del móvil o del ordenador para escanear documentos a PDF.',
    accept: 'image', multiple: true, mode: 'panel', action: 'Crear PDF',
    options: () => `
      <div class="field"><label class="lbl">Acabado</label>
        ${choice('filter', 'mejorar', 'Documento mejorado', 'Más contraste y fondo más limpio', true)}
        ${choice('filter', 'bn', 'Blanco y negro', 'Ideal para texto: archivos muy ligeros')}
        ${choice('filter', 'gris', 'Escala de grises')}
        ${choice('filter', 'color', 'Original')}
      </div>
      ${select('size', 'Tamaño de página', [['a4', 'A4'], ['letter', 'Carta'], ['fit', 'Igual que la foto']], 'a4')}`,
    async mount({ main, toast }) {
      main.innerHTML = `
        <div class="scan">
          <div class="scan-cam" hidden>
            <video autoplay playsinline muted></video>
            <div class="toolbar"><button type="button" class="pick scan-shot">Capturar página</button>
              <button type="button" class="btn-sec scan-stop">Cerrar cámara</button></div>
          </div>
          <div class="toolbar scan-actions">
            <button type="button" class="btn-sec scan-start">Activar cámara</button>
            <label class="btn-sec">Subir fotos<input type="file" accept="image/*" capture="environment" multiple hidden></label>
          </div>
          <p class="work-hint">Las páginas aparecerán aquí. Puedes girarlas o quitarlas antes de crear el PDF.</p>
          <div class="items scan-pages"></div>
        </div>`;
      const state = { pages: [], stream: null };
      const list = main.querySelector('.scan-pages');
      const video = main.querySelector('video');
      const cam = main.querySelector('.scan-cam');
      const draw = () => {
        list.innerHTML = '';
        state.pages.forEach((p, i) => {
          const el = document.createElement('div');
          el.className = 'item scan-page';
          const prev = rotateCanvas(p.thumb, p.rot);
          el.innerHTML = `<button class="item-x" title="Quitar" aria-label="Quitar">×</button><div class="thumb"></div>
            <div class="item-name">Página ${i + 1}</div>
            <div class="toolbar"><button type="button" data-a="l" title="Girar a la izquierda">↺</button><button type="button" data-a="r" title="Girar a la derecha">↻</button>
            ${i ? '<button type="button" data-a="up" title="Mover antes">←</button>' : ''}</div>`;
          el.querySelector('.thumb').appendChild(prev);
          el.addEventListener('click', (e) => {
            const a = e.target.closest('[data-a]')?.dataset.a;
            if (e.target.closest('.item-x')) state.pages.splice(i, 1);
            else if (a === 'l') p.rot = (p.rot + 270) % 360;
            else if (a === 'r') p.rot = (p.rot + 90) % 360;
            else if (a === 'up') [state.pages[i - 1], state.pages[i]] = [state.pages[i], state.pages[i - 1]];
            else return;
            draw();
          });
          list.appendChild(el);
        });
        if (window.ConverterUI) window.ConverterUI.refreshSummary();
      };
      const addCanvas = (c) => {
        const k = 220 / Math.max(c.width, c.height);
        const t = document.createElement('canvas');
        t.width = Math.round(c.width * k); t.height = Math.round(c.height * k);
        t.getContext('2d').drawImage(c, 0, 0, t.width, t.height);
        state.pages.push({ canvas: c, thumb: t, rot: 0 });
      };
      state.addFiles = async (files) => {
        for (const f of files) { try { addCanvas(await fileToCanvas(f)); } catch (_) { toast(`No se pudo leer «${f.name}».`); } }
        draw();
      };
      const stop = () => { if (state.stream) state.stream.getTracks().forEach((t) => t.stop()); state.stream = null; cam.hidden = true; };
      main.querySelector('.scan-start').addEventListener('click', async () => {
        try {
          state.stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1440 } }, audio: false });
          video.srcObject = state.stream;
          cam.hidden = false;
        } catch (_) {
          toast('No se pudo acceder a la cámara. Revisa los permisos o sube fotos.');
        }
      });
      main.querySelector('.scan-shot').addEventListener('click', () => {
        if (!video.videoWidth) return;
        const c = document.createElement('canvas');
        c.width = video.videoWidth; c.height = video.videoHeight;
        c.getContext('2d').drawImage(video, 0, 0);
        addCanvas(c);
        draw();
      });
      main.querySelector('.scan-stop').addEventListener('click', stop);
      main.querySelector('.scan-actions input').addEventListener('change', (e) => { state.addFiles([...e.target.files]); e.target.value = ''; });
      state.destroy = stop;
      state.pause = stop; // al cambiar de herramienta se apaga la cámara; las fotos se quedan
      state.isEmpty = () => !state.pages.length;
      state.summary = () => `${state.pages.length} página${state.pages.length === 1 ? '' : 's'}`;
      return state;
    },
    onFiles: (state, files) => state.addFiles(files),
    validate: ({ custom }) => (custom && custom.pages.length ? null : 'Captura o sube al menos una página.'),
    async run({ custom, opts, progress }) {
      const doc = await PDFLib.PDFDocument.create();
      const sizes = { a4: [595.28, 841.89], letter: [612, 792] };
      const n = custom.pages.length;
      for (let i = 0; i < n; i++) {
        progress(i / n, `Procesando página ${i + 1} de ${n}…`);
        const p = custom.pages[i];
        const c = scanFilter(rotateCanvas(p.canvas, p.rot), opts.filter || 'mejorar');
        const blob = opts.filter === 'bn' ? await U.toBlob(c, 'image/png') : await U.toBlob(c, 'image/jpeg', 0.85);
        const img = opts.filter === 'bn' ? await doc.embedPng(await blob.arrayBuffer()) : await doc.embedJpg(await blob.arrayBuffer());
        let pw, ph;
        if (opts.size === 'fit') { pw = c.width * 0.36; ph = c.height * 0.36; } // ~200 ppp
        else { [pw, ph] = sizes[opts.size] || sizes.a4; if (c.width > c.height) [pw, ph] = [ph, pw]; }
        const k = Math.min(pw / c.width, ph / c.height);
        const w = c.width * k, h = c.height * k;
        doc.addPage([pw, ph]).drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
      }
      if (custom.destroy) custom.destroy();
      return [{ name: 'Escaneo.pdf', blob: U.pdfBlob(await doc.save()) }];
    },
  },

  /* ───────────── Flujos de trabajo ───────────── */
  {
    id: 'flujos-de-trabajo', name: 'Flujos de trabajo', cat: 'flujos', icon: 'workflow', color: '#7c3aed',
    desc: 'Encadena varias herramientas y aplícalas de una vez a uno o muchos PDF. Guarda tus flujos para reutilizarlos.',
    accept: 'pdf', multiple: true, reorder: true, mode: 'files', action: 'Ejecutar flujo',
    hint: 'Cada paso se aplica a todos los archivos, en el orden de la lista.',
    options: () => '<div id="wf"></div>',
    setupOptions(form, { toast }) {
      const root = form.querySelector('#wf');
      let uid = 0;
      const steps = [];
      root.innerHTML = `
        <div class="field"><label class="lbl">Flujos guardados</label>
          <div class="inline"><select class="wf-saved"><option value="">— Flujo nuevo —</option></select><button type="button" class="wf-del" title="Borrar el flujo guardado">Borrar</button></div></div>
        <div class="wf-steps"></div>
        <div class="field"><label class="lbl">Añadir paso</label>
          <div class="inline"><select class="wf-add">${WF_STEPS.map((s) => `<option value="${s.id}">${s.name}</option>`).join('')}</select><button type="button" class="wf-add-btn">Añadir</button></div></div>
        <div class="field"><label class="lbl">Guardar este flujo</label>
          <div class="inline"><input class="wf-name" placeholder="Nombre del flujo"><button type="button" class="wf-save">Guardar</button></div></div>`;
      const list = root.querySelector('.wf-steps');
      const savedSel = root.querySelector('.wf-saved');
      const fillSaved = () => {
        savedSel.innerHTML = '<option value="">— Flujo nuevo —</option>' + wfSaved().map((w, i) => `<option value="${i}">${escHtml(w.name)}</option>`).join('');
      };
      const readStep = (st) => {
        const o = {};
        for (const el of st.el.querySelectorAll('[name]')) {
          const k = el.name.split('__')[1];
          if (el.type === 'checkbox') o[k] = el.checked;
          else if (el.type === 'radio') { if (el.checked) o[k] = el.value; }
          else if (el.type === 'file') o[k] = el.files[0] || null;
          else o[k] = el.value;
        }
        return o;
      };
      const syncStep = (st) => {
        const o = readStep(st);
        st.el.querySelectorAll('[data-wshow]').forEach((el) => { const [k, v] = el.dataset.wshow.split('='); el.hidden = o[k] !== v; });
      };
      const addStep = (id, values) => {
        const def = WF_STEPS.find((s) => s.id === id);
        if (!def) return;
        const st = { uid: ++uid, def };
        const el = document.createElement('div');
        el.className = 'wf-step';
        el.innerHTML = `<div class="wf-head"><b class="wf-n"></b><span>${def.name}</span>
          <span class="wf-btns"><button type="button" data-a="up" title="Subir">↑</button><button type="button" data-a="down" title="Bajar">↓</button><button type="button" data-a="del" title="Quitar">×</button></span></div>
          <div class="wf-opts opts">${def.options().replace(/data-show=/g, 'data-wshow=')}</div>`;
        // nombres e identificadores únicos por paso
        el.querySelectorAll('[name]').forEach((f) => { f.name = `s${st.uid}__${f.name}`; });
        el.querySelectorAll('[id]').forEach((f) => {
          const nid = `s${st.uid}-${f.id}`;
          el.querySelectorAll(`label[for="${f.id}"]`).forEach((l) => l.setAttribute('for', nid));
          f.id = nid;
        });
        if (values) {
          for (const f of el.querySelectorAll('[name]')) {
            const k = f.name.split('__')[1];
            if (!(k in values)) continue;
            if (f.type === 'checkbox') f.checked = !!values[k];
            else if (f.type === 'radio') f.checked = f.value === values[k];
            else if (f.type !== 'file') f.value = values[k];
          }
        }
        st.el = el;
        el.addEventListener('input', () => syncStep(st));
        el.addEventListener('change', () => syncStep(st));
        el.querySelector('.wf-btns').addEventListener('click', (e) => {
          const a = e.target.closest('[data-a]')?.dataset.a;
          const i = steps.indexOf(st);
          if (a === 'del') steps.splice(i, 1);
          else if (a === 'up' && i > 0) [steps[i - 1], steps[i]] = [steps[i], steps[i - 1]];
          else if (a === 'down' && i < steps.length - 1) [steps[i + 1], steps[i]] = [steps[i], steps[i + 1]];
          draw();
        });
        steps.push(st);
        syncStep(st);
        draw();
      };
      const draw = () => {
        list.innerHTML = steps.length ? '' : '<small class="note">Añade pasos para construir tu flujo.</small>';
        steps.forEach((st, i) => { st.el.querySelector('.wf-n').textContent = `${i + 1}.`; list.appendChild(st.el); });
      };
      root.querySelector('.wf-add-btn').addEventListener('click', () => addStep(root.querySelector('.wf-add').value));
      root.querySelector('.wf-save').addEventListener('click', () => {
        const name = root.querySelector('.wf-name').value.trim();
        if (!name) return toast('Escribe un nombre para el flujo.');
        if (!steps.length) return toast('Añade al menos un paso.');
        const data = wfSaved().filter((w) => w.name !== name);
        // las contraseñas y archivos no se guardan
        data.push({ name, steps: steps.map((st) => { const o = readStep(st); delete o.pw; delete o.pw2; delete o.image; return { id: st.def.id, opts: o }; }) });
        localStorage.setItem(WF_KEY, JSON.stringify(data));
        fillSaved();
        savedSel.value = String(data.length - 1);
        toast(`Flujo «${name}» guardado en este navegador.`);
      });
      savedSel.addEventListener('change', () => {
        const w = savedSel.value === '' ? null : wfSaved()[+savedSel.value];
        steps.length = 0;
        if (w) { w.steps.forEach((s) => addStep(s.id, s.opts)); root.querySelector('.wf-name').value = w.name; }
        else root.querySelector('.wf-name').value = '';
        draw();
      });
      root.querySelector('.wf-del').addEventListener('click', () => {
        const i = savedSel.value;
        if (i === '') return;
        const data = wfSaved();
        data.splice(+i, 1);
        localStorage.setItem(WF_KEY, JSON.stringify(data));
        fillSaved();
      });
      fillSaved();
      draw();
      root.__wf = { steps, readStep };
    },
    validate: ({ files, form }) => {
      const wf = (form || document).querySelector('#wf')?.__wf;
      if (!wf || !wf.steps.length) return 'Añade al menos un paso al flujo.';
      if (!files.length) return 'Añade algún PDF.';
      const pIdx = wf.steps.findIndex((s) => s.def.last);
      if (pIdx >= 0 && pIdx < wf.steps.length - 1) return '«Proteger PDF» tiene que ser el último paso: después el archivo queda cifrado.';
      for (const st of wf.steps) {
        const msg = st.def.validate && st.def.validate(wf.readStep(st), files.map((f) => f.file));
        if (msg) return `Paso «${st.def.name}»: ${msg}`;
      }
      return null;
    },
    async run({ files, progress, form }) {
      const wf = (form || document).querySelector('#wf').__wf;
      let current = files.map((f) => f.file);
      const n = wf.steps.length;
      for (let i = 0; i < n; i++) {
        const st = wf.steps[i];
        const sub = (f, t) => progress((i + Math.min(1, f)) / n, `Paso ${i + 1} de ${n} · ${st.def.name}${t ? ': ' + t : ''}`);
        sub(0);
        const out = await st.def.run(current, wf.readStep(st), sub);
        current = out.map((o) => (o instanceof File ? o : asPdfFile(o)));
      }
      return { files: current.map((f) => ({ name: f.name, blob: f })), title: 'Flujo completado', message: `${n} paso${n > 1 ? 's' : ''} aplicado${n > 1 ? 's' : ''} · ${current.length} archivo${current.length > 1 ? 's' : ''} resultante${current.length > 1 ? 's' : ''}.` };
    },
  }
);
