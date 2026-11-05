'use strict';

/* Iconos (trazos SVG 24x24) */
const ICONS = {
  merge: '<path d="M8 4H6a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h2"/><path d="M16 4h2a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-2"/><path d="M9 12h6M12 9v6"/>',
  split: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4 8.12 15.88M14.47 14.48 20 20M8.12 8.12 12 12"/>',
  trash: '<path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M10 11v6M14 11v6"/>',
  extract: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M9 15h6M12 12l3 3-3 3"/>',
  organize: '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
  compress: '<path d="M4 14h6v6M20 10h-6V4M14 10l7-7M3 21l7-7"/>',
  image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="9" cy="9" r="2"/><path d="m21 15-3.1-3.1a2 2 0 0 0-2.8 0L6 21"/>',
  fileDown: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M12 18v-6M9 15l3 3 3-3"/>',
  text: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/>',
  rotate: '<path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/>',
  rotateLeft: '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/>',
  watermark: '<path d="M12 22a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5C6 11.1 5 13 5 15a7 7 0 0 0 7 7z"/>',
  hash: '<path d="M4 9h16M4 15h16M10 3 8 21M16 3l-2 18"/>',
};

/* Fragmentos de formulario reutilizables */
const choice = (name, value, label, sub, checked) => `
  <label class="choice"><input type="radio" name="${name}" value="${value}" ${checked ? 'checked' : ''}>
    <span><b>${label}</b>${sub ? `<small>${sub}</small>` : ''}</span></label>`;

const select = (name, label, opts, selected) => `
  <div class="field"><label class="lbl">${label}</label>
    <select name="${name}">${opts
      .map(([v, t]) => `<option value="${v}" ${v === selected ? 'selected' : ''}>${t}</option>`)
      .join('')}</select></div>`;

const { PDFDocument, StandardFonts, degrees } = PDFLib;

/** Dibuja texto centrado en (cx, cy) y girado `angle` grados. */
function drawCentered(page, text, { cx, cy, font, size, color, opacity, angle }) {
  const w = font.widthOfTextAtSize(text, size);
  const h = size * 0.7;
  const r = (angle * Math.PI) / 180;
  page.drawText(text, {
    x: cx - (w / 2) * Math.cos(r) + (h / 2) * Math.sin(r),
    y: cy - (w / 2) * Math.sin(r) - (h / 2) * Math.cos(r),
    size, font, color, opacity, rotate: degrees(angle),
  });
}

/** Igual que drawCentered pero para imágenes. */
function drawImageCentered(page, img, { cx, cy, w, h, opacity, angle }) {
  const r = (angle * Math.PI) / 180;
  page.drawImage(img, {
    x: cx - (w / 2) * Math.cos(r) + (h / 2) * Math.sin(r),
    y: cy - (w / 2) * Math.sin(r) - (h / 2) * Math.cos(r),
    width: w, height: h, opacity, rotate: degrees(angle),
  });
}

function fontError(e) {
  if (/encode|WinAnsi/i.test(String(e && e.message))) {
    return new Error('El texto contiene caracteres no compatibles (por ejemplo emojis). Usa letras, números y signos normales.');
  }
  return e;
}

/** Extrae las imágenes incrustadas en un PDF (vía pdf.js). */
async function extractImages(pdf, base, png, progress) {
  const out = [];
  const seen = new Set();
  const { OPS } = pdfjsLib;
  const imgOps = new Set([OPS.paintImageXObject, OPS.paintJpegXObject, OPS.paintImageXObjectRepeat].filter(Boolean));
  for (let i = 1; i <= pdf.numPages; i++) {
    progress((i - 1) / pdf.numPages, `Buscando imágenes en la página ${i}…`);
    const page = await pdf.getPage(i);
    const ops = await page.getOperatorList();
    for (let k = 0; k < ops.fnArray.length; k++) {
      if (!imgOps.has(ops.fnArray[k])) continue;
      const name = ops.argsArray[k][0];
      if (typeof name !== 'string' || seen.has(name)) continue;
      seen.add(name);
      const img = await U.pdfObj(page, name);
      if (!img || img.width < 16 || img.height < 16) continue;
      const canvas = U.pdfImageToCanvas(img);
      if (!canvas) continue;
      if (!png) {
        // fondo blanco para las transparencias al pasar a JPG
        const c2 = document.createElement('canvas');
        c2.width = canvas.width; c2.height = canvas.height;
        const x2 = c2.getContext('2d');
        x2.fillStyle = '#fff'; x2.fillRect(0, 0, c2.width, c2.height); x2.drawImage(canvas, 0, 0);
        out.push({ name: `${base}_img${out.length + 1}.jpg`, blob: await U.toBlob(c2, 'image/jpeg', 0.92) });
      } else {
        out.push({ name: `${base}_img${out.length + 1}.png`, blob: await U.toBlob(canvas, 'image/png') });
      }
    }
    page.cleanup();
  }
  return out;
}

const TOOLS = [
  /* ───────────── ORGANIZAR ───────────── */
  {
    id: 'unir-pdf', name: 'Unir PDF', cat: 'organizar', icon: 'merge', color: '#e5484d',
    desc: 'Combina varios PDF en un único documento, en el orden que elijas.',
    accept: 'pdf', multiple: true, reorder: true, mode: 'files', action: 'Unir PDF',
    hint: 'Arrastra las tarjetas para cambiar el orden de los archivos.',
    validate: ({ files }) => (files.length < 2 ? 'Añade al menos 2 archivos PDF para unirlos.' : null),
    async run({ files, progress }) {
      const out = await PDFDocument.create();
      for (let i = 0; i < files.length; i++) {
        progress(i / files.length, `Uniendo ${files[i].file.name}…`);
        const src = await U.pdfLib(await U.bytes(files[i].file));
        const pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach((p) => out.addPage(p));
      }
      progress(1, 'Guardando…');
      return [{ name: 'Converter_unido.pdf', blob: U.pdfBlob(await out.save()) }];
    },
  },
  {
    id: 'dividir-pdf', name: 'Dividir PDF', cat: 'organizar', icon: 'split', color: '#f76b15',
    desc: 'Separa un PDF en varios archivos por rangos de páginas o página a página.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Dividir PDF',
    options: () => `
      <div class="field"><label class="lbl">Modo de división</label>
        ${choice('mode', 'ranges', 'Por rangos', 'Elige qué páginas va en cada archivo', true)}
        ${choice('mode', 'fixed', 'Cada N páginas', 'Trozos del mismo tamaño')}
        ${choice('mode', 'all', 'Todas las páginas', 'Un PDF por cada página')}
      </div>
      <div class="field" data-show="mode=ranges"><label class="lbl" for="o-ranges">Rangos</label>
        <input id="o-ranges" name="ranges" placeholder="Ej.: 1-3, 4, 5-8" autocomplete="off">
        <small>Cada rango se guardará como un PDF.</small>
        <label class="check"><input type="checkbox" name="join"> Unir todos los rangos en un solo PDF</label>
      </div>
      <div class="field" data-show="mode=fixed"><label class="lbl" for="o-n">Páginas por archivo</label>
        <input id="o-n" type="number" name="n" min="1" value="2"></div>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const base = U.baseName(file.name);
      const src = await U.pdfLib(await U.bytes(file));
      const total = src.getPageCount();
      let groups;
      if (opts.mode === 'all') {
        groups = Array.from({ length: total }, (_, i) => [i + 1, i + 1]);
      } else if (opts.mode === 'fixed') {
        const n = Math.max(1, parseInt(opts.n, 10) || 1);
        groups = [];
        for (let a = 1; a <= total; a += n) groups.push([a, Math.min(total, a + n - 1)]);
      } else {
        groups = U.parseRanges(opts.ranges, total);
      }
      const toList = ([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => ({ index: a - 1 + i }));

      if (opts.mode === 'ranges' && opts.join) {
        const doc = await U.pickPages(src, groups.flatMap(toList));
        return [{ name: `${base}_extracto.pdf`, blob: U.pdfBlob(await doc.save()) }];
      }
      const out = [];
      for (let i = 0; i < groups.length; i++) {
        const [a, b] = groups[i];
        progress(i / groups.length, `Creando archivo ${i + 1} de ${groups.length}…`);
        const doc = await U.pickPages(src, toList(groups[i]));
        out.push({ name: `${base}_${a === b ? 'pag' + a : a + '-' + b}.pdf`, blob: U.pdfBlob(await doc.save()) });
      }
      return { files: out, message: `Se han creado ${out.length} archivo${out.length > 1 ? 's' : ''}.` };
    },
  },
  {
    id: 'eliminar-paginas', name: 'Eliminar páginas', cat: 'organizar', icon: 'trash', color: '#d6409f',
    desc: 'Quita las páginas que no necesites de tu PDF.',
    accept: 'pdf', multiple: false, mode: 'pages', action: 'Eliminar páginas',
    pages: { select: 'delete' },
    hint: 'Haz clic en las páginas que quieres eliminar.',
    validate: ({ pages }) => {
      const n = pages.filter((p) => p.selected).length;
      if (!n) return 'Selecciona al menos una página para eliminar.';
      if (n === pages.length) return 'No puedes eliminar todas las páginas.';
      return null;
    },
    async run({ file, bytes, pages }) {
      const src = await U.pdfLib(bytes);
      const doc = await U.pickPages(src, pages.filter((p) => !p.selected));
      return [{ name: `${U.baseName(file.name)}_editado.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'extraer-paginas', name: 'Extraer páginas', cat: 'organizar', icon: 'extract', color: '#8e4ec6',
    desc: 'Saca las páginas que elijas a un PDF nuevo.',
    accept: 'pdf', multiple: false, mode: 'pages', action: 'Extraer páginas',
    pages: { select: 'keep' },
    hint: 'Haz clic en las páginas que quieres extraer.',
    options: () => `<label class="check"><input type="checkbox" name="separate"> Guardar cada página en un PDF distinto</label>`,
    validate: ({ pages }) => (pages.some((p) => p.selected) ? null : 'Selecciona al menos una página para extraer.'),
    async run({ file, bytes, pages, opts }) {
      const base = U.baseName(file.name);
      const src = await U.pdfLib(bytes);
      const sel = pages.filter((p) => p.selected);
      if (!opts.separate) {
        const doc = await U.pickPages(src, sel);
        return [{ name: `${base}_extracto.pdf`, blob: U.pdfBlob(await doc.save()) }];
      }
      const out = [];
      for (const p of sel) {
        const doc = await U.pickPages(src, [p]);
        out.push({ name: `${base}_pag${p.index + 1}.pdf`, blob: U.pdfBlob(await doc.save()) });
      }
      return out;
    },
  },
  {
    id: 'organizar-pdf', name: 'Organizar PDF', cat: 'organizar', icon: 'organize', color: '#3e63dd',
    desc: 'Reordena, gira y elimina páginas visualmente.',
    accept: 'pdf', multiple: false, mode: 'pages', action: 'Guardar cambios',
    pages: { reorder: true, rotate: true, del: true, reset: true },
    hint: 'Arrastra las páginas para reordenarlas. Usa los botones de cada página para girarla o eliminarla.',
    validate: ({ pages }) => (pages.length ? null : 'El documento debe tener al menos una página.'),
    async run({ file, bytes, pages }) {
      const src = await U.pdfLib(bytes);
      const doc = await U.pickPages(src, pages);
      return [{ name: `${U.baseName(file.name)}_organizado.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },

  /* ───────────── OPTIMIZAR ───────────── */
  {
    id: 'comprimir-pdf', name: 'Comprimir PDF', cat: 'optimizar', icon: 'compress', color: '#30a46c',
    desc: 'Reduce el tamaño de tu PDF manteniendo una buena calidad.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Comprimir PDF',
    options: () => `
      <div class="field"><label class="lbl">Nivel de compresión</label>
        ${choice('level', 'baja', 'Baja', 'Máxima calidad, menos reducción')}
        ${choice('level', 'media', 'Recomendada', 'Buen equilibrio entre calidad y tamaño', true)}
        ${choice('level', 'alta', 'Extrema', 'Archivo mínimo, menos calidad')}
      </div>
      <small class="note">Para comprimir, las páginas se convierten en imágenes optimizadas: el texto dejará de poder seleccionarse.</small>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const bytes = await U.bytes(file);
      const L = { baja: { scale: 2, q: 0.8 }, media: { scale: 1.5, q: 0.62 }, alta: { scale: 1.1, q: 0.42 } }[opts.level || 'media'];
      const pdf = await U.pdfJs(bytes);
      const out = await PDFDocument.create();
      for (let i = 1; i <= pdf.numPages; i++) {
        progress((i - 1) / pdf.numPages, `Comprimiendo página ${i} de ${pdf.numPages}…`);
        const { canvas, width, height } = await U.render(pdf, i, { scale: L.scale });
        const jpg = await U.toBlob(canvas, 'image/jpeg', L.q);
        canvas.width = canvas.height = 0;
        const img = await out.embedJpg(await jpg.arrayBuffer());
        out.addPage([width, height]).drawImage(img, { x: 0, y: 0, width, height });
      }
      pdf.destroy();
      progress(1, 'Guardando…');
      const candidates = [await out.save()];
      try {
        candidates.push(await (await U.pdfLib(bytes)).save({ useObjectStreams: true }));
      } catch (_) { /* si no se puede optimizar sin pérdidas, seguimos */ }
      const best = candidates.reduce((a, b) => (b.length < a.length ? b : a));
      const name = `${U.baseName(file.name)}_comprimido.pdf`;
      if (best.length >= bytes.length) {
        return {
          files: [{ name, blob: U.pdfBlob(bytes) }],
          title: 'Tu PDF ya estaba optimizado',
          message: 'No hemos conseguido reducirlo más. Prueba con el nivel «Extrema».',
        };
      }
      const pct = Math.round((1 - best.length / bytes.length) * 100);
      return {
        files: [{ name, blob: U.pdfBlob(best) }],
        title: `¡Tu PDF ocupa un ${pct}% menos!`,
        message: `${U.fmtSize(bytes.length)} → ${U.fmtSize(best.length)}`,
      };
    },
  },

  /* ───────────── CONVERTIR ───────────── */
  {
    id: 'pdf-a-jpg', name: 'PDF a JPG', cat: 'convertir', icon: 'image', color: '#ffb224',
    desc: 'Convierte cada página en una imagen JPG o PNG, o extrae las imágenes que contiene.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Convertir a imagen',
    options: () => `
      <div class="field"><label class="lbl">Qué quieres obtener</label>
        ${choice('what', 'pages', 'Páginas a imagen', 'Cada página se convierte en una imagen', true)}
        ${choice('what', 'images', 'Extraer imágenes', 'Saca las fotos e imágenes que hay dentro')}
      </div>
      <div class="field"><label class="lbl">Formato</label>
        ${choice('format', 'jpg', 'JPG', 'Archivos más ligeros', true)}
        ${choice('format', 'png', 'PNG', 'Sin pérdida de calidad')}
      </div>
      <div data-show="what=pages">
      ${select('quality', 'Resolución', [['1.5', 'Normal (108 ppp)'], ['2.78', 'Alta (200 ppp)'], ['4.17', 'Muy alta (300 ppp)']], '2.78')}
      </div>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const base = U.baseName(file.name);
      const png = opts.format === 'png';
      const pdf = await U.pdfJs(await U.bytes(file));
      if (opts.what === 'images') {
        const out = await extractImages(pdf, base, png, progress);
        pdf.destroy();
        if (!out.length) throw new Error('Este PDF no contiene imágenes que se puedan extraer.');
        return { files: out, message: `Se han extraído ${out.length} ${out.length > 1 ? "imágenes" : "imagen"}.` };
      }
      const out = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        progress((i - 1) / pdf.numPages, `Convirtiendo página ${i} de ${pdf.numPages}…`);
        const { canvas } = await U.render(pdf, i, { scale: parseFloat(opts.quality) || 2 });
        const blob = await U.toBlob(canvas, png ? 'image/png' : 'image/jpeg', 0.9);
        canvas.width = canvas.height = 0;
        out.push({ name: `${base}_pag${i}.${png ? 'png' : 'jpg'}`, blob });
      }
      pdf.destroy();
      return out;
    },
  },
  {
    id: 'jpg-a-pdf', name: 'JPG a PDF', cat: 'convertir', icon: 'fileDown', color: '#ffc53d',
    desc: 'Convierte imágenes JPG, PNG, WEBP… en un PDF.',
    accept: 'image', multiple: true, reorder: true, mode: 'files', action: 'Convertir a PDF',
    hint: 'Arrastra las imágenes para cambiar su orden en el PDF.',
    options: () => `
      ${select('size', 'Tamaño de página', [['a4', 'A4'], ['letter', 'Carta (US Letter)'], ['fit', 'Igual que la imagen']], 'a4')}
      ${select('orient', 'Orientación', [['auto', 'Automática'], ['portrait', 'Vertical'], ['landscape', 'Horizontal']], 'auto')}
      ${select('margin', 'Márgenes', [['0', 'Sin márgenes'], ['20', 'Pequeños'], ['50', 'Grandes']], '20')}`,
    async run({ files, opts, progress }) {
      const doc = await PDFDocument.create();
      const sizes = { a4: [595.28, 841.89], letter: [612, 792] };
      const margin = parseFloat(opts.margin) || 0;
      for (let i = 0; i < files.length; i++) {
        progress(i / files.length, `Añadiendo imagen ${i + 1} de ${files.length}…`);
        const im = await U.imageForPdf(files[i].file);
        const img = im.kind === 'png' ? await doc.embedPng(im.bytes) : await doc.embedJpg(im.bytes);
        let pw, ph;
        if (opts.size === 'fit') {
          pw = im.w * 0.75 + margin * 2;
          ph = im.h * 0.75 + margin * 2;
        } else {
          [pw, ph] = sizes[opts.size] || sizes.a4;
          const landscape = opts.orient === 'landscape' || (opts.orient === 'auto' && im.w > im.h);
          if (landscape) [pw, ph] = [ph, pw];
        }
        const s = Math.min((pw - margin * 2) / im.w, (ph - margin * 2) / im.h);
        const w = im.w * s;
        const h = im.h * s;
        doc.addPage([pw, ph]).drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
      }
      progress(1, 'Guardando…');
      const name = files.length === 1 ? `${U.baseName(files[0].file.name)}.pdf` : 'Converter_imagenes.pdf';
      return [{ name, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'pdf-a-texto', name: 'PDF a texto', cat: 'convertir', icon: 'text', color: '#12a594',
    desc: 'Extrae todo el texto de un PDF a un archivo .txt.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Extraer texto',
    options: () => `<label class="check"><input type="checkbox" name="marks" checked> Indicar dónde empieza cada página</label>
      <small class="note">Los PDF escaneados (fotos de papel) no contienen texto que se pueda extraer.</small>`,
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const pdf = await U.pdfJs(await U.bytes(file));
      const parts = [];
      let chars = 0;
      for (let i = 1; i <= pdf.numPages; i++) {
        progress((i - 1) / pdf.numPages, `Leyendo página ${i} de ${pdf.numPages}…`);
        const page = await pdf.getPage(i);
        const tc = await page.getTextContent();
        const txt = tc.items.map((it) => (it.str || '') + (it.hasEOL ? '\n' : '')).join('').trim();
        chars += txt.length;
        parts.push(opts.marks ? `──────── Página ${i} ────────\n\n${txt}` : txt);
      }
      pdf.destroy();
      if (!chars) throw new Error('Este PDF no contiene texto seleccionable (probablemente sea un documento escaneado).');
      const blob = new Blob([parts.join('\n\n')], { type: 'text/plain;charset=utf-8' });
      return { files: [{ name: `${U.baseName(file.name)}.txt`, blob }], message: `${chars.toLocaleString('es')} caracteres extraídos.` };
    },
  },

  /* ───────────── EDITAR ───────────── */
  {
    id: 'rotar-pdf', name: 'Rotar PDF', cat: 'editar', icon: 'rotate', color: '#0090ff',
    desc: 'Gira todas las páginas o solo las que quieras.',
    accept: 'pdf', multiple: false, mode: 'pages', action: 'Rotar PDF',
    pages: { rotate: true, rotateAll: true },
    hint: 'Gira todas las páginas a la vez o pulsa el botón de cada página.',
    validate: ({ pages }) => (pages.some((p) => p.rotation) ? null : 'Gira al menos una página.'),
    async run({ file, bytes, pages }) {
      const src = await U.pdfLib(bytes);
      const doc = await U.pickPages(src, pages);
      return [{ name: `${U.baseName(file.name)}_rotado.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'marca-de-agua', name: 'Marca de agua', cat: 'editar', icon: 'watermark', color: '#6e56cf',
    desc: 'Añade un texto o una imagen (como tu logo) como marca de agua.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Añadir marca de agua',
    options: () => `
      <div class="field"><label class="lbl">Tipo</label>
        ${choice('kind', 'text', 'Texto', '', true)}
        ${choice('kind', 'image', 'Imagen', 'Por ejemplo, tu logotipo')}
      </div>
      <div data-show="kind=text" class="opts">
        <div class="field"><label class="lbl" for="o-text">Texto</label>
          <input id="o-text" name="text" value="CONFIDENCIAL" maxlength="80"></div>
        <div class="row">
          <div class="field"><label class="lbl" for="o-size">Tamaño</label>
            <input id="o-size" type="number" name="size" min="8" max="200" value="60"></div>
          <div class="field"><label class="lbl" for="o-color">Color</label>
            <input id="o-color" type="color" name="color" value="#e5484d"></div>
        </div>
      </div>
      <div data-show="kind=image" class="opts">
        <div class="field"><label class="lbl" for="o-img">Imagen</label>
          <input id="o-img" type="file" name="image" accept="image/*"></div>
        ${select('scale', 'Tamaño de la imagen', [['0.25', 'Pequeña (25% del ancho)'], ['0.5', 'Mediana (50%)'], ['0.8', 'Grande (80%)']], '0.5')}
      </div>
      <div class="field"><label class="lbl" for="o-op">Opacidad</label>
        <input id="o-op" type="range" name="opacity" min="0.05" max="1" step="0.05" value="0.3"></div>
      ${select('angle', 'Inclinación', [['45', 'Diagonal (45°)'], ['0', 'Horizontal'], ['-45', 'Diagonal inversa']], '45')}
      ${select('layout', 'Posición', [['center', 'Centrada'], ['tile', 'Mosaico (repetida)']], 'center')}`,
    validate: ({ opts }) => {
      if (opts.kind === 'image') return opts.image ? null : 'Elige la imagen para la marca de agua.';
      return opts.text.trim() ? null : 'Escribe el texto de la marca de agua.';
    },
    async run({ files, opts, progress }) {
      const file = files[0].file;
      const doc = await U.pdfLib(await U.bytes(file));
      const font = await doc.embedFont(StandardFonts.HelveticaBold);
      const text = opts.text.trim();
      const size = Math.min(200, Math.max(8, parseFloat(opts.size) || 60));
      const style = { font, size, color: U.color(opts.color), opacity: parseFloat(opts.opacity) || 0.3 };
      const pages = doc.getPages();
      if (opts.kind === 'image') {
        let im;
        try {
          im = await Editor.imageToPng(opts.image);
        } catch (_) {
          throw new Error('No se pudo leer la imagen elegida.');
        }
        const img = await doc.embedPng(await (await fetch(im.url)).arrayBuffer());
        const opacity = parseFloat(opts.opacity) || 0.3;
        pages.forEach((page) => {
          const box = page.getCropBox();
          const angle = parseFloat(opts.angle) + page.getRotation().angle;
          const w = Math.min(box.width, box.height) * (parseFloat(opts.scale) || 0.5);
          const h = (w * im.h) / im.w;
          if (opts.layout === 'tile') {
            const diag = Math.hypot(box.width, box.height);
            for (let y = box.y - diag / 2; y < box.y + box.height + diag / 2; y += h * 1.8) {
              for (let x = box.x - diag / 2; x < box.x + box.width + diag / 2; x += w * 1.6) {
                drawImageCentered(page, img, { cx: x, cy: y, w, h, opacity, angle });
              }
            }
          } else {
            drawImageCentered(page, img, { cx: box.x + box.width / 2, cy: box.y + box.height / 2, w, h, opacity, angle });
          }
        });
        return [{ name: `${U.baseName(file.name)}_marca.pdf`, blob: U.pdfBlob(await doc.save()) }];
      }
      try {
        pages.forEach((page, i) => {
          progress(i / pages.length, `Página ${i + 1} de ${pages.length}…`);
          const box = page.getCropBox();
          // Compensamos la rotación propia de la página para que el texto se vea como se espera.
          const angle = parseFloat(opts.angle) + page.getRotation().angle;
          if (opts.layout === 'tile') {
            const tw = font.widthOfTextAtSize(text, size);
            const stepX = tw + size * 2;
            const stepY = size * 3.5;
            const diag = Math.hypot(box.width, box.height);
            for (let y = box.y - diag / 2; y < box.y + box.height + diag / 2; y += stepY) {
              const row = Math.round((y - box.y) / stepY);
              for (let x = box.x - diag / 2 + (row % 2 ? stepX / 2 : 0); x < box.x + box.width + diag / 2; x += stepX) {
                drawCentered(page, text, { ...style, cx: x, cy: y, angle });
              }
            }
          } else {
            drawCentered(page, text, { ...style, cx: box.x + box.width / 2, cy: box.y + box.height / 2, angle });
          }
        });
      } catch (e) {
        throw fontError(e);
      }
      return [{ name: `${U.baseName(file.name)}_marca.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
  {
    id: 'numeros-de-pagina', name: 'Números de página', cat: 'editar', icon: 'hash', color: '#ad7f58',
    desc: 'Numera las páginas de tu PDF con el formato y la posición que quieras.',
    accept: 'pdf', multiple: false, mode: 'files', action: 'Numerar páginas',
    options: () => `
      ${select('pos', 'Posición', [['bc', 'Abajo, centro'], ['br', 'Abajo, derecha'], ['bl', 'Abajo, izquierda'], ['tc', 'Arriba, centro'], ['tr', 'Arriba, derecha'], ['tl', 'Arriba, izquierda']], 'bc')}
      ${select('fmt', 'Formato', [['n', '1'], ['n/N', '1 / 10'], ['pag', 'Página 1'], ['pagN', 'Página 1 de 10']], 'n')}
      <div class="row">
        <div class="field"><label class="lbl" for="o-start">Empezar en</label>
          <input id="o-start" type="number" name="start" min="0" value="1"></div>
        <div class="field"><label class="lbl" for="o-fs">Tamaño</label>
          <input id="o-fs" type="number" name="size" min="6" max="48" value="11"></div>
      </div>
      <label class="check"><input type="checkbox" name="skipFirst"> No numerar la primera página (portada)</label>`,
    async run({ files, opts }) {
      const file = files[0].file;
      const doc = await U.pdfLib(await U.bytes(file));
      const font = await doc.embedFont(StandardFonts.Helvetica);
      const size = Math.min(48, Math.max(6, parseFloat(opts.size) || 11));
      const start = parseInt(opts.start, 10);
      const first = Number.isFinite(start) ? start : 1;
      const pages = doc.getPages();
      const skip = opts.skipFirst ? 1 : 0;
      const total = first + pages.length - skip - 1;
      const label = (n) => ({ n: `${n}`, 'n/N': `${n} / ${total}`, pag: `Página ${n}`, pagN: `Página ${n} de ${total}` })[opts.fmt] || `${n}`;
      const [v, h] = opts.pos;
      const m = 28;
      pages.forEach((page, i) => {
        if (i < skip) return;
        const txt = label(first + i - skip);
        const box = page.getCropBox();
        const rot = ((page.getRotation().angle % 360) + 360) % 360;
        // posición calculada sobre la página tal como se ve (con su giro)…
        const VW = rot % 180 ? box.height : box.width;
        const VH = rot % 180 ? box.width : box.height;
        const w = font.widthOfTextAtSize(txt, size);
        const vx = h === 'l' ? m : h === 'r' ? VW - m - w : (VW - w) / 2;
        const vy = v === 't' ? VH - m - size * 0.7 : m;
        // …y llevada a las coordenadas del PDF
        const [px, py] = rot === 90 ? [box.width - vy, vx] : rot === 180 ? [box.width - vx, box.height - vy] : rot === 270 ? [vy, box.height - vx] : [vx, vy];
        page.drawText(txt, { x: box.x + px, y: box.y + py, size, font, color: PDFLib.rgb(0.2, 0.2, 0.2), rotate: degrees(rot) });
      });
      return [{ name: `${U.baseName(file.name)}_numerado.pdf`, blob: U.pdfBlob(await doc.save()) }];
    },
  },
];
