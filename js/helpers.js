'use strict';

pdfjsLib.GlobalWorkerOptions.workerSrc =
  'vendor/pdf.worker.min.js';

/** Utilidades compartidas por todas las herramientas. */
const U = {
  fmtSize(b) {
    if (b < 1024) return b + ' B';
    if (b < 1048576) return (b / 1024).toFixed(1).replace('.', ',') + ' KB';
    return (b / 1048576).toFixed(2).replace('.', ',') + ' MB';
  },

  baseName(name) {
    return String(name).replace(/\.[^.]+$/, '') || 'documento';
  },

  async bytes(file) {
    return new Uint8Array(await file.arrayBuffer());
  },

  /** Abre un PDF con pdf-lib (para modificarlo). */
  async pdfLib(bytes) {
    try {
      return await PDFLib.PDFDocument.load(bytes);
    } catch (e) {
      if (/encrypt/i.test(String(e && e.message))) {
        throw new Error('Este PDF está protegido o cifrado y no se puede modificar.');
      }
      throw new Error('No se pudo leer el PDF. Puede que esté dañado.');
    }
  },

  /** Abre un PDF con pdf.js (para renderizarlo o leer su texto). */
  async pdfJs(bytes) {
    try {
      // pdf.js transfiere el buffer al worker, así que le pasamos una copia.
      return await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
    } catch (e) {
      if (e && e.name === 'PasswordException') {
        throw new Error('Este PDF está protegido con contraseña.');
      }
      throw new Error('No se pudo leer el PDF. Puede que esté dañado.');
    }
  },

  /**
   * Renderiza una página a un canvas.
   * - scale: escala fija (1 = 72 ppp)
   * - box: tamaño máximo en px (miniaturas)
   * Devuelve también el tamaño de la página en puntos (con su rotación aplicada).
   */
  async render(pdf, num, { scale, box }) {
    const page = await pdf.getPage(num);
    const base = page.getViewport({ scale: 1 });
    const s = scale || Math.min(box / base.width, box / base.height);
    const vp = page.getViewport({ scale: s });
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.floor(vp.width));
    canvas.height = Math.max(1, Math.floor(vp.height));
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    page.cleanup();
    return { canvas, width: base.width, height: base.height };
  },

  toBlob(canvas, type = 'image/jpeg', quality = 0.92) {
    return new Promise((resolve, reject) =>
      canvas.toBlob(
        (b) => (b ? resolve(b) : reject(new Error('No se pudo generar la imagen.'))),
        type,
        quality
      )
    );
  },

  /** "1-3, 5, 8-9" -> [[1,3],[5,5],[8,9]] (base 1, validado). */
  parseRanges(str, max) {
    const parts = String(str || '').split(/[,;]/).map((s) => s.trim()).filter(Boolean);
    if (!parts.length) throw new Error('Escribe al menos un rango de páginas, por ejemplo: 1-3, 5');
    return parts.map((p) => {
      const m = p.match(/^(\d+)\s*(?:-\s*(\d+))?$/);
      if (!m) throw new Error(`Rango no válido: "${p}". Usa el formato 1-3, 5, 7-9.`);
      const a = +m[1];
      const b = m[2] ? +m[2] : a;
      if (a < 1 || b < 1 || a > max || b > max) {
        throw new Error(`El rango "${p}" está fuera del documento (tiene ${max} páginas).`);
      }
      if (a > b) throw new Error(`El rango "${p}" está al revés. Escribe primero el número menor.`);
      return [a, b];
    });
  },

  color(hex) {
    const n = parseInt(String(hex).replace('#', ''), 16) || 0;
    return PDFLib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
  },

  pdfBlob(bytes) {
    return new Blob([bytes], { type: 'application/pdf' });
  },

  /** Crea un PDF nuevo con las páginas indicadas [{index, rotation}] en ese orden. */
  async pickPages(src, list) {
    const out = await PDFLib.PDFDocument.create();
    const copied = await out.copyPages(src, list.map((p) => p.index));
    copied.forEach((page, i) => {
      const extra = list[i].rotation || 0;
      if (extra) {
        const angle = (((page.getRotation().angle + extra) % 360) + 360) % 360;
        page.setRotation(PDFLib.degrees(angle));
      }
      out.addPage(page);
    });
    return out;
  },

  /** Convierte cualquier imagen a PNG/JPG embebible en un PDF (respeta la orientación EXIF). */
  async imageForPdf(file) {
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      const w = img.naturalWidth;
      const h = img.naturalHeight;
      if (file.type === 'image/png') {
        return { kind: 'png', bytes: await U.bytes(file), w, h };
      }
      const canvas = document.createElement('canvas');
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0);
      const blob = await U.toBlob(canvas, 'image/jpeg', 0.92);
      canvas.width = canvas.height = 0;
      return { kind: 'jpg', bytes: new Uint8Array(await blob.arrayBuffer()), w, h };
    } catch (e) {
      throw new Error(`No se pudo leer la imagen "${file.name}". Formato no compatible.`);
    } finally {
      URL.revokeObjectURL(url);
    }
  },

  async zip(files) {
    const zip = new JSZip();
    const used = new Map();
    for (const f of files) {
      let name = f.name;
      const n = used.get(name) || 0;
      used.set(name, n + 1);
      if (n) name = name.replace(/(\.[^.]+)?$/, ` (${n})$1`);
      zip.file(name, f.blob);
    }
    return zip.generateAsync({ type: 'blob' });
  },

  /** Carga un script externo solo cuando se necesita. */
  loadScript(src) {
    U._scripts = U._scripts || {};
    if (!U._scripts[src]) {
      U._scripts[src] = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolve;
        s.onerror = () => {
          delete U._scripts[src];
          reject(new Error('No se pudo descargar un componente necesario. Revisa tu conexión a internet.'));
        };
        document.head.appendChild(s);
      });
    }
    return U._scripts[src];
  },

  /** Versión ampliada de pdf-lib con cifrado/descifrado (se carga bajo demanda). */
  async pdfLibCrypto() {
    if (U._crypto) return U._crypto;
    const original = window.PDFLib;
    await U.loadScript('vendor/pdf-lib-cantoo.min.js');
    U._crypto = window.PDFLib;
    window.PDFLib = original; // no tocamos la librería que usa el resto de la app
    return U._crypto;
  },

  /** Comprueba con pdf.js si el PDF necesita contraseña y si la dada es correcta. */
  async checkPassword(bytes, password) {
    try {
      const pdf = await pdfjsLib.getDocument({ data: bytes.slice(), password }).promise;
      const n = pdf.numPages;
      pdf.destroy();
      return n;
    } catch (e) {
      if (e && e.name === 'PasswordException') {
        throw new Error(password ? 'La contraseña no es correcta.' : 'Este PDF necesita contraseña para abrirse. Escríbela en el panel.');
      }
      throw new Error('No se pudo leer el PDF. Puede que esté dañado.');
    }
  },

  /** Convierte un objeto de imagen de pdf.js (bitmap o píxeles) en un canvas. */
  pdfImageToCanvas(img) {
    if (!img || !img.width || !img.height) return null;
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    if (img.bitmap) {
      ctx.drawImage(img.bitmap, 0, 0);
      return canvas;
    }
    if (!img.data) return null;
    const id = ctx.createImageData(img.width, img.height);
    const d = img.data;
    const n = img.width * img.height;
    if (img.kind === 3 || d.length === n * 4) id.data.set(d.subarray ? d.subarray(0, n * 4) : d.slice(0, n * 4));
    else if (img.kind === 2 || d.length === n * 3) {
      for (let p = 0, q = 0; p < n; p++, q += 3) { id.data[p * 4] = d[q]; id.data[p * 4 + 1] = d[q + 1]; id.data[p * 4 + 2] = d[q + 2]; id.data[p * 4 + 3] = 255; }
    } else if (img.kind === 1) {
      const rowBytes = (img.width + 7) >> 3;
      for (let y = 0; y < img.height; y++) for (let x = 0; x < img.width; x++) {
        const v = (d[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1 ? 255 : 0;
        const p = (y * img.width + x) * 4;
        id.data[p] = id.data[p + 1] = id.data[p + 2] = v; id.data[p + 3] = 255;
      }
    } else return null;
    ctx.putImageData(id, 0, 0);
    return canvas;
  },

  /** Obtiene un objeto (imagen/fuente) ya cargado por pdf.js para una página. */
  pdfObj(page, name, timeout = 4000) {
    const store = name.startsWith('g_') ? page.commonObjs : page.objs;
    return new Promise((res) => {
      try { store.get(name, res); } catch (_) { res(null); }
      setTimeout(() => res(null), timeout);
    });
  },

  download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
  },
};
