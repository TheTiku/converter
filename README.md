# Converter

Herramientas PDF que funcionan en el navegador: unir, dividir, comprimir, convertir (Word, Excel, PowerPoint, imágenes, HTML), editar, firmar, proteger, OCR y más. Todo se procesa en el dispositivo del usuario: los archivos no se suben a ningún servidor.

Es una aplicación web instalable (PWA) que funciona sin conexión.
https://thetiku.github.io/converter/

## Probar en local

```bash
python3 -m http.server 8080
```

Abre `http://localhost:8080` en Chrome o Edge.

## Publicar cambios

Al cambiar archivos de la app, sube `VERSION` en `sw.js` para que los usuarios reciban la versión nueva guardada sin conexión.

## Estructura

- `index.html`, `css/`, `js/` — la aplicación
- `vendor/` — librerías incluidas (pdf-lib, pdf.js, JSZip, SheetJS, PptxGenJS, fontkit, html2canvas, Tesseract.js)
- `sw.js`, `manifest.webmanifest`, `icons/` — instalación y uso sin conexión
