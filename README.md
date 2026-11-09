# Facturador de alquileres

Página web que genera, a partir del listado mensual de alquileres en Excel, los comprobantes
**X** (liquidaciones, no fiscales) de cada inquilino con la estética de las facturas de ARCA,
y una lista de los **honorarios** que hay que facturar oficialmente en ARCA.

Todo corre en el navegador: el Excel no se sube a ningún servidor.

## Cómo se usa
1. Abrir la página.
2. Elegir el Excel del mes y revisar inquilinos e importes.
3. Tocar **Generar facturas**: se descarga un zip con un PDF por inquilino
   (ORIGINAL y DUPLICADO) y `HACER_EN_ARCA_<mes>.xlsx`.
4. Hacer en ARCA los honorarios listados (cada dato tiene un botón "Copiar").

## Estructura
- `index.html`, `app.js`: la página y la lectura del Excel.
- `factura.js`: dibujo del PDF.
- `lib/`: jsPDF, SheetJS y JSZip, incluidos para no depender de internet.
- `python/`: versión de línea de comandos y lector de facturas PDF de ARCA (`leer_arca.py`).

## Privacidad
El `.gitignore` excluye Excels, PDFs y zips: en este repositorio nunca van datos de inquilinos.

## Pendiente
- Leer el formato real del "Listado mensual alquileres".
- Conectar con Dropbox (leer el Excel y guardar los PDFs en la carpeta del mes).
