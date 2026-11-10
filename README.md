# Facturador de alquileres

Página web que genera, a partir del listado mensual de alquileres en Excel, los comprobantes
**X** (liquidaciones, no fiscales) de cada inquilino con la estética de las facturas de ARCA,
y una lista de los **honorarios** que hay que facturar oficialmente en ARCA.

Todo corre en el navegador: el Excel no se sube a ningún servidor.

## Cómo se usa
1. Abrir la página y elegir el `LISTADO MENSUAL ALQUILERES.xls`.
2. Elegir el mes a liquidar y revisar: la página marca los ajustes de alquiler del mes,
   contratos vencidos, ajustes que tocaban y no se hicieron, filas sin DNI, etc.
   Se puede destildar a quien no se quiera emitir.
3. Tocar **Generar liquidaciones**: se descarga un zip con un PDF por inquilino
   (ORIGINAL y DUPLICADO) y `RESUMEN_<mes>.xlsx` (hojas "Liquidaciones" y "Hacer en ARCA").
4. Hacer en ARCA los honorarios listados (cada dato tiene un botón "Copiar").

## Cómo lee el listado
Hoja `ALQUILERES`; fila 1 con un mes por columna. Cada inquilino empieza en la fila cuya
columna D dice "Alquiler" (columna A: `APELLIDO NOMBRE (Propietario) DNI ...`) y termina en
"TOTAL". Cada concepto con importe en la columna del mes pasa a la liquidación; "Alquiler"
se muestra como "Por cuenta y orden alq. <mes>". Los conceptos que dicen "honor..." van a la
lista de ARCA y no a la liquidación. Los datos del emisor están en `config.js`.

Numeración: AAMM + orden dentro del mes (diciembre 2026: 26120001, 26120002, ...), así no
depende de la PC que se use.

## Estructura
- `index.html`, `app.js`: la página.
- `listado.js`: lectura del listado mensual.
- `config.js`: datos del emisor.
- `factura.js`: dibujo del PDF.
- `lib/`: jsPDF, SheetJS y JSZip, incluidos para no depender de internet.
- `python/`: versión de línea de comandos y lector de facturas PDF de ARCA (`leer_arca.py`).

## Privacidad
El `.gitignore` excluye Excels, PDFs y zips: en este repositorio nunca van datos de inquilinos.

## Pendiente
- Conectar con Dropbox (leer el Excel y guardar los PDFs en la carpeta del mes).
