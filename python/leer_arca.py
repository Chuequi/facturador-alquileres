"""Lee facturas PDF descargadas de ARCA y arma la lista de clientes (inquilinos).

Uso:  python leer_arca.py CARPETA_CON_PDFS [salida.xlsx]

Por cada PDF extrae: tipo/número, fecha, documento, nombre, condición IVA,
domicilio y los renglones (concepto + importe). Genera un Excel con dos hojas:
  - Clientes: una fila por documento (DNI/CUIT), con los datos tal como figuran en ARCA.
  - Facturas: una fila por renglón facturado, para comparar con el listado mensual.
"""
import re
import sys
from pathlib import Path

import pypdf
from openpyxl import Workbook

from facturar import formato_hoja

SEP = re.compile(r"\s{2,}")


def importe(s):
    return float(s.replace(".", "").replace(",", ".")) if "," in s else float(s)


def leer_factura(ruta):
    texto = pypdf.PdfReader(ruta).pages[0].extract_text(extraction_mode="layout")
    lineas = [SEP.split(l.strip()) for l in texto.splitlines() if l.strip()]
    f = {"archivo": Path(ruta).name, "items": []}
    en_items = False
    for i, p in enumerate(lineas):
        linea = "  ".join(p)
        if len(p) >= 2 and p[1] == "FACTURA" or (len(p) >= 3 and p[-1] == "FACTURA"):
            f["letra"] = p[-2]
        if "Comp. Nro:" in linea:
            f["punto_venta"], f["numero"] = p[1], p[3]
        if "Fecha de Emisión:" in linea:
            f["fecha"] = p[-1]
        # El primer "CUIT:"/"DNI:" al inicio de línea después del período es el cliente
        m = re.match(r"(CUIT|DNI|CUIL|CDI|Pasaporte):\s*(\S+)\s+Apellido y Nombre / Razón Social:\s*(.*)", linea.replace("  ", " "))
        if m and "nro_doc" not in f:
            f["tipo_doc"], f["nro_doc"], f["nombre"] = m.group(1), m.group(2), m.group(3).strip()
            sig = "  ".join(lineas[i + 1])
            m2 = re.match(r"Condición frente al IVA:\s+(.*?)\s{2,}Domicilio(?: Comercial)?:\s+(.*)", sig)
            if m2:
                f["condicion_iva"], f["domicilio"] = m2.group(1).strip(), m2.group(2).strip()
        if p[0] == "Código":
            en_items = True
            continue
        if en_items:
            if p[0].startswith(("Importe", "Subtotal")):
                en_items = False
                continue
            if len(p) >= 4 and re.fullmatch(r"\d+,\d\d", p[1]):
                f["items"].append({"concepto": p[0], "importe": importe(p[3])})
        if p[0] == "Importe Total: $":
            f["total"] = importe(p[1])
    return f


def main(carpeta, salida):
    facturas = []
    for pdf in sorted(Path(carpeta).glob("*.pdf")):
        try:
            facturas.append(leer_factura(pdf))
        except Exception as e:  # un PDF raro no frena al resto
            print(f"No pude leer {pdf.name}: {e}")

    wb = Workbook()
    cl = wb.active
    cl.title = "Clientes"
    cl.append(["tipo_doc", "nro_doc", "nombre", "condicion_iva", "domicilio", "ultima_factura"])
    vistos = {}
    for f in facturas:
        if "nro_doc" in f:
            vistos[f["nro_doc"]] = f  # se queda con la última
    for f in vistos.values():
        cl.append([f["tipo_doc"], f["nro_doc"], f["nombre"], f.get("condicion_iva"), f.get("domicilio"),
                   f"{f.get('letra', '')} {f.get('punto_venta', '')}-{f.get('numero', '')}"])
    formato_hoja(cl, [9, 14, 34, 24, 42, 20])

    fa = wb.create_sheet("Facturas")
    fa.append(["archivo", "letra", "numero", "fecha", "nro_doc", "nombre", "concepto", "importe", "total_factura"])
    for f in facturas:
        for it in f["items"]:
            fa.append([f["archivo"], f.get("letra"), f.get("numero"), f.get("fecha"), f.get("nro_doc"),
                       f.get("nombre"), it["concepto"], it["importe"], f.get("total")])
        suma = round(sum(it["importe"] for it in f["items"]), 2)
        if f.get("total") is not None and abs(suma - f["total"]) > 0.01 and f.get("letra") != "A":
            print(f"Atención {f['archivo']}: los renglones suman {suma} y el total es {f['total']}")
    for row in fa.iter_rows(min_row=2):
        row[7].number_format = row[8].number_format = "#,##0.00"
    formato_hoja(fa, [38, 6, 11, 11, 13, 30, 34, 14, 14])
    wb.save(salida)
    print(f"{len(facturas)} facturas leídas, {len(vistos)} clientes -> {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "clientes_arca.xlsx")
