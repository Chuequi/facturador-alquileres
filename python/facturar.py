"""Facturador de alquileres (comprobantes X, no fiscales).

Flujo mensual:
  1) python facturar.py preparar 2026-12
       Lee contratos.xlsx, aplica los ajustes que tocan este mes y crea
       salida/2026-12/liquidacion_2026-12.xlsx para revisar (se pueden corregir
       importes a mano o poner "no" en la columna Emitir).
  2) python facturar.py generar 2026-12 [--fecha 03/12/2026]
       Lee la liquidación revisada, genera un PDF por contrato en salida/2026-12/
       y actualiza contratos.xlsx (importes nuevos, último ajuste, próximo número).
"""
import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from factura_pdf import dibujar_comprobante

BASE = Path(__file__).resolve().parent
CONTRATOS = BASE / "contratos.xlsx"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


# ---------------------------------------------------------------- utilidades
def parse_mes(s):
    """'2026-12' o '12/2026' o una fecha de Excel -> (año, mes)."""
    if isinstance(s, (datetime, date)):
        return s.year, s.month
    s = str(s).strip()
    if "-" in s:
        a, m = s.split("-")[:2]
    else:
        m, a = s.split("/")[-2:]
    return int(a), int(m)


def meses_entre(a, b):
    return (b[0] - a[0]) * 12 + (b[1] - a[1])


def si(v):
    return str(v or "").strip().lower() in ("si", "sí", "s", "x", "1", "true", "verdadero")


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace("$", "").replace(".", "").replace(",", ".").strip())


def leer_tabla(ws):
    filas = list(ws.iter_rows(values_only=True))
    cab = [str(c).strip() if c is not None else "" for c in filas[0]]
    out = []
    for f in filas[1:]:
        if all(v in (None, "") for v in f):
            continue
        out.append(dict(zip(cab, f)))
    return out


def leer_emisor(wb):
    em = {}
    for fila in wb["Emisor"].iter_rows(min_row=2, values_only=True):
        if fila[0]:
            em[str(fila[0]).strip()] = fila[1]
    return em


def formato_hoja(ws, anchos):
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="44546A")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------- preparar
def calcular(concepto, periodo):
    """Devuelve (importe_nuevo, descripcion_ajuste, se_ajusta)."""
    actual = num(concepto.get("importe_actual"))
    tipo = str(concepto.get("tipo_ajuste") or "ninguno").strip().lower()
    valor = num(concepto.get("valor_ajuste"))
    cada = int(num(concepto.get("cada_meses")) or 1)
    ult = concepto.get("ultimo_ajuste")
    if tipo.startswith("ning") or not valor:
        return actual, "", False
    if ult:
        transcurridos = meses_entre(parse_mes(ult), periodo)
        if transcurridos < cada:
            return actual, f"(próximo ajuste en {cada - transcurridos} mes/es)", False
    if tipo.startswith("porc"):
        return round(actual * (1 + valor / 100), 2), f"+{valor:g}%", True
    if tipo.startswith("monto") or tipo.startswith("fijo"):
        return round(actual + valor, 2), f"+${valor:,.2f}", True
    raise ValueError(f"tipo_ajuste desconocido: {tipo!r}")


def preparar(periodo_txt):
    periodo = parse_mes(periodo_txt)
    etiqueta = f"{periodo[0]}-{periodo[1]:02d}"
    wb = load_workbook(CONTRATOS, data_only=True)
    contratos = {str(c["id"]): c for c in leer_tabla(wb["Contratos"]) if si(c.get("activo", "si"))}
    conceptos = leer_tabla(wb["Conceptos"])

    out = Workbook()
    ws = out.active
    ws.title = "Liquidacion"
    cab = ["Emitir", "Contrato", "Inquilino", "Concepto", "Importe anterior", "Ajuste",
           "Importe a facturar", "Incluye IVA", "Fila concepto"]
    ws.append(cab)
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    for i, cpt in enumerate(conceptos, start=2):  # i = fila en la hoja Conceptos
        cid = str(cpt["contrato_id"])
        if cid not in contratos:
            continue
        nuevo, desc, ajusta = calcular(cpt, periodo)
        ws.append(["si", cid, contratos[cid]["inquilino"], cpt["concepto"],
                   num(cpt["importe_actual"]), desc, nuevo,
                   "si" if si(cpt.get("incluye_iva")) else "no", i])
        if ajusta:
            for c in ws[ws.max_row]:
                c.fill = amarillo
    for fila in ws.iter_rows(min_row=2):
        fila[4].number_format = fila[6].number_format = '#,##0.00'
    formato_hoja(ws, [8, 10, 30, 38, 16, 22, 18, 11, 9])
    ws.column_dimensions["I"].hidden = True

    destino = BASE / "salida" / etiqueta
    destino.mkdir(parents=True, exist_ok=True)
    ruta = destino / f"liquidacion_{etiqueta}.xlsx"
    out.save(ruta)
    print(f"Liquidación lista para revisar: {ruta}")
    print("Las filas en amarillo tienen ajuste este mes. Corregí 'Importe a facturar' si hace falta.")


# ---------------------------------------------------------------- generar
def generar(periodo_txt, fecha_txt=None, sin_actualizar=False):
    periodo = parse_mes(periodo_txt)
    etiqueta = f"{periodo[0]}-{periodo[1]:02d}"
    fecha = datetime.strptime(fecha_txt, "%d/%m/%Y").date() if fecha_txt else date.today()
    destino = BASE / "salida" / etiqueta
    liq = leer_tabla(load_workbook(destino / f"liquidacion_{etiqueta}.xlsx", data_only=True)["Liquidacion"])

    wb = load_workbook(CONTRATOS)
    wbv = load_workbook(CONTRATOS, data_only=True)
    emisor = leer_emisor(wbv)
    contratos = {str(c["id"]): c for c in leer_tabla(wbv["Contratos"])}
    mes_txt = f"{MESES[periodo[1] - 1]}/{periodo[0]}"

    # agrupar ítems por contrato respetando el orden
    grupos = {}
    for r in liq:
        if not si(r["Emitir"]):
            continue
        grupos.setdefault(str(r["Contrato"]), []).append(r)

    numero = int(num(emisor.get("proximo_numero")) or 1)
    copias = [c.strip() for c in str(emisor.get("copias") or "ORIGINAL").split(",") if c.strip()]
    ws_cpt = wb["Conceptos"]
    cab_cpt = [c.value for c in ws_cpt[1]]
    col_imp = cab_cpt.index("importe_actual") + 1
    col_ult = cab_cpt.index("ultimo_ajuste") + 1

    resumen = []
    en_arca = []
    for cid, filas in grupos.items():
        ct = contratos[cid]
        todos = [{"concepto": str(r["Concepto"]).replace("{mes}", mes_txt),
                  "importe": num(r["Importe a facturar"]),
                  "incluye_iva": si(r["Incluye IVA"])} for r in filas]
        cliente = {"tipo_doc": ct.get("tipo_doc") or "DNI", "nro_doc": ct.get("nro_doc") or "",
                   "nombre": ct["inquilino"], "condicion_iva": ct.get("condicion_iva") or "Consumidor Final",
                   "domicilio": ct.get("domicilio") or "", "condicion_venta": ct.get("condicion_venta") or "Contado"}
        # Los honorarios se facturan oficialmente en ARCA: van a la lista aparte, no al PDF X
        items = [it for it in todos if not es_honorario(it["concepto"])]
        for it in todos:
            if es_honorario(it["concepto"]):
                en_arca.append((cliente, it))
        if items:
            nombre = "".join(ch for ch in str(ct["inquilino"]) if ch.isalnum() or ch in " -_").strip().replace(" ", "_")
            pdf = destino / f"X_{int(emisor['punto_venta']):05d}_{numero:08d}_{nombre}.pdf"
            total = dibujar_comprobante(pdf, emisor, cliente, items, numero, fecha, copias)
            resumen.append((numero, ct["inquilino"], total, pdf.name))
            numero += 1

        if not sin_actualizar:
            for r in filas:
                fila = int(r["Fila concepto"])
                anterior = num(r["Importe anterior"])
                nuevo = num(r["Importe a facturar"])
                if abs(nuevo - anterior) > 0.004:
                    ws_cpt.cell(fila, col_imp, nuevo)
                    ws_cpt.cell(fila, col_ult, etiqueta)

    if not sin_actualizar:
        ws_em = wb["Emisor"]
        for fila in ws_em.iter_rows(min_row=2):
            if fila[0].value == "proximo_numero":
                fila[1].value = numero
        wb.save(CONTRATOS)

    if en_arca:
        lista_arca(destino / f"HACER_EN_ARCA_{etiqueta}.xlsx", en_arca)
    print(f"{len(resumen)} comprobantes en {destino}")
    for n, inq, tot, nombre in resumen:
        print(f"  {n:08d}  {inq:<35} $ {tot:>14,.2f}  {nombre}")
    print(f"  {'':8}  {'TOTAL':<35} $ {sum(r[2] for r in resumen):>14,.2f}")



def es_honorario(concepto):
    return str(concepto).strip().lower().startswith("honorario")


def lista_arca(ruta, en_arca):
    """Excel 'Estos hacerlos en ARCA': un renglón por honorario, con los datos listos para copiar."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Hacer en ARCA"
    ws.append(["Hecho", "Inquilino", "Tipo doc", "Nro doc", "Condición IVA", "Concepto", "Importe (con IVA)", "Neto sin IVA (para factura A)"])
    for cliente, it in en_arca:
        ws.append(["", cliente["nombre"], cliente["tipo_doc"], str(cliente["nro_doc"]),
                   cliente["condicion_iva"], it["concepto"], fmt_arca(it["importe"]),
                   fmt_arca(round(it["importe"] / 1.21, 2))])
    formato_hoja(ws, [8, 32, 9, 14, 20, 26, 16, 18])
    wb.save(ruta)
    print(f"Estos hacerlos en ARCA ({len(en_arca)}): {ruta}")
    for cliente, it in en_arca:
        print(f"  {cliente['nombre']:<35} {it['concepto']:<25} {fmt_arca(it['importe'])}")


def fmt_arca(n):
    """Como se tipea en ARCA: sin separador de miles, coma decimal (ej. 363000,00)."""
    return f"{n:.2f}".replace(".", ",")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("preparar")
    p1.add_argument("periodo", help="mes a facturar, ej. 2026-12")
    p2 = sub.add_parser("generar")
    p2.add_argument("periodo")
    p2.add_argument("--fecha", help="fecha de emisión dd/mm/aaaa (por defecto hoy)")
    p2.add_argument("--sin-actualizar", action="store_true",
                    help="no modificar contratos.xlsx (para pruebas)")
    a = ap.parse_args()
    if a.cmd == "preparar":
        preparar(a.periodo)
    else:
        generar(a.periodo, a.fecha, a.sin_actualizar)

