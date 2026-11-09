"""Dibuja un comprobante tipo "X" (no fiscal) con la estética de las facturas de ARCA.

Uso como módulo:  dibujar_comprobante(ruta_pdf, emisor, cliente, items, numero, fecha)
"""
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color, black, white

W, H = A4  # 595 x 842 pt
GRIS = Color(0.82, 0.82, 0.82)
GRIS_TEXTO = Color(0.35, 0.35, 0.35)

X0, X1 = 15, 582  # bordes izquierdo y derecho de los recuadros


def fmt(n):
    """Formato de importe como en ARCA: sin separador de miles, coma decimal."""
    return f"{n:.2f}".replace(".", ",")


class Hoja:
    """Wrapper de canvas con coordenadas medidas desde arriba (como en el PDF original)."""

    def __init__(self, c):
        self.c = c

    def txt(self, x, y, s, size=8, bold=False, italic=False, align="l", color=black):
        f = "Helvetica"
        if bold and italic:
            f = "Helvetica-BoldOblique"
        elif bold:
            f = "Helvetica-Bold"
        elif italic:
            f = "Helvetica-Oblique"
        c = self.c
        c.setFont(f, size)
        c.setFillColor(color)
        yy = H - y
        if align == "r":
            c.drawRightString(x, yy, s)
        elif align == "c":
            c.drawCentredString(x, yy, s)
        else:
            c.drawString(x, yy, s)
        c.setFillColor(black)

    def par(self, x, y, etiqueta, valor, size=8, gap=4):
        """'Etiqueta: valor' con la etiqueta en negrita."""
        self.txt(x, y, etiqueta, size, bold=True)
        w = self.c.stringWidth(etiqueta, "Helvetica-Bold", size)
        self.txt(x + w + gap, y, valor, size)

    def rect(self, x, y, w, h, lw=1, fill=None):
        c = self.c
        c.setLineWidth(lw)
        if fill is not None:
            c.setFillColor(fill)
        c.rect(x, H - y - h, w, h, stroke=1, fill=1 if fill is not None else 0)
        c.setFillColor(black)

    def linea(self, x1, y1, x2, y2, lw=1):
        self.c.setLineWidth(lw)
        self.c.line(x1, H - y1, x2, H - y2)


def _pagina(h, copia, emisor, cliente, items, numero, fecha, pag, total_pags):
    # ---------- Cabecera ----------
    h.rect(X0, 23, X1 - X0, 26)                       # fila ORIGINAL
    h.txt(W / 2, 42, copia, 12, bold=True, align="c")
    h.rect(X0, 49, X1 - X0, 121)                      # bloque emisor
    h.linea(298, 90, 298, 170)                        # divisor vertical
    h.rect(274, 49, 48, 41, fill=white)               # recuadro de la letra
    h.txt(298, 76, emisor.get("letra", "X"), 22, bold=True, align="c")
    h.txt(298, 84.5, "DOC. NO VÁLIDO", 4.6, bold=True, align="c")
    h.txt(298, 88.5, "COMO FACTURA", 4.6, bold=True, align="c")

    h.txt(146, 78, emisor["razon_social"], 9, bold=True, align="c")
    h.txt(341, 80, emisor.get("titulo") or "LIQUIDACIÓN", 17, bold=True)

    h.txt(341, 101, "Punto de Venta:", 7.5, bold=True)
    h.txt(417, 101.5, f"{int(emisor['punto_venta']):05d}", 9, bold=True)
    h.txt(461, 101, "Comp. Nro:", 7.5, bold=True)
    h.txt(517, 101.5, f"{int(numero):08d}", 9, bold=True)
    h.txt(341, 117, "Fecha de Emisión:", 8, bold=True)
    h.txt(428, 117.5, fecha.strftime("%d/%m/%Y"), 9, bold=True)

    h.par(21, 117, "Razón Social:", emisor["razon_social"])
    h.par(21, 141, "Domicilio Comercial:", emisor["domicilio"])
    h.par(21, 166, "Condición frente al IVA:", emisor["condicion_iva"], gap=7)
    h.par(341, 139, "CUIT:", str(emisor["cuit"]))
    h.par(341, 151, "Ingresos Brutos:", str(emisor.get("ingresos_brutos", "")), gap=6)
    h.txt(341, 163, "Fecha de Inicio de Actividades:", 8, bold=True)
    h.txt(486, 163, str(emisor.get("inicio_actividades", "")), 8)

    # ---------- Período ----------
    h.rect(X0, 170, X1 - X0, 22, lw=1.3)
    f = fecha.strftime("%d/%m/%Y")
    desde = cliente.get("periodo_desde") or f
    hasta = cliente.get("periodo_hasta") or f
    venc = cliente.get("vencimiento") or f
    h.txt(21, 186, "Período Facturado Desde:", 9, bold=True)
    h.txt(159, 186, desde, 9)
    h.txt(232, 186, "Hasta:", 9, bold=True)
    h.txt(265, 186, hasta, 9)
    h.txt(363, 186, "Fecha de Vto. para el pago:", 9, bold=True)
    h.txt(495, 186, venc, 9)

    # ---------- Cliente ----------
    h.rect(X0, 194, X1 - X0, 63)
    h.par(21, 205, f"{cliente['tipo_doc']}:", str(cliente["nro_doc"]), size=7)
    h.par(222, 205, "Apellido y Nombre / Razón Social:", cliente["nombre"], size=7, gap=1)
    h.txt(21, 221.5, "Condición frente al IVA:", 7, bold=True)
    h.txt(131, 221.5, cliente.get("condicion_iva", "Consumidor Final"), 7)
    h.par(312, 221.5, "Domicilio:", cliente.get("domicilio", ""), size=7, gap=1)
    h.txt(21, 241.5, "Condición de venta:", 7, bold=True)
    h.txt(113, 241.5, cliente.get("condicion_venta", "Contado"), 7)

    # ---------- Tabla de ítems ----------
    cols = [(X0, 52), (52, 210), (210, 268), (268, 300), (300, 382), (382, 420), (420, 495), (495, X1)]
    titulos = ["Código", "Producto / Servicio", "Cantidad", "U. Medida", "Precio Unit.",
               "% Bonif", "Imp. Bonif.", "Subtotal"]
    for (a, b), t in zip(cols, titulos):
        h.rect(a, 259, b - a, 16, lw=0.6, fill=GRIS)
        small = t in ("U. Medida", "Precio Unit.", "% Bonif", "Imp. Bonif.", "Subtotal")
        if t in ("Código", "Producto / Servicio"):
            h.txt(a + 4, 270, t, 7, bold=True)
        else:
            h.txt((a + b) / 2, 270, t, 5.8 if small else 7, bold=True, align="c")

    y = 290
    for it in items:
        cant = it.get("cantidad", 1)
        precio = it["importe"]
        if it.get("codigo"):
            h.txt(21, y, str(it["codigo"]), 7)
        h.txt(57, y, it["concepto"], 7)
        h.txt(259, y, fmt(cant), 7, align="r")
        h.txt(268, y, "unidades", 5.8)
        h.txt(382, y, fmt(precio), 7, align="r")
        h.txt(392, y, "0,00", 7)
        h.txt(487, y, "0,00", 7, align="r")
        h.txt(579, y, fmt(cant * precio), 7, align="r")
        y += 16

    # ---------- Totales ----------
    total = round(sum(it.get("cantidad", 1) * it["importe"] for it in items), 2)
    iva_cont = round(sum(it.get("cantidad", 1) * it["importe"] * 21 / 121
                         for it in items if it.get("incluye_iva")), 2)
    h.rect(X0, 516, X1 - X0, 154)
    h.par(445, 562, "Subtotal:", "$", gap=2)
    h.txt(573, 562, fmt(total), 8, bold=True, align="r")
    h.par(383, 580, "Importe Otros Tributos:", "$", gap=2)
    h.txt(573, 580, "0,00", 8, bold=True, align="r")
    h.txt(416, 600, "Importe Total:", 9.5, bold=True)
    h.txt(485, 600, "$", 9.5, bold=True)
    h.txt(573, 600, fmt(total), 9.5, bold=True, align="r")
    if iva_cont:
        h.txt(19, 626, "Régimen de Transparencia Fiscal al Consumidor (Ley 27.743)", 7.5, bold=True, italic=True)
        h.linea(19, 629, 297, 629, lw=0.8)
        h.txt(145, 644, "IVA Contenido: $", 7.5, bold=True)
        h.txt(287, 644, fmt(iva_cont), 7.5, bold=True, align="r")

    # ---------- Pie ----------
    h.txt(286, 695, f"Pág. {pag}/{total_pags}", 8.5, bold=True, align="c")
    return total


def dibujar_comprobante(ruta_pdf, emisor, cliente, items, numero, fecha,
                        copias=("ORIGINAL", "DUPLICADO")):
    c = canvas.Canvas(str(ruta_pdf), pagesize=A4)
    c.setTitle(f"Comprobante X {int(emisor['punto_venta']):05d}-{int(numero):08d}")
    c.setAuthor(emisor["razon_social"])
    total = 0
    for copia in copias:
        total = _pagina(Hoja(c), copia, emisor, cliente, items, numero, fecha, 1, 1)
        c.showPage()
    c.save()
    return total
