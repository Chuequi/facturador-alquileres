// Dibuja un comprobante "X" (no fiscal) con la estética de las facturas de ARCA.
// Coordenadas en puntos, medidas desde arriba (igual que el PDF original de ARCA).
/* global jspdf */

const W = 595.28, X0 = 15, X1 = 582;

function fmt(n) {
  // Formato ARCA: sin separador de miles, coma decimal
  return (Math.round(n * 100) / 100).toFixed(2).replace(".", ",");
}

function hoja(doc) {
  const font = (bold, italic) =>
    doc.setFont("helvetica", bold && italic ? "bolditalic" : bold ? "bold" : italic ? "italic" : "normal");
  return {
    txt(x, y, s, size = 8, o = {}) {
      font(o.bold, o.italic);
      doc.setFontSize(size);
      doc.setTextColor(...(o.color || [0, 0, 0]));
      const align = o.align === "r" ? "right" : o.align === "c" ? "center" : "left";
      doc.text(String(s ?? ""), x, y, { align });
      doc.setTextColor(0, 0, 0);
    },
    par(x, y, etiqueta, valor, size = 8, gap = 4) {
      this.txt(x, y, etiqueta, size, { bold: true });
      font(true, false);
      doc.setFontSize(size);
      const w = doc.getTextWidth(etiqueta);
      this.txt(x + w + gap, y, valor, size);
    },
    rect(x, y, w, h, lw = 1, fill = null) {
      doc.setLineWidth(lw);
      doc.setDrawColor(0, 0, 0);
      if (fill) { doc.setFillColor(...fill); doc.rect(x, y, w, h, "FD"); }
      else doc.rect(x, y, w, h, "S");
    },
    linea(x1, y1, x2, y2, lw = 1) { doc.setLineWidth(lw); doc.line(x1, y1, x2, y2); },
  };
}

function pagina(h, copia, emisor, cliente, items, numero, fecha) {
  const pv = String(emisor.punto_venta).padStart(5, "0");
  const nro = String(numero).padStart(8, "0");

  // Cabecera
  h.rect(X0, 23, X1 - X0, 26);
  h.txt(W / 2, 42, copia, 12, { bold: true, align: "c" });
  h.rect(X0, 49, X1 - X0, 121);
  h.linea(298, 90, 298, 170);
  h.rect(274, 49, 48, 41, 1, [255, 255, 255]);
  h.txt(298, 76, emisor.letra || "X", 22, { bold: true, align: "c" });
  h.txt(298, 84.5, "DOC. NO VÁLIDO", 4.6, { bold: true, align: "c" });
  h.txt(298, 88.5, "COMO FACTURA", 4.6, { bold: true, align: "c" });
  h.txt(146, 78, emisor.razon_social, 9, { bold: true, align: "c" });
  h.txt(341, 80, emisor.titulo || "LIQUIDACIÓN", 17, { bold: true });

  h.txt(341, 101, "Punto de Venta:", 7.5, { bold: true });
  h.txt(417, 101.5, pv, 9, { bold: true });
  h.txt(461, 101, "Comp. Nro:", 7.5, { bold: true });
  h.txt(517, 101.5, nro, 9, { bold: true });
  h.txt(341, 117, "Fecha de Emisión:", 8, { bold: true });
  h.txt(428, 117.5, fecha, 9, { bold: true });

  h.par(21, 117, "Razón Social:", emisor.razon_social);
  h.par(21, 141, "Domicilio Comercial:", emisor.domicilio);
  h.par(21, 166, "Condición frente al IVA:", emisor.condicion_iva, 8, 7);
  h.par(341, 139, "CUIT:", emisor.cuit);
  h.par(341, 151, "Ingresos Brutos:", emisor.ingresos_brutos || "", 8, 6);
  h.txt(341, 163, "Fecha de Inicio de Actividades:", 8, { bold: true });
  h.txt(486, 163, emisor.inicio_actividades || "", 8);

  // Período
  h.rect(X0, 170, X1 - X0, 22, 1.3);
  h.txt(21, 186, "Período Facturado Desde:", 9, { bold: true });
  h.txt(159, 186, cliente.periodo_desde || fecha, 9);
  h.txt(232, 186, "Hasta:", 9, { bold: true });
  h.txt(265, 186, cliente.periodo_hasta || fecha, 9);
  h.txt(363, 186, "Fecha de Vto. para el pago:", 9, { bold: true });
  h.txt(495, 186, cliente.vencimiento || fecha, 9);

  // Cliente
  h.rect(X0, 194, X1 - X0, 63);
  h.par(21, 205, `${cliente.tipo_doc}:`, cliente.nro_doc, 7);
  h.par(222, 205, "Apellido y Nombre / Razón Social:", cliente.nombre, 7, 1);
  h.txt(21, 221.5, "Condición frente al IVA:", 7, { bold: true });
  h.txt(131, 221.5, cliente.condicion_iva || "Consumidor Final", 7);
  h.par(312, 221.5, "Domicilio:", cliente.domicilio || "", 7, 1);
  h.txt(21, 241.5, "Condición de venta:", 7, { bold: true });
  h.txt(113, 241.5, cliente.condicion_venta || "Contado", 7);

  // Tabla
  const cols = [[X0, 52], [52, 210], [210, 268], [268, 300], [300, 382], [382, 420], [420, 495], [495, X1]];
  const tit = ["Código", "Producto / Servicio", "Cantidad", "U. Medida", "Precio Unit.", "% Bonif", "Imp. Bonif.", "Subtotal"];
  cols.forEach(([a, b], i) => {
    h.rect(a, 259, b - a, 16, 0.6, [209, 209, 209]);
    if (i < 2) h.txt(a + 4, 270, tit[i], 7, { bold: true });
    else h.txt((a + b) / 2, 270, tit[i], i === 2 ? 7 : 5.8, { bold: true, align: "c" });
  });
  let y = 290, total = 0;
  for (const it of items) {
    const cant = it.cantidad || 1, sub = cant * it.importe;
    total += sub;
    h.txt(57, y, it.concepto, 7);
    h.txt(259, y, fmt(cant), 7, { align: "r" });
    h.txt(268, y, "unidades", 5.8);
    h.txt(382, y, fmt(it.importe), 7, { align: "r" });
    h.txt(392, y, "0,00", 7);
    h.txt(487, y, "0,00", 7, { align: "r" });
    h.txt(579, y, fmt(sub), 7, { align: "r" });
    y += 16;
  }

  // Totales
  h.rect(X0, 516, X1 - X0, 154);
  h.par(445, 562, "Subtotal:", "$", 8, 2);
  h.txt(573, 562, fmt(total), 8, { bold: true, align: "r" });
  h.par(383, 580, "Importe Otros Tributos:", "$", 8, 2);
  h.txt(573, 580, "0,00", 8, { bold: true, align: "r" });
  h.txt(416, 600, "Importe Total:", 9.5, { bold: true });
  h.txt(485, 600, "$", 9.5, { bold: true });
  h.txt(573, 600, fmt(total), 9.5, { bold: true, align: "r" });

  // Pie
  h.txt(286, 695, "Pág. 1/1", 8.5, { bold: true, align: "c" });
  return Math.round(total * 100) / 100;
}

// Devuelve { blob, total }
function crearComprobante(emisor, cliente, items, numero, fecha, copias = ["ORIGINAL", "DUPLICADO"]) {
  const doc = new jspdf.jsPDF({ unit: "pt", format: "a4" });
  doc.setProperties({
    title: `Comprobante X ${String(emisor.punto_venta).padStart(5, "0")}-${String(numero).padStart(8, "0")}`,
    author: emisor.razon_social,
  });
  let total = 0;
  copias.forEach((c, i) => {
    if (i > 0) doc.addPage();
    total = pagina(hoja(doc), c, emisor, cliente, items, numero, fecha);
  });
  return { blob: doc.output("blob"), total };
}

window.Factura = { crearComprobante, fmt };
