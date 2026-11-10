// Lee la hoja "ALQUILERES" del "Listado mensual alquileres".
//
// Formato de la planilla (así la usa la inmobiliaria):
//   - Fila 1: desde la columna E, un mes por columna (fecha de Excel).
//   - Cada inquilino es un bloque de filas. El bloque empieza en la fila cuya columna D dice
//     "Alquiler" (o "ALQUILER", "Alquiler + coch.", ...). En esa fila, la columna A tiene
//     "APELLIDO NOMBRE (Propietario) DNI 12.345.678".
//   - Columna D: concepto (Tasa, Aguas, Expensas, com. Bco., Honorarios 2/3, ...).
//     Columna del mes: importe. El bloque termina en la fila "TOTAL" (a veces sin rótulo).
//   - Columna A, filas siguientes: domicilio, notas ("FACTURA A"...) y "VTO. dd/mm/aaaa"
//     (fin del contrato). Columnas B/C: "Cuatrimestral", "Vto. 01/11/2026" (fechas de ajuste).
(function (raiz) {
  const s = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const esHonorario = (c) => /honor/i.test(c);

  function columnaDelMes(cab, anio, mes, XLSX) {
    for (let j = 0; j < cab.length; j++) {
      const v = cab[j];
      if (typeof v === "number" && v > 30000) {
        const d = XLSX.SSF.parse_date_code(v);
        if (d.y === anio && d.m === mes) return j;
      }
    }
    return -1;
  }

  // "ACEVEDO ANDRES( Raimondi) DNI 23.622.453" -> {nombre, propietario, tipo_doc, nro_doc}
  function parseEncabezado(txt) {
    const t = s(txt);
    const propietario = (t.match(/\(([^)]*)\)/) || [])[1]?.trim() || "";
    let tipo = /CUIT|CUIL/i.test(t) ? "CUIT" : "DNI";
    // número: el primer grupo de 7 a 11 dígitos (permitiendo puntos, comas, guiones y espacios)
    const m = t.replace(/\([^)]*\)/g, " ").match(/(\d[\d.,\- ]{5,}\d)/);
    let nro = m ? m[1].replace(/\D/g, "") : "";
    if (nro.length === 11) tipo = "CUIT";
    const nombre = s(t.replace(/\([^)]*\)/g, " ")
      .split(/\b(D\.?\s?N\.?\s?I\.?|CUIT|CUIL|dni)\b|\d{1,3}[.,]\d{3}/i)[0]
      .replace(/[-–:,]+\s*$/, ""));
    return { nombre, propietario, tipo_doc: tipo, nro_doc: nro };
  }

  function fecha(txt) {
    const m = s(txt).match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{2,4})/);
    if (!m) return null;
    let a = +m[3]; if (a < 100) a += 2000;
    return { d: +m[1], m: +m[2], a };
  }

  // Devuelve { mesEncontrado, inquilinos: [...], omitidos: [...] }
  function leerListado(wb, anio, mes, XLSX) {
    const nombreHoja = wb.SheetNames.find((n) => /^alquileres$/i.test(n.trim()));
    if (!nombreHoja) throw new Error('No encontré la hoja "ALQUILERES".');
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { header: 1, defval: "", raw: true });
    const col = columnaDelMes(filas[0], anio, mes, XLSX);
    if (col < 0) throw new Error(`La planilla no tiene la columna de ${String(mes).padStart(2, "0")}/${anio}.`);

    const bloques = [];
    let b = null;
    filas.forEach((r, i) => {
      if (i === 0) return;
      const a = s(r[0]), d = s(r[3]), v = r[col], vAnt = r[col - 1];
      if (/^alq/i.test(d)) {
        b = { fila: i + 1, ...parseEncabezado(a), encabezado: a, lineas: [], notas: [], ajustes: [],
              alquilerAnterior: typeof vAnt === "number" ? vAnt : null, cerrado: false };
        bloques.push(b);
      }
      if (!b || b.cerrado) return;
      if (i + 1 !== b.fila && a) b.notas.push(a);
      for (const x of [s(r[1]), s(r[2])]) if (x) (fecha(x) ? b.ajustes : b.notas).push(x);
      if (/^total/i.test(d)) { b.total = typeof v === "number" ? v : null; b.cerrado = true; return; }
      if (typeof v === "number" && v !== 0) {
        if (!d) {
          // fila sin rótulo: si es la suma de lo anterior, es el TOTAL
          const suma = b.lineas.reduce((t, x) => t + x.importe, 0);
          const sinHon = b.lineas.filter((x) => !esHonorario(x.concepto)).reduce((t, x) => t + x.importe, 0);
          if (Math.abs(v - suma) < 1 || Math.abs(v - sinHon) < 1) { b.total = v; b.cerrado = true; return; }
          b.lineas.push({ concepto: "", importe: v, fila: i + 1 });
        } else {
          b.lineas.push({ concepto: d, importe: v, fila: i + 1 });
        }
      }
    });

    const inquilinos = [], omitidos = [];
    for (const x of bloques) {
      const alq = x.lineas.find((l) => /^alq/i.test(l.concepto));
      if (!x.nombre || !alq) {
        omitidos.push({ fila: x.fila, encabezado: x.encabezado || "(sin nombre)",
          motivo: !x.nombre ? "sin nombre de inquilino" : "sin alquiler este mes" });
        continue;
      }
      const avisos = [];
      const sinRotulo = x.lineas.filter((l) => !l.concepto);
      if (sinRotulo.length) avisos.push(`importe sin concepto en fila ${sinRotulo.map((l) => l.fila).join(", ")}`);
      const sumaX = x.lineas.filter((l) => !esHonorario(l.concepto)).reduce((t, l) => t + l.importe, 0);
      const suma = x.lineas.reduce((t, l) => t + l.importe, 0);
      if (x.total != null && Math.abs(x.total - sumaX) >= 1 && Math.abs(x.total - suma) >= 1)
        avisos.push(`el TOTAL del Excel dice ${x.total}, la suma da ${suma}`);
      if (!x.nro_doc) avisos.push("sin DNI/CUIT");
      const vto = x.notas.map((n) => /^vto/i.test(n) && fecha(n)).find(Boolean);
      if (vto && (vto.a < anio || (vto.a === anio && vto.m < mes))) avisos.push(`contrato vencido (${vto.d}/${vto.m}/${vto.a})`);
      const ajuste = x.ajustes.map(fecha).some((f) => f && f.a === anio && f.m === mes);
      if (ajuste && x.alquilerAnterior != null && Math.abs(alq.importe - x.alquilerAnterior) < 1)
        avisos.push("este mes tocaba ajuste y el alquiler no cambió");
      const notaTxt = x.notas.join(" ");
      inquilinos.push({
        fila: x.fila,
        cliente: {
          tipo_doc: x.tipo_doc, nro_doc: x.nro_doc, nombre: x.nombre,
          condicion_iva: /factura\s*a\b/i.test(notaTxt) || /^3/.test(x.nro_doc) ? "IVA Responsable Inscripto" : "Consumidor Final",
          domicilio: "SANTA FE", condicion_venta: "Contado",
        },
        propietario: x.propietario,
        items: x.lineas.filter((l) => l.concepto).map((l) => ({
          concepto: /^alq/i.test(l.concepto) ? "Por cuenta y orden alq. {mes}" : l.concepto,
          importe: Math.round(l.importe * 100) / 100,
        })),
        alquiler: alq.importe, alquilerAnterior: x.alquilerAnterior,
        ajusteEsteMes: ajuste, avisos,
      });
    }
    return { inquilinos, omitidos };
  }

  const api = { leerListado, parseEncabezado, esHonorario };
  if (typeof module !== "undefined") module.exports = api; else raiz.Listado = api;
})(typeof window !== "undefined" ? window : globalThis);
