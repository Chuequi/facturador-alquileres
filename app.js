/* global XLSX, JSZip, Factura */
// Lee el Excel, arma los comprobantes por inquilino y separa los honorarios.
// NOTA: por ahora entiende el formato de prueba (hojas Emisor / Contratos / Conceptos).
// Cuando llegue el Excel real del "Listado mensual alquileres", se cambia solo leerExcel().

const $ = (id) => document.getElementById(id);
const MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
let datos = null; // { emisor, inquilinos: [{cliente, items:[{concepto, importe}]}] }

const esHonorario = (c) => String(c).trim().toLowerCase().startsWith("honorario");
const num = (v) => typeof v === "number" ? v : parseFloat(String(v ?? "0").replace(/[$\s.]/g, "").replace(",", ".")) || 0;
const pesos = (n) => "$ " + n.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function hoy() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function fechaAR(iso) { const [a, m, d] = iso.split("-"); return `${d}/${m}/${a}`; }

function leerExcel(wb) {
  const tabla = (n) => {
    if (!wb.Sheets[n]) throw new Error(`El Excel no tiene la hoja "${n}".`);
    return XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "", raw: true });
  };
  const emisor = {};
  for (const r of tabla("Emisor")) if (r.campo) emisor[String(r.campo).trim()] = String(r.valor);
  const contratos = tabla("Contratos").filter((c) => !/^no$/i.test(String(c.activo).trim()));
  const conceptos = tabla("Conceptos");
  const inquilinos = contratos.map((c) => ({
    cliente: {
      tipo_doc: c.tipo_doc || "DNI", nro_doc: String(c.nro_doc), nombre: c.inquilino,
      condicion_iva: c.condicion_iva || "Consumidor Final", domicilio: c.domicilio || "",
      condicion_venta: c.condicion_venta || "Contado",
    },
    items: conceptos.filter((k) => String(k.contrato_id) === String(c.id) && num(k.importe_actual) !== 0)
      .map((k) => ({ concepto: String(k.concepto), importe: num(k.importe_actual) })),
  })).filter((i) => i.items.length);
  return { emisor, inquilinos };
}

function conMes(concepto) {
  const [a, m] = $("fecha").value.split("-");
  return concepto.replace("{mes}", `${MESES[+m - 1]}/${a}`);
}

function mostrarRevision() {
  const filas = datos.inquilinos.map((i) => {
    const x = i.items.filter((it) => !esHonorario(it.concepto));
    const h = i.items.filter((it) => esHonorario(it.concepto));
    const total = x.reduce((s, it) => s + it.importe, 0);
    return `<tr><td><strong>${esc(i.cliente.nombre)}</strong><br><span class="sub">${esc(i.cliente.tipo_doc)} ${esc(i.cliente.nro_doc)}</span></td>
      <td>${x.map((it) => esc(conMes(it.concepto))).join("<br>")}${h.length ? `<br><span class="tag">+ ${h.length} honorario/s para ARCA</span>` : ""}</td>
      <td class="n">${x.length ? pesos(total) : "—"}</td></tr>`;
  });
  $("tabla").innerHTML = `<tr><th>Inquilino</th><th>Conceptos</th><th class="n">Total X</th></tr>${filas.join("")}`;
  ["paso2", "paso3"].forEach((p) => $(p).classList.remove("oculto"));
  $("paso4").classList.add("oculto");
  $("listo").classList.add("oculto");
}

async function cargar(file) {
  $("error").classList.add("oculto");
  try {
    const wb = XLSX.read(await file.arrayBuffer());
    datos = leerExcel(wb);
    if (!datos.inquilinos.length) throw new Error("No encontré inquilinos con importes en el Excel.");
    if (datos.emisor.proximo_numero) $("numero").value = parseInt(datos.emisor.proximo_numero, 10) || 1;
    $("nombreArchivo").textContent = `${file.name} · ${datos.inquilinos.length} inquilinos`;
    mostrarRevision();
  } catch (e) {
    datos = null;
    $("error").textContent = "No pude leer el archivo: " + e.message;
    $("error").classList.remove("oculto");
  }
}

async function generar() {
  const btn = $("generar");
  btn.disabled = true; btn.textContent = "Generando…";
  try {
    const fecha = fechaAR($("fecha").value);
    const etiqueta = $("fecha").value.slice(0, 7);
    let numero = parseInt($("numero").value, 10) || 1;
    const copias = String(datos.emisor.copias || "ORIGINAL, DUPLICADO").split(",").map((s) => s.trim()).filter(Boolean);
    const pv = String(datos.emisor.punto_venta).padStart(5, "0");
    const zip = new JSZip();
    const carpeta = zip.folder(`Facturas ${etiqueta}`);
    const enArca = [];
    let cantidad = 0, suma = 0;

    for (const inq of datos.inquilinos) {
      const items = inq.items.map((it) => ({ ...it, concepto: conMes(it.concepto) }));
      items.filter((it) => esHonorario(it.concepto)).forEach((it) => enArca.push({ cliente: inq.cliente, ...it }));
      const x = items.filter((it) => !esHonorario(it.concepto));
      if (!x.length) continue;
      const { blob, total } = Factura.crearComprobante(datos.emisor, inq.cliente, x, numero, fecha, copias);
      const nombre = inq.cliente.nombre.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^\w -]/g, "").trim().replace(/\s+/g, "_");
      carpeta.file(`X_${pv}_${String(numero).padStart(8, "0")}_${nombre}.pdf`, blob);
      numero++; cantidad++; suma += total;
    }

    if (enArca.length) {
      const ws = XLSX.utils.aoa_to_sheet([
        ["Hecho", "Inquilino", "Tipo doc", "Nro doc", "Condición IVA", "Concepto", "Importe (con IVA)", "Neto sin IVA (para factura A)"],
        ...enArca.map((r) => ["", r.cliente.nombre, r.cliente.tipo_doc, r.cliente.nro_doc, r.cliente.condicion_iva,
          r.concepto, Factura.fmt(r.importe), Factura.fmt(r.importe / 1.21)]),
      ]);
      ws["!cols"] = [8, 32, 9, 14, 20, 26, 16, 18].map((w) => ({ wch: w }));
      const libro = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(libro, ws, "Hacer en ARCA");
      carpeta.file(`HACER_EN_ARCA_${etiqueta}.xlsx`, XLSX.write(libro, { type: "array", bookType: "xlsx" }));
    }

    const contenido = await zip.generateAsync({ type: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(contenido);
    a.download = `Facturas ${etiqueta}.zip`;
    document.body.appendChild(a); a.click(); a.remove();

    $("listo").innerHTML = `Listo: ${cantidad} facturas X por ${pesos(suma)}. Se descargó <strong>Facturas ${etiqueta}.zip</strong>.<br>
      El próximo mes empezá desde el número <strong>${numero}</strong>.`;
    $("listo").classList.remove("oculto");
    mostrarArca(enArca);
  } catch (e) {
    $("error").textContent = "Algo salió mal al generar: " + e.message;
    $("error").classList.remove("oculto");
  } finally {
    btn.disabled = false; btn.textContent = "Generar facturas";
  }
}

function mostrarArca(lista) {
  if (!lista.length) { $("paso4").classList.add("oculto"); return; }
  const boton = (v) => `<button class="btn chico" data-copiar="${esc(v)}">Copiar</button>`;
  $("arca").innerHTML = `<tr><th>Inquilino</th><th>Documento</th><th>Concepto</th><th class="n">Importe</th></tr>` +
    lista.map((r) => `<tr><td>${esc(r.cliente.nombre)}</td>
      <td>${esc(r.cliente.tipo_doc)} ${esc(r.cliente.nro_doc)} ${boton(r.cliente.nro_doc)}</td>
      <td>${esc(r.concepto)} ${boton(r.concepto)}</td>
      <td class="n">${Factura.fmt(r.importe)} ${boton(Factura.fmt(r.importe))}</td></tr>`).join("");
  $("paso4").classList.remove("oculto");
}

document.addEventListener("click", async (e) => {
  const v = e.target.dataset?.copiar;
  if (v === undefined) return;
  try { await navigator.clipboard.writeText(v); } catch { /* sin permiso de portapapeles */ }
  e.target.textContent = "¡Copiado!";
  setTimeout(() => (e.target.textContent = "Copiar"), 1500);
});

$("fecha").value = hoy();
$("fecha").addEventListener("change", () => datos && mostrarRevision());
$("archivo").addEventListener("change", (e) => e.target.files[0] && cargar(e.target.files[0]));
$("generar").addEventListener("click", generar);
const zona = $("zona");
zona.addEventListener("dragover", (e) => { e.preventDefault(); zona.style.borderColor = "var(--accent)"; });
zona.addEventListener("dragleave", () => (zona.style.borderColor = ""));
zona.addEventListener("drop", (e) => { e.preventDefault(); zona.style.borderColor = ""; e.dataTransfer.files[0] && cargar(e.dataTransfer.files[0]); });
