/* global XLSX, JSZip, Factura, Listado, EMISOR */
// Lee el "Listado mensual alquileres", arma una liquidación X por inquilino
// y separa los honorarios para hacerlos en ARCA.

const $ = (id) => document.getElementById(id);
const MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
  "septiembre", "octubre", "noviembre", "diciembre"];
let libro = null;   // el Excel leído
let datos = null;   // { inquilinos, omitidos }

const esHonorario = (c) => Listado.esHonorario(c);
const pesos = (n) => "$ " + n.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const guardar = (k, v) => { try { localStorage.setItem(k, v); } catch { /* sin almacenamiento */ } };
const leer = (k) => { try { return localStorage.getItem(k); } catch { return null; } };

function hoyISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
const fechaAR = (iso) => iso.split("-").reverse().join("/");
const periodo = () => { const [a, m] = $("mes").value.split("-").map(Number); return { a, m }; };
const etiquetaMes = () => { const { a, m } = periodo(); return `${MESES[m - 1]}/${a}`; };
const conMes = (c) => c.replace("{mes}", etiquetaMes());

function llenarMeses() {
  const d = new Date();
  const opciones = [];
  for (let k = -2; k <= 2; k++) {
    const x = new Date(d.getFullYear(), d.getMonth() + k, 1);
    const v = `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}`;
    opciones.push(`<option value="${v}"${k === 0 ? " selected" : ""}>${MESES_LARGO[x.getMonth()]} ${x.getFullYear()}</option>`);
  }
  $("mes").innerHTML = opciones.join("");
}

function procesar() {
  $("error").classList.add("oculto");
  if (!libro) return;
  try {
    const { a, m } = periodo();
    datos = Listado.leerListado(libro, a, m, XLSX);
    datos.inquilinos.forEach((i) => (i.emitir = true));
    $("nombreArchivo").textContent = `${libro._nombre} · ${datos.inquilinos.length} inquilinos`;
    mostrarRevision();
  } catch (e) {
    datos = null;
    ["paso2", "paso3", "paso4"].forEach((p) => $(p).classList.add("oculto"));
    mostrarError("No pude leer el archivo: " + e.message);
  }
}

function mostrarError(t) { $("error").textContent = t; $("error").classList.remove("oculto"); }

function mostrarRevision() {
  const conAvisos = datos.inquilinos.filter((i) => i.avisos.length).length;
  const filas = datos.inquilinos.map((i, n) => {
    const x = i.items.filter((it) => !esHonorario(it.concepto));
    const h = i.items.filter((it) => esHonorario(it.concepto));
    const total = x.reduce((s, it) => s + it.importe, 0);
    let ajuste = "";
    if (i.alquilerAnterior && Math.abs(i.alquiler - i.alquilerAnterior) >= 1) {
      const pct = ((i.alquiler / i.alquilerAnterior - 1) * 100).toFixed(1).replace(".", ",");
      ajuste = `<span class="tag azul">alquiler ${pesos(i.alquilerAnterior)} → ${pesos(i.alquiler)} (${pct}%)</span>`;
    }
    const avisos = i.avisos.map((a) => `<span class="tag rojo">⚠ ${esc(a)}</span>`).join(" ");
    return `<tr class="${i.emitir ? "" : "apagado"}">
      <td><input type="checkbox" data-i="${n}" ${i.emitir ? "checked" : ""} aria-label="Emitir"></td>
      <td class="inq"><strong>${esc(i.cliente.nombre)}</strong><br><span class="sub">${esc(i.cliente.tipo_doc)} ${esc(i.cliente.nro_doc) || "—"}</span>
        <div class="tags">${ajuste} ${avisos}${h.length ? ` <span class="tag">+ honorarios para ARCA</span>` : ""}</div></td>
      <td class="conceptos">${x.map((it) => `<div class="linea"><span>${esc(conMes(it.concepto))}</span><span>${pesos(it.importe)}</span></div>`).join("")}</td>
      <td class="n">${pesos(total)}</td></tr>`;
  });
  $("resumenRevision").innerHTML = conAvisos
    ? `<span class="tag rojo">⚠ ${conAvisos} con algo para revisar</span> Destildá los que no quieras emitir.`
    : "Destildá los que no quieras emitir.";
  $("tabla").innerHTML = `<tr><th></th><th>Inquilino</th><th>Conceptos</th><th class="n">Total</th></tr>${filas.join("")}`;
  $("omitidos").innerHTML = datos.omitidos.length
    ? `<details><summary>${datos.omitidos.length} renglones del Excel que no se emiten</summary><ul>${
      datos.omitidos.map((o) => `<li>Fila ${o.fila}: ${esc(o.encabezado)} <span class="sub">(${esc(o.motivo)})</span></li>`).join("")}</ul></details>`
    : "";
  ["paso2", "paso3"].forEach((p) => $(p).classList.remove("oculto"));
  $("paso4").classList.add("oculto");
  $("listo").classList.add("oculto");
}

async function cargar(file) {
  $("error").classList.add("oculto");
  try {
    libro = XLSX.read(await file.arrayBuffer());
    libro._nombre = file.name;
    procesar();
  } catch (e) {
    libro = null;
    mostrarError("No pude abrir el archivo: " + e.message);
  }
}

async function generar() {
  const btn = $("generar");
  btn.disabled = true; btn.textContent = "Generando…";
  try {
    const fecha = fechaAR($("fecha").value);
    const etiqueta = $("mes").value;
    let numero = parseInt($("numero").value, 10) || 1;
    const copias = String(EMISOR.copias).split(",").map((s) => s.trim()).filter(Boolean);
    const pv = String(EMISOR.punto_venta).padStart(5, "0");
    const zip = new JSZip();
    const carpeta = zip.folder(`Liquidaciones ${etiqueta}`);
    const enArca = [], resumen = [["Número", "Inquilino", "Documento", "Propietario", "Total"]];
    let cantidad = 0, suma = 0;

    for (const inq of datos.inquilinos.filter((i) => i.emitir)) {
      const items = inq.items.map((it) => ({ ...it, concepto: conMes(it.concepto) }));
      items.filter((it) => esHonorario(it.concepto)).forEach((it) => enArca.push({ cliente: inq.cliente, ...it }));
      const x = items.filter((it) => !esHonorario(it.concepto));
      if (!x.length) continue;
      const { blob, total } = Factura.crearComprobante(EMISOR, inq.cliente, x, numero, fecha, copias);
      const nombre = inq.cliente.nombre.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^\w -]/g, "").trim().replace(/\s+/g, "_");
      carpeta.file(`X_${pv}_${String(numero).padStart(8, "0")}_${nombre}.pdf`, blob);
      resumen.push([numero, inq.cliente.nombre, `${inq.cliente.tipo_doc} ${inq.cliente.nro_doc}`, inq.propietario, total]);
      numero++; cantidad++; suma += total;
    }

    const lib = XLSX.utils.book_new();
    const wsR = XLSX.utils.aoa_to_sheet(resumen);
    wsR["!cols"] = [8, 32, 16, 22, 14].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(lib, wsR, "Liquidaciones");
    if (enArca.length) {
      const ws = XLSX.utils.aoa_to_sheet([
        ["Hecho", "Inquilino", "Tipo doc", "Nro doc", "Condición IVA", "Concepto", "Importe (con IVA)", "Neto sin IVA (para factura A)"],
        ...enArca.map((r) => ["", r.cliente.nombre, r.cliente.tipo_doc, r.cliente.nro_doc, r.cliente.condicion_iva,
          r.concepto, Factura.fmt(r.importe), Factura.fmt(r.importe / 1.21)]),
      ]);
      ws["!cols"] = [8, 32, 9, 14, 20, 26, 16, 18].map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(lib, ws, "Hacer en ARCA");
    }
    carpeta.file(`RESUMEN_${etiqueta}.xlsx`, XLSX.write(lib, { type: "array", bookType: "xlsx" }));

    const contenido = await zip.generateAsync({ type: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(contenido);
    a.download = `Liquidaciones ${etiqueta}.zip`;
    document.body.appendChild(a); a.click(); a.remove();

    guardar("proximo_numero", String(numero));
    $("listo").innerHTML = `Listo: ${cantidad} liquidaciones por ${pesos(suma)}. Se descargó <strong>Liquidaciones ${etiqueta}.zip</strong>.<br>
      El próximo mes empezá desde el número <strong>${numero}</strong>.`;
    $("listo").classList.remove("oculto");
    mostrarArca(enArca);
  } catch (e) {
    mostrarError("Algo salió mal al generar: " + e.message);
  } finally {
    btn.disabled = false; btn.textContent = "Generar liquidaciones";
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
document.addEventListener("change", (e) => {
  const i = e.target.dataset?.i;
  if (i === undefined || !datos) return;
  datos.inquilinos[+i].emitir = e.target.checked;
  e.target.closest("tr").classList.toggle("apagado", !e.target.checked);
});

llenarMeses();
$("fecha").value = hoyISO();
$("numero").value = leer("proximo_numero") || 1;
$("mes").addEventListener("change", procesar);
$("archivo").addEventListener("change", (e) => e.target.files[0] && cargar(e.target.files[0]));
$("generar").addEventListener("click", generar);
const zona = $("zona");
zona.addEventListener("dragover", (e) => { e.preventDefault(); zona.style.borderColor = "var(--accent)"; });
zona.addEventListener("dragleave", () => (zona.style.borderColor = ""));
zona.addEventListener("drop", (e) => { e.preventDefault(); zona.style.borderColor = ""; e.dataTransfer.files[0] && cargar(e.dataTransfer.files[0]); });
