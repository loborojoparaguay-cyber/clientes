/*************************************************************
 *  LoborojoPy - API para el Panel de Clientes
 *  ----------------------------------------------------------
 *  Pegá TODO este código en el editor de Apps Script de tu
 *  planilla:  Extensiones -> Apps Script  (borrá lo que haya)
 *
 *  Columnas esperadas en la hoja (en cualquier orden, la
 *  primera fila deben ser los títulos):
 *    Cliente | Contraseña | Vencimiento | Pago | Servidor
 *
 *  Luego: Implementar -> Nueva implementación -> "Aplicación web"
 *    - Ejecutar como:  Yo
 *    - Quién tiene acceso:  Cualquier persona
 *  Copiá la URL que termina en /exec y pegala en el panel.
 *************************************************************/

// Si tu pestaña NO se llama "Clientes", cambiá el nombre acá.
// Dejalo en "" para usar la primera hoja del archivo.
var NOMBRE_HOJA = "";

/* ============ Puntos de entrada ============ */
function doGet(e) {
  return responder({ ok: true, clientes: leerClientes() });
}

function doPost(e) {
  var datos = {};
  try { datos = JSON.parse(e.postData.contents); } catch (err) {}
  var r;
  switch (datos.accion) {
    case "agregar":  r = agregarCliente(datos); break;
    case "renovar":  r = renovarCliente(datos); break;
    case "servidor": r = cambiarServidor(datos); break;
    case "eliminar": r = eliminarCliente(datos); break;
    default:         r = { ok: false, error: "Accion desconocida" };
  }
  return responder(r);
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ Utilidades de hoja ============ */
function hoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return (NOMBRE_HOJA ? ss.getSheetByName(NOMBRE_HOJA) : null) || ss.getSheets()[0];
}

// Detecta en qué columna (1-based) está cada dato, por el título.
function columnas(h) {
  var ancho = Math.max(h.getLastColumn(), 5);
  var tit = h.getRange(1, 1, 1, ancho).getValues()[0]
             .map(function (x) { return String(x).trim().toLowerCase(); });
  function buscar(claves, pormellado) {
    for (var i = 0; i < tit.length; i++)
      for (var j = 0; j < claves.length; j++)
        if (tit[i].indexOf(claves[j]) > -1) return i + 1;
    return pormellado; // posición por defecto si no hay título
  }
  return {
    nombre:   buscar(["cliente", "nombre", "usuario"], 1),
    pass:     buscar(["contra", "clave", "pass"], 2),
    vence:    buscar(["venc", "vence", "fecha", "expira"], 3),
    pago:     buscar(["pago", "abon"], 4),
    servidor: buscar(["servidor", "server", "estado"], 5)
  };
}

function fmtFecha(d) {
  if (!(d instanceof Date)) {
    var p = parseFecha(d);
    if (!p) return String(d || "");
    d = p;
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseFecha(txt) {
  if (txt instanceof Date) return txt;
  if (!txt) return null;
  txt = String(txt).trim();
  var m;
  if ((m = txt.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/)))
    return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = txt.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/))) {
    var a = +m[3]; if (a < 100) a += 2000;
    return new Date(a, +m[2] - 1, +m[1]);
  }
  var d = new Date(txt);
  return isNaN(d) ? null : d;
}

/* ============ Lectura ============ */
function leerClientes() {
  var h = hoja(), col = columnas(h);
  var ultima = h.getLastRow();
  if (ultima < 2) return [];
  var ancho = h.getLastColumn();
  var filas = h.getRange(2, 1, ultima - 1, ancho).getValues();
  var out = [];
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    var nombre = f[col.nombre - 1];
    if (String(nombre).trim() === "") continue;
    out.push({
      fila: i + 2,
      nombre: String(nombre),
      // La contraseña NO se envía al panel, por seguridad.
      vence: fmtFecha(f[col.vence - 1]),
      pago: /pag|si|ok/i.test(String(f[col.pago - 1])) ? "pagó" : "debe",
      servidor: /inact|no|off/i.test(String(f[col.servidor - 1])) ? "inactivo" : "activo"
    });
  }
  return out;
}

/* ============ Escritura ============ */
function agregarCliente(d) {
  var h = hoja(), col = columnas(h);
  var ancho = Math.max(h.getLastColumn(), 5);
  var fila = new Array(ancho).fill("");
  fila[col.nombre - 1]   = d.nombre || "";
  fila[col.pass - 1]     = d.pass || "";
  fila[col.vence - 1]    = fmtFecha(d.vence);
  fila[col.pago - 1]     = d.pago || "debe";
  fila[col.servidor - 1] = d.servidor || "activo";
  h.appendRow(fila);
  return { ok: true, clientes: leerClientes() };
}

// Renueva: suma dias al vencimiento (desde hoy si ya venció),
// y opcionalmente marca "pagó" y activa el servidor.
function renovarCliente(d) {
  var h = hoja(), col = columnas(h);
  var fila = +d.fila;
  if (!fila || fila < 2) return { ok: false, error: "Fila invalida" };

  var actual = parseFecha(h.getRange(fila, col.vence).getValue());
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var base = (!actual || actual < hoy) ? hoy : actual;
  var dias = parseInt(d.dias, 10) || 0;
  base.setDate(base.getDate() + dias);

  h.getRange(fila, col.vence).setValue(fmtFecha(base));
  if (d.marcarPago) h.getRange(fila, col.pago).setValue("pagó");
  if (d.activar)    h.getRange(fila, col.servidor).setValue("activo");
  return { ok: true, clientes: leerClientes() };
}

function cambiarServidor(d) {
  var h = hoja(), col = columnas(h);
  var fila = +d.fila;
  if (!fila || fila < 2) return { ok: false, error: "Fila invalida" };
  h.getRange(fila, col.servidor).setValue(d.estado === "activo" ? "activo" : "inactivo");
  return { ok: true, clientes: leerClientes() };
}

function eliminarCliente(d) {
  var h = hoja();
  var fila = +d.fila;
  if (!fila || fila < 2) return { ok: false, error: "Fila invalida" };
  h.deleteRow(fila);
  return { ok: true, clientes: leerClientes() };
}
