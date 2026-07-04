/*************************************************************
 *  LoborojoPy - API para el Panel de Clientes
 *  ----------------------------------------------------------
 *  Pegá TODO este código en el editor de Apps Script de tu
 *  planilla:  Extensiones -> Apps Script  (borrá lo que haya)
 *
 *  Columnas de la hoja (detectadas por su título, en cualquier orden):
 *    Usuario | Contraseña | Vence | Pago (si/no) | Debe (si/no) | Activado (si/no)
 *
 *  Luego: Implementar -> Nueva implementación -> "Aplicación web"
 *    - Ejecutar como:  Yo
 *    - Quién tiene acceso:  Cualquier persona
 *  Copiá la URL que termina en /exec y pegala en el panel.
 *************************************************************/

// Si tu pestaña NO se llama "Clientes", cambiá el nombre acá.
// Dejalo en "" para usar la primera hoja del archivo.
var NOMBRE_HOJA = "";

/* =========================================================
 *  AVISOS AUTOMÁTICOS DE VENCIMIENTO
 *  Completá SOLO el/los medio(s) que quieras usar y dejá
 *  los demás vacíos ("").  Más abajo se explica cómo activar
 *  el aviso diario automático.
 * ========================================================= */

// 1) EMAIL (lo más simple, no requiere nada extra)
var EMAIL_AVISOS = "";          // ej: "loborojopy@gmail.com"

// 2) TELEGRAM (gratis) - creá un bot con @BotFather
var TELEGRAM_TOKEN   = "";      // token que te da BotFather
var TELEGRAM_CHAT_ID = "";      // tu chat id (ver instrucciones)

// 3) WHATSAPP (gratis para uso personal vía CallMeBot)
var WHATSAPP_PHONE  = "";       // tu numero con codigo pais, ej: "595981234567"
var WHATSAPP_APIKEY = "";       // apikey que te da CallMeBot

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
  var ancho = Math.max(h.getLastColumn(), 7);
  var tit = h.getRange(1, 1, 1, ancho).getValues()[0]
             .map(function (x) { return String(x).trim().toLowerCase(); });
  function buscar(claves, pordefecto) {
    for (var i = 0; i < tit.length; i++)
      for (var j = 0; j < claves.length; j++)
        if (tit[i] && tit[i].indexOf(claves[j]) > -1) return i + 1;
    return pordefecto; // posición por defecto si no se encuentra el título
  }
  return {
    nombre:   buscar(["usuario", "cliente", "nombre"], 1),
    pass:     buscar(["contra", "clave", "pass"], 2),
    vence:    buscar(["vence", "venc", "vencim"], 3),
    pago:     buscar(["pago", "pagó", "abon"], 5),
    debe:     buscar(["debe", "deuda", "adeuda"], 6),
    activado: buscar(["activ", "servidor", "server", "estado"], 7)
  };
}

// Interpreta un valor de tipo si/no (acepta si, sí, s, x, true, 1, ok).
function esSi(v) {
  var t = String(v).trim().toLowerCase();
  return t === "si" || t === "sí" || t === "s" || t === "x" ||
         t === "true" || t === "1" || t === "ok" || t === "activo" || t === "pagó";
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
    var pagoSi = esSi(f[col.pago - 1]);        // columna "Pago"  (si/no)
    var debeSi = esSi(f[col.debe - 1]);        // columna "Debe"  (si/no)
    var activoSi = esSi(f[col.activado - 1]);  // columna "Activado" (si/no)
    out.push({
      fila: i + 2,
      nombre: String(nombre),
      // La contraseña NO se envía al panel, por seguridad.
      vence: fmtFecha(f[col.vence - 1]),
      // Si "Debe" está en si -> debe. Si no, y "Pago" está en si -> pagó. Si no -> debe.
      pago: debeSi ? "debe" : (pagoSi ? "pagó" : "debe"),
      servidor: activoSi ? "activo" : "inactivo"
    });
  }
  return out;
}

/* ============ Escritura ============ */
function agregarCliente(d) {
  var h = hoja(), col = columnas(h);
  var ancho = Math.max(h.getLastColumn(), 7);
  var fila = new Array(ancho).fill("");
  var debe = (d.pago === "debe");
  fila[col.nombre - 1]   = d.nombre || "";
  fila[col.pass - 1]     = d.pass || "";
  fila[col.vence - 1]    = fmtFecha(d.vence);
  fila[col.pago - 1]     = debe ? "no" : "si";                     // Pago
  fila[col.debe - 1]     = debe ? "si" : "no";                     // Debe
  fila[col.activado - 1] = (d.servidor === "inactivo") ? "no" : "si"; // Activado
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
  if (d.marcarPago) {
    h.getRange(fila, col.pago).setValue("si");  // marca Pago = si
    h.getRange(fila, col.debe).setValue("no");  // y Debe = no
  }
  if (d.activar) h.getRange(fila, col.activado).setValue("si"); // Activado = si
  return { ok: true, clientes: leerClientes() };
}

function cambiarServidor(d) {
  var h = hoja(), col = columnas(h);
  var fila = +d.fila;
  if (!fila || fila < 2) return { ok: false, error: "Fila invalida" };
  h.getRange(fila, col.activado).setValue(d.estado === "activo" ? "si" : "no");
  return { ok: true, clientes: leerClientes() };
}

function eliminarCliente(d) {
  var h = hoja();
  var fila = +d.fila;
  if (!fila || fila < 2) return { ok: false, error: "Fila invalida" };
  h.deleteRow(fila);
  return { ok: true, clientes: leerClientes() };
}


/* =========================================================
 *  Revisa vencimientos y envía el aviso.
 *  --> Esta es la función que hay que programar para que
 *      corra sola todos los días (ver instrucciones abajo).
 * ========================================================= */
function revisarVencimientos() {
  var clientes = leerClientes();
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var venceHoy = [], venceManana = [], vencidos = [];

  clientes.forEach(function (c) {
    var v = parseFecha(c.vence);
    if (!v) return;
    v.setHours(0, 0, 0, 0);
    var dias = Math.round((v - hoy) / 86400000);
    var etiqueta = c.nombre.trim() + " (" + c.vence + ")";
    if (dias === 0) venceHoy.push(etiqueta);
    else if (dias === 1) venceManana.push(etiqueta);
    else if (dias < 0) vencidos.push(etiqueta + " - hace " + Math.abs(dias) + " día(s)");
  });

  if (!venceHoy.length && !venceManana.length && !vencidos.length) return; // nada que avisar

  var lineas = ["🐺 LoborojoPy - Avisos de vencimiento", ""];
  if (venceManana.length) lineas.push("⏰ Vencen MAÑANA:\n- " + venceManana.join("\n- "), "");
  if (venceHoy.length)    lineas.push("🔴 Vencen HOY:\n- " + venceHoy.join("\n- "), "");
  if (vencidos.length)    lineas.push("❌ Ya VENCIDOS:\n- " + vencidos.join("\n- "), "");
  var mensaje = lineas.join("\n");

  if (EMAIL_AVISOS)                          enviarEmail(mensaje);
  if (TELEGRAM_TOKEN && TELEGRAM_CHAT_ID)    enviarTelegram(mensaje);
  if (WHATSAPP_PHONE && WHATSAPP_APIKEY)     enviarWhatsApp(mensaje);
}

function enviarEmail(texto) {
  MailApp.sendEmail(EMAIL_AVISOS, "🐺 LoborojoPy - Avisos de vencimiento", texto);
}

function enviarTelegram(texto) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  UrlFetchApp.fetch(url, {
    method: "post",
    payload: { chat_id: TELEGRAM_CHAT_ID, text: texto },
    muteHttpExceptions: true
  });
}

function enviarWhatsApp(texto) {
  var url = "https://api.callmebot.com/whatsapp.php?phone=" + WHATSAPP_PHONE +
            "&text=" + encodeURIComponent(texto) + "&apikey=" + WHATSAPP_APIKEY;
  UrlFetchApp.fetch(url, { muteHttpExceptions: true });
}

/* Función de prueba: ejecutala a mano una vez para probar el envío. */
function probarAviso() {
  var msg = "🐺 LoborojoPy - Prueba de aviso.\nSi ves este mensaje, ¡los avisos funcionan!";
  if (EMAIL_AVISOS)                       enviarEmail(msg);
  if (TELEGRAM_TOKEN && TELEGRAM_CHAT_ID) enviarTelegram(msg);
  if (WHATSAPP_PHONE && WHATSAPP_APIKEY)  enviarWhatsApp(msg);
}
