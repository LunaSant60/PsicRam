// Apps Script del formulario de datos. El mismo script muestra el formulario
// (archivo «Formulario» del proyecto) y guarda las respuestas.
//
// La hoja de cálculo usa dos pestañas, que se crean solas si no existen:
//   «Accesos»:    Nombre | Código | Usado | Enlace
//   «Respuestas»: se llena sola; la fila de encabezados se crea si está vacía.
//
// Para invitar a alguien: escribe su nombre en la columna A de «Accesos» y usa
// el menú Accesos › Generar enlaces. Cada enlace sirve para un solo envío.

var HOJA_ACCESOS = "Accesos";
var HOJA_RESPUESTAS = "Respuestas";

// URL de la implementación (Implementar › Administrar implementaciones).
var FORM_URL = "https://script.google.com/macros/s/AKfycbzyTBjmP5KxniCPpAdkrliDGbgBQDIt23h0k6O07PBaGzCEuSMW1phjN1Arc1nyw8bOMw/exec";

var CAMPOS = [
  { nombre: "nombre",    titulo: "Nombre completo",           patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]+$/ },
  { nombre: "calle",     titulo: "Calle",                     patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]+$/ },
  { nombre: "numero",    titulo: "Número",                    patron: /^[0-9]+$/ },
  { nombre: "interior",  titulo: "Número interior",           patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]*$/, opcional: true },
  { nombre: "colonia",   titulo: "Colonia o Fraccionamiento", patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 .,#\-]+$/ },
  { nombre: "postal",    titulo: "Código postal",             patron: /^[0-9]{1,5}$/ },
  { nombre: "municipio", titulo: "Municipio",                 patron: /^.{1,40}$/ },
  { nombre: "clave",     titulo: "Clave de elector",          patron: /^[A-Za-z0-9]{18}$/ },
  { nombre: "seccion",   titulo: "Sección",                   patron: /^[0-9]{1,4}$/ },
  { nombre: "ocr",       titulo: "OCR",                       patron: /^[0-9]{13}$/ },
  { nombre: "telefono",  titulo: "Teléfono",                  patron: /^[0-9]{10}$/ },
  { nombre: "correo",    titulo: "Correo electrónico",        patron: /^[^\s@]+@[^\s@]+\.[^\s@]+$/ }
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Accesos")
    .addItem("Generar enlaces", "generarEnlaces")
    .addToUi();
}

// Llena Código y Enlace en cada fila de «Accesos» que tenga nombre y no tenga código.
function generarEnlaces() {
  var hoja = obtenerHojaAccesos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return;
  var filas = hoja.getRange(2, 1, ultima - 1, 4).getValues();
  filas.forEach(function(fila) {
    if (fila[0] && !fila[1]) {
      fila[1] = Utilities.getUuid().replace(/-/g, "");
    }
    if (fila[1]) {
      fila[3] = FORM_URL + "?acceso=" + fila[1];
    }
  });
  hoja.getRange(2, 1, filas.length, 4).setValues(filas);
}

// Muestra el formulario solo si el enlace trae un código válido y sin usar.
function doGet(e) {
  // «c» es un parámetro reservado de Apps Script: Google rechaza la URL antes
  // de llegar a doGet. Por eso el código va en «acceso».
  if (e.parameter.modo === "estres") {
    return HtmlService.createHtmlOutputFromFile("Estres").setTitle("Prueba de estrés");
  }
  var codigo = String(e.parameter.acceso || "").trim();
  var acceso = buscarAcceso_(codigo);
  var plantilla = HtmlService.createTemplateFromFile("Formulario");
  plantilla.estado = !codigo ? "falta" : !acceso ? "invalido" : acceso.usado ? "usado" : "ok";
  plantilla.codigo = plantilla.estado === "ok" ? codigo : "";
  return plantilla.evaluate()
    .setTitle("Formulario de datos")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// El formulario la llama con google.script.run al presionar «Enviar formulario».
function enviarDatos(datos) {
  datos = datos || {};
  var valores = [];
  for (var i = 0; i < CAMPOS.length; i++) {
    var campo = CAMPOS[i];
    var valor = String(datos[campo.nombre] || "").trim();
    if ((!valor && !campo.opcional) || (valor && !campo.patron.test(valor))) {
      return { ok: false, error: "datos", message: "Revisa el campo «" + campo.titulo + "»." };
    }
    // El apóstrofo evita que Sheets quite ceros a la izquierda (OCR, CP, etc.).
    valores.push("'" + valor);
  }

  // Solo la consulta y la escritura van dentro del candado, para que muchos
  // envíos a la vez esperen su turno poco tiempo y nunca se use dos veces un enlace.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: "ocupado" };
  try {
    var acceso = buscarAcceso_(datos.codigo);
    if (!acceso) return { ok: false, error: "invalido" };
    if (acceso.usado) return { ok: false, error: "usado" };

    obtenerHojaRespuestas_().appendRow([new Date(), acceso.nombre].concat(valores));
    obtenerHojaAccesos_().getRange(acceso.fila, 3).setValue(new Date());
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function buscarAcceso_(codigo) {
  codigo = String(codigo || "").trim();
  if (!codigo) return null;
  var hoja = obtenerHojaAccesos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return null;
  var filas = hoja.getRange(2, 1, ultima - 1, 3).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][1]) === codigo) {
      return { fila: i + 2, nombre: filas[i][0], usado: !!filas[i][2] };
    }
  }
  return null;
}

function obtenerHojaAccesos_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(HOJA_ACCESOS);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_ACCESOS);
    hoja.appendRow(["Nombre", "Código", "Usado", "Enlace"]);
  }
  return hoja;
}

function obtenerHojaRespuestas_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(HOJA_RESPUESTAS);
  if (!hoja) hoja = libro.insertSheet(HOJA_RESPUESTAS);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha", "Invitado"].concat(CAMPOS.map(function(c) { return c.titulo; })));
  }
  return hoja;
}

// ---------- Prueba de estrés ----------
// 1. Ejecuta crearAccesosDePrueba desde el editor (crea 200 enlaces de prueba).
// 2. Abre la URL de la implementación con «?modo=estres» y presiona «Iniciar».
// 3. Ejecuta borrarPruebas para quitar los renglones de prueba de las dos pestañas.
var PRUEBA_NOMBRE = "Prueba estrés";
var PRUEBA_TOTAL = 200;

function crearAccesosDePrueba() {
  borrarPruebas();
  var filas = [];
  for (var i = 1; i <= PRUEBA_TOTAL; i++) {
    var n = ("00" + i).slice(-3);
    filas.push([PRUEBA_NOMBRE + " " + n, "PRUEBA" + n, "", ""]);
  }
  var hoja = obtenerHojaAccesos_();
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 4).setValues(filas);
}

function borrarPruebas() {
  [obtenerHojaAccesos_(), obtenerHojaRespuestas_()].forEach(function(hoja) {
    var ultima = hoja.getLastRow();
    if (ultima < 2) return;
    var col = hoja.getName() === HOJA_ACCESOS ? 1 : 2;
    var nombres = hoja.getRange(2, col, ultima - 1, 1).getValues();
    // De abajo hacia arriba, borrando cada bloque seguido de renglones de prueba.
    var fin = -1;
    for (var i = nombres.length - 1; i >= -1; i--) {
      var esPrueba = i >= 0 && String(nombres[i][0]).indexOf(PRUEBA_NOMBRE) === 0;
      if (esPrueba && fin < 0) fin = i;
      if (!esPrueba && fin >= 0) {
        hoja.deleteRows(i + 3, fin - i);
        fin = -1;
      }
    }
  });
}
