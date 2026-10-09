// Apps Script del formulario de datos. El mismo script muestra el formulario
// (archivo «Formulario» del proyecto) y guarda las respuestas.
//
// La hoja de cálculo usa dos pestañas, que se crean solas si no existen:
//   «Accesos»:    Nombre | Código | Usado | Enlace | Ilimitado
//   «Respuestas»: se llena sola; la fila de encabezados se crea si está vacía.
// Las fotos se guardan en la carpeta privada «Formulario - Fotos» de tu Drive y
// en «Respuestas» queda el enlace a cada una.
//
// Para invitar a alguien: escribe su nombre en la columna A de «Accesos» y usa
// el menú Accesos › Generar enlaces. Cada enlace sirve para un solo envío, salvo
// que su renglón diga «Sí» en la columna Ilimitado: ese enlace sirve para siempre
// (en Usado queda la fecha del último envío) hasta que borres el renglón o el «Sí».

var HOJA_ACCESOS = "Accesos";
var HOJA_RESPUESTAS = "Respuestas";

// URL de la implementación (Implementar › Administrar implementaciones).
var FORM_URL = "https://script.google.com/macros/s/AKfycbzyTBjmP5KxniCPpAdkrliDGbgBQDIt23h0k6O07PBaGzCEuSMW1phjN1Arc1nyw8bOMw/exec";

var CAMPOS = [
  { nombre: "nombres",   titulo: "Nombre(s)",                 patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]+$/ },
  // Los dos apellidos son opcionales (hay quien solo tiene uno).
  { nombre: "paterno",   titulo: "Apellido paterno",          patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]*$/, opcional: true },
  { nombre: "materno",   titulo: "Apellido materno",          patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]*$/, opcional: true },
  { nombre: "calle",     titulo: "Calle",                     patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 .,#\-]+$/ },
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

// Fotos que pide el formulario; cada una llega ya reducida a JPEG desde el navegador.
var FOTOS = [
  { nombre: "foto1", titulo: "Foto 1" },
  { nombre: "foto2", titulo: "Foto 2" }
];
var CARPETA_FOTOS = "Formulario - Fotos";
var MAX_FOTO_BYTES = 5 * 1024 * 1024;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Accesos")
    .addItem("Generar enlaces", "generarEnlaces")
    .addItem("Acomodar columnas de Respuestas", "acomodarRespuestas")
    .addItem("Ver carpeta de fotos", "verCarpetaFotos")
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
  plantilla.ilimitado = plantilla.estado === "ok" && acceso.ilimitado;
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
    valores.push(valor ? "'" + valor : "");
  }

  var fotos = [];
  for (var j = 0; j < FOTOS.length; j++) {
    var foto = leerFoto_(datos[FOTOS[j].nombre]);
    if (!foto) return { ok: false, error: "datos", message: "Revisa la foto «" + FOTOS[j].titulo + "»." };
    fotos.push(foto);
  }

  // Guardar las fotos es lo más lento, así que va antes del candado y solo si
  // el enlace sirve; si al final el enlace ya no sirve, se mandan a la papelera.
  var previo = buscarAcceso_(datos.codigo);
  if (!previo) return { ok: false, error: "invalido" };
  if (previo.usado) return { ok: false, error: "usado" };
  var archivos = guardarFotos_(fotos, previo.nombre, datos);

  // Solo la consulta y la escritura van dentro del candado, para que muchos
  // envíos a la vez esperen su turno poco tiempo y nunca se use dos veces un enlace.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    tirarArchivos_(archivos);
    return { ok: false, error: "ocupado" };
  }
  try {
    var acceso = buscarAcceso_(datos.codigo);
    if (!acceso || acceso.usado) {
      tirarArchivos_(archivos);
      return { ok: false, error: acceso ? "usado" : "invalido" };
    }
    var enlaces = archivos.map(function(archivo) { return archivo.getUrl(); });
    obtenerHojaRespuestas_().appendRow([new Date(), acceso.nombre].concat(valores, enlaces));
    obtenerHojaAccesos_().getRange(acceso.fila, 3).setValue(new Date());
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Devuelve los bytes de la foto si es un JPEG válido y no demasiado grande.
function leerFoto_(base64) {
  if (typeof base64 !== "string" || !base64 || base64.length > MAX_FOTO_BYTES * 4 / 3) return null;
  try {
    var bytes = Utilities.base64Decode(base64);
    // Todo JPEG empieza con FF D8 FF.
    if (bytes.length < 3 || bytes[0] !== -1 || bytes[1] !== -40 || bytes[2] !== -1) return null;
    return bytes;
  } catch (err) {
    return null;
  }
}

function guardarFotos_(fotos, invitado, datos) {
  var carpeta = carpetaFotos_();
  var nombre = [datos.nombres, datos.paterno, datos.materno]
    .map(function(v) { return String(v || "").trim(); })
    .filter(String).join(" ");
  var base = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH.mm.ss") +
    " - " + invitado + " - " + nombre;
  return fotos.map(function(bytes, i) {
    return carpeta.createFile(Utilities.newBlob(bytes, "image/jpeg", base + " - " + FOTOS[i].titulo + ".jpg"));
  });
}

function tirarArchivos_(archivos) {
  archivos.forEach(function(archivo) { archivo.setTrashed(true); });
}

// La carpeta se crea la primera vez y se recuerda por su ID; no se comparte con nadie.
function carpetaFotos_() {
  var propiedades = PropertiesService.getScriptProperties();
  var id = propiedades.getProperty("CARPETA_FOTOS");
  if (id) {
    try {
      var existente = DriveApp.getFolderById(id);
      if (!existente.isTrashed()) return existente;
    } catch (err) {}
  }
  var carpeta = DriveApp.createFolder(CARPETA_FOTOS);
  propiedades.setProperty("CARPETA_FOTOS", carpeta.getId());
  return carpeta;
}

// Menú Accesos › Ver carpeta de fotos. Ejecutarla una vez desde el editor también
// sirve para dar el permiso de Drive antes de publicar.
function verCarpetaFotos() {
  var url = carpetaFotos_().getUrl();
  try {
    SpreadsheetApp.getUi().alert("Carpeta de fotos (privada):\n" + url);
  } catch (err) {
    Logger.log(url);
  }
}

function buscarAcceso_(codigo) {
  codigo = String(codigo || "").trim();
  if (!codigo) return null;
  var hoja = obtenerHojaAccesos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return null;
  var filas = hoja.getRange(2, 1, ultima - 1, 5).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][1]) === codigo) {
      var ilimitado = filas[i][4] === true || /^(s[ií]|x)$/i.test(String(filas[i][4]).trim());
      return { fila: i + 2, nombre: filas[i][0], usado: !!filas[i][2] && !ilimitado, ilimitado: ilimitado };
    }
  }
  return null;
}

function obtenerHojaAccesos_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(HOJA_ACCESOS);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_ACCESOS);
    hoja.appendRow(["Nombre", "Código", "Usado", "Enlace", "Ilimitado"]);
  }
  return hoja;
}

function encabezadosRespuestas_() {
  return ["Fecha", "Invitado"]
    .concat(CAMPOS.map(function(c) { return c.titulo; }))
    .concat(FOTOS.map(function(f) { return f.titulo; }));
}

function obtenerHojaRespuestas_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(HOJA_RESPUESTAS);
  if (!hoja) hoja = libro.insertSheet(HOJA_RESPUESTAS);
  var encabezados = encabezadosRespuestas_();
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(encabezados);
    darFormatoRespuestas_(hoja);
  } else {
    // Si cambiaron los campos del formulario, reacomoda antes de escribir para
    // que ninguna respuesta quede en la columna equivocada.
    var actuales = hoja.getRange(1, 1, 1, encabezados.length).getValues()[0];
    if (actuales.join("|") !== encabezados.join("|")) acomodarHoja_(hoja);
  }
  return hoja;
}

// Ancho y formato fijos para que la fecha no salga como ##### al bajar a Excel.
function darFormatoRespuestas_(hoja) {
  hoja.setColumnWidth(1, 170);
  hoja.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm");
}

// Menú Accesos › Acomodar columnas de Respuestas.
function acomodarRespuestas() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_RESPUESTAS);
  if (hoja && hoja.getLastRow() > 0) acomodarHoja_(hoja);
}

// Nombres viejos de columnas que hoy se llaman distinto.
var ALIAS_COLUMNAS = { "Nombre completo": "Nombre(s)" };

// Reescribe la pestaña con los encabezados actuales y cada dato en su columna.
// Un renglón con datos más allá de los encabezados viejos se escribió con los
// campos actuales y se deja tal cual; los demás se acomodan según los
// encabezados que tenía la pestaña. Antes deja un respaldo.
function acomodarHoja_(hoja) {
  var libro = hoja.getParent();
  var copia = hoja.copyTo(libro);
  copia.setName(HOJA_RESPUESTAS + " respaldo " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH.mm"));

  var datos = hoja.getDataRange().getValues();
  var viejos = datos[0].map(function(t) { t = String(t).trim(); return ALIAS_COLUMNAS[t] || t; });
  var nuevos = encabezadosRespuestas_();
  var anchoViejo = datos[0].filter(String).length;

  var filas = datos.slice(1).filter(function(fila) {
    return fila.some(function(v) { return v !== ""; });
  }).map(function(fila) {
    var actual = fila.slice(anchoViejo).some(function(v) { return v !== ""; });
    var acomodada = nuevos.map(function(titulo, k) {
      var i = actual ? k : viejos.indexOf(titulo);
      return i >= 0 && i < fila.length ? fila[i] : "";
    });
    // Igual que en enviarDatos: el apóstrofo conserva ceros a la izquierda.
    return acomodada.map(function(v, i) {
      return i < 2 || v === "" ? v : "'" + v;
    });
  });

  hoja.clear();
  darFormatoRespuestas_(hoja);
  hoja.getRange(1, 1, 1, nuevos.length).setValues([nuevos]);
  if (filas.length) hoja.getRange(2, 1, filas.length, nuevos.length).setValues(filas);
  SpreadsheetApp.flush();
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
  var archivos = carpetaFotos_().getFiles();
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    if (archivo.getName().indexOf(PRUEBA_NOMBRE) >= 0) archivo.setTrashed(true);
  }
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
