// Apps Script del formulario de datos. El mismo script muestra el formulario
// (archivo «Formulario» del proyecto) y guarda las respuestas.
//
// La hoja de cálculo usa dos pestañas, que se crean solas si no existen:
//   «Accesos»:    Nombre | Código | Usado | Enlace | Ilimitado | Abierto | Modificar
//   «Respuestas»: se llena sola; la fila de encabezados se crea si está vacía.
// Las fotos se guardan en la carpeta privada «Formulario - Fotos» de tu Drive y
// en «Respuestas» queda el enlace a cada una.
//
// Para invitar a alguien: escribe su nombre en la columna A de «Accesos» y usa
// el menú Accesos › Generar enlaces. Cada enlace sirve para un solo envío, salvo
// que su renglón diga «Sí» en la columna Ilimitado: ese enlace sirve para siempre
// (en Usado queda la fecha del último envío) hasta que borres el renglón o el «Sí».
// Si el renglón dice «Sí» en la columna Abierto, ese enlace abre el formulario
// con todos los campos opcionales (lo que se escriba se valida igual). Todo llega
// a la misma pestaña «Respuestas».
// Si dice «Sí» en la columna Modificar, el enlace abre el módulo Modificaciones:
// busca a una persona por nombre y deja corregir o completar su registro
// (fotos incluidas). Ese enlace ve todos los registros, compártelo con cuidado.
// Su primera columna, «Consecutivo», numera cada
// registro (1, 2, 3...), y la segunda, «Registro», dice RC (Registro
// Completo: trae todo lo que pide el formulario definitivo, fotos incluidas)
// o RI (Registro Incompleto: le falta algo), sin importar por qué enlace llegó.

var HOJA_ACCESOS = "Accesos";
var HOJA_RESPUESTAS = "Respuestas";

// URL de la implementación (Implementar › Administrar implementaciones).
var FORM_URL = "https://script.google.com/macros/s/AKfycby2PsQgH911FicEgApc58_V3F2NfemPVBLQi4t1w5muFxrwQ1fLxCY5Z57VgaAmaK612w/exec";

var CAMPOS = [
  { nombre: "tipo",      titulo: "Tipo de representante",     patron: /^(RG|RC)$/ },
  { nombre: "nombres",   titulo: "Nombre(s)",                 patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]+$/ },
  // Los dos apellidos son opcionales (hay quien solo tiene uno).
  { nombre: "paterno",   titulo: "Apellido paterno",          patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]*$/, opcional: true },
  { nombre: "materno",   titulo: "Apellido materno",          patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]*$/, opcional: true },
  { nombre: "calle",     titulo: "Calle",                     patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 .,#\-]+$/ },
  { nombre: "numero",    titulo: "Número",                    patron: /^[0-9]+$/ },
  { nombre: "interior",  titulo: "Número interior",           patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 .,#\-]*$/, opcional: true },
  { nombre: "colonia",   titulo: "Colonia o Fraccionamiento", patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 .,#\-]+$/ },
  { nombre: "postal",    titulo: "Código postal",             patron: /^[0-9]{1,5}$/ },
  { nombre: "municipio", titulo: "Municipio",                 patron: /^.{1,40}$/ },
  { nombre: "clave",     titulo: "Clave de elector",          patron: /^[A-Za-z0-9]{18}$/ },
  { nombre: "seccion",   titulo: "Sección",                   patron: /^[0-9]{1,4}$/ },
  { nombre: "ocr",       titulo: "OCR",                       patron: /^[0-9]{13}$/ },
  { nombre: "telefono",  titulo: "Teléfono",                  patron: /^[0-9]{10}$/ },
  { nombre: "correo",    titulo: "Correo electrónico",        patron: /^[^\s@]+@[^\s@]+\.[^\s@]+$/, opcional: true },
  // Va al final del formulario y de la hoja, después de las fotos.
  { nombre: "seccionTrabajo", titulo: "Sección de trabajo",   patron: /^[0-9]{1,4}$/, alFinal: true },
  { nombre: "capturadoPor", titulo: "Registro capturado por",  patron: /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ ]+$/, alFinal: true }
];

// Fotos que pide el formulario; cada una llega ya reducida a JPEG desde el
// navegador. Las dos son obligatorias en el formulario definitivo (no en el abierto).
var FOTOS = [
  { nombre: "foto1", titulo: "Foto 1", obligatoria: true },
  { nombre: "foto2", titulo: "Foto 2", obligatoria: true }
];
var CARPETA_FOTOS = "Formulario - Fotos";
var MAX_FOTO_BYTES = 5 * 1024 * 1024;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Accesos")
    .addItem("Generar enlaces", "generarEnlaces")
    .addItem("Acomodar columnas de Respuestas", "acomodarRespuestas")
    .addItem("Separar nombres y apellidos", "separarNombres")
    .addItem("Ver carpeta de fotos", "verCarpetaFotos")
    .addToUi();
}

// Llena Código y Enlace en cada fila de «Accesos» que tenga nombre y no tenga código.
function generarEnlaces() {
  var hoja = obtenerHojaAccesos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return;
  // Las hojas de antes no tienen la columna Abierto.
  if (!hoja.getRange(1, 6).getValue()) hoja.getRange(1, 6).setValue("Abierto");
  if (!hoja.getRange(1, 7).getValue()) hoja.getRange(1, 7).setValue("Modificar");
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
  plantilla.abierto = plantilla.estado === "ok" && acceso.abierto;
  plantilla.modificar = plantilla.estado === "ok" && acceso.modificar;
  // ALLOWALL deja insertar el formulario en otra página (por ejemplo Google Sites),
  // donde ya no aparece el aviso de Google Apps Script.
  return plantilla.evaluate()
    .setTitle("Plataforma de Captura")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// El formulario la llama con google.script.run al presionar «Enviar formulario».
function enviarDatos(datos) {
  datos = datos || {};
  // Primero el enlace: de él depende si los campos son obligatorios.
  var previo = buscarAcceso_(datos.codigo);
  if (!previo) return { ok: false, error: "invalido" };
  if (previo.usado) return { ok: false, error: "usado" };

  var valores = [];
  for (var i = 0; i < CAMPOS.length; i++) {
    var campo = CAMPOS[i];
    var valor = String(datos[campo.nombre] || "").trim();
    if ((!valor && !campo.opcional && !previo.abierto) || (valor && !campo.patron.test(valor))) {
      return { ok: false, error: "datos", message: "Revisa el campo «" + campo.titulo + "»." };
    }
    // El apóstrofo evita que Sheets quite ceros a la izquierda (OCR, CP, etc.).
    valores.push(valor ? "'" + valor : "");
  }

  repartirNombre_(valores);

  // En el formulario abierto puede faltar cualquier campo, pero no todos.
  var algo = valores.some(String) || FOTOS.some(function(f) { return datos[f.nombre]; });
  if (!algo) return { ok: false, error: "datos", message: "Llena al menos un campo." };

  // Una foto vacía se omite; una que sí viene pero no es válida se rechaza.
  var fotos = [];
  for (var j = 0; j < FOTOS.length; j++) {
    var enviada = datos[FOTOS[j].nombre];
    if (!enviada && FOTOS[j].obligatoria && !previo.abierto) {
      return { ok: false, error: "datos", message: "Sube la foto «" + FOTOS[j].titulo + "»." };
    }
    var foto = enviada ? leerFoto_(enviada) : null;
    if (enviada && !foto) return { ok: false, error: "datos", message: "Revisa la foto «" + FOTOS[j].titulo + "»." };
    fotos.push(foto);
  }

  // Guardar las fotos es lo más lento, así que va antes del candado y solo si
  // el enlace sirve; si al final el enlace ya no sirve, se mandan a la papelera.
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
    var enlaces = archivos.map(function(archivo) { return archivo ? archivo.getUrl() : ""; });
    var registro = estaCompleto_(function(titulo) {
      var k = CAMPOS.map(function(c) { return c.titulo; }).indexOf(titulo);
      if (k >= 0) return valores[k];
      k = FOTOS.map(function(f) { return f.titulo; }).indexOf(titulo);
      return k >= 0 ? enlaces[k] : "";
    }) ? "RC" : "RI";
    var hoja = obtenerHojaRespuestas_();
    var numero = siguienteNumero_(hoja);
    // Cada dato en la columna de su título (no todos los campos van seguidos).
    var enc = encabezadosRespuestas_();
    var renglon = enc.map(function() { return ""; });
    renglon[0] = numero;
    renglon[1] = registro;
    renglon[2] = new Date();
    renglon[3] = acceso.nombre;
    CAMPOS.forEach(function(c, k) { renglon[enc.indexOf(c.titulo)] = valores[k]; });
    FOTOS.forEach(function(f, k) { renglon[enc.indexOf(f.titulo)] = enlaces[k]; });
    hoja.appendRow(renglon);
    // Que el consecutivo se vea como número aunque la columna traiga formato de fecha.
    hoja.getRange(hoja.getLastRow(), 1).setNumberFormat("0");
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
    if (!bytes) return null;
    return carpeta.createFile(Utilities.newBlob(bytes, "image/jpeg", base + " - " + FOTOS[i].titulo + ".jpg"));
  });
}

function tirarArchivos_(archivos) {
  archivos.forEach(function(archivo) { if (archivo) archivo.setTrashed(true); });
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
  var filas = hoja.getRange(2, 1, ultima - 1, 7).getValues();
  for (var i = 0; i < filas.length; i++) {
    if (String(filas[i][1]) === codigo) {
      var ilimitado = esSi_(filas[i][4]);
      var modificar = esSi_(filas[i][6]);
      return { fila: i + 2, nombre: filas[i][0], usado: !!filas[i][2] && !ilimitado && !modificar,
        ilimitado: ilimitado, abierto: esSi_(filas[i][5]), modificar: modificar };
    }
  }
  return null;
}

// «Sí», «Si», «x» o una casilla marcada.
function esSi_(valor) {
  return valor === true || /^(s[ií]|x)$/i.test(String(valor).trim());
}

function obtenerHojaAccesos_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(HOJA_ACCESOS);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_ACCESOS);
    hoja.appendRow(["Nombre", "Código", "Usado", "Enlace", "Ilimitado", "Abierto", "Modificar"]);
  }
  return hoja;
}

function encabezadosRespuestas_() {
  var titulo = function(x) { return x.titulo; };
  return ["Consecutivo", "Registro", "Fecha", "Invitado"]
    .concat(CAMPOS.filter(function(c) { return !c.alFinal; }).map(titulo))
    .concat(FOTOS.map(titulo))
    .concat(CAMPOS.filter(function(c) { return c.alFinal; }).map(titulo))
    .concat(["Última modificación"]);
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

// true si el registro trae todo lo que pide el formulario definitivo: los
// campos obligatorios y las fotos obligatorias. «valor» da lo guardado por título.
function estaCompleto_(valor) {
  var titulos = CAMPOS.filter(function(c) { return !c.opcional; })
    .concat(FOTOS.filter(function(f) { return f.obligatoria; }))
    .map(function(x) { return x.titulo; });
  return titulos.every(function(t) { return String(valor(t) || "").replace(/^'/, "").trim() !== ""; });
}

// El número que sigue al del último renglón (1 si no hay registros). Se toma
// del último y no del total, para no repetir números si se borra un renglón.
function siguienteNumero_(hoja) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return 1;
  var anterior = Number(hoja.getRange(ultima, 1).getValue());
  return (anterior > 0 ? anterior : ultima - 1) + 1;
}

// Ancho y formato fijos para que la fecha no salga como ##### al bajar a Excel.
function darFormatoRespuestas_(hoja) {
  hoja.setColumnWidth(1, 50);
  // La columna A antes era la fecha; sin esto el número 79 se ve como 19/03/1900.
  hoja.getRange("A:A").setNumberFormat("0");
  hoja.setColumnWidth(2, 80);
  hoja.setColumnWidth(3, 170);
  hoja.getRange("C:C").setNumberFormat("dd/mm/yyyy hh:mm");
}

// Menú Accesos › Acomodar columnas de Respuestas.
function acomodarRespuestas() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_RESPUESTAS);
  if (hoja && hoja.getLastRow() > 0) acomodarHoja_(hoja);
}

// Nombres viejos de columnas que hoy se llaman distinto.
var ALIAS_COLUMNAS = { "Nombre completo": "Nombre(s)", "No.": "Consecutivo", "Sección 2": "Sección de trabajo" };

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
  var ultimoNumero = 0;

  var filas = datos.slice(1).filter(function(fila) {
    return fila.some(function(v) { return v !== ""; });
  }).map(function(fila) {
    var actual = fila.slice(anchoViejo).some(function(v) { return v !== ""; });
    var acomodada = nuevos.map(function(titulo, k) {
      var i = actual ? k : viejos.indexOf(titulo);
      return i >= 0 && i < fila.length ? fila[i] : "";
    });
    // Los renglones sin número se numeran en orden, después del último que ya tenga.
    var numero = Number(acomodada[0]);
    if (numero > 0) ultimoNumero = Math.max(ultimoNumero, numero);
    else acomodada[0] = ++ultimoNumero;
    // RC o RI se vuelve a calcular con lo que trae cada renglón.
    acomodada[nuevos.indexOf("Registro")] = estaCompleto_(function(titulo) {
      return acomodada[nuevos.indexOf(titulo)];
    }) ? "RC" : "RI";
    // Igual que en enviarDatos: el apóstrofo conserva ceros a la izquierda.
    return acomodada.map(function(v, i) {
      return i < 4 || v === "" ? v : "'" + v;
    });
  });

  hoja.clear();
  darFormatoRespuestas_(hoja);
  hoja.getRange(1, 1, 1, nuevos.length).setValues([nuevos]);
  if (filas.length) hoja.getRange(2, 1, filas.length, nuevos.length).setValues(filas);
  SpreadsheetApp.flush();
}

// ---------- Separar nombres ----------
// Palabras que forman parte del nombre o apellido que les sigue («de la Cruz»).
var PARTICULAS = ["de", "del", "la", "las", "los", "y", "san", "santa", "van", "von", "da", "di"];

function sinAcentos_(texto) {
  return texto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Reparte un nombre completo en { nombres, paterno, materno }:
//   2 pedazos: nombre | paterno
//   3 pedazos: nombre | paterno | materno
//   4 o más:   los dos últimos son paterno y materno; lo demás, nombres.
// Con un solo pedazo devuelve null (no hay nada que separar).
function partirNombre_(completo) {
  // «TavaresMojica» → «Tavares Mojica».
  var palabras = String(completo).replace(/([a-záéíóúüñ])([A-ZÁÉÍÓÚÜÑ])/g, "$1 $2")
    .trim().split(/\s+/).filter(String);
  // Las partículas se pegan a la palabra que sigue: «de», «la», «Cruz» → «de la Cruz».
  var pedazos = [];
  var pendiente = [];
  palabras.forEach(function(p) {
    pendiente.push(p);
    if (PARTICULAS.indexOf(sinAcentos_(p)) < 0) {
      pedazos.push(pendiente.join(" "));
      pendiente = [];
    }
  });
  if (pendiente.length) {
    if (pedazos.length) pedazos[pedazos.length - 1] += " " + pendiente.join(" ");
    else pedazos.push(pendiente.join(" "));
  }
  if (pedazos.length < 2) return null;

  var corte = Math.max(1, pedazos.length - 2);
  return {
    nombres: pedazos.slice(0, corte).join(" "),
    paterno: pedazos[corte] || "",
    materno: pedazos.slice(corte + 1).join(" ")
  };
}

function indiceCampo_(nombre) {
  for (var i = 0; i < CAMPOS.length; i++) if (CAMPOS[i].nombre === nombre) return i;
  return -1;
}

// Si escribió nombre y apellidos juntos en «Nombre(s)», se reparten.
function repartirNombre_(valores) {
  var iNombres = indiceCampo_("nombres");
  if (valores[iNombres + 1] || valores[iNombres + 2]) return;
  var partes = partirNombre_(String(valores[iNombres]).replace(/^'/, ""));
  if (!partes) return;
  valores[iNombres] = "'" + partes.nombres;
  valores[iNombres + 1] = "'" + partes.paterno;
  valores[iNombres + 2] = partes.materno ? "'" + partes.materno : "";
}

// Menú Accesos › Separar nombres y apellidos. En los renglones de «Respuestas»
// sin apellidos, reparte el nombre completo escrito en «Nombre(s)». Antes deja
// un respaldo de la pestaña.
function separarNombres() {
  var hoja = obtenerHojaRespuestas_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return;
  var col = 5 + indiceCampo_("nombres"); // después de Consecutivo, Registro, Fecha e Invitado
  var rango = hoja.getRange(2, col, ultima - 1, 3);
  var filas = rango.getValues();
  var cambios = 0;
  var nuevas = filas.map(function(fila) {
    var partes = !String(fila[1]).trim() && !String(fila[2]).trim() && partirNombre_(fila[0]);
    if (!partes) return fila.map(function(v) { return v === "" ? "" : "'" + v; });
    cambios++;
    return ["'" + partes.nombres, "'" + partes.paterno, partes.materno ? "'" + partes.materno : ""];
  });
  if (!cambios) {
    SpreadsheetApp.getUi().alert("No hay nombres que separar.");
    return;
  }
  hoja.copyTo(hoja.getParent()).setName(HOJA_RESPUESTAS + " respaldo " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH.mm"));
  rango.setValues(nuevas);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Listo: se separaron " + cambios + " nombres. " +
    "Hay una copia de antes en la pestaña de respaldo.");
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
    var col = hoja.getName() === HOJA_ACCESOS ? 1 : 4;
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

// ---------- Modificaciones ----------
// Solo para enlaces con «Sí» en la columna Modificar de «Accesos».

function accesoModificar_(codigo) {
  var acceso = buscarAcceso_(codigo);
  return acceso && acceso.modificar ? acceso : null;
}

// Para comparar nombres sin importar acentos, mayúsculas ni signos.
function normalizar_(texto) {
  return sinAcentos_(String(texto || "")).replace(/[^a-z0-9ñ ]/g, " ").replace(/\s+/g, " ").trim();
}

// Busca por nombre completo o incompleto: cada palabra escrita debe aparecer
// en el nombre, en cualquier orden («juan vargas» encuentra a Juan Ibarra Vargas).
function buscarRegistros(codigo, texto) {
  if (!accesoModificar_(codigo)) return { ok: false, error: "invalido" };
  var palabras = normalizar_(texto).split(" ").filter(String);
  if (!palabras.length || normalizar_(texto).length < 2) {
    return { ok: false, message: "Escribe al menos 2 letras del nombre." };
  }
  var datos = obtenerHojaRespuestas_().getDataRange().getValues();
  var enc = datos[0];
  var col = function(t) { return enc.indexOf(t); };
  var iNom = col("Nombre(s)"), iPat = col("Apellido paterno"), iMat = col("Apellido materno");
  var resultados = [];
  for (var r = 1; r < datos.length; r++) {
    var fila = datos[r];
    var nombre = [fila[iNom], fila[iPat], fila[iMat]].map(String).filter(String).join(" ");
    var normal = normalizar_(nombre);
    if (!normal || !palabras.every(function(p) { return normal.indexOf(p) >= 0; })) continue;
    resultados.push({
      consecutivo: fila[col("Consecutivo")],
      nombre: nombre,
      seccion: String(fila[col("Sección")]),
      municipio: String(fila[col("Municipio")]),
      registro: String(fila[col("Registro")])
    });
  }
  return { ok: true, total: resultados.length, resultados: resultados.slice(0, 30) };
}

// Renglón (1, 2...) de «Respuestas» con ese consecutivo, o 0 si no está.
function filaDeConsecutivo_(hoja, consecutivo) {
  var n = Number(consecutivo);
  var ultima = hoja.getLastRow();
  if (!(n > 0) || ultima < 2) return 0;
  var columna = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  for (var i = 0; i < columna.length; i++) if (Number(columna[i][0]) === n) return i + 2;
  return 0;
}

function obtenerRegistro(codigo, consecutivo) {
  if (!accesoModificar_(codigo)) return { ok: false, error: "invalido" };
  var hoja = obtenerHojaRespuestas_();
  var fila = filaDeConsecutivo_(hoja, consecutivo);
  if (!fila) return { ok: false, message: "No se encontró ese registro. Búscalo otra vez." };
  var enc = encabezadosRespuestas_();
  var valores = hoja.getRange(fila, 1, 1, enc.length).getValues()[0];
  var campos = {};
  CAMPOS.forEach(function(c) { campos[c.nombre] = String(valores[enc.indexOf(c.titulo)]); });
  var fotos = {};
  FOTOS.forEach(function(f) { fotos[f.nombre] = String(valores[enc.indexOf(f.titulo)]) !== ""; });
  return { ok: true, consecutivo: valores[0], registro: valores[1], campos: campos, fotos: fotos };
}

// Guarda los cambios de un registro. Un campo vacío queda vacío; una foto que
// no se manda conserva la que ya tenía.
function guardarModificacion(datos) {
  datos = datos || {};
  var previo = accesoModificar_(datos.codigo);
  if (!previo) return { ok: false, error: "invalido" };

  var valores = [];
  for (var i = 0; i < CAMPOS.length; i++) {
    var campo = CAMPOS[i];
    var valor = String(datos[campo.nombre] || "").trim();
    if (valor && !campo.patron.test(valor)) {
      return { ok: false, error: "datos", message: "Revisa el campo «" + campo.titulo + "»." };
    }
    valores.push(valor ? "'" + valor : "");
  }
  if (!valores.some(String)) return { ok: false, error: "datos", message: "El registro no puede quedar vacío." };
  repartirNombre_(valores);

  var fotos = [];
  for (var j = 0; j < FOTOS.length; j++) {
    var enviada = datos[FOTOS[j].nombre];
    var foto = enviada ? leerFoto_(enviada) : null;
    if (enviada && !foto) return { ok: false, error: "datos", message: "Revisa la foto «" + FOTOS[j].titulo + "»." };
    fotos.push(foto);
  }
  var archivos = guardarFotos_(fotos, previo.nombre + " (modificación)", datos);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    tirarArchivos_(archivos);
    return { ok: false, error: "ocupado" };
  }
  try {
    var hoja = obtenerHojaRespuestas_();
    var fila = filaDeConsecutivo_(hoja, datos.consecutivo);
    if (!fila) {
      tirarArchivos_(archivos);
      return { ok: false, message: "No se encontró ese registro. Búscalo otra vez." };
    }
    var enc = encabezadosRespuestas_();
    var rango = hoja.getRange(fila, 1, 1, enc.length);
    var actual = rango.getValues()[0];
    var nuevo = actual.map(function(v, k) {
      // Igual que al guardar: el apóstrofo conserva ceros a la izquierda.
      return k < 4 || v === "" || v instanceof Date ? v : "'" + v;
    });
    CAMPOS.forEach(function(c, k) { nuevo[enc.indexOf(c.titulo)] = valores[k]; });
    FOTOS.forEach(function(f, k) {
      if (archivos[k]) nuevo[enc.indexOf(f.titulo)] = archivos[k].getUrl();
    });
    nuevo[enc.indexOf("Registro")] = estaCompleto_(function(t) { return nuevo[enc.indexOf(t)]; }) ? "RC" : "RI";
    nuevo[enc.indexOf("Última modificación")] = "'" +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + " - " + previo.nombre;
    rango.setValues([nuevo]);
    SpreadsheetApp.flush();
    return { ok: true, consecutivo: actual[0], registro: nuevo[enc.indexOf("Registro")] };
  } finally {
    lock.releaseLock();
  }
}
