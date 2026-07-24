/*************************************************************************
 *  INNOBA - Webhook de sincronizacion  (COTIZACIONES + RESERVAS)
 *  --------------------------------------------------------------------
 *  - Recibe cotizaciones del HTML de los clientes (como hasta ahora).
 *  - Recibe y entrega RESERVAS entre los computadores del equipo.
 *  Guarda cada registro como una fila en una hoja de calculo que este
 *  mismo script crea automaticamente la primera vez (no hay que crearla).
 *
 *  Como instalarlo:
 *   1) Abre https://script.google.com  y entra a TU proyecto actual
 *      (el que ya usas para el webhook de cotizaciones).
 *   2) Borra TODO el codigo que tengas y pega ESTE codigo completo.
 *   3) Guarda (icono del disquete).
 *   4) Implementar > Administrar implementaciones > (la que ya existe) >
 *      icono del lapiz (Editar) > Version: "Nueva version" > Implementar.
 *      *** IMPORTANTE: edita la implementacion EXISTENTE, asi la URL /exec
 *          NO cambia y no hay que tocar el programa. ***
 *   5) La primera vez te pedira autorizar los permisos: acepta.
 *************************************************************************/

// Clave para LEER (solo el programa .exe la tiene; el HTML nunca lee).
var LEER_KEY = 'inb_9f3Kx72Qp_seg2026';

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var tipo = _tipo(body.tipo);
    var id = String(body.id || body.uid || body.web_id || body.numero ||
                    Utilities.getUuid());
    // Solicitud de RESERVA desde la web: guarda los pasaportes en Drive y deja el link
    if (tipo === 'solicitudes' && body.pasajeros) {
      body.pasajeros = _guardarPasaportes(body.pasajeros, id);
    }
    _guardar(tipo, id, body);   // un borrado llega como marca {_accion:'borrar'} y solo
    return _json({ ok: true, id: id, tipo: tipo });   // sobrescribe la fila (tombstone)
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

/* Guarda los pasaportes (base64) en una carpeta de Drive y reemplaza el archivo
   por su LINK, para no exceder el limite de una celda de la hoja. */
function _guardarPasaportes(pasajeros, id) {
  // Si Drive no esta autorizado, NO se pierde la solicitud: se guarda igual sin archivos.
  var folder = null;
  try {
    var props = PropertiesService.getScriptProperties();
    var fid = props.getProperty('DRIVE_FOLDER');
    if (fid) {
      try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; }
    }
    if (!folder) {
      folder = DriveApp.createFolder('INNOBA Pasaportes');
      props.setProperty('DRIVE_FOLDER', folder.getId());
    }
  } catch (e) {
    folder = null;   // sin permiso de Drive
  }
  for (var i = 0; i < pasajeros.length; i++) {
    var p = pasajeros[i];
    if (p && p.archivo_b64) {
      if (folder) {
        try {
          var partes = String(p.archivo_b64).split(',');
          var datos = partes.length > 1 ? partes[1] : partes[0];
          var mime = (String(p.archivo_b64).match(/data:([^;]+);/) || [null, 'image/jpeg'])[1];
          var nombre = (p.archivo_nombre || ('pasaporte_' + (i + 1)));
          var blob = Utilities.newBlob(Utilities.base64Decode(datos), mime,
                                       id + '_' + nombre);
          var f = folder.createFile(blob);
          f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          p.pasaporte_url = f.getUrl();
        } catch (err) {
          p.pasaporte_url = 'ERROR: ' + err;
        }
      } else {
        p.pasaporte_url = 'PENDIENTE (falta autorizar Drive en el script)';
      }
      delete p.archivo_b64;   // nunca guardar el archivo dentro de la hoja
    }
  }
  return pasajeros;
}

/* Ejecuta esta funcion UNA VEZ desde el editor (boton Ejecutar) para autorizar
   Google Drive. Crea la carpeta donde se guardaran los pasaportes. */
function AUTORIZAR_DRIVE() {
  var props = PropertiesService.getScriptProperties();
  var folder = DriveApp.createFolder('INNOBA Pasaportes');
  props.setProperty('DRIVE_FOLDER', folder.getId());
  return 'Carpeta creada: ' + folder.getUrl();
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.key !== LEER_KEY) {
    return _json({ error: 'no autorizado' });
  }
  var tipo = _tipo(p.tipo);
  return _json({ items: _leer(tipo) });
}

// 'reserva'->reservas ; 'solicitud'->solicitudes ; cualquier otra cosa -> cotizaciones
function _tipo(t) {
  t = String(t || '').toLowerCase();
  if (t === 'reserva' || t === 'reservas') return 'reservas';
  if (t === 'solicitud' || t === 'solicitudes') return 'solicitudes';
  return 'cotizaciones';
}

function _hoja(nombre) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('INNOBA Sync (no borrar)');
    props.setProperty('SHEET_ID', ss.getId());
  }
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(['id', 'fecha', 'json']);
  }
  return sh;
}

function _guardar(tipo, id, obj) {
  var sh = _hoja(tipo);
  var datos = sh.getDataRange().getValues();
  var texto = JSON.stringify(obj);
  // Si el id ya existe, ACTUALIZA esa fila (para reservas que cambian de estado).
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(id)) {
      sh.getRange(i + 1, 2).setValue(new Date());
      sh.getRange(i + 1, 3).setValue(texto);
      return;
    }
  }
  sh.appendRow([id, new Date(), texto]);
}

function _leer(tipo) {
  var sh = _hoja(tipo);
  var datos = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < datos.length; i++) {
    try { out.push(JSON.parse(datos[i][2])); } catch (err) {}
  }
  return out;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
