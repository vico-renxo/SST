// ============================================================
//  ComunicadosCode.js — Gestión de Comunicados / Eventos
//  Hoja de cálculo: PERSONAL (misma que login)
//  Hoja interna:    "COMUNICADOS"
//  Columnas:
//    A=ID  B=Título  C=Tipo  D=URL/Contenido  E=Descripción
//    F=FechaDesde  G=FechaHasta  H=Activo  I=CreadoPor  J=FechaCreacion
// ============================================================

var COMUNICADOS_SHEET = 'COMUNICADOS';

function _getHojaComunicados() {
  var ss = getSpreadsheetPersonal();
  var hoja = ss.getSheetByName(COMUNICADOS_SHEET);
  if (!hoja) {
    hoja = ss.insertSheet(COMUNICADOS_SHEET);
    hoja.appendRow(['ID','Título','Tipo','URL_Contenido','Descripción',
                    'FechaDesde','FechaHasta','Activo','CreadoPor','FechaCreacion']);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function _rowToComunicado(r) {
  var TZ = Session.getScriptTimeZone();
  var fmt = function(d) {
    if (!d || !(d instanceof Date)) return '';
    return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  };
  return {
    id:          String(r[0] || ''),
    titulo:      String(r[1] || ''),
    tipo:        String(r[2] || 'texto'),        // 'video' | 'imagen' | 'texto'
    contenido:   String(r[3] || ''),
    descripcion: String(r[4] || ''),
    fechaDesde:  fmt(r[5]),
    fechaHasta:  fmt(r[6]),
    activo:      String(r[7] || '').toUpperCase() === 'SI',
    creadoPor:   String(r[8] || ''),
    fechaCreacion: fmt(r[9])
  };
}

// ── Listar todos (admin) ─────────────────────────────────────────────────
function obtenerComunicados() {
  try {
    var hoja = _getHojaComunicados();
    var last = hoja.getLastRow();
    if (last < 2) return { ok: true, items: [] };
    var data = hoja.getRange(2, 1, last - 1, 10).getValues();
    var items = data
      .filter(function(r) { return r[0] !== ''; })
      .map(_rowToComunicado)
      .reverse();
    return { ok: true, items: items };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Comunicado activo para hoy (para trabajadores) ───────────────────────
// Retorna el comunicado activo cuyo rango de fechas incluye hoy
function obtenerComunicadoActivo() {
  try {
    var hoja = _getHojaComunicados();
    var last = hoja.getLastRow();
    if (last < 2) return { ok: true, comunicado: null };

    var data = hoja.getRange(2, 1, last - 1, 10).getValues();
    var TZ   = Session.getScriptTimeZone();
    var hoyStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

    for (var i = data.length - 1; i >= 0; i--) {
      var r = data[i];
      if (!r[0]) continue;
      var activo = String(r[7] || '').toUpperCase() === 'SI';
      if (!activo) continue;

      var desde = r[5] instanceof Date ? Utilities.formatDate(r[5], TZ, 'yyyy-MM-dd') : String(r[5] || '');
      var hasta = r[6] instanceof Date ? Utilities.formatDate(r[6], TZ, 'yyyy-MM-dd') : String(r[6] || '');

      if ((!desde || hoyStr >= desde) && (!hasta || hoyStr <= hasta)) {
        return { ok: true, comunicado: _rowToComunicado(r) };
      }
    }
    return { ok: true, comunicado: null };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Guardar (crear o editar) ─────────────────────────────────────────────
function guardarComunicado(obj) {
  try {
    var hoja = _getHojaComunicados();
    var isEdit = !!(obj.id && obj.id.trim() !== '');

    var parseDate = function(s) {
      if (!s) return '';
      var d = new Date(s + 'T00:00:00');
      return isNaN(d) ? '' : d;
    };

    var fila = [
      isEdit ? obj.id : Utilities.getUuid().substring(0, 8).toUpperCase(),
      obj.titulo      || '',
      obj.tipo        || 'texto',
      obj.contenido   || '',
      obj.descripcion || '',
      parseDate(obj.fechaDesde),
      parseDate(obj.fechaHasta),
      (obj.activo === true || obj.activo === 'SI') ? 'SI' : 'NO',
      obj.creadoPor   || '',
      isEdit ? (parseDate(obj.fechaCreacion) || new Date()) : new Date()
    ];

    if (isEdit) {
      var last = hoja.getLastRow();
      if (last < 2) return { ok: false, error: 'No encontrado' };
      var ids = hoja.getRange(2, 1, last - 1, 1).getValues().flat();
      var idx = ids.findIndex(function(id) { return String(id) === String(obj.id); });
      if (idx === -1) return { ok: false, error: 'ID no encontrado' };
      hoja.getRange(idx + 2, 1, 1, 10).setValues([fila]);
    } else {
      hoja.appendRow(fila);
    }
    return { ok: true, id: fila[0] };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Eliminar ─────────────────────────────────────────────────────────────
function eliminarComunicado(id) {
  try {
    var hoja = _getHojaComunicados();
    var last = hoja.getLastRow();
    if (last < 2) return { ok: false, error: 'No encontrado' };
    var ids = hoja.getRange(2, 1, last - 1, 1).getValues().flat();
    var idx = ids.findIndex(function(x) { return String(x) === String(id); });
    if (idx === -1) return { ok: false, error: 'ID no encontrado' };
    hoja.deleteRow(idx + 2);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Subir imagen local a Drive ────────────────────────────────────────────
// Carpeta destino: https://drive.google.com/drive/folders/1_a0rg1PK13NtkDLQ-y-tqKslo9aOc8DV
var COMUNICADOS_FOLDER_ID = '1_a0rg1PK13NtkDLQ-y-tqKslo9aOc8DV';

function subirImagenComunicado(base64, nombre, mimeType) {
  try {
    var folder = DriveApp.getFolderById(COMUNICADOS_FOLDER_ID);
    var blob   = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nombre);
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    // Thumbnail URL es la más confiable para mostrar en <img> sin redirecciones
    var url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w2000';
    return { ok: true, url: url, fileId: fileId };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Toggle activo/inactivo ────────────────────────────────────────────────
function toggleActivoComunicado(id, activo) {
  try {
    var hoja = _getHojaComunicados();
    var last = hoja.getLastRow();
    if (last < 2) return { ok: false, error: 'No encontrado' };
    var ids = hoja.getRange(2, 1, last - 1, 1).getValues().flat();
    var idx = ids.findIndex(function(x) { return String(x) === String(id); });
    if (idx === -1) return { ok: false, error: 'ID no encontrado' };
    hoja.getRange(idx + 2, 8).setValue(activo ? 'SI' : 'NO');
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}