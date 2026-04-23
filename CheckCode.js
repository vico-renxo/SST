/* DEFINE GLOBAL VARIABLES, CHANGE THESE VARIABLES TO MATCH WITH YOUR SHEET */

const folderimgcheck = "13qGGx2VJRlbcPSw9b2Ldn10HlJQJg0zd" // ARCHIVO 2 "Imagen inspecciones"
let folderpdfcheck = "1Be7s5TlJRS6sj0f6NxhGPK9EqFJcVM6N" // ARCHIVO 3 "PDF inspecciones borrar"

let cachedCheck = null;
function getCheckSpreadsheet() {
  if (!cachedCheck) {
    cachedCheck = SpreadsheetApp.openById(SPREADSHEET_IDS.check);
  }
  return cachedCheck;
}

function getDatosRegistroCheck(offset, limit, filtroMes, filtroEquipo) {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("B DATOS");
    const lastRow = hoja.getLastRow();
    const datos = hoja.getRange(1, 1, lastRow, 98).getDisplayValues();

    const headers = datos[0].slice(0, 98);
    const registros = datos.slice(1).map(fila => fila.slice(0, 98));

    const lowerFiltroEquipo = (filtroEquipo || "").toLowerCase();
    const mesNum = (filtroMes && filtroMes !== "Todos" && filtroMes !== "todos") ? parseInt(filtroMes, 10) : null;

    const filtrados = registros.filter(fila => {
      // Filtro por equipo (col C = index 2)
      if (lowerFiltroEquipo && lowerFiltroEquipo !== "todos") {
        if (fila[2].toLowerCase() !== lowerFiltroEquipo) return false;
      }

      // Filtro por mes (col J = index 9) — sin restricción de año
      if (mesNum !== null) {
        var ts = _parseFechaCheck(fila[9]);
        if (ts === 0) return false;
        if (new Date(ts).getMonth() + 1 !== mesNum) return false;
      }

      return true;
    });

    // Ordenar por fecha descendente (col J = index 9, formato dd/MM/yyyy o similar)
    var colFecha = headers.indexOf('Fecha');
    if (colFecha === -1) colFecha = 9; // fallback a col J
    filtrados.sort(function(a, b) {
      var fa = _parseFechaCheck(a[colFecha]);
      var fb = _parseFechaCheck(b[colFecha]);
      return fb - fa; // descendente: más reciente primero
    });

    const paginados = filtrados.slice(offset, offset + limit);

    return {
      headers,
      data: paginados,
      total: filtrados.length
    };
  } catch (error) {
    Logger.log("⚠️ Error en getDatosRegistroCheck: " + error.message);
    return {
      headers: [],
      data: [],
      total: 0,
      error: error.message
    };
  }
}

// Parser de fecha texto → timestamp para ordenar
// Acepta: "26/2/2026 15:17:58", "26/02/2026", "2026-02-26", etc.
function _parseFechaCheck(txt) {
  if (!txt) return 0;
  var s = String(txt).trim();
  // dd/MM/yyyy HH:mm:ss o dd/MM/yyyy
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    return new Date(
      parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]),
      parseInt(m[4] || 0), parseInt(m[5] || 0), parseInt(m[6] || 0)
    ).getTime() || 0;
  }
  // yyyy-MM-dd
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3])).getTime() || 0;
  // Fallback
  var d = new Date(s);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function getEquiposCheck() {
  try {
    var ss = getCheckSpreadsheet();
    var invSheet = ss.getSheetByName('INVENTARIO');
    var lastRow = invSheet.getLastRow();
    if (lastRow < 3) return [];
    var data = invSheet.getRange(3, 4, lastRow - 2, 1).getDisplayValues(); // col D = equipos, desde fila 3
    var unique = {};
    data.forEach(function(r) {
      var v = String(r[0] || '').trim();
      if (v) unique[v] = true;
    });
    return Object.keys(unique).sort();
  } catch(e) { return []; }
}

function getHeadersCheck() {
  const hoja = getCheckSpreadsheet().getSheetByName("B DATOS");
  // ✅ Hasta columna 98 (CT)
  return hoja.getRange(1, 1, 1, 98).getValues()[0]; 
}

function globalVariables() {
  var spreadsheet = getCheckSpreadsheet();

  return {
    spreadsheetId : spreadsheet.getId(),
    dataRage      : 'B DATOS!A2:O',
    idRange       : 'B DATOS!A2:A',
    lastCol       : 'O',
    sheetID       : '675860866'
  };
}

function updateCell(value) {
  var sheetName = "ACTUAL";
  var cellAddress = "J2";
  var sheet = getCheckSpreadsheet().getSheetByName(sheetName);
  if (sheet) {
    sheet.getRange(cellAddress).setValue(value);
  }
}

function getAccessPasswords() {
  const sheet = getCheckSpreadsheet().getSheetByName('Acceso');
  const colB = sheet.getRange('B2:B').getValues().flat().filter(String);
  const colC = sheet.getRange('C2:C').getValues().flat().filter(String);
  return {
    loginPasswords: colB,
    deletePasswords: colC
  };
}

function readData(spreadsheetId, range) {
  var result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

function deleteData(ID) { 
  var startIndex = getRowIndexByID(ID);
  
  // ✅ ANTES DE BORRAR: Eliminar fotos asociadas del Drive
  try {
    var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
    var rowNum = startIndex + 1; // startIndex es 0-based (fila header=0), rowNum es 1-based para getRange
    var rowData = sheet.getRange(rowNum, 1, 1, 98).getValues()[0];
    
    // Recopilar todas las URLs de imágenes de esta fila
    var urlsToDelete = [];
    
    // Columna 11 (K) = Imagen principal
    if (rowData[10]) urlsToDelete.push(String(rowData[10]).trim());
    // Columna 12 (L) = Responsable (nombre de texto, no URL — se omite del borrado)
    
    // Columnas 69-98 (BQ-CT) = Fotos de observación y subsanación por sección
    // Soporta nuevo formato codificado "num::url~~num::url" y formato antiguo (URL directa)
    for (var col = 68; col < 98; col++) {
      var cellVal = String(rowData[col] || '').trim();
      if (!cellVal) continue;
      if (cellVal.includes('::')) {
        cellVal.split('~~').forEach(function(part) {
          var si = part.indexOf('::');
          if (si !== -1) urlsToDelete.push(part.substring(si + 2).trim());
        });
      } else {
        urlsToDelete.push(cellVal);
      }
    }
    
    // Eliminar cada archivo del Drive
    urlsToDelete.forEach(function(url) {
      if (!url || url === '' || url === 'NA') return;
      try {
        var fileId = '';
        // Formato: https://lh5.googleusercontent.com/d/FILE_ID
        if (url.includes('googleusercontent.com/d/')) {
          fileId = url.split('/d/')[1].split(/[?#\/]/)[0];
        }
        // Formato: https://drive.google.com/file/d/FILE_ID/view
        else if (url.includes('drive.google.com/file/d/')) {
          fileId = url.split('/d/')[1].split('/')[0];
        }
        // Formato: https://drive.google.com/open?id=FILE_ID
        else if (url.includes('id=')) {
          fileId = url.split('id=')[1].split('&')[0];
        }
        
        if (fileId) {
          DriveApp.getFileById(fileId).setTrashed(true);
          Logger.log("🗑️ Foto eliminada: " + fileId);
        }
      } catch (e) {
        Logger.log("⚠️ No se pudo eliminar foto: " + url + " - " + e.message);
      }
    });
    
    Logger.log("✅ " + urlsToDelete.filter(u => u && u !== '' && u !== 'NA').length + " fotos procesadas para eliminación del check " + ID);
  } catch (e) {
    Logger.log("⚠️ Error al eliminar fotos del Drive: " + e.message);
  }
  
  // ✅ BORRAR LA FILA
  var deleteRange = {
    "sheetId"     : globalVariables().sheetID,
    "dimension"   : "ROWS",
    "startIndex"  : startIndex,
    "endIndex"    : startIndex + 1
  };
  
  var deleteRequest = [{"deleteDimension":{"range":deleteRange}}];
  Sheets.Spreadsheets.batchUpdate({"requests": deleteRequest}, globalVariables().spreadsheetId);
}

function getRowIndexByID(id) {
  if(id) {
    var idList = readData(globalVariables().spreadsheetId, globalVariables().idRange);
    for(var i = 0; i < idList.length; i++) {
      if(id == idList[i][0]) {
        var rowIndex = parseInt(i + 1);
        return rowIndex;
      }
    }
  }
}

function setStatusCheck(){
  let sst = getCheckSpreadsheet().getSheetByName('B DATOS')
  let totalCheck1 = sst.getRange("T1").getValue();
  let totalCheck2 = sst.getRange("U1").getValue();

  return[totalCheck1, totalCheck2]
}

function getURL() {
  return ScriptApp.getService().getUrl();
}

function searchData(obj) {
  const sheet = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const dataToSearch = sheet.getRange(1, 1, lastRow, 3).getDisplayValues();

  const output = [];

  for (let i = 0; i < dataToSearch.length; i++) {
    if (dataToSearch[i].includes(obj.ad3)) {
      output.push(allData[i]);
    }
  }

  return output;
}

function saveDataCheck(obj) {
  try {
    var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');

    // Remover filtro activo si existe (evita error "No se admite esta operación en un rango con una fila filtrada")
    try { var _f = sheet.getFilter(); if (_f) _f.remove(); } catch(e) {}

    var folder = DriveApp.getFolderById(folderimgcheck);
    var imageUrl = '';

    var idEquipo = String(obj.ad4).trim();

    // 1) Subir imagen si existe
    if (obj.imageData) {
      var imageData = Utilities.base64Decode(obj.imageData.split(',')[1]);
      var blob = Utilities.newBlob(imageData, MimeType.PNG, obj.ad3 + ".png");
      var file = folder.createFile(blob);
      var fileId = file.getId();
      imageUrl = "https://lh5.googleusercontent.com/d/" + fileId;
    } else {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idColValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
        var imgColValues = sheet.getRange(2, 10, lastRow - 1, 1).getValues();

        for (var i = idColValues.length - 1; i >= 0; i--) {
          var idCelda = String(idColValues[i][0]).trim();
          var urlCelda = imgColValues[i][0];
          if (idCelda === idEquipo && urlCelda) {
            imageUrl = urlCelda;
            break;
          }
        }
      }
    }

    // 2) Registrar fila en "B DATOS"
    var timestamp = Math.floor(new Date().getTime() / 1000);
    var newValue = timestamp;
    var status = obj.checked.includes("No") ? "Abierto" : "Conforme";

    // Tipo de inspección (Planeada / No Planeada) — campo N
    var tipoInspeccion = (obj.tipoInspeccion && obj.tipoInspeccion.trim()) ? obj.tipoInspeccion.trim() : 'Planeada';

    var checked = "";
    if (obj.checked && typeof obj.checked === 'string') {
      checked = obj.checked;
    } else if (Array.isArray(obj.checked)) {
      checked = obj.checked.join(",");
    }

    // ✅ GUARDAR 15 COLUMNAS (A-O)
    var rowData = [
      newValue,
      obj.ad1,          // col B = Empresa
      obj.ad3,          // col C = Equipo
      obj.ad4,          // col D = Código/Placa
      obj.ad2,          // col E = Área
      obj.ad9,          // col F = Proceso
      obj.adSup || '',  // col G = Supervisor de Operaciones
      obj.ad5,          // col H = Lugar
      obj.ad7,          // col I = Plan de Acción
      new Date(),       // col J = Fecha
      imageUrl,         // col K = Imagen principal
      obj.ad6,          // col L = Responsable (persona presente)
      status,           // col M = Estado
      tipoInspeccion,   // col N = Tipo de Inspección (Planeada / No Planeada)
      checked           // col O = Items
    ];

    sheet.appendRow(rowData);

    // ✅ GUARDAR OBSERVACIONES POR SECCIÓN EN COLUMNAS BQ-CT (69-98)
    var lastRow = sheet.getLastRow();
    agregarFotosSeccion(obj, lastRow, folder);

    // ✅ GUARDAR FIRMA DEL RESPONSABLE EN COLUMNA CU (99)
    if (obj.firmaData) {
      try {
        var firmaImageData = Utilities.base64Decode(obj.firmaData.split(',')[1]);
        var firmaBlob = Utilities.newBlob(firmaImageData, MimeType.PNG, obj.ad3 + "_firma.png");
        var firmaFile = folder.createFile(firmaBlob);
        var firmaFileId = firmaFile.getId();
        var firmaUrl = "https://lh5.googleusercontent.com/d/" + firmaFileId;
        sheet.getRange(lastRow, 99).setValue(firmaUrl);  // col CU
      } catch (firmaError) {
        Logger.log("Error guardando firma: " + firmaError.message);
      }
    }

    // ✅ GUARDAR FIRMA DEL SUPERVISOR EN COLUMNA CV (100)
    if (obj.firmaSupervisor) {
      try {
        sheet.getRange(lastRow, 100).setValue(obj.firmaSupervisor);  // col CV = URL del perfil
      } catch (firmaSupError) {
        Logger.log("Error guardando firma supervisor: " + firmaSupError.message);
      }
    }

    setFormula();

    if (obj.checked.includes("No")) {
      sendChecklistEmail(obj, newValue, imageUrl, status);
    }

    return { 
      columnAValue: newValue,
      success: true,
      clearForm: true
    };

  } catch (error) {
    Logger.log("Error en saveDataCheck: " + error.message);
    return {
      success: false,
      error: error.message,
      clearForm: false
    };
  }
}

// =============================================
// HELPER: Parsear celda codificada "num::value~~num::value"
// Retorna [{num, value}]  — compatible con formato antiguo (URL sin '::')
// =============================================
function parseObsCell(cellValue) {
  if (!cellValue) return [];
  var s = String(cellValue).trim();
  if (!s) return [];
  if (!s.includes('::')) return [{ num: '', value: s }];  // formato antiguo
  return s.split('~~').filter(Boolean).map(function(part) {
    var si = part.indexOf('::');
    if (si === -1) return { num: '', value: part };
    return { num: part.substring(0, si), value: part.substring(si + 2) };
  });
}

// =============================================
// FUNCIÓN: Guardar fotos/comentarios por ítem en columnas BQ-CT
// Nuevo formato: obs.items = [{num, comment, image}]
// Codificación: "num::url~~num2::url2" por celda
// =============================================
function agregarFotosSeccion(obj, fila, folder) {
  var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');

  var sectionObsList = [];
  if (obj.sectionObs) {
    try {
      sectionObsList = JSON.parse(obj.sectionObs);
    } catch (e) {
      Logger.log("⚠️ Error parseando sectionObs: " + e.message);
      return;
    }
  }

  if (sectionObsList.length === 0) return;

  sectionObsList.forEach(function(obs) {
    var seccion = obs.section + 1;  // 0-based → 1-based
    if (seccion < 1 || seccion > 10) return;

    var colFoto       = 69 + (seccion - 1) * 3;  // BQ,BT,BW...
    var colComentario = 71 + (seccion - 1) * 3;  // BS,BV,BY...

    var fotoPartes     = [];
    var comentPartes   = [];

    var items = obs.items || [];
    items.forEach(function(item) {
      // 📸 Subir foto y guardar URL codificada
      if (item.image && item.image !== '') {
        try {
          var base64Data = item.image.split(',')[1];
          var imageData  = Utilities.base64Decode(base64Data);
          var blob       = Utilities.newBlob(imageData, MimeType.JPEG,
                             obj.ad3 + '_obs_sec' + seccion + '_item' + item.num + '.jpg');
          var file       = folder.createFile(blob);
          var photoUrl   = 'https://lh5.googleusercontent.com/d/' + file.getId();
          fotoPartes.push(item.num + '::' + photoUrl);
          Logger.log('✅ Foto sección ' + seccion + ' ítem ' + item.num + ' guardada');
        } catch (e) {
          Logger.log('❌ Error foto sec ' + seccion + ' ítem ' + item.num + ': ' + e.message);
        }
      }
      // 📝 Guardar comentario codificado
      if (item.comment && item.comment.trim() !== '') {
        comentPartes.push(item.num + '::' + item.comment.trim());
      }
    });

    if (fotoPartes.length > 0)   sheet.getRange(fila, colFoto).setValue(fotoPartes.join('~~'));
    if (comentPartes.length > 0) sheet.getRange(fila, colComentario).setValue(comentPartes.join('~~'));
    // colSubsanacion (colFoto+1) queda vacía — se llenará en seguimiento
  });
}

// ✅ FUNCIÓN: Guardar seguimiento (fotos de levantamiento/subsanación)
function saveDataCheckSeguimiento(objData) {
  try {
    const checkId = objData.checkId;
    const sheet = getCheckSpreadsheet().getSheetByName('B DATOS');

    // Remover filtro activo si existe
    try { var _f = sheet.getFilter(); if (_f) _f.remove(); } catch(e) {}

    const folder = DriveApp.getFolderById(folderimgcheck);
    
    if (!checkId) {
      throw new Error("ID del check no proporcionado");
    }
    
    // Buscar la fila del check
    const data = sheet.getRange("A2:A").getValues();
    let rowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == checkId) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error("Check no encontrado");
    }
    
    // ✅ GUARDAR FOTOS DE SUBSANACIÓN en columnas 70,73,76... (BR,BU,BX...)
    // Nuevo formato: objData['subsanacion-{s}'] = JSON array [{num, b64}]
    // Formato legado: objData['fotoLevantamiento-{N}'] = base64 collage único
    for (let s = 0; s < 10; s++) {
      const seccion = s + 1;
      const newKey  = 'subsanacion-' + s;
      const legKey  = 'fotoLevantamiento-' + seccion;

      if (objData[newKey]) {
        // Nuevo formato: per-item
        try {
          const items = JSON.parse(objData[newKey]);
          const encodedParts = [];
          items.forEach(function(item) {
            try {
              const imageData = Utilities.base64Decode(item.b64.split(',')[1]);
              const blob = Utilities.newBlob(imageData, MimeType.JPEG,
                'sub_' + checkId + '_sec' + seccion + '_item' + item.num + '.jpg');
              const file    = folder.createFile(blob);
              const photoUrl = 'https://lh5.googleusercontent.com/d/' + file.getId();
              encodedParts.push(item.num + '::' + photoUrl);
              Logger.log('✅ Sub foto sec ' + seccion + ' ítem ' + item.num);
            } catch (e) {
              Logger.log('❌ Error sub foto sec ' + seccion + ' ítem ' + item.num + ': ' + e.message);
            }
          });
          if (encodedParts.length > 0) {
            const colSubsanacion = 70 + s * 3;
            sheet.getRange(rowIndex, colSubsanacion).setValue(encodedParts.join('~~'));
          }
        } catch (e) {
          Logger.log('❌ Error parseando subsanacion-' + s + ': ' + e.message);
        }
      } else if (objData[legKey] && objData[legKey] !== '') {
        // Formato legado (collage único por sección)
        try {
          const imageData = Utilities.base64Decode(objData[legKey].split(',')[1]);
          const blob = Utilities.newBlob(imageData, MimeType.PNG,
            'levantamiento_' + checkId + '_seccion' + seccion + '.png');
          const file    = folder.createFile(blob);
          const photoUrl = 'https://lh5.googleusercontent.com/d/' + file.getId();
          const colSubsanacion = 70 + s * 3;
          sheet.getRange(rowIndex, colSubsanacion).setValue(photoUrl);
          Logger.log('✅ Sub foto sec ' + seccion + ' (formato legado)');
        } catch (error) {
          Logger.log('❌ Error sub foto sec ' + seccion + ' legado: ' + error.message);
        }
      }
    }
    
    // ✅ CALCULAR Y ACTUALIZAR ESTADO AUTOMÁTICO
    const estadoActualizado = calcularEstadoCheck(checkId, sheet, rowIndex);
    sheet.getRange(rowIndex, 13).setValue(estadoActualizado);
    
    return {
      success: true,
      checkId: checkId,
      estado: estadoActualizado,
      mensaje: "Seguimiento guardado correctamente"
    };
    
  } catch (error) {
    Logger.log("❌ Error en saveDataCheckSeguimiento: " + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

// ✅ NUEVA FUNCIÓN: Marcar ítem como "En Gestión del Responsable" (solo admin)
function setGestionResponsable(objData) {
  try {
    const checkId  = objData.checkId;
    const sIdx     = parseInt(objData.sIdx);
    const itemNum  = String(objData.itemNum);
    const habilitar = objData.habilitar === true;

    const sheet    = getCheckSpreadsheet().getSheetByName('B DATOS');

    // Remover filtro activo si existe
    try { var _f2 = sheet.getFilter(); if (_f2) _f2.remove(); } catch(e) {}

    const rowIndex0 = getRowIndexByID(checkId); // 0-based (batchUpdate), necesita +1 para sheet.getRange
    if (rowIndex0 == null || rowIndex0 === -1) throw new Error('Check no encontrado: ' + checkId);
    const rowIndex = rowIndex0 + 1; // convertir a 1-based para sheet.getRange

    const colSubsanacion = 70 + sIdx * 3;
    const subRaw  = String(sheet.getRange(rowIndex, colSubsanacion).getValue() || '').trim();
    const subItems = parseObsCell(subRaw);

    // Filtrar el ítem actual y reconstruir la celda
    let filtrados = subItems.filter(function(i) { return i.num !== itemNum; });
    if (habilitar) {
      filtrados.push({ num: itemNum, value: 'EN_GESTION' });
    }

    const newValue = filtrados.map(function(i) { return i.num + '::' + i.value; }).join('~~');
    sheet.getRange(rowIndex, colSubsanacion).setValue(newValue);

    // Recalcular y guardar estado
    const estadoActualizado = calcularEstadoCheck(checkId, sheet, rowIndex);
    sheet.getRange(rowIndex, 13).setValue(estadoActualizado);

    Logger.log('✅ GR ' + (habilitar ? 'habilitado' : 'deshabilitado') + ' para check ' + checkId + ' sec ' + sIdx + ' ítem ' + itemNum);
    return { success: true, estado: estadoActualizado };
  } catch (e) {
    Logger.log('❌ Error en setGestionResponsable: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ✅ Calcular estado automático del check (conteo a nivel de ítem)
function calcularEstadoCheck(checkId, sheet, rowIndex) {
  let totalObservaciones = 0;
  let observacionesLevantadas = 0;

  for (let seccion = 1; seccion <= 10; seccion++) {
    const colFoto        = 69 + (seccion - 1) * 3;
    const colSubsanacion = 70 + (seccion - 1) * 3;

    const fotoRaw = String(sheet.getRange(rowIndex, colFoto).getValue() || '').trim();
    const subRaw  = String(sheet.getRange(rowIndex, colSubsanacion).getValue() || '').trim();

    if (!fotoRaw) continue;

    const obsItems = parseObsCell(fotoRaw);
    const subItems = parseObsCell(subRaw);

    // Construir set de ítems subsanados por num
    const subNums = new Set(subItems.map(function(i) { return i.num; }));

    obsItems.forEach(function(obsItem) {
      if (!obsItem.value) return;
      totalObservaciones++;
      if (subNums.has(obsItem.num)) observacionesLevantadas++;
    });
  }

  let estado;
  if (totalObservaciones === 0)                            estado = 'Conforme';
  else if (observacionesLevantadas === 0)                  estado = 'Abierto';
  else if (observacionesLevantadas < totalObservaciones)   estado = 'En Proceso';
  else                                                      estado = 'Cerrado';

  Logger.log('✅ Estado: ' + estado + ' (' + observacionesLevantadas + '/' + totalObservaciones + ')');
  return estado;
}

function sendChecklistEmail(obj, newValue, imageUrl, status) {
  const ss = getCheckSpreadsheet();
  
  const menuSheet = ss.getSheetByName('MENÚ');
  const recipient = menuSheet.getRange("B24").getValue().trim();
  if (!recipient) return;

  const itemsSheet = ss.getSheetByName('CHECK LIST');
  const lastRow = itemsSheet.getLastRow();
  const itemsData = itemsSheet.getRange(1, 1, lastRow, 3).getValues();

  const equipo = obj.ad3;

  const items = itemsData
    .filter(row => row[1] === equipo)
    .map(row => row[2]);

  const compList = obj.checked.split(',').map(c => c.trim());

  const itemsHtml = compList.map((valor, i) => {
    const item = items[i];
    if (!item) return '';
    if (valor === 'No') return `<div style="color:#d9534f;"><b>X</b> ${i+1}. ${item}</div>`;
    if (valor === 'Si') return `<div style="color:#5cb85c;"><b>✓</b> ${i+1}. ${item}</div>`;
    return `<div style="color:#0275d8;"><b>O</b> ${i+1}. ${item}</div>`;
  }).join('');

  const subject = "⚠️ Alerta: " + equipo + " No conforme";

  const body = `
    <div style="font-family: Arial; max-width:700px; margin:auto; padding:20px; background:#f9f9f9; border:1px solid #ccc; border-radius:8px;">
      <h2 style="color:#d9534f;">🛑 Alerta Check List – No Conforme</h2>
      <p>Estimado equipo,</p>
      <p>Se ha registrado una lista de verificación con observaciones:</p>
      <table style="width:100%; font-size:14px; margin-top:15px;">
        <tr><td><b>ID</b></td><td>${newValue}</td></tr>
        <tr><td><b>Empresa</b></td><td>${obj.ad1}</td></tr>
        <tr><td><b>Equipo</b></td><td>${equipo}</td></tr>
        <tr><td><b>Código/Placa</b></td><td>${obj.ad4}</td></tr>
        <tr><td><b>Área</b></td><td>${obj.ad2}</td></tr>
        <tr><td><b>Supervisor de Operaciones</b></td><td>${obj.adSup || '-'}</td></tr>
        <tr><td><b>Responsable de Área</b></td><td>${obj.ad6}</td></tr>
        <tr><td><b>Proceso</b></td><td>${obj.ad9}</td></tr>
        <tr><td><b>Lugar</b></td><td>${obj.ad5}</td></tr>
        <tr><td><b>Plan de Acción</b></td><td>${obj.ad7}</td></tr>
        <tr><td><b>Fecha</b></td><td>${new Date().toLocaleString()}</td></tr>
        <tr><td><b>Estado</b></td><td style="color:${status === 'Abierto' ? '#d9534f' : '#5cb85c'};"><b>${status}</b></td></tr>
      </table>
      ${ imageUrl ? `<div style="margin-top:20px;"><b>Imagen registrada:</b><br><img src="${imageUrl}" style="max-width:100%; border-radius:4px;"></div>` : '' }
      <div style="margin-top:25px;">
        <h3 style="color:#0275d8;">Lista de Verificación</h3>
        <div style="font-size:14px;">${itemsHtml}</div>
      </div>
      <p style="margin-top:25px;">Revisar y tomar las acciones correspondientes.</p>
      <hr style="margin-top:30px;">
      <p style="font-size:12px; color:#666;">Mensaje generado automáticamente.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: body
  });
}

function setFormula() {
  var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
  var lastRow = sheet.getLastRow();

  var rangeToCopy = sheet.getRange(lastRow-1, 16);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 16));

  var rangeToCopy = sheet.getRange(lastRow-1, 17);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 17));

  var rangeToCopy = sheet.getRange(lastRow-1, 18);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 18));

  var rangeToCopy = sheet.getRange(lastRow-1, 19);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 19));
}

function getDataCheckList(user) {
  const sheet = getCheckSpreadsheet().getSheetByName("HISTORIAL");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 6).getDisplayValues();

  const result = data.filter(r => r[4] === user.ad3 && r[5] === user.ad4);
  return result;
}

function getDropDownarray(cargo) {
  const checkSS  = getCheckSpreadsheet();
  const sheet    = checkSS.getSheetByName("INVENTARIO");
  const lastRow  = sheet.getLastRow();
  const data     = sheet.getRange(1, 1, lastRow, 18).getDisplayValues();

  const filteredData = data.slice(2).map(row =>
    row.slice(1).map(cell => String(cell).trim().replace(/\s+/g, ' '))
  );

  // ── Leer B DATOS → última inspección por (código, lugar) ─────────────────
  // { codigo: { lugar: Date } }
  const ultimaInspPorLugar = {};
  try {
    const bSheet = checkSS.getSheetByName('B DATOS');
    const bLast  = bSheet ? bSheet.getLastRow() : 0;
    if (bLast > 1) {
      bSheet.getRange(2, 1, bLast - 1, 15).getValues().forEach(row => {
        const cod   = String(row[3] || '').trim(); // col D = código
        const lugar = String(row[7] || '').trim(); // col H = lugar
        const fecha = row[9];                      // col J = fecha
        if (cod && fecha instanceof Date) {
          if (!ultimaInspPorLugar[cod]) ultimaInspPorLugar[cod] = {};
          const key = lugar || '__sin_lugar__';
          if (!ultimaInspPorLugar[cod][key] || fecha > ultimaInspPorLugar[cod][key]) {
            ultimaInspPorLugar[cod][key] = new Date(fecha);
          }
        }
      });
    }
  } catch(e) {
    Logger.log('getDropDownarray: error leyendo B DATOS — ' + e.message);
  }

  // ── Hoy a medianoche ────────────────────────────────────────────────────
  const hoy     = new Date();
  const hoyDate = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  // ── Parser inline de frecuencia (evita dependencia cruzada de archivos) ─
  function parseFreq(val) {
    if (val === null || val === undefined || val === '') return -1;
    if (typeof val === 'number') return val;
    const s = String(val).trim().toLowerCase().replace(',', '.');
    const n = parseFloat(s);
    if (!isNaN(n)) return n;
    const m = s.match(/[ck]\/(\d+)/) || s.match(/cada\s+(\d+)/);
    if (m) return parseInt(m[1], 10);
    if (s.includes('diario'))    return 1;
    if (s.includes('semanal'))   return 7;
    if (s.includes('quincenal')) return 15;
    if (s.includes('mensual'))   return 30;
    if (s.includes('bimestral')) return 60;
    if (s.includes('trimestral'))return 90;
    if (s.includes('semestral')) return 180;
    if (s.includes('anual'))     return 365;
    return -1; // desconocido → no filtrar
  }

  const cargoLower   = cargo ? String(cargo).trim().toLowerCase() : '';
  const esSupervisor = cargoLower.includes('supervisor');

  const result = [];
  filteredData.forEach(row => {
    // 1. Retirados → excluir completamente (no aparecen en ningún dropdown)
    if ((row[14] || '').toLowerCase() === 'retirado') return;

    // 2. Filtro por cargo (col F = B-index 4) → excluir completamente
    if (cargoLower && !esSupervisor) {
      const cargoCol = String(row[4] || '').trim().toLowerCase();
      if (cargoCol) {
        const ok = cargoCol.split(',').some(c => {
          const ct = c.trim();
          return ct && cargoLower.includes(ct);
        });
        if (!ok) return;
      }
    }

    // 3. Calcular disponibilidad por período → flag en índice 17
    //    La fila SIEMPRE se incluye (para empresa/área), solo el equipo se oculta si !disponible
    const codigo     = String(row[3]  || '').trim();
    const freq       = parseFreq(row[11]);
    const lugaresStr = String(row[10] || '').trim();
    const allLugares = lugaresStr
      ? lugaresStr.split(',').map(l => l.trim()).filter(l => l)
      : [];

    const insp = ultimaInspPorLugar[codigo] || {};
    let disponible = true;
    let availLugares;

    if (allLugares.length > 0) {
      // Equipo con lugares definidos → filtrar por período
      if (freq === 0) {
        availLugares = allLugares.filter(lug => !insp[lug]);
      } else if (freq > 0) {
        availLugares = allLugares.filter(lug => {
          const ultima = insp[lug];
          if (!ultima) return true;
          const ultimaDate = new Date(ultima.getFullYear(), ultima.getMonth(), ultima.getDate());
          return _estaDisponibleInsp(freq, ultimaDate, hoyDate);
        });
      } else {
        availLugares = allLugares.slice(); // freq=-1 desconocido: todos disponibles
      }
      disponible = availLugares.length > 0; // visible si al menos un lugar pendiente
    } else {
      // Sin lugares definidos → chequeo a nivel equipo con clave __sin_lugar__
      availLugares = [];
      if (freq === 0) {
        disponible = !insp['__sin_lugar__'];
      } else if (freq > 0) {
        const ultima = insp['__sin_lugar__'];
        if (ultima) {
          const ultimaDate = new Date(ultima.getFullYear(), ultima.getMonth(), ultima.getDate());
          if (!_estaDisponibleInsp(freq, ultimaDate, hoyDate)) disponible = false;
        }
      }
      // freq=-1: disponible = true (ya es el valor por defecto)
    }

    const rowConFlag = row.slice();
    rowConFlag[17] = disponible;    // flag de visibilidad para el dropdown equipo
    rowConFlag[18] = availLugares;  // lugares aún pendientes para el dropdown lugar
    result.push(rowConFlag);
  });

  return result;
}

function getAdditionalInfoByValue(value) {
  const sheet = getCheckSpreadsheet().getSheetByName("INVENTARIO");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 14).getDisplayValues();

  for (let i = 2; i < data.length; i++) {
    if (data[i][4] == value) {
      return {
        message1: data[i][7],
        message2: data[i][13]
      };
    }
  }

  return null;
}

function muatData() { 
  var datasheet = getCheckSpreadsheet().getSheetByName("ACTUAL"); 
  var mydata = datasheet.getRange(4,1, datasheet.getLastRow()-3,5).getValues();  
   mydata = mydata.filter(row => row.some(cell => cell !== ''));
  var kolomdata = 0;  
  
 for(var i = 0; i < mydata.length; i++){          
        
        var data = new Date(mydata[i][kolomdata]) ;      

          data.setDate(data.getDate());

        var d = data.getDate();
        var m = data.getMonth() + 1;
        var a = data.getFullYear();
        
        if(d < 10){
           var d = "0" + d;
        }

        if(m < 10){
           var m = "0" + m;
        }
}

 return mydata;
}

function getColumnAValue() {
    var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return { columnAValue: 1, success: true, clearForm: true };
    // Batch: lee columna A completa en una sola llamada (evita N+1 getRange individual)
    var colA = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] !== '') {
        return { columnAValue: colA[i][0] + 1, success: true, clearForm: true };
      }
    }
    return { columnAValue: 1, success: true, clearForm: true };
}

// ══════════════════════════════════════════════════════════════════════════════
// PDF DESDE HTML  –  Sin dependencia de la hoja FORMATO
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Convierte una URL lh5.googleusercontent.com/d/FILE_ID
 * a una data-URI base64 incrustada en el HTML para el PDF.
 */
function convertirUrlParaPDF(url) {
  if (!url || url.trim() === '') return '';
  try {
    var match = url.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
    if (!match) return '';
    var fileId = match[1];
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var mime = blob.getContentType() || 'image/jpeg';
    return 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log('convertirUrlParaPDF error (' + url + '): ' + e.message);
    return '';
  }
}

/**
 * Construye el HTML completo del registro de inspección y lo convierte a PDF.
 * Sube el archivo a Drive y devuelve la URL pública.
 */
function generarPDFdesdeHTML(recordId) {
  var ss = getCheckSpreadsheet();

  // ── 1. Leer fila de B DATOS ──────────────────────────────────────────────
  var bSheet = ss.getSheetByName('B DATOS');
  var bLastRow = bSheet.getLastRow();
  var bData = bSheet.getRange(2, 1, bLastRow - 1, 100).getDisplayValues();
  var fila = null;
  for (var i = 0; i < bData.length; i++) {
    if (String(bData[i][0]).trim() === String(recordId).trim()) {
      fila = bData[i]; break;
    }
  }
  if (!fila) throw new Error('Registro no encontrado: ' + recordId);

  var empresa    = fila[1];
  var equipo     = fila[2];
  var codPlaca   = fila[3];
  var area       = fila[4];
  var proceso    = fila[5];
  var supervisor = fila[6];
  var lugar      = fila[7];
  var planAccion = fila[8];
  var fechaHora  = fila[9];          // getDisplayValues → string con fecha y hora
  var responsable     = fila[11];
  var estado          = fila[12];
  var tipoInspeccion  = fila[13] || 'Planeada';   // col N
  var itemsCSV        = fila[14] || '';
  var firmaUrl        = (fila[98] || '').trim();  // col CU (index 98) - firma trabajador
  var firmaSupUrl     = (fila[99] || '').trim();  // col CV (index 99) - firma supervisor

  // Separar fecha y hora
  var partes   = String(fechaHora).split(' ');
  var fechaStr = partes[0] || fechaHora;
  var horaStr  = partes.slice(1).join(' ') || '';

  // ── Leer datos de empresa desde INFO EMPRESA (lookup por nombre) ──────────
  var ruc          = '';
  var actividadEco = '';
  var domicilioEmp = '';
  var objetivoEmp  = '';
  var sedeEmp      = '';
  var clienteEmp   = '';
  try {
    var infoEmpSh = getSpreadsheetPersonal().getSheetByName('INFO EMPRESA');
    if (infoEmpSh) {
      var infoEmpData = infoEmpSh.getRange(2, 1, infoEmpSh.getLastRow() - 1, 7).getDisplayValues();
      for (var ie = 0; ie < infoEmpData.length; ie++) {
        if (String(infoEmpData[ie][0]).trim().toLowerCase() === String(empresa).trim().toLowerCase()) {
          ruc          = (infoEmpData[ie][1] || '').trim();
          actividadEco = (infoEmpData[ie][2] || '').trim();
          domicilioEmp = (infoEmpData[ie][3] || '').trim();
          objetivoEmp  = (infoEmpData[ie][4] || '').trim();
          sedeEmp      = (infoEmpData[ie][5] || '').trim();
          clienteEmp   = (infoEmpData[ie][6] || '').trim();
          break;
        }
      }
    }
  } catch(e) {}

  // ── Contar personal activo ────────────────────────────────────────────────
  var numTrabajadores = contarPersonalActivo();

  // ── 2. Leer CHECK LIST items para este equipo ────────────────────────────
  var clSheet = ss.getSheetByName('CHECK LIST');
  var clData  = clSheet.getRange(1, 1, clSheet.getLastRow(), 3).getDisplayValues();
  var checkItems = clData.filter(function(row) { return row[1] === equipo; });

  // ── 3. Leer INVENTARIO → frecuencia (col M = índice 12) ─────────────────
  var invSheet = ss.getSheetByName('INVENTARIO');
  var invData  = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 18).getDisplayValues();
  var frecuencia = '';
  for (var i = 0; i < invData.length; i++) {
    if (String(invData[i][4]).trim() === String(codPlaca).trim()) {
      frecuencia = invData[i][12] || ''; break;
    }
  }
  // Objetivo viene de INFO EMPRESA col E
  var objetivo = objetivoEmp || '';

  // ── 4. Parsear compliance + fotos de sección ────────────────────────────
  var compliance = itemsCSV.split(',').map(function(v) { return v.trim(); });

  // Columnas (0-based): foto=68+s*3, sub=69+s*3, comentario=70+s*3  (s=0..9)
  // Soporta nuevo formato "num::url~~num::url" y formato antiguo (URL directa)
  var secFotos = [];
  for (var s = 0; s < 10; s++) {
    var fotoRaw   = (fila[68 + s * 3] || '').trim();
    var subRaw    = (fila[69 + s * 3] || '').trim();
    var comentRaw = (fila[70 + s * 3] || '').trim();

    var fotoItems   = parseObsCell(fotoRaw);
    var subItems    = parseObsCell(subRaw);
    var comentItems = parseObsCell(comentRaw);

    var subByNum    = {};
    subItems.forEach(function(it) { subByNum[it.num] = it.value; });
    var comentByNum = {};
    comentItems.forEach(function(it) { comentByNum[it.num] = it.value; });

    // Unir nums con foto y nums solo con comentario
    var allNums = {};
    fotoItems.forEach(function(it) { if (it.num) allNums[it.num] = true; });
    comentItems.forEach(function(it) { if (it.num) allNums[it.num] = true; });
    var fotoByNum = {};
    fotoItems.forEach(function(it) { if (it.num) fotoByNum[it.num] = it.value; });

    var items = Object.keys(allNums).map(function(num) {
      var rawSub = subByNum[num] || '';
      return {
        num    : num,
        foto   : convertirUrlParaPDF(fotoByNum[num] || ''),
        sub    : rawSub === 'EN_GESTION' ? 'EN_GESTION' : convertirUrlParaPDF(rawSub),
        coment : comentByNum[num] || ''
      };
    });

    secFotos.push({
      items  : items,
      // Legado (num='' = URL directa sin codificación)
      foto   : fotoItems.length === 1 && !fotoItems[0].num ? convertirUrlParaPDF(fotoItems[0].value) : '',
      sub    : subItems.length  === 1 && !subItems[0].num  ? (subItems[0].value === 'EN_GESTION' ? 'EN_GESTION' : convertirUrlParaPDF(subItems[0].value)) : '',
      coment : comentItems.length === 1 && !comentItems[0].num ? comentItems[0].value : ''
    });
  }

  // Firmas
  var firmaB64    = convertirUrlParaPDF(firmaUrl);     // firma trabajador

  // Buscar en PERSONAL: firma supervisor + cargos (col G = index 6)
  var cargoSupervisor  = '';
  var cargoResponsable = '';
  try {
    var persSh = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    var persData = persSh.getRange(2, 1, persSh.getLastRow() - 1, 18).getDisplayValues();
    for (var p = 0; p < persData.length; p++) {
      var nombrePers = String(persData[p][2]).trim().toLowerCase();
      // Firma y cargo del supervisor
      if (supervisor && nombrePers === String(supervisor).trim().toLowerCase()) {
        if (!firmaSupUrl) firmaSupUrl = (persData[p][17] || '').trim();
        cargoSupervisor = (persData[p][6] || '').trim();  // col G = cargo
      }
      // Cargo del responsable del área
      if (responsable && nombrePers === String(responsable).trim().toLowerCase()) {
        cargoResponsable = (persData[p][6] || '').trim();  // col G = cargo
      }
    }
  } catch(e) {}
  var firmaSupB64 = convertirUrlParaPDF(firmaSupUrl);  // firma supervisor

  // ── 5. Armar filas del checklist ─────────────────────────────────────────
  // itemNum es GLOBAL (igual que data-itemnum en Check.html, que no resetea por sección)
  var secCounter   = -1;
  var globalItemNum = 0;
  var checkRows    = [];
  for (var i = 0; i < checkItems.length; i++) {
    var txt    = checkItems[i][2] || '';
    var isTit  = /^\d+\./.test(txt.trim()) && (compliance[i] === 'NA' || compliance[i] === undefined);
    if (isTit) secCounter++;
    var secIdx  = Math.max(secCounter, 0);
    var itemNum = null;
    if (!isTit) { globalItemNum++; itemNum = globalItemNum; }
    checkRows.push({
      type  : isTit ? 'section' : 'item',
      text  : txt,
      valor : compliance[i] || 'NA',
      secIdx: secIdx,
      itemNum: itemNum
    });
  }

  // Índice de observaciones: secItemLookup[secIdx][numGlobal] = {foto,sub,coment}
  var secItemLookup = [];
  for (var s = 0; s < secFotos.length; s++) {
    var byNum = {};
    var sf0 = secFotos[s];
    if (sf0 && sf0.items) sf0.items.forEach(function(it) { byNum[String(it.num)] = it; });
    secItemLookup.push(byNum);
  }

  // ── 6. Estado → color ────────────────────────────────────────────────────
  var estadoColor = { 'Conforme':'#27ae60','Abierto':'#e74c3c',
                      'En Proceso':'#e67e22','Cerrado':'#2980b9' };
  var stColor = estadoColor[estado] || '#555';

  // ── 7. Construir HTML ────────────────────────────────────────────────────
  var css = '<style>'
  + '* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; box-sizing:border-box; margin:0; padding:0 }'
  + 'body { font-family:Arial,sans-serif; font-size:8pt; color:#222; padding:8px }'
  + 'table { border-collapse:collapse; width:100%; margin-bottom:4px }'
  + 'td,th { border:1px solid #999; padding:3px 5px; vertical-align:middle }'
  + '.lbl { background:#d9d9d9 !important; font-weight:bold; font-size:7pt }'
  + '.val { font-size:8pt }'
  + '.sec-head { background:#bfc9d4 !important; font-weight:bold; text-align:center; font-size:8pt; padding:4px }'
  + '.ev-row td { background:#f9f9f9 !important }'
  + '.si { color:#1a7a1a !important; font-weight:bold }'
  + '.no { color:#cc0000 !important; font-weight:bold }'
  + '.na { color:#777 !important }'
  + '.foto-cell { text-align:center; width:90px }'
  + '.foto-cell img { width:80px; height:62px; object-fit:cover; border:1px solid #ccc }'
  + '.ck-head { background:#2c5aa0 !important; color:#fff !important; font-weight:bold; font-size:7.5pt; text-align:center }'
  + '@media print { @page { size:A4 landscape; margin:8mm } tr { page-break-inside:avoid } }'
  + '</style>';


  // ── HEADER INSTITUCIONAL ─────────────────────────────────────────────────
  var hdr = '<table><tr>' +
    '<td rowspan="2" style="width:14%;text-align:center;font-size:20pt;font-weight:bold;color:#e2001a;letter-spacing:-1px;border:1px solid #666">Adecco</td>' +
    '<td colspan="4" style="text-align:center;font-weight:bold;font-size:9pt;background:#d9d9d9;border:1px solid #666">' +
      'SISTEMA DE GESTIÓN DE SEGURIDAD, SALUD OCUPACIONAL Y MEDIO AMBIENTE</td>' +
    '<td rowspan="2" style="width:18%;font-size:7pt;line-height:1.7;border:1px solid #666;padding:4px">' +
      '<b>Código:</b> SOOMA - FR002<br><b>Revisión:</b> 18/03/2025<br><b>Versión:</b> V03</td>' +
    '</tr><tr>' +
    '<td colspan="4" style="text-align:center;font-weight:bold;font-size:11pt;border:1px solid #666">' +
      'REGISTRO DE INSPECCIÓN / VERIFICACIÓN</td>' +
    '</tr></table>';

  // ── DATOS GENERALES (datos desde INFO EMPRESA con fallback) ─────────────
  var esPlaneada = String(tipoInspeccion).trim().toLowerCase() !== 'no planeada';
  var domicilioPDF = domicilioEmp || lugar;
  var sedePDF      = sedeEmp || lugar;
  var clientePDF   = clienteEmp || area;

  var datos = '<table>' +
    // Fila 1: labels
    '<tr>' +
      '<td class="lbl" style="width:20%">RAZÓN SOCIAL</td>' +
      '<td class="lbl" style="width:11%">RUC</td>' +
      '<td class="lbl" style="width:27%" colspan="2">DOMICILIO (Dirección, distrito, departamento, provincia)</td>' +
      '<td class="lbl" style="width:27%">ACTIVIDAD ECONÓMICA</td>' +
      '<td class="lbl" style="width:15%">Nº DE TRABAJADORES</td>' +
    '</tr>' +
    // Fila 2: valores
    '<tr>' +
      '<td class="val">' + empresa + '</td>' +
      '<td class="val">' + (ruc || '') + '</td>' +
      '<td class="val" colspan="2">' + domicilioPDF + '</td>' +
      '<td class="val" style="font-size:7pt">' + (actividadEco || '') + '</td>' +
      '<td class="val" style="text-align:center;font-weight:bold">' + (numTrabajadores !== '' ? numTrabajadores : '') + '</td>' +
    '</tr>' +
    // Fila 3: Cliente + ID
    '<tr>' +
      '<td class="lbl">CLIENTE</td>' +
      '<td class="val" colspan="3">' + clientePDF + '</td>' +
      '<td class="lbl" style="background:#ffe066">ID</td>' +
      '<td class="val" style="background:#ffe066;font-weight:bold;text-align:center">' + recordId + '</td>' +
    '</tr>' +
    // Fila 4: Sede (sola en su fila)
    '<tr>' +
      '<td class="lbl">SEDE</td>' +
      '<td class="val" colspan="5">' + sedePDF + '</td>' +
    '</tr>' +
    // Fila 5: Responsable de la Inspección + Cargo
    '<tr>' +
      '<td class="lbl">RESPONSABLE DE LA INSPECCIÓN</td>' +
      '<td class="val" colspan="2">' + supervisor + '</td>' +
      '<td class="lbl">CARGO DEL INSPECTOR</td>' +
      '<td class="val" colspan="2">' + (cargoSupervisor || '').toUpperCase() + '</td>' +
    '</tr>' +
    // Fila 6: Área Inspeccionada + Equipo / Insumos / Inspección
    '<tr>' +
      '<td class="lbl">ÁREA INSPECCIONADA</td>' +
      '<td class="val" colspan="2">' + (lugar || area || '') + '</td>' +
      '<td class="lbl">EQUIPO / INSUMOS / INSPECCIÓN</td>' +
      '<td class="val" colspan="2">' + (equipo || '') + '</td>' +
    '</tr>' +
    // Fila 7: Responsable del Área + Cargo del Responsable
    '<tr>' +
      '<td class="lbl">RESPONSABLE DEL ÁREA</td>' +
      '<td class="val" colspan="2">' + responsable + '</td>' +
      '<td class="lbl">CARGO</td>' +
      '<td class="val" colspan="2">' + (cargoResponsable || '').toUpperCase() + '</td>' +
    '</tr>' +
    // Fila 8: Tipo de Inspección + Fecha
    '<tr>' +
      '<td class="lbl" colspan="2">TIPO DE INSPECCIÓN</td>' +
      '<td class="val" style="font-size:7.5pt" colspan="2">' +
        'PLANEADA (' + (esPlaneada ? 'X' : '&nbsp;') + ')&nbsp;&nbsp;&nbsp;' +
        'NO PLANEADA (' + (!esPlaneada ? 'X' : '&nbsp;') + ')&nbsp;&nbsp;&nbsp;' +
        'OTROS (Especificar)' +
      '</td>' +
      '<td class="lbl">FECHA DE INSPECCIÓN</td>' +
      '<td class="val">' + fechaStr + '</td>' +
    '</tr>' +
    // Fila 9: Frecuencia + Hora
    '<tr>' +
      '<td class="lbl" colspan="2">FRECUENCIA DE INSPECCIÓN</td>' +
      '<td class="val" colspan="2">' + (frecuencia || '—') + '</td>' +
      '<td class="lbl">HORA DE INSPECCIÓN</td>' +
      '<td class="val">' + horaStr + '</td>' +
    '</tr>' +
    '</table>';

  // OBJETIVO
  var objHtml = '<table>' +
    '<tr><td class="lbl" style="text-align:center">OBJETIVO DE LA INSPECCIÓN</td></tr>' +
    '<tr><td class="val" style="padding:5px 8px;line-height:1.5">' + (objetivo || '—') + '</td></tr>' +
    '</table>';

  // CHECKLIST TABLE
  var ckHtml = '<table>' +
    '<tr>' +
      '<th class="ck-head" style="width:45%">Ítem Verificado</th>' +
      '<th class="ck-head" style="width:12%">Cumplimiento</th>' +
      '<th class="ck-head" style="width:13%">Evidencia</th>' +
      '<th class="ck-head" style="width:13%">Subsanación</th>' +
      '<th class="ck-head" style="width:17%">Comentario</th>' +
    '</tr>';

  for (var r = 0; r < checkRows.length; r++) {
    var row = checkRows[r];
    if (row.type === 'section') {
      ckHtml += '<tr><td colspan="5" class="sec-head">' + row.text + '</td></tr>';
    } else {
      var v = row.valor;
      var vClass = v === 'Si' ? 'si' : (v === 'No' ? 'no' : 'na');
      var vLabel = v === 'Si' ? 'Sí' : (v === 'No' ? 'No' : 'NA');

      // Observación inline para ítems "No"
      var obsItem = null;
      if (v === 'No' && row.itemNum !== null) {
        var lookup = secItemLookup[row.secIdx] || {};
        obsItem = lookup[String(row.itemNum)] || null;
        // Fallback: legacy (sección sin items individuales)
        if (!obsItem) {
          var sfLeg = secFotos[row.secIdx] || {};
          if (!(sfLeg.items && sfLeg.items.length > 0) && (sfLeg.foto || sfLeg.sub || sfLeg.coment)) {
            obsItem = { foto: sfLeg.foto, sub: sfLeg.sub, coment: sfLeg.coment };
          }
        }
      }

      var evCell  = (obsItem && obsItem.foto) ? '<td class="foto-cell"><img src="' + obsItem.foto + '"></td>' : '<td></td>';
      var subCell;
      if (obsItem && obsItem.sub === 'EN_GESTION') {
        subCell = '<td style="text-align:center;font-size:7pt;color:#6c757d;font-weight:bold;vertical-align:middle;padding:4px;">En Gesti\u00f3n</td>';
      } else if (obsItem && obsItem.sub) {
        subCell = '<td class="foto-cell"><img src="' + obsItem.sub + '"></td>';
      } else {
        subCell = '<td></td>';
      }
      var comCell = (obsItem && obsItem.coment) ? '<td style="font-size:7pt">' + obsItem.coment + '</td>'        : '<td></td>';

      ckHtml += '<tr>' +
        '<td style="font-size:7.5pt">' + row.text + '</td>' +
        '<td style="text-align:center" class="' + vClass + '">' + vLabel + '</td>' +
        evCell + subCell + comCell +
        '</tr>';
    }
  }
  ckHtml += '</table>';

  // PLAN DE ACCIÓN + ESTADO
  var planHtml = '<table><tr>' +
    '<td class="lbl" style="width:15%">PLAN DE ACCIÓN</td>' +
    '<td class="val" style="width:65%">' + (planAccion || '—') + '</td>' +
    '<td class="lbl" style="width:10%;text-align:center">ESTADO</td>' +
    '<td style="width:10%;text-align:center;background:' + stColor + ';color:#fff;font-weight:bold">' + estado + '</td>' +
    '</tr></table>';

  // SECCIÓN DE FIRMAS (trabajador + supervisor)
  var firmasTrabajadorImg = firmaB64
    ? '<img src="' + firmaB64 + '" style="max-height:60px;max-width:140px">'
    : '<span style="color:#aaa;font-size:7pt">Sin firma</span>';
  var firmasSupervisorImg = firmaSupB64
    ? '<img src="' + firmaSupB64 + '" style="max-height:60px;max-width:140px">'
    : '<span style="color:#aaa;font-size:7pt">Sin firma</span>';

  var firmasHtml = '<table style="margin-top:8px">' +
    '<tr>' +
      '<td style="width:50%;text-align:center;border:1px solid #999;padding:10px;vertical-align:bottom;height:80px">' +
        firmasTrabajadorImg +
      '</td>' +
      '<td style="width:50%;text-align:center;border:1px solid #999;padding:10px;vertical-align:bottom;height:80px">' +
        firmasSupervisorImg +
      '</td>' +
    '</tr>' +
    '<tr>' +
      '<td style="text-align:center;border:1px solid #999;padding:4px;font-weight:bold;font-size:8pt">' +
        (responsable || '—') +
      '</td>' +
      '<td style="text-align:center;border:1px solid #999;padding:4px;font-weight:bold;font-size:8pt">' +
        (supervisor || '—') +
      '</td>' +
    '</tr>' +
    '<tr>' +
      '<td style="text-align:center;border:1px solid #999;padding:3px;font-size:7pt;background:#d9d9d9">Responsable del Área</td>' +
      '<td style="text-align:center;border:1px solid #999;padding:3px;font-size:7pt;background:#d9d9d9">Supervisor de Operaciones</td>' +
    '</tr>' +
    '</table>';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' + css + '</head><body>' +
    hdr + datos + objHtml + ckHtml + planHtml + firmasHtml +
    '</body></html>';

  // ── 8. Convertir a PDF y subir ───────────────────────────────────────────
  var _sanear = function(s) {
    return String(s || '').replace(/[^\w\sáéíóúÁÉÍÓÚñÑüÜ-]/g, '').trim().replace(/\s+/g, '_');
  };
  var _equiNombre = _sanear(equipo || area || 'Inspeccion');
  var _lugarNombre = _sanear(lugar);
  var _fecha = String(fechaStr).replace(/\//g, '-').replace(/\s/g, '');
  var _partes = ['CHECK_LIST', _equiNombre, _lugarNombre, _fecha].filter(Boolean);
  var pdfBlob = Utilities.newBlob(html, 'text/html', 'reporte.html')
                         .getAs(MimeType.PDF)
                         .setName(_partes.join('_') + '.pdf');
  var folder  = DriveApp.getFolderById(folderpdfcheck);
  var file    = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ── getPdfUrl actualizado → usa generarPDFdesdeHTML ─────────────────────────
function getPdfUrl(columnAValue) {
  return generarPDFdesdeHTML(columnAValue);
}

function setIDAndGetLinks(recordId) {
    var sheet = getCheckSpreadsheet().getSheetByName('FORMATO');
    sheet.getRange('D5').setValue(recordId);
    
    SpreadsheetApp.flush();
    Utilities.sleep(50);

    var sheetId = sheet.getSheetId();
    var url = getCheckSpreadsheet().getUrl().replace(/edit$/, '');
    
    var exportPdfUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:D55';
    var token = ScriptApp.getOAuthToken();
    var responsePdf = UrlFetchApp.fetch(exportPdfUrl, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    var blobPdf = responsePdf.getBlob().setName('PDF_RAC_' + recordId + '.pdf');
    var folder = DriveApp.getFolderById(folderpdfcheck);
    var filePdf = folder.createFile(blobPdf);
    filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var pdfUrl = filePdf.getUrl();
    return {
        pdfUrl: pdfUrl
    };
}

function generarPDF(recordId) {
  var sheet = getCheckSpreadsheet().getSheetByName('FORMATO');
  sheet.getRange('D5').setValue(recordId);
  
  SpreadsheetApp.flush();
  Utilities.sleep(50);

  var sheetId = sheet.getSheetId();
  var url = getCheckSpreadsheet().getUrl().replace(/edit$/, '');
  
  var exportpdflink = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:D55';
  var token = ScriptApp.getOAuthToken();
  var responsePdf = UrlFetchApp.fetch(exportpdflink, {
      headers: {
          'Authorization': 'Bearer ' + token
      }
  });
  var blobPdf = responsePdf.getBlob().setName('PDF_RAC_' + recordId + '.pdf');
  var folder = DriveApp.getFolderById(folderpdfcheck);
  var filePdf = folder.createFile(blobPdf);
  filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var pdflink = filePdf.getUrl();
  return {
      pdflink: pdflink
  };
}

function getData2() {
  const sheet = getCheckSpreadsheet().getSheetByName("B DATOS");
  const lastRow = sheet.getLastRow();
  // ✅ Hasta columna 98 (CT)
  const data = sheet.getRange(1, 1, lastRow, 98).getDisplayValues(); 
  
  console.log(data);
  return data;
}

function getItemsData() {
  const sheet = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 3).getDisplayValues();
  return data;
}

function uploadImageToDrive(fileData, fileName) {
  var folder = DriveApp.getFolderById(folderimgcheck);
  var blob = Utilities.newBlob(Utilities.base64Decode(fileData.split(',')[1]), 'image/png', fileName);
  var file = folder.createFile(blob);
  var fileUrl = "https://lh5.googleusercontent.com/d/" + file.getId();
  return fileUrl;
}

function updateData(updatedData) {
  try {
    const sheet = getCheckSpreadsheet().getSheetByName("B DATOS");
    if (!sheet) throw new Error('La hoja "B DATOS" no existe.');

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const idToFind = updatedData[0];

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == idToFind) {
        const targetRow = i + 2;
        const updateValues = [updatedData.slice(1, 15)];
        sheet.getRange(targetRow, 2, 1, updateValues[0].length).setValues(updateValues);
        break;
      }
    }
  } catch (error) {
    Logger.log("Error en updateData: " + error.message);
    return {
      success: false,
      error: error.message,
      clearForm: false
    };
  }
}

function obtenerDatosChecklist() {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = hoja.getLastRow();
  if (lastRow < 1) return { headersCheck: [], filas: [] };

  const datos = hoja.getRange(1, 1, lastRow, 3).getValues();
  const datosComoTexto = datos.map(fila => fila.map(celda => String(celda || "")));

  return {
    headersCheck: datosComoTexto[0],
    filas: datosComoTexto.slice(1)
  };
}

function obtenerOpcionesInventario() {
  const hoja = getSpreadsheetPersonal().getSheetByName("LISTAS");
  const valores = hoja.getRange("M2:M" + hoja.getLastRow()).getValues().flat();
  return [...new Set(valores.filter(v => v))];
}

function getResponsablesPersonal() {
  const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  // Lee cols C-L (10 cols): C=nombre(0), …, L=estado activo(9)
  const data = hoja.getRange(2, 3, lastRow - 1, 10).getValues();
  return [...new Set(
    data
      .filter(r => { const e = String(r[9] || '').trim().toUpperCase(); return e === 'SI' || e === 'ACTIVO'; })
      .map(r => String(r[0] || '').trim())
      .filter(v => v)
  )].sort();
}

function getSupervisoresOperaciones() {
  const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  // Lee cols C-L (10 cols): C=nombre(0), D(1), E(2), F(3), G=cargo(4), …, L=estado activo(9)
  const data = hoja.getRange(2, 3, lastRow - 1, 10).getValues();
  return [...new Set(
    data
      .filter(r => {
        const estado = String(r[9] || '').trim().toUpperCase();
        const cargo  = String(r[4] || '').trim().toLowerCase();
        return (estado === 'SI' || estado === 'ACTIVO') && cargo.includes('supervisor');
      })
      .map(r => String(r[0] || '').trim())
      .filter(v => v)
  )].sort();
}

function contarPersonalActivo() {
  try {
    var hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return 0;
    var data = hoja.getRange(2, 3, lastRow - 1, 1).getValues(); // col C = nombres
    return data.filter(function(r) { return String(r[0] || '').trim() !== ''; }).length;
  } catch(e) { return ''; }
}

function agregarChecklist(data) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  if (!data) return;

  const varias = Array.isArray(data[0]) && data.length > 0;
  const filas = varias ? data : [data];

  const lastRow = hoja.getLastRow();
  let maxId = 0;
  if (lastRow >= 2) {
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .map(v => Number(v))
      .filter(n => !isNaN(n) && isFinite(n));
    if (ids.length) maxId = Math.max.apply(null, ids);
  }

  let nextId = maxId + 1;
  const filasConId = filas.map(row => {
    const r = row.slice();
    if (r.length === 0) {
      return null;
    }
    if (r[0] === undefined || r[0] === null || r[0] === "" || (typeof r[0] === "number" && isNaN(r[0]))) {
      r[0] = nextId++;
    } else {
      const maybeNum = Number(r[0]);
      if (!isNaN(maybeNum)) {
        r[0] = maybeNum;
        if (maybeNum >= nextId) nextId = maybeNum + 1;
      }
    }
    if (r.length < 3) {
      while (r.length < 3) r.push("");
    }
    return r;
  }).filter(Boolean);

  if (filasConId.length === 0) return;

  const startRow = hoja.getLastRow() + 1;
  hoja.getRange(startRow, 1, filasConId.length, filasConId[0].length).setValues(filasConId);
  return { success: true, clearForm: true };
}

function eliminarChecklist(rowIndex) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  hoja.deleteRow(rowIndex + 2);
  return { success: true, clearForm: true };
}

function obtenerItemsPorEquipo(equipo) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  const datos = hoja.getRange(2, 1, lastRow - 1, 3).getValues();

  return datos
    .filter(r => String(r[1]).trim() === String(equipo).trim())
    .map(r => ({ id: Number(r[0]) || 0, item: String(r[2] || "").trim() }));
}

function actualizarChecklistPorEquipo(equipo, itemsJson) {
  if (!equipo) throw new Error("Equipo no especificado.");

  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const nuevosItems = JSON.parse(itemsJson || "[]");
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return;

  const datos = hoja.getRange(2, 1, lastRow - 1, 3).getValues()
    .map((r, i) => ({
      id: Number(r[0]) || 0,
      eq: String(r[1]).trim(),
      item: String(r[2]).trim(),
      fila: i + 2
    }))
    .filter(r => r.eq === String(equipo).trim());

  const existentesPorID = new Map(datos.map(r => [r.id, r]));
  const usadosIDs = new Set();
  const eliminaciones = [];

  nuevosItems.forEach(n => {
    if (n.id && existentesPorID.has(n.id)) {
      const filaExistente = existentesPorID.get(n.id);
      usadosIDs.add(n.id);
      if (n.item !== filaExistente.item) {
        hoja.getRange(filaExistente.fila, 3).setValue(n.item);
      }
    }
  });

  datos.forEach(r => {
    if (!usadosIDs.has(r.id)) eliminaciones.push(r.fila);
  });
  eliminaciones.sort((a, b) => b - a).forEach(f => hoja.deleteRow(f));

  const nuevos = nuevosItems.filter(n => !n.id || !existentesPorID.has(n.id));
  if (nuevos.length > 0) {
    const nextId = (hoja.getRange(hoja.getLastRow(), 1).getValue() || 0) + 1;
    const registros = nuevos.map((n, i) => [nextId + i, equipo, n.item]);
    hoja.getRange(hoja.getLastRow() + 1, 1, registros.length, 3).setValues(registros);
  }

  return {
    status: "ok",
    modificados: usadosIDs.size,
    agregados: nuevos.length,
    eliminados: eliminaciones.length,
    success: true,
    clearForm: true
  };
}

function obtenerInventarioServerSide(offset = 0, limit = 30, terminoBusqueda = "") {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  const headers = hoja.getRange(2, 1, 1, hoja.getLastColumn()).getValues()[0];
  const ultimaFila = hoja.getLastRow();

  const totalFilas = ultimaFila - 2;
  const filasLeidas = hoja.getRange(3, 1, totalFilas, hoja.getLastColumn()).getValues();
  const fechaCols = [13, 16, 17];

  let datos = filasLeidas.reverse();

  if (terminoBusqueda && terminoBusqueda.trim() !== "") {
    const filtro = terminoBusqueda.toLowerCase();
    datos = datos.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(filtro))
    );
  }

  const paginados = datos.slice(offset, offset + limit);

  const formateados = paginados.map(fila =>
    fila.map((celda, i) => {
      if (fechaCols.includes(i) && celda instanceof Date) {
        return Utilities.formatDate(celda, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return celda;
    })
  );

  return {
    headersInvet: headers,
    filas: formateados,
    total: datos.length
  };
}

function agregarInventario(data) {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  data = parsearFechas(data);
  data[0] = hoja.getLastRow() - 1;
  hoja.appendRow(data);
  return { success: true, clearForm: true };
}

function actualizarInventario(data) {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  data = parsearFechas(data);
  const fila = parseInt(data[0], 10) + 2;
  hoja.getRange(fila, 1, 1, data.length).setValues([data]);
  return { success: true, clearForm: true };
}

function eliminarInventarioPorNum(num) {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  const fila = parseInt(num, 10) + 2;
  hoja.deleteRow(fila);
  
  const datos = hoja.getRange(3, 1, hoja.getLastRow() - 2, 1).getValues();
  datos.forEach((_, i) => {
    hoja.getRange(i + 3, 1).setValue(i + 1);
  });
  return { success: true, clearForm: true };
}

function parsearFechas(data) {
  const fechaIndices = [13, 16, 17];
  fechaIndices.forEach(i => {
    if (data[i]) {
      const partes = data[i].split("/");
      if (partes.length === 3) {
        data[i] = new Date(`${partes[2]}-${partes[1]}-${partes[0]}`);
      }
    }
  });
  return data;
}

function obtenerEquiposSinChecklist() {
  const hojaInventario = getCheckSpreadsheet().getSheetByName("INVENTARIO");
  const hojaChecklist = getCheckSpreadsheet().getSheetByName("CHECK LIST");

  const limpiarTexto = (texto) => String(texto).trim().toLowerCase();

  const valoresInventario = hojaInventario.getRange("D3:D" + hojaInventario.getLastRow()).getValues().flat();
  const inventarioLimpio = valoresInventario
    .map(limpiarTexto)
    .filter(e => e !== "");

  const valoresChecklist = hojaChecklist.getRange("B2:B" + hojaChecklist.getLastRow()).getValues().flat();
  const checklistLimpio = valoresChecklist
    .map(limpiarTexto)
    .filter(e => e !== "");

  const faltantes = valoresInventario.filter((equipo) => {
    const equipoLimpio = limpiarTexto(equipo);
    return equipoLimpio && !checklistLimpio.includes(equipoLimpio);
  });

  const faltantesUnicos = [...new Set(faltantes.map(limpiarTexto))];

  if (faltantesUnicos.length > 0) {
  }

  return faltantesUnicos;
}

function generarItemsConGemini(base64DataUrl, textoBase, numItems) {

  const hayArchivo = !!base64DataUrl;
  const hayTexto = !!textoBase && textoBase.trim() !== "";

  let prompt = `
Eres un experto en seguridad industrial. Genera una lista de ${
    numItems ? numItems + " " : ""
  }ítems de verificación para un checklist técnico de equipos.

Cada ítem debe tener formato de pregunta breve y clara, con foco en cumplimiento:
Ejemplos:
- "¿Manómetro: Con presión adecuada?"
- "¿Manguera: En buen estado sin daños físicos?"
- "¿Etiqueta de inspección vigente?"
- "¿Válvula principal: Sin fugas visibles?"

`;

  if (hayTexto && hayArchivo) {
    prompt += `
Analiza el documento adjunto y el siguiente texto descriptivo:
"""${textoBase}"""
`;
  } else if (hayArchivo) {
    prompt += `
Analiza el documento adjunto y genera los ítems relevantes del checklist.
`;
  } else if (hayTexto) {
    prompt += `
Basado en este texto o descripción del equipo:
"""${textoBase}"""
`;
  }

  prompt += `
Devuelve un JSON puro con este formato exacto:
[
  {"item": "¿Texto del ítem 1?"},
  {"item": "¿Texto del ítem 2?"},
  ...
]
Solo responde con el JSON, sin explicaciones ni comentarios adicionales.
`;

  const parts = [{ text: prompt }];
  if (hayArchivo) {
    const mimeType = base64DataUrl.match(/^data:(.*?);/)[1];
    const base64 = base64DataUrl.split(",")[1];
    parts.push({ inlineData: { mimeType, data: base64 } });
  }

  const payload = { contents: [{ parts }] };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const texto = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

  const inicio = texto.indexOf("[");
  const fin = texto.lastIndexOf("]");
  if (inicio === -1 || fin === -1) throw new Error("Respuesta inválida de Gemini.");

  const limpio = texto.substring(inicio, fin + 1);
  const arr = JSON.parse(limpio);

  return JSON.stringify(arr.map(p => ({ item: p.item || p.pregunta || "" })));
}

function obtenerItemsPorEquipoConSeparadores(equipo) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  
  const datos = hoja.getRange(2, 1, lastRow - 1, 3).getValues();
  
  return datos
    .filter(r => String(r[1]).trim() === String(equipo).trim())
    .map(r => {
      const item = String(r[2] || "").trim();
      
      const esNumeroConPunto = /^\d+\.\s*.+$/.test(item);
      const tieneInterrogacion = item.includes('¿') || item.includes('?');
      
      const esTitulo = esNumeroConPunto && !tieneInterrogacion;
      
      const esSubtitulo = !esTitulo;
      
      return {
        id: Number(r[0]) || 0,
        item: item,
        esSeparador: esTitulo,
        seleccionable: esSubtitulo,
        tipo: esTitulo ? 'titulo' : 'subtitulo'
      };
    });
}
// ============================================================
//  HELPERS DE PERÍODOS CALENDARIO — usados por disponibilidad,
//  cumplimiento y alertas de inspecciones.
//  Reglas:
//    semanal  → semana calendario lunes–domingo
//    mensual  → mes calendario 1° al último día
//    bimestral→ bimestre calendario
//    trimestral → trimestre calendario (ene-mar, abr-jun, jul-sep, oct-dic)
//    semestral  → semestre calendario (ene-jun, jul-dic)
//    anual      → año calendario
//    otro (ej. cada 4 días) → rolling (freq días completos, día de inspección no se cuenta)
// ============================================================

/** Retorna el tipo de período calendario según freq (número de días),
 *  o null si es rolling (cada N días libre). */
function _freqEsCalendario(freq) {
  if (freq === 7)   return 'semana';
  if (freq === 30)  return 'mes';
  if (freq === 60)  return 'bimestre';
  if (freq === 90)  return 'trimestre';
  if (freq === 180) return 'semestre';
  if (freq === 365) return 'anio';
  return null;
}

/** Lunes de la semana a la que pertenece la fecha d. */
function _lunesDeSemana(d) {
  var day  = d.getDay(); // 0=Dom,1=Lun,...,6=Sab
  var diff = (day === 0) ? -6 : (1 - day);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff);
}

/** true si fecha1 y fecha2 caen en el mismo período calendario. */
function _mismoPeriodoCalInsp(periodo, fecha1, fecha2) {
  switch (periodo) {
    case 'semana':
      return _lunesDeSemana(fecha1).getTime() === _lunesDeSemana(fecha2).getTime();
    case 'mes':
      return fecha1.getMonth()  === fecha2.getMonth()  &&
             fecha1.getFullYear() === fecha2.getFullYear();
    case 'bimestre':
      return Math.floor(fecha1.getMonth() / 2) === Math.floor(fecha2.getMonth() / 2) &&
             fecha1.getFullYear() === fecha2.getFullYear();
    case 'trimestre':
      return Math.floor(fecha1.getMonth() / 3) === Math.floor(fecha2.getMonth() / 3) &&
             fecha1.getFullYear() === fecha2.getFullYear();
    case 'semestre':
      return Math.floor(fecha1.getMonth() / 6) === Math.floor(fecha2.getMonth() / 6) &&
             fecha1.getFullYear() === fecha2.getFullYear();
    case 'anio':
      return fecha1.getFullYear() === fecha2.getFullYear();
    default:
      return false;
  }
}

/** Inicio del período actual (para calcular días vencido). */
function _inicioPeriodoActualInsp(periodo, hoy) {
  switch (periodo) {
    case 'semana':
      return _lunesDeSemana(hoy);
    case 'mes':
      return new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    case 'bimestre':
      return new Date(hoy.getFullYear(), Math.floor(hoy.getMonth() / 2) * 2, 1);
    case 'trimestre':
      return new Date(hoy.getFullYear(), Math.floor(hoy.getMonth() / 3) * 3, 1);
    case 'semestre':
      return new Date(hoy.getFullYear(), Math.floor(hoy.getMonth() / 6) * 6, 1);
    case 'anio':
      return new Date(hoy.getFullYear(), 0, 1);
    default:
      return hoy;
  }
}

/**
 * Disponibilidad de una inspección:
 *   - Si es período calendario → disponible si ultimaDate NO cae en el período actual de hoyDate.
 *   - Si es rolling (cada N días) → disponible cuando hoyDate >= ultimaDate + freq.
 *     El día de inspección NO se cuenta, se avanza freq días completos.
 * @param {number} freq - diasFrecuencia parseado
 * @param {Date}   ultimaDate - fecha normalizada (sin hora) de la última inspección
 * @param {Date}   hoyDate    - fecha normalizada (sin hora) de hoy
 * @returns {boolean}
 */
function _estaDisponibleInsp(freq, ultimaDate, hoyDate) {
  var periodo = _freqEsCalendario(freq);
  if (periodo) {
    return !_mismoPeriodoCalInsp(periodo, ultimaDate, hoyDate);
  }
  // Rolling: disponible cuando hoy >= ultima + freq días completos
  var nextDue = new Date(ultimaDate);
  nextDue.setDate(nextDue.getDate() + freq);
  return hoyDate >= nextDue;
}

/**
 * Cumplimiento al guardar un checklist:
 *   - Si ya se hizo en el período actual → "Cumple"
 *   - Si no se hizo → "No cumple"
 * @param {number} diasFrecuencia
 * @param {Date|null} ultimaFecha - última inspección (puede ser null)
 * @param {Date} hoy
 * @returns {string}
 */
function _calcularCumplimientoInsp(diasFrecuencia, ultimaFecha, hoy) {
  if (diasFrecuencia === 0) return 'Única vez';
  if (!ultimaFecha) return 'Primera vez';

  var periodo = _freqEsCalendario(diasFrecuencia);
  if (periodo) {
    // Cumple si la última inspección fue en el MISMO período que hoy
    return _mismoPeriodoCalInsp(periodo, ultimaFecha, hoy) ? 'Cumple' : 'No cumple';
  }
  // Rolling: cumple si hoy <= ultima + diasFrecuencia
  var fechaLimite = new Date(ultimaFecha);
  fechaLimite.setDate(fechaLimite.getDate() + diasFrecuencia);
  return (hoy > fechaLimite) ? 'No cumple' : 'Cumple';
}

/**
 * Días vencido para alertas:
 *   - Calendario: días desde inicio del período sin inspección (o 0 si ya se hizo).
 *   - Rolling: días desde que quedó disponible sin inspección.
 * @returns {number} negativo = aún no vence, 0 = vence hoy, positivo = vencido
 */
function _diasVencidoInsp(freq, ultimaDate, hoyDate) {
  var periodo = _freqEsCalendario(freq);
  if (periodo) {
    if (_mismoPeriodoCalInsp(periodo, ultimaDate, hoyDate)) return -999; // ya hecho este período
    var inicio = _inicioPeriodoActualInsp(periodo, hoyDate);
    return Math.floor((hoyDate - inicio) / 86400000);
  }
  // Rolling: el día de inspección NO se cuenta, se avanza freq días completos
  var nextDue = new Date(ultimaDate);
  nextDue.setDate(nextDue.getDate() + freq);
  return Math.floor((hoyDate - nextDue) / 86400000);
}