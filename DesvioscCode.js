const desvios = SpreadsheetApp.openById("1eIJfA7dAlkQ1rXcRGC2qSFnvZ-jYIPn8cA_TbUZcWZE")
const folderimginspe = "1G_RoN2wmOzoso4zupMfi240S7y1criFk"   //INSPECCIONES - IMG
const folderpdfinspe = "199bLX65wB1pkFT9U2b-YnpxPm9OZmdCm"    //INSPECCIONES - PDF
// var logo = desvios.getSheetByName('MENÚ').getRange('B18:B18').getValues()

let cachedDesvios = null;
function getDesviosSpreadsheet() {
  if (!cachedDesvios) {
    cachedDesvios = SpreadsheetApp.openById(SPREADSHEET_IDS.desvios);
  }
  return cachedDesvios;
}

// Parser robusto: acepta "DD/MM/YYYY", "DD/MM/YYYY HH:mm:ss",
// "YYYY-MM-DD", "YYYY-MM-DD HH:mm", "YYYY-MM-DDTHH:mm"
// Retorna { m, y } o null
function _parseMesAnioDesvio(txt) {
  if (!txt) return null;
  var s = String(txt).trim();
  var m;
  // dd/MM/yyyy ...
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return { m: parseInt(m[2], 10), y: parseInt(m[3], 10) };
  // yyyy-MM-dd ...
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return { m: parseInt(m[2], 10), y: parseInt(m[1], 10) };
  return null;
}

// === CACHE de la lectura completa de la hoja ===
// Evita re-leer A:U en cada paginación / cambio de chip
const _DESVIOS_CACHE_KEY = 'desvios_full_v2';
const _DESVIOS_CACHE_TTL = 300; // 5 minutos

function _readDesviosFullCached() {
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get(_DESVIOS_CACHE_KEY);
    if (raw) {
      try { return JSON.parse(raw); } catch (e) { /* fallthrough */ }
    }
  } catch (e) {
    Logger.log('Cache read error: ' + e.message);
  }
  // Miss → leer hoja
  const hoja = getDesviosSpreadsheet().getSheetByName("B DATOS");
  const lastRow = hoja.getLastRow();
  if (lastRow < 1) return { headers: [], registros: [] };
  const datos = hoja.getRange(1, 1, lastRow, 21).getDisplayValues();
  const result = { headers: datos[0], registros: datos.slice(1) };
  // Guardar en cache (si entra en 100KB)
  try {
    const json = JSON.stringify(result);
    if (json.length < 95 * 1024) {
      CacheService.getScriptCache().put(_DESVIOS_CACHE_KEY, json, _DESVIOS_CACHE_TTL);
    }
  } catch (e) { /* skip cache */ }
  return result;
}

function _invalidateDesviosCache() {
  try { CacheService.getScriptCache().remove(_DESVIOS_CACHE_KEY); } catch (e) {}
}

function getDatosRegistro(offset, limit, filtroMes) {
  try {
    const { headers, registros } = _readDesviosFullCached();

    const mesNum = (filtroMes && filtroMes !== "Todos" && filtroMes !== "todos") ? parseInt(filtroMes, 10) : null;

    const filtrados = mesNum === null ? registros : registros.filter(fila => {
      var parsed = _parseMesAnioDesvio(fila[10]);
      if (!parsed) return false;
      // Solo filtramos por mes (sin restricción de año) para que el usuario vea
      // todos los registros del mes seleccionado, sin importar el año
      return parsed.m === mesNum;
    });

    const start = Math.max(filtrados.length - offset - limit, 0);
    const end = filtrados.length - offset;
    const paginados = filtrados.slice(start, end).reverse();

    return {
      headers,
      data: paginados,
      total: filtrados.length
    };
  } catch (error) {
    Logger.log("⚠️ Error en getDatosRegistro: " + error.message);
    return {
      headers: [],
      data: [],
      total: 0,
      error: error.message
    };
  }
}

function getHeaders() {
  const hoja = getDesviosSpreadsheet().getSheetByName("B DATOS");
  return hoja.getRange(1, 1, 1, 21).getValues()[0]; // Fila 1, columnas A a U (21 columnas)
}

function globalVariablesDesvios() {
  const spreadsheet = getDesviosSpreadsheet();

  return {
    spreadsheetId : spreadsheet.getId(), // 👍 Más robusto
    dataRage      : 'B DATOS!A2:U',
    idRange       : 'B DATOS!A2:A',
    lastCol       : 'U',
    insertRange   : 'B DATOS!A1:U1',
    sheetID       : spreadsheet.getSheetByName("B DATOS").getSheetId() // 💡 opcional
  };
}

function processFormDesvios(formObject) {
  try {
    const values = getFormValuesDesvios(formObject);
    const id = values[0][0];
    if (formObject.RecId && checkIDDesvios(formObject.RecId)) {
      updateDataDesvios(values, globalVariablesDesvios().spreadsheetId, getRangeByIDDesvios(formObject.RecId));
    } else {
      appendDataDesvios(values, globalVariablesDesvios().spreadsheetId, globalVariablesDesvios().insertRange);
    }
    return id; // <-- Devuelve ID al frontend como éxito
  } catch (e) {
    throw new Error("Error en el procesamiento del formulario: " + e.message);
  }
}

var folder = DriveApp.getFolderById(folderimginspe);

// FUNCIÓN 4 - Extrae los valores del formulario y genera un array para su almacenamiento.(FUNCIONA)
function getFormValuesDesvios(formObject) {

  // Fecha y hora SIN AM/PM (formato 24h)
  var now = new Date();
  var formattedDate = Utilities.formatDate(
    now,
    'America/Lima',
    'dd/MM/yyyy HH:mm:ss'
  );

  // Verificar si se subió imagen anotada (canvas), original o conservar anterior
  let imageUrl = "";
  if (formObject.annotatedData && formObject.annotatedData.length > 10) {
    // 🎨 Imagen con anotaciones dibujadas — subir a Drive desde base64
    try {
      var annotatedBlob = Utilities.newBlob(
        Utilities.base64Decode(formObject.annotatedData),
        'image/jpeg',
        'desvio_anotado_' + Date.now() + '.jpg'
      );
      var file1Ann = folder.createFile(annotatedBlob);
      imageUrl = "https://lh5.googleusercontent.com/d/" + file1Ann.getId();
    } catch (eAnn) {
      Logger.log("Error subiendo imagen anotada, usando original: " + eAnn);
      if (formObject.myFile1 && formObject.myFile1.length > 0) {
        imageUrl = "https://lh5.googleusercontent.com/d/" + folder.createFile(formObject.myFile1).getId();
      } else {
        imageUrl = formObject.imgAnterior1 || "";
      }
    }
  } else if (formObject.myFile1 && formObject.myFile1.length > 0) {
    let file1 = folder.createFile(formObject.myFile1);
    imageUrl = "https://lh5.googleusercontent.com/d/" + file1.getId();
  } else {
    imageUrl = formObject.imgAnterior1 || ""; // conservar imagen anterior
  }

  // Verificar si se subió una nueva imagen de cierre
  let imageUrl2 = "";
  if (formObject.myFile2 && formObject.myFile2.length > 0) {
    let file2 = folder.createFile(formObject.myFile2);
    imageUrl2 = "https://lh5.googleusercontent.com/d/" + file2.getId();
  } else {
    imageUrl2 = formObject.imgAnterior2 || ""; // conservar imagen anterior
  }

  let values;

  if (formObject.RecId && checkIDDesvios(formObject.RecId)) {
    // Edición
    values = [[
      formObject.RecId.toString(),
      formObject.nameDesvios,
      formObject.nombreDesvios,
      formObject.num,
      formObject.classroom,
      formObject.gender,
      formObject.address,
      formObject.emailDesvios,
      formObject.descripcion,
      formObject.resp,
      formObject.fechaDesvios,
      formObject.clasi,
      formObject.amonestado,
      formObject.procesoDesvios,
      formObject.acciones,
      formObject.estado,
      formObject.plan,
      imageUrl,        // Columna Q
      formattedDate,   // Columna R (fecha sin AM/PM)
      imageUrl2,       // Columna S
      formObject.estadoDesvios
    ]];
  } else {
    // Nuevo registro
    values = [[
      new Date().getTime().toString(),
      formObject.nameDesvios,
      formObject.nombreDesvios,
      formObject.num,
      formObject.classroom,
      formObject.gender,
      formObject.address,
      formObject.emailDesvios,
      formObject.descripcion,
      formObject.resp,
      formObject.fechaDesvios,
      formObject.clasi,
      formObject.amonestado,
      formObject.procesoDesvios,
      formObject.acciones,
      formObject.estado,
      formObject.plan,
      imageUrl,
      formattedDate,
      imageUrl2,
      formObject.estadoDesvios
    ]];
  }

  return values;
}

// FUNCIÓN 5 - Agrega una nueva fila a la hoja de cálculo.
function appendDataDesvios(values, spreadsheetId, range) {
  var valueRange = Sheets.newValueRange();
  valueRange.values = values;
  Sheets.Spreadsheets.Values.append(valueRange, spreadsheetId, range, {
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS"
  });
  _invalidateDesviosCache();

  // Procesos en segundo plano con manejo de errores silencioso
  try {
    emailDesvios();
  } catch (e) {
    Logger.log("Error en emailDesvios: " + e.message);
  }

  try {
    setDateDesvios();
  } catch (e) {
    Logger.log("Error en setDateDesvios: " + e.message);
  }
}

// FUNCIÓN 6 - Obtiene datos de una hoja de cálculo y devuelve los valores
function readDataDesvios(spreadsheetId,range){
  var result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

// FUNCIÓN 7 - Actualiza una fila existente.
function updateDataDesvios(values, spreadsheetId, range) {
  try {
    var valueRange = Sheets.newValueRange();
    valueRange.values = values;
    Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, range, {
      valueInputOption: "RAW"
    });

    setDateDesvios();
    _invalidateDesviosCache();

  } catch (e) {
    throw new Error("Error al actualizar datos: " + e.message);
  }
}

// FUNCIÓN 8 - Elimina una fila con un ID específico.
function deleteDataDesvios(ID){ 
  var startIndex = getRowIndexByIDDesvios(ID);
  
  var deleteRange = {
                      "sheetId"     : globalVariablesDesvios().sheetID,
                      "dimension"   : "ROWS",
                      "startIndex"  : startIndex,
                      "endIndex"    : startIndex+1
                    }
  
  var deleteRequest= [{"deleteDimension":{"range":deleteRange}}];
  Sheets.Spreadsheets.batchUpdate({"requests": deleteRequest}, globalVariablesDesvios().spreadsheetId);
  _invalidateDesviosCache();
}

// FUNCIÓN 9 - Verifica si un ID existe en la base de datos, usando readDataDesvios.
function checkIDDesvios(ID){
  var idList = readDataDesvios(globalVariablesDesvios()
  .spreadsheetId,globalVariablesDesvios().idRange,)
  .reduce(function(a,b){
    return a.concat(b);
    });
  return idList.includes(ID);
}

// FUNCIÓN 10 - Obtiene el rango de celdas de un registro específico.
function getRangeByIDDesvios(id){
  if(id){
    var idList = readDataDesvios(globalVariablesDesvios().spreadsheetId,globalVariablesDesvios().idRange);
    for(var i=0;i<idList.length;i++){
      if(id==idList[i][0]){
        return 'B DATOS!A'+(i+2)+':'+globalVariablesDesvios().lastCol+(i+2);
      }
    }
  }
}

// FUNCIÓN 11 - Retorna un registro si el ID existe.
function getRecordByIdDesvios(id){
  if(id && checkIDDesvios(id)){
    var result = readDataDesvios(globalVariablesDesvios().spreadsheetId,getRangeByIDDesvios(id));
    return result;
  }
}

// FUNCIÓN 12 - Obtiene el índice de fila de un ID.
function getRowIndexByIDDesvios(id){
  if(id){
    var idList = readDataDesvios(globalVariablesDesvios().spreadsheetId,globalVariablesDesvios().idRange);
    for(var i=0;i<idList.length;i++){
      if(id==idList[i][0]){
        var rowIndex = parseInt(i+1);
        return rowIndex;
      }
    }
  }
}

function getTotalDesvios() {
  const ss = SpreadsheetApp.openById(globalVariablesDesvios().spreadsheetId);
  const sheet = ss.getSheetByName("B DATOS");
  return sheet.getLastRow() - 1; // Sin contar encabezado
}

//FUNCIÓN 18 - Formatea la columna R como fecha y hora
function setDateDesvios() {
  var sheet = getDesviosSpreadsheet().getSheetByName('B DATOS');
  var lastRow = sheet.getLastRow();
  var dateRange = sheet.getRange('S2:S' + lastRow);
  dateRange.setNumberFormat('d/M/yyyy, H:mm:ss');
}

// FUNCIÓN 30 - Busca un valor en la columna B de GENERAL y devuelve el valor de la columna C correspondiente.
function onInputChange(searchtext) {
  //var spreadsheetId = '14MX0wRMyTTM1y6kVNbk6tuqVKpsoycwDraUviTYRkuQ';
  var sheetName = 'PERSONAL';
  var range = 'B:C';
  var sheet = getSpreadsheetPersonal().getSheetByName(sheetName);
  var data = sheet.getRange(range).getValues();

  var nombreDesvios = "";
  data.forEach(function(row) {
    if (row[0] === searchtext) {
      nombreDesvios = row[1];
    }
  });
  return nombreDesvios;
}

function getPasswordsDesvios() {
  const hoja = getDesviosSpreadsheet().getSheetByName("Acceso");
  const ultimaFila = hoja.getLastRow();
  const colE = hoja.getRange("D2:D" + ultimaFila).getValues().flat(); // Para eliminar
  const colD = hoja.getRange("C2:C" + ultimaFila).getValues().flat(); // Para editar

  return {
    eliminar: colE,
    editar: colD
  };
}


// FUNCIÓN 33 - Obtiene un valor de la celda Z1 en la hoja B DATOS. Relación: Puede ser usada en reportes o actualizaciones de estado.
function setStatusDesvios(){
  let sst = getDesviosSpreadsheet().getSheetByName('B DATOS')
  let total1 = sst.getRange("AA1").getValue();
  let total2 = sst.getRange("AB1").getValue();

  //Logger.log([total1,total2])
  return[total1, total2]
}

// FUNCIÓN 34 - Escribe un recordId en V5, genera un PDF de la hoja FICHA RAC T1, lo sube a Google Drive y devuelve el enlace.
function setIDAndGetLink(recordId) {
  var sheet = getDesviosSpreadsheet().getSheetByName('FICHA RAC T1');
  sheet.getRange('L5').setValue(recordId);
  
  // Asegurarse de que los cambios se apliquen
  SpreadsheetApp.flush();
  Utilities.sleep(150);

  var sheetId = sheet.getSheetId();
  var url = getDesviosSpreadsheet().getUrl().replace(/edit$/, '');
  // var exportUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:AB38&size=0&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
  var exportUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=B1:P35';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  var blob = response.getBlob().setName(sheet.getName() + '.pdf');
  
  // Obtener la carpeta específica por ID o nombre
  //var folderId = '1DnQvjDVgN7E0FEGBMCZQLppV62hhme6f';
  var folder = DriveApp.getFolderById(folderpdfinspe);
  
  // Crear el archivo en la carpeta especificada
  var file = folder.createFile(blob);

  // Hacer que el archivo sea público
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}


//Inspecciones:

function getFilteredData(filters) {
  const sheet = getDesviosSpreadsheet().getSheetByName("B DATOS");
  const lastRow = sheet.getLastRow();

  // Leer solo hasta columna U (índice 20, o columna 21)
  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues(); // Sin encabezado

  // Índices de columnas necesarias para mostrar
  const showIndexes = [8, 5, 4, 13, 7, 17, 19, 6, 9, 20, 10, 16];

  const fechaInicial = filters.fechaInicial ? new Date(filters.fechaInicial) : null;
  const fechaFinal = filters.fechaFinal ? new Date(filters.fechaFinal) : null;
  if (fechaFinal) fechaFinal.setDate(fechaFinal.getDate() + 1); // incluir todo el día final

  const match = (row) => {
    const fechaRegistro = new Date(row[10]); // Columna K

    return (!filters.empresa         || row[3].toString().includes(filters.empresa)) &&
           (!filters.lugar           || row[4].toString().includes(filters.lugar)) &&
           (!filters.blanco          || row[5].toString().includes(filters.blanco)) &&
           (!filters.proceso         || row[13].toString().includes(filters.proceso)) &&
           (!filters.nombre          || row[2].toString().includes(filters.nombre)) &&
           (!filters.responsable     || row[9].toString().includes(filters.responsable)) &&
           (!filters.clasificacion   || row[11].toString().includes(filters.clasificacion)) &&
           (!filters.amonestado      || row[12].toString().includes(filters.amonestado)) &&
           (!filters.filtroestadoinpse || row[20].toString().includes(filters.filtroestadoinpse)) &&
           (!fechaInicial || fechaRegistro >= fechaInicial) &&
           (!fechaFinal   || fechaRegistro < fechaFinal);
  };

  const filtered = data.filter(match);
  const result = filtered.map(row => showIndexes.map(i => row[i]));
  return result;
}



//PDF INSPECCIONES
function guardarDatosYGenerarPDF(datos) {
  const hoja = getDesviosSpreadsheet().getSheetByName('INSPECCIÓN');

  hoja.getRange('B6').setValue(datos.empresa);
  hoja.getRange('G12').setValue(datos.fechaInicial);
  hoja.getRange('I12').setValue(datos.fechaFinal);
  hoja.getRange('C9').setValue(datos.lugar);
  hoja.getRange('D8').setValue(datos.inspeccionadoPor);
  hoja.getRange('I11').setValue(datos.blanco);
  hoja.getRange('C10').setValue(datos.responsableArea);
  hoja.getRange('I10').setValue(datos.filtroestadoinpse);

  return generarPDFInspeccion();
}

function generarPDFInspeccion() { 
  const sheet = getDesviosSpreadsheet().getSheetByName('INSPECCIÓN');
  const sheetId = sheet.getSheetId();
  const url = getDesviosSpreadsheet().getUrl().replace(/edit$/, '');

  // Buscar la última fila con fecha válida en la columna N
  const valuesN = sheet.getRange("J1:J").getValues();
  let lastRow = 0;

  for (let i = 0; i < valuesN.length; i++) {
    const val = valuesN[i][0];
    const isDate = Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val);
    if (isDate) {
      lastRow = i + 1;
    }
  }

  if (lastRow < 1) {
    throw new Error('No se encontraron fechas válidas en la columna J.');
  }

  // Agregar el parámetro range para limitar la exportación
  const exportPdfUrl = url + 'export?format=pdf' +
    '&size=A4' +
    '&portrait=false' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false' +
    '&range=A1:N' + lastRow + // <- Aquí está la clave
    '&gid=' + sheetId;

  const token = ScriptApp.getOAuthToken();
  const responsePdf = UrlFetchApp.fetch(exportPdfUrl, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  const blobPdf = responsePdf.getBlob().setName('PDF_RAC.pdf');
  const folder = DriveApp.getFolderById(folderpdfcheck);
  const filePdf = folder.createFile(blobPdf);
  filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return filePdf.getUrl();
}


// FUNCIÓN 35: Enviar correo basado en datos de la hoja de cálculo
function emailDesvios() { 
  const ss = getDesviosSpreadsheet();
  const sheetDatos = ss.getSheetByName("B DATOS");
  const sheetMenu = ss.getSheetByName("MENÚ");

  const emailDesvios = sheetMenu.getRange("B21").getValue();
  const lastRow = sheetDatos.getLastRow();
  const values = sheetDatos.getRange(lastRow, 1, 1, 23).getValues();

  for (let i = 0; i < values.length; i++) { 
    const row = values[i];
    const value = row[6] ? row[6].toString().toLowerCase().trim() : '';

    // ✅ Detecta "Alto", "ALTO", "alto", etc.
    if (value === 'alto') { 
      const Empresa = row[3];
      const Lugar = row[4];
      const Blanco = row[5];
      const Subestandar = row[7];
      const Potencial = row[6];
      const Descripcion = row[8];
      const Imagenes = row[17];
      const Medida = row[10];
      const Responsable = row[9];
      const Estado = row[15];
      const DaysAging = row[11];
      const Reportante = row[2];
      const Org = row[16];
      const Reportado = row[12];
      const Proceso = row[13];
      const Correccion = row[14];
      const Imagen2 = row[19];
      const Situcion = row[20];
      const emailAddress = emailDesvios; 

      // ✅ Construir el cuerpo del correo en HTML
      const htmlMessage = `
        <div style="font-family: Arial, sans-serif; color: #333; max-width:700px;">
          <div style="background-color: #f44336; color: white; padding: 10px 20px; border-radius: 5px 5px 0 0;">
            <h2 style="margin: 0;">⚠️ Reporte de Acto / Condición Insegura - Potencial ${Potencial}</h2>
          </div>
          <div style="border: 1px solid #ddd; border-top: none; border-radius: 0 0 5px 5px; padding: 20px; background-color: #fafafa;">
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
              <tbody>
                <tr><td style="font-weight: bold; width: 30%;">Blanco:</td><td>${Blanco}</td></tr>
                <tr><td style="font-weight: bold;">Empresa:</td><td>${Empresa}</td></tr>
                <tr><td style="font-weight: bold;">Ubicación:</td><td>${Lugar}</td></tr>
                <tr><td style="font-weight: bold;">Subestándar:</td><td>${Subestandar}</td></tr>
                <tr><td style="font-weight: bold;">Descripción:</td><td>${Descripcion}</td></tr>
                <tr><td style="font-weight: bold;">Reportado por:</td><td>${Reportante}</td></tr>
                <tr><td style="font-weight: bold;">Fecha de registro:</td><td>${Medida}</td></tr>
                <tr><td style="font-weight: bold;">Tipo:</td><td>${DaysAging}</td></tr>
                <tr><td style="font-weight: bold;">Proceso:</td><td>${Proceso}</td></tr>
                <tr><td style="font-weight: bold;">Medidas correctivas:</td><td>${Correccion}</td></tr>
                <tr><td style="font-weight: bold;">Responsable:</td><td>${Responsable}</td></tr>
                <tr><td style="font-weight: bold;">Reportado:</td><td>${Reportado}</td></tr>
                <tr><td style="font-weight: bold;">Plan de acción:</td><td>${Org}</td></tr>
                <tr><td style="font-weight: bold;">Riesgo crítico:</td><td>${Estado}</td></tr>
                <tr><td style="font-weight: bold;">Situación:</td><td><strong style="color: red;">${Situcion}</strong></td></tr>
              </tbody>
            </table>

            ${Imagenes ? `
              <div style="margin-top: 20px;">
                <strong>☀︎ Imagen de la observación:</strong><br>
                <img src="${Imagenes}" alt="Imagen de la observación" style="max-width: 100%; border: 1px solid #ccc; border-radius: 5px; margin-top: 10px;">
              </div>` : ''}

            ${Imagen2 ? `
              <div style="margin-top: 20px;">
                <strong>☀︎ Imagen complementaria:</strong><br>
                <img src="${Imagen2}" alt="Imagen complementaria" style="max-width: 100%; border: 1px solid #ccc; border-radius: 5px; margin-top: 10px;">
              </div>` : ''}

            <p style="margin-top: 30px;">Saludos cordiales,<br><strong>Equipo de Seguridad</strong></p>
          </div>
        </div>
      `;

      const subject = `⚠️ Alerta: ${Blanco}, ${Subestandar}, de potencial ${Potencial}`;

      // ✅ Enviar el correo
      GmailApp.sendEmail(emailAddress, subject, '', { htmlBody: htmlMessage });
    }
  }
}

// FUNCIÓN 40 
/*
Similar a geminiAPI2, pero además permite personalizar el análisis con un texto dinámico tomado de la celda B3 de la hoja "ANALISIS". Luego, concatena los datos filtrados y el texto personalizado antes de enviarlos a la API de Gemini, y finalmente guarda el resultado en B4
*/
function geminiAPI3() { 
  const spreadsheet = getDesviosSpreadsheet();
  
  // Hojas de trabajo
  const sheetBD = spreadsheet.getSheetByName('B DATOS');
  const sheetANALISIS = spreadsheet.getSheetByName('ANALISIS');

  // Obtener el rango de fechas desde la hoja "ANALISIS"
  const fechaInicio = new Date(sheetANALISIS.getRange('B1').getValue());
  const fechaFin = new Date(sheetANALISIS.getRange('B2').getValue());

  // Obtener el último valor de la columna R (fechas) de la hoja "B DATOS"
  const lastRowBD = sheetBD.getLastRow();
  const rangoFechasBD = sheetBD.getRange(`S2:S${lastRowBD}`).getValues(); // Fechas en la columna R

  // Obtener los valores de la columna I que estén en el rango de fechas
  let concatenatedText = '';
  for (let i = 0; i < rangoFechasBD.length; i++) {
    const fechaBD = new Date(rangoFechasBD[i][0]);
    
    if (fechaBD >= fechaInicio && fechaBD <= fechaFin) {
      const valorI = sheetBD.getRange(`I${i + 2}`).getValue(); // Valor de la columna I en la fila correspondiente
      if (valorI) { // Si hay valor en la columna I
        concatenatedText += valorI + '. ';
      }
    }
  }

  // Obtener el texto dinámico desde la celda B3
  const textoAnalisis = sheetANALISIS.getRange('B3').getValue();

  // Preparar la solicitud solo si hay contenido en concatenatedText
  const cellB4 = sheetANALISIS.getRange('B4'); // Celda donde se mostrará el resultado
  if (concatenatedText) {
    const payload = {
      "contents": [
        {"parts": [
          { 
            "text": `${textoAnalisis}, de la siguiente base de datos ${concatenatedText}`
          }
        ]}
      ]
    };

    const params = {
      'contentType': 'application/json',
      'method': 'post',
      'payload': JSON.stringify(payload)
    };

    try {
      const response = UrlFetchApp.fetch(geminiUrl, params);
      const data = JSON.parse(response);
      const responseText = data.candidates[0].content.parts[0].text;

      // Escribir la respuesta en la hoja "ANALISIS", celda B4
      cellB4.setValue(responseText);

      // Opcional: Imprimir la respuesta en el log
      console.log(`Análisis general: ${responseText}`);
    } catch (error) {
      const errorMessage = `Error al obtener el análisis: ${error}`;
      cellB4.setValue(errorMessage); // Mostrar mensaje de error en la celda B4
      console.error(errorMessage);
    }
  } else {
    const noDataMessage = 'No hay datos en el rango de fechas especificado para analizar.';
    cellB4.setValue(noDataMessage); // Mostrar mensaje de no datos en la celda B4
    console.log(noDataMessage);
  }
}



//USADO PARA REALIZAR LA PETICIÓN A LA IA CUANDO SE PRESIONE EL BOTÓN "MEDIDAS CORRECTIVAS"
function geminiAPI4(concatenatedText) { 
  let responseText = "No se obtuvo respuesta."; // Inicializa con un valor predeterminado

  // Preparar la solicitud solo si hay contenido en concatenatedText
  if (concatenatedText) {
    const payload = {
      "contents": [
        {"parts": [
          { 
            "text": `Describe, en máximo 200 caracteres y sin comentarios introductorios, las Medidas Correctivas Inmediatas a realizar en base al siguiente reporte: ${concatenatedText}`
          }
        ]}
      ]
    };

    const params = {
      'contentType': 'application/json',
      'method': 'post',
      'payload': JSON.stringify(payload)
    };

    try {
      const response = UrlFetchApp.fetch(geminiUrl, params);
      const data = JSON.parse(response.getContentText()); // Obtener el texto y parsear JSON
      responseText = data.candidates[0]?.content?.parts[0]?.text || "No se obtuvo una respuesta válida.";

      // Escribir la respuesta en la celda J22 (descomentar si se usa en una hoja de cálculo)
      // const cellF = sheet.getRange('J22');
      // cellF.setValue(responseText);

      // Opcional: Imprimir la respuesta en el log
      console.log(`Análisis general: ${responseText}`);
    } catch (error) {
      console.error(`Error al obtener el análisis: ${error.message}`);
    }
  } else {
    console.log('No hay datos en la columna I para analizar.');
  }

  return responseText;
}


//USADO PARA REALIZAR LA PETICIÓN A LA IA CUANDO SE PRESIONE EL BOTÓN "GENERAR PLAN DE ACCIÓN"
function geminiAPI5(concatenatedText) { 
  let responseText = "No se obtuvo respuesta."; // Inicializa con un valor predeterminado

  // Preparar la solicitud solo si hay contenido en concatenatedText
  if (concatenatedText) {
    const payload = {
      "contents": [
        {"parts": [
          { 
            "text": `Describe, en máximo 200 caracteres y sin comentarios introductorios, un plan de acción global teniendo como referencia las gerarquias de controles (Eliminación, Sustitución, Controles de ingenieria, Controles administrativos y EPP) para poder contrarrestar este caso ${concatenatedText} y otros similares. Nota la geraquia de controles solo es una referencia no necesitas enumerar o listar`
          }
        ]}
      ]
    };

    const params = {
      'contentType': 'application/json',
      'method': 'post',
      'payload': JSON.stringify(payload)
    };

    try {
      const response = UrlFetchApp.fetch(geminiUrl, params);
      const data = JSON.parse(response.getContentText()); // Obtener el texto y parsear JSON
      responseText = data.candidates[0]?.content?.parts[0]?.text || "No se obtuvo una respuesta válida.";

      // Escribir la respuesta en la celda J22 (descomentar si se usa en una hoja de cálculo)
      // const cellF = sheet.getRange('J22');
      // cellF.setValue(responseText);

      // Opcional: Imprimir la respuesta en el log
      console.log(`Análisis general: ${responseText}`);
    } catch (error) {
      console.error(`Error al obtener el análisis: ${error.message}`);
    }
  } else {
    console.log('No hay datos en la columna I para analizar.');
  }

  return responseText;
}




// FUNCIÓN 2 -  CAPAZ DE USAR GEMINI PARA ANALIZAR IMÁGENES 
function describirImagen(imageUrl) {
  //Para poder analizar la imagen, necesito la URL directa de la imagen (que generalmente termina en .jpg, .jpeg, .png, .gif, etc.). No funcionará adecuadamente si le entregamos links de diferente formato al mencionado.
  //FUNCIONA, PERO NO ES USADA EN ESTA APLICACIÓN, PUES LOS LINKS QUE SE GENERAN NO TIENEN EL FORMATO DESEADO
  const apiUrl = geminiUrl;
  //const apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro-vision:generateContent?key=" + API_KEY;
  const requestBody = {
    contents: [
      {
        parts: [
          {
            inline_data: {
              mime_type: "image/jpeg", // O el tipo de MIME adecuado para tu imagen
              data: Utilities.base64Encode(UrlFetchApp.fetch(imageUrl).getBlob().getBytes())
            }
          },
          {
            text: "Describe la siguiente imagen en detalle. ¿Qué elementos ves? ¿Cuál crees que es el tema principal? Describe el entorno y cualquier otra característica relevante."
          }
        ]
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    console.log("Respuesta de la API:", response.getContentText()); // Para depuración
    const json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      return "No se pudo obtener una descripción de la imagen.";
    }
  } catch (error) {
    console.error("Error al analizar la imagen:", error);
    return "Hubo un error al intentar analizar la imagen.";
  }
}


//FUNCIÓN 3 - Analiza una imagen a partir de su contenido en base64 utilizando la API de Gemini y devuelve una descripción.
function describirImagenBase64(base64Image, mimeType) {
  //FUNCIÓN ACTUALMENTE USADA
  //RECIBE UNA IMAGEN CODIFICADA EN BASE64 Y LA DESCRIBE (ENTREGA UN TEXTO COMO SALIDA)
  const apiUrl = geminiUrl;
  const requestBody = {
    contents: [
      {
        parts: [
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Image
            }
          },
          {
            text: "Analiza las zonas marcadas con anotaciones visuales (flechas, círculos, rectángulos o trazos de colores) en esta imagen. Si no hay anotaciones, analiza toda la imagen. Describe en máximo 200 caracteres las condiciones inseguras observables y sus posibles consecuencias como incidentes, accidentes o impacto ambiental. Sé específico y directo. Sin introducciones."
          }
        ]
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody)
  };

  try {
    const response = UrlFetchApp.fetch(geminiUrl, options);
    console.log("Respuesta de la API (base64):", response.getContentText()); // Para depuración
    const json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      return "No se pudo obtener una descripción de la imagen.";
    }
  } catch (error) {
    console.error("Error al analizar la imagen (base64):", error);
    return "Hubo un error al intentar analizar la imagen.";
  }
}