
// --- CONSTANTES ---
const EMPLOYEES_SS_ID = "1SrkbAD8aoLGCCr8oMh0yRp3iiRl0Du4WEpUU88zOCOc";
const SPREADSHEET_ID = "12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE"; 

// ✅ NUEVAS CONSTANTES PARA GUARDADO JSON
const FOLDER_DB_ID = "17tKcRGZtUjE0HwosxlGrycFWIJ20aaS8"; // CARPETA ROL DE TURNOS
const DB_FILENAME = "rol_turnos.json";
const DEPT_CONFIG_FILENAME = "department_config.json";

// --- CACHE PARA OPTIMIZACIÓN ---
let cachedEmployeesSheet = null;
let cachedMainSheet = null;


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getEmployeesSheet() {
  if (!cachedEmployeesSheet) {
    cachedEmployeesSheet = SpreadsheetApp.openById(EMPLOYEES_SS_ID);
  }
  return cachedEmployeesSheet;
}

function getMainSheet() {
  if (!cachedMainSheet) {
    cachedMainSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return cachedMainSheet;
}

// --- FUNCIÓN PRINCIPAL: OBTENER EMPLEADOS ---
function getEmployeesFromDB() {
  try {
    // Fuente única: hoja PERSONAL del spreadsheet principal
    var shPersonal = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    if (!shPersonal) {
      Logger.log("getEmployeesFromDB: hoja PERSONAL no encontrada");
      return JSON.stringify([]);
    }

    var personalData = shPersonal.getDataRange().getValues();
    var empList = [];

    for (var i = 1; i < personalData.length; i++) {
      var pr         = personalData[i];
      var id         = pr[0];                                       // Col A: ID
      var dni        = String(pr[1]  || '').trim();                 // Col B: DNI
      var rawName    = String(pr[2]  || '').trim();                 // Col C: Nombre
      var puesto     = String(pr[6]  || '').trim();                 // Col G: Cargo
      var situacion  = String(pr[11] || '').trim().toUpperCase();   // Col L: ACTIVO/SI/CESADO…
      var foto       = String(pr[14] || '').trim();                 // Col O: Foto
      var autorizado = String(pr[15] || '').trim().toUpperCase();   // Col P: SI/NO acceso
      var fechaCese  = pr[18];                                      // Col S: Fecha de Cese

      if (!id || !rawName) continue;

      // Bloqueado del login → siempre excluir
      if (autorizado === 'NO') continue;

      // Si no es ACTIVO ni SI (LIQUIDADO, CESADO, RETIRADO, BAJA, etc.)
      if (situacion && situacion !== 'ACTIVO' && situacion !== 'SI') {
        // Tiene fecha de cese → incluir como inactivo hasta que pase esa fecha
        if (fechaCese instanceof Date && !isNaN(fechaCese)) {
          var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
          var cese = new Date(fechaCese); cese.setHours(0, 0, 0, 0);
          if (hoy > cese) {
            // Pasó la fecha de cese → agregar con activo:false (turnos visibles, sin sidebar)
            empList.push({
              id: id.toString(), dni: dni,
              name: generateShortName(rawName), fullName: rawName,
              jobTitle: puesto, num: id, foto: foto, activo: false
            });
            continue;
          }
          // Si hoy <= cese: sigue apareciendo normalmente hasta su último día
        } else {
          continue; // Sin fecha de cese → excluir completamente
        }
      }

      empList.push({
        id: id.toString(), dni: dni,
        name: generateShortName(rawName), // Nombre corto (Web)
        fullName: rawName,                // Nombre completo (Excel)
        jobTitle: puesto, num: id, foto: foto
      });
    }

    return JSON.stringify(empList);
  } catch (e) {
    Logger.log("Error leyendo DB: " + e.toString());
    return JSON.stringify([]);
  }
}

// --- GENERADOR DE NOMBRES CORTOS (VERSIÓN FINAL) ---
function generateShortName(fullName) {
  if (!fullName) return "";
  
  // Dividimos por espacios
  // Tu data: "Caracela Flores Victor Renzo"
  let parts = fullName.trim().split(" ");
  
  if (parts.length >= 2) {
    // Tomamos el último trozo como Nombre (Renzo)
    let firstName = parts[parts.length - 1]; 
    // Tomamos el primer trozo como Apellido (Caracela)
    let lastName = parts[0]; 
    
    // Devolvemos: "R. Caracela" (Sin cortar)
    return firstName.charAt(0) + ". " + lastName;
  } else {
    return fullName;
  }
}

// --- FUNCIÓN PRINCIPAL: GUARDAR REPORTE EN EXCEL ---
function saveFullReport(payload) {
  // ✅ CAMBIA ESTE ID POR EL DE TU HOJA
  var targetSpreadsheetID = "12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE"; // <-- IMPORTANTE: Cambia esto
  
  Logger.log("🔍 Intentando abrir hoja: " + targetSpreadsheetID);
  
  var ss;
  try {
    ss = SpreadsheetApp.openById(targetSpreadsheetID);
    Logger.log("✅ Hoja abierta exitosamente: " + ss.getName());
  } catch (e) {
    Logger.log("❌ ERROR abriendo hoja: " + e.toString());
    throw new Error("❌ No puedo acceder a la hoja " + targetSpreadsheetID + ". Verifica el ID y permisos.");
  }
  
  var timestamp = Utilities.formatDate(new Date(), "GMT-5", "yyyy-MM-dd HH:mm:ss");
  var savedCount = 0;
  
  // --- GUARDAR DETALLES ---
  if(payload.details && payload.details.length > 0) {
    Logger.log("📝 Guardando " + payload.details.length + " registros de detalle...");
    
    var sheetDet = ss.getSheetByName("BD_Detalle");
    if (!sheetDet) {
      Logger.log("⚠️ Creando nueva pestaña BD_Detalle");
      sheetDet = ss.insertSheet("BD_Detalle");
      sheetDet.appendRow(["FECHA", "LUNES_SEMANA", "ID_EMP", "DNI", "NOMBRE", "ZONA", "TURNO", "HORAS", "TIPO", "CODIGO", "REGISTRADO_EL"]);
    }
    
    try {
      var rowsDet = payload.details.map(function(r) { return r.concat([timestamp]); });
      sheetDet.getRange(sheetDet.getLastRow() + 1, 1, rowsDet.length, rowsDet[0].length).setValues(rowsDet);
      savedCount += rowsDet.length;
      Logger.log("✅ Detalles guardados correctamente");
    } catch (e) {
      Logger.log("❌ Error guardando detalles: " + e.toString());
      throw e;
    }
  }

  // --- GUARDAR RESUMEN ---
  if(payload.summary && payload.summary.length > 0) {
    Logger.log("📊 Guardando " + payload.summary.length + " registros de resumen...");
    
    var sheetRes = ss.getSheetByName("BD_Resumen_Semanal");
    if (!sheetRes) {
      Logger.log("⚠️ Creando nueva pestaña BD_Resumen_Semanal");
      sheetRes = ss.insertSheet("BD_Resumen_Semanal");
      sheetRes.appendRow(["LUNES_SEMANA", "ID_EMP", "DNI", "NOMBRE", "HH_TOTAL", "HH_REGULAR", "HH_EXTRA", "NOCHES", "DIAS_TRAB", "DIAS_DESC", "TIENE_AUS", "DETALLE_AUS", "ESTADO", "REGISTRADO_EL"]);
    }

    try {
      var rowsRes = payload.summary.map(function(r) { return r.concat([timestamp]); });
      sheetRes.getRange(sheetRes.getLastRow() + 1, 1, rowsRes.length, rowsRes[0].length).setValues(rowsRes);
      Logger.log("✅ Resumen guardado correctamente");
    } catch (e) {
      Logger.log("❌ Error guardando resumen: " + e.toString());
      throw e;
    }
  }
  
  var finalMessage = "✅ ÉXITO TOTAL: " + savedCount + " registros guardados en " + ss.getName();
  Logger.log(finalMessage);
  return finalMessage;
}




function testSpreadsheetAccess() {
  try {
    var ss = SpreadsheetApp.openById(EMPLOYEES_SS_ID);
    Logger.log("✅ Conexión exitosa con: " + ss.getName());
  } catch (e) {
    Logger.log("❌ ERROR de conexión: " + e.toString());
  }
}

// ✅ NUEVA FUNCIÓN: LEER DATOS DE LA PESTAÑA "MOF"
function getMOFConfigData() {
  try {
    // Usamos el mismo ID que usas para empleados
    var ss = SpreadsheetApp.openById(EMPLOYEES_SS_ID); 
    var sheet = ss.getSheetByName("MOF");
    
    if (!sheet) {
      // Fallback si no existe la pestaña, para que no rompa
      return JSON.stringify([]);
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]);

    // Leemos desde la fila 2 hasta la última, Columnas B (2) y C (3)
    // getRange(fila, col, numFilas, numCols) -> B es col 2.
    // Para asegurar, leemos A, B, C (cols 1, 2, 3) y filtramos luego.
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); 
    
    var rolesList = [];
    var seen = new Set();

    data.forEach(row => {
      // Asumiendo estructura visual: A=ID, B=AREA, C=CARGO
      var area = row[1]; // Columna B
      var cargo = row[2]; // Columna C

      if (area && cargo) {
        var cargoClean = cargo.toString().trim().toUpperCase();
        var areaClean = area.toString().trim().toUpperCase();

        // Creamos una lista única de Cargos con su Área asociada
        // Esto servirá para validar y para la configuración maestra
        if (!seen.has(cargoClean)) {
          rolesList.push({
            name: cargoClean,     // El nombre del Rol/Cargo
            department: areaClean // El Área a la que pertenece
          });
          seen.add(cargoClean);
        }
      }
    });

    return JSON.stringify(rolesList);

  } catch (e) {
    Logger.log("❌ Error leyendo MOF: " + e.toString());
    return JSON.stringify([]);
  }
}

// ✅ ============= NUEVAS FUNCIONES PARA GUARDADO JSON =============

// --- FUNCIÓN PARA GUARDAR JSON EN DRIVE ---
function savePlannerToDrive(jsonString) {
  try {
    var props = PropertiesService.getScriptProperties();
    var fileId = props.getProperty("DB_FILE_ID_V4");
    var file;

    // --- INTENTO RÁPIDO (Acceso directo por ID) ---
    if (fileId) {
      try {
        file = DriveApp.getFileById(fileId);
        
        // ✅ NUEVA VALIDACIÓN: Si está en la papelera, lo tratamos como borrado
        if (file.isTrashed()) {
           throw new Error("El archivo está en la papelera");
        }
        
        file.setContent(jsonString);
        return "⚡ Guardado Rápido (ID) - " + Utilities.formatDate(new Date(), "GMT-5", "HH:mm:ss");
      } catch (e) {
        // Si falla o está en la papelera, borramos la memoria para crear uno nuevo
        props.deleteProperty("DB_FILE_ID_V4");
      }
    }

    // ... (El resto de la función sigue igual: busca por nombre o crea uno nuevo) ...
    var folder;
    try {
      folder = DriveApp.getFolderById(FOLDER_DB_ID);
    } catch (e) {
      return "❌ Error: Carpeta no encontrada.";
    }

    var files = folder.getFilesByName(DB_FILENAME);
    
    if (files.hasNext()) {
      file = files.next();
      file.setContent(jsonString);
    } else {
      file = folder.createFile(DB_FILENAME, jsonString, MimeType.PLAIN_TEXT);
    }

    props.setProperty("DB_FILE_ID_V4", file.getId());
    return "✅ Guardado (Nuevo ID) - " + Utilities.formatDate(new Date(), "GMT-5", "HH:mm:ss");

  } catch (e) {
    return "❌ Error crítico: " + e.toString();
  }
}



// --- FUNCIÓN PARA CARGAR JSON DESDE DRIVE ---
function loadPlannerFromDrive() {
  try {
    let folder;
    try { 
      folder = DriveApp.getFolderById(FOLDER_DB_ID); 
    } catch (e) { 
      Logger.log("⚠️ Carpeta no encontrada");
      return null; 
    }
    
    const files = folder.getFilesByName(DB_FILENAME);
    if (files.hasNext()) {
      let content = files.next().getBlob().getDataAsString();
      Logger.log("✅ Datos JSON cargados exitosamente");
      return content;
    } else {
      Logger.log("ℹ️ No se encontró archivo de base de datos");
      return null;
    }
  } catch (e) { 
    Logger.log("❌ Error cargando datos: " + e.toString());
    return null; 
  }
}

// --- FUNCIÓN PARA GUARDAR CONFIGURACIÓN DE DEPARTAMENTOS ---
function saveDepartmentConfig(config) {
  try {
    let folder = DriveApp.getFolderById(FOLDER_DB_ID);
    let files = folder.getFilesByName(DEPT_CONFIG_FILENAME);
    let content = JSON.stringify(config, null, 2);
    
    if (files.hasNext()) {
      files.next().setContent(content);
    } else {
      folder.createFile(DEPT_CONFIG_FILENAME, content, MimeType.PLAIN_TEXT);
    }
    
    return "✅ Configuración de departamentos guardada";
  } catch(e) { 
    return "❌ Error guardando configuración: " + e.toString(); 
  }
}

// --- FUNCIÓN PARA CARGAR CONFIGURACIÓN DE DEPARTAMENTOS ---
function getDepartmentConfig() {
  try {
    let folder = DriveApp.getFolderById(FOLDER_DB_ID);
    let files = folder.getFilesByName(DEPT_CONFIG_FILENAME);
    if (files.hasNext()) {
      return JSON.parse(files.next().getBlob().getDataAsString());
    }
  } catch(e) { 
    Logger.log("⚠️ Configuración de departamentos no encontrada"); 
  }
  return {};
}

// --- FUNCIÓN DE PRUEBA PARA VERIFICAR GUARDADO ---
function testJSONSave() {
  try {
    // Datos de prueba
    var testData = {
      test: true,
      timestamp: new Date().toISOString(),
      message: "Prueba de guardado JSON"
    };
    
    var result = savePlannerToDrive(JSON.stringify(testData, null, 2));
    Logger.log("Resultado del test: " + result);
    
    // Intentar cargar
    var loaded = loadPlannerFromDrive();
    if (loaded) {
      var parsed = JSON.parse(loaded);
      Logger.log("✅ Test exitoso - Datos cargados: " + parsed.message);
    } else {
      Logger.log("⚠️ No se pudieron cargar los datos de prueba");
    }
    
  } catch (e) {
    Logger.log("❌ Error en test: " + e.toString());
  }
}

function getRolInitialData() {
  return {
    employees: JSON.parse(getEmployeesFromDB()), // Usamos tu función existente
    mof: JSON.parse(getMOFConfigData())          // Usamos tu función existente
  };
}