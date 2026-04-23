// =============================================
// PASSO - Programa Anual de SSO (Backend)
// =============================================

// Helper: dias en un mes
function _diasEnMes(mes, anio) {
  return new Date(anio, mes, 0).getDate();
}

// Helper: dias laborables (lun-vie) en un mes
function _diasLaboralesMes(mes, anio) {
  var count = 0;
  var totalDias = _diasEnMes(mes, anio);
  for (var d = 1; d <= totalDias; d++) {
    var dow = new Date(anio, mes - 1, d).getDay();
    if (dow !== 0 && dow !== 6) count++;
  }
  return count;
}

// Helper: normalizar texto (quitar acentos y minúsculas)
function _normalizarPASSO(str) {
  return String(str || "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Helper: verificar si un tipo de la hoja coincide con el buscado
function _tipoCoincide(valorHoja, tipoBuscado) {
  var a = _normalizarPASSO(valorHoja);
  var b = _normalizarPASSO(tipoBuscado);
  if (!a || !b) return false;
  if (a === b) return true;
  var minLen = Math.min(a.length, b.length, 6);
  return a.substring(0, minLen) === b.substring(0, minLen);
}

// Helper: parsear frecuencia TEXTO a objeto de configuración
function _parsearFrecuenciaTexto(texto) {
  if (!texto) return null;
  var t = _normalizarPASSO(texto);
  if (!t) return null;

  if (t === "diaria" || t === "diario") return { label: "DIARIA",      cadaDias: 1,   clase: "diario"     };
  if (t === "semanal")                  return { label: "SEMANAL",     cadaDias: 7,   clase: "semanal"    };
  if (t === "quincenal")                return { label: "QUINCENAL",   cadaDias: 15,  clase: "2semanas"   };
  if (t === "mensual")                  return { label: "MENSUAL",     cadaMeses: 1,  clase: "mensual"    };
  if (t === "bimestral")                return { label: "BIMESTRAL",   cadaMeses: 2,  clase: "trimestral" };
  if (t === "trimestral")               return { label: "TRIMESTRAL",  cadaMeses: 3,  clase: "trimestral" };
  if (t === "semestral")                return { label: "SEMESTRAL",   cadaMeses: 6,  clase: "semestral"  };
  if (t === "anual")                    return { label: "ANUAL",       cadaMeses: 12, clase: "anual"      };

  var matchDias = t.match(/cada\s*(\d+)\s*dias?/);
  if (matchDias) {
    var d = parseInt(matchDias[1]);
    return { label: "C/" + d + " DÍAS", cadaDias: d, clase: d <= 7 ? "semanal" : "2semanas" };
  }

  var matchSem = t.match(/(?:c\/?|cada\s*)(\d+)\s*semanas?/);
  if (matchSem) {
    var s = parseInt(matchSem[1]);
    return { label: "C/" + s + " SEM", cadaDias: s * 7, clase: "2semanas" };
  }

  var matchMes = t.match(/(?:c\/?|cada\s*)(\d+)\s*meses?/);
  if (matchMes) {
    var mm = parseInt(matchMes[1]);
    return { label: "C/" + mm + " MESES", cadaMeses: mm, clase: mm <= 3 ? "trimestral" : "semestral" };
  }

  // 🔥 PUNTO 1: Fallback número puro → isNaN ya filtró antes, pero si llega aquí lo ignoramos
  var numMatch = t.match(/^(\d+)$/);
  if (numMatch) {
    var n = parseInt(numMatch[1]);
    if (n <= 1)   return { label: "DIARIA",      cadaDias: 1,   clase: "diario"     };
    if (n <= 7)   return { label: "SEMANAL",     cadaDias: 7,   clase: "semanal"    };
    if (n <= 15)  return { label: "QUINCENAL",   cadaDias: 15,  clase: "2semanas"   };
    if (n <= 31)  return { label: "MENSUAL",     cadaMeses: 1,  clase: "mensual"    };
    if (n <= 92)  return { label: "TRIMESTRAL",  cadaMeses: 3,  clase: "trimestral" };
    if (n <= 183) return { label: "SEMESTRAL",   cadaMeses: 6,  clase: "semestral"  };
    return { label: "ANUAL", cadaMeses: 12, clase: "anual" };
  }

  return null;
}

// Helper: cantidad de lunes (semanas completas lun-dom) cuyo lunes cae en el mes
function _semanasEnMes(mes, anio) {
  var count = 0;
  var diasMes = _diasEnMes(mes + 1, anio);
  for (var d = 1; d <= diasMes; d++) {
    if (new Date(anio, mes, d).getDay() === 1) count++; // contar lunes = inicio de semana
  }
  return count;
}

// Helper: generar conteo exacto por mes simulando fechas reales desde fechaInicio
// "Cada N días" = el día de inspección NO se cuenta, se avanza N días calendario
// Ejemplo: cada 4 días desde 3/mar → 3, 7, 11, 15... (saltos de N = 4 días)
function _calcularProgramadoAnualExacto(cadaDias, fechaInicio, anio) {
  var conteo = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  var finAnio = new Date(anio, 11, 31);
  var cursor = new Date(fechaInicio.getTime());

  // Si fechaInicio es de un año anterior, avanzar hasta llegar al año actual
  while (cursor.getFullYear() < anio) {
    cursor.setDate(cursor.getDate() + cadaDias);
  }
  // Si se pasó al año siguiente, no hay inspecciones este año
  if (cursor.getFullYear() > anio) return conteo;

  // Recorrer sumando al mes correspondiente
  while (cursor <= finAnio && cursor.getFullYear() === anio) {
    conteo[cursor.getMonth()]++;
    cursor.setDate(cursor.getDate() + cadaDias);
  }
  return conteo;
}

// Helper: calcular programados en un mes según frecuencia — lógica de calendario
// Reglas:
//   semanal (cadaDias=7)       → semanas cuyo LUNES cae en el mes
//   mensual (cadaMeses=1)      → 1 si es el mes en cuestión
//   bimestral (cadaMeses=2)    → 1 si mes es el 1° del bimestre (ene,mar,may,jul,sep,nov → basado en mes%2)
//   trimestral (cadaMeses=3)   → 1 si mes es inicio de trimestre (0,3,6,9)
//   semestral (cadaMeses=6)    → 1 si mes es inicio de semestre (0,6)
//   anual (cadaMeses=12)       → 1 solo en enero (mes=0)
//   diario (cadaDias=1)        → días laborales
//   otra periodicidad días     → cálculo exacto si hay fechaInicio, sino Math.ceil
function _calcularProgramadoMes(frecInfo, mes, anio) {
  if (!frecInfo) return 0;

  if (frecInfo.cadaDias) {
    if (frecInfo.cadaDias === 1) return _diasLaboralesMes(mes + 1, anio);
    if (frecInfo.cadaDias === 7) return _semanasEnMes(mes, anio);  // ← semanas calendario

    // Cálculo exacto si tenemos fecha de inicio
    if (frecInfo.fechaInicio instanceof Date) {
      if (!frecInfo._conteoExacto) {
        frecInfo._conteoExacto = _calcularProgramadoAnualExacto(frecInfo.cadaDias, frecInfo.fechaInicio, anio);
      }
      return frecInfo._conteoExacto[mes];
    }

    // Fallback: aproximación
    var diasMes = _diasEnMes(mes + 1, anio);
    return Math.max(1, Math.ceil(diasMes / frecInfo.cadaDias));
  }

  if (frecInfo.cadaMeses) {
    var cm = frecInfo.cadaMeses;
    if (cm === 1)  return 1;                                         // mensual: todos los meses
    if (cm === 2)  return (mes % 2 === 0)  ? 1 : 0;                 // bimestral: ene,mar,may,jul,sep,nov
    if (cm === 3)  return (mes % 3 === 0)  ? 1 : 0;                 // trimestral: ene,abr,jul,oct
    if (cm === 6)  return (mes % 6 === 0)  ? 1 : 0;                 // semestral: ene,jul
    if (cm >= 12)  return (mes === 0)      ? 1 : 0;                 // anual: solo enero
    return (mes % cm === 0) ? 1 : 0;
  }

  return 0;
}

// Helper: obtener o crear hoja REUNIONES
function _getHojaReuniones() {
  var ss = getSpreadsheetCapacitaciones();
  var hoja = ss.getSheetByName("REUNIONES");
  if (!hoja) {
    hoja = ss.insertSheet("REUNIONES");
    hoja.appendRow(["ID", "Tema", "Responsable", "Gerencia", "Area",
                    "MesesProgramados", "MesesEjecutados", "Anio"]);
  }
  return hoja;
}

// =============================================
// 1. INSPECCIONES - CON MEJORAS APLICADAS
// =============================================
function obtenerDatosPASSOInspecciones() {
  try {
    var anio = new Date().getFullYear();
    var ssCheck = getCheckSpreadsheet();
    var hojaInv = ssCheck.getSheetByName("INVENTARIO");
    if (!hojaInv) return { actividades: [], error: "Hoja INVENTARIO no encontrada" };

    var lastRowInv = hojaInv.getLastRow();
    if (lastRowInv < 3) return { actividades: [], debug: "Solo " + lastRowInv + " filas en INVENTARIO" };

    var lastColInv = Math.max(hojaInv.getLastColumn(), 17);
    var invData = hojaInv.getRange(3, 1, lastRowInv - 2, lastColInv).getValues();

    var equipos = [];
    var debugInfo = [];

    for (var i = 0; i < invData.length; i++) {

      // Col P (index 15) = estado
      var estado = (invData[i].length > 15) ? _normalizarPASSO(invData[i][15]) : "";
      if (estado === "retirado") continue;

      // 🔥 PUNTO 4: Normalizar nombre equipo desde INVENTARIO
      var nombreEquipo = String(invData[i][3] || "").trim().toUpperCase();   // Col D
      var codigo       = String(invData[i][6] || "").trim();                 // Col G
      var area         = String(invData[i][2] || "").trim();                 // Col C
      var frecTexto    = String(invData[i][12] || "").trim();                // Col M
      var codigoInterno = String(invData[i][4] || "").trim();

      // 🔥 PUNTO 1: Leer Col L y filtrar números puros con isNaN + normalizar a mayúsculas
      var lugaresRaw = String(invData[i][11] || "").trim();                  // Col L
      var lugares = lugaresRaw
        ? lugaresRaw.split(",")
            .map(function(l) { return l.trim().toUpperCase(); })
            .filter(function(l) { return l !== "" && isNaN(l); })
        : [];

      if (i < 5) debugInfo.push(
        "fila" + (i + 3) + ": equipo=\"" + nombreEquipo +
        "\" frec=\"" + frecTexto +
        "\" lugares=" + lugares.length
      );

      var frecInfo = _parsearFrecuenciaTexto(frecTexto);

      // Fallback frecuencia
      if (!frecInfo) {
        var frecNum = invData[i][13];
        if (typeof frecNum === "number" && frecNum > 0) {
          frecInfo = _parsearFrecuenciaTexto(String(frecNum));
        } else {
          var parsed = parseInt(String(frecNum).replace(/[^\d]/g, ""), 10);
          if (parsed > 0) frecInfo = _parsearFrecuenciaTexto(String(parsed));
        }
      }

      if (!frecInfo) continue;
      if (!nombreEquipo) continue;

      // Col Q (index 16) = fecha de inicio de inspección para cálculo exacto
      var fechaInicioRaw = (invData[i].length > 16) ? invData[i][16] : null;
      if (fechaInicioRaw instanceof Date) {
        frecInfo.fechaInicio = fechaInicioRaw;
      }

      equipos.push({
        gerencia:    "GERENCIA DE OPERACIONES",
        area:        area || "SIN ÁREA",
        nombre:      nombreEquipo,   // 🔥 ya en mayúsculas
        codigo:      codigo,
        codigoInterno: codigoInterno,
        responsable: "VICTOR CARACELA",
        frecuencia:  frecInfo.label,
        frecClase:   frecInfo.clase,
        frecInfo:    frecInfo,
        lugares:     lugares         // 🔥 ya en mayúsculas, sin números
      });
    }

    // =============================================
    // 🔥 LEER B DATOS — cruzar por equipo + lugar + mes
    // PUNTO 2: Normalizar lugar desde B DATOS
    // PUNTO 3: Normalizar nombre equipo desde B DATOS
    // PUNTO 6: Clave siempre en mayúsculas
    // =============================================
    var hojaBD = ssCheck.getSheetByName("B DATOS");
    var ejecPorEquipoLugarMes = {};

    if (hojaBD && hojaBD.getLastRow() > 1) {
      var bdLastCol = Math.max(hojaBD.getLastColumn(), 10);
      var bdData = hojaBD.getRange(2, 1, hojaBD.getLastRow() - 1, bdLastCol).getValues();

      for (var k = 0; k < bdData.length; k++) {
        // 🔥 PUNTO 3: Normalizar nombre equipo en B DATOS
        var eqNombre = String(bdData[k][2] || "").trim().toUpperCase();   // Col C

        // 🔥 PUNTO 2: Normalizar lugar en B DATOS
        var lugar    = String(bdData[k][7] || "").trim().toUpperCase();   // Col H

        var fecha    = bdData[k][9];                                       // Col J

        if (!eqNombre || !(fecha instanceof Date)) continue;
        if (fecha.getFullYear() !== anio) continue;

        // 🔥 PUNTO 6: Clave triple siempre en mayúsculas
        var clave = eqNombre + "||" + lugar + "||" + fecha.getMonth();
        ejecPorEquipoLugarMes[clave] = (ejecPorEquipoLugarMes[clave] || 0) + 1;
      }
    }

    // =============================================
    // 🔥 PUNTO 5: CONSTRUIR ACTIVIDADES
    // Una sola fila por equipo → programado × cantidad de lugares
    // =============================================
    var actividades = [];

    equipos.forEach(function(eq) {
      var meses = [];

      for (var m = 0; m < 12; m++) {
        var frecBase   = _calcularProgramadoMes(eq.frecInfo, m, anio);
        var programado = 0;
        var ejecutado  = 0;

        if (eq.lugares.length > 0) {
          // 🔥 Programado = frecuencia base × cantidad de lugares
          programado = frecBase * eq.lugares.length;

          // 🔥 Ejecutado = suma real por cada lugar (cruce exacto)
          eq.lugares.forEach(function(lugar) {
            var clave = eq.nombre + "||" + lugar + "||" + m;
            ejecutado += ejecPorEquipoLugarMes[clave] || 0;
          });

        } else {
          // Sin lugares: comportamiento normal
          programado = frecBase;

          Object.keys(ejecPorEquipoLugarMes).forEach(function(k) {
            var partes = k.split("||");
            if (partes[0] === eq.nombre && parseInt(partes[2]) === m) {
              ejecutado += ejecPorEquipoLugarMes[k];
            }
          });
        }

        meses.push({ programado: programado, ejecutado: ejecutado });
      }

      actividades.push({
        gerencia:    eq.gerencia,
        area:        eq.area,
        nombre:      eq.nombre,
        codigo:      eq.codigo,
        codigoInterno: eq.codigoInterno, 
        lugares:     eq.lugares.length > 0 ? eq.lugares.join(", ") : null, // 🔥 todos en un campo
        responsable: eq.responsable,
        frecuencia:  eq.frecuencia,
        frecClase:   eq.frecClase,
        meses:       meses
      });
    });

    var actividadesAgrupadas = agruparInspeccionesParaVista(actividades);  // ← LLAMADA

    return {
      actividades: actividadesAgrupadas,   // ← agrupado, no el original
      anio:        anio,
      debug:       debugInfo.join(" | ")
    };

  } catch (error) {
    Logger.log("Error PASSO Inspecciones: " + error.message + " | " + error.stack);
    return { actividades: [], error: error.message };
  }
}


// =============================================
// 2. CAPACITACIONES / ENTRENAMIENTOS
// =============================================
function _convertirFrecuenciaPASSO(dias) {
  dias = parseInt(dias);
  if (isNaN(dias) || dias <= 0) return "—";
  if (dias >= 360) return "ANUAL";
  if (dias >= 180) return "SEMESTRAL";
  if (dias >= 90)  return "TRIMESTRAL";
  if (dias >= 30)  return "MENSUAL";
  if (dias >= 14)  return "QUINCENAL";
  return dias + " DÍAS";
}

function getPassoDesdeMatriz() {
  var ssCap = getSpreadsheetCapacitaciones();
  var hojaMatriz = ssCap.getSheetByName("Matriz");
  var hojaBD     = ssCap.getSheetByName("B DATOS");

  // 🔥 DETECTAR ÚLTIMA COLUMNA CON DATOS EN FILA 17 (nombres de cursos)
  // Estructura Matriz: filas 2-16 = propiedades, fila 17 = nombres de cursos, fila 18+ = cargos
  // Fila 10 = Recurrente (nueva)
  var fila17 = hojaMatriz.getRange(17, 1, 1, hojaMatriz.getLastColumn()).getValues()[0];
  var lastCol = 0;
  for (var i = fila17.length - 1; i >= 0; i--) {
    if (fila17[i] && String(fila17[i]).trim() !== "") {
      lastCol = i + 1;
      break;
    }
  }
  if (lastCol < 5) lastCol = hojaMatriz.getLastColumn();

  Logger.log("🔥 Última columna detectada: " + lastCol);

  // Fila 2  = Tipo de Programa
  // Fila 4  = Gerencia
  // Fila 5  = Área
  // Fila 6  = Responsable
  // Fila 8  = Programación del curso (fechas)
  // Fila 9  = Vigencia (Meses)
  // Fila 10 = Recurrente (nueva)
  // Fila 17 = Nombres de cursos/temas
  var tipos             = hojaMatriz.getRange(2,  5, 1, lastCol - 4).getValues()[0];
  var responsables      = hojaMatriz.getRange(6,  5, 1, lastCol - 4).getValues()[0];
  var gerencias         = hojaMatriz.getRange(4,  5, 1, lastCol - 4).getValues()[0];
  var fechasProgramadas = hojaMatriz.getRange(8,  5, 1, lastCol - 4).getValues()[0];
  var areas             = hojaMatriz.getRange(5,  5, 1, lastCol - 4).getValues()[0];
  var vigencias         = hojaMatriz.getRange(9,  5, 1, lastCol - 4).getValues()[0];
  var esRecurrentes     = hojaMatriz.getRange(10, 5, 1, lastCol - 4).getValues()[0];
  var capacitaciones    = hojaMatriz.getRange(17, 5, 1, lastCol - 4).getValues()[0];

  Logger.log("🔥 Total capacitaciones leídas: " +
    capacitaciones.filter(function(c) { return c; }).length);

  // Leer B DATOS con 16 columnas para acceder a cols O (14) y P (15)
  var bdDatos = hojaBD.getLastRow() > 1
    ? hojaBD.getRange(2, 1, hojaBD.getLastRow() - 1, 16).getValues()
    : [];

  // Construir mapa de ciclos activos desde filas ACTIVACION
  // { temaNormalizado: [{ inicio: Date, fin: Date, modulo: Number }] }
  var activacionesMap = {};
  bdDatos.forEach(function(row) {
    // Fila de control: formato nuevo (col J = "ACTIVACION") o anterior (col A = "ACTIVACION")
    var esActivacion = String(row[9]).trim().toUpperCase() === "ACTIVACION" ||
                       String(row[0]).trim().toUpperCase() === "ACTIVACION";
    if (!esActivacion) return;
    var temaNorm = String(row[4]).trim().toUpperCase();
    var cicloStr = String(row[15] || "");
    var partes   = cicloStr.split(" - ");
    if (partes.length !== 2) return;
    var parseD = function(s) {
      var m = s.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      return m ? new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) : null;
    };
    var ini = parseD(partes[0]);
    var fin = parseD(partes[1]);
    if (!ini || !fin) return;
    if (!activacionesMap[temaNorm]) activacionesMap[temaNorm] = [];
    activacionesMap[temaNorm].push({ inicio: ini, fin: fin, modulo: parseInt(row[14]) || 0 });
  });

  var anioActual = new Date().getFullYear();
  var capacitacionesMap = {};

  // 🔥 LEER CAPACITACIONES — meses programados calculados desde vigencia (como inspecciones)
  // Vigencia (meses) en Matriz fila 9 → determina cada cuántos meses se repite el curso
  // Ej: vigencia=12 → anual (1 mes/año), vigencia=6 → semestral (2 meses), vigencia=2 → bimestral (6 meses)
  capacitaciones.forEach(function(cap, i) {
    if (!cap || String(cap).trim() === "") return;

    var capNormalizado  = String(cap).trim().toUpperCase();
    var tipo            = tipos[i] || "SIN CLASIFICAR";
    var fechaProgramada = fechasProgramadas[i];
    var vigMeses        = parseInt(vigencias[i]) || 12; // meses de vigencia/frecuencia

    // Normalizar fecha programada
    if (typeof fechaProgramada === 'string' && fechaProgramada.trim() !== '') {
      var mf = fechaProgramada.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}:\d{2}))?/);
      if (mf) {
        fechaProgramada = new Date(mf[3] + '-' + (mf[2].length < 2 ? '0' : '') + mf[2] + '-' + (mf[1].length < 2 ? '0' : '') + mf[1] + 'T' + (mf[4] || '00:00'));
      }
    }

    // Sin fecha válida en Matriz fila 8 → mesInicio = -1 (sin meses programados)
    var mesInicio = -1;
    if (fechaProgramada instanceof Date && !isNaN(fechaProgramada.getTime())) {
      mesInicio = fechaProgramada.getMonth();
    }

    if (!capacitacionesMap[capNormalizado]) {
      var rec = esRecurrentes[i];
      capacitacionesMap[capNormalizado] = {
        tipo:             tipo,
        gerencia:         gerencias[i] || "",
        area:             areas[i] || "",
        capacitacion:     cap,
        responsable:      responsables[i] || "",
        frecuencia:       "",
        mesesProgramados: [],
        esRecurrente:     (rec === true || String(rec).toUpperCase() === "VERDADERO"),
        vigenciaMeses:    vigMeses,
        fechaInicioCiclo: fechaProgramada
      };

      // Calcular meses programados: solo si hay fecha programada en la Matriz (fila 8).
      // Sin fecha → mesesProgramados queda [] (sin "P" en ningún mes).
      var meses = capacitacionesMap[capNormalizado].mesesProgramados;
      if (mesInicio >= 0) {
        // Hacia adelante desde mesInicio
        for (var m = mesInicio; m < 12; m += vigMeses) {
          if (meses.indexOf(m) === -1) meses.push(m);
        }
        // Hacia atrás desde mesInicio (meses anteriores del año en curso)
        for (var m = mesInicio - vigMeses; m >= 0; m -= vigMeses) {
          if (meses.indexOf(m) === -1) meses.push(m);
        }
        meses.sort(function(a, b) { return a - b; });
      }

      // Frecuencia legible
      var labels = {1:"MENSUAL",2:"BIMESTRAL",3:"TRIMESTRAL",4:"CUATRIMESTRAL",6:"SEMESTRAL",12:"ANUAL"};
      capacitacionesMap[capNormalizado].frecuencia = labels[vigMeses] || ("CADA " + vigMeses + " MESES");
    }
  });

  // ── Mapa de DNIs activos: solo estos cuentan en numerador Y denominador ──
  var dnisActivos = {};          // { dni: true } para lookup O(1)
  var totalTrabajadores = 0;
  try {
    var hojaPersonal = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    var lastRowPers  = hojaPersonal.getLastRow();
    if (lastRowPers > 1) {
      // Col B (índice 1) = DNI, Col L (índice 11) = estado activo
      var personalData = hojaPersonal.getRange(2, 1, lastRowPers - 1, 12).getValues();
      for (var pi = 0; pi < personalData.length; pi++) {
        var est = String(personalData[pi][11] || '').trim().toUpperCase(); // Col L
        var dni = String(personalData[pi][1]  || '').trim();               // Col B
        if (!dni) continue;
        // Activo = solo si el estado ES 'SI' o 'ACTIVO' (excluye CESADO, RETIRADO, INACTIVO, BAJA, etc.)
        if (est !== 'SI' && est !== 'ACTIVO') continue;
        dnisActivos[dni] = true;
        totalTrabajadores++;
      }
    }
  } catch(eP) {
    Logger.log('No se pudo leer PERSONAL para trabajadores activos: ' + eP.message);
  }
  if (totalTrabajadores < 1) totalTrabajadores = 1; // evitar división por cero
  Logger.log('Trabajadores activos para % capacitaciones: ' + totalTrabajadores);

  var agrupado = {};

  // 🔥 CALCULAR CUMPLIMIENTO DESDE B DATOS — porcentaje real de asistentes
  Object.keys(capacitacionesMap).forEach(function(capNormalizado) {
    var info = capacitacionesMap[capNormalizado];
    var tipo = info.tipo;

    if (!agrupado[tipo]) agrupado[tipo] = [];

    var meses = new Array(12).fill(0);

    Logger.log("🔥 " + info.capacitacion +
      " | Meses programados: " + info.mesesProgramados.join(",") +
      " | Frecuencia: " + info.frecuencia);

    // Marcar meses programados como -1 (P)
    info.mesesProgramados.forEach(function(mesIdx) {
      meses[mesIdx] = -1;
    });

    // Contar DNIs únicos APROBADOS por mes → calcular % real
    info.mesesProgramados.forEach(function(mesIdx) {
      var dnisAprobados = {};

      // Para cursos recurrentes: obtener ciclo activo desde activacionesMap
      var inicioCiclo = null;
      var finCiclo    = null;
      if (info.esRecurrente) {
        var hoyLocal  = new Date();
        var ciclosRec = activacionesMap[capNormalizado] || [];
        var cicloActivo = null;
        for (var ci = 0; ci < ciclosRec.length; ci++) {
          if (hoyLocal >= ciclosRec[ci].inicio && hoyLocal <= ciclosRec[ci].fin) {
            cicloActivo = ciclosRec[ci];
            break;
          }
        }
        if (cicloActivo) {
          inicioCiclo = cicloActivo.inicio;
          finCiclo    = cicloActivo.fin;
        }
      }

      bdDatos.forEach(function(row) {
        var dniBD    = String(row[0] || "").trim();
        var estadoBD = String(row[9] || "").trim().toUpperCase();
        // Saltar filas de control (ambos formatos: col A o col J)
        if (estadoBD === "ACTIVACION" || dniBD.toUpperCase() === "ACTIVACION") return;
        var temaBD   = String(row[4] || "").trim().toUpperCase();
        // Col O (index 14) = origen: solo contar "Matriz" para capacitaciones programadas
        var origenBD = String(row[14] || "").trim().toUpperCase();
        if (origenBD !== "MATRIZ") return;
        var fecha    = row[7];

        if (dniBD === "" || !dnisActivos[dniBD] || temaBD !== capNormalizado || estadoBD !== "APROBADO" || !(fecha instanceof Date)) return;

        var enPeriodo;
        if (info.esRecurrente && inicioCiclo && finCiclo) {
          enPeriodo = fecha >= inicioCiclo && fecha <= finCiclo;
        } else {
          enPeriodo = fecha.getMonth() === mesIdx && fecha.getFullYear() === anioActual;
        }

        if (enPeriodo) dnisAprobados[dniBD] = true;
      });

      var totalAprobados = Object.keys(dnisAprobados).length;
      if (totalAprobados > 0) {
        var pct = Math.round((totalAprobados / totalTrabajadores) * 100);
        meses[mesIdx] = Math.min(pct, 100); // 100 = "E", <100 = porcentaje
      }
    });

    agrupado[tipo].push({
      gerencia:     info.gerencia,
      area:         info.area,
      capacitacion: info.capacitacion,
      responsable:  info.responsable,
      frecuencia:   info.frecuencia,
      meses:        meses
    });
  });

  // ── Actividades NO programadas: B DATOS con col O = "Sin programar" ──
  // Se diferencia por la columna O (index 14), NO por si el nombre existe en la Matriz.
  // Un mismo curso puede tener registros "Matriz" y "Sin programar".
  var noProgramadasMap = {};
  bdDatos.forEach(function(row) {
    var dniBD    = String(row[0] || "").trim();
    var estadoBD = String(row[9] || "").trim().toUpperCase();
    if (estadoBD === "ACTIVACION" || dniBD.toUpperCase() === "ACTIVACION") return;
    if (!dniBD || !dnisActivos[dniBD]) return;
    if (estadoBD !== "APROBADO") return;

    var temaBD = String(row[4] || "").trim();
    if (!temaBD) return;
    var temaNorm = temaBD.toUpperCase();
    // Col O (index 14): solo incluir "Sin programar"
    var origenBD = String(row[14] || "").trim().toUpperCase();
    if (origenBD !== "SIN PROGRAMAR") return;

    var fecha = row[7];
    if (!(fecha instanceof Date) || isNaN(fecha.getTime())) return;
    var mes = fecha.getMonth();

    if (!noProgramadasMap[temaNorm]) {
      noProgramadasMap[temaNorm] = {
        capacitacion: temaBD,
        area:         String(row[5] || ""),
        gerencia:     "",
        responsable:  String(row[11] || ""),
        frecuencia:   "NO PROGRAMADO",
        meses:        new Array(12).fill(0),
        _dnisMes:     {}
      };
    }
    var entry = noProgramadasMap[temaNorm];
    if (!entry._dnisMes[mes]) entry._dnisMes[mes] = {};
    entry._dnisMes[mes][dniBD] = true;
  });

  // Calcular porcentajes por mes para cada actividad no programada
  var actividadesNoProgramadas = Object.keys(noProgramadasMap).map(function(k) {
    var entry = noProgramadasMap[k];
    Object.keys(entry._dnisMes).forEach(function(mes) {
      var count = Object.keys(entry._dnisMes[mes]).length;
      entry.meses[parseInt(mes)] = Math.min(Math.round((count / totalTrabajadores) * 100), 100);
    });
    delete entry._dnisMes;
    return entry;
  });

  return { agrupado: agrupado, actividadesNoProgramadas: actividadesNoProgramadas };
}


// =============================================
// 3. REUNIONES - CRUD
// =============================================
function obtenerReuniones() {
  try {
    var hoja    = _getHojaReuniones();
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { reuniones: [] };

    var datos    = hoja.getRange(2, 1, lastRow - 1, 8).getValues();
    var reuniones = datos.map(function(r) {
      return {
        id:               r[0],
        tema:             r[1],
        responsable:      r[2],
        gerencia:         r[3],
        area:             r[4],
        mesesProgramados: String(r[5] || ""),
        mesesEjecutados:  String(r[6] || ""),
        anio:             r[7]
      };
    });

    return { reuniones: reuniones };
  } catch (error) {
    Logger.log("Error en obtenerReuniones: " + error.message);
    return { reuniones: [], error: error.message };
  }
}

function agregarReunion(data) {
  try {
    var hoja = _getHojaReuniones();
    var id   = "R" + Date.now();
    var fila = [
      id,
      data[0] || "",
      data[1] || "",
      data[2] || "",
      data[3] || "",
      data[4] || "",
      "",
      data[5] || new Date().getFullYear()
    ];
    hoja.appendRow(fila);
    return { success: true, id: id };
  } catch (error) {
    Logger.log("Error en agregarReunion: " + error.message);
    return { success: false, error: error.message };
  }
}

function actualizarReunion(data) {
  try {
    var hoja    = _getHojaReuniones();
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { success: false, error: "No hay reuniones" };

    var ids       = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    var idBuscado = String(data[0]).trim();

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === idBuscado) {
        var fila = [data[0], data[1], data[2], data[3],
                    data[4], data[5], data[6], data[7]];
        hoja.getRange(i + 2, 1, 1, 8).setValues([fila]);
        return { success: true };
      }
    }
    return { success: false, error: "Reunión no encontrada" };
  } catch (error) {
    Logger.log("Error en actualizarReunion: " + error.message);
    return { success: false, error: error.message };
  }
}

function eliminarReunion(id) {
  try {
    var hoja    = _getHojaReuniones();
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { success: false, error: "No hay reuniones" };

    var ids       = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    var idBuscado = String(id).trim();

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === idBuscado) {
        hoja.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, error: "Reunión no encontrada" };
  } catch (error) {
    Logger.log("Error en eliminarReunion: " + error.message);
    return { success: false, error: error.message };
  }
}

function marcarReunionEjecutada(id, mesEjecutado) {
  try {
    var hoja    = _getHojaReuniones();
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { success: false, error: "No hay reuniones" };

    var datos     = hoja.getRange(2, 1, lastRow - 1, 8).getValues();
    var idBuscado = String(id).trim();

    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === idBuscado) {
        var mesStr             = String(mesEjecutado);
        var ejecutadosActuales = String(datos[i][6] || "")
          .split(",")
          .map(function(s) { return s.trim(); })
          .filter(function(s) { return s !== ""; });

        var idx = ejecutadosActuales.indexOf(mesStr);
        if (idx >= 0) {
          ejecutadosActuales.splice(idx, 1);
        } else {
          ejecutadosActuales.push(mesStr);
        }

        var nuevosEjecutados = ejecutadosActuales
          .sort(function(a, b) { return Number(a) - Number(b); })
          .join(",");

        hoja.getRange(i + 2, 7).setValue(nuevosEjecutados);
        return { success: true, mesesEjecutados: nuevosEjecutados };
      }
    }
    return { success: false, error: "Reunión no encontrada" };
  } catch (error) {
    Logger.log("Error en marcarReunionEjecutada: " + error.message);
    return { success: false, error: error.message };
  }
}


// =============================================
// 4. DATOS COMBINADOS PASSO
// =============================================
function obtenerDatosPASSOCompleto() {
  try {
    var inspecciones  = obtenerDatosPASSOInspecciones();
    var reunionesData = obtenerReuniones();

    var passoResult = { agrupado: {}, actividadesNoProgramadas: [] };
    var capError    = null;

    try {
      passoResult = getPassoDesdeMatriz();
    } catch (e) {
      capError = e.message;
      Logger.log("Error getPassoDesdeMatriz: " + e.message + " | " + e.stack);
    }

    var capacitaciones         = [];
    var entrenamientos         = [];
    var actividadesNoProgramadas = passoResult.actividadesNoProgramadas || [];

    // Cursos de la Matriz → pestaña Capacitaciones
    Object.keys(passoResult.agrupado || {}).forEach(function(tipo) {
      capacitaciones = capacitaciones.concat(passoResult.agrupado[tipo]);
    });

    var anio = new Date().getFullYear();

    var reunionesActividades = (reunionesData.reuniones || [])
      .filter(function(r) { return Number(r.anio) === anio || !r.anio; })
      .map(function(r) {

        var mesesProg = String(r.mesesProgramados || "")
          .split(",")
          .map(function(s) { return s.trim(); })
          .filter(function(s) { return s !== ""; });

        var mesesEjec = String(r.mesesEjecutados || "")
          .split(",")
          .map(function(s) { return s.trim(); })
          .filter(function(s) { return s !== ""; });

        var meses = [];
        for (var m = 0; m < 12; m++) {
          meses.push({
            programado: mesesProg.indexOf(String(m + 1)) >= 0,
            ejecutado:  mesesEjec.indexOf(String(m + 1)) >= 0
          });
        }

        return {
          id:               r.id,
          nombre:           r.tema,
          responsable:      r.responsable,
          gerencia:         r.gerencia,
          area:             r.area,
          mesesProgramados: r.mesesProgramados,
          mesesEjecutados:  r.mesesEjecutados,
          meses:            meses
        };
      });

    return {
      inspecciones:             inspecciones.actividades || [],
      capacitaciones:           capacitaciones,
      actividadesNoProgramadas: actividadesNoProgramadas,
      reuniones:                reunionesActividades,
      reunionesRaw:             reunionesData.reuniones || [],
      entrenamientos:           entrenamientos,
      anio:                     anio,
      debug: {
        inspCount:  (inspecciones.actividades || []).length,
        capCount:   capacitaciones.length,
        reunCount:  reunionesActividades.length,
        entrCount:  entrenamientos.length,
        inspError:  inspecciones.error || null,
        capError:   capError,
        inspDebug:  inspecciones.debug || "",
        capDebug:   "Total desde Matriz: " + capacitaciones.length,
        entrDebug:  ""
      }
    };

  } catch (error) {
    Logger.log("Error en obtenerDatosPASSOCompleto: " + error.message);
    return {
      inspecciones:   [],
      capacitaciones: [],
      reuniones:      [],
      reunionesRaw:   [],
      entrenamientos: [],
      error:          error.message
    };
  }
}

function agruparInspeccionesParaVista(items) {
  var grupos = {};
  
  items.forEach(function(item) {
    // 🔑 Clave visual = nombre (Col D) + descripcion (Col G)
    var claveVista = (item.nombre || '') + '||' + (item.codigo || '');
    
    if (!grupos[claveVista]) {
      grupos[claveVista] = {
        nombre:      item.nombre,
        codigo:      item.codigo,      // Col G → descripción visible
        codigosInt:  [],               // Col E → IDs internos (ocultos)
        responsable: item.responsable,
        frecuencia:  item.frecuencia,
        frecClase:   item.frecClase,
        area:        item.area,
        gerencia:    item.gerencia,
        meses:       item.meses.map(function() { 
                       return { programado: 0, ejecutado: 0 }; 
                     })
      };
    }
    
    // ✅ Guardar código interno (Col E) sin mostrarlo
    grupos[claveVista].codigosInt.push(item.codigoInterno); // Col E
    
    // ✅ Acumular programado y ejecutado por mes
    item.meses.forEach(function(m, idx) {
      grupos[claveVista].meses[idx].programado += (m.programado || 0);
      grupos[claveVista].meses[idx].ejecutado  += (m.ejecutado  || 0);
    });
  });
  
  return Object.keys(grupos).map(function(k) { return grupos[k]; });
}