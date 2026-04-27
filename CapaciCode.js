//Inicio Capacitaciones                                                                                                                   
    //let ss = SpreadsheetApp.openById("1MyXsN09Jf23dcimniDLrDu3luDgasCk8EbLHOp3gzGw")                                                        
    const foldercharlas = "1IWmNW4wMZbC43QHrdivlRtAwnTHa3c2v"; //CARPETA DE CHARLAS                                            
     const foldefirmascap = "1GcIoeFFtpZ6EISt0w5R1byNkDy5I3ugi";  //CARPETA DE FIRMA CAPCITACIONES                                                                                          
    let cachedCapacitaciones = null;                                                                                                          
    function getSpreadsheetCapacitaciones() {                                                                                                 
      if (!cachedCapacitaciones) {                                                                                                            
        cachedCapacitaciones = SpreadsheetApp.openById("1Ev5_B3jMtjy_xXt13NYBXYwFA-maFAeLSKfiCFIsMQo"); //HOJA DE CALCULO CAPACITACIONES                                       
      }                                                                                                                                       
      return cachedCapacitaciones;                                                                                                            
    }                                                                                                                                         
                                                                                                                                              
// ─── Helper: parsea "DD/MM/YYYY - DD/MM/YYYY" → { inicio, fin } ────────────
function _parsearRangoCiclo(cicloStr) {
  const partes = String(cicloStr || "").split(" - ");
  if (partes.length !== 2) return null;
  const parseD = s => {
    const m = s.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    return m ? new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) : null;
  };
  const inicio = parseD(partes[0]);
  const fin    = parseD(partes[1]);
  return (inicio && fin) ? { inicio, fin } : null;
}

// ─── Helper: sumar N meses a una fecha sin cruzar fin de mes ───────────────
function _sumarMeses(fecha, meses) {
  const d = new Date(fecha);
  d.setMonth(d.getMonth() + meses);
  return d;
}

function cargarDatosGlobales() {
  const ssCap = getSpreadsheetCapacitaciones();
  const ssPersonal = getSpreadsheetPersonal();
  const hojaMatriz = ssCap.getSheetByName("Matriz");
  const hojaPersonal = ssPersonal.getSheetByName("PERSONAL");
  const hojaBDatos = ssCap.getSheetByName("B DATOS");

  if (!hojaMatriz)   throw new Error("No se encontró la hoja 'Matriz' en el spreadsheet de Capacitaciones.");
  if (!hojaPersonal) throw new Error("No se encontró la hoja 'PERSONAL' en el spreadsheet de Personal.");
  if (!hojaBDatos)   throw new Error("No se encontró la hoja 'B DATOS' en el spreadsheet de Capacitaciones.");

  const lastRowMatriz = hojaMatriz.getLastRow();
  const lastColMatriz = hojaMatriz.getLastColumn();
  if (lastRowMatriz < 1 || lastColMatriz < 1)
    throw new Error("La hoja 'Matriz' está vacía.");

  const matrizDatos = hojaMatriz.getRange(1, 1, lastRowMatriz, lastColMatriz).getValues();

  // Fila 10 (índice 9) → Recurrente (nueva); Fila 17 (índice 16) → Temas; filas 18+ → Cargos
  const cursos = (matrizDatos[16] || []).slice(4);
  const cargos = matrizDatos.slice(17).map(f => f[3]).filter(Boolean);
  const matriz = matrizDatos.slice(17).map(f => f.slice(4, 4 + cursos.length));

  const lastRowPersonal = hojaPersonal.getLastRow();
  const personalDatos = lastRowPersonal > 0
    ? hojaPersonal.getRange(1, 1, lastRowPersonal, 12).getValues()
    : [];
  const personalCargos = personalDatos.slice(1).map(f => f[6]);

  const lastRowBDatos = hojaBDatos.getLastRow();
  const bdDatos = lastRowBDatos > 0
    ? hojaBDatos.getRange(1, 1, lastRowBDatos, 16).getValues()
    : [];

  return { matrizDatos, cursos, cargos, matriz, personalDatos, personalCargos, bdDatos };
}
                                                                                                                                     
 function obtenerDatosPorDNI(dni) {
  const datos = cargarDatosGlobales();
  const { personalDatos, matrizDatos, bdDatos } = datos;

  // 🔹 Buscar persona por DNI
  let persona = null;
  for (let i = 1; i < personalDatos.length; i++) {
    if (personalDatos[i][1].toString().trim() === dni.toString().trim()) {
      persona = {
        nombre: personalDatos[i][2],
        empresa: personalDatos[i][4],
        cargo: personalDatos[i][6]
      };
      break;
    }
  }

  if (!persona) return { encontrado: false };

  // 🔹 Extraer filas horizontales de la matriz
  const tiposProg = matrizDatos[1].slice(4);                             // Fila 2  → Tipo de Programa
  const tiposCapacitacion = matrizDatos[2].slice(4);                     // Fila 3  → Tipo de Capacitación
  const gerencias = matrizDatos[3].slice(4);                             // Fila 4  → Gerencia
  const areas = matrizDatos[4].slice(4);                                 // Fila 5  → Área
  const responsables = matrizDatos[5].slice(4);                          // Fila 6  → Responsable
  const capacitadores = matrizDatos[6].slice(4);                         // Fila 7  → Capacitador
  const programaciones = matrizDatos[7].slice(4);                        // Fila 8  → Programación del curso
  const temporalidades = matrizDatos[8].slice(4);                        // Fila 9  → Vigencia (Meses)
  const esRecurrente = matrizDatos[9].slice(4);                           // Fila 10 → Recurrente (NUEVA)
  const duraciones = matrizDatos[10].slice(4).map(d => parseInt(d));    // Fila 11 → Duración (Minutos)
  const horasLectivas = matrizDatos[11].slice(4);                       // Fila 12 → Horas lectivas (Minutos)
  const puntajesMin = matrizDatos[12].slice(4).map(p => parseFloat(p)); // Fila 13 → Puntaje mínimo
  const imagen = matrizDatos[13].slice(4);                              // Fila 14 → Imagen
  const links = matrizDatos[14].slice(4);                               // Fila 15 → Link
  const tieneCertificacion = matrizDatos[15].slice(4);                  // Fila 16 → Certificación
  const cursos = matrizDatos[16].slice(4);                              // Fila 17 → Temas

  // 🔹 Buscar la fila del cargo
  const filaCursos = matrizDatos.find(f => f[3]?.toString().toLowerCase().trim() === persona.cargo.toLowerCase().trim());
  if (!filaCursos) return { encontrado: true, persona, cursos: [] };

  // 🔹 Indexar evaluaciones (excluye filas de control ACTIVACION)
  const evaluacionesMap = {};
  // 🔹 Indexar activaciones de ciclo: { temaNormalizado: [{ modulo, inicio, fin }] }
  const activacionesMap = {};
  for (let i = 1; i < bdDatos.length; i++) {
    const row = bdDatos[i];
    const dniBD  = String(row[0] || "").trim();
    const tema   = String(row[4] || "").trim();
    const colO   = String(row[14] || "").trim().toUpperCase(); // MODULO (col O)
    const colP   = String(row[15] || "").trim();               // CICLO  (col P)

    // Fila de control de ciclo: detectar formato nuevo (col J) y formato anterior (col A)
    const estadoBD = String(row[9] || "").trim().toUpperCase();
    const esActivacion = estadoBD === "ACTIVACION" || dniBD.toUpperCase() === "ACTIVACION";
    if (esActivacion) {
      const temaNorm = tema.toLowerCase();
      if (!activacionesMap[temaNorm]) activacionesMap[temaNorm] = [];
      const rango = _parsearRangoCiclo(colP);
      if (rango) activacionesMap[temaNorm].push({ modulo: parseInt(colO) || 0, inicio: rango.inicio, fin: rango.fin });
      continue;
    }
    // Fila de evaluación normal
    const puntaje = row[6];
    const fecha   = row[7];
    if (!evaluacionesMap[dniBD]) evaluacionesMap[dniBD] = {};
    if (!evaluacionesMap[dniBD][tema]) evaluacionesMap[dniBD][tema] = [];
    evaluacionesMap[dniBD][tema].push({ puntaje: parseFloat(puntaje), fecha: new Date(fecha) });
  }

  const hoy = new Date();
  const cursosEncontrados = [];

  for (let j = 4; j < filaCursos.length; j++) {
    if (filaCursos[j] !== true && filaCursos[j] !== "VERDADERO") continue;

    const temaCurso = String(cursos[j - 4] || '');
    const evaluaciones = evaluacionesMap[dni]?.[temaCurso] || [];

    const intentosHoy = evaluaciones.filter(ev => new Date(ev.fecha).toDateString() === hoy.toDateString()).length;

    // Mejor puntaje y última fecha
    let mejorPuntaje = null;
    let fechaEvaluacion = null;
    for (const ev of evaluaciones) {
      if (
        mejorPuntaje === null ||
        ev.puntaje > mejorPuntaje ||
        (ev.puntaje === mejorPuntaje && ev.fecha > fechaEvaluacion)
      ) {
        mejorPuntaje = ev.puntaje;
        fechaEvaluacion = ev.fecha;
      }
    }

    // Programación (debe calcularse ANTES del bloque Estado)
    const prog = programaciones[j - 4];
    let textoProgramacion = "";
    let fechaProg = null;

    if (prog instanceof Date && !isNaN(prog)) {
      fechaProg = prog;
    } else if (typeof prog === "string" && prog.trim() !== "") {
      const match = prog.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(\d{1,2}:\d{2}(?::\d{2})?)?/);
      if (match) {
        const [_, d, m, y, hms] = match;
        fechaProg = new Date(`${m}/${d}/${y} ${hms || "00:00"}`);
      }
    }

    if (fechaProg instanceof Date && !isNaN(fechaProg)) {
      const fin = new Date(fechaProg.getTime() + duraciones[j - 4] * 60000);
      if (hoy >= fechaProg && hoy <= fin) {
        const fechaOculta = Utilities.formatDate(fechaProg, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        textoProgramacion = `En vivo ((🔴))|${fechaOculta}`;
      } else if (hoy < fechaProg) {
        textoProgramacion = Utilities.formatDate(fechaProg, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      } else {
        textoProgramacion = "No previsto";
      }
    } else {
      textoProgramacion = prog ? prog.toString() : "";
    }

    // Estado — basado en ACTIVACION rows (tanto recurrentes como no recurrentes)
    // La fila 8 (Programación) define el mes de inicio del primer ciclo.
    // Recurrente=TRUE: ciclos de vigencia meses hasta 31/12.
    // Recurrente=FALSE: un solo ciclo desde el 1ro del mes hasta 31/12.
    let estadoCurso = "Pendiente";
    const esRec = esRecurrente[j - 4] === true || String(esRecurrente[j - 4]).toUpperCase() === "VERDADERO";
    const temaNorm = temaCurso.toLowerCase();

    const activaciones = activacionesMap[temaNorm] || [];
    if (activaciones.length === 0) {
      // Sin ciclos activados — usar lógica legacy para no romper datos anteriores
      if (mejorPuntaje !== null && fechaEvaluacion instanceof Date && !isNaN(fechaEvaluacion)) {
        const vencimiento = _sumarMeses(fechaEvaluacion, parseInt(temporalidades[j - 4]) || 0);
        estadoCurso = hoy > vencimiento
          ? "Caducado"
          : mejorPuntaje >= puntajesMin[j - 4] ? "Aprobado" : "Reprobado";
      }
    } else {
      const cicloActivo = activaciones.find(a => hoy >= a.inicio && hoy <= a.fin);
      if (!cicloActivo) {
        estadoCurso = "Ciclo Finalizado";
      } else {
        // Buscar evaluaciones del empleado dentro del ciclo activo
        const evsCiclo = evaluaciones.filter(ev => ev.fecha >= cicloActivo.inicio && ev.fecha <= cicloActivo.fin);
        if (evsCiclo.length > 0) {
          const mejor = evsCiclo.reduce((b, ev) => ev.puntaje > b.puntaje ? ev : b, evsCiclo[0]);
          estadoCurso = mejor.puntaje >= puntajesMin[j - 4] ? "Aprobado" : "Reprobado";
        }
        // else: Pendiente (hay ciclo activo pero no rindió)
      }
    }

    // 🔹 Agregar al resultado
    cursosEncontrados.push({
      tema: temaCurso,
      tipoProg: tiposProg[j - 4] || '',
      tipoCapacitacion: tiposCapacitacion[j - 4] || '',
      responsable: responsables[j - 4] || '',
      gerencia: gerencias[j - 4] || '',
      link: links[j - 4],
      programacion: textoProgramacion,
      image: imagen[j - 4],
      puntajeMinimo: puntajesMin[j - 4],
      duracion: duraciones[j - 4],
      horasLectivas: horasLectivas[j - 4],
      temporabilidad: temporalidades[j - 4],
      tieneCertificacion: (tieneCertificacion[j - 4] === true || String(tieneCertificacion[j - 4]).toUpperCase() === "VERDADERO"),
      puntaje: mejorPuntaje !== null ? mejorPuntaje : "-",
      area: areas[j - 4],
      capacitador: capacitadores[j - 4],
      estado: estadoCurso,
      intentosHoy,
      fecha: fechaEvaluacion
        ? Utilities.formatDate(fechaEvaluacion, Session.getScriptTimeZone(), "dd/MM/yyyy")
        : "-"
    });
  }

  return { encontrado: true, persona, cursos: cursosEncontrados };
}
                                                                                                                                         
                                                                                                                                              
    //QUIZZ                                                                                                                                   
    /** Nombre hojas */                                                                                                                       
    var quizData = "Examen"; //                                                                                                               
    var bd = "B DATOS"; //                                                                                                                    
                                                                                                                                              
    /** ******** Cuestionarios  ******** **/                                                                                                  
    function getDataQuestion(selectedValue) {                                                                                                 
      const sheet = getSpreadsheetCapacitaciones().getSheetByName(quizData);                                                                  
      const lastRow = sheet.getLastRow();                                                                                                     
      const data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues() // A2:K                                                           
        .filter(d => d[0] !== "" && d[2] === selectedValue); // d[2] = Tema                                                                   
                                                                                                                                              
      const maxQ = data.length;                                                                                                               
      const correctAnswer = [];                                                                                                               
      const pointValues = [];                                                                                                                 
                                                                                                                                              
      const radioLists = data.map((d, index) => {                                                                                             
        const pregunta = d[3];                                                                                                                
        const urlImagen = d[4];                                                                                                               
        const opciones = [d[5], d[6], d[7], d[8]];                                                                                            
        const correcta = d[9]; // valor 1-4                                                                                                   
        const puntos = parseFloat(d[10]) || 0;                                                                                                
        const id = index + 1;                                                                                                                 
                                                                                                                                              
        correctAnswer.push(correcta);                                                                                                         
        pointValues.push(puntos);                                                                                                             
                                                                                                                                              
        let imgHtml = urlImagen ? `<img class="img-fluid cat mt-2 mb-3" src="${urlImagen}" alt="imagen">` : "";                               
                                                                                                                                              
        return `                                                                                                                              
          <div id="${id}" class="fade-in-page mt-4" style="display:none">                                                                     
            <hr>                                                                                                                              
            <div class="row mt-2">                                                                                                            
              <label class="radio-label mt-2">                                                                                                
                <div><span class="inner-label1">${pregunta}</span></div>                                                                      
              </label>                                                                                                                        
              <div class="text-center"><span class="inner-label1">${imgHtml}</span></div>                                                     
                                                                                                                                              
              <label class="radio-label choice mt-4" style="display:none">                                                                    
                <input name="q${id}" type="radio" id="x${id}" value="0" checked>                                                              
                <span class="inner-label"></span>                                                                                             
              </label>                                                                                                                        
                                                                                                                                              
              ${opciones.map((op, i) => `                                                                                                     
                <label class="radio-label choice mt-4${i === 3 ? ' mb-2' : ''}">                                                              
                  <input name="q${id}" type="radio" id="${String.fromCharCode(97 + i)}${id}" value="${i + 1}">                                
                  <span class="inner-label">${op}</span>                                                                                      
                </label>                                                                                                                      
              `).join('')}                                                                                                                    
            </div>                                                                                                                            
          </div>                                                                                                                              
        `;                                                                                                                                    
      });                                                                                                                                     
                                                                                                                                              
      return [maxQ, correctAnswer, radioLists.join(""), pointValues];                                                                         
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
                                                                                                                                              
  function recordData(
  dni, nombre, cargo, empresa,
  tema, area, point, duracion,
  Ans, firmaBase64, capacitadorForm, duracionForm, detalle // 👈 nuevo parámetro
) {
  const ss = getSpreadsheetCapacitaciones();
  const hojaBD = ss.getSheetByName("B DATOS");
  const hojaMatriz = ss.getSheetByName("Matriz");
  const now = new Date();

  try {
    // === 1️⃣ Buscar datos del curso ===
    const numCursos = hojaMatriz.getLastColumn() - 4;
    const cursos = hojaMatriz.getRange(17, 5, 1, numCursos).getValues()[0];        // Fila 17 → Temas
    const puntajesMin = hojaMatriz.getRange(13, 5, 1, numCursos).getValues()[0];  // Fila 13 → Puntaje mínimo
    const temporalidades = hojaMatriz.getRange(9, 5, 1, numCursos).getValues()[0]; // Fila 9  → Vigencia (Meses)
    const capacitadores = hojaMatriz.getRange(7, 5, 1, numCursos).getValues()[0];  // Fila 7  → Capacitador
    const horasLectivas = hojaMatriz.getRange(12, 5, 1, numCursos).getValues()[0]; // Fila 12 → Horas lectivas
    const esRecurrentes = hojaMatriz.getRange(10, 5, 1, numCursos).getValues()[0]; // Fila 10 → Recurrente
    const programaciones = hojaMatriz.getRange(8, 5, 1, numCursos).getValues()[0]; // Fila 8  → Programación
    const duracionesMin = hojaMatriz.getRange(11, 5, 1, numCursos).getValues()[0]; // Fila 11 → Duración (Min)

    let puntajeMinimo = null;
    let temporalidad = null;
    let capacitador = "";
    let horasLectiva = "";
    let duracionMin = "";
    let encontradoEnMatriz = false;

    for (let j = 0; j < cursos.length; j++) {
      if (String(cursos[j]).toLowerCase().trim() === String(tema).toLowerCase().trim()) {
        puntajeMinimo = parseFloat(puntajesMin[j]);
        temporalidad = parseInt(temporalidades[j]);
        capacitador = String(capacitadores[j] || "");
        horasLectiva = horasLectivas[j] || "";
        duracionMin = duracionesMin[j] || "";
        encontradoEnMatriz = true;
        break;
      }
    }

    if (!capacitador && capacitadorForm) capacitador = capacitadorForm;
    if (!horasLectiva && duracionForm) horasLectiva = duracionForm;

    // Limpiar horasLectiva: guardar solo el número (ej: "80 min" → 80)
    if (horasLectiva !== "" && horasLectiva !== null && horasLectiva !== undefined) {
      const numHL = parseFloat(String(horasLectiva).replace(/[^0-9.]/g, ""));
      if (!isNaN(numHL)) horasLectiva = numHL;
    }

    // Si el tema no está en la Matriz, buscar puntajeMinimo en hoja TEMAS (col M = índice 12)
    let horaInicioTema = null;
    if (!encontradoEnMatriz) {
      const hojaTemas = ss.getSheetByName("TEMAS");
      const lastRowTemas = hojaTemas.getLastRow();
      if (lastRowTemas > 1) {
        const temasDatos = hojaTemas.getRange(2, 1, lastRowTemas - 1, 13).getValues();
        const temaNormBusq = String(tema).toLowerCase().trim();
        for (let t = 0; t < temasDatos.length; t++) {
          if (String(temasDatos[t][1]).toLowerCase().trim() === temaNormBusq) {
            const pm = parseFloat(temasDatos[t][12]);
            if (!isNaN(pm)) puntajeMinimo = pm;
            horaInicioTema = temasDatos[t][8] || null;
            if (!duracionMin) duracionMin = String(temasDatos[t][4] || "");
            break;
          }
        }
      }
    }

    // === 1b. Leer MODULO y CICLO desde fila ACTIVACION activa en B DATOS ===
    // Si el tema NO está en la Matriz → actividad no programada → "SIN PROGRAMAR"
    let modulo = encontradoEnMatriz ? "MATRIZ" : "SIN PROGRAMAR";
    let ciclo  = "";
    let fechaCicloInicio = null;
    const lastRowBD = hojaBD.getLastRow();
    if (lastRowBD > 1) {
      const bdAll = hojaBD.getRange(2, 1, lastRowBD - 1, 16).getValues();
      const temaNorm = String(tema).trim().toLowerCase();
      for (const row of bdAll) {
        // Detectar fila de control: formato nuevo (col J) o anterior (col A)
        const esAct = String(row[9]).trim().toUpperCase() === "ACTIVACION" ||
                      String(row[0]).trim().toUpperCase() === "ACTIVACION";
        if (!esAct) continue;
        if (String(row[4]).trim().toLowerCase() !== temaNorm) continue;
        const rango = _parsearRangoCiclo(String(row[15] || ""));
        if (rango && now >= rango.inicio && now <= rango.fin) {
          ciclo = "Ciclo " + (parseInt(row[14]) || 1);
          fechaCicloInicio = rango.inicio;
          break;
        }
      }
    }

    // === 2️⃣ Determinar estado ===
    let estadoFinal = "Pendiente";
    if (point === null || point === undefined || point === "") {
      // Sin examen — registro directo como Aprobado
      estadoFinal = "Aprobado";
    } else if (!isNaN(point) && point !== "-") {
      if (puntajeMinimo !== null && !isNaN(puntajeMinimo)) {
        estadoFinal = point >= puntajeMinimo ? "Aprobado" : "Reprobado";
      } else {
        // Sin puntaje mínimo configurado → cualquier puntaje aprueba
        estadoFinal = "Aprobado";
      }
    }

    // Fecha de habilitación del curso (fecha en que se activó el ciclo, no la del examen)
    let fechaHabilitacion = "";
    if (fechaCicloInicio instanceof Date) {
      fechaHabilitacion = fechaCicloInicio;
    } else if (horaInicioTema) {
      const d = horaInicioTema instanceof Date ? horaInicioTema : new Date(horaInicioTema);
      if (!isNaN(d.getTime())) fechaHabilitacion = d;
    }

    // === 3️⃣ Subir firma si existe ===
    let urlFirma = "";
    if (firmaBase64 && firmaBase64.startsWith("data:image/")) {
      const folder = DriveApp.getFolderById(foldefirmascap);
      const nombreArchivo = `${dni}_firma_${Date.now()}.png`;
      const blob = Utilities.newBlob(
        Utilities.base64Decode(firmaBase64.split(",")[1]),
        "image/png",
        nombreArchivo
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFirma = `https://lh5.googleusercontent.com/d/${file.getId()}`;
    }

    // === 4️⃣ Registrar datos ===
    const filaNueva = [
      "'" + dni,
      nombre,
      cargo,
      empresa,
      tema,
      area,
      point,
      now,
      horasLectiva || "",
      estadoFinal,
      temporalidad || "",
      capacitador || "",
      detalle || "",   // columna M → Comentarios
      urlFirma || "",  // columna N → Firma
      modulo || "",    // columna O → Módulo (ciclo N° para recurrente)
      ciclo || ""      // columna P → Ciclo (rango de fechas del ciclo)
    ].concat(Ans || []);

    hojaBD.appendRow(filaNueva);

    return ["success", point, estadoFinal];
  } catch (error) {
    Logger.log("Error en recordData: " + error);
    return ["error", error.toString()];
  }
}

function _registrarEnRegistroFirmas(ss, fecha, tema, dni, nombre, cargo, empresa, area, capacitador, duracion, comentarios, firmaUrl, codigoRegistro) {
  try {
    let hoja = ss.getSheetByName("REGISTRO FIRMAS");
    if (!hoja) {
      hoja = ss.insertSheet("REGISTRO FIRMAS");
      const encabezados = ["Fecha", "Tema", "DNI", "Nombre", "Cargo", "Empresa", "Área", "Capacitador", "Duración (min)", "Comentarios", "Firma URL", "Código Registro"];
      hoja.appendRow(encabezados);
      const headerRange = hoja.getRange(1, 1, 1, encabezados.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1a73e8");
      headerRange.setFontColor("#ffffff");
      hoja.setFrozenRows(1);
      hoja.setColumnWidth(1, 110);
      hoja.setColumnWidth(2, 200);
      hoja.setColumnWidth(3, 100);
      hoja.setColumnWidth(4, 200);
      hoja.setColumnWidth(5, 150);
      hoja.setColumnWidth(6, 150);
      hoja.setColumnWidth(7, 130);
      hoja.setColumnWidth(8, 180);
      hoja.setColumnWidth(9, 110);
      hoja.setColumnWidth(10, 200);
      hoja.setColumnWidth(11, 300);
      hoja.setColumnWidth(12, 120);
    }
    hoja.appendRow([
      fecha || "",
      tema || "",
      "'" + String(dni).replace(/^'/, ""),
      nombre || "",
      cargo || "",
      empresa || "",
      area || "",
      capacitador || "",
      duracion || "",
      comentarios || "",
      firmaUrl || "",
      codigoRegistro || ""
    ]);
  } catch (e) {
    Logger.log("Error en _registrarEnRegistroFirmas: " + e);
  }
}

/**
 * Registra asistencia SOLO en REGISTRO FIRMAS (sin B DATOS).
 * Usado por el flujo Registro de Firma (código + firma, sin examen).
 */
function registrarSoloFirma(dni, nombre, cargo, empresa, tema, area, firmaBase64, capacitador, duracion, comentarios, codigoRegistro) {
  try {
    const ss = getSpreadsheetCapacitaciones();
    const now = new Date();

    // Subir firma a Drive
    let urlFirma = "";
    if (firmaBase64 && firmaBase64.startsWith("data:image/")) {
      const folder = DriveApp.getFolderById(foldefirmascap);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(firmaBase64.split(",")[1]),
        "image/png",
        `${dni}_firma_${Date.now()}.png`
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFirma = `https://lh5.googleusercontent.com/d/${file.getId()}`;
    }

    _registrarEnRegistroFirmas(ss, now, tema, String(dni).replace(/^'/, ""), nombre, cargo, empresa, area, capacitador, duracion || "", comentarios || "", urlFirma, codigoRegistro || "");

    return ["success"];
  } catch (e) {
    Logger.log("Error en registrarSoloFirma: " + e);
    return ["error", e.toString()];
  }
}

    //ESTRELLAS                                                                                                                               
    function guardarCalificacionEnFila(dni, calificacion) {                                                                                   
      const sheet = getSpreadsheetCapacitaciones().getSheetByName("B DATOS");                                                                 
      const lastRow = sheet.getLastRow();                                                                                                     
                                                                                                                                              
      // Solo leemos la columna A (donde está el DNI)                                                                                         
      const dniCol = sheet.getRange(1, 1, lastRow, 1).getValues(); // Columna A                                                               
                                                                                                                                              
      for (let i = lastRow - 1; i >= 0; i--) {                                                                                                
        if (dniCol[i][0] == dni) {                                                                                                            
          sheet.getRange(i + 1, 13).setValue(calificacion); // Columna M (13)                                                                 
          return true;                                                                                                                        
        }                                                                                                                                     
      }                                                                                                                                       
      return false;                                                                                                                           
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    function getAllTopics() {                                                                                                                 
      const sheet = getSpreadsheetCapacitaciones().getSheetByName(quizData);                                                                  
      const data = sheet.getRange(2, 3, sheet.getLastRow()-1).getValues(); // Suponiendo que los temas están en la columna C                  
      const uniqueTopics = [...new Set(data.flat())].filter(String);                                                                          
      return uniqueTopics;                                                                                                                    
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
                                                                                                                                              
    //Configuracion matriz                                                                                                                    
function obtenerMatrizInvertida() {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  if (!hoja) return null;

  const lastCol = hoja.getLastColumn();
  // leemos 16 filas horizontales (D2..D17: 15 datos + 1 índice)
  const rango = hoja.getRange(2, 4, 16, lastCol - 3).getValues(); // D2 en adelante, 16 filas

  const headers = rango.map(fila => fila[0]);
  const datos = rango.map(fila => fila.slice(1));
  const indice = datos.pop(); // la última fila del bloque será el índice (temas)

  const filas = datos[0].map((_, i) => {
    const id = indice[i];
    if (!id) return null;

    const fila = datos.map(f => f[i]);
    fila.unshift(id);

    return fila.map(celda =>
      celda instanceof Date
        ? Utilities.formatDate(celda, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
        : celda
    );
  }).filter(Boolean);

  headers.pop();
  headers.unshift("Curso");

  return { headers, filas };
}
                                                                                                                                     
                                                                                                                                              
function agregarRegistro(nuevaColumna) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");

  // 16 valores: 1 nombre del curso + 15 propiedades (filas 2..16)
  if (nuevaColumna.length !== 16) throw new Error("Se requieren 16 valores: 1 curso + 15 datos");

  const datos = nuevaColumna.slice(1); // Las 15 propiedades (para filas 2..16)
  const curso = nuevaColumna[0];       // El nombre del curso

  // ── UPSERT: buscar si el curso ya existe en fila 17 (desde col E = col 5) ──
  const lastColActual = hoja.getLastColumn();
  let col = -1; // columna donde escribir
  let esNuevo = true;

  if (lastColActual >= 5) {
    const row17 = hoja.getRange(17, 5, 1, lastColActual - 4).getValues()[0];
    const idx = row17.findIndex(function(c) {
      return String(c || '').trim().toLowerCase() === String(curso).trim().toLowerCase();
    });
    if (idx !== -1) {
      col = 5 + idx; // columna real del curso existente → PISAR
      esNuevo = false;
    }
  }

  if (col === -1) {
    col = lastColActual + 1; // curso nuevo → agregar al final
  }

  // Escribe filas 2 a 16 en la columna correspondiente
  hoja.getRange(2, col, 15, 1).setValues(datos.map(d => [d]));

  // Escribe el índice del curso en la fila 17
  hoja.getRange(17, col).setValue(curso);

  // Aplica validación de checkbox en fila 10 (Recurrente) y fila 16 (Certificación)
  const checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  hoja.getRange(10, col).setDataValidation(checkboxValidation);
  hoja.getRange(16, col).setDataValidation(checkboxValidation);

  // Aplica validación de checkbox en filas 18+ (cargos) — solo si es nueva columna
  if (esNuevo) {
    const lastRow = hoja.getLastRow();
    if (lastRow >= 18) {
      hoja.getRange(18, col, lastRow - 17, 1).setDataValidation(checkboxValidation);
    }
  }

  // Auto-generar ciclos anuales solo para cursos NUEVOS
  // (los existentes ya tienen sus ciclos en B DATOS)
  if (esNuevo) {
    let fechaInicioRaw = nuevaColumna[7];
    let fechaInicio = null;
    if (fechaInicioRaw instanceof Date && !isNaN(fechaInicioRaw)) {
      fechaInicio = fechaInicioRaw;
    } else if (typeof fechaInicioRaw === "string" && fechaInicioRaw.trim()) {
      const mf = fechaInicioRaw.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (mf) fechaInicio = new Date(parseInt(mf[1]), parseInt(mf[2]) - 1, parseInt(mf[3]));
    }
    generarCiclosAnuales(curso, fechaInicio);
  }
}

// ─── Reparar cursos sin ciclos activados ─────────────────────────────────────
// Recorre la Matriz, y para cada curso que NO tiene fila ACTIVACION en B DATOS,
// genera todos los ciclos anuales usando la fecha de programación.
// Aplica a cursos recurrentes Y no recurrentes.
function repararCiclosFaltantes() {
  const ss    = getSpreadsheetCapacitaciones();
  const hoja  = ss.getSheetByName("Matriz");
  const hojaBD = ss.getSheetByName("B DATOS");
  if (!hoja || !hojaBD) return { reparados: 0, mensaje: "Hoja no encontrada" };

  const lastCol = hoja.getLastColumn();
  if (lastCol < 5) return { reparados: 0, mensaje: "Sin cursos en Matriz" };

  // Leer headers de la Matriz (columnas desde E en adelante, índice 1 = col E)
  const progRow   = hoja.getRange(8,  5, 1, lastCol - 4).getValues()[0]; // fila 8: programación
  const cursosRow = hoja.getRange(17, 5, 1, lastCol - 4).getValues()[0]; // fila 17: nombres

  // Leer B DATOS para ver qué cursos ya tienen ACTIVACION
  const lastRowBD = hojaBD.getLastRow();
  const cursosConCiclo = new Set();
  if (lastRowBD > 1) {
    const bdAll = hojaBD.getRange(2, 1, lastRowBD - 1, 16).getValues();
    bdAll.forEach(function(row) {
      const col9 = String(row[9] || "").trim().toUpperCase();
      const col0 = String(row[0] || "").trim().toUpperCase();
      if (col9 === "ACTIVACION" || col0 === "ACTIVACION") {
        cursosConCiclo.add(String(row[4]).trim().toLowerCase());
      }
    });
  }

  let reparados = 0;
  for (let i = 0; i < cursosRow.length; i++) {
    const nombreCurso = String(cursosRow[i] || "").trim();
    if (!nombreCurso) continue;
    if (cursosConCiclo.has(nombreCurso.toLowerCase())) continue; // ya tiene ciclos

    // Parsear fecha de programación
    let fechaInicio = null;
    const raw = progRow[i];
    if (raw instanceof Date && !isNaN(raw)) {
      fechaInicio = new Date(raw);
    } else if (typeof raw === "string" && raw.trim()) {
      const m = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) fechaInicio = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
    }
    generarCiclosAnuales(nombreCurso, fechaInicio);
    reparados++;
  }
  return { reparados, mensaje: reparados + " curso(s) reparados" };
}

// ─── Generar TODOS los ciclos anuales para un curso ──────────────────────────
// La fila 8 (Programación) define el MES en que empieza el primer ciclo.
// • Recurrente=TRUE  → ciclos consecutivos de "vigencia" meses, cortados al 31/12.
// • Recurrente=FALSE → un solo ciclo: 1ro del mes de programación → 31/12.
function generarCiclosAnuales(nombreCurso, fechaProgramacion) {
  const ssCap  = getSpreadsheetCapacitaciones();
  const hojaBD = ssCap.getSheetByName("B DATOS");
  const hojaM  = ssCap.getSheetByName("Matriz");

  // Leer datos del curso desde Matriz
  const lastCol   = hojaM.getLastColumn();
  if (lastCol < 5) throw new Error("Sin cursos en Matriz");
  const cursosRow = hojaM.getRange(17, 5, 1, lastCol - 4).getValues()[0];
  const vigRow    = hojaM.getRange(9,  5, 1, lastCol - 4).getValues()[0];
  const recRow    = hojaM.getRange(10, 5, 1, lastCol - 4).getValues()[0];

  const idx = cursosRow.findIndex(c => String(c).trim().toLowerCase() === String(nombreCurso).trim().toLowerCase());
  if (idx < 0) throw new Error("Curso no encontrado en Matriz: " + nombreCurso);

  const vigMeses = parseInt(vigRow[idx]) || 1;
  const esRec    = recRow[idx] === true || String(recRow[idx]).toUpperCase() === "VERDADERO";

  // Parsear fecha de programación
  let fp = fechaProgramacion;
  if (typeof fp === "string" && fp.trim()) {
    const m = fp.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) fp = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  }
  if (!(fp instanceof Date) || isNaN(fp)) fp = new Date();

  const anio     = fp.getFullYear();
  const mesInicio = fp.getMonth(); // 0-based
  const finAnio  = new Date(anio, 11, 31, 23, 59, 59); // 31 de diciembre
  const tz       = Session.getScriptTimeZone();
  const ciclosCreados = [];

  if (!esRec) {
    // ── NO RECURRENTE (anual): un solo ciclo → 1ro del mes hasta 31/12 ──
    const inicio = new Date(anio, mesInicio, 1);
    const cicloStr = Utilities.formatDate(inicio, tz, "dd/MM/yyyy") + " - " +
                     Utilities.formatDate(finAnio, tz, "dd/MM/yyyy");
    const fila = new Array(16).fill("");
    fila[4]  = nombreCurso;
    fila[7]  = new Date();
    fila[9]  = "ACTIVACION";
    fila[14] = 1;
    fila[15] = cicloStr;
    hojaBD.appendRow(fila);
    ciclosCreados.push(1);
  } else {
    // ── RECURRENTE: ciclos de vigMeses, desde el mes de programación ─────
    let numCiclo  = 0;
    let mesActual = mesInicio;

    while (true) {
      const inicio = new Date(anio, mesActual, 1);
      if (inicio > finAnio) break; // ya pasó diciembre

      numCiclo++;

      // Fin teórico del ciclo: inicio + vigencia meses − 1 día
      let fin = _sumarMeses(inicio, vigMeses);
      fin.setDate(fin.getDate() - 1);
      fin.setHours(23, 59, 59, 0);

      // Cortar en 31/12 si excede el año
      if (fin > finAnio) fin = new Date(finAnio);

      const cicloStr = Utilities.formatDate(inicio, tz, "dd/MM/yyyy") + " - " +
                       Utilities.formatDate(fin, tz, "dd/MM/yyyy");
      const fila = new Array(16).fill("");
      fila[4]  = nombreCurso;
      fila[7]  = new Date();
      fila[9]  = "ACTIVACION";
      fila[14] = numCiclo;
      fila[15] = cicloStr;
      hojaBD.appendRow(fila);
      ciclosCreados.push(numCiclo);

      mesActual += vigMeses;
    }
  }

  return ciclosCreados;
}

                                                                                                                                     
function actualizarRegistro(columnaEditada) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  // los cursos están en la fila 17, desde columna E (col 5)
  const datos = hoja.getRange(17, 5, 1, hoja.getLastColumn() - 4).getValues()[0]; // fila 17, desde E

  const curso = columnaEditada[0];
  const colIndex = datos.indexOf(curso);

  if (colIndex === -1) throw new Error("Curso no encontrado");

  const datosNuevos = columnaEditada.slice(1); // sin el nombre del curso
  if (datosNuevos.length !== 15) throw new Error("Se requieren 15 valores para actualizar (filas 2..16)");

  const col = 5 + colIndex; // columna real donde está el curso

  // Actualiza filas 2 a 16 (15 filas)
  hoja.getRange(2, col, 15, 1).setValues(datosNuevos.map(d => [d]));

  // Actualiza el nombre del curso en fila 17
  hoja.getRange(17, col).setValue(curso);
}
                                                                                                                                
function eliminarRegistroPorCurso(nombreCurso) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  if (!hoja) return;

  // D2 en adelante, altura = 16 filas (2..17)
  const rango = hoja.getRange(2, 4, 16, hoja.getLastColumn() - 3); // D2 en adelante
  const datos = rango.getValues();

  // Buscar la columna del índice (última fila del bloque = fila 17)
  const filaIndice = datos[datos.length - 1]; // esto apunta a la fila que contiene los nombres de curso (fila 17)
  const colIndex = filaIndice.indexOf(nombreCurso);

  if (colIndex === -1) return;

  // Borrar verticalmente todos los valores en esa columna (filas 2..15)
  for (let fila = 0; fila < datos.length; fila++) {
    hoja.getRange(fila + 2, 4 + colIndex).setValue(""); // columna D + colIndex
  }
}
                                                                                                                                  
    //EXAMEN                                                                                                                                  
    function obtenerMatrizExamenPaginado(offset, limit, filtro = "") {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    return { headers2: [], filas: [], total: 0 };
  }

  // Leer columnas A–K (11 columnas)
  const datos = hoja.getRange(1, 1, lastRow, 11).getValues();
  const [headers2, ...filas] = datos;

  let filtradas = filas;
  if (filtro) {
    const texto = filtro.toLowerCase();
    filtradas = filas.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(texto))
    );
  }

  const paginadas = filtradas.slice(offset, offset + limit);
  return {
    headers2,
    filas: paginadas,
    total: filtradas.length
  };
}

/**
 * 🔹 Recupera todas las preguntas pertenecientes a un tema específico.
 */
function obtenerPreguntasPorTema(tema) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];

  const data = hoja.getRange(2, 1, lastRow - 1, 11).getValues();
  return data
    .filter(r => r[2] === tema)
    .map(r => ({
      id: r[0],
      pregunta: r[3],
      url: r[4],
      opciones: [r[5], r[6], r[7], r[8]],
      correcta: r[9],
      puntos: r[10]
    }));
}

/**
 * 🔹 Crea o actualiza múltiples preguntas de un mismo tema.
 * Si es edición: elimina físicamente las que ya no están en el modal.
 */
function guardarPreguntasMultiples(lista, idsOriginales, esEdicion) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  const dataExistente = lastRow > 1
    ? hoja.getRange(2, 1, lastRow - 1, 11).getValues()
    : [];

  if (esEdicion && idsOriginales?.length) {
    // 🔸 Identificar IDs que ya no existen en la nueva lista y eliminarlas físicamente
    const idsEliminar = idsOriginales.filter(id => !lista.some(p => p[0] === id));
    if (idsEliminar.length > 0) {
      // Buscar sus posiciones de fila (de abajo hacia arriba para no desajustar índices)
      const filasEliminar = [];
      dataExistente.forEach((r, i) => {
        if (idsEliminar.includes(r[0])) filasEliminar.push(i + 2);
      });
      filasEliminar.sort((a, b) => b - a).forEach(fila => hoja.deleteRow(fila));
    }
  }

  // 🔸 Actualizar o insertar cada pregunta
  lista.forEach(data => {
    if (!data[0]) {
      // Nueva pregunta → generar ID único
      const idUnico = "E" + Date.now().toString().slice(-7) + Math.floor(Math.random() * 100);
      data[0] = idUnico;
      hoja.appendRow(data);
    } else {
      // Buscar y actualizar si existe
      const index = dataExistente.findIndex(r => r[0] === data[0]);
      if (index >= 0) {
        hoja.getRange(index + 2, 1, 1, data.length).setValues([data]);
      } else {
        hoja.appendRow(data);
      }
    }
  });

  return true;
}

/**
 * 🔹 Elimina una pregunta por su ID.
 */
function eliminarPreguntaPorID(id) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) {
      hoja.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}
                                                                                                                                
function obtenerOpcionesFormulario() {
  const ss = getSpreadsheetCapacitaciones();
  const hojaMatriz = ss.getSheetByName("Matriz");
  const hojaTemas = ss.getSheetByName("TEMAS");
  
  // --- 1. PROCESAR MATRIZ (Horizontal: Fila 5=Área, 17=Temas) ---
  const lastCol = hojaMatriz.getLastColumn();
  // Leemos desde la columna E (5) hasta el final
  const filaAreas = hojaMatriz.getRange(5, 5, 1, lastCol - 4).getValues()[0];
  const filaTemas = hojaMatriz.getRange(17, 5, 1, lastCol - 4).getValues()[0];

  let datosMatrizRelacion = [];
  let areasUnicasMatriz = [];

  filaAreas.forEach((area, i) => {
    const tema = filaTemas[i];
    if (area && tema) {
      const areaLimpia = area.toString().trim();
      const temaLimpio = tema.toString().trim();
      
      // Guardamos la relación para filtrar después
      datosMatrizRelacion.push([temaLimpio, areaLimpia]);
      
      // Lista para el primer select
      if (!areasUnicasMatriz.includes(areaLimpia)) {
        areasUnicasMatriz.push(areaLimpia);
      }
    }
  });

  // --- 2. PROCESAR HOJA TEMAS (Vertical) ---
  const lastRowTemas = hojaTemas.getLastRow();
  let datosTemasSheet = [];
  if (lastRowTemas > 1) {
    // Col B (Temas), Col C (Área)
    datosTemasSheet = hojaTemas.getRange(2, 2, lastRowTemas - 1, 2).getValues()
      .map(f => [f[0].toString().trim(), f[1].toString().trim()]);
  }

  return { 
    opciones1: areasUnicasMatriz.sort(), // Áreas de la Matriz
    datosMatrizSheet: datosMatrizRelacion, // Nueva relación [Tema, Área] de Matriz
    datosTemasSheet: datosTemasSheet       // Relación [Tema, Área] de hoja TEMAS
  };
}
                                                                                                                                     
    //MATRIZ CHECK                                                                                                                            
    function parseFecha(fechaStr) {                                                                                                           
      if (fechaStr instanceof Date) return fechaStr;                                                                                          
                                                                                                                                              
      if (typeof fechaStr === "string") {                                                                                                     
        const partes = fechaStr.split('/');                                                                                                   
        if (partes.length < 3) return new Date('');                                                                                           
        const [dia, mes, anioHora] = partes;                                                                                                  
        const [anio, hora] = anioHora.split(' ');                                                                                             
        return new Date(`${mes}/${dia}/${anio} ${hora || '00:00:00'}`);                                                                       
      }                                                                                                                                       
                                                                                                                                              
      return new Date('');                                                                                                                    
    }                                                                                                                                         
                                                                                                                                              
    function obtenerMatrizPermisos() {
  const datos = cargarDatosGlobales();

  // Mapear cursos no vacíos conservando su posición original (para calcular colSheet correctamente)
  const cursosConIdx = datos.cursos
    .map((c, i) => ({ curso: String(c).trim(), colIdx: i }))
    .filter(({ curso }) => curso !== '');
  const cursos = cursosConIdx.map(({ curso }) => curso);
  const colIndices = cursosConIdx.map(({ colIdx }) => colIdx); // índice original en la hoja (col E = colIdx 0)

  const cargos = datos.cargos.filter(String);
  // Reconstruir la matriz usando solo las columnas de cursos no vacíos
  const matriz = datos.matriz.map(row => colIndices.map(i => row[i]));
  const personalCargos = datos.personalCargos;
  const bdDatos = datos.bdDatos.slice(1);
  bdDatos.shift();

  const idxDNI = 0;             // Columna A
  const idxCargo = 2;           // Columna C
  const idxTema = 4;            // Columna E
  const idxPuntaje = 6;         // Columna G
  const idxFecha = 7;           // Columna H
  const idxHoras = 8;           // Columna I (Carga Horaria en minutos)
  const idxEstatus = 9;         // Columna J
  const idxTemporalidad = 10;   // Columna K

  const ahora = new Date();
  const mejoresPorDNIyCurso = {};

  bdDatos.forEach(row => {
    const dni = row[idxDNI];
    const curso = row[idxTema];
    const estatus = row[idxEstatus];
    const puntaje = parseFloat(row[idxPuntaje]) || 0;
    const fechaEval = parseFecha(row[idxFecha]);
    const dias = parseInt(row[idxTemporalidad]) || 0;

    if (estatus !== "Aprobado") return;

    const fechaLimite = new Date(fechaEval);
    fechaLimite.setDate(fechaEval.getDate() + dias);
    if (ahora > fechaLimite) return;

    const clave = `${dni}_${curso}`;
    if (!mejoresPorDNIyCurso[clave] || puntaje > (parseFloat(mejoresPorDNIyCurso[clave][idxPuntaje]) || 0)) {
      mejoresPorDNIyCurso[clave] = row;
    }
  });

  const datosFiltrados = Object.values(mejoresPorDNIyCurso);

  // === Cálculo de resumen por curso ===
  const resumen = cursos.map(curso => {
    const colIndex = cursos.indexOf(curso);
    const cargosAsignados = cargos.filter((_, i) => matriz[i][colIndex] === true);
    const personasAsignadas = personalCargos.filter(cargo => cargosAsignados.includes(cargo)).length;
    const personasAprobadas = datosFiltrados.filter(row =>
      cargosAsignados.includes(row[idxCargo]) && row[idxTema] === curso
    ).length;

    return { curso, asignados: personasAsignadas, aprobados: personasAprobadas };
  });

  // === Totales globales ===
  const totalProgramados = resumen.reduce((sum, item) => sum + item.asignados, 0);
  const totalAprobados = resumen.reduce((sum, item) => sum + item.aprobados, 0);

  // === Nuevos cálculos ===
  const totalCursosProgramados = cursos.length;
  const totalCursosRealizados = resumen.filter(r => r.aprobados > 0).length;
  // Trabajadores activos: DNI no vacío (col B) y sin valor en col L (liquidado)
  // Solo cuenta trabajadores con CONDICIÓN = ACTIVO (col L, índice 11)
  const totalTrabajadores = datos.personalDatos.slice(1).filter(r =>
    r[1] && String(r[1]).trim() !== '' &&
    String(r[11] || '').trim().toUpperCase() === 'ACTIVO'
  ).length;
  const totalHoras = datosFiltrados.reduce((sum, r) => sum + (parseFloat(r[idxHoras]) || 0), 0) / 60; // a horas

  return {
    cursos,
    cargos,
    matriz,
    colIndices,  // posiciones reales de columna (0-based desde col E) para calcular colSheet
    resumen,
    totales: {
      cursosProgramados: totalCursosProgramados,
      cursosRealizados: totalCursosRealizados,
      trabajadores: totalTrabajadores,
      programados: totalProgramados,
      aprobados: totalAprobados,
      horas: totalHoras.toFixed(1) // redondear a 1 decimal
    }
  };
}


                                                                                                                                    
 function obtenerDetalleCurso(curso) {
  const ssCap = getSpreadsheetCapacitaciones();
  const hojaBDatos = ssCap.getSheetByName("B DATOS");
  const hojaMatriz = ssCap.getSheetByName("Matriz");
  const lastRow = hojaBDatos.getLastRow();
  const ahora = new Date();

  // 🔹 Leer solo columnas A–K (1–11) desde fila 2
  const datos = hojaBDatos.getRange(2, 1, lastRow - 1, 11).getValues();

  // 🔹 Leer cursos (fila 17), temporalidades (fila 9), certificación (fila 16)
  const cursos = hojaMatriz.getRange("E17:17").getValues()[0];
  const temporalidades = hojaMatriz.getRange("E9:9").getValues()[0];
  const certificaciones = hojaMatriz.getRange("E16:16").getValues()[0];

  // 🔹 Ubicar el índice del curso actual
  const colIndex = cursos.indexOf(curso);
  const tieneCertificacion = certificaciones[colIndex] === true || certificaciones[colIndex] === "VERDADERO";

  // Índices de columnas en B DATOS
  const idxDni = 0;             // Columna A
  const idxNombre = 1;          // Columna B
  const idxCargo = 2;           // Columna C
  const idxEmpresa = 3;         // Columna D
  const idxTema = 4;            // Columna E
  const idxPuntaje = 6;         // Columna G
  const idxFecha = 7;           // Columna H
  const idxEstatus = 9;         // Columna J
  const idxTemporalidad = 10;   // Columna K

  const mejoresPorDNI = {};

  // 🔹 Filtrar los mejores registros aprobados y vigentes
  datos.forEach(row => {
    if (row[idxTema] !== curso) return;
    if (row[idxEstatus] !== "Aprobado") return;

    const dni = row[idxDni];
    const puntaje = parseFloat(row[idxPuntaje]) || 0;
    const fechaEval = parseFecha(row[idxFecha]);
    const dias = parseInt(row[idxTemporalidad]) || 0;

    const fechaLimite = new Date(fechaEval);
    fechaLimite.setDate(fechaEval.getDate() + dias);
    if (ahora > fechaLimite) return;

    if (!mejoresPorDNI[dni] || puntaje > (parseFloat(mejoresPorDNI[dni][idxPuntaje]) || 0)) {
      mejoresPorDNI[dni] = row;
    }
  });

  // 🔹 Convertir los resultados en un arreglo de objetos formateados
  return Object.values(mejoresPorDNI).map(row => {
    let fechaCruda = row[idxFecha];
    let fechaFormateada = "";

    if (fechaCruda instanceof Date) {
      const dia = fechaCruda.getDate().toString().padStart(2, '0');
      const mes = (fechaCruda.getMonth() + 1).toString().padStart(2, '0');
      const anio = fechaCruda.getFullYear();
      fechaFormateada = `${dia}/${mes}/${anio}`;
    } else if (typeof fechaCruda === "string") {
      fechaFormateada = fechaCruda.split(" ")[0];
    }

    return {
      dni: row[idxDni],
      nombre: row[idxNombre],
      cargo: row[idxCargo],
      empresa: row[idxEmpresa],
      puntaje: row[idxPuntaje],
      estatus: row[idxEstatus],
      fecha: fechaFormateada,
      certificacion: tieneCertificacion ? "Con certificación" : "Curso sin certificación"
    };
  });
}
                                                                                        
                                                                                                                                              
    function actualizarCeldaCheckbox(fila, columna, valor) {                                                                                  
      const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");                                                                   
      hoja.getRange(fila, columna).setValue(valor);                                                                                           
    }                                                                                                                                         
                                                                                                                                              
    function actualizarAsignadosPorCurso(fila, columna, valor) {
  const ssCap = getSpreadsheetCapacitaciones();
  const hojaMatriz = ssCap.getSheetByName("Matriz");
  const hojaBDatos = ssCap.getSheetByName("B DATOS");
  const hojaPersonal = getSpreadsheetPersonal().getSheetByName("PERSONAL");

  // 🔹 Obtener todos los cursos (fila 17)
  const cursos = hojaMatriz.getRange("E17:17").getValues()[0];
  const lastRowMatriz = hojaMatriz.getLastRow();

  // 🔹 Calcular índice de columna dentro del rango de cursos
  const colIndex = columna - 5;
  if (colIndex < 0 || colIndex >= cursos.length) return;

  const curso = cursos[colIndex];

  // ✅ Actualizar el valor del checkbox
  hojaMatriz.getRange(fila, columna).setValue(valor);
  SpreadsheetApp.flush();

  // ✅ Leer la matriz desde fila 18 en adelante (cargos verticales)
  const cargosRange = hojaMatriz.getRange(18, 4, lastRowMatriz - 17, cursos.length + 1).getValues();
  const cargos = cargosRange.map(row => row[0]);
  const matriz = cargosRange.map(row => row.slice(1));

  // ✅ Cargos que tienen el curso marcado como asignado
  const cargosAsignados = new Set(
    cargos.filter((_, i) => matriz[i][colIndex] === true)
  );

  // ✅ Personal → Columna G contiene el cargo
  const personalCargos = hojaPersonal.getRange(2, 7, hojaPersonal.getLastRow() - 1).getValues().flat();
  const personasAsignadas = personalCargos.filter(cargo => cargosAsignados.has(cargo)).length;

  // ✅ Leer B DATOS (A–K)
  const lastRowBD = hojaBDatos.getLastRow();
  const datos = hojaBDatos.getRange(2, 1, lastRowBD - 1, 11).getValues();

  const ahora = new Date();
  const mejoresPorDNIyCurso = {};

  // ✅ Filtrar los mejores resultados por persona y curso
  for (const row of datos) {
    const [dni, , cargo, , tema, , puntajeStr, fechaStr, , estatus, diasStr] = row;

    if (tema !== curso || estatus !== "Aprobado") continue;

    const puntaje = parseFloat(puntajeStr) || 0;
    const fechaEval = parseFecha(fechaStr);
    const dias = parseInt(diasStr) || 0;

    const fechaLimite = new Date(fechaEval);
    fechaLimite.setDate(fechaEval.getDate() + dias);
    if (ahora > fechaLimite) continue;

    const clave = `${dni}_${tema}`;
    if (!mejoresPorDNIyCurso[clave] || puntaje > (parseFloat(mejoresPorDNIyCurso[clave][6]) || 0)) {
      mejoresPorDNIyCurso[clave] = row;
    }
  }

  const datosFiltrados = Object.values(mejoresPorDNIyCurso);

  // ✅ Contar aprobados entre los cargos asignados
  const personasAprobadas = datosFiltrados.filter(row =>
    cargosAsignados.has(row[2]) && row[4] === curso
  ).length;

  // ✅ Nueva salida con curso y valores actualizados
  return {
    curso,
    asignados: personasAsignadas,
    aprobados: personasAprobadas
  };
}
                                                                                                                               
                                                                           
    //BUSCADOR GENERAL                                                                                                                        
    function getDatosDesdeBDATOS(start = 0, length = 30, search = "", fechaDesde = "", fechaHasta = "") {                                     
      const sheetBDATOS = getSpreadsheetCapacitaciones().getSheetByName("B DATOS");                                                           
      const sheetMatriz = getSpreadsheetCapacitaciones().getSheetByName("Matriz");                                                            
      if (!sheetBDATOS || !sheetMatriz) throw new Error('Faltan hojas');                                                                      
                                                                                                                                              
      const data = sheetBDATOS.getRange("A1:M" + sheetBDATOS.getLastRow()).getValues();                                                       
      const headers = data[0];                                                                                                                
      headers.push("Aprobados / Previstos");                                                                                                  
      const hoy = new Date();                                                                                                                 
                                                                                                                                              
      const temasPorID = {};                                                                                                                  
      const cargoPorID = {};                                                                                                                  
                                                                                                                                              
      for (let i = 1; i < data.length; i++) {                                                                                                 
        const [id, , cargo, , tema, , , fecha, , estado, temporalidad] = data[i];                                                             
        if (!cargoPorID[id]) cargoPorID[id] = cargo;                                                                                          
                                                                                                                                              
        if (estado !== "Aprobado" || !(fecha instanceof Date) || isNaN(temporalidad)) continue;                                               
                                                                                                                                              
        const dias = Math.floor((hoy - fecha) / (1000 * 60 * 60 * 24));                                                                       
        if (dias >= 0 && dias <= parseInt(temporalidad)) {                                                                                    
          if (!temasPorID[id]) temasPorID[id] = new Set();                                                                                    
          temasPorID[id].add(tema);                                                                                                           
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      const cargosMatriz = sheetMatriz.getRange("D18:D" + sheetMatriz.getLastRow()).getValues().flat().filter(c => c);
      const filaTemas = sheetMatriz.getRange("E17:17").getValues()[0];
      const cantidadTemas = filaTemas.filter(t => t).length;
      const matrizValores = sheetMatriz.getRange(18, 5, cargosMatriz.length, cantidadTemas).getValues();                                      
                                                                                                                                              
      const previstosPorCargo = {};                                                                                                           
      cargosMatriz.forEach((cargo, idx) => {                                                                                                  
        const fila = matrizValores[idx];                                                                                                      
        const verdaderos = fila.filter(val => val === true).length;                                                                           
        previstosPorCargo[cargo] = verdaderos;                                                                                                
      });                                                                                                                                     
                                                                                                                                              
      let filas = [];                                                                                                                         
                                                                                                                                              
      for (let i = 1; i < data.length; i++) {                                                                                                 
        const fila = [...data[i]];                                                                                                            
        const id = fila[0];                                                                                                                   
        const estado = fila[9];                                                                                                               
        const fecha = fila[7];                                                                                                                
                                                                                                                                              
        if (fecha instanceof Date) {                                                                                                          
          fila[7] = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");                                          
          fila._fechaOrden = fecha;                                                                                                           
        } else {                                                                                                                              
          fila._fechaOrden = new Date("1900-01-01");                                                                                          
        }                                                                                                                                     
                                                                                                                                              
        if (estado === "Aprobado") {                                                                                                          
          const aprobados = temasPorID[id] ? temasPorID[id].size : 0;                                                                         
          const cargo = cargoPorID[id] || "";                                                                                                 
          const previstos = previstosPorCargo[cargo] || 0;                                                                                    
          fila.push(`${aprobados}/${previstos}`);                                                                                             
          filas.push(fila);                                                                                                                   
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      // Filtro por texto                                                                                                                     
      if (search) {                                                                                                                           
        const query = search.toLowerCase();                                                                                                   
        filas = filas.filter(row => row.some(cell => String(cell).toLowerCase().includes(query)));                                            
      }                                                                                                                                       
                                                                                                                                              
      // Filtro por fecha (columna 7 ya formateada)                                                                                           
      if (fechaDesde || fechaHasta) {                                                                                                         
        const desde = fechaDesde ? new Date(fechaDesde) : null;                                                                               
        const hasta = fechaHasta ? new Date(fechaHasta) : null;                                                                               
        filas = filas.filter(row => {                                                                                                         
          const f = row._fechaOrden;                                                                                                          
          return (!desde || f >= desde) && (!hasta || f <= hasta);                                                                            
        });                                                                                                                                   
      }                                                                                                                                       
                                                                                                                                              
      // Ordenar                                                                                                                              
      filas.sort((a, b) => b._fechaOrden - a._fechaOrden);                                                                                    
      filas = filas.map(row => {                                                                                                              
        delete row._fechaOrden;                                                                                                               
        return row;                                                                                                                           
      });                                                                                                                                     
                                                                                                                                              
      const totalFiltrado = filas.length;                                                                                                     
      const paginated = filas.slice(start, start + length);                                                                                   
                                                                                                                                              
      return {                                                                                                                                
        headers: headers,                                                                                                                     
        data: paginated,                                                                                                                      
        total: totalFiltrado                                                                                                                  
      };                                                                                                                                      
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    function insertarYObtenerDatosCumplimiento(valor) {                                                                                       
      const hoja = getSpreadsheetCapacitaciones().getSheetByName("CUMPLIMIENTO🧍‍♂️");                                                          
      hoja.getRange("B3").setValue(valor);                                                                                                    
                                                                                                                                              
      const datosColA = hoja.getRange("A:A").getValues();                                                                                     
      let ultimaFilaValida = 0;                                                                                                               
                                                                                                                                              
      for (let i = datosColA.length - 1; i >= 0; i--) {                                                                                       
        const valorCelda = datosColA[i][0];                                                                                                   
        if (typeof valorCelda === "number" && valorCelda > 1) {                                                                               
          ultimaFilaValida = i + 1; // porque i es índice 0-based                                                                             
          break;                                                                                                                              
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      if (ultimaFilaValida === 0) return []; // no hay filas válidas                                                                          
                                                                                                                                              
      const rango = hoja.getRange(`A1:H${ultimaFilaValida}`);                                                                                 
      return rango.getValues();                                                                                                               
    } 

                                                                                                                                             
    //CHARLAS                                                                                                                                                                                                                                                                    
// function saveData(value) {
//   const ws = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");
//   const randomId = Math.floor(Math.random() * 1e8);
//   const fila = [randomId].concat(value);
//   ws.appendRow(fila);
// }
    function saveData(value) {
  const ws = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");
  const randomId = Math.floor(Math.random() * 1e8);

  // Forzar campo "Comentarios" como texto (posición 11 si empieza desde 0)
  if (value[10] !== undefined && value[10] !== null) {
    value[10] = "'- " + value[10].toString();
  }

  const fila = [randomId].concat(value);
  ws.appendRow(fila);
}
                                                                                                                           
                                                                                                                                              
                                                                                                                                              
    function uploadFilesToDrive(files) {                                                                                                      
      var folder = DriveApp.getFolderById(foldercharlas); //ARCHIVO 1 "Charlas"                                                               
      var urls = files.map(function(file) {                                                                                                   
        var contentType = file.data.match(/^data:(.*?);/)[1];                                                                                 
        var base64Data = file.data.split(',')[1];                                                                                             
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, file.filename);                                         
        var createdFile = folder.createFile(blob);                                                                                            
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);                                                   
        return createdFile.getUrl();                                                                                                          
      });                                                                                                                                     
      return urls;                                                                                                                            
    }                                                                                                                                              
                                                                                                                                              
    //BUSCADOR CHARLAS                                                                                                                        
    function getDatosDesdeCharlas(start = 0, length = 30, search = "", mes = "Todos") {
      const sheet = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");
      const lastRow = sheet.getLastRow();

      // Solo columnas A–M → col 1 a 14
      const data = sheet.getRange(1, 1, lastRow, 14).getValues();
      const headers = data[0];
      let rows = data.slice(1);

      const mesNum = (mes && mes !== "Todos" && mes !== "todos") ? parseInt(mes, 10) : null;

      // Procesar fechas (columna B) y filtrar
      rows = rows.map(row => {
        const fecha = row[1];
        if (fecha instanceof Date) {
          row[1] = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
          row._fechaOrden = fecha;
        } else {
          row._fechaOrden = new Date("1900-01-01");
        }
        return row;
      });

      // Filtro por mes (columna B = índice 1) — sin restricción de año
      if (mesNum !== null) {
        rows = rows.filter(row => {
          const f = row._fechaOrden;
          return (f.getMonth() + 1) === mesNum;
        });
      }

      // Filtro por texto
      if (search) {
        const s = search.toLowerCase();
        rows = rows.filter(row => row.some(cell => String(cell).toLowerCase().includes(s)));
      }

      // Ordenar por fecha descendente
      rows.sort((a, b) => b._fechaOrden - a._fechaOrden);

      // Eliminar campo auxiliar
      rows = rows.map(row => { delete row._fechaOrden; return row; });

      const totalFiltrado = rows.length;
      const paginated = rows.slice(start, start + length);

      return { headers: headers, data: paginated, total: totalFiltrado };
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    //GUARDAR EDICION CHARLAS                                                                                                                 
    function actualizarFilaPorID(id, nuevosDatos) {                                                                                           
      const sheet = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");                                                                
      const lastRow = sheet.getLastRow();                                                                                                     
                                                                                                                                              
      // Solo columnas A–M → col 1 a 14                                                                                                       
      const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();                                                                         
                                                                                                                                              
      for (let i = 0; i < data.length; i++) {                                                                                                 
        if (String(data[i][0]).trim() === String(id).trim()) {                                                                                
          sheet.getRange(i + 2, 1, 1, nuevosDatos.length).setValues([nuevosDatos]);                                                           
          return true;                                                                                                                        
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      throw new Error("ID no encontrado");                                                                                                    
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    //ELIMINA CHARLA                                                                                                                          
    function eliminarRegistroPorId(id) {                                                                                                      
      const hoja = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");                                                                 
      const lastRow = hoja.getLastRow();                                                                                                      
                                                                                                                                              
      // Solo columnas A–M (1 a 14), sin encabezado                                                                                           
      const datos = hoja.getRange(2, 1, lastRow - 1, 14).getValues();                                                                         
                                                                                                                                              
      for (let i = 0; i < datos.length; i++) {                                                                                                
        if (String(datos[i][0]).trim() === String(id).trim()) {                                                                               
          hoja.deleteRow(i + 2); // +2 porque datos empieza en fila 2 y el índice i es base 0                                                 
          return `Registro con ID ${id} eliminado.`;                                                                                          
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      throw new Error(`No se encontró el registro con ID ${id}`);                                                                             
    }                                                                                                                                         
   
//CREA CON BLOB
function generarCertificadoDesdeNombreYTema(nombre, tema) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName('CERTIFICADO');
  hoja.getRange('D9').setValue(nombre);
  hoja.getRange('D13').setValue(tema);

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  const sheetId = hoja.getSheetId();
  const url = getSpreadsheetCapacitaciones().getUrl().replace(/edit$/, '');
  const exportUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:G24&portrait=false';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token }
  });

  const blob = response.getBlob().setName(`Certificado - ${nombre}.pdf`);
  const base64 = Utilities.base64Encode(blob.getBytes());

  return base64;
}


//IA CAPACITACIONES
//const API_KEY2 = "AIzaSyBlm8NHhMDagHHREeGrqLWRInfcHK6Y_bw";
function analizarArchivoConGemini(base64DataUrl) {
  //const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`;
  
  const mimeType = base64DataUrl.match(/^data:(.*?);/)[1];
  const base64 = base64DataUrl.split(',')[1];

  // 🔁 Obtener las listas reales desde la hoja 'LISTAS'
  const listas = getTodasLasListas();

  // 🧠 Construir dinámicamente el prompt con esas listas
  const prompt = `
Extrae los siguientes campos desde el contenido del archivo adjunto.
Usa exactamente los valores disponibles en los siguientes menús:

Empresas válidas: ${listas.empresas.join(", ")}
Lugares válidos: ${listas.lugares.join(", ")}
Tipo de formación válidas: ${listas.capacitaciones.join(", ")}
Gestiones válidas: ${listas.areas.join(", ")}
Registrado por válidas: ${listas.trabajadores.join(", ")}

Devuelve los campos con este formato:

Fecha:  
Tema:  
Lugar:  
Tipo de formación:  
Capacitador:  
Empresa:  
Gestión:  
Duración (min):  
Asistentes:  
Comentarios:  
Registrado por:
`;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inlineData: {
            mimeType: mimeType,
            data: base64
          }
        }
      ]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(geminiUrl, options);
  const json = JSON.parse(response.getContentText());

  if (json.candidates && json.candidates[0]?.content?.parts?.[0]?.text) {
    return json.candidates[0].content.parts[0].text;
  } else {
    return `⚠️ Error en respuesta de Gemini:\n${JSON.stringify(json)}`;
  }
}
/**
 * 🔹 Usa Gemini 2.5 Flash para generar preguntas con o sin archivo, con o sin texto adicional.
 * @param {string|null} base64DataUrl - Archivo en base64 (PDF, Word, etc) o null si no hay archivo.
 * @param {number} numPreguntas - Cantidad de preguntas a generar.
 * @param {string|null} textoBase - Texto de contexto o instrucciones del usuario.
 * @return {string} JSON con formato [{pregunta, respuestas[], correcta}]
 */
function analizarConGemini(base64DataUrl, numPreguntas, textoBase) {
  const hayArchivo = !!base64DataUrl;
  const hayTexto = !!textoBase && textoBase.trim() !== "";

  // 🔹 Construir el prompt dinámico según lo que el usuario proporcione
  let prompt = "";

  if (hayTexto && hayArchivo) {
    prompt = `
El usuario proporcionó un archivo y una instrucción adicional.
Analiza el documento adjunto teniendo en cuenta lo siguiente:
"""${textoBase}"""
Genera ${numPreguntas} preguntas de opción múltiple relevantes al contexto indicado.
`;
  } else if (hayArchivo) {
    prompt = `
Analiza el siguiente documento y genera ${numPreguntas} preguntas de opción múltiple.
`;
  } else if (hayTexto) {
    prompt = `
Genera ${numPreguntas} preguntas de opción múltiple basadas en el siguiente texto:
"""${textoBase}"""
`;
  } else {
    throw new Error("No se recibió ni texto ni archivo para procesar.");
  }

  prompt += `
Cada pregunta debe tener:
- 1 texto de pregunta.
- 4 posibles respuestas.
- 1 número que indique cuál es la correcta (1 a 4).

El formato de salida debe ser JSON puro, sin texto adicional, así:
[
  {
    "pregunta": "¿Cuál es el resultado de 2+2?",
    "respuestas": ["1","2","3","4"],
    "correcta": 4
  }
]
Devuelve **solo el JSON**, sin explicaciones ni texto fuera del arreglo.
`;

  // 🔹 Crear el payload dependiendo de si hay archivo o no
  const parts = [{ text: prompt }];
  if (hayArchivo) {
    const mimeType = base64DataUrl.match(/^data:(.*?);/)[1];
    const base64 = base64DataUrl.split(",")[1];
    parts.push({ inlineData: { mimeType, data: base64 } });
  }

  const payload = { contents: [{ parts }] };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;

  // 🔹 Llamada a Gemini
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const texto = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

  const inicio = texto.indexOf("[");
  const fin = texto.lastIndexOf("]");
  if (inicio === -1 || fin === -1) throw new Error("Respuesta inválida de Gemini.");

  // 🔹 Normalización del JSON
  try {
    const arr = JSON.parse(texto.substring(inicio, fin + 1));
    arr.forEach(p => {
      let c = p.correcta;
      if (typeof c === "string") {
        c = c.trim().toUpperCase();
        if (["A", "B", "C", "D"].includes(c)) c = ["A","B","C","D"].indexOf(c) + 1;
        else if (/^\d$/.test(c)) c = parseInt(c);
        else c = 1;
      }
      if (isNaN(c) || c < 1 || c > 4) c = 1;
      p.correcta = c;
    });
    return JSON.stringify(arr);
  } catch (e) {
    throw new Error("Error al interpretar el JSON generado por Gemini: " + e);
  }
}


///FORMULARIO DE CAPACITACIONES
// -----------------------------
// MÓDULO TEMAS (Server side)
// -----------------------------
var temasData = "TEMAS";

/**
 * Para cada registro en la lista, devuelve cuántos trabajadores firmaron (REGISTRO FIRMAS)
 * y el total de activos, para mostrar el progreso en RegistrosCap.
 * Input: [ { codigo, tema, fecha } ]
 * Output: { [codigo]: { firmados, total, pct } }
 */
function obtenerConteoPorRegistros(registros) {
  const ss = getSpreadsheetCapacitaciones();

  // Leer REGISTRO FIRMAS: col B=Tema (idx 1), col L=Código Registro (idx 11)
  const hRF = ss.getSheetByName("REGISTRO FIRMAS");
  let rfData = [];
  if (hRF && hRF.getLastRow() > 1) {
    const numCols = Math.max(hRF.getLastColumn(), 12);
    rfData = hRF.getRange(2, 1, hRF.getLastRow() - 1, numCols).getValues();
  }

  // Contar trabajadores activos
  const hP = getSpreadsheetPersonal().getSheetByName("PERSONAL");
  let totalActivos = 0;
  if (hP.getLastRow() > 1) {
    const estados = hP.getRange(2, 16, hP.getLastRow() - 1, 1).getValues();
    totalActivos = estados.filter(r => String(r[0] || "").toUpperCase() === "ACTIVO").length;
  }
  if (totalActivos === 0) totalActivos = 1;

  const result = {};
  (registros || []).forEach(function(reg) {
    const codigoNorm = String(reg.codigo || "").trim().toUpperCase();
    const temaNorm   = String(reg.tema   || "").trim().toLowerCase();
    let firmados = 0;

    // Primary: match by Código Registro (col 12, index 11) — new records
    rfData.forEach(function(row) {
      const rowCodigo = String(row[11] || "").trim().toUpperCase();
      if (rowCodigo && rowCodigo === codigoNorm) firmados++;
    });

    // Fallback: if nothing matched by código, count by tema name (old records without código)
    if (firmados === 0) {
      rfData.forEach(function(row) {
        const rowCodigo = String(row[11] || "").trim();
        const rowTema   = String(row[1]  || "").trim().toLowerCase();
        if (!rowCodigo && rowTema === temaNorm) firmados++;
      });
    }

    result[reg.codigo] = {
      firmados: firmados,
      total:    totalActivos,
      pct:      Math.round(firmados / totalActivos * 100)
    };
  });

  return result;
}

/**
 * Devuelve encabezados, filas paginadas y total (claves: headersTemas, filas, total).
 */
function obtenerMatrizTemasPaginado(offset, limit, filtro = "") {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);
  const lastRow = hoja.getLastRow();

  if (lastRow < 2) {
    return { headersTemas: [], filas: [], total: 0 };
  }

  // Leer encabezados desde la primera fila (todas las columnas presentes)
  const numCols = hoja.getLastColumn();
  const headersTemas = hoja.getRange(1, 1, 1, numCols).getDisplayValues()[0];

  // Leer solo datos desde la fila 2 hasta la última (sin encabezado)
  const numFilas = lastRow - 1;
  const datos = hoja.getRange(2, 1, numFilas, numCols).getDisplayValues();

  // Invertimos los datos para mostrar los últimos primero
  let filas = datos.reverse();

  // Filtro si se proporciona
  if (filtro) {
    const texto = String(filtro).toLowerCase();
    filas = filas.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(texto))
    );
  }

  // Paginación
  const paginadas = filas.slice(offset, offset + limit);

  return {
    headersTemas,
    filas: paginadas,
    total: filas.length
  };
}

/**
 * Agrega un nuevo registro en TEMAS.
 * data: objeto con campos: tema, area, capacitador, duracion, intentos, validez, estado, horaInicio, horaFin, examen, valoracion
 * Genera Código automático (prefijo "T" + 7 caracteres).
 */
function agregarTema(data) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);

  // Generar código único (T + 7 caracteres alfanum)
  const codigo = generarCodigoTema();
  const fila = [
    codigo,                         // Código
    data.tema || "",                // Temas
    data.area || "",                // Área
    data.capacitador || "",         // Capacitador
    data.duracion || "",            // Duración (Min)
    data.intentos || "0",           // Intentos (Veces)
    data.validez || "",             // Validez (Min)
    data.estado || "ABIERTO",        // Estado
    data.horaInicio || "",          // HoraInicio
    data.horaFin || "",             // HoraFin
    data.examen || "No",            // Examen
    data.valoracion || "No",        // Valoración
    data.puntajeMinimo || "0"       // PuntajeMinimo
  ];

  hoja.appendRow(fila);
  return codigo;
}

/**
 * Actualiza un tema según su Código (data.codigo).
 * data debe contener las 12 columnas en formato de objeto (ver arriba).
 */
function actualizarTema(data) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);

  // Generar NUEVO código (como solicitaste)
  const nuevoCodigo = generarCodigoTema();
  const codigoBuscado = String(data.codigo || "").trim();

  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const codigos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoBuscado) {

      const fila = [
        nuevoCodigo,                // Nuevo código reemplaza al anterior
        data.tema || "",
        data.area || "",
        data.capacitador || "",
        data.duracion || "",
        data.intentos || "0",
        data.validez || "",
        data.estado || "Activo",
        data.horaInicio || "",
        data.horaFin || "",
        data.examen || "No",
        data.valoracion || "No",
        data.puntajeMinimo || "0"   // PuntajeMinimo
      ];

      hoja.getRange(i + 2, 1, 1, fila.length).setValues([fila]);

      // Retornar el nuevo código para el Swal
      return nuevoCodigo;
    }
  }
  return null;
}

/**
 * Elimina tema por Código.
 */
function eliminarTemaPorCodigo(codigo) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const codigos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  const buscado = String(codigo).trim();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === buscado) {
      hoja.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

/**
 * Cierra un registro de asistencia (Estado → CERRADO).
 */
function cerrarRegistroManual(codigo) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const codigos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  const buscado = String(codigo).trim();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === buscado) {
      hoja.getRange(i + 2, 8).setValue("CERRADO"); // col H = Estado
      return true;
    }
  }
  return false;
}

/**
 * Genera código 'T' + 7 chars alfanum
 */
function generarCodigoTema() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let out = "T";
  for (let i = 0; i < 7; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

// ✅ Obtener tema por código
function getTemaPorCodigo(codigo, dni) {
  const ss = getSpreadsheetCapacitaciones();
  const hoja = ss.getSheetByName("TEMAS");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { error: "No hay datos en la hoja TEMAS." };

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 11).getValues();
  const codigoBuscado = String(codigo).trim().toUpperCase();
  const ahora = new Date();

  for (let i = datos.length - 1; i >= 0; i--) {
    const [
      codigoTema,  // A (0)
      tema,        // B (1)
      area,        // C (2)
      capacitador, // D (3)
      duracion,    // E (4)
      intentos,    // F (5)
      ,            // G (6) - validez
      estado,      // H (7) - ABIERTO/CERRADO (nuevo) o Activo (antiguo)
      fechaInicio, // I (8)
      fechaFin,    // J (9)
      status       // K (10) - Examen Sí/No
    ] = datos[i];

    if (String(codigoTema).toUpperCase() !== codigoBuscado) continue;

    const estadoNorm = String(estado || "").trim().toUpperCase();

    // Registro explícitamente cerrado → rechazar
    if (estadoNorm === "CERRADO") {
      return { error: "Este registro está CERRADO. No se pueden registrar más firmas." };
    }

    // Registro ABIERTO (nuevo flujo) → sin validación de fechas
    // Registro con otro estado (Activo, etc.) → validar por rango de fechas (flujo antiguo)
    if (estadoNorm !== "ABIERTO") {
      const inicio = fechaInicio ? new Date(fechaInicio) : null;
      const fin    = fechaFin    ? new Date(fechaFin)    : null;
      if (!inicio || !fin) {
        return { error: "El curso no tiene fechas válidas configuradas." };
      }
      if (ahora < inicio || ahora > fin) {
        return {
          error: `El curso no está disponible en este rango:\nInicio: ${inicio.toLocaleString()}\nFin: ${fin.toLocaleString()}`
        };
      }
    }

    // Verificar si el DNI ya firmó este tema hoy
    let yaFirmo = false;
    if (dni) {
      const hojaBD = ss.getSheetByName("B DATOS");
      const ultimaFilaBD = hojaBD.getLastRow();
      if (ultimaFilaBD > 1) {
        const bdDatos = hojaBD.getRange(2, 1, ultimaFilaBD - 1, 8).getValues();
        const temaNorm = String(tema).trim().toLowerCase();
        const dniNorm  = String(dni).replace(/^'/, "").trim();
        for (const row of bdDatos) {
          const rowDni   = String(row[0]).replace(/^'/, "").trim();
          const rowTema  = String(row[4]).trim().toLowerCase();
          const rowFecha = row[7] ? new Date(row[7]) : null;
          if (rowDni === dniNorm && rowTema === temaNorm && rowFecha &&
              rowFecha.getFullYear() === ahora.getFullYear() &&
              rowFecha.getMonth()    === ahora.getMonth()    &&
              rowFecha.getDate()     === ahora.getDate()) {
            yaFirmo = true;
            break;
          }
        }
      }
    }

    return { tema, area, capacitador, duracion, intentos, status, yaFirmo };
  }

  return { error: "Código no encontrado." };
}



// ✅ Forzar actualización manual del caché
function actualizarCacheTemas() {
  const cache = CacheService.getScriptCache();
  cache.remove("TEMAS_CACHE");
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("TEMAS");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return "Sin datos para actualizar";

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 8).getValues();
  cache.put("TEMAS_CACHE", JSON.stringify(datos), 300);
  return "Cache actualizado correctamente";
}

function getTemasDesdeBD() {  
  const cache = CacheService.getScriptCache();

  // === 1. Intentar obtener desde caché ===
  const cacheLista = cache.get("lista_temas");
  if (cacheLista) {
    try {
      return JSON.parse(cacheLista);  // Respuesta instantánea
    } catch (err) {}
  }

  // === 2. Leer desde hoja ===
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("TEMAS");
  const lastRow = hoja.getLastRow();

  if (lastRow < 2) return [];

  // Leer solo columna B (temas)
  const valores = hoja.getRange(2, 2, lastRow - 1, 1).getValues();

  // Procesar
  const temasUnicos = [...new Set(valores.flat()
    .map(v => v && v.toString().trim())
    .filter(Boolean)
  )].sort();

  // === 3. Guardar en caché 10 min ===
  cache.put("lista_temas", JSON.stringify(temasUnicos), 600);

  return temasUnicos;
}

//  DASHBOARD LABORAL - Reportes Laborales
//
//  LOGICA:
//  - TOTAL siempre = Personal ACTIVO de hoja PERSONAL (denominador)
//  - TEMAS = todos los unicos de B DATOS columna E
//  - Para cada trabajador ACTIVO x cada tema:
//      busca en B DATOS col A (DNI) + col E (Tema) -> col J (Estado)
//      si no tiene registro -> Pendiente (nunca asistio)
//  - % cumplimiento = Aprobados / (total activos x total temas)
// ============================================================
function obtenerDashboardLaboral() {
  const ssCap      = getSpreadsheetCapacitaciones();
  const ssPersonal = getSpreadsheetPersonal();

  // ── 1. PERSONAL: activos e inactivos ─────────────────────
  const hojaPersonal = ssPersonal.getSheetByName('PERSONAL');
  const lastRowP = hojaPersonal.getLastRow();
  const personalRaw = lastRowP > 1
    ? hojaPersonal.getRange(1, 1, lastRowP, 12).getValues()
    : [[]];

  // Todos los trabajadores con DNI
  const todoPersonal = personalRaw.slice(1).filter(r =>
    r[1] && String(r[1]).trim() !== ''
  ).map(r => ({
    dni:     String(r[1]).trim().replace(/^'/, ''),
    nombre:  String(r[2] || '').trim(),
    cargo:   String(r[6] || '').trim(),
    empresa: String(r[4] || '').trim(),
    estado:  String(r[11] || '').trim().toUpperCase()  // ACTIVO / INACTIVO
  }));

  // Solo ACTIVOS = denominador fijo para todas las metricas
  const personalActivo = todoPersonal.filter(r => r.estado === 'ACTIVO');

  // ── 2. B DATOS: todos los registros ──────────────────────
  // Col: A(0)=DNI  B(1)=Nombre  C(2)=Cargo  D(3)=Empresa  E(4)=Tema
  //      F(5)=Area  G(6)=Puntaje  H(7)=Fecha  I(8)=Horas
  //      J(9)=Estado(Aprobado/Reprobado)  K(10)=Temporalidad(dias)
  //      ...  O(14)=Origen (Matriz / Sin programar)
  const hojaBD    = ssCap.getSheetByName('B DATOS');
  const lastRowBD = hojaBD.getLastRow();
  const bdRows = lastRowBD > 1
    ? hojaBD.getRange(2, 1, lastRowBD - 1, 16).getValues()
    : [];

  const ahora = new Date();

  // ── 3. Indexar B DATOS: "dni|tema" -> mejor estado ───────
  // Prioridad: Aprobado-vigente(4) > Reprobado(3) > Caducado(2) > Pendiente(1)
  const mejorPorDniTema = {};
  const temasSet        = new Set();

  bdRows.forEach(function(row) {
    var dni        = String(row[0] || '').trim().replace(/^'/, '');
    var tema       = String(row[4] || '').trim();
    var estatusRaw = String(row[9] || '').trim();
    var dias       = parseInt(row[10]) || 365;
    var puntaje    = parseFloat(row[6]) || 0;
    var fechaRaw   = row[7];
    var origenBD   = String(row[14] || '').trim().toUpperCase();

    if (!dni || !tema) return;
    // Solo capacitaciones de Matriz (col O = "MATRIZ")
    if (origenBD !== 'MATRIZ') return;
    temasSet.add(tema);

    var fecha  = parseFecha(fechaRaw);
    var estado = estatusRaw || 'Pendiente';

    if (estatusRaw === 'Aprobado' && fecha instanceof Date && !isNaN(fecha)) {
      var venc = new Date(fecha);
      venc.setDate(venc.getDate() + dias);
      if (ahora > venc) estado = 'Caducado';
    }

    var clave = dni + '|' + tema;
    var prio  = { 'Aprobado': 4, 'Reprobado': 3, 'Caducado': 2, 'Pendiente': 1 };
    var actual = mejorPorDniTema[clave];
    if (!actual || (prio[estado] || 0) > (prio[actual.estado] || 0)) {
      mejorPorDniTema[clave] = {
        estado: estado,
        puntaje: puntaje,
        fecha: (fecha instanceof Date && !isNaN(fecha))
          ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy')
          : ''
      };
    }
  });

  var todosLosTemas = Array.from(temasSet).sort();

  // ── 4. Cruzar: PERSONAL ACTIVO x TODOS LOS TEMAS ─────────
  // Para cada activo, iterar todos los temas de B DATOS.
  // Sin registro en B DATOS = Pendiente (nunca asistio).
  var cumplimientoPorTrabajador = personalActivo.map(function(trab) {
    var detalle = todosLosTemas.map(function(tema) {
      var reg = mejorPorDniTema[trab.dni + '|' + tema];
      return {
        tema:    tema,
        estado:  reg ? reg.estado  : 'Pendiente',
        fecha:   reg ? reg.fecha   : '',
        puntaje: reg ? reg.puntaje : 0
      };
    });

    var ap = detalle.filter(function(d) { return d.estado === 'Aprobado';  }).length;
    var rp = detalle.filter(function(d) { return d.estado === 'Reprobado'; }).length;
    var cd = detalle.filter(function(d) { return d.estado === 'Caducado';  }).length;
    var pd = detalle.filter(function(d) { return d.estado === 'Pendiente'; }).length;
    var tot = todosLosTemas.length;

    return {
      dni:      trab.dni,
      nombre:   trab.nombre,
      cargo:    trab.cargo,
      empresa:  trab.empresa,
      estadoPersonal: trab.estado,
      temasTotal:  tot,
      aprobados:   ap,
      reprobados:  rp,
      caducados:   cd,
      pendientes:  pd,
      porcentaje:  tot > 0 ? Math.round((ap / tot) * 100) : 0,
      detalle: detalle
    };
  }).sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });

  // ── 5. Resumen por tema (denominador = total ACTIVOS) ─────
  var totalActivos = personalActivo.length;
  var resumenPorTema = todosLosTemas.map(function(tema) {
    var aprobados = 0, reprobados = 0, caducados = 0, pendientes = 0;
    personalActivo.forEach(function(trab) {
      var reg = mejorPorDniTema[trab.dni + '|' + tema];
      var est = reg ? reg.estado : 'Pendiente';
      if      (est === 'Aprobado')  aprobados++;
      else if (est === 'Reprobado') reprobados++;
      else if (est === 'Caducado')  caducados++;
      else                          pendientes++;
    });
    return {
      tema:       tema,
      total:      totalActivos,      // denominador fijo
      aprobados:  aprobados,
      reprobados: reprobados,
      caducados:  caducados,
      pendientes: pendientes,
      registros:  aprobados + reprobados + caducados  // los que al menos asistieron
    };
  });

  // ── 6. Incumplidores: activos con al menos 1 no-Aprobado ─
  var incumplidores = cumplimientoPorTrabajador
    .filter(function(t) { return t.reprobados > 0 || t.caducados > 0 || t.pendientes > 0; })
    .sort(function(a, b) { return a.porcentaje - b.porcentaje; });

  // ── 7. KPIs globales ─────────────────────────────────────
  // Denominador = total activos * total temas (cross join completo)
  var totalCeldas    = totalActivos * todosLosTemas.length;
  var totalAprobados = cumplimientoPorTrabajador.reduce(function(s, t) { return s + t.aprobados; }, 0);

  // ── 8. Vacunas INFO ADICIONAL ─────────────────────────────

    const vacunasHeaders = [
    'EMO','Tétano & Difteria (1)','Tétano & Difteria (2)','Tétano & Difteria (3)',
    'Hepatitis B (1)','Hepatitis B (2)','Hepatitis B (3)',
    'Neumococo','Influenza',
    'COVID-19 (1)','COVID-19 (2)','COVID-19 (3)','COVID-19 (4)',
    'Carnet de Sanidad'
  ];
  const resumenVacunas = {};
  vacunasHeaders.forEach(v => resumenVacunas[v] = { vacuna: v, conVacuna: 0, sinVacuna: 0 });
  const vacunasData = [];

  try {
    const hojaIA = ssPersonal.getSheetByName('INFO ADICIONAL');
    if (hojaIA && hojaIA.getLastRow() > 1) {
      const raw = hojaIA.getRange(2, 1, hojaIA.getLastRow() - 1, 16).getValues();
      raw.forEach(row => {
        const dni    = String(row[0] || '').trim().replace(/^'/, '');
        const nombre = String(row[1] || '').trim();
        if (!dni) return;
        const vacunas = {};
        vacunasHeaders.forEach((v, i) => {
          const val   = row[i + 2];
          const tiene = val !== null && val !== '' && val !== false && String(val).trim() !== '';
          vacunas[v]  = { valor: String(val || '').trim(), tiene };
          if (tiene) resumenVacunas[v].conVacuna++;
          else       resumenVacunas[v].sinVacuna++;
        });
        vacunasData.push({ dni, nombre, vacunas });
      });
    }
  } catch(e) { Logger.log('INFO ADICIONAL error: ' + e); }

  return {
    kpis: {
      totalTrabajadores:      totalActivos,
      totalTemas:             todosLosTemas.length,
      totalCeldas:            totalCeldas,
      totalAprobados:         totalAprobados,
      porcentajeCumplimiento: totalCeldas > 0 ? Math.round((totalAprobados / totalCeldas) * 100) : 0,
      totalIncumplidores:     incumplidores.length
    },
    todosLosTemas:             todosLosTemas,
    resumenPorTema:            resumenPorTema,
    cumplimientoPorTrabajador: cumplimientoPorTrabajador,
    incumplidores:             incumplidores,
    resumenVacunas:            Object.values(resumenVacunas),
    vacunasData:               vacunasData,
    vacunasHeaders:            vacunasHeaders
  };
}

// ═══════════════════════════════════════════════════════════════════
// ASISTENTE DE VOZ — funciones de consulta para Cloudflare Worker
// ═══════════════════════════════════════════════════════════════════

function _asst_buscarRegistros(p) {
  const ss      = getSpreadsheetCapacitaciones();
  const hoja    = ss.getSheetByName("B DATOS");
  if (!hoja) return { registros: [], total: 0, error: "Hoja B DATOS no encontrada" };
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return { registros: [], total: 0 };

  const datos = hoja.getRange(2, 1, lastRow - 1, 14).getDisplayValues();
  const query = String(p.query || "").toLowerCase().trim();
  const dniF  = String(p.dni  || "").replace(/^'/, "").trim();
  const temaF = String(p.tema || "").toLowerCase().trim();

  const filtrados = datos.filter(r => {
    if (String(r[9]).toUpperCase() === "ACTIVACION") return false;
    if (dniF  && !String(r[0]).replace(/^'/, "").includes(dniF)) return false;
    if (temaF && !String(r[4]).toLowerCase().includes(temaF))    return false;
    if (query) {
      const fila = r.join(" ").toLowerCase();
      if (!fila.includes(query)) return false;
    }
    return true;
  }).slice(0, 20).map(r => ({
    dni: String(r[0]).replace(/^'/, ""),
    nombre: r[1], cargo: r[2], empresa: r[3],
    tema: r[4], area: r[5], puntaje: r[6],
    fecha: r[7], estado: r[9], capacitador: r[11],
    comentarios: r[12]
  }));

  return { registros: filtrados, total: filtrados.length };
}

function _asst_verificarTema(p) {
  const ss   = getSpreadsheetCapacitaciones();
  const tema = String(p.tema || "").toLowerCase().trim();

  // Buscar en Matriz
  const hMat = ss.getSheetByName("Matriz");
  if (!hMat) return { encontrado: false, tema: p.tema, error: "Hoja Matriz no encontrada" };
  const numC = hMat.getLastColumn() - 4;
  if (numC > 0) {
    const temas = hMat.getRange(17, 5, 1, numC).getValues()[0];
    for (let i = 0; i < temas.length; i++) {
      if (String(temas[i]).toLowerCase().trim() === tema) {
        const pmArr  = hMat.getRange(13, 5 + i, 1, 1).getValues();
        const durArr = hMat.getRange(11, 5 + i, 1, 1).getValues();
        const capArr = hMat.getRange(7,  5 + i, 1, 1).getValues();
        const pm = parseFloat(pmArr[0][0]) || 0;
        return {
          encontrado: true, fuente: "Matriz",
          tema: temas[i],
          puntajeMinimo: pm,
          tieneExamen: pm > 0,
          duracionMin: durArr[0][0],
          capacitador: capArr[0][0]
        };
      }
    }
  }

  // Buscar en TEMAS
  const hT = ss.getSheetByName("TEMAS");
  if (!hT) return { encontrado: false, tema: p.tema };
  const lr = hT.getLastRow();
  if (lr > 1) {
    const filas = hT.getRange(2, 1, lr - 1, 11).getValues();
    for (const f of filas) {
      if (String(f[1]).toLowerCase().trim() === tema) {
        const inicio = f[8] ? new Date(f[8]) : null;
        const fin    = f[9] ? new Date(f[9]) : null;
        const ahora  = new Date();
        const activo = inicio && fin && ahora >= inicio && ahora <= fin;
        return {
          encontrado: true, fuente: "TEMAS",
          codigo: f[0], tema: f[1], area: f[2],
          capacitador: f[3], duracionMin: f[4],
          tieneExamen: f[10],
          activo: activo,
          horaInicio: inicio ? inicio.toLocaleDateString("es-PE") : "",
          horaFin: fin ? fin.toLocaleDateString("es-PE") : ""
        };
      }
    }
  }

  return { encontrado: false, tema: p.tema };
}

function _asst_estadoTrabajador(p) {
  const ss  = getSpreadsheetCapacitaciones();
  const dni = String(p.dni || "").replace(/^'/, "").trim();
  if (!dni) return { error: "DNI requerido" };

  const hoja = ss.getSheetByName("B DATOS");
  if (!hoja) return { dni, registros: [], error: "Hoja B DATOS no encontrada" };
  const lr   = hoja.getLastRow();
  if (lr < 2) return { dni, registros: [] };

  const datos = hoja.getRange(2, 1, lr - 1, 11).getDisplayValues();
  const regs  = datos.filter(r => {
    if (String(r[9]).toUpperCase() === "ACTIVACION") return false;
    return String(r[0]).replace(/^'/, "").trim() === dni;
  }).map(r => ({
    tema: r[4], estado: r[9], puntaje: r[6], fecha: r[7], capacitador: r[11]
  }));

  const aprobados  = regs.filter(r => r.estado === "Aprobado").length;
  const reprobados = regs.filter(r => r.estado === "Reprobado").length;
  const nombre     = regs.length > 0 ? (datos.find(r => String(r[0]).replace(/^'/, "").trim() === dni) || [])[1] || "" : "";

  return { dni, nombre, totalRegistros: regs.length, aprobados, reprobados, registros: regs.slice(0, 10) };
}

function _asst_resumenCumplimiento(p) {
  const ss   = getSpreadsheetCapacitaciones();
  const hoja = ss.getSheetByName("B DATOS");
  if (!hoja) return { resumen: [], error: "Hoja B DATOS no encontrada" };
  const lr   = hoja.getLastRow();
  if (lr < 2) return { resumen: [] };

  const temaF = String(p.tema || "").toLowerCase().trim();
  const datos = hoja.getRange(2, 1, lr - 1, 10).getDisplayValues();

  const map = {};
  for (const r of datos) {
    if (String(r[9]).toUpperCase() === "ACTIVACION") continue;
    const t = String(r[4]).trim();
    if (temaF && !t.toLowerCase().includes(temaF)) continue;
    if (!map[t]) map[t] = { tema: t, aprobados: 0, reprobados: 0, pendientes: 0 };
    if (r[9] === "Aprobado")   map[t].aprobados++;
    else if (r[9] === "Reprobado") map[t].reprobados++;
    else map[t].pendientes++;
  }

  return { resumen: Object.values(map) };
}

function _asst_temasActivos(p) {
  const ss  = getSpreadsheetCapacitaciones();
  const ahora = new Date();
  const activos = [];

  // Desde Matriz (tienen ciclos ACTIVACION en B DATOS)
  const hBD = ss.getSheetByName("B DATOS");
  const lr  = hBD.getLastRow();
  if (lr > 1) {
    const datos = hBD.getRange(2, 1, lr - 1, 16).getValues();
    for (const r of datos) {
      const esAct = String(r[9]).toUpperCase() === "ACTIVACION" || String(r[0]).toUpperCase() === "ACTIVACION";
      if (!esAct) continue;
      const rango = _parsearRangoCiclo(String(r[15] || ""));
      if (rango && ahora >= rango.inicio && ahora <= rango.fin) {
        activos.push({ tema: String(r[4]), fuente: "Matriz", ciclo: String(r[15]) });
      }
    }
  }

  // Desde TEMAS (con fechas activas)
  const hT = ss.getSheetByName("TEMAS");
  const lrT = hT.getLastRow();
  if (lrT > 1) {
    const filas = hT.getRange(2, 1, lrT - 1, 10).getValues();
    for (const f of filas) {
      const inicio = f[8] ? new Date(f[8]) : null;
      const fin    = f[9] ? new Date(f[9]) : null;
      if (inicio && fin && ahora >= inicio && ahora <= fin) {
        activos.push({ tema: String(f[1]), fuente: "TEMAS", codigo: String(f[0]), capacitador: String(f[3]) });
      }
    }
  }

  return { activos, total: activos.length };
}

function _asst_registroFirmas(p) {
  const ss   = getSpreadsheetCapacitaciones();
  const hoja = ss.getSheetByName("REGISTRO FIRMAS");
  if (!hoja) return { registros: [], total: 0, mensaje: "Hoja REGISTRO FIRMAS aún vacía" };

  const lr = hoja.getLastRow();
  if (lr < 2) return { registros: [], total: 0 };

  const datos = hoja.getRange(2, 1, lr - 1, 11).getDisplayValues();
  const temaF = String(p.tema || "").toLowerCase().trim();

  const filtrados = datos.filter(r => {
    if (temaF && !String(r[1]).toLowerCase().includes(temaF)) return false;
    return true;
  }).slice(0, 20).map(r => ({
    fecha: r[0], tema: r[1], dni: String(r[2]).replace(/^'/, ""),
    nombre: r[3], cargo: r[4], empresa: r[5], area: r[6],
    capacitador: r[7], duracion: r[8], comentarios: r[9], firmaUrl: r[10]
  }));

  return { registros: filtrados, total: lr - 1 };
}

// ──────────────────────────────────────────────────────────────────────────────
// ASISTENTE IA — Módulos extendidos (Personal, EPP, Eventos, Desvíos, Checklist, IPERC)
// ──────────────────────────────────────────────────────────────────────────────

function _asst_consultarPersonal(p) {
  const hoja = getSpreadsheetPersonal().getSheetByName("PERSONAL");
  if (!hoja) return { error: "Hoja PERSONAL no encontrada" };
  const lr = hoja.getLastRow();
  if (lr < 2) return { trabajadores: [], total: 0 };

  const datos    = hoja.getRange(2, 1, lr - 1, 12).getDisplayValues();
  const dniF     = String(p.dni     || "").trim().replace(/^'/, "");
  const queryF   = String(p.query   || "").toLowerCase().trim();
  const empresaF = String(p.empresa || "").toLowerCase().trim();
  const cargoF   = String(p.cargo   || "").toLowerCase().trim();
  const estadoF  = String(p.estado  || "").toLowerCase().trim();

  const filtrados = datos.filter(r => {
    const dni      = String(r[1] || "").replace(/^'/, "");
    const nombre   = String(r[2] || "").toLowerCase();
    const empresa  = String(r[4] || "").toLowerCase();
    const cargo    = String(r[6] || "").toLowerCase();
    const condicion = String(r[11] || "").toLowerCase();
    if (dniF     && !dni.includes(dniF)) return false;
    if (queryF   && !nombre.includes(queryF) && !dni.includes(queryF)) return false;
    if (empresaF && !empresa.includes(empresaF)) return false;
    if (cargoF   && !cargo.includes(cargoF)) return false;
    if (estadoF  && !condicion.includes(estadoF)) return false;
    return true;
  }).slice(0, 20).map(r => ({
    dni:       String(r[1] || "").replace(/^'/, ""),
    nombre:    r[2], empresa: r[4], cargo: r[6], condicion: r[11]
  }));

  const totalActivos   = datos.filter(r => String(r[11] || "").toLowerCase().includes("activ")).length;
  const totalInactivos = datos.filter(r => String(r[11] || "").toLowerCase().includes("inactiv")).length;

  return { trabajadores: filtrados, encontrados: filtrados.length, totalActivos, totalInactivos, totalGeneral: lr - 1 };
}

function _asst_consultarVacunasEmo(p) {
  const hoja = getSpreadsheetPersonal().getSheetByName("INFO ADICIONAL");
  if (!hoja) return { error: "Hoja INFO ADICIONAL no encontrada" };
  const lr = hoja.getLastRow();
  if (lr < 2) return { registros: [], total: 0 };

  const vacunasHeaders = [
    "EMO","Tétano & Difteria (1)","Tétano & Difteria (2)","Tétano & Difteria (3)",
    "Hepatitis B (1)","Hepatitis B (2)","Hepatitis B (3)",
    "Neumococo","Influenza",
    "COVID-19 (1)","COVID-19 (2)","COVID-19 (3)","COVID-19 (4)","Carnet de Sanidad"
  ];

  const datos   = hoja.getRange(2, 1, lr - 1, 16).getValues();
  const dniF    = String(p.dni    || "").trim().replace(/^'/, "");
  const vacunaF = String(p.vacuna || "").toLowerCase().trim();
  const soloSin = String(p.soloSin || "") === "true";

  const registros = datos.filter(r => {
    const dni = String(r[0] || "").replace(/^'/, "");
    return !dniF || dni.includes(dniF);
  }).map(r => {
    const dni = String(r[0] || "").replace(/^'/, "");
    const nombre = String(r[1] || "").trim();
    const vacunas = {};
    vacunasHeaders.forEach((v, i) => {
      const val = r[i + 2];
      vacunas[v] = (val !== null && val !== "" && val !== false && String(val).trim() !== "")
        ? String(val).trim() : null;
    });
    return { dni, nombre, vacunas };
  });

  const resumen = {};
  vacunasHeaders.forEach(v => {
    resumen[v] = {
      conVacuna: registros.filter(r => r.vacunas[v]).length,
      sinVacuna: registros.filter(r => !r.vacunas[v]).length
    };
  });

  const keyBuscada = vacunaF ? vacunasHeaders.find(v => v.toLowerCase().includes(vacunaF)) : null;
  const resultado  = soloSin
    ? registros.filter(r => keyBuscada ? !r.vacunas[keyBuscada] : Object.values(r.vacunas).some(v => !v))
    : registros;

  return { registros: resultado.slice(0, 20), resumen, total: lr - 1, vacunasHeaders };
}

function _asst_consultarEppStock(p) {
  const ss   = getSpreadsheetEPP();
  const hoja = ss ? ss.getSheetByName("STOCK") : null;
  if (!hoja) return { error: "Hoja STOCK EPP no encontrada" };
  const lr = hoja.getLastRow();
  if (lr < 2) return { items: [], total: 0 };

  const datos      = hoja.getRange(2, 1, lr - 1, 10).getDisplayValues();
  const productoF  = String(p.producto  || "").toLowerCase().trim();
  const almacenF   = String(p.almacen   || "").toLowerCase().trim();
  const soloAlerta = String(p.soloAlerta || "") === "true";

  const filtrados = datos.filter(r => {
    if (productoF && !String(r[2] || "").toLowerCase().includes(productoF)
                  && !String(r[3] || "").toLowerCase().includes(productoF)) return false;
    if (almacenF  && !String(r[1] || "").toLowerCase().includes(almacenF)) return false;
    if (soloAlerta && (parseFloat(r[5]) || 0) > (parseFloat(r[6]) || 0)) return false;
    return true;
  }).slice(0, 25).map(r => ({
    almacen: r[1], producto: r[2], variante: r[3], categoria: r[4],
    stock: r[5], stockMinimo: r[6], precio: r[7],
    alerta: (parseFloat(r[5]) || 0) <= (parseFloat(r[6]) || 0)
  }));

  const totalConAlerta = datos.filter(r => (parseFloat(r[5]) || 0) <= (parseFloat(r[6]) || 0)).length;
  return { items: filtrados, encontrados: filtrados.length, totalConAlerta, total: lr - 1 };
}

function _asst_consultarEppRegistro(p) {
  const ss   = getSpreadsheetEPP();
  const hoja = ss ? ss.getSheetByName("REGISTRO") : null;
  if (!hoja) return { error: "Hoja REGISTRO EPP no encontrada" };
  const lr = hoja.getLastRow();
  if (lr < 2) return { registros: [], total: 0 };

  const datos      = hoja.getRange(2, 1, lr - 1, 26).getDisplayValues();
  const dniF       = String(p.dni       || "").trim().replace(/^'/, "");
  const productoF  = String(p.producto  || "").toLowerCase().trim();
  const estadoF    = String(p.estado    || "").toLowerCase().trim();
  const proxVencer = String(p.proxVencer || "") === "true";
  const hoy        = new Date();

  const filtrados = datos.filter(r => {
    if (dniF      && !String(r[6]  || "").replace(/^'/, "").includes(dniF))       return false;
    if (productoF && !String(r[4]  || "").toLowerCase().includes(productoF))       return false;
    if (estadoF   && !String(r[24] || "").toLowerCase().includes(estadoF))         return false;
    if (proxVencer) {
      const fv   = new Date(r[20]);
      if (isNaN(fv.getTime())) return false;
      const dias = (fv - hoy) / 86400000;
      if (dias < 0 || dias > 30) return false;
    }
    return true;
  }).slice(0, 20).map(r => ({
    fecha: r[1], producto: r[4], variante: r[5],
    dni: String(r[6] || "").replace(/^'/, ""), nombres: r[7],
    empresa: r[8], cargo: r[9], cantidad: r[10],
    devolvible: r[16], fechaVencimiento: r[20], estado: r[24]
  }));

  return { registros: filtrados, encontrados: filtrados.length, total: lr - 1 };
}

function _asst_consultarEventos(p) {
  const ss   = getSpreadsheetAccidentes();
  const hoja = ss ? ss.getSheetByName("B DATOS") : null;
  if (!hoja) return { error: "Hoja EVENTOS B DATOS no encontrada" };
  const lr = hoja.getLastRow();
  if (lr < 2) return { eventos: [], total: 0 };

  const datos    = hoja.getRange(2, 1, lr - 1, 21).getDisplayValues();
  const dniF     = String(p.dni     || "").trim().replace(/^'/, "");
  const empresaF = String(p.empresa || "").toLowerCase().trim();
  const tipoF    = String(p.tipo    || "").toLowerCase().trim();
  const estadoF  = String(p.estado  || "").toLowerCase().trim();
  const desdeF   = p.fechaDesde ? new Date(p.fechaDesde) : null;
  const hastaF   = p.fechaHasta ? new Date(p.fechaHasta) : null;

  const filtrados = datos.filter(r => {
    if (dniF     && !String(r[2] || "").replace(/^'/, "").includes(dniF)) return false;
    if (empresaF && !String(r[5] || "").toLowerCase().includes(empresaF)) return false;
    if (tipoF    && !String(r[8] || "").toLowerCase().includes(tipoF)
                 && !String(r[9] || "").toLowerCase().includes(tipoF))    return false;
    if (estadoF  && !String(r[18] || "").toLowerCase().includes(estadoF)) return false;
    if (desdeF || hastaF) {
      const f = new Date(r[1]);
      if (isNaN(f.getTime())) return false;
      if (desdeF && f < desdeF) return false;
      if (hastaF && f > hastaF) return false;
    }
    return true;
  }).slice(0, 20).map(r => ({
    id: r[0], fecha: r[1], dni: String(r[2] || "").replace(/^'/, ""),
    nombre: r[3], cargo: r[4], empresa: r[5],
    lugar: r[6], tipoEvento: r[8], tipo: r[9],
    descripcion: r[11], fechaInicio: r[12], fechaFin: r[13],
    diasDescanso: r[14], responsable: r[15], estado: r[18]
  }));

  const resumenTipos = {};
  datos.forEach(r => {
    const t = String(r[8] || "Sin tipo").trim();
    resumenTipos[t] = (resumenTipos[t] || 0) + 1;
  });

  return { eventos: filtrados, encontrados: filtrados.length, total: lr - 1, resumenPorTipo: resumenTipos };
}

function _asst_consultarDesvios(p) {
  const hoja = getDesviosSpreadsheet().getSheetByName("B DATOS");
  if (!hoja) return { error: "Hoja DESVIOS B DATOS no encontrada" };
  const lr = hoja.getLastRow();
  if (lr < 2) return { desvios: [], total: 0 };

  const datos    = hoja.getRange(2, 1, lr - 1, 21).getDisplayValues();
  const queryF   = String(p.query        || "").toLowerCase().trim();
  const estadoF  = String(p.estado       || "").toLowerCase().trim();
  const clasifF  = String(p.clasificacion|| "").toLowerCase().trim();
  const respF    = String(p.responsable  || "").toLowerCase().trim();

  const filtrados = datos.filter(r => {
    if (queryF  && !String(r[2] || "").toLowerCase().includes(queryF)
               && !String(r[4] || "").toLowerCase().includes(queryF)) return false;
    if (estadoF && !String(r[15] || "").toLowerCase().includes(estadoF)
               && !String(r[20] || "").toLowerCase().includes(estadoF)) return false;
    if (clasifF && !String(r[11] || "").toLowerCase().includes(clasifF)) return false;
    if (respF   && !String(r[9]  || "").toLowerCase().includes(respF)) return false;
    return true;
  }).slice(0, 20).map(r => ({
    id: r[0], nombre: r[2], equipo: r[4],
    descripcion: r[8], responsable: r[9], fecha: r[10],
    clasificacion: r[11], amonestado: r[12], proceso: r[13],
    estado: r[15], estadoDesvio: r[20]
  }));

  const resumenEstado = {};
  datos.forEach(r => {
    const e = String(r[15] || "Sin estado").trim();
    resumenEstado[e] = (resumenEstado[e] || 0) + 1;
  });

  return { desvios: filtrados, encontrados: filtrados.length, total: lr - 1, resumenPorEstado: resumenEstado };
}

function _asst_consultarChecklist(p) {
  try {
    const ss   = SpreadsheetApp.openById("1NR4VtBUqO6DkM_rSjNqC8m19-QPjrd_IW1aEmmsUD6U");
    const hoja = ss.getSheetByName("B DATOS");
    if (!hoja) return { error: "Hoja CHECKLIST B DATOS no encontrada" };
    const lr = hoja.getLastRow();
    if (lr < 2) return { evaluaciones: [], total: 0 };

    const datos     = hoja.getRange(2, 1, lr - 1, 16).getDisplayValues();
    const equipoF   = String(p.equipo   || "").toLowerCase().trim();
    const evaluadoF = String(p.evaluado || "").toLowerCase().trim();
    const estadoF   = String(p.estado   || "").toLowerCase().trim();

    const filtrados = datos.filter(r => {
      if (equipoF   && !String(r[2] || "").toLowerCase().includes(equipoF))  return false;
      if (evaluadoF && !String(r[5] || "").toLowerCase().includes(evaluadoF)) return false;
      if (estadoF   && !String(r[12] || "").toLowerCase().includes(estadoF))  return false;
      return true;
    }).slice(0, 15).map(r => {
      const numericas = String(r[15] || "").split(",").map(v => {
        const t = v.trim();
        if (t === "si") return 5; if (t === "no") return 1; if (t === "na") return null;
        const n = parseFloat(t); return isNaN(n) ? null : n;
      }).filter(v => v !== null);

      const pctBuenos = numericas.length > 0
        ? Math.round(numericas.filter(v => v >= 4).length / numericas.length * 100) : null;
      const clasificacion = pctBuenos === null ? "Sin datos"
        : pctBuenos >= 70 ? "Bueno" : pctBuenos >= 50 ? "Regular" : "Deficiente";

      return {
        fecha: r[9], equipo: r[2], codigo: r[3], area: r[4],
        evaluado: r[5], evaluador: r[6], lugar: r[7], estado: r[12],
        promedio: numericas.length > 0
          ? (numericas.reduce((a, b) => a + b, 0) / numericas.length).toFixed(1) : null,
        pctBuenos, clasificacion, totalItems: numericas.length
      };
    });

    const resumenClasif = { Bueno: 0, Regular: 0, Deficiente: 0, "Sin datos": 0 };
    filtrados.forEach(r => { resumenClasif[r.clasificacion] = (resumenClasif[r.clasificacion] || 0) + 1; });

    return { evaluaciones: filtrados, encontrados: filtrados.length, total: lr - 1, resumenClasificacion: resumenClasif };
  } catch(e) {
    return { error: "Error al leer checklist: " + e.toString() };
  }
}

function _asst_consultarIperc(p) {
  const ss   = ipercGetSpreadsheet_();
  const hoja = ss ? ss.getSheetByName("DATOS") : null;
  if (!hoja) return { error: "Hoja IPERC DATOS no encontrada" };
  const lr = hoja.getLastRow();
  if (lr <= IPERC_FILA_INICIO) return { riesgos: [], total: 0 };

  const filas    = lr - IPERC_FILA_INICIO;
  const datos    = hoja.getRange(IPERC_FILA_INICIO, 1, filas, 23).getDisplayValues();
  const procesoF = String(p.proceso || "").toLowerCase().trim();
  const areaF    = String(p.area    || "").toLowerCase().trim();
  const nivelF   = String(p.nivel   || "").toLowerCase().trim();

  const filtrados = datos.filter(r => {
    if (!String(r[1] || "").trim()) return false;
    if (procesoF && !String(r[1] || "").toLowerCase().includes(procesoF)) return false;
    if (areaF    && !String(r[2] || "").toLowerCase().includes(areaF))    return false;
    if (nivelF   && !String(r[11] || "").toLowerCase().includes(nivelF)
               && !String(r[20] || "").toLowerCase().includes(nivelF))    return false;
    return true;
  }).slice(0, 20).map(r => ({
    proceso: r[1], area: r[2], tarea: r[3], puesto: r[4],
    peligro: r[6], riesgo: r[7],
    scoreInicial: r[10], nivelInicial: r[11],
    epp: r[16], nivelResidual: r[20], accion: r[21], responsable: r[22]
  }));

  const resumenNivel = {};
  filtrados.forEach(r => {
    const n = r.nivelResidual || r.nivelInicial || "Sin nivel";
    resumenNivel[n] = (resumenNivel[n] || 0) + 1;
  });

  return { riesgos: filtrados, encontrados: filtrados.length, total: filas, resumenPorNivel: resumenNivel };
}
