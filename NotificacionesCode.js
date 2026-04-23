// ============================================================
//  NotificacionesCode.js — Push Notifications vía Cloudflare Worker
//  Envía notificaciones push a dispositivos de trabajadores
// ============================================================

// URL de tu Cloudflare Worker (ACTUALIZAR con tu dominio real)
const PUSH_WORKER_URL = 'https://viczul.com';

// Token secreto para autenticar llamadas GAS → Worker
// IMPORTANTE: Configurar el mismo valor como variable de entorno
// PUSH_AUTH_TOKEN en tu Cloudflare Worker
const PUSH_AUTH_TOKEN = 'adecco_push_2026_secret_token_xyz123';

/**
 * Enviar push a UN trabajador por DNI
 */
function enviarPushNotification(dni, title, body, tag) {
  try {
    if (!dni || !title) return { ok: false, error: 'dni y title requeridos' };

    const payload = {
      token: PUSH_AUTH_TOKEN,
      dni: String(dni).trim(),
      title: title,
      body: body || '',
      tag: tag || 'general',
      url: '/'
    };

    const resp = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/send', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    Logger.log('Push enviado a ' + dni + ': ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('Error enviando push: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Enviar push a VARIOS trabajadores por DNI[]
 */
function enviarPushBulk(dnis, title, body, tag) {
  try {
    if (!Array.isArray(dnis) || !dnis.length || !title) {
      return { ok: false, error: 'dnis[] y title requeridos' };
    }

    const payload = {
      token: PUSH_AUTH_TOKEN,
      dnis: dnis.map(d => String(d).trim()),
      title: title,
      body: body || '',
      tag: tag || 'general',
      url: '/'
    };

    const resp = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/send-bulk', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    Logger.log('Push bulk enviado a ' + dnis.length + ' usuarios: ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('Error enviando push bulk: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================
//  FUNCIONES DE NOTIFICACIÓN POR MÓDULO (automáticas)
// ============================================================

function notificarEntregaEPP(dni, producto, variante) {
  const desc = producto + (variante ? ' (' + variante + ')' : '');
  return enviarPushNotification(
    dni,
    'EPP Asignado',
    'Se te asignó: ' + desc + '. Ingresa para firmar la recepción.',
    'epp-entrega'
  );
}

function notificarConfirmacionEPP(dniSupervisor, trabajadorNombre, producto, accion) {
  const titulo = accion === 'confirmado' ? 'EPP Confirmado' : 'EPP Rechazado';
  const cuerpo = trabajadorNombre + ' ' + accion + ' la recepción de: ' + producto;
  return enviarPushNotification(dniSupervisor, titulo, cuerpo, 'epp-confirmacion');
}

function notificarCapacitacion(dnis, tema, fecha) {
  const body = fecha
    ? 'Capacitación: ' + tema + ' programada para ' + fecha + '. Revisa tu app.'
    : 'Tienes una capacitación asignada: ' + tema + '. Revisa tu app.';
  return enviarPushBulk(dnis, 'Capacitación', body, 'capacitacion');
}

function enviarNotificacionManual(dni, titulo, mensaje) {
  return enviarPushNotification(dni, titulo, mensaje, 'manual');
}

function notificarATodos(titulo, mensaje) {
  try {
    const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'No hay trabajadores' };

    const data = hoja.getRange(2, 1, lastRow - 1, 12).getValues();
    const dnis = [];

    for (let i = 0; i < data.length; i++) {
      const estado = (data[i][11] || '').toString().toUpperCase(); // Col L = estado activo
      const dni = (data[i][1] || '').toString().trim(); // Col B = DNI
      if (dni && (estado === 'SI' || estado === 'ACTIVO')) {
        dnis.push(dni);
      }
    }

    if (!dnis.length) return { ok: false, error: 'No hay trabajadores activos' };
    return enviarPushBulk(dnis, titulo, mensaje, 'general');
  } catch (e) {
    Logger.log('Error notificando a todos: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================
//  FUNCIONES MANUALES — Llamadas desde el frontend (admin)
//  Usadas por google.script.run desde EPPMaestro y Capacitaciones
// ============================================================

/**
 * Enviar alerta manual de EPP desde el panel de administración
 * @param {string} modo - "individual" o "todos"
 * @param {string} [dni] - DNI del trabajador (solo si modo="individual")
 * @param {string} titulo - Título de la alerta
 * @param {string} mensaje - Mensaje de la alerta
 * @returns {Object} { ok, sent, ... }
 */
function enviarAlertaEPP(modo, dni, titulo, mensaje) {
  try {
    if (modo === 'todos') {
      return notificarATodos(
        titulo || 'Alerta EPP',
        mensaje || 'Revisa tu módulo de EPP. Tienes actualizaciones pendientes.'
      );
    } else {
      if (!dni) return { ok: false, error: 'DNI requerido para notificación individual' };
      return enviarPushNotification(
        dni,
        titulo || 'Alerta EPP',
        mensaje || 'Revisa tu módulo de EPP. Tienes actualizaciones pendientes.',
        'epp-manual'
      );
    }
  } catch (e) {
    Logger.log('Error en alerta EPP manual: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Enviar alerta manual de Capacitaciones desde el panel de administración
 * @param {string} modo - "individual", "seleccion" o "todos"
 * @param {string|string[]} dniOrDnis - DNI o array de DNIs
 * @param {string} titulo - Título de la alerta
 * @param {string} mensaje - Mensaje de la alerta
 * @returns {Object} { ok, sent, ... }
 */
function enviarAlertaCapacitacion(modo, dniOrDnis, titulo, mensaje) {
  try {
    const tit = titulo || 'Alerta Capacitación';
    const msg = mensaje || 'Tienes una capacitación pendiente. Revisa tu app.';

    if (modo === 'todos') {
      return notificarATodos(tit, msg);
    } else if (modo === 'seleccion' && Array.isArray(dniOrDnis)) {
      return enviarPushBulk(dniOrDnis, tit, msg, 'cap-manual');
    } else {
      if (!dniOrDnis) return { ok: false, error: 'DNI requerido' };
      return enviarPushNotification(String(dniOrDnis), tit, msg, 'cap-manual');
    }
  } catch (e) {
    Logger.log('Error en alerta Cap manual: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Obtener lista de trabajadores activos para los selects de notificación
 * Devuelve [{dni, nombre, cargo, empresa}]
 */
function obtenerTrabajadoresParaNotificar() {
  try {
    const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return [];

    const data = hoja.getRange(2, 1, lastRow - 1, 12).getValues();
    const trabajadores = [];

    for (let i = 0; i < data.length; i++) {
      const estado = (data[i][11] || '').toString().toUpperCase(); // Col L = estado activo
      const dni = (data[i][1] || '').toString().trim(); // Col B = DNI
      const nombre = (data[i][2] || '').toString().trim(); // Col C = Nombre
      const cargo = (data[i][3] || '').toString().trim(); // Col D = Cargo
      const empresa = (data[i][4] || '').toString().trim(); // Col E = Empresa
      if (dni && (estado === 'SI' || estado === 'ACTIVO')) {
        trabajadores.push({ dni: dni, nombre: nombre, cargo: cargo, empresa: empresa });
      }
    }

    return trabajadores.sort(function(a, b) {
      return (a.nombre || '').localeCompare(b.nombre || '');
    });
  } catch (e) {
    Logger.log('Error obteniendo trabajadores: ' + e.message);
    return [];
  }
}

// ============================================================
//  TEST — Ejecutar desde el editor GAS para diagnosticar conexión
//  En el editor: seleccionar esta función y clic en ▶ Run
// ============================================================
function testPushConnection() {
  // Test 1: Verificar que el Worker responde (GET sin auth)
  Logger.log('=== TEST 1: Verificar Worker ===');
  try {
    var resp1 = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/test', {
      muteHttpExceptions: true
    });
    Logger.log('Status: ' + resp1.getResponseCode());
    Logger.log('Body: ' + resp1.getContentText());
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }

  // Test 2: Enviar push de prueba (POST con auth)
  Logger.log('=== TEST 2: Enviar push con auth ===');
  Logger.log('URL: ' + PUSH_WORKER_URL + '/api/push/send');
  Logger.log('Token que envío (longitud): ' + PUSH_AUTH_TOKEN.length);
  Logger.log('Token que envío (primeros 10): ' + PUSH_AUTH_TOKEN.substring(0, 10));
  try {
    var payload = {
      token: PUSH_AUTH_TOKEN,
      dni: '44366329',
      title: 'Test de conexión',
      body: 'Si ves esto, la conexión funciona',
      tag: 'test'
    };
    Logger.log('Payload: ' + JSON.stringify(payload));
    var resp2 = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/send', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('Status: ' + resp2.getResponseCode());
    Logger.log('Body: ' + resp2.getContentText());
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }

  Logger.log('=== FIN DE TESTS ===');
}


// ============================================================
//  ALERTAS AUTOMÁTICAS DE INSPECCIONES — CHECK LIST
//  Se ejecuta cada 4 horas via time-driven trigger.
//  Envía Push Notification (nativa) a trabajadores en turno
//  + resumen Telegram al administrador.
//
//  Lógica por tipo de frecuencia:
//    Semanal (7d)  → disponible si última inspección NO fue esta semana (lun-dom)
//    Mensual (30d) → disponible si última inspección NO fue este mes calendario
//    Trimestral (90d) → NO fue en el trimestre actual (ene-mar/abr-jun/jul-sep/oct-dic)
//    Semestral (180d) → NO fue en el semestre actual (ene-jun/jul-dic)
//    Anual (365d)  → NO fue este año calendario
//    Cada N días   → rolling: última + (N-1) días (sin cambios)
//
//  Activar:  configurarTriggerAlertasInspeccion()  (una sola vez)
//  Desactivar: eliminarTriggerAlertasInspeccion()
//  Test manual: testAlertasInspeccion()
// ============================================================

var ROL_SS_ID_ALERTAS = '12h2yVs0NlD3h3zMYl_93o7ohOKzurxcPZXifoTyVigE';
var TZ_ALERTAS = 'GMT-5';

function verificarYEnviarAlertasInspeccion() {
  try {
    var ahora    = new Date();
    var hoyStr   = Utilities.formatDate(ahora, TZ_ALERTAS, 'yyyy-MM-dd');
    var horaStr  = Utilities.formatDate(ahora, TZ_ALERTAS, 'HH:mm');
    var mesStr   = Utilities.formatDate(ahora, TZ_ALERTAS, 'yyyy-MM');
    var hoyDate  = _soloFechaInsp(ahora);

    // ── 1. Leer INVENTARIO (fila 3+, cols A-P) ──────────────────────────
    var checkSS  = getCheckSpreadsheet();
    var invSheet = checkSS.getSheetByName('INVENTARIO');
    var invLast  = invSheet.getLastRow();
    if (invLast < 3) { Logger.log('INVENTARIO vacío.'); return; }

    var invData = invSheet.getRange(3, 1, invLast - 2, 16).getValues();
    var equipos = [];
    for (var i = 0; i < invData.length; i++) {
      var r = invData[i];
      var empresa  = String(r[1] || '').trim();
      var area     = String(r[2] || '').trim();
      var equipo   = String(r[3] || '').trim();
      var codigo   = String(r[4] || '').trim();
      var cargo    = String(r[5] || '').trim();  // Col F (A-index 5) = Cargo Responsable
      var diasFreq = _parseDiasFreqInsp(r[12]);  // Col M (A-index 12) = diasFrecuencia
      var estado   = String(r[15] || '').trim().toLowerCase(); // Col P (A-index 15) = estado
      var lugaresRaw = String(r[11] || '').trim(); // Col L (A-index 11) = lugares (elección múltiple)
      var lugares = lugaresRaw
        ? lugaresRaw.split(',').map(function(l){ return l.trim().toUpperCase(); }).filter(Boolean)
        : [];
      if (!codigo || !equipo) continue;
      if (estado === 'retirado') continue;
      if (isNaN(diasFreq) || diasFreq <= 0) continue;
      equipos.push({ empresa: empresa, area: area, equipo: equipo,
                     codigo: codigo, diasFrecuencia: diasFreq, cargo: cargo, lugares: lugares });
    }
    Logger.log('Equipos con frecuencia: ' + equipos.length + ' / Total filas: ' + invData.length);
    if (equipos.length === 0) {
      Logger.log('Sin equipos con frecuencia definida. Verificar columna M (Frecuencia) del INVENTARIO.');
      return;
    }

    // ── 2. Leer B DATOS → última inspección por código||lugar ──────────────────
    var bSheet = checkSS.getSheetByName('B DATOS');
    var bLast  = bSheet.getLastRow();
    var bData  = bLast > 1 ? bSheet.getRange(2, 1, bLast - 1, 15).getValues() : [];
    var ultimaInsp = {};
    var conteoMes  = {};
    for (var j = 0; j < bData.length; j++) {
      var bd = bData[j];
      var bdCod   = String(bd[3] || '').trim();
      var bdFecha = bd[9];
      if (!bdCod || !(bdFecha instanceof Date)) continue;
      var bdLugar = String(bd[7] || '').trim().toUpperCase(); // Col H (index 7) = lugar
      var bdKey   = bdLugar ? (bdCod + '||' + bdLugar) : bdCod;
      if (!ultimaInsp[bdKey] || bdFecha > ultimaInsp[bdKey]) ultimaInsp[bdKey] = new Date(bdFecha);
      if (Utilities.formatDate(bdFecha, TZ_ALERTAS, 'yyyy-MM') === mesStr)
        conteoMes[bdKey] = (conteoMes[bdKey] || 0) + 1;
    }

    // ── 3. Calcular equipos vencidos (por ubicación si aplica) ─────────────────
    var alertas = [];
    var totalMensuales = 0, cumplMensuales = 0;
    for (var k = 0; k < equipos.length; k++) {
      var eq = equipos[k];
      var periodo = _freqEsCalendario(eq.diasFrecuencia);
      var tipo;
      if (periodo === 'semana') {
        tipo = 'SEMANAL';
      } else if (periodo === 'mes' || (!periodo && eq.diasFrecuencia >= 28 && eq.diasFrecuencia <= 31)) {
        tipo = 'MENSUAL';
      } else if (periodo === 'bimestre') {
        tipo = 'BIMESTRAL';
      } else if (periodo === 'trimestre') {
        tipo = 'TRIMESTRAL';
      } else if (periodo === 'semestre') {
        tipo = 'SEMESTRAL';
      } else if (periodo === 'anio') {
        tipo = 'ANUAL';
      } else {
        tipo = 'PERIÓDICA';
      }

      // Iterar por cada lugar (o una vez si no hay lugares)
      var lugaresCheck = eq.lugares.length > 0 ? eq.lugares : [null];
      for (var li = 0; li < lugaresCheck.length; li++) {
        var lugar = lugaresCheck[li];
        var eqKey = lugar ? (eq.codigo + '||' + lugar) : eq.codigo;
        var ultima = ultimaInsp[eqKey] || null;
        var diasVencido;

        if (tipo === 'MENSUAL' || tipo === 'BIMESTRAL' || tipo === 'TRIMESTRAL' ||
            tipo === 'SEMESTRAL' || tipo === 'ANUAL') {
          totalMensuales++;
          if ((conteoMes[eqKey] || 0) > 0) cumplMensuales++;
        }

        if (!ultima) {
          // Sin registro previo → vencido desde el inicio del período actual
          var periodoActual = periodo || 'mes';
          var inicioPer = _inicioPeriodoActualInsp(periodoActual, hoyDate);
          diasVencido = Math.floor((hoyDate - inicioPer) / 86400000);
        } else {
          var ultimaDateN = _soloFechaInsp(ultima);
          diasVencido = _diasVencidoInsp(eq.diasFrecuencia, ultimaDateN, hoyDate);
        }

        if (diasVencido < 0) continue;
        alertas.push({ equipo: eq.equipo + (lugar ? ' — ' + lugar : ''),
                       codigo: eq.codigo, area: eq.area, empresa: eq.empresa,
                       diasFreq: eq.diasFrecuencia, diasVencido: diasVencido, tipo: tipo,
                       cargo: eq.cargo, lugar: lugar || '',
                       ultimaFecha: ultima ? Utilities.formatDate(ultima, TZ_ALERTAS, 'dd/MM/yyyy') : 'Sin registro' });
      }
    }
    if (alertas.length === 0) { Logger.log('✅ Sin inspecciones vencidas a las ' + horaStr); return; }

    // ── 4. Trabajadores en turno HOY (desde rol_turnos.json) ────────────
    var _turnoData   = _getTrabajadoresPorFecha(hoyStr);
    var trabajadoresHoy = _turnoData.trabajadoresHoy;
    var dnisEnTurno     = _turnoData.dnisEnTurno;
    var dnisEnTurnoSet  = {};
    dnisEnTurno.forEach(function(d) { dnisEnTurnoSet[d] = true; });

    // ── 5. Construir mensaje ─────────────────────────────────────────────
    alertas.sort(function(a, b) { return b.diasVencido - a.diasVencido; });
    var porArea = {};
    alertas.forEach(function(a) {
      var k2 = a.area.toUpperCase() || 'SIN ÁREA';
      if (!porArea[k2]) porArea[k2] = [];
      porArea[k2].push(a);
    });

    var lineas = [
      '🔔 <b>ALERTA — INSPECCIONES PENDIENTES</b>',
      '⏰ ' + Utilities.formatDate(ahora, TZ_ALERTAS, 'dd/MM/yyyy HH:mm') + ' hrs',
      '⚠️ <b>' + alertas.length + '</b> inspección(es) vencida(s)', ''
    ];
    // Texto plano para push notification
    var pushLines = [alertas.length + ' inspección(es) vencida(s):'];

    Object.keys(porArea).forEach(function(areaKey) {
      var items  = porArea[areaKey];
      var trabAr = trabajadoresHoy[areaKey] || trabajadoresHoy['GENERAL'] || [];
      lineas.push('━━━━━━━━━━━━━━━━━━');
      lineas.push('📍 <b>' + areaKey + '</b>');
      if (trabAr.length > 0) {
        lineas.push('👷 <b>En turno:</b> ' + trabAr.map(function(w) {
          return w.nombre + (w.turno ? ' [' + w.turno + ']' : '');
        }).join(', '));
      } else {
        lineas.push('👷 <i>Sin personal en turno registrado hoy</i>');
      }
      lineas.push('');
      items.forEach(function(it) {
        var urgencia = it.diasVencido === 0 ? '🟡 HOY' : it.diasVencido <= 2 ? '🟠 +' + it.diasVencido + 'd' : '🔴 +' + it.diasVencido + 'd';
        var freqLabel = it.tipo === 'SEMANAL'    ? 'Semanal (lun-dom)'
                       : it.tipo === 'MENSUAL'    ? 'Mensual (mes calendario)'
                       : it.tipo === 'BIMESTRAL'  ? 'Bimestral'
                       : it.tipo === 'TRIMESTRAL' ? 'Trimestral'
                       : it.tipo === 'SEMESTRAL'  ? 'Semestral'
                       : it.tipo === 'ANUAL'      ? 'Anual'
                       : 'Cada ' + it.diasFreq + 'd';
        lineas.push(urgencia + ' — <code>' + it.codigo + '</code> <b>' + it.equipo + '</b>');
        lineas.push('   📋 ' + freqLabel + ' | Última: ' + it.ultimaFecha);
        pushLines.push(it.equipo + ' (' + it.codigo + ') - ' + freqLabel);
      });
      lineas.push('');
    });

    if (totalMensuales > 0) {
      var pctMensual = Math.round((cumplMensuales / totalMensuales) * 100);
      var diaDelMes  = hoyDate.getDate();
      var diasEnMes  = new Date(hoyDate.getFullYear(), hoyDate.getMonth() + 1, 0).getDate();
      var pctEsperado = Math.max(10, Math.round((diaDelMes / diasEnMes) * 100));
      if (pctMensual < pctEsperado) {
        lineas.push('━━━━━━━━━━━━━━━━━━');
        lineas.push('📊 <b>Cumplimiento mensual:</b> ' + pctMensual + '% / Esperado: ' + pctEsperado + '%');
        lineas.push('   (' + cumplMensuales + '/' + totalMensuales + ' mensuales realizadas)');
        lineas.push('');
      }
    }
    lineas.push('<i>🤖 ADECCO — Próxima verificación en 4 horas</i>');

    // ── 6. Enviar Push Notification por ZONA/POSTA + CARGO ────────────────
    // Filtros combinados:
    //   a) POSTA: se compara el lugar del equipo contra la ZONA del trabajador
    //      Y también contra su TURNO/sub (que puede contener el nombre de la posta).
    //      Ej: zona="SECTOR A", turno="POSTA SUR" → coincide con lugar="POSTA SUR"
    //   b) CARGO: si el equipo tiene cargo(s) asignados en Col F, solo los
    //      trabajadores con ese cargo lo ven
    //   c) ADMIN: zona ADMIN* recibe TODAS las alertas sin filtro
    var cargoPorDni = _getCargoPorDniInsp(); // { dni: cargo_lower }
    var dnisNotificados = {};  // { dni: [equipos] }
    alertas.forEach(function(it) {
      var lugarAlerta = (it.lugar || '').toUpperCase();
      var cargosEquipo = String(it.cargo || '').split(',').map(function(c) { return c.trim().toLowerCase(); }).filter(Boolean);
      dnisEnTurno.forEach(function(dni) {
        var puesto = _getPuestoDeTrabajador(dni, trabajadoresHoy);
        if (!puesto) return;
        var esAdmin = _esZonaAdmin(puesto.zona);
        // Admin → recibe todo
        if (!esAdmin) {
          if (!lugarAlerta) return; // Sin lugar definido → solo admin
          // Filtro posta: coincide si zona O turno/sub del trabajador coinciden con el lugar
          var matchLugar = _zonaMatchesLugar(puesto.zona, lugarAlerta) ||
                           (puesto.turno && _zonaMatchesLugar(puesto.turno, lugarAlerta));
          if (!matchLugar) return;
          // Filtro cargo: si el equipo tiene cargo asignado, el trabajador debe tenerlo
          if (cargosEquipo.length > 0) {
            var dniCargo = (cargoPorDni[dni] || '').toLowerCase();
            var matchCargo = cargosEquipo.some(function(c) {
              return dniCargo.indexOf(c) >= 0 || c.indexOf(dniCargo.split(' ')[0]) >= 0;
            });
            if (!matchCargo) return;
          }
        }
        if (!dnisNotificados[dni]) dnisNotificados[dni] = [];
        dnisNotificados[dni].push(it.equipo);
      });
    });

    // 6b. Agregar alertas de capacitaciones pendientes (PASSO) al push
    var passoPendientes = _getCapacitacionesPendientes();
    dnisEnTurno.forEach(function(dni) {
      var caps = passoPendientes.pendientesPorDni[dni];
      if (!caps || caps.length === 0) return;
      if (!dnisNotificados[dni]) dnisNotificados[dni] = [];
      caps.forEach(function(c) {
        dnisNotificados[dni].push('CAPACITACIÓN: ' + c.capacitacion + ' (' + c.frecuencia + ')');
      });
    });

    // Enviar push a cada trabajador
    Object.keys(dnisNotificados).forEach(function(dni) {
      var equiposList = dnisNotificados[dni];
      var pushBody = equiposList.length + ' pendiente(s):\n' + equiposList.join('\n');
      enviarPushNotification(dni, 'Pendientes — ADECCO', pushBody, 'inspeccion-alerta');
    });
    Logger.log('Push enviado a ' + Object.keys(dnisNotificados).length + ' trabajadores (por zona)');

    // ── 7. Enviar resumen Telegram al administrador ──────────────────────
    var resultado = enviarTelegram(lineas.join('\n'));
    Logger.log('Telegram enviado: ' + JSON.stringify(resultado) + ' | Vencidos: ' + alertas.length);

  } catch (e) {
    Logger.log('❌ Error en verificarYEnviarAlertasInspeccion: ' + e.toString());
  }
}

function configurarTriggerAlertasInspeccion() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'verificarYEnviarAlertasInspeccion') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('verificarYEnviarAlertasInspeccion').timeBased().everyHours(4).create();
  Logger.log('✅ Trigger configurado: verificarYEnviarAlertasInspeccion cada 4 horas');
  return '✅ Trigger configurado cada 4 horas';
}

function eliminarTriggerAlertasInspeccion() {
  var count = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'verificarYEnviarAlertasInspeccion') { ScriptApp.deleteTrigger(t); count++; }
  });
  Logger.log('Eliminados ' + count + ' trigger(s)');
  return 'Eliminados ' + count + ' trigger(s)';
}

function testAlertasInspeccion() { verificarYEnviarAlertasInspeccion(); }

// ── Diagnóstico: imprimir primeras filas del INVENTARIO ──────────────────────
// Ejecutar desde el editor GAS para ver qué hay en cada columna relevante
function debugInventarioFrecuencias() {
  var sheet = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  var lastRow = sheet.getLastRow();
  Logger.log('INVENTARIO — Última fila: ' + lastRow);
  if (lastRow < 3) { Logger.log('INVENTARIO vacío (< 3 filas).'); return; }
  var headers = sheet.getRange(2, 1, 1, 16).getValues()[0];
  Logger.log('CABECERAS A→P: ' + headers.map(function(h, i) {
    return String.fromCharCode(65 + i) + '="' + h + '"';
  }).join(' | '));
  var rows = sheet.getRange(3, 1, Math.min(lastRow - 2, 5), 16).getValues();
  rows.forEach(function(r, idx) {
    Logger.log('Fila ' + (idx + 3) + ': D="' + r[3] + '" E="' + r[4] + '" F="' + r[5] +
               '" M(freq)="' + r[12] + '" P(estado)="' + r[15] + '"' +
               ' → parsedFreq=' + _parseDiasFreqInsp(r[12]));
  });
}

// ── Parser robusto de diasFrecuencia (texto o número) ───────────────────────
// Acepta: 7 (número), "7" (texto), "C/7 DÍAS", "SEMANAL", "MENSUAL", etc.
function _parseDiasFreqInsp(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  var s = String(val).trim().toLowerCase().replace(',', '.');
  var n = parseFloat(s);
  if (!isNaN(n) && n > 0) return n;
  var m = s.match(/[ck]\/(\d+)/) || s.match(/cada\s+(\d+)/);
  if (m) return parseInt(m[1], 10);
  if (s.includes('diario') || s.includes('daily'))    return 1;
  if (s.includes('semanal') || s.includes('weekly'))  return 7;
  if (s.includes('quincenal'))                        return 15;
  if (s.includes('mensual') || s.includes('monthly')) return 30;
  if (s.includes('bimestral'))                        return 60;
  if (s.includes('trimestral'))                       return 90;
  if (s.includes('semestral'))                        return 180;
  if (s.includes('anual') || s.includes('yearly'))   return 365;
  return 0;
}

// ── Mapa DNI → cargo (minúscula) desde hoja PERSONAL ───────────────────────
function _getCargoPorDniInsp() {
  var hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return {};
  var data = hoja.getRange(2, 2, lastRow - 1, 6).getValues(); // cols B:G
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var dni   = String(data[i][0] || '').trim();          // col B = DNI
    var cargo = String(data[i][5] || '').trim().toLowerCase(); // col G = cargo
    if (dni) map[dni] = cargo;
  }
  return map;
}

function _soloFechaInsp(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

// ── Helper: ¿La zona del turno coincide con el lugar del equipo? ────────────
// Maneja abreviaciones (C1 ↔ CONCENTRADORA 1), typos menores (LAYDONW ↔ LAYDOWN)
function _zonaMatchesLugar(zona, lugar) {
  var z = (zona || '').toUpperCase().replace(/\s+/g, ' ').trim();
  var l = (lugar || '').toUpperCase().replace(/\s+/g, ' ').trim();
  if (!z || !l) return false;
  if (z === l) return true;
  if (z.indexOf(l) >= 0 || l.indexOf(z) >= 0) return true;
  // Extraer parte después de "POSTA" para comparar keywords
  var extractKey = function(s) {
    var m = s.match(/POSTA\s+(.+)/);
    return m ? m[1].trim() : s;
  };
  var zk = extractKey(z);
  var lk = extractKey(l);
  if (zk === lk) return true;
  if (zk.indexOf(lk) >= 0 || lk.indexOf(zk) >= 0) return true;
  // Coincidencia por primer carácter + número: "C1" ↔ "CONCENTRADORA 1"
  var zNums = zk.match(/\d+/g) || [];
  var lNums = lk.match(/\d+/g) || [];
  if (zNums.length > 0 && lNums.length > 0 && zNums[0] === lNums[0]) {
    if (zk.charAt(0) === lk.charAt(0)) return true;
  }
  // Primeros 3 caracteres iguales: "LAYDOWN" ↔ "LAYDONW"
  if (zk.length >= 3 && lk.length >= 3 && zk.substring(0, 3) === lk.substring(0, 3)) return true;
  return false;
}

// ── Helper: ¿Es zona de administrador? ────────────────────────────────────────
// Solo la zona ADMIN* (ADMNISTRATIVO, ADMINISTRATIVO) bypasea todos los filtros.
function _esZonaAdmin(zona) {
  var z = (zona || '').toUpperCase();
  return z.indexOf('ADMIN') >= 0;
}

// ── Helper: Obtener zona de un trabajador desde trabajadoresHoy ─────────────
function _getZonaDeTrabajador(dni, trabajadoresHoy) {
  var zonas = Object.keys(trabajadoresHoy);
  for (var i = 0; i < zonas.length; i++) {
    var workers = trabajadoresHoy[zonas[i]];
    for (var j = 0; j < workers.length; j++) {
      if (workers[j].dni === dni) return zonas[i];
    }
  }
  return null;
}

// ── Helper: Obtener { zona, turno } de un trabajador ──────────────────────
// El turno/sub puede contener el nombre de la posta (ej. "POSTA SUR", "POSTA C1")
// cuando la estructura del Rol usa subs como puntos de inspección.
function _getPuestoDeTrabajador(dni, trabajadoresHoy) {
  var zonas = Object.keys(trabajadoresHoy);
  for (var i = 0; i < zonas.length; i++) {
    var workers = trabajadoresHoy[zonas[i]];
    for (var j = 0; j < workers.length; j++) {
      if (workers[j].dni === dni) {
        return { zona: zonas[i], turno: String(workers[j].turno || '').trim().toUpperCase() };
      }
    }
  }
  return null;
}

// ============================================================
//  HELPER: Capacitaciones pendientes del mes actual (PASSO)
//  Devuelve { pendientesPorDni: {dni: [{cap,freq}]}, resumen: [{cap,freq,total,completaron,faltan,faltanNombres}] }
// ============================================================
function _getCapacitacionesPendientes() {
  var result = { pendientesPorDni: {}, resumen: [] };
  try {
    var ssCap = getSpreadsheetCapacitaciones();
    var hojaMatriz = ssCap.getSheetByName('Matriz');
    var hojaBD     = ssCap.getSheetByName('B DATOS');
    var mesActual  = new Date().getMonth(); // 0-based
    var anioActual = new Date().getFullYear();

    // Leer estructura Matriz
    var matrizLastRow = hojaMatriz.getLastRow();
    var fila17 = hojaMatriz.getRange(17, 1, 1, hojaMatriz.getLastColumn()).getValues()[0];
    var lastCol = 0;
    for (var i = fila17.length - 1; i >= 0; i--) {
      if (fila17[i] && String(fila17[i]).trim() !== '') { lastCol = i + 1; break; }
    }
    if (lastCol < 5) lastCol = hojaMatriz.getLastColumn();

    var capacitaciones    = hojaMatriz.getRange(17, 5, 1, lastCol - 4).getValues()[0];
    var fechasProgramadas = hojaMatriz.getRange(8,  5, 1, lastCol - 4).getValues()[0];
    var vigencias         = hojaMatriz.getRange(9,  5, 1, lastCol - 4).getValues()[0];
    var esRecurrentes     = hojaMatriz.getRange(10, 5, 1, lastCol - 4).getValues()[0];

    // Leer filas 18+ de la Matriz: col D = nombre del cargo, cols E+ = habilitación por capacitación
    // Un valor no vacío / TRUE / "X" / "SI" en la celda indica que ese cargo TIENE esa capacitación activa.
    // cargosHabilitadosPorCap[capNorm] = Set-like object { cargoNorm: true }
    var cargosHabilitadosPorCap = {}; // { capNombreNorm: { cargoNorm: true } }
    if (matrizLastRow >= 18) {
      var cargoFilas = hojaMatriz.getRange(18, 1, matrizLastRow - 17, lastCol).getValues();
      cargoFilas.forEach(function(fila) {
        var cargoNom = String(fila[3] || '').trim().toLowerCase(); // col D (índice 3)
        if (!cargoNom) return;
        for (var ci = 0; ci < capacitaciones.length; ci++) {
          var capNom = String(capacitaciones[ci] || '').trim().toUpperCase();
          if (!capNom) continue;
          var celda = fila[4 + ci]; // col E en adelante
          var habilitado = celda === true || celda === 1 ||
                           String(celda).trim().toUpperCase() === 'VERDADERO' ||
                           String(celda).trim().toUpperCase() === 'TRUE' ||
                           String(celda).trim().toUpperCase() === 'X' ||
                           String(celda).trim().toUpperCase() === 'SI' ||
                           String(celda).trim().toUpperCase() === 'SÍ';
          if (habilitado) {
            if (!cargosHabilitadosPorCap[capNom]) cargosHabilitadosPorCap[capNom] = {};
            cargosHabilitadosPorCap[capNom][cargoNom] = true;
          }
        }
      });
    }

    // Construir lista de capacitaciones programadas para el mes actual
    var capsMesActual = []; // [{nombre, nombreNorm, vigMeses, esRecurrente}]
    capacitaciones.forEach(function(cap, idx) {
      if (!cap || String(cap).trim() === '') return;
      var nombre     = String(cap).trim();
      var nombreNorm = nombre.toUpperCase();
      var vigMeses   = parseInt(vigencias[idx]) || 12;
      var rec        = esRecurrentes[idx];
      var esRec      = (rec === true || String(rec).toUpperCase() === 'VERDADERO');

      // Calcular si está programada para mesActual
      var fechaProg = fechasProgramadas[idx];
      if (typeof fechaProg === 'string' && fechaProg.trim() !== '') {
        var mf = fechaProg.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mf) fechaProg = new Date(parseInt(mf[3]), parseInt(mf[2]) - 1, parseInt(mf[1]));
      }
      if (!(fechaProg instanceof Date) || isNaN(fechaProg.getTime())) return;

      var mesInicio = fechaProg.getMonth();
      var mesesProg = [];
      for (var m = mesInicio; m < 12; m += vigMeses) mesesProg.push(m);
      for (var m2 = mesInicio - vigMeses; m2 >= 0; m2 -= vigMeses) mesesProg.push(m2);

      if (mesesProg.indexOf(mesActual) === -1) return; // No programada este mes

      var labels = {1:'MENSUAL',2:'BIMESTRAL',3:'TRIMESTRAL',6:'SEMESTRAL',12:'ANUAL'};
      capsMesActual.push({
        nombre: nombre,
        nombreNorm: nombreNorm,
        vigMeses: vigMeses,
        esRecurrente: esRec,
        frecuencia: labels[vigMeses] || ('CADA ' + vigMeses + ' MESES')
      });
    });

    if (capsMesActual.length === 0) return result;

    // Leer trabajadores activos (cols A-G + col F situación + col P autorizado)
    var hojaPersonal = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    var lastRowPers  = hojaPersonal.getLastRow();
    var dnisActivos  = {}; // { dni: { nombre, cargo } }
    if (lastRowPers > 1) {
      var personalData = hojaPersonal.getRange(2, 1, lastRowPers - 1, 16).getValues();
      for (var pi = 0; pi < personalData.length; pi++) {
        var situacion = String(personalData[pi][5]  || '').trim().toUpperCase(); // col F: ACTIVO/LIQUIDADO
        var autorizado = String(personalData[pi][15] || '').trim().toUpperCase(); // col P: SI/NO
        var dni  = String(personalData[pi][1] || '').trim(); // col B
        var nom  = String(personalData[pi][2] || '').trim(); // col C
        var cargo = String(personalData[pi][6] || '').trim().toLowerCase(); // col G
        // Solo trabajadores activos y autorizados
        if (dni && situacion !== 'LIQUIDADO' && autorizado !== 'NO') {
          dnisActivos[dni] = { nombre: nom, cargo: cargo };
        }
      }
    }

    // Leer B DATOS capacitaciones — activaciones y completados
    var bdLast = hojaBD.getLastRow();
    var bdData = bdLast > 1 ? hojaBD.getRange(2, 1, bdLast - 1, 16).getValues() : [];

    // Construir activaciones para recurrentes
    var activacionesMap = {};
    bdData.forEach(function(row) {
      var esAct = String(row[9]).trim().toUpperCase() === 'ACTIVACION' ||
                  String(row[0]).trim().toUpperCase() === 'ACTIVACION';
      if (!esAct) return;
      var temaNorm = String(row[4]).trim().toUpperCase();
      var cicloStr = String(row[15] || '');
      var partes   = cicloStr.split(' - ');
      if (partes.length !== 2) return;
      var parseD = function(s) {
        var m = s.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        return m ? new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) : null;
      };
      var ini = parseD(partes[0]);
      var fin = parseD(partes[1]);
      if (ini && fin) {
        if (!activacionesMap[temaNorm]) activacionesMap[temaNorm] = [];
        activacionesMap[temaNorm].push({ inicio: ini, fin: fin });
      }
    });

    // Para cada capacitación del mes, encontrar quién la completó
    capsMesActual.forEach(function(capInfo) {
      var dnisCompletaron = {};
      var hoy = new Date();

      // Buscar ciclo activo si es recurrente
      var inicioCiclo = null, finCiclo = null;
      if (capInfo.esRecurrente) {
        var ciclos = activacionesMap[capInfo.nombreNorm] || [];
        for (var ci = 0; ci < ciclos.length; ci++) {
          if (hoy >= ciclos[ci].inicio && hoy <= ciclos[ci].fin) {
            inicioCiclo = ciclos[ci].inicio;
            finCiclo    = ciclos[ci].fin;
            break;
          }
        }
      }

      bdData.forEach(function(row) {
        var dniBD    = String(row[0] || '').trim();
        var estadoBD = String(row[9] || '').trim().toUpperCase();
        if (estadoBD === 'ACTIVACION' || dniBD.toUpperCase() === 'ACTIVACION') return;
        var temaBD = String(row[4] || '').trim().toUpperCase();
        var fecha  = row[7];
        // Col O (index 14): solo contar "Matriz" para capacitaciones programadas
        var origenBD = String(row[14] || '').trim().toUpperCase();
        if (origenBD !== 'MATRIZ') return;
        if (!dniBD || !dnisActivos[dniBD] || temaBD !== capInfo.nombreNorm || estadoBD !== 'APROBADO' || !(fecha instanceof Date)) return;

        var enPeriodo;
        if (capInfo.esRecurrente && inicioCiclo && finCiclo) {
          enPeriodo = fecha >= inicioCiclo && fecha <= finCiclo;
        } else {
          enPeriodo = fecha.getMonth() === mesActual && fecha.getFullYear() === anioActual;
        }
        if (enPeriodo) dnisCompletaron[dniBD] = true;
      });

      // Determinar trabajadores destinatarios: solo cargos habilitados en la Matriz
      var cargosHabCap = cargosHabilitadosPorCap[capInfo.nombreNorm] || null;
      // Si no hay ningún cargo habilitado en la Matriz para esta cap, se asume que aplica a todos
      var filtrarPorCargo = cargosHabCap && Object.keys(cargosHabCap).length > 0;

      var dnisDestinatarios = Object.keys(dnisActivos).filter(function(dni) {
        if (!filtrarPorCargo) return true;
        var cargoDni = (dnisActivos[dni].cargo || '').toLowerCase().trim();
        // Verificar si algún cargo habilitado coincide con el cargo del trabajador
        return Object.keys(cargosHabCap).some(function(cargoHab) {
          return cargoDni === cargoHab ||
                 cargoDni.indexOf(cargoHab) >= 0 ||
                 cargoHab.indexOf(cargoDni.split(' ')[0]) >= 0;
        });
      });

      var totalActivos     = dnisDestinatarios.length;
      var totalCompletaron = Object.keys(dnisCompletaron).length;
      var faltanDnis       = [];
      var faltanNombres    = [];

      dnisDestinatarios.forEach(function(dni) {
        if (!dnisCompletaron[dni]) {
          faltanDnis.push(dni);
          faltanNombres.push((dnisActivos[dni] && dnisActivos[dni].nombre) || dni);
        }
      });

      if (faltanDnis.length > 0) {
        result.resumen.push({
          capacitacion: capInfo.nombre,
          frecuencia:   capInfo.frecuencia,
          total:        totalActivos,
          completaron:  totalCompletaron,
          faltan:       faltanDnis.length,
          faltanDnis:   faltanDnis,
          faltanNombres: faltanNombres
        });

        faltanDnis.forEach(function(dni) {
          if (!result.pendientesPorDni[dni]) result.pendientesPorDni[dni] = [];
          result.pendientesPorDni[dni].push({ capacitacion: capInfo.nombre, frecuencia: capInfo.frecuencia });
        });
      }
    });

  } catch(e) {
    Logger.log('Error en _getCapacitacionesPendientes: ' + e.message);
  }
  return result;
}

// ============================================================
//  HELPER: Trabajadores en turno para una fecha ISO (yyyy-MM-dd)
//  Lee rol_turnos.json (Drive) — fuente real de los turnos
//  Devuelve { trabajadoresHoy: {zona:[{nombre,turno,dni}]}, dnisEnTurno: [] }
// ============================================================
function _getTrabajadoresPorFecha(fechaISO) {
  var trabajadoresHoy = {};
  var dnisEnTurno     = [];
  try {
    var plannerStr = loadPlannerFromDrive();
    if (!plannerStr) { Logger.log('⚠️ rol_turnos.json vacío'); return { trabajadoresHoy: trabajadoresHoy, dnisEnTurno: dnisEnTurno }; }

    var planner = JSON.parse(plannerStr);
    var dbPlan  = planner.db       || {};
    var struct  = planner.structure || [];

    // Mapa empId → { dni, fullName }
    var empsRaw = JSON.parse(getEmployeesFromDB());
    var empById = {};
    empsRaw.forEach(function(e) { empById[String(e.id)] = e; });

    Object.keys(dbPlan).forEach(function(cellKey) {
      // cellKey = "2026-03-06_subId" (subId puede contener guiones bajos)
      var underIdx = cellKey.indexOf('_');
      if (underIdx === -1) return;
      var dateStr = cellKey.substring(0, underIdx);
      if (dateStr !== fechaISO) return;

      var subId = cellKey.substring(underIdx + 1);

      // Resolver zona y nombre del turno desde structure
      var zoneName  = 'GENERAL';
      var shiftName = subId;
      for (var ai = 0; ai < struct.length; ai++) {
        var area = struct[ai];
        var subs = area.subs || [];
        for (var si = 0; si < subs.length; si++) {
          if (subs[si].id === subId) {
            zoneName  = String(area.name || 'GENERAL').trim().toUpperCase();
            shiftName = String(subs[si].name || subId).trim();
            break;
          }
        }
      }

      var items = dbPlan[cellKey];
      for (var ii = 0; ii < items.length; ii++) {
        var item = items[ii];
        if (item.type !== 'WORK') continue;        // ignorar ABSENCE, LOCKED, AVAILABILITY
        if (!item.hours || item.hours <= 0) continue;
        var emp = empById[String(item.empId)];
        if (!emp) continue;

        var dni    = String(emp.dni      || '').trim();
        var nombre = String(emp.fullName || emp.name || '').trim();

        if (!trabajadoresHoy[zoneName]) trabajadoresHoy[zoneName] = [];
        trabajadoresHoy[zoneName].push({ nombre: nombre, turno: shiftName, dni: dni });
        if (dni && dnisEnTurno.indexOf(dni) === -1) dnisEnTurno.push(dni);
      }
    });

    Logger.log('✅ Turno hoy (' + fechaISO + '): ' + dnisEnTurno.length + ' trabajadores en ' + Object.keys(trabajadoresHoy).length + ' zonas');
  } catch(e) {
    Logger.log('❌ _getTrabajadoresPorFecha: ' + e.message);
  }
  return { trabajadoresHoy: trabajadoresHoy, dnisEnTurno: dnisEnTurno };
}
// ============================================================
//  SIMULADOR DE ALERTAS — devuelve el diagnóstico SIN enviar
//  Llamado desde el frontend vía google.script.run
// ============================================================
function simularAlertasInspeccion() {
  try {
    var ahora   = new Date();
    var hoyStr  = Utilities.formatDate(ahora, TZ_ALERTAS, 'yyyy-MM-dd');
    var mesStr  = Utilities.formatDate(ahora, TZ_ALERTAS, 'yyyy-MM');
    var hoyDate = _soloFechaInsp(ahora);

    // ── 1. INVENTARIO ────────────────────────────────────────────────────
    var checkSS  = getCheckSpreadsheet();
    var invSheet = checkSS.getSheetByName('INVENTARIO');
    var invLast  = invSheet.getLastRow();
    if (invLast < 3) return { ok: false, error: 'INVENTARIO vacío' };

    var invData = invSheet.getRange(3, 1, invLast - 2, 16).getValues();
    var equipos = [];
    for (var i = 0; i < invData.length; i++) {
      var r      = invData[i];
      var codigo = String(r[4] || '').trim();
      var equipo = String(r[3] || '').trim();
      var estado = String(r[15] || '').trim().toLowerCase();
      var freq   = _parseDiasFreqInsp(r[12]);
      var lugaresRaw = String(r[11] || '').trim(); // Col L (index 11) = lugares
      var lugares = lugaresRaw
        ? lugaresRaw.split(',').map(function(l){ return l.trim().toUpperCase(); }).filter(Boolean)
        : [];
      if (!codigo || !equipo || estado === 'retirado') continue;
      if (isNaN(freq) || freq <= 0) continue;
      equipos.push({
        empresa: String(r[1] || '').trim(),
        area:    String(r[2] || '').trim(),
        equipo:  equipo,
        codigo:  codigo,
        cargo:   String(r[5] || '').trim(),
        diasFrecuencia: freq,
        lugares: lugares
      });
    }

    // ── 2. B DATOS → última inspección por código||lugar ────────────────────────
    var bSheet = checkSS.getSheetByName('B DATOS');
    var bLast  = bSheet.getLastRow();
    var bData  = bLast > 1 ? bSheet.getRange(2, 1, bLast - 1, 15).getValues() : [];
    var ultimaInsp = {};
    for (var j = 0; j < bData.length; j++) {
      var bd = bData[j];
      var bdCod   = String(bd[3] || '').trim();
      var bdFecha = bd[9];
      if (!bdCod || !(bdFecha instanceof Date)) continue;
      var bdLugar = String(bd[7] || '').trim().toUpperCase(); // Col H (index 7) = lugar
      var bdKey   = bdLugar ? (bdCod + '||' + bdLugar) : bdCod;
      if (!ultimaInsp[bdKey] || bdFecha > ultimaInsp[bdKey]) ultimaInsp[bdKey] = new Date(bdFecha);
    }

    // ── 3. Calcular estado de cada equipo (por ubicación si aplica) ────────────────────────────────
    var TZ = TZ_ALERTAS;
    var fmt = function(d) { return d ? Utilities.formatDate(d, TZ, 'dd/MM/yyyy') : 'Sin registro'; };

    var vencidos = [], proximos = [], alDia = [];

    equipos.forEach(function(eq) {
      var periodo = _freqEsCalendario(eq.diasFrecuencia);
      var tipo    = periodo === 'semana'    ? 'SEMANAL'
                  : periodo === 'mes'       ? 'MENSUAL'
                  : periodo === 'bimestre'  ? 'BIMESTRAL'
                  : periodo === 'trimestre' ? 'TRIMESTRAL'
                  : periodo === 'semestre'  ? 'SEMESTRAL'
                  : periodo === 'anio'      ? 'ANUAL'
                  : 'CADA ' + eq.diasFrecuencia + 'd';

      // Iterar por cada lugar (o una vez si no hay lugares)
      var lugaresCheck = eq.lugares.length > 0 ? eq.lugares : [null];
      lugaresCheck.forEach(function(lugar) {
        var eqKey  = lugar ? (eq.codigo + '||' + lugar) : eq.codigo;
        var ultima = ultimaInsp[eqKey] || null;
        var diasVencido;

        if (!ultima) {
          var inicioPer = _inicioPeriodoActualInsp(periodo || 'mes', hoyDate);
          diasVencido = Math.floor((hoyDate - inicioPer) / 86400000);
        } else {
          diasVencido = _diasVencidoInsp(eq.diasFrecuencia, _soloFechaInsp(ultima), hoyDate);
        }

        // Calcular próxima fecha (para mostrar en "próximos")
        var proxFecha = '';
        if (ultima && !periodo) {
          var nd = new Date(_soloFechaInsp(ultima));
          nd.setDate(nd.getDate() + eq.diasFrecuencia);
          proxFecha = fmt(nd);
        }

        var item = {
          equipo:       eq.equipo + (lugar ? ' — ' + lugar : ''),
          codigo:       eq.codigo,
          area:         eq.area,
          empresa:      eq.empresa,
          tipo:         tipo,
          diasFreq:     eq.diasFrecuencia,
          ultimaFecha:  fmt(ultima),
          proxFecha:    proxFecha,
          diasVencido:  diasVencido,
          cargo:        eq.cargo,
          lugar:        lugar || ''
        };

        if (diasVencido < 0) {
          // Faltan días: dentro de 0-3 días = próximo
          if (diasVencido >= -3) proximos.push(item);
          else alDia.push(item);
        } else {
          vencidos.push(item);
        }
      });
    });

    vencidos.sort(function(a,b){ return b.diasVencido - a.diasVencido; });
    proximos.sort(function(a,b){ return b.diasVencido - a.diasVencido; }); // diasVencido negativo, más cercano a 0 primero

    // ── 4. Trabajadores en turno HOY (desde rol_turnos.json) ─────────────
    var _turnoData2  = _getTrabajadoresPorFecha(hoyStr);
    var trabajadoresHoy = _turnoData2.trabajadoresHoy;
    var dnisEnTurno     = _turnoData2.dnisEnTurno;


    // ── 5. Mapear qué trabajador recibiría push por equipo vencido ────────
    // Filtros combinados: ZONA/POSTA + CARGO. Admin recibe TODO.
    var cargoPorDni2 = _getCargoPorDniInsp(); // { dni: cargo_lower }
    var pushPorTrabajador = {}; // { "nombre [zona/turno]": [equipos] }

    vencidos.forEach(function(it) {
      var lugarAlerta = (it.lugar || '').toUpperCase();
      var cargosEquipo = String(it.cargo || '').split(',').map(function(c) { return c.trim().toLowerCase(); }).filter(Boolean);
      dnisEnTurno.forEach(function(dni) {
        var puesto2 = _getPuestoDeTrabajador(dni, trabajadoresHoy);
        if (!puesto2) return;
        var esAdmin = _esZonaAdmin(puesto2.zona);
        if (!esAdmin) {
          if (!lugarAlerta) return;
          // Filtro posta: zona O turno/sub del trabajador debe coincidir con el lugar
          var matchLugar2 = _zonaMatchesLugar(puesto2.zona, lugarAlerta) ||
                            (puesto2.turno && _zonaMatchesLugar(puesto2.turno, lugarAlerta));
          if (!matchLugar2) return;
          // Filtro cargo
          if (cargosEquipo.length > 0) {
            var dniCargo = (cargoPorDni2[dni] || '').toLowerCase();
            var matchCargo = cargosEquipo.some(function(c) {
              return dniCargo.indexOf(c) >= 0 || c.indexOf(dniCargo.split(' ')[0]) >= 0;
            });
            if (!matchCargo) return;
          }
        }
        // Buscar nombre — mostrar turno/posta si es informativo
        var nombre = dni;
        var etiqueta = puesto2.turno || puesto2.zona;
        (trabajadoresHoy[puesto2.zona] || []).forEach(function(w) {
          if (w.dni === dni) nombre = w.nombre + ' [' + etiqueta + ']';
        });
        if (!pushPorTrabajador[nombre]) pushPorTrabajador[nombre] = [];
        pushPorTrabajador[nombre].push(it.equipo + ' (' + it.codigo + ')');
      });
    });

    // ── 5b. Capacitaciones pendientes (PASSO) ─────────────────────────────
    var passoPendientes = _getCapacitacionesPendientes();
    // Agregar al push de cada trabajador en turno sus capacitaciones pendientes
    var passoPorTrabajador = {}; // { "nombre [zona]": [{cap,freq}] } — solo para el simulador
    dnisEnTurno.forEach(function(dni) {
      var caps = passoPendientes.pendientesPorDni[dni];
      if (!caps || caps.length === 0) return;
      var zonaWorker = _getZonaDeTrabajador(dni, trabajadoresHoy);
      var nombre = dni;
      if (zonaWorker) {
        (trabajadoresHoy[zonaWorker] || []).forEach(function(w) {
          if (w.dni === dni) nombre = w.nombre + ' [' + zonaWorker + ']';
        });
      }
      passoPorTrabajador[nombre] = caps;
    });

    // ── 6. Construir preview del mensaje Telegram ────────────────────────
    var telegramPreview = [];
    if (vencidos.length > 0) {
      telegramPreview.push('🔔 ALERTA — INSPECCIONES PENDIENTES');
      telegramPreview.push('⏰ ' + Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy HH:mm') + ' hrs');
      telegramPreview.push('⚠️ ' + vencidos.length + ' inspección(es) vencida(s)');
      telegramPreview.push('');
      var porArea = {};
      vencidos.forEach(function(a) {
        var k = (a.area || 'SIN ÁREA').toUpperCase();
        if (!porArea[k]) porArea[k] = [];
        porArea[k].push(a);
      });
      Object.keys(porArea).forEach(function(areaKey) {
        var trabAr = trabajadoresHoy[areaKey] || trabajadoresHoy['GENERAL'] || [];
        telegramPreview.push('━━━━━━━━━━━━━━━');
        telegramPreview.push('📍 ' + areaKey);
        telegramPreview.push('👷 En turno: ' + (trabAr.length ? trabAr.map(function(w){ return w.nombre; }).join(', ') : 'Sin personal registrado hoy'));
        porArea[areaKey].forEach(function(it) {
          var urg = it.diasVencido === 0 ? '🟡 HOY' : it.diasVencido <= 2 ? '🟠 +' + it.diasVencido + 'd' : '🔴 +' + it.diasVencido + 'd';
          telegramPreview.push(urg + ' — ' + it.codigo + ' ' + it.equipo + ' | ' + it.tipo + ' | Última: ' + it.ultimaFecha);
        });
      });
    } else {
      telegramPreview.push('✅ Sin inspecciones vencidas hoy.');
    }

    return {
      ok:              true,
      fecha:           Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy HH:mm'),
      totalEquipos:    equipos.length,
      vencidos:        vencidos,
      proximos:        proximos,
      alDia:           alDia.length,
      trabajadoresHoy: trabajadoresHoy,
      dnisEnTurno:     dnisEnTurno.length,
      pushPorTrabajador: pushPorTrabajador,
      telegramPreview: telegramPreview.join('\n'),
      // PASSO — capacitaciones pendientes
      passoResumen:        passoPendientes.resumen,       // [{cap,freq,total,completaron,faltan,faltanNombres}]
      passoPorTrabajador:  passoPorTrabajador             // {"nombre [zona]": [{cap,freq}]}
    };

  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}