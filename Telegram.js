/**
 * ============================================
 * MÓDULO DE NOTIFICACIONES TELEGRAM
 * ============================================
 * Sistema centralizado para enviar notificaciones
 * desde Google Apps Script a Telegram
 */

// ========== CONFIGURACIÓN ==========
// Los valores se leen desde Script Properties (TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID).
// Los hardcodeados son fallback temporal — configura las Script Properties y elimínalos.
const _tgProps = PropertiesService.getScriptProperties();
const TELEGRAM_CONFIG = {
  botToken: _tgProps.getProperty('TELEGRAM_BOT_TOKEN') || '8316348321:AAHyx9OczZdtoNuYi8OzPXx868c1tzhhwmc',
  chatId:   _tgProps.getProperty('TELEGRAM_CHAT_ID')   || '6725665354',
  apiUrl: 'https://api.telegram.org/bot'
};

/**
 * Función principal para enviar mensajes a Telegram
 * @param {string} mensaje - Texto del mensaje
 * @param {Object} opciones - Opciones adicionales (chatId, parseMode, etc)
 * @return {Object} Resultado del envío
 */
function enviarTelegram(mensaje, opciones = {}) {
  try {
    const chatId = opciones.chatId || TELEGRAM_CONFIG.chatId;
    const parseMode = opciones.parseMode || 'HTML';
    const disableNotification = opciones.silencioso || false;
    
    const url = `${TELEGRAM_CONFIG.apiUrl}${TELEGRAM_CONFIG.botToken}/sendMessage`;
    
    const payload = {
      chat_id: chatId,
      text: mensaje,
      parse_mode: parseMode,
      disable_notification: disableNotification
    };
    
    // Si hay botones inline
    if (opciones.botones) {
      payload.reply_markup = {
        inline_keyboard: opciones.botones
      };
    }
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      return { success: true, messageId: result.result.message_id };
    } else {
      Logger.log('Error Telegram: ' + result.description);
      return { success: false, error: result.description };
    }
    
  } catch (error) {
    Logger.log('Error enviando a Telegram: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Enviar documento/archivo a Telegram
 * @param {string} fileId - ID del archivo en Google Drive
 * @param {string} caption - Descripción del archivo
 */
function enviarDocumentoTelegram(fileId, caption = '') {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    
    const url = `${TELEGRAM_CONFIG.apiUrl}${TELEGRAM_CONFIG.botToken}/sendDocument`;
    
    const formData = {
      chat_id: TELEGRAM_CONFIG.chatId,
      caption: caption
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: {
        ...formData,
        document: blob
      },
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    return result.ok ? { success: true } : { success: false, error: result.description };
    
  } catch (error) {
    Logger.log('Error enviando documento: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Enviar imagen a Telegram
 * @param {string} imageUrl - URL de la imagen o ID de Drive
 * @param {string} caption - Descripción de la imagen
 */
function enviarImagenTelegram(imageUrl, caption = '') {
  try {
    const url = `${TELEGRAM_CONFIG.apiUrl}${TELEGRAM_CONFIG.botToken}/sendPhoto`;
    
    const payload = {
      chat_id: TELEGRAM_CONFIG.chatId,
      photo: imageUrl,
      caption: caption
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    return result.ok ? { success: true } : { success: false, error: result.description };
    
  } catch (error) {
    Logger.log('Error enviando imagen: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ========== NOTIFICACIONES ESPECÍFICAS DEL SISTEMA ==========

/**
 * Notificar nuevo usuario registrado
 */
function notificarNuevoUsuario(datos) {
  const mensaje = `
🆕 <b>Nuevo Usuario Registrado</b>

👤 <b>Nombre:</b> ${datos.nombre}
🆔 <b>Usuario:</b> ${datos.usuario}
🏢 <b>Empresa:</b> ${datos.empresa}
💼 <b>Cargo:</b> ${datos.cargo}
📧 <b>Email:</b> ${datos.email}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

<i>Sistema BIOX-SIG</i>
  `.trim();
  
  const botones = [[
    { text: '👁️ Ver Usuario', url: 'https://www.iassoma.com/ccl' }
  ]];
  
  return enviarTelegram(mensaje, { botones });
}

/**
 * Notificar usuario actualizado
 */
function notificarUsuarioActualizado(datos) {
  const mensaje = `
✏️ <b>Usuario Actualizado</b>

👤 <b>Usuario:</b> ${datos.usuario}
📝 <b>Cambios realizados por:</b> ${datos.modificadoPor || 'Sistema'}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificar usuario eliminado
 */
function notificarUsuarioEliminado(usuario) {
  const mensaje = `
🗑️ <b>Usuario Eliminado</b>

👤 <b>Usuario:</b> ${usuario}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

⚠️ <i>Esta acción es irreversible</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificar nuevo registro HHT
 */
function notificarNuevoHHT(datos) {
  const mensaje = `
📋 <b>Nuevo Registro HHT</b>

🆔 <b>ID:</b> ${datos.id || 'N/A'}
📍 <b>Lugar:</b> ${datos.lugar || 'N/A'}
👷 <b>Inspector:</b> ${datos.inspector || 'N/A'}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

<i>Sistema BIOX-SIG - HHT</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificar nuevo desvío detectado
 */
function notificarDesvio(datos) {
  const mensaje = `
⚠️ <b>NUEVO DESVÍO DETECTADO</b>

📋 <b>Tipo:</b> ${datos.tipo || 'N/A'}
📍 <b>Lugar:</b> ${datos.lugar || 'N/A'}
🔴 <b>Nivel:</b> ${datos.nivel || 'N/A'}
👤 <b>Reportado por:</b> ${datos.reportadoPor || 'N/A'}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

${datos.descripcion ? '📝 <b>Descripción:</b>\n' + datos.descripcion : ''}

<i>Requiere atención inmediata</i>
  `.trim();
  
  const botones = [[
    { text: '🔍 Ver Detalles', url: 'https://www.iassoma.com/ccl' }
  ]];
  
  return enviarTelegram(mensaje, { botones });
}

/**
 * Notificar capacitación programada
 */
function notificarCapacitacion(datos) {
  const mensaje = `
📚 <b>Capacitación Programada</b>

📖 <b>Tema:</b> ${datos.tema || 'N/A'}
👥 <b>Participantes:</b> ${datos.participantes || 'N/A'}
📍 <b>Lugar:</b> ${datos.lugar || 'N/A'}
📅 <b>Fecha:</b> ${datos.fecha || 'N/A'}
⏰ <b>Hora:</b> ${datos.hora || 'N/A'}

<i>Sistema BIOX-SIG - Capacitaciones</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificar checklist completado
 */
function notificarChecklistCompletado(datos) {
  const mensaje = `
✅ <b>Checklist Completado</b>

📋 <b>Tipo:</b> ${datos.tipo || 'N/A'}
👤 <b>Completado por:</b> ${datos.usuario || 'N/A'}
📊 <b>Resultado:</b> ${datos.resultado || 'N/A'}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

<i>Sistema BIOX-SIG</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificar evento/incidente
 */
function notificarEvento(datos) {
  const severidad = datos.severidad || 'MEDIA';
  const emoji = severidad === 'ALTA' ? '🔴' : severidad === 'MEDIA' ? '🟡' : '🟢';
  
  const mensaje = `
${emoji} <b>EVENTO REGISTRADO</b>

📋 <b>Tipo:</b> ${datos.tipo || 'N/A'}
🔴 <b>Severidad:</b> ${severidad}
📍 <b>Lugar:</b> ${datos.lugar || 'N/A'}
👤 <b>Reportado por:</b> ${datos.reportadoPor || 'N/A'}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

${datos.descripcion ? '📝 <b>Descripción:</b>\n' + datos.descripcion : ''}

<i>Sistema BIOX-SIG - Eventos</i>
  `.trim();
  
  const botones = [[
    { text: '📊 Ver Dashboard', url: 'https://www.iassoma.com/ccl' }
  ]];
  
  return enviarTelegram(mensaje, { botones });
}

/**
 * Notificar EPP asignado/entregado
 */
function notificarEPP(datos) {
  const mensaje = `
🦺 <b>EPP Registrado</b>

👤 <b>Trabajador:</b> ${datos.trabajador || 'N/A'}
🛡️ <b>Equipo:</b> ${datos.equipo || 'N/A'}
📦 <b>Cantidad:</b> ${datos.cantidad || 'N/A'}
📅 <b>Fecha entrega:</b> ${new Date().toLocaleDateString('es-PE')}

<i>Sistema BIOX-SIG - EPP</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificar login de usuario (opcional, puede ser muy frecuente)
 */
function notificarLogin(nombre) {
  const mensaje = `
🔐 <b>Inicio de Sesión</b>

👤 ${nombre}
📅 ${new Date().toLocaleString('es-PE')}
  `.trim();
  
  // Enviar silenciosamente para no molestar
  return enviarTelegram(mensaje, { silencioso: true });
}

/**
 * Notificar actualización de mapa de riesgos
 */
function notificarMapaRiesgos(datos) {
  const accion = datos.accion || 'actualizado'; // creado, actualizado, eliminado
  
  const mensaje = `
🗺️ <b>Mapa de Riesgos ${accion}</b>

📋 <b>Título:</b> ${datos.titulo || 'N/A'}
👤 <b>Autor:</b> ${datos.autor || 'N/A'}
📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

<i>Sistema BIOX-SIG - Mapas</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Notificación de error del sistema (para debugging)
 */
function notificarError(error, contexto = '') {
  const mensaje = `
🚨 <b>ERROR DEL SISTEMA</b>

⚠️ <b>Contexto:</b> ${contexto}
📝 <b>Error:</b> ${error.toString()}
📅 <b>Fecha:</b> ${new Date().toLocaleString('es-PE')}

<i>Requiere atención técnica</i>
  `.trim();
  
  return enviarTelegram(mensaje);
}

/**
 * Resumen diario automático (llamar con trigger)
 */
function enviarResumenDiario() {
  try {
    // Obtener estadísticas del día
    const hoja = getSpreadsheetPersonal().getSheetByName('Log');
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    
    const data = hoja.getDataRange().getValues();
    const loginsDia = data.filter(row => {
      const fecha = new Date(row[0]);
      return fecha >= hoy;
    }).length;
    
    const mensaje = `
📊 <b>Resumen Diario - BIOX-SIG</b>

📅 <b>Fecha:</b> ${new Date().toLocaleDateString('es-PE')}

📈 <b>Estadísticas del día:</b>
• 🔐 Inicios de sesión: ${loginsDia}

<i>Sistema BIOX-SIG</i>
    `.trim();
    
    return enviarTelegram(mensaje);
    
  } catch (error) {
    Logger.log('Error en resumen diario: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Configurar trigger para resumen diario
 * Ejecutar esta función una vez para configurar
 */
function configurarTriggerDiario() {
  // Eliminar triggers existentes para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarResumenDiario') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger para las 8:00 AM
  ScriptApp.newTrigger('enviarResumenDiario')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  
  Logger.log('Trigger diario configurado exitosamente');
}

/**
 * Función de prueba
 */
function testTelegram() {
  const mensaje = `
🧪 <b>Prueba de Conexión</b>

✅ El sistema de notificaciones Telegram está funcionando correctamente.

📅 ${new Date().toLocaleString('es-PE')}
  `.trim();
  
  const resultado = enviarTelegram(mensaje);
  Logger.log('Resultado prueba: ' + JSON.stringify(resultado));
  return resultado;
}