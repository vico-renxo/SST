/**
 * Obtiene alertas de vencimiento de EPP para el trabajador logueado.
 *
 * ALGORITMO:
 *   Por cada producto (clave = dni|producto, sin variante), toma la entrega
 *   MÁS RECIENTE (por fecha de entrega).
 *   - Si esa entrega no tiene fecha de vencimiento → cubierto, sin alerta.
 *   - Si tiene fecha futura → cubierto, sin alerta.
 *   - Si está vencida (≤0 días) → VENCIDO.
 *   - Si vence en ≤15 días → VENCE EN N DÍAS.
 *
 * Agrupa por producto SIN variante: una entrega nueva de cualquier variante
 * del mismo producto cancela la alerta de variantes anteriores.
 *
 * Filtra por MATRIZ: solo productos asignados al cargo actual del trabajador.
 * Admin (ADMIN_EMAIL en ScriptProperties) omite los filtros de DNI y MATRIZ.
 */
function obtenerAlertasVencimientos(dniLogin) {
  try {
    const ss     = getSpreadsheetEPP();
    const shReg  = ss.getSheetByName(SHEPP.REGISTRO);
    const correoActual = Session.getActiveUser().getEmail();

    // ADMIN_EMAIL solo desde ScriptProperties — sin fallback al correo activo
    // (el fallback haría esAdmin=true para todos los usuarios)
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || '';
    const esAdmin = (adminEmail !== '' && correoActual === adminEmail) || !dniLogin;

    const data = shReg.getDataRange().getValues();
    const hoy  = new Date();

    const COL_DNI      = IDX.REG.DNI - 1;
    const COL_PRODUCTO = IDX.REG.PRODUCTO - 1;
    const COL_VENC     = IDX.REG.FECHA_VENCIMIENTO - 1;
    const COL_OP       = IDX.REG.OPERACION - 1;
    const COL_NOMBRES  = IDX.REG.NOMBRES - 1;
    const COL_CARGO    = IDX.REG.CARGO - 1;
    const COL_FECHA    = IDX.REG.FECHA - 1;

    // PASO 1: Por cada dni|producto, quedarse solo con la entrega MÁS RECIENTE
    // (por fecha de entrega, no por fecha de vencimiento).
    // Clave sin variante: una entrega nueva de cualquier variante del mismo
    // producto cancela la alerta de variantes anteriores.
    let cargoDelDni = '';
    const ultimasPorProducto = {}; // clave → { fila, fechaEntrega }

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      if (!esAdmin && _str(fila[COL_DNI]) !== _str(dniLogin)) continue;
      if (_str(fila[COL_OP]) !== 'Entrega') continue;

      const dni      = _str(fila[COL_DNI]);
      const producto = _str(fila[COL_PRODUCTO]);
      if (!producto) continue;

      if (!cargoDelDni && _str(fila[COL_DNI]) === _str(dniLogin)) {
        cargoDelDni = _str(fila[COL_CARGO]).toUpperCase().trim();
      }

      const clave = dni + '|' + producto;
      const fechaEntrega = new Date(fila[COL_FECHA]);
      if (isNaN(fechaEntrega.getTime())) continue;

      if (!ultimasPorProducto[clave] || fechaEntrega > ultimasPorProducto[clave].fechaEntrega) {
        ultimasPorProducto[clave] = { fila, fechaEntrega };
      }
    }

    // PASO 2: Construir set de productos en la MATRIZ del cargo
    let productosEnMatriz = null;
    if (!esAdmin && cargoDelDni) {
      const grid = _readMatrizGrid_();
      productosEnMatriz = new Set();
      for (const pb in grid.byProduct) {
        const cargos = (grid.byProduct[pb].previstoCargos || [])
          .map(c => c.trim().toUpperCase());
        if (cargos.some(c => c === cargoDelDni || cargoDelDni.includes(c) || c.includes(cargoDelDni))) {
          productosEnMatriz.add(pb.trim().toUpperCase());
        }
      }
      Logger.log('AlertasVenc — cargo: ' + cargoDelDni + ' | productos en matriz: ' + productosEnMatriz.size);
    }

    // PASO 3: Generar alertas solo de la entrega más reciente de cada producto
    const alertas = [];

    for (const clave in ultimasPorProducto) {
      const { fila } = ultimasPorProducto[clave];
      const producto = _str(fila[COL_PRODUCTO]);

      // Filtro MATRIZ: solo EPPs del cargo actual
      if (productosEnMatriz && !productosEnMatriz.has(producto.trim().toUpperCase())) continue;

      const fechaVencRaw = fila[COL_VENC];
      if (!fechaVencRaw) continue; // Sin fecha de vencimiento → cubierto indefinidamente

      const fechaVenc = new Date(fechaVencRaw);
      if (isNaN(fechaVenc.getTime())) continue;

      const diffDias = Math.ceil((fechaVenc - hoy) / (1000 * 60 * 60 * 24));
      const fechaFormateada = Utilities.formatDate(fechaVenc, 'GMT-5', 'dd/MM/yyyy');

      if (diffDias <= 0) {
        alertas.push({
          producto,
          trabajador: _str(fila[COL_NOMBRES]),
          fecha:  fechaFormateada,
          estado: 'VENCIDO',
          clase:  'fila-vencida',
          badge:  'bg-rojo'
        });
      } else if (diffDias <= 15) {
        alertas.push({
          producto,
          trabajador: _str(fila[COL_NOMBRES]),
          fecha:  fechaFormateada,
          estado: `VENCE EN ${diffDias} DÍAS`,
          clase:  'fila-proxima',
          badge:  'bg-naranja'
        });
      }
    }

    return alertas.sort((a, b) => (a.badge === 'bg-rojo' ? -1 : 1)).slice(0, 15);

  } catch (e) {
    Logger.log('Error en alertas: ' + e.message);
    return [];
  }
}

/**
 * Verifica si un trabajador tiene EPPs vencidos o entregas pendientes de firma.
 */
function verificarAlertasCompletas(dniLogin) {
  try {
    const alertasVenc = obtenerAlertasVencimientos(dniLogin);
    const pendientes  = (typeof obtenerEntregasPendientes === 'function')
      ? obtenerEntregasPendientes(dniLogin)
      : [];
    return {
      tieneVencimientos: alertasVenc && alertasVenc.length > 0,
      tienePendientes:   pendientes  && pendientes.length  > 0,
      totalPendientes:   pendientes  ? pendientes.length   : 0
    };
  } catch (e) {
    Logger.log('Error en verificarAlertasCompletas: ' + e.message);
    return { tieneVencimientos: false, tienePendientes: false, totalPendientes: 0 };
  }
}
