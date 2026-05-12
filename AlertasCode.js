/**
 * Obtiene alertas de vencimiento de EPP para el trabajador logueado.
 *
 * ALGORITMO:
 *   Por cada producto (sin distinguir variante), recopila TODAS las entregas.
 *   - Si ALGUNA tiene vencimiento futuro → el trabajador está cubierto, sin alerta.
 *   - Si TODAS están vencidas → alerta con la fecha más reciente (última expirada).
 *
 * Filtra por MATRIZ: solo productos asignados al cargo actual del trabajador.
 * Admin (ADMIN_EMAIL) omite el filtro de MATRIZ.
 */
function obtenerAlertasVencimientos(dniLogin) {
  try {
    const ss  = getSpreadsheetEPP();
    const shReg = ss.getSheetByName(SHEPP.REGISTRO);
    const correoActual = Session.getActiveUser().getEmail();
    const ADMIN_EMAIL  = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
                         || Session.getActiveUser().getEmail();
    const esAdmin = (correoActual === ADMIN_EMAIL || !dniLogin);

    const data = shReg.getDataRange().getValues();
    const hoy  = new Date();

    const COL_DNI      = IDX.REG.DNI - 1;
    const COL_PRODUCTO = IDX.REG.PRODUCTO - 1;
    const COL_VENC     = IDX.REG.FECHA_VENCIMIENTO - 1;
    const COL_OP       = IDX.REG.OPERACION - 1;
    const COL_NOMBRES  = IDX.REG.NOMBRES - 1;
    const COL_CARGO    = IDX.REG.CARGO - 1;

    // PASO 1: Recopilar TODAS las entregas por dni|producto (sin variante)
    let cargoDelDni = '';
    const entregasPorProducto = {}; // clave → [{ fila, fechaVenc }]

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      if (!esAdmin && _str(fila[COL_DNI]) !== _str(dniLogin)) continue;
      if (_str(fila[COL_OP]) !== 'Entrega') continue;

      const dni      = _str(fila[COL_DNI]);
      const producto = _str(fila[COL_PRODUCTO]);
      if (!producto) continue;

      const clave = dni + '|' + producto;

      if (!cargoDelDni && _str(fila[COL_DNI]) === _str(dniLogin)) {
        cargoDelDni = _str(fila[COL_CARGO]).toUpperCase().trim();
      }

      const fechaVencRaw = fila[COL_VENC];
      if (!fechaVencRaw) continue;
      const fechaVenc = new Date(fechaVencRaw);
      if (isNaN(fechaVenc.getTime())) continue;

      if (!entregasPorProducto[clave]) entregasPorProducto[clave] = [];
      entregasPorProducto[clave].push({ fila, fechaVenc });
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

    // PASO 3: Por cada producto, verificar si hay ALGUNA entrega vigente
    const alertas = [];

    for (const clave in entregasPorProducto) {
      const entregas = entregasPorProducto[clave];
      if (!entregas.length) continue;

      const producto = _str(entregas[0].fila[COL_PRODUCTO]);

      // Filtro MATRIZ: solo EPPs del cargo actual
      if (productosEnMatriz && !productosEnMatriz.has(producto.trim().toUpperCase())) continue;

      // Si ALGUNA entrega tiene vencimiento futuro → cubierto, sin alerta
      const tienePlazoVigente = entregas.some(e => e.fechaVenc > hoy);
      if (tienePlazoVigente) continue;

      // Todas vencidas → alerta con la de vencimiento más reciente
      entregas.sort((a, b) => b.fechaVenc - a.fechaVenc);
      const { fila, fechaVenc } = entregas[0];

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
