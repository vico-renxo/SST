/**
 * Obtiene alertas de vencimiento para el trabajador logueado.
 * - Agrupa por DNI + Producto (sin variante) → evita falsos VENCIDO cuando hay
 *   una entrega nueva con distinta variante que cancela la anterior.
 * - Filtra por MATRIZ: solo muestra productos asignados al cargo del usuario.
 * - Solo muestra la ÚLTIMA entrega por producto.
 */
function obtenerAlertasVencimientos(dniLogin) {
  try {
    const ss = getSpreadsheetEPP();
    const shReg = ss.getSheetByName(SHEPP.REGISTRO);
    const correoActual = Session.getActiveUser().getEmail();
    const ADMIN_EMAIL = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
                        || Session.getActiveUser().getEmail();
    const esAdmin = (correoActual === ADMIN_EMAIL || !dniLogin);

    const data = shReg.getDataRange().getValues();
    const hoy  = new Date();

    const COL_DNI      = IDX.REG.DNI - 1;
    const COL_PRODUCTO = IDX.REG.PRODUCTO - 1;
    const COL_VENC     = IDX.REG.FECHA_VENCIMIENTO - 1;
    const COL_OP       = IDX.REG.OPERACION - 1;
    const COL_NOMBRES  = IDX.REG.NOMBRES - 1;
    const COL_FECHA    = IDX.REG.FECHA - 1;
    const COL_CARGO    = IDX.REG.CARGO - 1;

    // PASO 1: Agrupar por DNI + Producto (sin variante).
    // Si existe una entrega nueva (ej. variante "UND") y otra vieja (variante "")
    // para el mismo producto, la nueva cancela la anterior.
    let cargoDelDni = '';
    const ultimasPorProducto = {};

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      if (!esAdmin && _str(fila[COL_DNI]) !== _str(dniLogin)) continue;
      if (_str(fila[COL_OP]) !== 'Entrega') continue;

      const dni      = _str(fila[COL_DNI]);
      const producto = _str(fila[COL_PRODUCTO]);
      const clave    = dni + '|' + producto; // sin variante

      if (!cargoDelDni && _str(fila[COL_DNI]) === _str(dniLogin)) {
        cargoDelDni = _str(fila[COL_CARGO]).toUpperCase().trim();
      }

      const fechaEntrega = new Date(fila[COL_FECHA]);
      if (isNaN(fechaEntrega.getTime())) continue;

      if (!ultimasPorProducto[clave] || fechaEntrega > ultimasPorProducto[clave].fechaEntrega) {
        ultimasPorProducto[clave] = { fila, fechaEntrega };
      }
    }

    // PASO 2: Construir set de productos asignados al cargo (MATRIZ).
    // Admin ve todo; trabajador solo ve sus EPPs según cargo.
    let productosEnMatriz = null;
    if (!esAdmin && cargoDelDni) {
      const grid = _readMatrizGrid_();
      productosEnMatriz = new Set();
      for (const pb in grid.byProduct) {
        const cargos = (grid.byProduct[pb].previstoCargos || [])
          .map(c => c.trim().toUpperCase());
        const match = cargos.some(c => c === cargoDelDni
          || cargoDelDni.includes(c)
          || c.includes(cargoDelDni));
        if (match) productosEnMatriz.add(pb.trim().toUpperCase());
      }
      Logger.log('AlertasVenc — cargo: ' + cargoDelDni + ' | productos en matriz: ' + productosEnMatriz.size);
    }

    // PASO 3: Generar alertas solo de entregas vencidas/próximas y en MATRIZ.
    const alertas = [];

    for (const clave in ultimasPorProducto) {
      const { fila } = ultimasPorProducto[clave];
      const producto = _str(fila[COL_PRODUCTO]);

      // Filtro MATRIZ
      if (productosEnMatriz && !productosEnMatriz.has(producto.trim().toUpperCase())) continue;

      const fechaVencRaw = fila[COL_VENC];
      if (!fechaVencRaw) continue;
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
