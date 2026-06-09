/**
 * LA POSTA — Backend Google Apps Script
 * backend.gs v4.9
 *
 * API REST para la app web de punto de venta.
 * Pegá este archivo completo en el editor de Apps Script
 * (reemplazando el contenido anterior).
 *
 * Después de pegar:
 * 1. Guardá (Ctrl+S)
 * 2. Implementar → Nueva implementación → Aplicación web
 * 3. Ejecutar como: Yo | Quién tiene acceso: Cualquiera
 * 4. Copiá la URL que te da — esa es la URL de la API
 */

// ── ID del Google Sheet ──────────────────────────────────────────────────────
const SHEET_ID = '1H4qiXfr5Wc6q45qQJlq14MMkEiWZQAvxTx6r2vYiNCs';

// ── Entrada GET ──────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      case 'getProductos':    return ok(getProductos());
      case 'getStock':        return ok(getStock());
      case 'getConfig':       return ok(getConfig());
      case 'getVentasHoy':    return ok(getVentasHoy());
      case 'getVentasFecha':  return ok(getVentasFecha(e.parameter.fecha));
      case 'getResumenHoy':   return ok(getResumenHoy());
      case 'getFondoCaja':    return ok(getFondoCaja(e.parameter.fecha));
      case 'getCompras':      return ok(getCompras());
      case 'getAjustes':      return ok(getAjustes());
      case 'getGastos':       return ok(getGastos());
      case 'getInforme':      return ok(getInforme(e.parameter.desde, e.parameter.hasta));
      case 'getDatosExport':  return ok(getDatosExport(e.parameter.desde, e.parameter.hasta));
      case 'getAccionistas':  return ok(getAccionistas(e.parameter.desde, e.parameter.hasta, e.parameter.modalidad));
      default:                return ok({ status: 'La Posta API activa' });
    }
  } catch (err) {
    return error(err.message);
  }
}

// ── Entrada POST ─────────────────────────────────────────────────────────────
function doPost(e) {
  // Lee desde FormData (campo 'payload') o desde postData según el cliente
  const raw = (e.parameter && e.parameter.payload) ? e.parameter.payload : e.postData.contents;
  const body = JSON.parse(raw);
  const action = body.action;
  try {
    switch (action) {
      case 'guardarVenta':    return ok(guardarVenta(body.data));
      case 'guardarCompra':   return ok(guardarCompra(body.data));
      case 'guardarAjuste':   return ok(guardarAjuste(body.data));
      case 'guardarProducto':      return ok(guardarProducto(body.data));
      case 'updateProducto':       return ok(updateProducto(body.data));
      case 'bulkImportProductos':  return ok(bulkImportProductos(body.data));
      case 'initStock':            return ok(initStock());
      case 'cerrarDia':          return ok(cerrarDia(body.data));
      case 'recalcularResumen':  return ok(recalcularResumen(body.data));
      case 'guardarConfig':      return ok(guardarConfig(body.data));
      case 'setFondoCaja':       return ok(setFondoCaja(body.data));
      case 'resetDatos':         return ok(resetDatos(body.data));
      case 'marcarCompraPagada': return ok(marcarCompraPagada(body.data));
      case 'registrarPago':      return ok(registrarPago(body.data));
      case 'guardarGasto':       return ok(guardarGasto(body.data));
      case 'marcarGastoPagado':  return ok(marcarGastoPagado(body.data));
      case 'anularVenta':        return ok(anularVenta(body.data));
      case 'borrarProducto':     return ok(borrarProducto(body.data));
      case 'borrarGasto':        return ok(borrarGasto(body.data));
      default:                return error('Acción desconocida: ' + action);
    }
  } catch (err) {
    return error(err.message);
  }
}

// ── Helpers de respuesta ─────────────────────────────────────────────────────
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// Convierte filas de Sheet a array de objetos usando la fila 1 como claves
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const val = row[i];
      // Convertir objetos Date a string YYYY-MM-DD para evitar problemas de comparación
      if (val instanceof Date && !isNaN(val)) {
        obj[h] = Utilities.formatDate(val, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
      } else {
        obj[h] = val;
      }
    });
    return obj;
  });
}

// ── Funciones GET ────────────────────────────────────────────────────────────

function getProductos() {
  return sheetToObjects(getSheet('Productos'));
}

function getConfig() {
  const sheet = getSheet('Configuracion');
  const data = sheet.getDataRange().getValues();
  // PIN_Socios (col B), PIN_Caja (col D); defaults si no existen
  const pinSocios = (data[1][1] !== undefined && data[1][1] !== '')
    ? String(data[1][1]).padStart(4, '0') : '5678';
  const pinCaja = (data[1][3] !== undefined && data[1][3] !== '')
    ? String(data[1][3]).padStart(4, '0') : '4321';
  return {
    PIN_Admin:   String(data[1][0]).padStart(4, '0'),
    PIN_Socios:  pinSocios,
    Meta_Diaria: data[1][2],
    PIN_Caja:    pinCaja
  };
}

/**
 * Guarda la configuración (PINs y meta diaria) en la hoja Configuracion.
 * data = { pinAdmin, pinSocios, pinCaja, metaDiaria }
 */
function guardarConfig(data) {
  const sheet = getSheet('Configuracion');
  // Asegura encabezados de las columnas nuevas
  sheet.getRange(1, 2).setValue('PIN_Socios');
  sheet.getRange(1, 4).setValue('PIN_Caja');
  sheet.getRange(2, 1, 1, 4).setValues([[
    String(data.pinAdmin).padStart(4, '0'),
    String(data.pinSocios || '5678').padStart(4, '0'),
    Number(data.metaDiaria) || 1500000,
    String(data.pinCaja || '4321').padStart(4, '0')
  ]]);
  return getConfig();
}

function getVentasHoy() {
  const hoy = fechaHoy();
  const ventas  = sheetToObjects(getSheet('BD_Ventas'))
    .filter(v => formatearFecha(v['Fecha']) === hoy && v['Estado'] !== 'CANCELADA');
  // Excluir del detalle los tickets de ventas anuladas
  const cancelados = ticketsCancelados();
  const detalle = sheetToObjects(getSheet('Detalle_Ventas'))
    .filter(d => formatearFecha(d['Fecha']) === hoy && !cancelados.has(String(d['Ticket'])));
  return { ventas, detalle };
}

// Conjunto de tickets de ventas anuladas (para excluir su detalle de los rankings)
function ticketsCancelados() {
  const set = {};
  sheetToObjects(getSheet('BD_Ventas')).forEach(v => {
    if (v['Estado'] === 'CANCELADA') set[String(v['Ticket'])] = true;
  });
  return { has: (t) => set[t] === true };
}

function getResumenHoy() {
  const hoy = fechaHoy();
  const resumen = sheetToObjects(getSheet('Resumen_Dia'));
  return resumen.find(r => r['Fecha'] === hoy) || null;
}

function getCompras() {
  const sheet = getSheet('Compras');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _fila: i + 2 }; // número de fila real en la hoja
    headers.forEach((h, j) => {
      const val = row[j];
      obj[h] = (val instanceof Date && !isNaN(val))
        ? Utilities.formatDate(val, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd')
        : val;
    });
    return obj;
  });
}

/**
 * Marca una compra como pagada (cambia Estado_Pago a 'Pagado').
 * data = { fila }
 */
// Asegura que la hoja Compras tenga la columna "Pagado" (monto abonado).
function asegurarColPagado(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf('Pagado');
  if (col < 0) {
    col = headers.length;            // próxima columna libre (0-based)
    sheet.getRange(1, col + 1).setValue('Pagado');
  }
  return col; // índice 0-based de la columna Pagado
}

// Marcar como pagada del todo = registrar un pago por el saldo completo.
function marcarCompraPagada(data) {
  return registrarPago({ fila: data.fila, todo: true });
}

/**
 * Registra un pago (total o parcial) a una compra.
 * data = { fila, monto } o { fila, todo:true }
 * Devuelve { pagado, saldo, estado }.
 */
function registrarPago(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Compras');
    const colPagado = asegurarColPagado(sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colTotal  = headers.indexOf('Costo_Total');
    const colEstado = headers.indexOf('Estado_Pago');

    const total  = Number(sheet.getRange(data.fila, colTotal + 1).getValue()) || 0;
    let pagado   = Number(sheet.getRange(data.fila, colPagado + 1).getValue()) || 0;
    const saldo  = total - pagado;

    let monto = data.todo ? saldo : (Number(data.monto) || 0);
    if (monto <= 0) throw new Error('El monto del pago debe ser mayor a 0.');
    if (monto > saldo) monto = saldo; // no se puede pagar más que el saldo

    const nuevoPagado = pagado + monto;
    sheet.getRange(data.fila, colPagado + 1).setValue(nuevoPagado);

    const estado = nuevoPagado >= total - 0.001 ? 'Pagado' : (nuevoPagado > 0 ? 'Parcial' : 'Pendiente');
    sheet.getRange(data.fila, colEstado + 1).setValue(estado);

    // Historial de pagos
    registrarHistorialPago(ss, sheet, data.fila, monto);

    return { ok: true, pagado: nuevoPagado, saldo: total - nuevoPagado, estado: estado };
  } finally {
    lock.releaseLock();
  }
}

// Guarda cada abono en la hoja Pagos_Proveedores (historial).
function registrarHistorialPago(ss, comprasSheet, fila, monto) {
  let hoja = ss.getSheetByName('Pagos_Proveedores');
  if (!hoja) {
    hoja = ss.insertSheet('Pagos_Proveedores');
    hoja.getRange(1, 1, 1, 5).setValues([['Fecha', 'Proveedor', 'Producto', 'Monto', 'Fila_Compra']]);
    hoja.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#d9d9d9');
    hoja.setFrozenRows(1);
  }
  const headers = comprasSheet.getRange(1, 1, 1, comprasSheet.getLastColumn()).getValues()[0];
  const prov = comprasSheet.getRange(fila, headers.indexOf('Proveedor') + 1).getValue();
  const prod = comprasSheet.getRange(fila, headers.indexOf('Producto') + 1).getValue();
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  hoja.appendRow([hoy, prov, prod, monto, fila]);
}

function getAjustes() {
  return sheetToObjects(getSheet('Ajustes_Stock'));
}

// ── Stock valorizado ────────────────────────────────────────────────────────
// Suma de (stock de cada producto × su costo actual). Capital en mercadería hoy.
function valorStock(ss) {
  const productos = sheetToObjects(ss.getSheetByName('Productos'));
  const costoDe = {};
  productos.forEach(p => { costoDe[String(p['Nombre']).trim()] = Number(p['Costo']) || 0; });

  const stockSheet = ss.getSheetByName('Stock_Actual');
  if (!stockSheet) return 0;

  let total = 0;
  sheetToObjects(stockSheet).forEach(r => {
    const s = Number(r['Stock_Actual']) || 0;
    const c = costoDe[String(r['Producto']).trim()] || 0;
    total += s * c;
  });
  return total;
}

// ── Informes ────────────────────────────────────────────────────────────────

/**
 * Calcula el informe completo de un rango de fechas (inclusive).
 * desde, hasta = 'yyyy-MM-dd'
 */
function getInforme(desde, hasta) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const enRango = (f) => {
    const fecha = formatearFecha(f);
    return fecha >= desde && fecha <= hasta;
  };

  // ── Ventas ──
  const ventas = sheetToObjects(ss.getSheetByName('BD_Ventas'))
    .filter(v => enRango(v['Fecha']) && v['Estado'] !== 'CANCELADA');

  const totalVentas = ventas.reduce((s, v) => s + (Number(v['Total Venta']) || 0), 0);
  const costoVentas = ventas.reduce((s, v) => s + (Number(v['Costo Total']) || 0), 0);
  const margenBruto = totalVentas - costoVentas;
  const totalRedondeo = ventas.reduce((s, v) => s + (Number(v['Redondeo']) || 0), 0);
  const cantTickets = ventas.length;
  const ticketProm  = cantTickets > 0 ? totalVentas / cantTickets : 0;

  const medios = {
    efectivo:      ventas.reduce((s, v) => s + (Number(v['Efectivo']) || 0), 0),
    mercadoPago:   ventas.reduce((s, v) => s + (Number(v['Mercado Pago']) || 0), 0),
    transferencia: ventas.reduce((s, v) => s + (Number(v['Transferencia']) || 0), 0),
    cuentaDni:     ventas.reduce((s, v) => s + (Number(v['Cuenta DNI']) || 0), 0),
    tarjeta:       ventas.reduce((s, v) => s + (Number(v['Tarjeta']) || 0), 0),
    otro:          ventas.reduce((s, v) => s + (Number(v['Otro']) || 0), 0)
  };

  // ── Detalle: ranking de productos y por categoría (excluye anuladas) ──
  const cancelados = ticketsCancelados();
  const detalle = sheetToObjects(ss.getSheetByName('Detalle_Ventas'))
    .filter(d => enRango(d['Fecha']) && !cancelados.has(String(d['Ticket'])));

  const prodMap = {};
  const catMap = {};
  detalle.forEach(d => {
    const nombre = d['Producto'] || '(sin nombre)';
    const cat = d['Categoría'] || '(sin categoría)';
    const subtotal = Number(d['Subtotal']) || 0;
    const margen = Number(d['Margen $']) || 0;
    const cantidad = Number(d['Cantidad']) || 0;

    if (!prodMap[nombre]) prodMap[nombre] = { nombre, categoria: cat, ingreso: 0, margen: 0, cantidad: 0 };
    prodMap[nombre].ingreso += subtotal;
    prodMap[nombre].margen += margen;
    prodMap[nombre].cantidad += cantidad;

    if (!catMap[cat]) catMap[cat] = { categoria: cat, ingreso: 0, margen: 0 };
    catMap[cat].ingreso += subtotal;
    catMap[cat].margen += margen;
  });

  const topProductos = Object.values(prodMap).sort((a, b) => b.ingreso - a.ingreso);
  const porCategoria = Object.values(catMap).sort((a, b) => b.ingreso - a.ingreso);

  // ── Gastos ──
  const gastos = getGastos().filter(g => enRango(g['Fecha']));
  const totalGastos = gastos.reduce((s, g) => s + (Number(g['Monto']) || 0), 0);
  const gastoCatMap = {};
  gastos.forEach(g => {
    const cat = g['Categoría'] || '(sin categoría)';
    if (!gastoCatMap[cat]) gastoCatMap[cat] = { categoria: cat, monto: 0 };
    gastoCatMap[cat].monto += Number(g['Monto']) || 0;
  });
  const gastosPorCategoria = Object.values(gastoCatMap).sort((a, b) => b.monto - a.monto);

  // ── Resultado neto = margen bruto - gastos ──
  const resultadoNeto = margenBruto - totalGastos;

  // ── Por día: agrupado por fecha ──
  const diaMap = {};
  ventas.forEach(v => {
    const fecha = formatearFecha(v['Fecha']);
    if (!diaMap[fecha]) diaMap[fecha] = { fecha, total: 0, efectivo: 0, mercadoPago: 0, transferencia: 0, cuentaDni: 0, tarjeta: 0, otro: 0, tickets: 0 };
    diaMap[fecha].total        += Number(v['Total Venta'])   || 0;
    diaMap[fecha].efectivo     += Number(v['Efectivo'])      || 0;
    diaMap[fecha].mercadoPago  += Number(v['Mercado Pago'])  || 0;
    diaMap[fecha].transferencia+= Number(v['Transferencia']) || 0;
    diaMap[fecha].cuentaDni    += Number(v['Cuenta DNI'])    || 0;
    diaMap[fecha].tarjeta      += Number(v['Tarjeta'])       || 0;
    diaMap[fecha].otro         += Number(v['Otro'])          || 0;
    diaMap[fecha].tickets++;
  });
  const porDia = Object.values(diaMap).sort((a, b) => a.fecha.localeCompare(b.fecha));

  return {
    desde, hasta,
    totalVentas, costoVentas, margenBruto, totalRedondeo,
    margenPct: totalVentas > 0 ? (margenBruto / totalVentas * 100) : 0,
    cantTickets, ticketProm,
    medios, porDia,
    topProductos, porCategoria,
    totalGastos, gastosPorCategoria,
    resultadoNeto,
    stockValorizado: valorStock(ss)
  };
}

/**
 * Estado de resultados comparativo por período.
 * Devuelve un array de períodos, cada uno con sus métricas.
 * modalidad = 'semanal' | 'mensual' | 'anual'
 */
function getAccionistas(desde, hasta, modalidad) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const enRango = (f) => {
    const fecha = formatearFecha(f);
    return fecha >= desde && fecha <= hasta;
  };

  // Clave y etiqueta del período según modalidad
  const claveDe = (fechaStr) => {
    const f = formatearFecha(fechaStr);
    if (modalidad === 'anual') return f.substring(0, 4);
    if (modalidad === 'semanal') {
      // Lunes de esa semana
      const d = new Date(f + 'T12:00:00');
      const day = d.getDay();              // 0 dom .. 6 sab
      const diff = (day === 0 ? -6 : 1 - day);
      d.setDate(d.getDate() + diff);
      return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
    }
    return f.substring(0, 7);              // mensual: YYYY-MM
  };

  const meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const labelDe = (clave) => {
    if (modalidad === 'anual') return clave;
    if (modalidad === 'semanal') {
      const p = clave.split('-');
      return `Sem ${p[2]}/${p[1]}`;
    }
    const p = clave.split('-');
    return `${meses[parseInt(p[1]) - 1]} ${p[0]}`;
  };

  // Acumular ventas y costo por período
  const acum = {};
  const get = (k) => { if (!acum[k]) acum[k] = { ventas:0, costoMP:0, gastos:0 }; return acum[k]; };

  sheetToObjects(ss.getSheetByName('BD_Ventas'))
    .filter(v => enRango(v['Fecha']) && v['Estado'] !== 'CANCELADA')
    .forEach(v => {
      const a = get(claveDe(v['Fecha']));
      a.ventas  += Number(v['Total Venta']) || 0;
      a.costoMP += Number(v['Costo Total']) || 0;
    });

  getGastos().filter(g => enRango(g['Fecha'])).forEach(g => {
    get(claveDe(g['Fecha'])).gastos += Number(g['Monto']) || 0;
  });

  // Armar array ordenado por clave de período
  const periodos = Object.keys(acum).sort().map(k => {
    const a = acum[k];
    const cm = a.ventas - a.costoMP;
    return {
      periodo: k,
      label: labelDe(k),
      ventas: a.ventas,
      costoMP: a.costoMP,
      contribMarginal: cm,
      gastos: a.gastos,
      utilidadNeta: cm - a.gastos
    };
  });

  return { periodos: periodos, stockValorizado: valorStock(ss) };
}

/**
 * Devuelve todos los datos del sistema para exportar a Excel.
 * Las tablas con fecha se filtran por rango; productos y stock van completos.
 */
function getDatosExport(desde, hasta) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const enRango = (f) => {
    const fecha = formatearFecha(f);
    return fecha >= desde && fecha <= hasta;
  };

  return {
    productos: sheetToObjects(ss.getSheetByName('Productos')),
    stock:     ss.getSheetByName('Stock_Actual') ? sheetToObjects(ss.getSheetByName('Stock_Actual')) : [],
    ventas:    sheetToObjects(ss.getSheetByName('BD_Ventas')).filter(v => enRango(v['Fecha'])),
    detalle:   sheetToObjects(ss.getSheetByName('Detalle_Ventas')).filter(d => enRango(d['Fecha'])),
    compras:   sheetToObjects(ss.getSheetByName('Compras')).filter(c => enRango(c['Fecha'])),
    gastos:    getGastos().filter(g => enRango(g['Fecha'])),
    ajustes:   sheetToObjects(ss.getSheetByName('Ajustes_Stock')).filter(a => enRango(a['Fecha']))
  };
}

// ── Gastos ──────────────────────────────────────────────────────────────────

// Devuelve la hoja Gastos, creándola con encabezados si no existe.
function getGastosSheet(ss) {
  let sheet = ss.getSheetByName('Gastos');
  if (!sheet) {
    sheet = ss.insertSheet('Gastos');
    sheet.getRange(1, 1, 1, 7).setValues([[
      'Fecha', 'Categoría', 'Detalle', 'Monto', 'Medio_Pago', 'Estado_Pago', 'Observación'
    ]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#d9d9d9');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getGastos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getGastosSheet(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _fila: i + 2 };
    headers.forEach((h, j) => {
      const val = row[j];
      obj[h] = (val instanceof Date && !isNaN(val))
        ? Utilities.formatDate(val, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd')
        : val;
    });
    return obj;
  });
}

/**
 * data = { fecha, categoria, detalle, monto, medioPago, estadoPago, observacion }
 */
function guardarGasto(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getGastosSheet(ss);
    sheet.appendRow([
      data.fecha,
      data.categoria,
      data.detalle,
      Number(data.monto) || 0,
      data.medioPago || '',
      data.estadoPago || 'Pagado',
      data.observacion || ''
    ]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function marcarGastoPagado(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getGastosSheet(ss);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colEstado = headers.indexOf('Estado_Pago');
    sheet.getRange(data.fila, colEstado + 1).setValue('Pagado');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── Funciones POST ───────────────────────────────────────────────────────────

/**
 * data = {
 *   ticket, fecha, hora, total, costoTotal, margen,
 *   medioPago, mediosPago: { efectivo, mercadoPago, transferencia, cuentaDni, tarjeta, otro },
 *   efectivoRecibido, vuelto, estado, nota,
 *   items: [{ producto, cantidad, unidad, precioUnit, subtotal, costoUnit, costoTotal, margen, categoria }]
 * }
 */
function guardarVenta(data) {
  // ── LOCK: evita que dos ventas simultáneas se pisen ──
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // espera hasta 30s por el turno

  try {
    // ── IDEMPOTENCIA: si esta venta (uuid) ya se procesó, devolver el mismo
    //    resultado sin volver a guardar. Evita duplicados por reintento. ──
    const props = PropertiesService.getScriptProperties();
    if (data.uuid) {
      const prev = props.getProperty('v_' + data.uuid);
      if (prev) return JSON.parse(prev);
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetVentas = ss.getSheetByName('BD_Ventas');

    // ── ID seguro: dentro del lock, lastRow es confiable ──
    const lastRow = sheetVentas.getLastRow();
    const id = lastRow; // fila - 1 encabezado

    // ── TICKET generado en el SERVIDOR (único global, ignora el del cliente) ──
    const ticket = generarTicketServidor(ss);

    const mp = data.mediosPago || {};
    sheetVentas.appendRow([
      id,
      ticket,
      data.fecha,
      data.hora,
      data.total,
      data.costoTotal || 0,
      data.margen || 0,
      data.total > 0 ? ((data.margen || 0) / data.total * 100).toFixed(2) : 0,
      data.medioPago,
      mp.efectivo || 0,
      mp.mercadoPago || 0,
      mp.transferencia || 0,
      mp.cuentaDni || 0,
      mp.tarjeta || 0,
      mp.otro || 0,
      data.efectivoRecibido || 0,
      data.vuelto || 0,
      data.estado || 'OK',
      data.nota || '',
      data.redondeo || 0
    ]);

    // ── Detalle_Ventas ──
    const sheetDetalle = ss.getSheetByName('Detalle_Ventas');
    data.items.forEach(item => {
      sheetDetalle.appendRow([
        ticket,
        data.fecha,
        data.hora,
        item.producto,
        item.cantidad,
        item.unidad,
        item.precioUnit,
        item.subtotal,
        item.costoUnit || 0,
        item.costoTotal || 0,
        item.margen || 0,
        item.categoria || ''
      ]);
    });

    // Actualizar resumen del día
    actualizarResumenDia(ss, data.fecha);

    // Descontar stock por cada ítem vendido
    data.items.forEach(item => {
      updateStock(ss, item.producto, -item.cantidad);
    });

    const resultado = { ticket: ticket, id: id };
    // Registrar uuid procesado para idempotencia
    if (data.uuid) props.setProperty('v_' + data.uuid, JSON.stringify(resultado));
    return resultado;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lista las ventas de una fecha (default hoy), con su estado.
 */
function getVentasFecha(fecha) {
  const f = fecha || fechaHoy();
  return sheetToObjects(getSheet('BD_Ventas')).filter(v => formatearFecha(v['Fecha']) === f);
}

/**
 * Anula una venta: la marca CANCELADA, devuelve el stock y recalcula el día.
 * data = { ticket }
 */
function anularVenta(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetV = ss.getSheetByName('BD_Ventas');
    const rows = sheetV.getDataRange().getValues();
    const headers = rows[0];
    const colTicket = headers.indexOf('Ticket');
    const colEstado = headers.indexOf('Estado');
    const colFecha  = headers.indexOf('Fecha');

    let fechaVenta = null;
    let encontrada = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][colTicket]) === String(data.ticket)) {
        if (String(rows[i][colEstado]) === 'CANCELADA') {
          throw new Error('Esta venta ya estaba anulada.');
        }
        sheetV.getRange(i + 1, colEstado + 1).setValue('CANCELADA');
        fechaVenta = formatearFecha(rows[i][colFecha]);
        encontrada = true;
        break;
      }
    }
    if (!encontrada) throw new Error('No se encontró la venta ' + data.ticket);

    // Devolver el stock de cada producto de esa venta
    const sheetD = ss.getSheetByName('Detalle_Ventas');
    const dRows = sheetD.getDataRange().getValues();
    const dh = dRows[0];
    const dcTicket = dh.indexOf('Ticket');
    const dcProd = dh.indexOf('Producto');
    const dcCant = dh.indexOf('Cantidad');
    for (let i = 1; i < dRows.length; i++) {
      if (String(dRows[i][dcTicket]) === String(data.ticket)) {
        updateStock(ss, dRows[i][dcProd], Number(dRows[i][dcCant]) || 0); // suma de vuelta
      }
    }

    // Recalcular el resumen del día
    if (fechaVenta) actualizarResumenDia(ss, fechaVenta);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Genera el próximo ticket único del día leyendo BD_Ventas.
 * Se llama SIEMPRE dentro de un lock para garantizar unicidad.
 */
function generarTicketServidor(ss) {
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyyMMdd');
  const sheet = ss.getSheetByName('BD_Ventas');
  const rows = sheet.getDataRange().getValues();
  let max = 0;
  rows.slice(1).forEach(row => {
    const t = String(row[1]); // columna Ticket
    if (t.startsWith(fecha)) {
      const num = parseInt(t.split('-')[1]) || 0;
      if (num > max) max = num;
    }
  });
  return `${fecha}-${String(max + 1).padStart(4, '0')}`;
}

/**
 * data = {
 *   fecha, proveedor, producto, categoria, cantidad, unidad,
 *   costoTotal, costoUnit, medioPago, estadoPago, observacion
 * }
 */
function guardarCompra(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const compras = ss.getSheetByName('Compras');
    const colPagado = asegurarColPagado(compras); // 0-based
    // Pagado inicial: si se carga como Pagado, queda saldado; sino 0 (debe todo)
    const estado = data.estadoPago || 'Pendiente';
    const pagadoInicial = (estado === 'Pagado') ? (Number(data.costoTotal) || 0) : 0;

    const fila = [
      data.fecha,
      data.proveedor,
      data.producto,
      data.categoria || '',
      data.cantidad,
      data.unidad,
      data.costoTotal,
      data.costoUnit,
      data.medioPago,
      estado,
      data.observacion || ''
    ];
    // Completar hasta la columna Pagado
    while (fila.length < colPagado) fila.push('');
    fila[colPagado] = pagadoInicial;
    compras.appendRow(fila);

    // Sumar stock por compra
    updateStock(ss, data.producto, data.cantidad);
    // Actualizar el costo del producto con el costo unitario de esta compra
    // (costo de reposición — mantiene el margen calculado correcto)
    if (data.costoUnit && data.costoUnit > 0) {
      actualizarCostoProducto(ss, data.producto, data.costoUnit);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualiza el campo Costo de un producto en la hoja Productos.
 */
function actualizarCostoProducto(ss, nombreProducto, nuevoCosto) {
  const sheet = ss.getSheetByName('Productos');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const colNombre = headers.indexOf('Nombre');
  const colCosto = headers.indexOf('Costo');
  if (colNombre < 0 || colCosto < 0) return;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][colNombre]).trim() === String(nombreProducto).trim()) {
      sheet.getRange(i + 1, colCosto + 1).setValue(nuevoCosto);
      return;
    }
  }
}

/**
 * data = { fecha, producto, tipo, cantidad, motivo, responsable }
 */
function guardarAjuste(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ss.getSheetByName('Ajustes_Stock').appendRow([
      data.fecha,
      data.producto,
      data.tipo,
      data.cantidad,
      data.motivo || '',
      data.responsable || ''
    ]);
    // cantidad ya viene con signo correcto (negativo para Salida/Merma)
    updateStock(ss, data.producto, data.cantidad);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * data = { nombre, categoria, unidad, precio, costo, activo,
 *           precioOferta, cantMinOferta, fechaFinOferta }
 */
function guardarProducto(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ss.getSheetByName('Productos').appendRow([
      data.nombre,
      data.categoria,
      data.unidad,
      data.precio,
      data.costo || 0,
      data.activo !== false ? 'SI' : 'NO',
      data.precioOferta || '',
      data.cantMinOferta || '',
      data.fechaFinOferta || '',
      data.precioOferta2 || '',
      data.cantMinOferta2 || ''
    ]);
    // Agregar al stock automáticamente (en 0) si no existe
    const stockSheet = ss.getSheetByName('Stock_Actual');
    if (stockSheet) {
      const existe = stockSheet.getDataRange().getValues()
        .slice(1).some(r => String(r[0]).trim() === String(data.nombre).trim());
      if (!existe) {
        const ahora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
        stockSheet.appendRow([data.nombre, 0, ahora]);
      }
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Borra un producto del catálogo y de Stock_Actual.
 * No toca las ventas históricas (guardan el nombre).
 * data = { nombre }
 */
function borrarProducto(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Productos');
    const rows = sheet.getDataRange().getValues();
    const colNombre = rows[0].indexOf('Nombre');
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][colNombre]).trim() === String(data.nombre).trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    // Borrar también de Stock_Actual
    const stock = ss.getSheetByName('Stock_Actual');
    if (stock) {
      const srows = stock.getDataRange().getValues();
      for (let i = srows.length - 1; i >= 1; i--) {
        if (String(srows[i][0]).trim() === String(data.nombre).trim()) {
          stock.deleteRow(i + 1);
        }
      }
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Borra un gasto por número de fila.
 * data = { fila }
 */
function borrarGasto(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getGastosSheet(ss);
    sheet.deleteRow(data.fila);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualiza un producto existente por nombre.
 * data = { nombreOriginal, ...campos a actualizar }
 */
function updateProducto(data) {
  const sheet = getSheet('Productos');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const colNombre = headers.indexOf('Nombre');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][colNombre] === data.nombreOriginal) {
      const map = {
        'Nombre': data.nombre,
        'Categoría': data.categoria,
        'Unidad': data.unidad,
        'Precio': data.precio,
        'Costo': data.costo,
        'Activo': data.activo !== false ? 'SI' : 'NO',
        'Precio_Oferta': data.precioOferta || '',
        'Cantidad_Min_Oferta': data.cantMinOferta || '',
        'Fecha_Fin_Oferta': data.fechaFinOferta || '',
        'Precio_Oferta_2': data.precioOferta2 || '',
        'Cantidad_Min_Oferta_2': data.cantMinOferta2 || ''
      };
      headers.forEach((h, col) => {
        if (map[h] !== undefined) {
          sheet.getRange(i + 1, col + 1).setValue(map[h]);
        }
      });
      return { ok: true };
    }
  }
  throw new Error('Producto no encontrado: ' + data.nombreOriginal);
}

/**
 * Recalcula y guarda el resumen del día en Resumen_Dia.
 */
function formatearFecha(val) {
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
  }
  return String(val || '').substring(0, 10);
}

function actualizarResumenDia(ss, fecha) {
  const ventas  = sheetToObjects(ss.getSheetByName('BD_Ventas')).filter(v => formatearFecha(v['Fecha']) === fecha && v['Estado'] !== 'CANCELADA');
  const config  = getConfig();
  const meta    = config.Meta_Diaria;

  const totalVentas = ventas.reduce((s, v) => s + (Number(v['Total Venta']) || 0), 0);
  const costoTotal  = ventas.reduce((s, v) => s + (Number(v['Costo Total']) || 0), 0);
  const margenAbs   = totalVentas - costoTotal;
  const margenPct   = totalVentas > 0 ? (margenAbs / totalVentas * 100).toFixed(2) : 0;
  const cantTickets = ventas.length;

  const efectivo      = ventas.reduce((s, v) => s + (Number(v['Efectivo']) || 0), 0);
  const mercadoPago   = ventas.reduce((s, v) => s + (Number(v['Mercado Pago']) || 0), 0);
  const transferencia = ventas.reduce((s, v) => s + (Number(v['Transferencia']) || 0), 0);
  const cuentaDni     = ventas.reduce((s, v) => s + (Number(v['Cuenta DNI']) || 0), 0);
  const tarjeta       = ventas.reduce((s, v) => s + (Number(v['Tarjeta']) || 0), 0);
  const otro          = ventas.reduce((s, v) => s + (Number(v['Otro']) || 0), 0);

  const cumplimiento = meta > 0 ? (totalVentas / meta * 100).toFixed(2) : 0;

  const sheetResumen = ss.getSheetByName('Resumen_Dia');
  const rows = sheetResumen.getDataRange().getValues();

  // Buscar si ya existe fila para esta fecha
  for (let i = 1; i < rows.length; i++) {
    if (formatearFecha(rows[i][0]) === fecha) {
      sheetResumen.getRange(i + 1, 1, 1, 14).setValues([[
        fecha, totalVentas, costoTotal, margenAbs, margenPct,
        cantTickets, efectivo, mercadoPago, transferencia,
        cuentaDni, tarjeta, otro, meta, cumplimiento
      ]]);
      return;
    }
  }

  // Si no existe, agregar fila nueva
  sheetResumen.appendRow([
    fecha, totalVentas, costoTotal, margenAbs, margenPct,
    cantTickets, efectivo, mercadoPago, transferencia,
    cuentaDni, tarjeta, otro, meta, cumplimiento
  ]);
}

/**
 * Cierre de día manual desde el panel admin.
 * data = { fecha, efectivoReal }
 */
function cerrarDia(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  actualizarResumenDia(ss, data.fecha);
  return getResumenHoy();
}

/**
 * Limpia filas duplicadas de Resumen_Dia y recalcula para una fecha dada.
 * data = { fecha }
 */
function recalcularResumen(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Resumen_Dia');
  const fecha = data.fecha || fechaHoy();

  // Eliminar todas las filas de esa fecha
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (formatearFecha(rows[i][0]) === fecha) {
      sheet.deleteRow(i + 1);
    }
  }

  // Recalcular desde cero
  actualizarResumenDia(ss, fecha);
  return getResumenHoy();
}

// ── Utilidades ───────────────────────────────────────────────────────────────

function fechaHoy() {
  return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');
}

/**
 * Genera el próximo número de ticket para el día.
 * Formato: YYYYMMDD-XXXX
 */
function generarTicket() {
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyyMMdd');
  const sheet = getSheet('BD_Ventas');
  const rows = sheet.getDataRange().getValues();

  let max = 0;
  rows.slice(1).forEach(row => {
    const ticket = String(row[1]);
    if (ticket.startsWith(fecha)) {
      const num = parseInt(ticket.split('-')[1]) || 0;
      if (num > max) max = num;
    }
  });

  return `${fecha}-${String(max + 1).padStart(4, '0')}`;
}

// ── Stock ─────────────────────────────────────────────────────────────────

function getStock() {
  return sheetToObjects(getSheet('Stock_Actual'));
}

/**
 * Actualiza el stock de un producto sumando delta (positivo o negativo).
 * Si el producto no existe en Stock_Actual, lo crea.
 */
function updateStock(ss, producto, delta) {
  const sheet = ss.getSheetByName('Stock_Actual');
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  const ahora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(producto).trim()) {
      const stockActual = Number(rows[i][1]) || 0;
      sheet.getRange(i + 1, 2).setValue(stockActual + delta);
      sheet.getRange(i + 1, 3).setValue(ahora);
      return;
    }
  }
  // No existe — crear fila nueva
  sheet.appendRow([producto, delta, ahora]);
}

/**
 * Inicializa la hoja Stock_Actual con todos los productos en stock 0
 * (solo los que no tengan fila todavía).
 */
function initStock() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Stock_Actual');

  if (!sheet) {
    sheet = ss.insertSheet('Stock_Actual');
    sheet.getRange(1, 1, 1, 3).setValues([['Producto', 'Stock_Actual', 'Ultima_Actualizacion']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#d9d9d9');
    sheet.setFrozenRows(1);
  }

  const productos = sheetToObjects(ss.getSheetByName('Productos'));
  const rows = sheet.getDataRange().getValues();
  const existentes = rows.slice(1).map(r => String(r[0]).trim());
  const ahora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');

  const nuevos = productos
    .filter(p => p['Nombre'] && !existentes.includes(String(p['Nombre']).trim()))
    .map(p => [p['Nombre'], 0, ahora]);

  if (nuevos.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, nuevos.length, 3).setValues(nuevos);
  }

  return { inicializados: nuevos.length, yaExistian: existentes.length };
}

/**
 * Importación masiva de productos desde lista de precios externa.
 * data = { productos: [...], limpiarAntes: true/false }
 */
function bulkImportProductos(data) {
  const sheet = getSheet('Productos');
  if (data.limpiarAntes) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  }
  const rows = data.productos.map(p => [
    p.nombre,
    p.categoria,
    p.unidad,
    p.precio,
    p.costo || 0,
    p.activo !== false ? 'SI' : 'NO',
    p.precioOferta || '',
    p.cantMinOferta || '',
    p.fechaFinOferta || ''
  ]);
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }
  return { importados: rows.length };
}

// ── Fondo de caja ───────────────────────────────────────────────────────────

/**
 * Guarda el fondo de caja inicial de un día.
 * data = { fecha, monto, responsable }
 */
function setFondoCaja(data) {
  const props = PropertiesService.getScriptProperties();
  const fecha = data.fecha || fechaHoy();
  props.setProperty('fondo_' + fecha, JSON.stringify({
    monto: Number(data.monto) || 0,
    responsable: data.responsable || '',
    hora: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'HH:mm')
  }));
  return { ok: true, fondo: Number(data.monto) || 0 };
}

function getFondoCaja(fecha) {
  const props = PropertiesService.getScriptProperties();
  const f = fecha || fechaHoy();
  const raw = props.getProperty('fondo_' + f);
  return raw ? JSON.parse(raw) : { monto: 0, responsable: '', hora: '' };
}

// ── Reset de datos (con PIN admin) ──────────────────────────────────────────

/**
 * Borra todos los datos transaccionales pero conserva Productos y Configuracion.
 * Requiere PIN admin correcto.
 * data = { pin }
 */
function resetDatos(data) {
  const config = getConfig();
  const pinAdmin = String(config.PIN_Admin).padStart(4, '0');
  if (String(data.pin).padStart(4, '0') !== pinAdmin) {
    throw new Error('PIN incorrecto. No se borró nada.');
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Hojas a limpiar (deja fila 1 de encabezados)
  ['BD_Ventas', 'Detalle_Ventas', 'Compras', 'Ajustes_Stock', 'Resumen_Dia'].forEach(nombre => {
    const sheet = ss.getSheetByName(nombre);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    }
  });

  // Stock_Actual a 0 (no borra productos, solo pone stock en cero)
  const stock = ss.getSheetByName('Stock_Actual');
  if (stock) {
    const lastRow = stock.getLastRow();
    if (lastRow > 1) {
      const ahora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
      for (let i = 2; i <= lastRow; i++) {
        stock.getRange(i, 2).setValue(0);
        stock.getRange(i, 3).setValue(ahora);
      }
    }
  }

  // Limpiar propiedades (fondos de caja y uuids de idempotencia)
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('fondo_') === 0 || k.indexOf('v_') === 0) props.deleteProperty(k);
  });

  return { ok: true, mensaje: 'Datos de prueba borrados. Productos y configuración conservados.' };
}

// Exponer generarTicket como acción GET
function doGet_extra(e) {
  if (e.parameter.action === 'generarTicket') {
    return ok({ ticket: generarTicket() });
  }
}
