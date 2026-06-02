/**
 * LA POSTA — Backend Google Apps Script
 * backend.gs v1.0
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
      case 'getConfig':       return ok(getConfig());
      case 'getVentasHoy':    return ok(getVentasHoy());
      case 'getResumenHoy':   return ok(getResumenHoy());
      case 'getCompras':      return ok(getCompras());
      case 'getAjustes':      return ok(getAjustes());
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
      case 'cerrarDia':          return ok(cerrarDia(body.data));
      case 'recalcularResumen':  return ok(recalcularResumen(body.data));
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
  return {
    PIN_Admin:    String(data[1][0]),
    PIN_Dueño:    String(data[1][1]).padStart(4, '0'),
    Meta_Diaria:  data[1][2]
  };
}

function getVentasHoy() {
  const hoy = fechaHoy();
  const ventas  = sheetToObjects(getSheet('BD_Ventas')).filter(v => v['Fecha'] === hoy);
  const detalle = sheetToObjects(getSheet('Detalle_Ventas')).filter(d => d['Fecha'] === hoy);
  return { ventas, detalle };
}

function getResumenHoy() {
  const hoy = fechaHoy();
  const resumen = sheetToObjects(getSheet('Resumen_Dia'));
  return resumen.find(r => r['Fecha'] === hoy) || null;
}

function getCompras() {
  return sheetToObjects(getSheet('Compras'));
}

function getAjustes() {
  return sheetToObjects(getSheet('Ajustes_Stock'));
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
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── BD_Ventas ──
  const sheetVentas = ss.getSheetByName('BD_Ventas');
  const lastRow = sheetVentas.getLastRow();
  const id = lastRow; // ID autoincremental (fila - 1 encabezado)

  const mp = data.mediosPago || {};
  sheetVentas.appendRow([
    id,
    data.ticket,
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
    data.nota || ''
  ]);

  // ── Detalle_Ventas ──
  const sheetDetalle = ss.getSheetByName('Detalle_Ventas');
  data.items.forEach(item => {
    sheetDetalle.appendRow([
      data.ticket,
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

  return { ticket: data.ticket, id };
}

/**
 * data = {
 *   fecha, proveedor, producto, categoria, cantidad, unidad,
 *   costoTotal, costoUnit, medioPago, estadoPago, observacion
 * }
 */
function guardarCompra(data) {
  getSheet('Compras').appendRow([
    data.fecha,
    data.proveedor,
    data.producto,
    data.categoria || '',
    data.cantidad,
    data.unidad,
    data.costoTotal,
    data.costoUnit,
    data.medioPago,
    data.estadoPago || 'Pagado',
    data.observacion || ''
  ]);
  return { ok: true };
}

/**
 * data = { fecha, producto, tipo, cantidad, motivo, responsable }
 */
function guardarAjuste(data) {
  getSheet('Ajustes_Stock').appendRow([
    data.fecha,
    data.producto,
    data.tipo,
    data.cantidad,
    data.motivo || '',
    data.responsable || ''
  ]);
  return { ok: true };
}

/**
 * data = { nombre, categoria, unidad, precio, costo, activo,
 *           precioOferta, cantMinOferta, fechaFinOferta }
 */
function guardarProducto(data) {
  getSheet('Productos').appendRow([
    data.nombre,
    data.categoria,
    data.unidad,
    data.precio,
    data.costo || 0,
    data.activo !== false ? 'SI' : 'NO',
    data.precioOferta || '',
    data.cantMinOferta || '',
    data.fechaFinOferta || ''
  ]);
  return { ok: true };
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
        'Fecha_Fin_Oferta': data.fechaFinOferta || ''
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

// Exponer generarTicket como acción GET
function doGet_extra(e) {
  if (e.parameter.action === 'generarTicket') {
    return ok({ ticket: generarTicket() });
  }
}
