/**
 * LA POSTA — setup_sheets.gs v4
 *
 * setupSheets()        → crea hojas desde cero (primera vez)
 * agregarColumnasOferta2() → agrega Precio_Oferta_2 y Cantidad_Min_Oferta_2
 *                            a una hoja Productos existente, sin borrar datos
 */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEETS = [
    { name: 'BD_Ventas',       headers: ['ID','Ticket','Fecha','Hora','Total Venta','Costo Total','Margen $','Margen %','Medio Pago','Efectivo','Mercado Pago','Transferencia','Cuenta DNI','Tarjeta','Otro','Efectivo Recibido','Vuelto','Estado','Cliente/Nota'] },
    { name: 'Detalle_Ventas',  headers: ['Ticket','Fecha','Hora','Producto','Cantidad','Unidad','Precio Unit.','Subtotal','Costo Unit.','Costo Total','Margen $','Categoría'] },
    { name: 'Productos',       headers: ['Nombre','Categoría','Unidad','Precio','Costo','Activo','Precio_Oferta','Cantidad_Min_Oferta','Fecha_Fin_Oferta','Precio_Oferta_2','Cantidad_Min_Oferta_2'] },
    { name: 'Compras',         headers: ['Fecha','Proveedor','Producto','Categoría','Cantidad','Unidad','Costo_Total','Costo_Unit','Medio_Pago','Estado_Pago','Observación'] },
    { name: 'Ajustes_Stock',   headers: ['Fecha','Producto','Tipo','Cantidad','Motivo','Responsable'] },
    { name: 'Configuracion',   headers: ['PIN_Admin','PIN_Dueño','Meta_Diaria'] },
    { name: 'Resumen_Dia',     headers: ['Fecha','Total_Ventas','Costo_Total','Margen_$','Margen_%','Cant_Tickets','Efectivo','Mercado_Pago','Transferencia','Cuenta_DNI','Tarjeta','Otro','Meta_Diaria','Cumplimiento_%'] }
  ];

  for (const def of SHEETS) {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) sheet = ss.insertSheet(def.name);
    sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
  }

  ss.getSheetByName('Configuracion').getRange('A2:C2').setValues([['1234','0000',1500000]]);

  Logger.log('OK — hojas creadas');
}

function agregarColumnasOferta2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Productos');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  ['Precio_Oferta_2', 'Cantidad_Min_Oferta_2'].forEach(col => {
    if (!headers.includes(col)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      Logger.log('Columna agregada: ' + col);
    } else {
      Logger.log('Ya existe: ' + col);
    }
  });
}
