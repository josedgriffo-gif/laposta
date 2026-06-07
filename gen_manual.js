const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
  TableOfContents, PageBreak, Header, Footer, PageNumber
} = require('docx');

const GREEN = "1B5E20";
const GREEN2 = "2E7D32";

// Helpers
const P = (text, opts = {}) => new Paragraph({ children: [new TextRun({ text, ...opts })], spacing: { after: 120 }, ...(opts.p || {}) });
const H1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
const bullet = (text) => new Paragraph({ numbering: { reference: "b", level: 0 }, children: [new TextRun(text)], spacing: { after: 60 } });
const num = (text) => new Paragraph({ numbering: { reference: "n", level: 0 }, children: [new TextRun(text)], spacing: { after: 60 } });
const space = () => new Paragraph({ children: [new TextRun("")], spacing: { after: 80 } });

function nota(text, color, fill) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: border, bottom: border, left: { style: BorderStyle.SINGLE, size: 18, color }, right: border },
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun(text)] })]
    })] })]
  });
}
const importante = (t) => nota(t, "EF5350", "FFEBEE");
const tip = (t) => nota(t, "66BB6A", "E8F5E9");
const aviso = (t) => nota(t, "FFB300", "FFF8E1");

// Tabla simple
function tabla(headers, rows) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const colW = Math.floor(9360 / headers.length);
  const widths = headers.map(() => colW);
  const headRow = new TableRow({ children: headers.map(h => new TableCell({
    borders, width: { size: colW, type: WidthType.DXA },
    shading: { fill: GREEN, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF" })] })]
  })) });
  const dataRows = rows.map(r => new TableRow({ children: r.map(c => new TableCell({
    borders, width: { size: colW, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun(c)] })]
  })) }));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: widths, rows: [headRow, ...dataRows] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: GREEN },
        paragraph: { spacing: { before: 320, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: GREEN2 },
        paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, font: "Arial", color: "555555" },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "n", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun("La Posta — Manual de usuario · Página "), new TextRun({ children: [PageNumber.CURRENT] })
    ] })] }) },
    children: [
      // ── PORTADA ──
      new Paragraph({ spacing: { before: 2600, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "🛒", size: 96 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 0 }, children: [new TextRun({ text: "LA POSTA", bold: true, size: 72, color: GREEN })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 }, children: [new TextRun({ text: "Sistema de Punto de Venta", size: 34, color: GREEN2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 0 }, children: [new TextRun({ text: "Manual de usuario", bold: true, size: 30 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 0 }, children: [new TextRun({ text: "Versión 3.4", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1800, after: 0 }, children: [new TextRun({ text: "App: josedgriffo-gif.github.io/laposta", size: 20, color: "888888" })] }),
      new Paragraph({ children: [new PageBreak()] }),

      // ── ÍNDICE ──
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Índice")] }),
      new TableOfContents("Tabla de contenido", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),

      // ══════════ PARTE 1: ADMINISTRADOR ══════════
      H1("Parte 1 — Manual del Administrador"),
      P("Este manual cubre todo el sistema. Está pensado para vos, que configurás y manejás La Posta. Para el uso de la caja por parte de un empleado, mirá la Parte 2 (Guía rápida para la Cajera)."),

      H2("1. Qué es y cómo funciona"),
      P("La Posta es una app web de punto de venta. No necesita servidores ni hosting pago. Funciona con tres piezas:"),
      bullet("La app (las pantallas que ves), publicada en internet."),
      bullet("Un Google Sheet que guarda todos los datos (es la base de datos)."),
      bullet("Un \"backend\" en Google Apps Script que conecta la app con el Sheet."),
      space(),
      P("Direcciones de la app:"),
      tabla(["Pantalla", "Dirección"], [
        ["Caja", "josedgriffo-gif.github.io/laposta/"],
        ["Nroshoy", ".../laposta/dueno.html"],
        ["Admin / Accionistas", ".../laposta/admin.html"],
      ]),
      tip("Consejo: instalá la app en la tablet (\"Agregar a pantalla de inicio\") y guardá las direcciones en favoritos para no perderlas."),

      H2("2. Niveles de acceso y claves (PIN)"),
      P("El sistema tiene 4 niveles. Cada uno entra con su PIN de 4 dígitos:"),
      tabla(["Nivel", "Quién lo usa", "Qué ve"], [
        ["Caja", "Cajera/empleado", "Solo facturar. No ve costos ni ganancias."],
        ["Nroshoy", "Encargado", "Resumen del día (ventas, meta, medios de pago)."],
        ["Accionistas", "Socios", "Informes + Estado de resultados (rentabilidad)."],
        ["Admin", "Vos", "Todo el sistema."],
      ]),
      space(),
      P("Los PIN se cambian en Admin → Configuración. Deben ser de 4 dígitos y distintos entre sí."),
      importante("Cambiá los PIN por defecto antes de usar el sistema en serio. El PIN Admin es la llave maestra: abre todo, incluida la caja."),

      H2("3. Primer inicio (configuración por única vez)"),
      P("Esto ya quedó configurado. Se documenta por si hay que rehacerlo o entenderlo."),
      num("El Google Sheet \"La Posta — Sistema\" tiene todas las hojas de datos."),
      num("El backend (Apps Script) está publicado como aplicación web."),
      num("La app está publicada en GitHub Pages."),
      space(),
      H3("Cómo actualizar el backend"),
      P("Cuando se cambia la lógica del servidor, hay que copiar el archivo backend.gs y pegarlo en Apps Script:"),
      num("Abrir backend.gs con el Bloc de notas → Ctrl+A → Ctrl+C."),
      num("En Apps Script: Ctrl+A → borrar → Ctrl+V → Ctrl+S."),
      num("Implementar → Administrar implementaciones → lápiz → Nueva versión → Implementar."),
      aviso("Regla simple: si solo cambió la app (las pantallas), alcanza con recargar (Ctrl+F5). Si cambió el backend, hay que copiar y pegar en Apps Script."),

      H2("4. Productos y ofertas"),
      P("En Admin → Productos administrás el catálogo."),
      bullet("Agregar: completá nombre, categoría, unidad y precio. El producto entra solo al stock en 0."),
      bullet("Unidad: escribí la que quieras (un, kg, caja, litros, bidón...). Las usadas aparecen como sugerencia."),
      bullet("Editar: con el botón \"Editar\" de cada fila."),
      bullet("Activar/Desactivar: un producto inactivo no aparece en la caja."),
      bullet("Ordenar y buscar: tocá los títulos de columna para ordenar; usá el buscador y los filtros."),
      space(),
      H3("Ofertas por cantidad"),
      P("Cargá Precio de oferta + Cantidad mínima (y fecha fin opcional). Cuando el cliente lleva esa cantidad o más, la caja aplica el precio de oferta automáticamente."),

      H2("5. Compras y stock"),
      P("En Admin → Compras registrás lo que le comprás a un proveedor. Al guardar:"),
      bullet("El stock del producto sube automáticamente."),
      bullet("El costo del producto se actualiza con el de esta compra (para que el margen sea real)."),
      space(),
      H3("Ajuste de stock"),
      P("En Admin → Ajuste stock corregís el inventario:"),
      bullet("Entrada / Corrección: suma stock."),
      bullet("Salida / Merma: resta stock (se guarda en negativo solo)."),
      tip("El stock se ve en la columna Stock de la pestaña Productos: verde si es positivo, naranja si es 0, rojo si es negativo. El stock negativo está permitido, solo avisa."),

      H2("6. Cuenta corriente con proveedores"),
      P("Cuando cargás una compra con estado Pendiente o Parcial, aparece en Admin → Cuenta corriente."),
      bullet("Muestra, por proveedor, cuánto le debés y el detalle de cada compra impaga."),
      bullet("Arriba, el total general adeudado."),
      bullet("El botón \"Pagar\" en cada compra la marca como pagada y la saca de la cuenta corriente."),

      H2("7. Gastos"),
      P("En Admin → Gastos registrás lo que no es mercadería: sueldos, luz, agua, gas, alquiler, impuestos, etc."),
      bullet("La categoría es libre con sugerencias (escribí nuevas cuando haga falta)."),
      bullet("Estado Pagado o Pendiente."),
      bullet("Estos gastos se descuentan en Informes y Accionistas para calcular la ganancia real."),

      H2("8. Anular una venta"),
      P("En Admin → Anular venta:"),
      num("Elegí la fecha y tocá \"Ver ventas del día\"."),
      num("Encontrá la venta y tocá \"Anular\" → confirmá dos veces."),
      space(),
      P("Al anular: la venta queda CANCELADA (no se borra, queda el registro), el stock vuelve al inventario y el día se recalcula."),

      H2("9. Cierre de caja"),
      P("En Admin → Cierre de caja:"),
      num("Apertura: al empezar el día, cargá el fondo de caja inicial (con cuánto efectivo arrancás)."),
      num("Durante el día ves el resumen: total, margen, tickets y desglose por medio de pago."),
      num("Al cerrar, ingresás el efectivo real contado y el sistema te dice si la caja cuadra. El esperado es: fondo inicial + ventas en efectivo."),

      H2("10. Informes y Estado de resultados (Accionistas)"),
      H3("Informes"),
      P("Elegís un período (hoy, semana, mes, etc.) y ves: ventas, costo, margen, gastos, resultado neto, ranking de productos, ventas por categoría y gastos por categoría."),
      H3("Accionistas (Estado de resultados)"),
      P("Comparativo período contra período. Elegís semana a semana, mes a mes o año a año, y muestra en columnas:"),
      P("Ventas → menos Costo MP → Contribución marginal → menos Gastos → Utilidad neta.", { bold: false }),
      tip("Estas dos pestañas las ven solo Admin y Accionistas, nunca la cajera."),

      H2("11. Exportar a Excel"),
      P("En Admin → Exportar bajás los datos a un archivo Excel en tu PC (no toca el sistema)."),
      num("Marcá qué querés: productos, ventas, compras, gastos, stock, ajustes."),
      num("Elegí el período."),
      num("Tocá \"Descargar Excel\": un archivo con una hoja por cada cosa."),
      P("Sirve para armar tablas dinámicas y análisis a tu gusto en tu propia computadora, sin riesgo de romper el sistema."),

      H2("12. La tablet: pantalla completa y prestarla"),
      H3("Pantalla completa"),
      num("Abrí la app en Chrome → menú (3 puntitos) → \"Agregar a pantalla de inicio\"."),
      num("Abrí la app desde el ícono \"LP\". Se ve sin la barra del navegador."),
      H3("Prestar la tablet sin exponer tus datos"),
      bullet("Creá un usuario aparte en Android (Ajustes → Usuarios) solo para el negocio."),
      bullet("En ese usuario, no agregues tus cuentas de Google."),
      bullet("Usá \"Fijar pantalla\" de Android para bloquear la tablet en la app."),
      bullet("Tu usuario personal queda protegido con tu PIN o huella."),

      H2("13. Reiniciar datos de prueba"),
      P("En Admin → Configuración → \"Reiniciar datos de prueba\". Borra todas las ventas, compras, gastos y ajustes, y pone el stock en cero. Conserva los productos y la configuración."),
      importante("Usalo solo cuando termines de probar, para arrancar de cero en producción. Pide confirmación y PIN."),

      // ══════════ PARTE 2: CAJERA ══════════
      new Paragraph({ children: [new PageBreak()] }),
      H1("Parte 2 — Guía rápida para la Cajera"),
      P("Esta parte es solo para facturar. Imprimila y dejala cerca de la caja."),

      H2("Abrir la caja"),
      num("Abrí la app (el ícono \"LP\" en la tablet)."),
      num("Ingresá el PIN de Caja que te dieron."),

      H2("Hacer una venta"),
      num("Tocá una categoría arriba, o usá el buscador para encontrar el producto."),
      num("Tocá el producto:"),
      bullet("Si es por unidad: se agrega de a uno. Ajustá con + / − en el carrito de la derecha."),
      bullet("Si es por peso (kg): se abre un teclado. Escribí los GRAMOS (500 = medio kilo, 1100 = 1 kilo cien). No hace falta el punto. También hay botones rápidos (250, 500, 750, 1 kg)."),
      num("Repetí para cada producto. El total se ve abajo a la derecha."),
      num("Cuando está todo, tocá COBRAR."),
      num("Elegí cómo paga:"),
      bullet("Efectivo: escribí cuánto te dio el cliente y la pantalla calcula el vuelto."),
      bullet("Mercado Pago, Transferencia, Cuenta DNI, Tarjeta: tocá el botón que corresponde."),
      bullet("Mixto: si paga con varios medios, repartí los montos hasta cubrir el total."),
      num("Tocá CONFIRMAR. Aparece el número de ticket. ¡Listo!"),

      H2("Cosas para saber"),
      tip("Si se corta internet, podés seguir vendiendo igual. Las ventas se guardan y se suben solas cuando vuelve el wifi."),
      aviso("Si te equivocaste y querés vaciar el carrito antes de cobrar, tocá \"Cancelar venta\"."),
      importante("Si ya cobraste una venta por error, NO la podés anular vos. Avisá al encargado: la anulación se hace desde el Panel Admin con clave."),
      space(),
      P("Las ofertas se aplican solas cuando el cliente lleva la cantidad necesaria. No tenés que hacer nada especial."),

      new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "— Fin del manual —", color: "888888" })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Manual_La_Posta_v3.4.docx", buffer);
  console.log("Manual generado OK");
});
