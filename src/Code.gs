/**
 * Agrega el menu "Cotizador BDB" al abrir el spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cotizador BDB')
    .addItem('Inicializar hojas', 'initializeProject')
    .addToUi();
}

/**
 * Crea (o reinicia) la Hoja Maestra y la hoja Cotizador. Se corre una vez
 * al preparar el spreadsheet, o cuando se quiera regresar a la estructura
 * base.
 */
function initializeProject() {
  var hojaMaestra = setupHojaMaestra_();
  setupCotizador_(hojaMaestra);
  SpreadsheetApp.getUi().alert(
    'Listo. Revisa "Hoja Maestra" (costos) y "Cotizador" (captura de rutas).'
  );
}

/**
 * Al elegir el Tipo de Unidad en la hoja Cotizador, coloca automaticamente
 * la formula COTIZAR() en esa fila para que el costo y la tarifa piso
 * aparezcan solos, sin que Sihanka tenga que copiar formulas a mano.
 */
function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== COTIZADOR) return;
    if (range.getRow() < 2 || range.getColumn() !== 3) return;

    var formulaCell = sheet.getRange(range.getRow(), 9); // columna I
    if (range.getValue() === '') {
      formulaCell.clearContent();
    } else {
      formulaCell.setFormula(
        '=COTIZAR(C' + range.getRow() + ',D' + range.getRow() + ',F' + range.getRow() +
        ',G' + range.getRow() + ',H' + range.getRow() + ')'
      );
    }
  } catch (err) {
    // No interrumpir la edicion del usuario si algo falla en el trigger.
  }
}
