/**
 * Agrega el menu "Cotizador BDB" al abrir el spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cotizador BDB')
    .addItem('Inicializar hojas', 'initializeProject')
    .addItem('Abrir dashboard', 'abrirDashboard')
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
 * Al elegir el Tipo de Unidad o el Puesto / Categoria de Sueldo en la hoja
 * Cotizador, coloca automaticamente la formula COTIZAR() en esa fila para
 * que el costo y la tarifa piso aparezcan solos, sin que Sihanka tenga que
 * copiar formulas a mano. Como son dos catalogos independientes, la
 * formula solo se coloca cuando ambos ya estan elegidos.
 */
function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== COTIZADOR) return;
    var row = range.getRow();
    if (row < 2 || (range.getColumn() !== 3 && range.getColumn() !== 4)) return;

    var tipoUnidad = sheet.getRange(row, 3).getValue();
    var puesto = sheet.getRange(row, 4).getValue();
    var formulaCell = sheet.getRange(row, 10); // columna J

    if (tipoUnidad === '' || puesto === '') {
      formulaCell.clearContent();
    } else {
      formulaCell.setFormula(
        '=COTIZAR(C' + row + ',D' + row + ',E' + row + ',G' + row + ',H' + row + ',I' + row + ')'
      );
    }
  } catch (err) {
    // No interrumpir la edicion del usuario si algo falla en el trigger.
  }
}
