/**
 * Agrega el menu "Cotizador BDB" al abrir el spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cotizador BDB')
    .addItem('Inicializar hojas (borra y reconstruye todo)', 'initializeProject')
    .addItem('Cargar/actualizar Tarifas Vigentes y MELI', 'cargarTarifasVigentes')
    .addSeparator()
    .addItem('Abrir dashboard Comercial', 'abrirDashboardComercial')
    .addItem('Abrir dashboard interno (costos)', 'abrirDashboardInterno')
    .addToUi();
}

/**
 * Crea (o reinicia) solo "Tarifas Vigentes", "MELI Estaciones" y
 * "MELI Tarifas", sin tocar Costos Unidad, Nomina, Zonas, Ruta-Zona-Puesto,
 * Cotizador ni Solicitudes. Util para cargar/actualizar esas 3 hojas sin
 * arriesgar los costos reales que ya se hayan capturado con
 * "Inicializar hojas".
 */
function cargarTarifasVigentes() {
  setupTarifasVigentes_();
  setupMeliEstaciones_();
  setupMeliTarifas_();
  SpreadsheetApp.getUi().alert('Listo. Se cargaron/actualizaron "Tarifas Vigentes", "MELI Estaciones" y "MELI Tarifas".');
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
