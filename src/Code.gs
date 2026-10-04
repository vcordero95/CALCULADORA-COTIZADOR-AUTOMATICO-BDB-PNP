/**
 * Agrega el menu "Cotizador BDB" al abrir el spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cotizador BDB')
    .addItem('Inicializar hojas (borra y reconstruye todo)', 'initializeProject')
    .addItem('Cargar/actualizar Tarifas Vigentes y MELI', 'cargarTarifasVigentes')
    .addItem('Agregar columna INEGI (sin borrar datos)', 'migrarCodigoVehiculoInegi_')
    .addItem('Activar actualizacion semanal de precios de combustible (INEGI)', 'activarActualizacionCombustibleInegi_')
    .addItem('Quitar Volumen y agregar Periodo de Cantidad en Solicitudes', 'migrarPeriodoCantidad_')
    .addSeparator()
    .addItem('Abrir dashboard Comercial', 'abrirDashboardComercial')
    .addItem('Abrir dashboard interno (costos)', 'abrirDashboardInterno')
    .addToUi();
}

/**
 * Crea (o reinicia) solo "Tarifas Vigentes", "MELI Estaciones",
 * "MELI Tarifas" y "Amazon Tarifas", sin tocar Costos Unidad, Nomina,
 * Zonas, Ruta-Zona-Puesto, Cotizador ni Solicitudes. Util para cargar/
 * actualizar esas hojas sin arriesgar los costos reales que ya se hayan
 * capturado con "Inicializar hojas". Ya no hace falta usar esto para
 * agregar tarifas nuevas del dia a dia: las busquedas leen directo de
 * la hoja, asi que esas se pueden agregar ahi a mano sin pasar por el
 * codigo. Este boton sigue siendo destructivo (borra y reescribe estas
 * 4 hojas desde el catalogo semilla del codigo), asi que solo se usa
 * cuando Vanessa confirma que no tiene capturas manuales pendientes en
 * ellas que se perderian.
 */
function cargarTarifasVigentes() {
  setupTarifasVigentes_();
  setupMeliEstaciones_();
  setupMeliTarifas_();
  setupAmazonTarifas_();
  SpreadsheetApp.getUi().alert('Listo. Se cargaron/actualizaron "Tarifas Vigentes", "MELI Estaciones", "MELI Tarifas" y "Amazon Tarifas".');
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
