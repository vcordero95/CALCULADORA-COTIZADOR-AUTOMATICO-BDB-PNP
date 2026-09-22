/**
 * Sirve el dashboard "Cotizador BDB" como Web App independiente. Acceso
 * restringido al dominio de BDB (ver src/appsscript.json: webapp.access).
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Cotizador BDB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Catalogos para llenar los dropdowns del dashboard: Tipo de Unidad y
 * Puesto / Categoria de Sueldo, tal como estan capturados en la Hoja
 * Maestra (cada uno en su propia tabla, independiente entre si).
 */
function obtenerCatalogos() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_MAESTRA);
  if (!hoja) {
    throw new Error('No existe la hoja "Hoja Maestra". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  var soloLlenos = function (valor) { return valor !== ''; };

  var unidades = hoja.getRange(FILA_DATOS_MAESTRA, 1, FILAS_MAESTRA, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  var puestos = hoja.getRange(FILA_DATOS_MAESTRA, 10, FILAS_MAESTRA, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  return { unidades: unidades, puestos: puestos };
}

/**
 * Calcula una cotizacion desde el dashboard, reusando COTIZAR() (la misma
 * formula que usa la hoja Cotizador). El margen llega del formulario en
 * porcentaje (20) y aqui se convierte a fraccion (0.20).
 */
function calcularCotizacionWeb(datos) {
  var margenFraccion = Number(datos.margen) / 100;
  var fila = COTIZAR(datos.tipoUnidad, datos.puesto, datos.km, datos.viajesMes, datos.casetas, margenFraccion)[0];

  return {
    sueldoMensual: fila[0],
    costoGasolinaKm: fila[1],
    rentaMensual: fila[2],
    mantenimientoMensual: fila[3],
    costoVariable: fila[4],
    costoFijoProrrateado: fila[5],
    costoTotal: fila[6],
    tarifaPiso: fila[7]
  };
}

/**
 * Guarda la cotizacion calculada como un renglon nuevo en la hoja
 * Cotizador, con los valores ya resueltos (no formulas), para que quede
 * como un registro fijo de lo que se cotizo al momento de la solicitud.
 */
function guardarCotizacionWeb(datos, resultado) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(COTIZADOR);
  if (!sheet) {
    throw new Error('No existe la hoja "Cotizador". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  var fila = siguienteFilaLibreCotizador_(sheet);
  var margenFraccion = Number(datos.margen) / 100;

  sheet.getRange(fila, 1, 1, 17).setValues([[
    new Date(), datos.rutaCliente || '', datos.tipoUnidad, datos.puesto,
    Number(datos.km) || 0, Number(datos.paradas) || 0, Number(datos.viajesMes) || 0,
    Number(datos.casetas) || 0, margenFraccion,
    resultado.sueldoMensual, resultado.costoGasolinaKm, resultado.rentaMensual, resultado.mantenimientoMensual,
    resultado.costoVariable, resultado.costoFijoProrrateado, resultado.costoTotal, resultado.tarifaPiso
  ]]);

  return true;
}

/**
 * Primera fila libre de la hoja Cotizador, segun la columna Tipo de
 * Unidad (que nunca viene pre-llenada, a diferencia del margen por
 * defecto que si se pre-llena en las 500 filas reservadas).
 */
function siguienteFilaLibreCotizador_(sheet) {
  var columna = sheet.getRange(2, 3, FILAS_COTIZADOR, 1).getValues();
  for (var i = 0; i < columna.length; i++) {
    if (columna[i][0] === '') {
      return 2 + i;
    }
  }
  return 2 + FILAS_COTIZADOR;
}

/**
 * Atajo desde el menu de la hoja de calculo para abrir el dashboard
 * publicado, sin tener que buscar la URL de la implementacion.
 */
function abrirDashboard() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Aun no hay una version publicada. Implementar > Nueva implementacion > Aplicacion web.');
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<a href="' + url + '" target="_blank" style="font-family:sans-serif;">Abrir Cotizador BDB</a>' +
    '<script>window.open(' + JSON.stringify(url) + ', "_blank");</script>'
  ).setWidth(300).setHeight(60);
  ui.showModalDialog(html, 'Cotizador BDB');
}
