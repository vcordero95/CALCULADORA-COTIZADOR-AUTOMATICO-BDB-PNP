/**
 * Sirve el dashboard como Web App independiente. Acceso restringido al
 * dominio de BDB (ver src/appsscript.json: webapp.access). Por default
 * sirve el dashboard de Comercial (solicitud de tarifa de cliente);
 * agregando ?vista=interna a la URL se sirve el dashboard de costos de
 * uso interno.
 */
function doGet(e) {
  var vista = e && e.parameter && e.parameter.vista;
  var archivo = vista === 'interna' ? 'Dashboard' : 'DashboardComercial';
  return HtmlService.createHtmlOutputFromFile(archivo)
    .setTitle('Cotizador BDB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Catalogos para el dashboard interno (costos): Tipo de Unidad y Puesto /
 * Categoria de Sueldo, cada uno de su propio catalogo.
 */
function obtenerCatalogos() {
  var soloLlenos = function (valor) { return valor !== ''; };

  var unidades = obtenerHojaCostosUnidad_().getRange(2, 1, FILAS_CATALOGO, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  var puestos = obtenerHojaNomina_().getRange(2, 1, FILAS_CATALOGO, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  return { unidades: unidades, puestos: puestos };
}

/**
 * Catalogos para el dashboard de Comercial: Zona/Ciudad y Tipo de Unidad
 * salen de sus catalogos; Tipo de Ruta, Frecuencia y Tipo de Cobro son
 * categorias fijas del negocio.
 */
function obtenerCatalogosComercial() {
  var soloLlenos = function (valor) { return valor !== ''; };

  var zonas = SpreadsheetApp.getActive().getSheetByName(HOJA_ZONAS)
    .getRange(2, 1, FILAS_CATALOGO, 1).getValues()
    .map(function (fila) { return fila[0]; }).filter(soloLlenos);

  var unidades = obtenerHojaCostosUnidad_().getRange(2, 1, FILAS_CATALOGO, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  return {
    zonas: zonas,
    unidades: unidades,
    tiposRuta: TIPOS_RUTA,
    frecuencias: FRECUENCIAS,
    tiposCobro: TIPOS_COBRO
  };
}

/**
 * Calcula una cotizacion desde el dashboard interno, reusando COTIZAR()
 * (la misma formula que usa la hoja Cotizador). El margen llega del
 * formulario en porcentaje (20) y aqui se convierte a fraccion (0.20).
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
 * Guarda la cotizacion calculada (dashboard interno) como un renglon
 * nuevo en la hoja Cotizador, con los valores ya resueltos (no formulas).
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
    if (columna[i][0] === '') return 2 + i;
  }
  return 2 + FILAS_COTIZADOR;
}

/**
 * Calcula una solicitud de tarifa de cliente desde el dashboard de
 * Comercial, reusando SOLICITAR_TARIFA(). El margen llega del formulario
 * en porcentaje (20) y aqui se convierte a fraccion (0.20).
 */
function calcularSolicitudWeb(datos) {
  var margenFraccion = Number(datos.margen) / 100;
  return SOLICITAR_TARIFA(
    datos.tipoRuta, datos.zona, datos.tipoUnidad, datos.frecuencia, datos.km,
    datos.tipoCobro, datos.cantidad, !!datos.requiereAuxiliar, margenFraccion
  );
}

/**
 * Guarda la solicitud calculada (dashboard de Comercial) como un renglon
 * nuevo en la hoja Solicitudes, con los valores ya resueltos.
 */
function guardarSolicitudWeb(datos, resultado) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(HOJA_SOLICITUDES);
  if (!sheet) {
    throw new Error('No existe la hoja "Solicitudes". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  var fila = sheet.getLastRow() + 1;
  var margenFraccion = Number(datos.margen) / 100;

  sheet.getRange(fila, 1, 1, 25).setValues([[
    new Date(), datos.rutaCliente || '', datos.tipoRuta, datos.zona, datos.tipoUnidad,
    datos.frecuencia, datos.volumen || '', Number(datos.km) || 0, datos.tipoCobro, Number(datos.cantidad) || '',
    datos.requiereAuxiliar ? 'Si' : 'No',
    resultado.puestoPrincipal, resultado.puestoAuxiliar, resultado.costoCasetas, resultado.viajesMes,
    resultado.sueldoMensual, resultado.rentaMensual, resultado.mantenimientoMensual, resultado.costoGasolinaKm,
    resultado.costoVariable, resultado.costoFijoProrrateado, resultado.costoTotal, margenFraccion,
    resultado.tarifaPiso, resultado.tarifaPorUnidad
  ]]);

  return true;
}

/**
 * Atajos desde el menu de la hoja de calculo para abrir cada dashboard
 * publicado, sin tener que buscar la URL de la implementacion.
 */
function abrirDashboardComercial() {
  abrirUrl_(ScriptApp.getService().getUrl());
}

function abrirDashboardInterno() {
  var base = ScriptApp.getService().getUrl();
  abrirUrl_(base ? base + '?vista=interna' : null);
}

function abrirUrl_(url) {
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
