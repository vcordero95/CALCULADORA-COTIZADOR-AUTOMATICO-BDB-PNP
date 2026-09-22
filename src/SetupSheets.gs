/**
 * Nombres de las hojas usadas por el cotizador. Cada catalogo vive en su
 * propia pestaña (mas facil de mantener que varias tablas en una sola
 * hoja por columnas).
 */
var HOJA_COSTOS_UNIDAD = 'Costos Unidad';
var HOJA_NOMINA = 'Nomina';
var HOJA_PUNTOS = 'Puntos';
var HOJA_CASETAS = 'Casetas';
var HOJA_RUTA_ZONA_PUESTO = 'Ruta-Zona-Puesto';
var HOJA_CONFIG = 'Config';
var COTIZADOR = 'Cotizador';
var HOJA_SOLICITUDES = 'Solicitudes';

/** Filas reservadas para cada catalogo (Costos Unidad, Nomina, Puntos). */
var FILAS_CATALOGO = 50;

/** Filas reservadas para combinaciones Tipo de Ruta + Destino, y para Casetas (Punto A + Punto B). */
var FILAS_RUTA_ZONA = 100;
var FILAS_CASETAS = 150;

/** Filas reservadas para captura en Cotizador y Solicitudes. */
var FILAS_COTIZADOR = 500;
var FILAS_SOLICITUDES = 500;

/**
 * Valores fijos de Tipo de Ruta, Frecuencia y Tipo de Cobro. Son
 * categorias estables del negocio (no un catalogo que crezca), por eso
 * van fijas en el codigo en vez de en una hoja.
 */
var TIPOS_RUTA = ['Local', 'Foráneo'];
var FRECUENCIAS = ['7x7', '6x7', '5x7', '4x7', '3x7', '2x7', '1x7'];
var TIPOS_COBRO = ['Por Ruta', 'Por Paquete', 'Por Parada', 'Por Palet'];

/**
 * Crea (o reinicia) todas las hojas del cotizador. Se corre una vez al
 * preparar el spreadsheet, o cuando se quiera regresar a la estructura
 * base.
 */
function initializeProject() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.alert(
    'Inicializar hojas',
    'Esto borra y reconstruye TODAS las hojas del cotizador (Costos Unidad, Nomina, Puntos, Casetas, ' +
    'Ruta-Zona-Puesto, Config, Cotizador, Solicitudes, Tarifas Vigentes, MELI Estaciones, MELI Tarifas), ' +
    'regresandolas a sus datos de ejemplo. Si ya capturaste costos reales, se perderan. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) return;

  setupConfig_();
  var hojaCostosUnidad = setupCostosUnidad_();
  var hojaNomina = setupNomina_();
  var hojaPuntos = setupPuntos_();
  setupCasetas_(hojaPuntos);
  setupRutaZonaPuesto_(hojaPuntos, hojaNomina);
  setupCotizador_(hojaCostosUnidad, hojaNomina);
  setupSolicitudes_();
  setupTarifasVigentes_();
  setupMeliEstaciones_();
  setupMeliTarifas_();
  SpreadsheetApp.getUi().alert(
    'Listo. Revisa "Costos Unidad", "Nomina", "Puntos", "Casetas", "Ruta-Zona-Puesto" y "Config" ' +
    '(catalogos); "Cotizador" / "Solicitudes" (registro de cotizaciones); y ' +
    '"Tarifas Vigentes" / "MELI Estaciones" / "MELI Tarifas" (referencia de tarifas actuales).'
  );
}

/**
 * "Config": precios de combustible, parte fiscal por periodo, y la
 * politica de margen (Piso / Objetivo) que usa el dashboard de Comercial
 * en vez de pedirle el margen a Comercial. Editable sin tocar el script
 * cuando cambien.
 */
function setupConfig_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_CONFIG) || ss.insertSheet(HOJA_CONFIG);
  sheet.clear();

  var datos = [
    ['Precio Diesel ($/L)', 24.5],
    ['Precio Gasolina ($/L)', 23.8],
    ['Fiscal Semanal', 2253.6],
    ['Fiscal Quincenal', 5190.6],
    ['Margen Piso', 0.2],
    ['Margen Objetivo', 0.3]
  ];
  sheet.getRange(1, 1, datos.length, 2).setValues(datos);
  sheet.getRange(1, 1, datos.length, 1).setFontWeight('bold');
  sheet.getRange(1, 2, 4, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(5, 2, 2, 1).setNumberFormat('0%');
  sheet.autoResizeColumns(1, 2);
  return sheet;
}

/**
 * "Costos Unidad": catalogo por Tipo de Unidad (renta, mantenimiento,
 * rendimiento y tipo de combustible). El costo de gasolina por km se
 * resuelve solo contra los precios de Config.
 */
function setupCostosUnidad_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_COSTOS_UNIDAD) || ss.insertSheet(HOJA_COSTOS_UNIDAD);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = [
    'Tipo de Unidad', 'Renta Mensual', 'Costo Diario', 'Mantenimiento Mensual', 'Mantenimiento Diario',
    'Rendimiento (km/L)', 'Tipo de Combustible', 'Costo Gasolina por KM'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1c2b4a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var combustibleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Diesel', 'Gasolina'], true).setAllowInvalid(false).build();
  sheet.getRange(2, 7, FILAS_CATALOGO, 1).setDataValidation(combustibleRule);

  var costoDiario = [];
  var mantenimientoDiario = [];
  var costoGasolinaKm = [];
  for (var i = 0; i < FILAS_CATALOGO; i++) {
    var r = 2 + i;
    costoDiario.push(['=IF(B' + r + '="","",B' + r + '/30.4)']);
    mantenimientoDiario.push(['=IF(D' + r + '="","",D' + r + '/30.4)']);
    costoGasolinaKm.push([
      '=IF(OR(F' + r + '="",G' + r + '=""),"",IF(G' + r + '="Diesel",Config!$B$1,Config!$B$2)/F' + r + ')'
    ]);
  }
  sheet.getRange(2, 3, FILAS_CATALOGO, 1).setFormulas(costoDiario);
  sheet.getRange(2, 5, FILAS_CATALOGO, 1).setFormulas(mantenimientoDiario);
  sheet.getRange(2, 8, FILAS_CATALOGO, 1).setFormulas(costoGasolinaKm);

  // Nombres alineados con el tarifario real de clientes (Tarifas Vigentes),
  // para que Tipo de Unidad sí cruce con esa referencia.
  var ejemplo = [
    ['Auto', 8000, 4000, 12.0, 'Gasolina'],
    ['Small Van - 1 tn', 11000, 5500, 8.0, 'Gasolina'],
    ['Large Van - 1.5 tn', 15000, 8000, 6.0, 'Diesel'],
    ['3.5 T caja seca', 18000, 10000, 4.5, 'Diesel']
  ];
  sheet.getRange(2, 1, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[0]]; }));
  sheet.getRange(2, 2, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[1]]; }));
  sheet.getRange(2, 4, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[2]]; }));
  sheet.getRange(2, 6, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[3]]; }));
  sheet.getRange(2, 7, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[4]]; }));

  sheet.getRange(2, 2, FILAS_CATALOGO, 4).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 6, FILAS_CATALOGO, 1).setNumberFormat('0.00');
  sheet.getRange(2, 8, FILAS_CATALOGO, 1).setNumberFormat('$#,##0.00');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * "Nomina": catalogo por Puesto / Categoria de Sueldo (NO por Tipo de
 * Unidad). Incluye puestos principales y de auxiliar; cual aplica a cada
 * ruta se resuelve en "Ruta-Zona-Puesto".
 */
function setupNomina_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_NOMINA) || ss.insertSheet(HOJA_NOMINA);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = [
    'Puesto / Categoría de Sueldo', 'Sueldo', 'Periodicidad de Sueldo',
    'Fiscal', 'Complemento', 'IMSS', 'Comisiones', 'Sueldo Mensual Total'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#2b4a3f').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var periodicidadRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Semanal', 'Quincenal'], true).setAllowInvalid(false).build();
  sheet.getRange(2, 3, FILAS_CATALOGO, 1).setDataValidation(periodicidadRule);

  var fiscal = [];
  var complemento = [];
  var imss = [];
  var comisiones = [];
  var sueldoMensualTotal = [];
  for (var i = 0; i < FILAS_CATALOGO; i++) {
    var r = 2 + i;
    fiscal.push(['=IF(C' + r + '="","",IF(C' + r + '="Semanal",Config!$B$3,Config!$B$4))']);
    complemento.push(['=IF(OR(B' + r + '="",D' + r + '=""),"",B' + r + '-D' + r + ')']);
    imss.push(['=IF(D' + r + '="","",D' + r + '*0.3)']);
    comisiones.push(['=IF(E' + r + '="","",E' + r + '*0.05)']);
    sueldoMensualTotal.push([
      '=IF(OR(B' + r + '="",C' + r + '="",F' + r + '="",G' + r + '=""),"",(B' + r + '+F' + r + '+G' + r + ')*IF(C' + r + '="Semanal",4.33,2))'
    ]);
  }
  sheet.getRange(2, 4, FILAS_CATALOGO, 1).setFormulas(fiscal);
  sheet.getRange(2, 5, FILAS_CATALOGO, 1).setFormulas(complemento);
  sheet.getRange(2, 6, FILAS_CATALOGO, 1).setFormulas(imss);
  sheet.getRange(2, 7, FILAS_CATALOGO, 1).setFormulas(comisiones);
  sheet.getRange(2, 8, FILAS_CATALOGO, 1).setFormulas(sueldoMensualTotal);

  var ejemplo = [
    ['Chofer Local', 3200, 'Semanal'],
    ['Chofer Foraneo', 7800, 'Quincenal'],
    ['Auxiliar Local', 1800, 'Semanal'],
    ['Auxiliar Foraneo', 4200, 'Quincenal']
  ];
  sheet.getRange(2, 1, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[0]]; }));
  sheet.getRange(2, 2, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[1]]; }));
  sheet.getRange(2, 3, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[2]]; }));

  sheet.getRange(2, 2, FILAS_CATALOGO, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 4, FILAS_CATALOGO, 5).setNumberFormat('$#,##0.00');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * "Puntos": catalogo de ciudades/puntos que se usan como Punto A
 * (origen) y Punto B (destino) de una ruta, y como Destino en
 * Ruta-Zona-Puesto.
 */
function setupPuntos_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_PUNTOS) || ss.insertSheet(HOJA_PUNTOS);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = ['Punto'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4a3a1c').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var ejemplo = [['CDMX'], ['Querétaro'], ['Puebla'], ['Toluca'], ['Monterrey'], ['Guadalajara']];
  sheet.getRange(2, 1, ejemplo.length, 1).setValues(ejemplo);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * "Casetas": costo de casetas entre Punto A y Punto B (no importa el
 * orden: A->B cuesta lo mismo que B->A). Comercial elige el origen y
 * destino de la ruta, y ese costo se resuelve solo.
 *
 * Los montos de ejemplo son de camion de 2 ejes, tomados de cobertura de
 * prensa sobre las tarifas CAPUFE 2026 (columna Fuente) — son un punto de
 * partida a validar/actualizar con el PDF oficial de CAPUFE para el tipo
 * de unidad real de la flota, no un dato exacto por vehiculo.
 */
function setupCasetas_(hojaPuntos) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_CASETAS) || ss.insertSheet(HOJA_CASETAS);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = ['Punto A', 'Punto B', 'Costo Casetas', 'Fuente'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4a3a1c').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var puntoRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaPuntos.getRange(2, 1, FILAS_CATALOGO, 1), true).setAllowInvalid(false).build();
  sheet.getRange(2, 1, FILAS_CASETAS, 1).setDataValidation(puntoRule);
  sheet.getRange(2, 2, FILAS_CASETAS, 1).setDataValidation(puntoRule);

  var ejemplo = [
    ['CDMX', 'Querétaro', 490, 'Camión 2 ejes, prensa CAPUFE abr-2026 (nmas.com.mx) — validar en PDF oficial CAPUFE'],
    ['CDMX', 'Puebla', 650, 'Camión 2-3 ejes, prensa CAPUFE abr-2026 (nmas.com.mx) — validar en PDF oficial CAPUFE'],
    ['CDMX', 'Toluca', 260, 'Camión/autobús 2 ejes, tramo La Marquesa, prensa CAPUFE abr-2026 (milenio.com) — validar en PDF oficial CAPUFE']
  ];
  sheet.getRange(2, 1, ejemplo.length, 4).setValues(ejemplo);
  sheet.getRange(2, 3, FILAS_CASETAS, 1).setNumberFormat('$#,##0.00');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * "Ruta-Zona-Puesto": por cada combinacion de Tipo de Ruta + Destino
 * (Punto B), que Puesto principal (y cual Auxiliar, si la ruta lo
 * necesita) aplica. Esto es lo que permite que Comercial elija Tipo de
 * Ruta + Punto A/Punto B sin tener que saber nada de nomina.
 */
function setupRutaZonaPuesto_(hojaPuntos, hojaNomina) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_RUTA_ZONA_PUESTO) || ss.insertSheet(HOJA_RUTA_ZONA_PUESTO);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = ['Tipo de Ruta', 'Destino (Punto B)', 'Puesto Principal', 'Puesto Auxiliar'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4a1c3a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var tipoRutaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TIPOS_RUTA, true).setAllowInvalid(false).build();
  sheet.getRange(2, 1, FILAS_RUTA_ZONA, 1).setDataValidation(tipoRutaRule);

  var puntoRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaPuntos.getRange(2, 1, FILAS_CATALOGO, 1), true).setAllowInvalid(false).build();
  sheet.getRange(2, 2, FILAS_RUTA_ZONA, 1).setDataValidation(puntoRule);

  var puestoRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaNomina.getRange(2, 1, FILAS_CATALOGO, 1), true).setAllowInvalid(false).build();
  sheet.getRange(2, 3, FILAS_RUTA_ZONA, 1).setDataValidation(puestoRule);
  sheet.getRange(2, 4, FILAS_RUTA_ZONA, 1).setDataValidation(puestoRule);

  var ejemplo = [
    ['Local', 'CDMX', 'Chofer Local', 'Auxiliar Local'],
    ['Foráneo', 'Monterrey', 'Chofer Foraneo', 'Auxiliar Foraneo'],
    ['Foráneo', 'Guadalajara', 'Chofer Foraneo', 'Auxiliar Foraneo']
  ];
  sheet.getRange(2, 1, ejemplo.length, 4).setValues(ejemplo);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * "Cotizador": vista de costo directo (uso interno de Vanessa). Tipo de
 * Unidad y Puesto / Categoria de Sueldo se eligen directamente, sin pasar
 * por Zona/Tipo de Ruta.
 */
function setupCotizador_(hojaCostosUnidad, hojaNomina) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(COTIZADOR) || ss.insertSheet(COTIZADOR);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = [
    'Fecha', 'Ruta / Cliente', 'Tipo de Unidad', 'Puesto / Categoría de Sueldo', 'Kilometros', 'Paradas',
    'Viajes al Mes', 'Costo Casetas', 'Margen',
    'Sueldo Mensual', 'Costo Gasolina/KM', 'Renta Mensual', 'Mantenimiento Mensual',
    'Costo Variable', 'Costo Fijo Prorrateado', 'Costo Total', 'Tarifa Piso'
  ];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1c2b4a')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var unidadRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaCostosUnidad.getRange(2, 1, FILAS_CATALOGO, 1), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, FILAS_COTIZADOR, 1).setDataValidation(unidadRule);

  var puestoRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaNomina.getRange(2, 1, FILAS_CATALOGO, 1), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, FILAS_COTIZADOR, 1).setDataValidation(puestoRule);

  sheet.getRange(2, 9, FILAS_COTIZADOR, 1).setValue(0.2).setNumberFormat('0%');
  sheet.getRange(2, 8, FILAS_COTIZADOR, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 10, FILAS_COTIZADOR, 8).setNumberFormat('$#,##0.00');

  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * "Solicitudes": registro de las solicitudes de tarifa de cliente hechas
 * por Comercial desde el dashboard (caracteristicas de ruta, no costos).
 * Solo se llena desde la Web App (guardarSolicitudWeb), como historial.
 */
function setupSolicitudes_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_SOLICITUDES) || ss.insertSheet(HOJA_SOLICITUDES);
  sheet.clear();

  var headers = [
    'Fecha', 'Cliente', 'Referencia de Ruta', 'Tipo de Ruta', 'Punto A', 'Punto B', 'Tipo de Unidad', 'Frecuencia', 'Volumen',
    'Kilometros', 'Tipo de Cobro', 'Cantidad', 'Requiere Auxiliar', 'Estacion MELI',
    'Puesto Principal', 'Puesto Auxiliar', 'Costo Casetas', 'Viajes al Mes',
    'Sueldo Mensual', 'Renta Mensual', 'Mantenimiento Mensual', 'Costo Gasolina/KM',
    'Costo Variable', 'Costo Fijo Prorrateado', 'Costo Total',
    'Margen Piso', 'Margen Objetivo', 'Tarifa Piso', 'Tarifa Objetivo',
    'Tarifa por Unidad (Piso)', 'Tarifa por Unidad (Objetivo)', 'Tarifa Vigente (referencia)'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1c2b4a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}
