/**
 * Nombres de las hojas usadas por el cotizador.
 */
var HOJA_MAESTRA = 'Hoja Maestra';
var COTIZADOR = 'Cotizador';

/**
 * Filas reservadas para captura en la hoja Cotizador (dropdown + margen
 * por defecto se aplican hasta esta fila).
 */
var FILAS_COTIZADOR = 500;

/**
 * Filas reservadas para catalogo en la Hoja Maestra (dropdowns y formulas
 * de costo se aplican hasta esta fila).
 */
var FILAS_MAESTRA = 50;

/**
 * Fila donde inician los datos de las tablas de la Hoja Maestra
 * (fila 1: titulo de cada tabla, fila 2: encabezados).
 */
var FILA_DATOS_MAESTRA = 3;

/**
 * Crea (o reinicia) la hoja "Hoja Maestra": dos catalogos independientes,
 * uno por Tipo de Unidad (renta, mantenimiento, combustible) y otro por
 * Puesto / Categoria de Sueldo (nomina). El sueldo NO depende de la unidad:
 * en el Cotizador se eligen ambos por separado para una misma ruta.
 *
 * Nomina: el sueldo se captura por periodo (Semanal o Quincenal). La parte
 * "Fiscal" es un monto fijo segun el periodo (ver bloque de referencia en
 * S1:T4), el "Complemento" es el resto del sueldo, el IMSS es 30% de la
 * parte fiscal y la comision es 5% del complemento; todo se convierte
 * despues a un costo mensual total.
 */
function setupHojaMaestra_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_MAESTRA) || ss.insertSheet(HOJA_MAESTRA);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  // Titulos de cada tabla (fila 1).
  sheet.getRange(1, 1, 1, 8).merge().setValue('Costos por Tipo de Unidad')
    .setFontWeight('bold').setBackground('#1c2b4a').setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.getRange(1, 10, 1, 8).merge().setValue('Nómina (por Puesto / Categoría de Sueldo)')
    .setFontWeight('bold').setBackground('#2b4a3f').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Encabezados (fila 2).
  var headersUnidad = [
    'Tipo de Unidad', 'Renta Mensual', 'Costo Diario', 'Mantenimiento Mensual', 'Mantenimiento Diario',
    'Rendimiento (km/L)', 'Tipo de Combustible', 'Costo Gasolina por KM'
  ];
  sheet.getRange(2, 1, 1, headersUnidad.length).setValues([headersUnidad])
    .setFontWeight('bold').setBackground('#1c2b4a').setFontColor('#ffffff');

  var headersNomina = [
    'Puesto / Categoría de Sueldo', 'Sueldo', 'Periodicidad de Sueldo',
    'Fiscal', 'Complemento', 'IMSS', 'Comisiones', 'Sueldo Mensual Total'
  ];
  sheet.getRange(2, 10, 1, headersNomina.length).setValues([headersNomina])
    .setFontWeight('bold').setBackground('#2b4a3f').setFontColor('#ffffff');

  sheet.setFrozenRows(2);

  // Bloque de referencia: precios de combustible y parte fiscal por
  // periodo. Editable sin tocar el script cuando cambien.
  sheet.getRange('S1:S4').setValues([
    ['Precio Diesel ($/L)'], ['Precio Gasolina ($/L)'], ['Fiscal Semanal'], ['Fiscal Quincenal']
  ]).setFontWeight('bold');
  sheet.getRange('T1:T4').setValues([[24.5], [23.8], [2253.6], [5190.6]]).setNumberFormat('$#,##0.00');

  // Listas desplegables.
  var combustibleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Diesel', 'Gasolina'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(FILA_DATOS_MAESTRA, 7, FILAS_MAESTRA, 1).setDataValidation(combustibleRule);

  var periodicidadRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Semanal', 'Quincenal'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(FILA_DATOS_MAESTRA, 12, FILAS_MAESTRA, 1).setDataValidation(periodicidadRule);

  // Formulas de costo, una por fila, hasta FILAS_MAESTRA.
  var costoDiario = [];
  var mantenimientoDiario = [];
  var costoGasolinaKm = [];
  var fiscal = [];
  var complemento = [];
  var imss = [];
  var comisiones = [];
  var sueldoMensualTotal = [];
  for (var i = 0; i < FILAS_MAESTRA; i++) {
    var r = FILA_DATOS_MAESTRA + i;
    costoDiario.push(['=IF(B' + r + '="","",B' + r + '/30.4)']);
    mantenimientoDiario.push(['=IF(D' + r + '="","",D' + r + '/30.4)']);
    costoGasolinaKm.push(['=IF(OR(F' + r + '="",G' + r + '=""),"",IF(G' + r + '="Diesel",$T$1,$T$2)/F' + r + ')']);
    fiscal.push(['=IF(L' + r + '="","",IF(L' + r + '="Semanal",$T$3,$T$4))']);
    complemento.push(['=IF(OR(K' + r + '="",M' + r + '=""),"",K' + r + '-M' + r + ')']);
    imss.push(['=IF(M' + r + '="","",M' + r + '*0.3)']);
    comisiones.push(['=IF(N' + r + '="","",N' + r + '*0.05)']);
    sueldoMensualTotal.push([
      '=IF(OR(K' + r + '="",L' + r + '="",O' + r + '="",P' + r + '=""),"",(K' + r + '+O' + r + '+P' + r + ')*IF(L' + r + '="Semanal",4.33,2))'
    ]);
  }
  sheet.getRange(FILA_DATOS_MAESTRA, 3, FILAS_MAESTRA, 1).setFormulas(costoDiario);
  sheet.getRange(FILA_DATOS_MAESTRA, 5, FILAS_MAESTRA, 1).setFormulas(mantenimientoDiario);
  sheet.getRange(FILA_DATOS_MAESTRA, 8, FILAS_MAESTRA, 1).setFormulas(costoGasolinaKm);
  sheet.getRange(FILA_DATOS_MAESTRA, 13, FILAS_MAESTRA, 1).setFormulas(fiscal);
  sheet.getRange(FILA_DATOS_MAESTRA, 14, FILAS_MAESTRA, 1).setFormulas(complemento);
  sheet.getRange(FILA_DATOS_MAESTRA, 15, FILAS_MAESTRA, 1).setFormulas(imss);
  sheet.getRange(FILA_DATOS_MAESTRA, 16, FILAS_MAESTRA, 1).setFormulas(comisiones);
  sheet.getRange(FILA_DATOS_MAESTRA, 17, FILAS_MAESTRA, 1).setFormulas(sueldoMensualTotal);

  // Datos de ejemplo (solo columnas de captura): reemplazar con los costos
  // reales de BDB. Las dos tablas son independientes entre si.
  var ejemploUnidad = [
    ['Torton', 15000, 8000, 3.0, 'Diesel'],
    ['Rabon', 12000, 6000, 3.6, 'Gasolina'],
    ['Tractocamion', 18000, 12000, 2.4, 'Diesel']
  ];
  sheet.getRange(FILA_DATOS_MAESTRA, 1, ejemploUnidad.length, 1).setValues(ejemploUnidad.map(function(f) { return [f[0]]; }));
  sheet.getRange(FILA_DATOS_MAESTRA, 2, ejemploUnidad.length, 1).setValues(ejemploUnidad.map(function(f) { return [f[1]]; }));
  sheet.getRange(FILA_DATOS_MAESTRA, 4, ejemploUnidad.length, 1).setValues(ejemploUnidad.map(function(f) { return [f[2]]; }));
  sheet.getRange(FILA_DATOS_MAESTRA, 6, ejemploUnidad.length, 1).setValues(ejemploUnidad.map(function(f) { return [f[3]]; }));
  sheet.getRange(FILA_DATOS_MAESTRA, 7, ejemploUnidad.length, 1).setValues(ejemploUnidad.map(function(f) { return [f[4]]; }));

  var ejemploNomina = [
    ['Chofer Local', 3200, 'Semanal'],
    ['Chofer Foraneo', 7800, 'Quincenal'],
    ['Operador Especializado', 9500, 'Quincenal']
  ];
  sheet.getRange(FILA_DATOS_MAESTRA, 10, ejemploNomina.length, 1).setValues(ejemploNomina.map(function(f) { return [f[0]]; }));
  sheet.getRange(FILA_DATOS_MAESTRA, 11, ejemploNomina.length, 1).setValues(ejemploNomina.map(function(f) { return [f[1]]; }));
  sheet.getRange(FILA_DATOS_MAESTRA, 12, ejemploNomina.length, 1).setValues(ejemploNomina.map(function(f) { return [f[2]]; }));

  // Formatos.
  sheet.getRange(FILA_DATOS_MAESTRA, 2, FILAS_MAESTRA, 4).setNumberFormat('$#,##0.00'); // Renta, Costo Diario, Mantenimiento Mensual, Mantenimiento Diario
  sheet.getRange(FILA_DATOS_MAESTRA, 6, FILAS_MAESTRA, 1).setNumberFormat('0.00'); // Rendimiento
  sheet.getRange(FILA_DATOS_MAESTRA, 8, FILAS_MAESTRA, 1).setNumberFormat('$#,##0.00'); // Costo Gasolina por KM
  sheet.getRange(FILA_DATOS_MAESTRA, 11, FILAS_MAESTRA, 1).setNumberFormat('$#,##0.00'); // Sueldo
  sheet.getRange(FILA_DATOS_MAESTRA, 13, FILAS_MAESTRA, 5).setNumberFormat('$#,##0.00'); // Fiscal, Complemento, IMSS, Comisiones, Sueldo Mensual Total

  sheet.autoResizeColumns(1, 20); // A:T (tabla Unidad, tabla Nomina y bloque de referencia)

  return sheet;
}

/**
 * Crea (o reinicia) la hoja "Cotizador": captura de una ruta nueva con
 * lista desplegable de Tipo de Unidad y de Puesto / Categoria de Sueldo
 * (independientes entre si) y las columnas calculadas por COTIZAR().
 */
function setupCotizador_(hojaMaestra) {
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

  // Listas desplegables tomadas de la Hoja Maestra. Google Sheets omite las
  // celdas vacias del rango al mostrar las opciones, asi que un rango fijo
  // no muestra huecos en blanco.
  var unidadRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaMaestra.getRange(FILA_DATOS_MAESTRA, 1, FILAS_MAESTRA, 1), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, FILAS_COTIZADOR, 1).setDataValidation(unidadRule);

  var puestoRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaMaestra.getRange(FILA_DATOS_MAESTRA, 10, FILAS_MAESTRA, 1), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, FILAS_COTIZADOR, 1).setDataValidation(puestoRule);

  // Margen por defecto (editable por fila).
  sheet.getRange(2, 9, FILAS_COTIZADOR, 1).setValue(0.2).setNumberFormat('0%');

  sheet.getRange(2, 8, FILAS_COTIZADOR, 1).setNumberFormat('$#,##0.00'); // Costo Casetas
  sheet.getRange(2, 10, FILAS_COTIZADOR, 8).setNumberFormat('$#,##0.00'); // Sueldo Mensual ... Tarifa Piso

  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}
