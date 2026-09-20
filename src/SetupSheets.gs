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
 * Crea (o reinicia) la hoja "Hoja Maestra": el catalogo de tipos de unidad
 * con sus costos base (renta, mantenimiento, combustible y nomina).
 *
 * Nomina: el sueldo se captura por periodo (Semanal o Quincenal). La parte
 * "Fiscal" es un monto fijo segun el periodo (ver bloque de referencia en
 * Q1:R4), el "Complemento" es el resto del sueldo, el IMSS es 30% de la
 * parte fiscal y la comision es 5% del complemento; todo se convierte
 * despues a un costo mensual total.
 */
function setupHojaMaestra_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_MAESTRA) || ss.insertSheet(HOJA_MAESTRA);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = [
    'Tipo de Unidad', 'Renta Mensual', 'Costo Diario', 'Mantenimiento Mensual', 'Mantenimiento Diario',
    'Rendimiento (km/L)', 'Tipo de Combustible', 'Costo Gasolina por KM',
    'Sueldo', 'Periodicidad de Sueldo', 'Fiscal', 'Complemento', 'IMSS', 'Comisiones', 'Sueldo Mensual Total'
  ];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1c2b4a')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Bloque de referencia: precios de combustible y parte fiscal por
  // periodo. Editable sin tocar el script cuando cambien.
  sheet.getRange('Q1:Q4').setValues([
    ['Precio Diesel ($/L)'], ['Precio Gasolina ($/L)'], ['Fiscal Semanal'], ['Fiscal Quincenal']
  ]).setFontWeight('bold');
  sheet.getRange('R1:R4').setValues([[24.5], [23.8], [2253.6], [5190.6]]).setNumberFormat('$#,##0.00');

  // Listas desplegables.
  var combustibleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Diesel', 'Gasolina'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, FILAS_MAESTRA, 1).setDataValidation(combustibleRule);

  var periodicidadRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Semanal', 'Quincenal'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 10, FILAS_MAESTRA, 1).setDataValidation(periodicidadRule);

  // Formulas de costo, una por fila, hasta FILAS_MAESTRA.
  var costoDiario = [];
  var mantenimientoDiario = [];
  var costoGasolinaKm = [];
  var fiscal = [];
  var complemento = [];
  var imss = [];
  var comisiones = [];
  var sueldoMensualTotal = [];
  for (var r = 2; r < 2 + FILAS_MAESTRA; r++) {
    costoDiario.push(['=IF(B' + r + '="","",B' + r + '/30.4)']);
    mantenimientoDiario.push(['=IF(D' + r + '="","",D' + r + '/30.4)']);
    costoGasolinaKm.push(['=IF(OR(F' + r + '="",G' + r + '=""),"",IF(G' + r + '="Diesel",$R$1,$R$2)/F' + r + ')']);
    fiscal.push(['=IF(J' + r + '="","",IF(J' + r + '="Semanal",$R$3,$R$4))']);
    complemento.push(['=IF(OR(I' + r + '="",K' + r + '=""),"",I' + r + '-K' + r + ')']);
    imss.push(['=IF(K' + r + '="","",K' + r + '*0.3)']);
    comisiones.push(['=IF(L' + r + '="","",L' + r + '*0.05)']);
    sueldoMensualTotal.push([
      '=IF(OR(I' + r + '="",J' + r + '="",M' + r + '="",N' + r + '=""),"",(I' + r + '+M' + r + '+N' + r + ')*IF(J' + r + '="Semanal",4.33,2))'
    ]);
  }
  sheet.getRange(2, 3, FILAS_MAESTRA, 1).setFormulas(costoDiario);
  sheet.getRange(2, 5, FILAS_MAESTRA, 1).setFormulas(mantenimientoDiario);
  sheet.getRange(2, 8, FILAS_MAESTRA, 1).setFormulas(costoGasolinaKm);
  sheet.getRange(2, 11, FILAS_MAESTRA, 1).setFormulas(fiscal);
  sheet.getRange(2, 12, FILAS_MAESTRA, 1).setFormulas(complemento);
  sheet.getRange(2, 13, FILAS_MAESTRA, 1).setFormulas(imss);
  sheet.getRange(2, 14, FILAS_MAESTRA, 1).setFormulas(comisiones);
  sheet.getRange(2, 15, FILAS_MAESTRA, 1).setFormulas(sueldoMensualTotal);

  // Datos de ejemplo (solo columnas de captura): reemplazar con los costos
  // reales de BDB. Tipo de Unidad, Renta Mensual, Mantenimiento Mensual,
  // Rendimiento, Tipo de Combustible, Sueldo, Periodicidad de Sueldo.
  var ejemplo = [
    ['Torton', 15000, 8000, 3.0, 'Diesel', 2600, 'Quincenal'],
    ['Rabon', 12000, 6000, 3.6, 'Gasolina', 1450, 'Semanal'],
    ['Tractocamion', 18000, 12000, 2.4, 'Diesel', 3200, 'Quincenal']
  ];
  sheet.getRange(2, 1, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[0]]; }));
  sheet.getRange(2, 2, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[1]]; }));
  sheet.getRange(2, 4, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[2]]; }));
  sheet.getRange(2, 6, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[3]]; }));
  sheet.getRange(2, 7, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[4]]; }));
  sheet.getRange(2, 9, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[5]]; }));
  sheet.getRange(2, 10, ejemplo.length, 1).setValues(ejemplo.map(function(f) { return [f[6]]; }));

  // Formatos.
  sheet.getRange(2, 2, FILAS_MAESTRA, 4).setNumberFormat('$#,##0.00'); // Renta, Costo Diario, Mantenimiento Mensual, Mantenimiento Diario
  sheet.getRange(2, 6, FILAS_MAESTRA, 1).setNumberFormat('0.00'); // Rendimiento
  sheet.getRange(2, 8, FILAS_MAESTRA, 1).setNumberFormat('$#,##0.00'); // Costo Gasolina por KM
  sheet.getRange(2, 9, FILAS_MAESTRA, 1).setNumberFormat('$#,##0.00'); // Sueldo
  sheet.getRange(2, 11, FILAS_MAESTRA, 5).setNumberFormat('$#,##0.00'); // Fiscal, Complemento, IMSS, Comisiones, Sueldo Mensual Total

  sheet.autoResizeColumns(1, headers.length);
  sheet.autoResizeColumns(17, 2); // Q:R

  return sheet;
}

/**
 * Crea (o reinicia) la hoja "Cotizador": captura de una ruta nueva con
 * lista desplegable de unidad y las columnas calculadas por COTIZAR().
 */
function setupCotizador_(hojaMaestra) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(COTIZADOR) || ss.insertSheet(COTIZADOR);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();

  var headers = [
    'Fecha', 'Ruta / Cliente', 'Tipo de Unidad', 'Kilometros', 'Paradas',
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

  // Lista desplegable de Tipo de Unidad tomada de la Hoja Maestra. Google
  // Sheets omite las celdas vacias del rango al mostrar las opciones, asi
  // que un rango fijo no muestra huecos en blanco.
  var validationRange = hojaMaestra.getRange(2, 1, FILAS_MAESTRA, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, FILAS_COTIZADOR, 1).setDataValidation(rule);

  // Margen por defecto (editable por fila).
  sheet.getRange(2, 8, FILAS_COTIZADOR, 1).setValue(0.2).setNumberFormat('0%');

  sheet.getRange(2, 7, FILAS_COTIZADOR, 1).setNumberFormat('$#,##0.00'); // Costo Casetas
  sheet.getRange(2, 9, FILAS_COTIZADOR, 8).setNumberFormat('$#,##0.00'); // Sueldo Mensual ... Tarifa Piso

  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}
