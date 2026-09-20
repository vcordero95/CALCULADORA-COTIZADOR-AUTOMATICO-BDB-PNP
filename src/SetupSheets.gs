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
 * Crea (o reinicia) la hoja "Hoja Maestra": el catálogo de tipos de unidad
 * con sus costos base (sueldo, gasolina por km y renta).
 */
function setupHojaMaestra_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HOJA_MAESTRA) || ss.insertSheet(HOJA_MAESTRA);
  sheet.clear();

  var headers = ['Tipo de Unidad', 'Sueldo Mensual', 'Costo Gasolina por KM', 'Renta Mensual'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1c2b4a')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Datos de ejemplo: reemplazar con los costos reales de BDB.
  var ejemplo = [
    ['Torton', 15000, 5.50, 8000],
    ['Rabon', 12000, 4.20, 6000],
    ['Tractocamion', 18000, 7.80, 12000]
  ];
  sheet.getRange(2, 1, ejemplo.length, headers.length).setValues(ejemplo);
  sheet.getRange(2, 2, ejemplo.length, 3).setNumberFormat('$#,##0.00');
  sheet.autoResizeColumns(1, headers.length);

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

  var headers = [
    'Fecha', 'Ruta / Cliente', 'Tipo de Unidad', 'Kilometros', 'Paradas',
    'Viajes al Mes', 'Costo Casetas', 'Margen',
    'Sueldo Mensual', 'Costo Gasolina/KM', 'Renta Mensual',
    'Costo Variable', 'Costo Fijo Prorrateado', 'Costo Total', 'Tarifa Piso'
  ];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1c2b4a')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Lista desplegable de Tipo de Unidad tomada de la Hoja Maestra.
  var lastMaestraRow = Math.max(hojaMaestra.getLastRow(), 2);
  var validationRange = hojaMaestra.getRange(2, 1, lastMaestraRow - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, FILAS_COTIZADOR, 1).setDataValidation(rule);

  // Margen por defecto (editable por fila).
  sheet.getRange(2, 8, FILAS_COTIZADOR, 1).setValue(0.2).setNumberFormat('0%');

  sheet.getRange(2, 7, FILAS_COTIZADOR, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 9, FILAS_COTIZADOR, 6).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 15, FILAS_COTIZADOR, 1).setNumberFormat('$#,##0.00');

  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}
