/**
 * Calcula el costo y la tarifa piso de una ruta.
 *
 * Busca en la Hoja Maestra, para el tipo de unidad dado, el costo de
 * gasolina por km (ya resuelto segun Diesel/Gasolina) y el sueldo mensual
 * total (nomina ya convertida a mensual), los suma a la renta y el
 * mantenimiento mensuales, prorratea ese costo fijo entre los viajes al
 * mes, suma las casetas de la ruta y aplica la tarifa piso:
 * Costo ÷ (1 - Margen).
 *
 * Se usa como formula de hoja de calculo, ej. =COTIZAR(C2,D2,F2,G2,H2),
 * y regresa un renglon con 8 columnas: Sueldo Mensual, Costo Gasolina/KM,
 * Renta Mensual, Mantenimiento Mensual, Costo Variable, Costo Fijo
 * Prorrateado, Costo Total, Tarifa Piso.
 *
 * @param {string} tipoUnidad Tipo de unidad, tal como aparece en la Hoja Maestra.
 * @param {number} km Kilometros de la ruta.
 * @param {number} viajesMes Viajes al mes, para prorratear los costos fijos.
 * @param {number} casetas Costo de casetas de la ruta o corredor.
 * @param {number} margen Margen objetivo sobre venta, como fraccion (0.20 = 20%).
 * @return {Array<Array<number>>} Renglon con los 8 valores calculados.
 * @customfunction
 */
function COTIZAR(tipoUnidad, km, viajesMes, casetas, margen) {
  if (tipoUnidad === '' || tipoUnidad === undefined || tipoUnidad === null) {
    return [['', '', '', '', '', '', '', '']];
  }

  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_MAESTRA);
  if (!hoja) {
    throw new Error('No existe la hoja "Hoja Maestra". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  var lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    throw new Error('La Hoja Maestra no tiene tipos de unidad capturados.');
  }

  // A: Tipo de Unidad, B: Renta Mensual, D: Mantenimiento Mensual,
  // H: Costo Gasolina por KM, O: Sueldo Mensual Total.
  var catalogo = hoja.getRange(2, 1, lastRow - 1, 15).getValues();
  var fila = null;
  for (var i = 0; i < catalogo.length; i++) {
    if (catalogo[i][0] === tipoUnidad) {
      fila = catalogo[i];
      break;
    }
  }
  if (!fila) {
    throw new Error('Tipo de unidad "' + tipoUnidad + '" no esta en la Hoja Maestra.');
  }
  if (fila[7] === '' || fila[14] === '') {
    throw new Error(
      'Completa Rendimiento/Tipo de Combustible y Sueldo/Periodicidad de "' + tipoUnidad + '" en la Hoja Maestra.'
    );
  }

  var rentaMensual = Number(fila[1]) || 0;
  var mantenimientoMensual = Number(fila[3]) || 0;
  var gasolinaKm = Number(fila[7]) || 0;
  var sueldoMensual = Number(fila[14]) || 0;

  var viajes = Number(viajesMes) || 0;
  if (viajes <= 0) {
    throw new Error('Falta capturar "Viajes al mes" (mayor a 0) para prorratear los costos fijos.');
  }

  var margenPct = Number(margen) || 0;
  if (margenPct >= 1) {
    throw new Error('El margen debe ser menor a 100%.');
  }

  var kilometros = Number(km) || 0;
  var costoCasetas = Number(casetas) || 0;

  var costoVariable = gasolinaKm * kilometros;
  var costoFijoMensual = sueldoMensual + rentaMensual + mantenimientoMensual;
  var costoFijoProrrateado = costoFijoMensual / viajes;
  var costoTotal = costoVariable + costoFijoProrrateado + costoCasetas;
  var tarifaPiso = costoTotal / (1 - margenPct);

  return [[sueldoMensual, gasolinaKm, rentaMensual, mantenimientoMensual, costoVariable, costoFijoProrrateado, costoTotal, tarifaPiso]];
}
