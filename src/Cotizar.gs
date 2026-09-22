/**
 * Calcula el costo y la tarifa piso de una ruta.
 *
 * El Tipo de Unidad y el Puesto / Categoria de Sueldo son independientes:
 * se buscan cada uno en su propia tabla de la Hoja Maestra (Costos por
 * Tipo de Unidad y Nomina). Se toma el costo de gasolina por km (ya
 * resuelto segun Diesel/Gasolina) y la renta y mantenimiento mensuales de
 * la unidad, se suma el sueldo mensual total del puesto elegido, se
 * prorratea ese costo fijo entre los viajes al mes, se suman las casetas
 * de la ruta y se aplica la tarifa piso: Costo ÷ (1 - Margen).
 *
 * Se usa como formula de hoja de calculo, ej.
 * =COTIZAR(C2,D2,E2,G2,H2,I2), y regresa un renglon con 8 columnas:
 * Sueldo Mensual, Costo Gasolina/KM, Renta Mensual, Mantenimiento Mensual,
 * Costo Variable, Costo Fijo Prorrateado, Costo Total, Tarifa Piso.
 *
 * @param {string} tipoUnidad Tipo de unidad, tal como aparece en la tabla "Costos por Tipo de Unidad".
 * @param {string} puesto Puesto / Categoria de sueldo, tal como aparece en la tabla "Nomina".
 * @param {number} km Kilometros de la ruta.
 * @param {number} viajesMes Viajes al mes, para prorratear los costos fijos.
 * @param {number} casetas Costo de casetas de la ruta o corredor.
 * @param {number} margen Margen objetivo sobre venta, como fraccion (0.20 = 20%).
 * @return {Array<Array<number>>} Renglon con los 8 valores calculados.
 * @customfunction
 */
function COTIZAR(tipoUnidad, puesto, km, viajesMes, casetas, margen) {
  if (tipoUnidad === '' || tipoUnidad === undefined || tipoUnidad === null ||
      puesto === '' || puesto === undefined || puesto === null) {
    return [['', '', '', '', '', '', '', '']];
  }

  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_MAESTRA);
  if (!hoja) {
    throw new Error('No existe la hoja "Hoja Maestra". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  // Tabla "Costos por Tipo de Unidad": A Tipo de Unidad, B Renta Mensual,
  // D Mantenimiento Mensual, H Costo Gasolina por KM.
  var catalogoUnidad = hoja.getRange(FILA_DATOS_MAESTRA, 1, FILAS_MAESTRA, 8).getValues();
  var filaUnidad = null;
  for (var i = 0; i < catalogoUnidad.length; i++) {
    if (catalogoUnidad[i][0] === tipoUnidad) {
      filaUnidad = catalogoUnidad[i];
      break;
    }
  }
  if (!filaUnidad) {
    throw new Error('Tipo de unidad "' + tipoUnidad + '" no esta en la tabla "Costos por Tipo de Unidad".');
  }
  if (filaUnidad[7] === '') {
    throw new Error('Completa Rendimiento y Tipo de Combustible de "' + tipoUnidad + '" en la Hoja Maestra.');
  }

  // Tabla "Nomina": J Puesto / Categoria de Sueldo, Q Sueldo Mensual Total.
  var catalogoNomina = hoja.getRange(FILA_DATOS_MAESTRA, 10, FILAS_MAESTRA, 8).getValues();
  var filaNomina = null;
  for (var j = 0; j < catalogoNomina.length; j++) {
    if (catalogoNomina[j][0] === puesto) {
      filaNomina = catalogoNomina[j];
      break;
    }
  }
  if (!filaNomina) {
    throw new Error('Puesto / categoria de sueldo "' + puesto + '" no esta en la tabla "Nomina".');
  }
  if (filaNomina[7] === '') {
    throw new Error('Completa Sueldo y Periodicidad de "' + puesto + '" en la Hoja Maestra.');
  }

  var rentaMensual = Number(filaUnidad[1]) || 0;
  var mantenimientoMensual = Number(filaUnidad[3]) || 0;
  var gasolinaKm = Number(filaUnidad[7]) || 0;
  var sueldoMensual = Number(filaNomina[7]) || 0;

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
