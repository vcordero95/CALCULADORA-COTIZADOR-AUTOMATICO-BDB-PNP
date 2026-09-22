/**
 * Matematica compartida de costo, una vez que ya se resolvieron renta,
 * mantenimiento, gasolina/km, sueldo mensual, km, viajes al mes y casetas
 * (sin importar si esos datos vinieron de COTIZAR(), que los toma
 * directo, o de SOLICITAR_TARIFA(), que los resuelve a partir de Tipo de
 * Ruta / Punto A / Punto B / Frecuencia). No incluye el margen: eso lo
 * aplica cada quien por separado con aplicarMargen_().
 */
function calcularCostoRuta_(rentaMensual, mantenimientoMensual, gasolinaKm, sueldoMensual, km, viajesMes, casetas) {
  var viajes = Number(viajesMes) || 0;
  if (viajes <= 0) {
    throw new Error('Falta capturar "Viajes al mes" (mayor a 0) para prorratear los costos fijos.');
  }

  var kilometros = Number(km) || 0;
  var costoCasetas = Number(casetas) || 0;

  var costoVariable = Number(gasolinaKm) * kilometros;
  var costoFijoMensual = Number(sueldoMensual) + Number(rentaMensual) + Number(mantenimientoMensual);
  var costoFijoProrrateado = costoFijoMensual / viajes;
  var costoTotal = costoVariable + costoFijoProrrateado + costoCasetas;

  return {
    costoVariable: costoVariable,
    costoFijoProrrateado: costoFijoProrrateado,
    costoTotal: costoTotal
  };
}

/** Tarifa = Costo ÷ (1 - Margen). Margen como fraccion (0.20 = 20%). */
function aplicarMargen_(costoTotal, margen) {
  var margenPct = Number(margen) || 0;
  if (margenPct >= 1) {
    throw new Error('El margen debe ser menor a 100%.');
  }
  return costoTotal / (1 - margenPct);
}

function obtenerHojaCostosUnidad_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_COSTOS_UNIDAD);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_COSTOS_UNIDAD + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return hoja;
}

function obtenerHojaNomina_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_NOMINA);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_NOMINA + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return hoja;
}

function obtenerHojaPuntos_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PUNTOS);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_PUNTOS + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return hoja;
}

/** Busca un Tipo de Unidad en la hoja Costos Unidad. Regresa null si no existe. */
function buscarUnidadPorTipo_(tipoUnidad) {
  var datos = obtenerHojaCostosUnidad_().getRange(2, 1, FILAS_CATALOGO, 8).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === tipoUnidad) return datos[i];
  }
  return null;
}

/** Busca un Puesto / Categoria de Sueldo en la hoja Nomina. Regresa null si no existe. */
function buscarPuestoEnNomina_(puesto) {
  var datos = obtenerHojaNomina_().getRange(2, 1, FILAS_CATALOGO, 8).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === puesto) return datos[i];
  }
  return null;
}

/**
 * Calcula el costo y la tarifa piso de una ruta a partir del Tipo de
 * Unidad y el Puesto / Categoria de Sueldo elegidos directamente (uso
 * interno, hoja Cotizador). Son independientes: cada uno se busca en su
 * propio catalogo.
 *
 * Se usa como formula de hoja de calculo, ej.
 * =COTIZAR(C2,D2,E2,G2,H2,I2), y regresa un renglon con 8 columnas:
 * Sueldo Mensual, Costo Gasolina/KM, Renta Mensual, Mantenimiento Mensual,
 * Costo Variable, Costo Fijo Prorrateado, Costo Total, Tarifa Piso.
 *
 * @param {string} tipoUnidad Tipo de unidad, tal como aparece en "Costos Unidad".
 * @param {string} puesto Puesto / Categoria de sueldo, tal como aparece en "Nomina".
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

  var filaUnidad = buscarUnidadPorTipo_(tipoUnidad);
  if (!filaUnidad) {
    throw new Error('Tipo de unidad "' + tipoUnidad + '" no esta en la hoja "' + HOJA_COSTOS_UNIDAD + '".');
  }
  if (filaUnidad[7] === '') {
    throw new Error('Completa Rendimiento y Tipo de Combustible de "' + tipoUnidad + '" en ' + HOJA_COSTOS_UNIDAD + '.');
  }

  var filaNomina = buscarPuestoEnNomina_(puesto);
  if (!filaNomina) {
    throw new Error('Puesto / categoria de sueldo "' + puesto + '" no esta en la hoja "' + HOJA_NOMINA + '".');
  }
  if (filaNomina[7] === '') {
    throw new Error('Completa Sueldo y Periodicidad de "' + puesto + '" en ' + HOJA_NOMINA + '.');
  }

  var rentaMensual = Number(filaUnidad[1]) || 0;
  var mantenimientoMensual = Number(filaUnidad[3]) || 0;
  var gasolinaKm = Number(filaUnidad[7]) || 0;
  var sueldoMensual = Number(filaNomina[7]) || 0;

  var r = calcularCostoRuta_(rentaMensual, mantenimientoMensual, gasolinaKm, sueldoMensual, km, viajesMes, casetas);
  var tarifaPiso = aplicarMargen_(r.costoTotal, margen);

  return [[sueldoMensual, gasolinaKm, rentaMensual, mantenimientoMensual, r.costoVariable, r.costoFijoProrrateado, r.costoTotal, tarifaPiso]];
}
