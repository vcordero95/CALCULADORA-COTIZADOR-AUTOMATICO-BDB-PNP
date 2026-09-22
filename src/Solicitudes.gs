/** Busca una Zona / Ciudad en la hoja Zonas. Regresa null si no existe. */
function buscarZona_(zona) {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_ZONAS);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_ZONAS + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  var datos = hoja.getRange(2, 1, FILAS_CATALOGO, 2).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === zona) return datos[i];
  }
  return null;
}

/**
 * Busca, para una combinacion de Tipo de Ruta + Zona, que Puesto
 * principal y cual Auxiliar aplican. Regresa null si esa combinacion no
 * esta definida en la hoja Ruta-Zona-Puesto.
 */
function buscarRutaZonaPuesto_(tipoRuta, zona) {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_RUTA_ZONA_PUESTO);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_RUTA_ZONA_PUESTO + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  var datos = hoja.getRange(2, 1, FILAS_RUTA_ZONA, 4).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === tipoRuta && datos[i][1] === zona) return datos[i];
  }
  return null;
}

/**
 * Convierte una Frecuencia tipo "7x7", "5x7", etc. al numero de veces por
 * semana (el primer numero del patron).
 */
function vecesPorSemana_(frecuencia) {
  var match = String(frecuencia).match(/^(\d+)/);
  if (!match) {
    throw new Error('Frecuencia "' + frecuencia + '" no es valida (usa el formato NxM, ej. 7x7).');
  }
  return Number(match[1]);
}

/**
 * Calcula la tarifa de una solicitud de Comercial a partir de las
 * caracteristicas de la ruta, no de costos: Tipo de Ruta + Zona resuelven
 * el Puesto principal (y el Auxiliar, si la ruta lo necesita) en la hoja
 * Ruta-Zona-Puesto; la Zona tambien resuelve el Costo de Casetas; la
 * Frecuencia resuelve los Viajes al Mes (veces por semana x 4.33). El
 * resto es el mismo calculo de costo/tarifa piso que usa COTIZAR().
 *
 * Si tipoCobro no es "Por Ruta", ademas se regresa la Tarifa Piso entre
 * la Cantidad (paquetes, paradas o palets) capturada.
 *
 * No es una funcion de hoja de calculo (@customfunction): la llama
 * calcularSolicitudWeb() desde el dashboard de Comercial.
 *
 * @return {Object} Desglose completo de la solicitud.
 */
function SOLICITAR_TARIFA(tipoRuta, zona, tipoUnidad, frecuencia, km, tipoCobro, cantidad, requiereAuxiliar, margen) {
  var filaUnidad = buscarUnidadPorTipo_(tipoUnidad);
  if (!filaUnidad) {
    throw new Error('Tipo de unidad "' + tipoUnidad + '" no esta en la hoja "' + HOJA_COSTOS_UNIDAD + '".');
  }
  if (filaUnidad[7] === '') {
    throw new Error('Completa Rendimiento y Tipo de Combustible de "' + tipoUnidad + '" en ' + HOJA_COSTOS_UNIDAD + '.');
  }

  var filaZona = buscarZona_(zona);
  if (!filaZona) {
    throw new Error('Zona / Ciudad "' + zona + '" no esta en la hoja "' + HOJA_ZONAS + '".');
  }

  var mapeo = buscarRutaZonaPuesto_(tipoRuta, zona);
  if (!mapeo) {
    throw new Error(
      'No hay un Puesto definido para Tipo de Ruta "' + tipoRuta + '" + Zona "' + zona +
      '" en la hoja "' + HOJA_RUTA_ZONA_PUESTO + '".'
    );
  }

  var puestoPrincipal = mapeo[2];
  var puestoAuxiliar = mapeo[3];

  var filaPrincipal = buscarPuestoEnNomina_(puestoPrincipal);
  if (!filaPrincipal) {
    throw new Error('Puesto "' + puestoPrincipal + '" (mapeado para ' + tipoRuta + ' + ' + zona + ') no esta en la hoja "' + HOJA_NOMINA + '".');
  }
  if (filaPrincipal[7] === '') {
    throw new Error('Completa Sueldo y Periodicidad de "' + puestoPrincipal + '" en ' + HOJA_NOMINA + '.');
  }

  var sueldoMensual = Number(filaPrincipal[7]) || 0;
  var puestoAuxiliarUsado = '';

  if (requiereAuxiliar) {
    if (!puestoAuxiliar) {
      throw new Error(
        'No hay Puesto Auxiliar definido para Tipo de Ruta "' + tipoRuta + '" + Zona "' + zona +
        '" en la hoja "' + HOJA_RUTA_ZONA_PUESTO + '".'
      );
    }
    var filaAuxiliar = buscarPuestoEnNomina_(puestoAuxiliar);
    if (!filaAuxiliar) {
      throw new Error('Puesto Auxiliar "' + puestoAuxiliar + '" no esta en la hoja "' + HOJA_NOMINA + '".');
    }
    if (filaAuxiliar[7] === '') {
      throw new Error('Completa Sueldo y Periodicidad de "' + puestoAuxiliar + '" en ' + HOJA_NOMINA + '.');
    }
    sueldoMensual += Number(filaAuxiliar[7]) || 0;
    puestoAuxiliarUsado = puestoAuxiliar;
  }

  var viajesMes = vecesPorSemana_(frecuencia) * 4.33;

  var rentaMensual = Number(filaUnidad[1]) || 0;
  var mantenimientoMensual = Number(filaUnidad[3]) || 0;
  var gasolinaKm = Number(filaUnidad[7]) || 0;
  var costoCasetas = Number(filaZona[1]) || 0;

  var r = calcularCostoRuta_(rentaMensual, mantenimientoMensual, gasolinaKm, sueldoMensual, km, viajesMes, costoCasetas, margen);

  var tarifaPorUnidad = '';
  if (tipoCobro !== 'Por Ruta') {
    var cantidadNum = Number(cantidad) || 0;
    if (cantidadNum <= 0) {
      throw new Error('Captura la Cantidad (paquetes/paradas/palets) para calcular la tarifa "' + tipoCobro + '".');
    }
    tarifaPorUnidad = r.tarifaPiso / cantidadNum;
  }

  return {
    puestoPrincipal: puestoPrincipal,
    puestoAuxiliar: puestoAuxiliarUsado,
    costoCasetas: costoCasetas,
    viajesMes: viajesMes,
    sueldoMensual: sueldoMensual,
    rentaMensual: rentaMensual,
    mantenimientoMensual: mantenimientoMensual,
    costoGasolinaKm: gasolinaKm,
    costoVariable: r.costoVariable,
    costoFijoProrrateado: r.costoFijoProrrateado,
    costoTotal: r.costoTotal,
    tarifaPiso: r.tarifaPiso,
    tarifaPorUnidad: tarifaPorUnidad
  };
}
