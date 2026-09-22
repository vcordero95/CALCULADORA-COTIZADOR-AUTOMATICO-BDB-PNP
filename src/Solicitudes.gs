/**
 * Busca el costo de casetas EXACTO entre dos puntos (Punto A y Punto B),
 * sin importar el orden (A->B cuesta lo mismo que B->A). Regresa null si
 * esa ruta no esta capturada en la hoja Casetas (no truena: el llamador
 * usa un estimado por km como respaldo).
 */
function buscarCasetasExacto_(puntoA, puntoB) {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CASETAS);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_CASETAS + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  var datos = hoja.getRange(2, 1, FILAS_CASETAS, 4).getValues();
  for (var i = 0; i < datos.length; i++) {
    var a = datos[i][0];
    var b = datos[i][1];
    if ((a === puntoA && b === puntoB) || (a === puntoB && b === puntoA)) {
      return datos[i];
    }
  }
  return null;
}

/** Costo de casetas estimado ($/km), tomado de Config, para cuando la ruta no esta en el catalogo. */
function obtenerCostoCasetasPorKm_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_CONFIG + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return Number(hoja.getRange(7, 2).getValue()) || 0;
}

/**
 * Resuelve el costo de casetas de una ruta. Si Tipo de Ruta es "Local"
 * (dentro de la misma ciudad) no lleva caseta. Si es Foraneo, Line Haul o
 * Media Milla (de Punto A a Punto B) si lleva: se busca el costo exacto
 * en la hoja Casetas, y si esa ruta especifica todavia no esta
 * capturada, se estima con kilometros x Costo Casetas Estimado ($/km) de
 * Config, para poder cotizar cualquier ruta del pais sin esperar a que el
 * catalogo este completo.
 */
function resolverCostoCasetas_(tipoRuta, puntoA, puntoB, km) {
  if (TIPOS_RUTA_CON_CASETA.indexOf(tipoRuta) === -1) {
    return { costo: 0, estimado: false, fuente: 'No aplica (Tipo de Ruta Local)' };
  }

  var filaExacta = (puntoA && puntoB) ? buscarCasetasExacto_(puntoA, puntoB) : null;
  if (filaExacta) {
    return { costo: Number(filaExacta[2]) || 0, estimado: false, fuente: filaExacta[3] || '' };
  }

  var costoPorKm = obtenerCostoCasetasPorKm_();
  var kilometros = Number(km) || 0;
  return {
    costo: costoPorKm * kilometros,
    estimado: true,
    fuente: 'Estimado a ' + costoPorKm + ' $/km (ruta no esta en el catalogo Casetas)'
  };
}

/**
 * Busca, para una combinacion de Tipo de Ruta + Destino (Punto B), que
 * Puesto principal y cual Auxiliar aplican. Regresa null si esa
 * combinacion no esta definida en la hoja Ruta-Zona-Puesto.
 */
function buscarRutaZonaPuesto_(tipoRuta, puntoB) {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_RUTA_ZONA_PUESTO);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_RUTA_ZONA_PUESTO + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  var datos = hoja.getRange(2, 1, FILAS_RUTA_ZONA, 4).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === tipoRuta && datos[i][1] === puntoB) return datos[i];
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

/** Margen Piso y Margen Objetivo, como fraccion, tomados de Config (politica fija de la empresa). */
function obtenerMargenesPolitica_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_CONFIG + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return {
    piso: Number(hoja.getRange(5, 2).getValue()) || 0,
    objetivo: Number(hoja.getRange(6, 2).getValue()) || 0
  };
}

/**
 * Calcula la tarifa de una solicitud de Comercial a partir de las
 * caracteristicas de la ruta, no de costos: Tipo de Ruta + Destino
 * (Punto B) resuelven el Puesto principal (y el Auxiliar, si la ruta lo
 * necesita) en la hoja Ruta-Zona-Puesto; la Frecuencia resuelve los
 * Viajes al Mes (veces por semana x 4.33).
 *
 * El Costo de Casetas depende del Tipo de Ruta: "Local" no lleva caseta;
 * "Foraneo", "Line Haul" y "Media Milla" (rutas de Punto A a Punto B) si
 * llevan, y se resuelve con Punto A + Punto B en la hoja Casetas — o, si
 * esa ruta especifica todavia no esta capturada, con un estimado por km
 * (ver resolverCostoCasetas_()) para poder cotizar cualquier ruta del
 * pais sin depender de tener el catalogo completo.
 *
 * El margen NO lo captura Comercial: se usa la politica fija de la
 * empresa (Margen Piso / Margen Objetivo en Config), y se regresan ambas
 * tarifas. Si tipoCobro no es "Por Ruta", ademas se regresa cada tarifa
 * entre la Cantidad (paquetes, paradas o palets) capturada.
 *
 * No es una funcion de hoja de calculo (@customfunction): la llama
 * calcularSolicitudWeb() desde el dashboard de Comercial.
 *
 * @return {Object} Desglose completo de la solicitud.
 */
function SOLICITAR_TARIFA(tipoRuta, puntoA, puntoB, tipoUnidad, frecuencia, km, tipoCobro, cantidad, requiereAuxiliar) {
  var filaUnidad = buscarUnidadPorTipo_(tipoUnidad);
  if (!filaUnidad) {
    throw new Error('Tipo de unidad "' + tipoUnidad + '" no esta en la hoja "' + HOJA_COSTOS_UNIDAD + '".');
  }
  if (filaUnidad[7] === '') {
    throw new Error('Completa Rendimiento y Tipo de Combustible de "' + tipoUnidad + '" en ' + HOJA_COSTOS_UNIDAD + '.');
  }

  var mapeo = buscarRutaZonaPuesto_(tipoRuta, puntoB);
  if (!mapeo) {
    throw new Error(
      'No hay un Puesto definido para Tipo de Ruta "' + tipoRuta + '" + Destino "' + puntoB +
      '" en la hoja "' + HOJA_RUTA_ZONA_PUESTO + '".'
    );
  }

  var puestoPrincipal = mapeo[2];
  var puestoAuxiliar = mapeo[3];

  var filaPrincipal = buscarPuestoEnNomina_(puestoPrincipal);
  if (!filaPrincipal) {
    throw new Error('Puesto "' + puestoPrincipal + '" (mapeado para ' + tipoRuta + ' + ' + puntoB + ') no esta en la hoja "' + HOJA_NOMINA + '".');
  }
  if (filaPrincipal[7] === '') {
    throw new Error('Completa Sueldo y Periodicidad de "' + puestoPrincipal + '" en ' + HOJA_NOMINA + '.');
  }

  var sueldoMensual = Number(filaPrincipal[7]) || 0;
  var puestoAuxiliarUsado = '';

  if (requiereAuxiliar) {
    if (!puestoAuxiliar) {
      throw new Error(
        'No hay Puesto Auxiliar definido para Tipo de Ruta "' + tipoRuta + '" + Destino "' + puntoB +
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
  var infoCasetas = resolverCostoCasetas_(tipoRuta, puntoA, puntoB, km);
  var costoCasetas = infoCasetas.costo;

  var r = calcularCostoRuta_(rentaMensual, mantenimientoMensual, gasolinaKm, sueldoMensual, km, viajesMes, costoCasetas);

  var margenes = obtenerMargenesPolitica_();
  var tarifaPiso = aplicarMargen_(r.costoTotal, margenes.piso);
  var tarifaObjetivo = aplicarMargen_(r.costoTotal, margenes.objetivo);

  var tarifaPorUnidadPiso = '';
  var tarifaPorUnidadObjetivo = '';
  if (tipoCobro !== 'Por Ruta') {
    var cantidadNum = Number(cantidad) || 0;
    if (cantidadNum <= 0) {
      throw new Error('Captura la Cantidad (paquetes/paradas/palets) para calcular la tarifa "' + tipoCobro + '".');
    }
    tarifaPorUnidadPiso = tarifaPiso / cantidadNum;
    tarifaPorUnidadObjetivo = tarifaObjetivo / cantidadNum;
  }

  return {
    puestoPrincipal: puestoPrincipal,
    puestoAuxiliar: puestoAuxiliarUsado,
    costoCasetas: costoCasetas,
    casetasEstimadas: infoCasetas.estimado,
    casetasFuente: infoCasetas.fuente,
    viajesMes: viajesMes,
    sueldoMensual: sueldoMensual,
    rentaMensual: rentaMensual,
    mantenimientoMensual: mantenimientoMensual,
    costoGasolinaKm: gasolinaKm,
    costoVariable: r.costoVariable,
    costoFijoProrrateado: r.costoFijoProrrateado,
    costoTotal: r.costoTotal,
    margenPiso: margenes.piso,
    margenObjetivo: margenes.objetivo,
    tarifaPiso: tarifaPiso,
    tarifaObjetivo: tarifaObjetivo,
    tarifaPorUnidadPiso: tarifaPorUnidadPiso,
    tarifaPorUnidadObjetivo: tarifaPorUnidadObjetivo
  };
}
