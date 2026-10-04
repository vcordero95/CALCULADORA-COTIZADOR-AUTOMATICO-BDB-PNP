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

/**
 * Agrega una ruta nueva a la hoja Casetas con el costo que se estimo por
 * km, para que la proxima vez que se pida esa misma ruta ya no sea un
 * estimado. Se llama solo al GUARDAR una solicitud (no en cada calculo),
 * para no llenar el catalogo con rutas que Comercial solo esta probando.
 * Si la ruta ya existe (alguien mas la guardo mientras tanto) no duplica.
 */
function agregarCasetaAprendida_(puntoA, puntoB, costo) {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CASETAS);
  if (!hoja) return;
  if (buscarCasetasExacto_(puntoA, puntoB)) return;

  var columnaA = hoja.getRange(2, 1, FILAS_CASETAS, 1).getValues();
  for (var i = 0; i < columnaA.length; i++) {
    if (columnaA[i][0] === '') {
      var fila = 2 + i;
      hoja.getRange(fila, 1, 1, 4).setValues([[
        puntoA, puntoB, costo,
        'Estimado automaticamente desde una solicitud guardada el ' +
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') +
          ' — validar/ajustar con el costo real de caseta.'
      ]]);
      return;
    }
  }
  // Catalogo lleno (sin filas libres hasta FILAS_CASETAS): no se agrega, se sigue estimando por km.
}

/** Costo de casetas estimado ($/km), tomado de Config, para cuando la ruta no esta en el catalogo ni la pudo resolver INEGI. */
function obtenerCostoCasetasPorKm_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_CONFIG + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return Number(hoja.getRange(7, 2).getValue()) || 0;
}

/** Token de la API de Ruteo de INEGI (SAKBE), tomado de Config. */
function obtenerTokenInegi_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_CONFIG + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  return String(hoja.getRange(8, 2).getValue() || '').trim();
}

/**
 * Busca el id_dest de un lugar en la API de Ruteo de INEGI (funcion
 * buscadestino). Regresa null si no hay match o si la llamada falla,
 * para que el llamador siga con el estimado por km sin romper la
 * cotizacion.
 */
function buscarIdDestinoInegi_(nombreLugar, token) {
  try {
    var respuesta = UrlFetchApp.fetch('https://gaia.inegi.org.mx/sakbe_v3.1/buscadestino', {
      method: 'post',
      payload: { buscar: nombreLugar, type: 'json', num: 1, key: token },
      muteHttpExceptions: true
    });
    if (respuesta.getResponseCode() !== 200) return null;
    var datos = JSON.parse(respuesta.getContentText());
    if (!datos.response || !datos.response.success || !datos.data || !datos.data.length) return null;
    return datos.data[0].id_dest;
  } catch (e) {
    return null;
  }
}

/**
 * Consulta en vivo el costo de caseta de una ruta Punto A -> Punto B en
 * la API de Ruteo de INEGI (SAKBE), usando el Codigo de Vehiculo INEGI
 * del Tipo de Unidad elegido (varia el costo real: no es lo mismo Auto
 * que Camion). Regresa null ante cualquier falla (lugar no encontrado,
 * error de red, respuesta invalida), nunca truena: el llamador cae al
 * estimado por km.
 */
function consultarCostoCasetaInegi_(puntoA, puntoB, codigoVehiculo) {
  var token = obtenerTokenInegi_();
  if (!token || !puntoA || !puntoB) return null;

  var idA = buscarIdDestinoInegi_(puntoA, token);
  var idB = buscarIdDestinoInegi_(puntoB, token);
  if (!idA || !idB) return null;

  try {
    var respuesta = UrlFetchApp.fetch('https://gaia.inegi.org.mx/sakbe_v3.1/cuota', {
      method: 'post',
      payload: { dest_i: idA, dest_f: idB, v: codigoVehiculo, type: 'json', key: token },
      muteHttpExceptions: true
    });
    if (respuesta.getResponseCode() !== 200) return null;
    var datos = JSON.parse(respuesta.getContentText());
    if (!datos.response || !datos.response.success || !datos.data) return null;
    return { costo: Number(datos.data.costo_caseta) || 0, km: Number(datos.data.long_km) || 0 };
  } catch (e) {
    return null;
  }
}

/**
 * Resuelve el costo de casetas de una ruta segun la Modalidad (no el
 * Alcance): "Line Haul" y "Media Milla" son de Punto A a Punto B y si
 * llevan caseta; el resto no. Cuando aplica, se consulta primero en
 * vivo la API de Ruteo de INEGI (variando segun el Codigo de Vehiculo
 * INEGI del Tipo de Unidad elegido), para siempre cotizar con el dato
 * mas actualizado posible; si INEGI no puede resolverla (lugar no
 * encontrado, error de red), se cae al costo exacto capturado en la
 * hoja Casetas (CAPUFE u otra fuente oficial); y solo si tampoco esta
 * ahi, se estima con kilometros x Costo Casetas Estimado ($/km) de
 * Config, para poder cotizar cualquier ruta del pais sin bloquear nunca
 * la cotizacion.
 */
function resolverCostoCasetas_(modalidad, puntoA, puntoB, km, codigoVehiculo) {
  if (MODALIDADES_CON_CASETA.indexOf(modalidad) === -1) {
    return { costo: 0, estimado: false, fuente: 'No aplica para Modalidad "' + modalidad + '"' };
  }

  var resultadoInegi = (puntoA && puntoB) ? consultarCostoCasetaInegi_(puntoA, puntoB, codigoVehiculo) : null;
  if (resultadoInegi) {
    return {
      costo: resultadoInegi.costo,
      estimado: true,
      fuente: 'API de Ruteo de INEGI (SAKBE), en vivo — ' + resultadoInegi.km + ' km'
    };
  }

  var filaExacta = (puntoA && puntoB) ? buscarCasetasExacto_(puntoA, puntoB) : null;
  if (filaExacta) {
    return { costo: Number(filaExacta[2]) || 0, estimado: false, fuente: (filaExacta[3] || '') + ' (INEGI no respondio, se uso el catalogo Casetas)' };
  }

  var costoPorKm = obtenerCostoCasetasPorKm_();
  var kilometros = Number(km) || 0;
  return {
    costo: costoPorKm * kilometros,
    estimado: true,
    fuente: 'Estimado a ' + costoPorKm + ' $/km (INEGI no respondio y la ruta no esta en el catalogo Casetas)'
  };
}

/**
 * Busca el Puesto Principal / Auxiliar para un Alcance (Local/Foraneo).
 * Es independiente de la Modalidad: cualquier Modalidad usa el mismo
 * Puesto segun su Alcance. Regresa null si ese Alcance no esta definido
 * en la hoja Alcance-Puesto.
 */
function buscarPuestoPorAlcance_(alcance) {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_ALCANCE_PUESTO);
  if (!hoja) {
    throw new Error('No existe la hoja "' + HOJA_ALCANCE_PUESTO + '". Ejecuta Cotizador BDB > Inicializar hojas.');
  }
  var datos = hoja.getRange(2, 1, FILAS_ALCANCE, 3).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === alcance) return datos[i];
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
 * caracteristicas de la ruta, no de costos. Se usa para todas las
 * Modalidades en modo "Tarifa nueva" excepto XPT, cuyo calculo siempre
 * es la tabla estandarizada por Nivel de MELI (ver el manejo de XPT en
 * WebApp.gs / calcularSolicitudWeb).
 *
 * Alcance (Local/Foraneo) resuelve el Puesto principal (y el Auxiliar,
 * si la ruta lo necesita) en la hoja Alcance-Puesto, sin importar la
 * Modalidad. La Frecuencia resuelve los Viajes al Mes (veces por semana
 * x 4.33).
 *
 * El Costo de Casetas depende de la Modalidad (ver resolverCostoCasetas_):
 * solo Line Haul y Media Milla llevan, resuelto con Punto A + Punto B en
 * la hoja Casetas o, si esa ruta especifica todavia no esta capturada,
 * con un estimado por km.
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
function SOLICITAR_TARIFA(alcance, modalidad, puntoA, puntoB, tipoUnidad, frecuencia, km, tipoCobro, cantidad, requiereAuxiliar) {
  var filaUnidad = buscarUnidadPorTipo_(tipoUnidad);
  if (!filaUnidad) {
    throw new Error('Tipo de unidad "' + tipoUnidad + '" no esta en la hoja "' + HOJA_COSTOS_UNIDAD + '".');
  }
  if (filaUnidad[7] === '') {
    throw new Error('Completa Rendimiento y Tipo de Combustible de "' + tipoUnidad + '" en ' + HOJA_COSTOS_UNIDAD + '.');
  }

  var mapeo = buscarPuestoPorAlcance_(alcance);
  if (!mapeo) {
    throw new Error(
      'No hay un Puesto definido para Alcance "' + alcance + '" en la hoja "' + HOJA_ALCANCE_PUESTO + '".'
    );
  }

  var puestoPrincipal = mapeo[1];
  var puestoAuxiliar = mapeo[2];

  var filaPrincipal = buscarPuestoEnNomina_(puestoPrincipal);
  if (!filaPrincipal) {
    throw new Error('Puesto "' + puestoPrincipal + '" (mapeado para Alcance ' + alcance + ') no esta en la hoja "' + HOJA_NOMINA + '".');
  }
  if (filaPrincipal[7] === '') {
    throw new Error('Completa Sueldo y Periodicidad de "' + puestoPrincipal + '" en ' + HOJA_NOMINA + '.');
  }

  var sueldoMensual = Number(filaPrincipal[7]) || 0;
  var puestoAuxiliarUsado = '';

  if (requiereAuxiliar) {
    if (!puestoAuxiliar) {
      throw new Error('No hay Puesto Auxiliar definido para Alcance "' + alcance + '" en la hoja "' + HOJA_ALCANCE_PUESTO + '".');
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
  var codigoVehiculo = filaUnidad[8];
  var infoCasetas = resolverCostoCasetas_(modalidad, puntoA, puntoB, km, codigoVehiculo);
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

