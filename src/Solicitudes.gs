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
 * Quita acentos y puntos de un nombre de lugar antes de mandarlo a
 * INEGI: su buscador de destinos no reconoce caracteres acentuados
 * ("Querétaro" no encuentra nada, pero "Queretaro" si) ni abreviaturas
 * con punto ("N.L." no encuentra nada, pero "NL" si) — confirmado
 * probando varias ciudades y estados en vivo contra su API. Sin esto,
 * cualquier Punto A/B con acento (la mayoria de los nombres en Mexico)
 * fallaba siempre y se caia al estimado por km.
 */
function normalizarParaInegi_(texto) {
  return String(texto || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\./g, '')
    .trim();
}

/**
 * Corre intentoFn() hasta "intentos" veces, esperando un poco mas entre
 * cada una, hasta que regrese algo distinto de null/undefined. Las
 * llamadas de Apps Script a INEGI fallan de forma intermitente — a
 * veces con un HTTP distinto de 200, pero tambien a veces con un 200
 * valido que trae el resultado vacio (success:false o data:[]) — asi
 * que el reintento tiene que cubrir AMBOS casos, no solo el status
 * HTTP, o se queda sin reintentar justo en el caso mas comun.
 */
function conReintentos_(intentoFn, intentos) {
  var maxIntentos = intentos || 3;
  for (var intento = 1; intento <= maxIntentos; intento++) {
    var resultado = intentoFn();
    if (resultado != null) return resultado;
    if (intento < maxIntentos) Utilities.sleep(400 * intento);
  }
  return null;
}

/**
 * Busca el id_dest de un lugar en la API de Ruteo de INEGI (funcion
 * buscadestino). Regresa null si no hay match o si la llamada falla
 * tras varios intentos, para que el llamador siga con el estimado por
 * km sin romper la cotizacion.
 */
function buscarIdDestinoInegi_(nombreLugar, token) {
  var nombreNormalizado = normalizarParaInegi_(nombreLugar);
  var payload = { buscar: nombreNormalizado, type: 'json', num: '1', key: token };

  return conReintentos_(function () {
    try {
      var respuesta = UrlFetchApp.fetch('https://gaia.inegi.org.mx/sakbe_v3.1/buscadestino', {
        method: 'post', payload: payload, muteHttpExceptions: true
      });
      if (respuesta.getResponseCode() !== 200) {
        console.error('INEGI buscadestino "' + nombreLugar + '" HTTP ' + respuesta.getResponseCode() + ': ' + respuesta.getContentText());
        return null;
      }
      var datos = JSON.parse(respuesta.getContentText());
      if (!datos.response || !datos.response.success || !datos.data || !datos.data.length) {
        console.error('INEGI buscadestino "' + nombreLugar + '" sin match: ' + respuesta.getContentText());
        return null;
      }
      return datos.data[0].id_dest;
    } catch (e) {
      console.error('INEGI buscadestino "' + nombreLugar + '" excepcion: ' + e.message);
      return null;
    }
  });
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
  if (!token) {
    console.error('INEGI: falta el Token INEGI en Config!B8.');
    return null;
  }
  if (!puntoA || !puntoB) return null;

  var idA = buscarIdDestinoInegi_(puntoA, token);
  var idB = buscarIdDestinoInegi_(puntoB, token);
  if (!idA || !idB) return null;

  var payload = { dest_i: idA, dest_f: idB, v: codigoVehiculo, type: 'json', key: token };

  return conReintentos_(function () {
    try {
      var respuesta = UrlFetchApp.fetch('https://gaia.inegi.org.mx/sakbe_v3.1/cuota', {
        method: 'post', payload: payload, muteHttpExceptions: true
      });
      if (respuesta.getResponseCode() !== 200) {
        console.error('INEGI cuota ' + puntoA + ' -> ' + puntoB + ' HTTP ' + respuesta.getResponseCode() + ': ' + respuesta.getContentText());
        return null;
      }

      var texto = respuesta.getContentText();
      try {
        var datos = JSON.parse(texto);
        if (!datos.response || !datos.response.success || !datos.data) {
          console.error('INEGI cuota ' + puntoA + ' -> ' + puntoB + ' fallo: ' + texto);
          return null;
        }
        return { costo: Number(datos.data.costo_caseta) || 0, km: Number(datos.data.long_km) || 0 };
      } catch (errorParse) {
        // Rutas muy largas regresan un geojson enorme; si el parse de
        // todo el documento falla (respuesta cortada o con algun
        // caracter que rompe el parser estricto), se rescatan los 2
        // numeros que realmente se necesitan directo del texto, sin
        // depender de que el documento completo sea JSON valido.
        var costoMatch = texto.match(/"costo_caseta":(-?[\d.]+)/);
        var kmMatch = texto.match(/"long_km":(-?[\d.]+)/);
        if (costoMatch && kmMatch) {
          console.error('INEGI cuota ' + puntoA + ' -> ' + puntoB + ' JSON invalido (' + errorParse.message + '), rescatado por texto. Longitud respuesta: ' + texto.length);
          return { costo: Number(costoMatch[1]) || 0, km: Number(kmMatch[1]) || 0 };
        }
        console.error('INEGI cuota ' + puntoA + ' -> ' + puntoB + ' JSON invalido y no se pudo rescatar: ' + errorParse.message + '. Longitud respuesta: ' + texto.length);
        return null;
      }
    } catch (e) {
      console.error('INEGI cuota ' + puntoA + ' -> ' + puntoB + ' excepcion: ' + e.message);
      return null;
    }
  });
}

/** Busca un tipo de combustible (ej. "Diésel", "Regular") en la lista que regresa la API de INEGI. Regresa null si no esta. */
function buscarCombustibleInegiPorTipo_(lista, tiposAceptados) {
  for (var i = 0; i < lista.length; i++) {
    if (tiposAceptados.indexOf(lista[i].tipo) !== -1) {
      return Number(lista[i].costo) || null;
    }
  }
  return null;
}

/**
 * Jala de la API de INEGI (funcion "combustible") el precio promedio
 * nacional de Diesel y Gasolina (que INEGI actualiza el primer dia
 * habil de cada semana) y lo escribe en Config!B1 / Config!B2. No
 * truena si INEGI falla (token invalido, sin red, etc.): en ese caso
 * se queda el ultimo precio capturado. Pensada para correr sola via un
 * trigger semanal (ver activarActualizacionCombustibleInegi_).
 */
function actualizarPreciosCombustibleInegi_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  if (!hoja) return;
  var token = obtenerTokenInegi_();
  if (!token) return;

  try {
    var respuesta = UrlFetchApp.fetch('https://gaia.inegi.org.mx/sakbe_v3.1/combustible', {
      method: 'post',
      payload: { type: 'json', key: token },
      muteHttpExceptions: true
    });
    if (respuesta.getResponseCode() !== 200) return;
    var datos = JSON.parse(respuesta.getContentText());
    if (!datos.response || !datos.response.success || !datos.data) return;

    var diesel = buscarCombustibleInegiPorTipo_(datos.data, ['Diésel', 'Diesel']);
    var gasolina = buscarCombustibleInegiPorTipo_(datos.data, ['Regular', 'Magna']);

    if (diesel != null) hoja.getRange(1, 2).setValue(diesel);
    if (gasolina != null) hoja.getRange(2, 2).setValue(gasolina);
  } catch (e) {
    // No interrumpir: si INEGI falla, se queda con el ultimo precio capturado.
  }
}

/**
 * Activa (una sola vez) el disparador semanal que mantiene actualizados
 * los precios de Diesel/Gasolina en Config con el dato de INEGI. Si ya
 * esta activo, no lo duplica. Se corre tambien de inmediato al activar,
 * para no esperar hasta el proximo lunes.
 */
function activarActualizacionCombustibleInegi_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'actualizarPreciosCombustibleInegi_') {
      SpreadsheetApp.getUi().alert('La actualizacion automatica semanal de combustibles ya esta activa.');
      return;
    }
  }

  ScriptApp.newTrigger('actualizarPreciosCombustibleInegi_')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  actualizarPreciosCombustibleInegi_();

  SpreadsheetApp.getUi().alert(
    'Listo. Los precios de Diesel y Gasolina en Config se van a actualizar solos cada lunes con el dato de INEGI ' +
    '(ya se actualizaron ahora mismo con el precio mas reciente).'
  );
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
    return { costo: 0, estimado: false, fuente: 'No aplica para Modalidad "' + modalidad + '"', km: null };
  }

  var resultadoInegi = (puntoA && puntoB) ? consultarCostoCasetaInegi_(puntoA, puntoB, codigoVehiculo) : null;
  if (resultadoInegi) {
    return {
      costo: resultadoInegi.costo,
      estimado: true,
      fuente: 'API de Ruteo de INEGI (SAKBE), en vivo',
      km: resultadoInegi.km
    };
  }

  var filaExacta = (puntoA && puntoB) ? buscarCasetasExacto_(puntoA, puntoB) : null;
  if (filaExacta) {
    return { costo: Number(filaExacta[2]) || 0, estimado: false, fuente: (filaExacta[3] || '') + ' (INEGI no respondio, se uso el catalogo Casetas)', km: null };
  }

  var costoPorKm = obtenerCostoCasetasPorKm_();
  var kilometros = Number(km) || 0;
  return {
    costo: costoPorKm * kilometros,
    estimado: true,
    fuente: 'Estimado a ' + costoPorKm + ' $/km (INEGI no respondio y la ruta no esta en el catalogo Casetas)',
    km: null
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
 * entre la Cantidad (paquetes, paradas o palets) capturada, convertida a
 * cantidad por viaje: si Comercial la captura "Por Semana" se divide
 * entre los viajes por semana de la Frecuencia; si la captura "Por Dia"
 * se usa tal cual (un viaje = un dia, igual que asume la Frecuencia).
 *
 * No es una funcion de hoja de calculo (@customfunction): la llama
 * calcularSolicitudWeb() desde el dashboard de Comercial.
 *
 * @return {Object} Desglose completo de la solicitud.
 */
function SOLICITAR_TARIFA(alcance, modalidad, puntoA, puntoB, tipoUnidad, frecuencia, km, tipoCobro, cantidad, periodoCantidad, requiereAuxiliar) {
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

  var vecesSemana = vecesPorSemana_(frecuencia);
  var viajesMes = vecesSemana * 4.33;

  var rentaMensual = Number(filaUnidad[1]) || 0;
  var mantenimientoMensual = Number(filaUnidad[3]) || 0;
  var gasolinaKm = Number(filaUnidad[7]) || 0;
  var codigoVehiculo = filaUnidad[8];
  var infoCasetas = resolverCostoCasetas_(modalidad, puntoA, puntoB, km, codigoVehiculo);
  var costoCasetas = infoCasetas.costo;

  // Si la API de Ruteo de INEGI pudo resolver la ruta, su kilometraje es
  // mas confiable que lo que Comercial haya tecleado a mano (esa es la
  // ruta real entre Punto A y Punto B, no una estimacion); si no, se usa
  // el km capturado.
  var kmUsado = infoCasetas.km != null ? infoCasetas.km : (Number(km) || 0);
  var kmFuente = infoCasetas.km != null ? 'Calculado por la API de Ruteo de INEGI' : 'Capturado por Comercial';

  if (kmUsado <= 0) {
    throw new Error(
      'No se pudieron calcular los kilometros: la API de Ruteo de INEGI no identifico "' + puntoA + '" / "' + puntoB +
      '" (intenta con el formato "Ciudad, Estado"), y tampoco se capturo un kilometraje manual. Captura los Kilometros a mano para continuar.'
    );
  }

  var r = calcularCostoRuta_(rentaMensual, mantenimientoMensual, gasolinaKm, sueldoMensual, kmUsado, viajesMes, costoCasetas);

  var margenes = obtenerMargenesPolitica_();
  var tarifaPiso = aplicarMargen_(r.costoTotal, margenes.piso);
  var tarifaObjetivo = aplicarMargen_(r.costoTotal, margenes.objetivo);

  var tarifaPorUnidadPiso = '';
  var tarifaPorUnidadObjetivo = '';
  if (tipoCobro !== 'Por Ruta') {
    var cantidadNum = Number(cantidad) || 0;
    if (cantidadNum <= 0) {
      throw new Error('Captura la Cantidad para calcular la tarifa "' + tipoCobro + '".');
    }
    var cantidadPorViaje = periodoCantidad === 'Por Semana' ? (cantidadNum / vecesSemana) : cantidadNum;
    tarifaPorUnidadPiso = tarifaPiso / cantidadPorViaje;
    tarifaPorUnidadObjetivo = tarifaObjetivo / cantidadPorViaje;
  }

  return {
    puestoPrincipal: puestoPrincipal,
    puestoAuxiliar: puestoAuxiliarUsado,
    costoCasetas: costoCasetas,
    casetasEstimadas: infoCasetas.estimado,
    casetasFuente: infoCasetas.fuente,
    km: kmUsado,
    kmFuente: kmFuente,
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

/** "$#,##0.00" para texto plano (correos), sin depender de formato de celda. */
function formatoMoneda_(valor) {
  var n = Number(valor) || 0;
  return '$' + n.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Avisa a Vanessa por correo cuando Comercial guarda una solicitud
 * nueva, con los datos clave para revisarla. Nunca interrumpe el
 * guardado si el correo falla (cuota de MailApp agotada, etc.).
 */
function enviarCorreoNuevaSolicitud_(fila, datos, resultado, correo, clienteRegistrado) {
  try {
    var esBusqueda = !!resultado.esBusquedaTarifa;
    var asunto = 'Nueva solicitud de tarifa — ' + clienteRegistrado + ' (fila ' + fila + ')';
    var cuerpo = 'Se guardo una solicitud nueva en el Cotizador BDB.\n\n' +
      'Solicitada por: ' + correo + '\n' +
      'Cliente: ' + clienteRegistrado + '\n' +
      'Referencia de ruta: ' + (datos.rutaCliente || '-') + '\n' +
      'Alcance / Modalidad: ' + (datos.alcance || '-') + ' / ' + (datos.modalidad || '-') + '\n' +
      (datos.puntoA ? ('Punto A / Punto B: ' + datos.puntoA + ' -> ' + datos.puntoB + '\n') : '') +
      'Tipo de Unidad: ' + (datos.tipoUnidad || '-') + '\n\n' +
      (esBusqueda
        ? 'Es busqueda de tarifa estandarizada (Service Partner/XPT), no calculo de costo.\n' +
          (resultado.tarifaMeli ? ('Tarifa vigente MELI (' + resultado.tarifaMeli.nivel + '): ' + formatoMoneda_(resultado.tarifaMeli.tarifa) + '\n') : '')
        : 'Tarifa Piso: ' + formatoMoneda_(resultado.tarifaPiso) + '\n' +
          'Tarifa Objetivo: ' + formatoMoneda_(resultado.tarifaObjetivo) + '\n') +
      '\nRevisala y cambia su Estatus (Pendiente/Aprobada/Rechazada) en la hoja Solicitudes, fila ' + fila + ', ' +
      'para que el sistema le avise el resultado a quien la pidio.';

    MailApp.sendEmail(CORREO_VANESSA, asunto, cuerpo);
  } catch (e) {
    // No interrumpir el guardado si el correo no se pudo enviar.
  }
}

/**
 * Disparador instalable (se activa desde el menu con
 * activarNotificacionesAprobacion_ en Code.gs): se dispara con
 * cualquier edicion de cualquier hoja, pero solo actua si fue la
 * columna Estatus de Solicitudes y el valor nuevo es "Aprobada" o
 * "Rechazada" — ahi le avisa por correo a quien pidio esa solicitud.
 * Nunca interrumpe la edicion si el correo falla.
 */
function onEditAprobacionSolicitud_(e) {
  try {
    var range = e.range;
    var hoja = range.getSheet();
    if (hoja.getName() !== HOJA_SOLICITUDES) return;
    if (range.getColumn() !== COLUMNA_ESTATUS_SOLICITUDES) return;
    var fila = range.getRow();
    if (fila < 2) return;

    var estatus = String(range.getValue() || '').trim();
    if (estatus !== 'Aprobada' && estatus !== 'Rechazada') return;

    enviarCorreoResultadoSolicitud_(hoja, fila, estatus);
  } catch (err) {
    // No interrumpir la edicion si falla el envio del correo.
  }
}

/** Correo a quien pidio la solicitud (identificado por su Correo Solicitante), avisando si fue Aprobada o Rechazada. */
function enviarCorreoResultadoSolicitud_(hoja, fila, estatus) {
  var datos = hoja.getRange(fila, 1, 1, COLUMNA_ESTATUS_SOLICITUDES).getValues()[0];
  var correo = String(datos[COLUMNA_CORREO_SOLICITANTE - 1] || '').trim();
  if (!correo) return;

  var cliente = datos[1], referencia = datos[2], alcance = datos[3], modalidad = datos[4];
  var tarifaPiso = datos[29], tarifaObjetivo = datos[30];

  var asunto = 'Tu solicitud de tarifa (' + cliente + ') fue ' + estatus.toLowerCase();
  var cuerpo = 'Hola,\n\n' +
    'Tu solicitud de tarifa para "' + cliente + '"' + (referencia ? ' (' + referencia + ')' : '') +
    ' [' + alcance + ' / ' + modalidad + '] fue ' + estatus.toUpperCase() + '.\n\n' +
    (estatus === 'Aprobada'
      ? 'Tarifa piso: ' + formatoMoneda_(tarifaPiso) + '\nTarifa objetivo: ' + formatoMoneda_(tarifaObjetivo) + '\n\n'
      : '') +
    'Este correo se genero automaticamente desde el Cotizador BDB.';

  try {
    MailApp.sendEmail(correo, asunto, cuerpo);
  } catch (e) {
    // No interrumpir si el correo no se puede enviar (direccion invalida, cuota, etc.).
  }
}

