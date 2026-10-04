/**
 * Sirve el dashboard como Web App independiente. Acceso restringido al
 * dominio de BDB (ver src/appsscript.json: webapp.access). Por default
 * sirve el dashboard de Comercial (solicitud de tarifa de cliente);
 * agregando ?vista=interna a la URL se sirve el dashboard de costos de
 * uso interno.
 */
function doGet(e) {
  var vista = e && e.parameter && e.parameter.vista;
  var archivo = vista === 'interna' ? 'Dashboard' : 'DashboardComercial';
  return HtmlService.createHtmlOutputFromFile(archivo)
    .setTitle('Cotizador BDB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Catalogos para el dashboard interno (costos): Tipo de Unidad y Puesto /
 * Categoria de Sueldo, cada uno de su propio catalogo.
 */
function obtenerCatalogos() {
  var soloLlenos = function (valor) { return valor !== ''; };

  var unidades = obtenerHojaCostosUnidad_().getRange(2, 1, FILAS_CATALOGO, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  var puestos = obtenerHojaNomina_().getRange(2, 1, FILAS_CATALOGO, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  return { unidades: unidades, puestos: puestos };
}

/**
 * Catalogos para el dashboard de Comercial: Puntos (origen/destino) y
 * Tipo de Unidad salen de sus catalogos; Alcance, Modalidad, Frecuencia y
 * Tipo de Cobro son categorias fijas del negocio. Cliente sale de
 * Tarifas Vigentes, mas la opcion fija "Nuevo Cliente" (para clientes que
 * aun no tienen tarifa registrada). Estaciones MELI solo aplica cuando
 * Cliente = Mercado Libre y Modalidad = XPT.
 */
function obtenerCatalogosComercial() {
  var soloLlenos = function (valor) { return valor !== ''; };

  var puntos = obtenerHojaPuntos_().getRange(2, 1, FILAS_PUNTOS, 1).getValues()
    .map(function (fila) { return fila[0]; }).filter(soloLlenos);

  var unidades = obtenerHojaCostosUnidad_().getRange(2, 1, FILAS_CATALOGO, 1)
    .getValues().map(function (fila) { return fila[0]; }).filter(soloLlenos);

  var clientesUnicos = {};
  _TARIFAS_VIGENTES_DATA.forEach(function (fila) { clientesUnicos[fila[1]] = true; });
  var clientes = Object.keys(clientesUnicos).sort();
  clientes.push('Nuevo Cliente');

  var estacionesMeli = _MELI_ESTACIONES_DATA.map(function (fila) { return fila[0]; });

  return {
    puntos: puntos,
    unidades: unidades,
    alcances: ALCANCES,
    modalidades: MODALIDADES,
    modalidadesConPuntoAB: MODALIDADES_CON_PUNTO_A_B,
    frecuencias: FRECUENCIAS,
    tiposCobro: TIPOS_COBRO,
    periodosCantidad: PERIODOS_CANTIDAD,
    clientes: clientes,
    estacionesMeli: estacionesMeli
  };
}

/**
 * Calcula una cotizacion desde el dashboard interno, reusando COTIZAR()
 * (la misma formula que usa la hoja Cotizador). El margen llega del
 * formulario en porcentaje (20) y aqui se convierte a fraccion (0.20).
 */
function calcularCotizacionWeb(datos) {
  var margenFraccion = Number(datos.margen) / 100;
  var fila = COTIZAR(datos.tipoUnidad, datos.puesto, datos.km, datos.viajesMes, datos.casetas, margenFraccion)[0];

  return {
    sueldoMensual: fila[0],
    costoGasolinaKm: fila[1],
    rentaMensual: fila[2],
    mantenimientoMensual: fila[3],
    costoVariable: fila[4],
    costoFijoProrrateado: fila[5],
    costoTotal: fila[6],
    tarifaPiso: fila[7]
  };
}

/**
 * Guarda la cotizacion calculada (dashboard interno) como un renglon
 * nuevo en la hoja Cotizador, con los valores ya resueltos (no formulas).
 */
function guardarCotizacionWeb(datos, resultado) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(COTIZADOR);
  if (!sheet) {
    throw new Error('No existe la hoja "Cotizador". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  var fila = siguienteFilaLibreCotizador_(sheet);
  var margenFraccion = Number(datos.margen) / 100;

  sheet.getRange(fila, 1, 1, 17).setValues([[
    new Date(), datos.rutaCliente || '', datos.tipoUnidad, datos.puesto,
    Number(datos.km) || 0, Number(datos.paradas) || 0, Number(datos.viajesMes) || 0,
    Number(datos.casetas) || 0, margenFraccion,
    resultado.sueldoMensual, resultado.costoGasolinaKm, resultado.rentaMensual, resultado.mantenimientoMensual,
    resultado.costoVariable, resultado.costoFijoProrrateado, resultado.costoTotal, resultado.tarifaPiso
  ]]);

  return true;
}

/**
 * Primera fila libre de la hoja Cotizador, segun la columna Tipo de
 * Unidad (que nunca viene pre-llenada, a diferencia del margen por
 * defecto que si se pre-llena en las 500 filas reservadas).
 */
function siguienteFilaLibreCotizador_(sheet) {
  var columna = sheet.getRange(2, 3, FILAS_COTIZADOR, 1).getValues();
  for (var i = 0; i < columna.length; i++) {
    if (columna[i][0] === '') return 2 + i;
  }
  return 2 + FILAS_COTIZADOR;
}

/**
 * Calcula una solicitud de tarifa de cliente (modo "Tarifa nueva" del
 * dashboard de Comercial: se calcula desde costo, no se busca). La unica
 * excepcion es "XPT" (solo Mercado Libre): su calculo siempre es la
 * misma tabla estandarizada por Nivel (Estacion -> Nivel -> MELI
 * Tarifas), sea la ruta nueva o repetida, asi que no pasa por Costos
 * Unidad + Nomina + Casetas como las demas.
 *
 * Cualquier otra Modalidad (Dedicada, Spot, Service Partner, Line Haul,
 * Media Milla) si pasa por el motor de costos completo, reusando
 * SOLICITAR_TARIFA() — incluida Service Partner, para poder proponerle
 * una tarifa a un cliente que todavia no tiene una negociada. Comercial
 * no captura margen: se usa la politica fija de la empresa (Margen Piso
 * / Margen Objetivo en Config) y se regresan ambas tarifas.
 *
 * En todos los casos se agregan, como referencia (no como el resultado
 * en si), las tarifas vigentes del Cliente elegido.
 *
 * Atrapa cualquier error y lo regresa como { error: mensaje } en vez de
 * dejarlo "throw": Safari tiene un bug conocido con google.script.run
 * donde un error lanzado del lado del servidor no siempre llega al
 * withFailureHandler() del dashboard, y el boton se queda en
 * "Calculando..." sin mostrar nada. Regresando el error como dato
 * normal, el dashboard lo puede mostrar sin depender de ese mecanismo.
 */
function calcularSolicitudWeb(datos) {
  try {
    if (datos.modalidad === 'XPT') {
      if (datos.cliente !== 'Mercado Libre') {
        throw new Error('XPT solo aplica para Mercado Libre.');
      }
      if (!datos.estacionMeli) {
        throw new Error('Selecciona la Estacion MELI para buscar la tarifa XPT.');
      }
      return {
        esBusquedaTarifa: true,
        referenciasVigentes: [],
        tarifaMeli: buscarTarifaMeliDedicada_(datos.estacionMeli, datos.tipoUnidad, datos.km)
      };
    }

    var resultado = SOLICITAR_TARIFA(
      datos.alcance, datos.modalidad, datos.puntoA, datos.puntoB, datos.tipoUnidad, datos.frecuencia, datos.km,
      datos.tipoCobro, datos.cantidad, datos.periodoCantidad, !!datos.requiereAuxiliar
    );

    resultado.esBusquedaTarifa = false;
    resultado.referenciasVigentes = (datos.cliente && datos.cliente !== 'Nuevo Cliente')
      ? buscarTarifasVigentesPorCliente_(datos.cliente) : [];
    resultado.tarifaMeli = null;

    return resultado;
  } catch (e) {
    return { error: e.message || String(e) };
  }
}

/**
 * Busca una tarifa YA EXISTENTE para un Cliente, sin pasar por el motor
 * de costos: solo consulta lo que ya esta capturado en Tarifas Vigentes
 * (y, si Cliente es Mercado Libre y se dan Estacion + Tipo de Unidad +
 * km, la tarifa exacta de ruta dedicada en MELI Tarifas). Es el modo
 * "Tarifa existente" del dashboard de Comercial.
 */
function buscarTarifaExistenteWeb(datos) {
  try {
    var referenciasVigentes = [];
    var tarifaMeli = null;

    if (datos.cliente && datos.cliente !== 'Nuevo Cliente') {
      referenciasVigentes = buscarTarifasVigentesPorCliente_(datos.cliente);

      if (datos.cliente === 'Mercado Libre' && datos.estacionMeli && datos.tipoUnidad && datos.km) {
        tarifaMeli = buscarTarifaMeliDedicada_(datos.estacionMeli, datos.tipoUnidad, datos.km);
      }
    }

    return { referenciasVigentes: referenciasVigentes, tarifaMeli: tarifaMeli };
  } catch (e) {
    return { error: e.message || String(e) };
  }
}

/**
 * Guarda la solicitud calculada (dashboard de Comercial) como un renglon
 * nuevo en la hoja Solicitudes, con los valores ya resueltos. Para
 * Service Partner / XPT (busqueda de tarifa, sin motor de costos) los
 * campos de costo quedan en blanco: no hay nada que inventar ahi.
 *
 * Regresa { ok: true } o { error: mensaje } en vez de "throw" / regresar
 * true a secas, por la misma razon que calcularSolicitudWeb: el bug de
 * Safari con google.script.run no entrega bien los errores lanzados.
 */
function guardarSolicitudWeb(datos, resultado) {
  try {
    return guardarSolicitudWeb_(datos, resultado);
  } catch (e) {
    return { error: e.message || String(e) };
  }
}

function guardarSolicitudWeb_(datos, resultado) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(HOJA_SOLICITUDES);
  if (!sheet) {
    throw new Error('No existe la hoja "Solicitudes". Ejecuta Cotizador BDB > Inicializar hojas.');
  }

  var fila = sheet.getLastRow() + 1;
  var clienteRegistrado = datos.cliente === 'Nuevo Cliente'
    ? 'Nuevo Cliente: ' + (datos.nuevoClienteNombre || '')
    : (datos.cliente || '');
  var tarifaVigenteRegistrada = resultado.tarifaMeli ? resultado.tarifaMeli.tarifa : '';

  sheet.getRange(fila, 1, 1, 34).setValues([[
    new Date(), clienteRegistrado, datos.rutaCliente || '', datos.alcance || '', datos.modalidad,
    datos.puntoA || '', datos.puntoB || '', datos.tipoUnidad,
    datos.frecuencia || '', Number(datos.km) || '', datos.tipoCobro || '', datos.periodoCantidad || '', Number(datos.cantidad) || '',
    datos.requiereAuxiliar ? 'Si' : 'No', datos.estacionMeli || '',
    resultado.puestoPrincipal || '', resultado.puestoAuxiliar || '', resultado.costoCasetas || '',
    resultado.casetasEstimadas ? 'Si' : 'No', resultado.viajesMes || '',
    resultado.sueldoMensual || '', resultado.rentaMensual || '', resultado.mantenimientoMensual || '', resultado.costoGasolinaKm || '',
    resultado.costoVariable || '', resultado.costoFijoProrrateado || '', resultado.costoTotal || '',
    resultado.margenPiso || '', resultado.margenObjetivo || '', resultado.tarifaPiso || '', resultado.tarifaObjetivo || '',
    resultado.tarifaPorUnidadPiso || '', resultado.tarifaPorUnidadObjetivo || '', tarifaVigenteRegistrada
  ]]);

  if (resultado.casetasEstimadas && datos.puntoA && datos.puntoB) {
    agregarCasetaAprendida_(datos.puntoA, datos.puntoB, resultado.costoCasetas);
  }

  return { ok: true };
}

/**
 * Atajos desde el menu de la hoja de calculo para abrir cada dashboard
 * publicado, sin tener que buscar la URL de la implementacion.
 */
function abrirDashboardComercial() {
  abrirUrl_(ScriptApp.getService().getUrl());
}

function abrirDashboardInterno() {
  var base = ScriptApp.getService().getUrl();
  abrirUrl_(base ? base + '?vista=interna' : null);
}

function abrirUrl_(url) {
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Aun no hay una version publicada. Implementar > Nueva implementacion > Aplicacion web.');
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<a href="' + url + '" target="_blank" style="font-family:sans-serif;">Abrir Cotizador BDB</a>' +
    '<script>window.open(' + JSON.stringify(url) + ', "_blank");</script>'
  ).setWidth(300).setHeight(60);
  ui.showModalDialog(html, 'Cotizador BDB');
}
