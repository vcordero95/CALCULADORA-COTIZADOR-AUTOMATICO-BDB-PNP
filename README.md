# Cotizador Automático BDB

Herramienta en Google Sheets + Apps Script para calcular el costo mínimo y la
tarifa piso de una ruta nueva, sin tener que juntar manualmente sueldos,
gasolina, renta y casetas desde varios archivos.

Origen y contexto completo del proyecto: [docs/Hoja de Proyecto IA - Cotizador Automatico TAREA.pdf](docs/Hoja%20de%20Proyecto%20IA%20-%20Cotizador%20Automatico%20TAREA.pdf).

## Roadmap (por semana)

| Semana | Qué se construye | Estado |
|---|---|---|
| 3 · El PMV | Hoja con lista desplegable de unidad, viajes al mes y km, que calcula costo y tarifa con `Costo ÷ (1 - Margen)` | 🚧 En construcción |
| 4 · Que aguante | Semáforos (20% Piso / 30% Objetivo), bloqueos si faltan datos | Casetas por ruta (Punto A/B) ya implementado; semáforo visual pendiente |
| 5 · Que lo use otro | Vista resumen para Miguel (5 datos clave) | Pendiente |
| 6 · Que se defienda | Prueba con 3 cotizaciones reales, medición de tiempo | Pendiente |

## Dos formas de cotizar

- **Vista interna (costos)** — Vanessa elige directo Tipo de Unidad y
  Puesto/Categoría de Sueldo. Es la hoja **Cotizador** y el dashboard
  `?vista=interna`.
- **Dashboard de Comercial (solicitud de tarifa de cliente)** — Comercial
  no ve costos ni margen. Primero elige qué necesita:
  - **Tarifa nueva** (calcular desde cero): captura Cliente, **Alcance**
    (Local/Foráneo) y **Modalidad** (Dedicada, Spot, Service Partner,
    Line Haul, Media Milla, XPT) por separado — cualquier Modalidad
    puede ser Local o Foránea —, Tipo de Unidad, Frecuencia, km, Tipo de
    Cobro y si necesita auxiliar. Si el Tipo de Cobro no es "Por Ruta"
    (Por Paquete/Por Parada/Por Palet), se despliega Cantidad con su
    Periodo (Por Día o Por Semana, según como lo haya pedido el
    cliente): si es Por Semana, se divide entre los viajes por semana de
    la Frecuencia para convertirla a cantidad por viaje antes de sacar
    la Tarifa por Unidad. Punto A / Punto B solo si la Modalidad es Line
    Haul, Media Milla,
    Service Partner o XPT (rutas de punto a punto; Dedicada y Spot no
    los piden). Según la Modalidad, el sistema calcula distinto:
    - **Dedicada, Spot, Line Haul, Media Milla**: motor de costos completo
      (resuelve el Puesto por Alcance, el costo de casetas — solo Line
      Haul/Media Milla — y los viajes al mes por frecuencia; calcula la
      Tarifa Piso 20% y la Tarifa Objetivo 30% con la política fija de
      margen de la empresa).
    - **Service Partner**: es tarifa de red ya negociada con el cliente,
      no un costo de ruta (ver Hoja 9 del tarifario original: Tarifa Base
      + Diferenciador Foráneo = Total). En vez de calcular, se **busca**
      la tarifa vigente de ese Cliente; si no hay una para ese punto, hay
      que negociarla con el cliente en vez de inventar un costo.
    - **XPT** (solo Mercado Libre): usa la misma tabla estandarizada de
      rutas dedicadas de MELI (Estación → Nivel → `MELI Tarifas`) que ya
      existía — se **busca**, no se calcula desde costo.

    Los resultados calculados (no las búsquedas) se guardan en la hoja
    **Solicitudes**.
  - **Tarifa existente** (buscar en el tarifario): elige Cliente y
    **Tipo de Servicio** (Last Mile, XPT, Dedicada, Spot, Service
    Partner, Line Haul, Media Milla, Helper — "Last Mile" y "XPT" solo
    aparecen para Mercado Libre). Si el Tipo de Servicio es Last Mile o
    XPT, captura Estación + Tipo de Unidad + km y se resuelve con el
    grid de Nivel de `MELI Tarifas` (son el mismo cálculo, separados
    porque para el negocio son dos servicios distintos). Para el resto,
    captura **Estación / Ruta** (texto libre, ej. "SCQ1", "Tapalpa",
    "NACIONAL") y se busca en `Tarifas Vigentes` filtrando por esa
    modalidad y esa estación/ruta — regresa solo la tarifa pedida, no
    todo el tarifario del cliente. Nada de esto recalcula: solo busca lo
    que ya está registrado.

  Ambos modos comparten el dashboard por default (sin `?vista=`).

La tarifa para proveedores de red externa **no** está incluida todavía
(queda para después, como marca el brief original).

### Comparación contra tarifas vigentes

- El dropdown **Cliente** sale de `Tarifas Vigentes` (un import del
  tarifario real de clientes) más la opción fija **"Nuevo Cliente"** (para
  cotizaciones de clientes que aún no tienen tarifa registrada — en ese
  caso no se muestra ninguna comparación).
- Para cualquier cliente, al calcular se listan **todas** sus tarifas
  vigentes registradas (Estado Vigencia = Activa, de cualquier Tipo De
  Servicio — Last Mile, Line Haul, Media Milla, marcado en la columna
  "Servicio") como referencia — no es un match exacto automático, porque
  los nombres de vehículo y rutas en el tarifario real no siempre calzan
  literal con los catálogos internos.
- Para **Mercado Libre + Por Ruta** (ruta dedicada), además aparece un
  dropdown de **Estación MELI**: con esa estación (nodo) + Tipo de Unidad
  + km sí se calcula una tarifa vigente exacta contra `MELI Tarifas`
  (Nivel L1-L4 × rango de km), usando el nivel de esa estación en
  `MELI Estaciones`.
- `Tarifas Vigentes`, `MELI Estaciones` y `MELI Tarifas` tienen datos
  reales de clientes: se cargan solo dentro del Google Sheet, nunca se
  suben a git/GitHub (ver `.gitignore`: `src/TarifasVigentes.gs`).

## Estructura del proyecto

```
src/
  appsscript.json      Manifiesto del proyecto de Apps Script (incluye acceso de la Web App)
  Code.gs               Menú personalizado, onEdit de la hoja Cotizador
  SetupSheets.gs        Crea/formatea todas las hojas del spreadsheet
  Cotizar.gs            COTIZAR(): costo/tarifa por Tipo de Unidad + Puesto (vista interna)
  Solicitudes.gs        SOLICITAR_TARIFA(): costo/tarifa por Alcance/Modalidad/Punto A-B/Frecuencia (Comercial)
  TarifasVigentes.gs    Datos reales de tarifas de clientes + MELI (gitignored, no se sube a GitHub)
  WebApp.gs             Sirve los dos dashboards y expone las funciones de calculo
  Dashboard.html         Dashboard interno (costos)
  DashboardComercial.html Dashboard de Comercial (solicitud de tarifa de cliente)
docs/
  Hoja de Proyecto IA - Cotizador Automatico TAREA.pdf   Brief original del proyecto
```

## Las hojas del spreadsheet

Catálogos (los mantiene Vanessa):

- **Costos Unidad** — por Tipo de Unidad: Renta Mensual, Mantenimiento
  Mensual, Rendimiento (km/L) y Tipo de Combustible. El Costo de Gasolina
  por KM se resuelve solo contra los precios de `Config`. También incluye
  el **Código Vehículo INEGI** (0 Motocicleta, 1 Automóvil, 2-4 Autobús de
  2 a 4 ejes, 5-12 Camión de 2 a 9 ejes), que se usa para consultar el
  costo de caseta real de cada Tipo de Unidad en la API de Ruteo de
  INEGI (ver `Casetas`).
- **Nomina** — por Puesto / Categoría de Sueldo (no por unidad), incluye
  puestos principales y de auxiliar:
  - **Fiscal**: monto fijo según el periodo (Semanal/Quincenal), tomado de `Config`.
  - **Complemento** = Sueldo − Fiscal.
  - **IMSS** = 30% de Fiscal.
  - **Comisiones** = 5% del Complemento.
  - **Sueldo Mensual Total** = (Sueldo + IMSS + Comisiones) × 4.33 si es Semanal, o × 2 si es Quincenal.
- **Puntos** — catálogo de ciudades usado como sugerencias para Punto A /
  Punto B. Viene precargado con
  cobertura nacional (las 32 capitales estatales + las plazas
  logísticas/industriales más relevantes de cada estado, ~150 ciudades en
  total). En el dashboard, Punto A / Punto B son **texto libre** (con
  estas ciudades como autocompletado): Comercial puede escribir cualquier
  ciudad del país, esté o no en este catálogo — no se bloquea si el
  destino real no aparece en la lista. Vanessa puede agregar más ciudades
  aquí cuando haga falta.
- **Casetas** — costo de casetas entre Punto A y Punto B (sin importar el
  orden). Trae ~18 corredores precargados con el costo de **camión de 2
  ejes (C2)** tomado directo del PDF oficial de CAPUFE, "Tarifas Vigentes
  2026" (Red FONADIN) — la fuente autoritativa, no prensa (columna
  Fuente cita el tramo oficial exacto). Los que dicen "suma de tramos"
  son corredores conocidos armados encadenando tramos oficiales
  consecutivos (ej. CDMX-Acapulco = México-Cuernavaca + Cuernavaca-
  Acapulco); vale la pena confirmar que la ruta real siga ese mismo
  camino. CDMX-Toluca no es red CAPUFE (va por un tramo concesionado
  aparte) y se dejó el dato de prensa. Este catálogo ya **no es la
  primera fuente que se consulta**: para cotizar siempre con el dato más
  actualizado posible, el sistema primero consulta **en vivo la API de
  Ruteo de INEGI (SAKBÉ)** con el Código Vehículo INEGI del Tipo de
  Unidad elegido (así que Auto y Camión dan un costo de caseta distinto
  para la misma ruta, como en la realidad). Solo si INEGI no puede
  resolverla (lugar no encontrado, falla de red) se usa el costo exacto
  de este catálogo, y si tampoco está aquí, se cae al estimado por
  kilómetros × Costo Casetas Estimado ($/km) de `Config` — tres niveles
  de respaldo: INEGI en vivo → catálogo exacto → estimado por km, para
  nunca bloquear una cotización.
  El token de INEGI vive en `Config` (fila "Token INEGI (API Ruteo
  SAKBE)"), editable sin tocar el script. Nota: INEGI es una fuente
  confiable pero no perfecta — en pruebas coincidió exacto con CAPUFE en
  algunos corredores (ej. CDMX-Puebla) y se desvió en otros (ej. CDMX-
  Querétaro, por la selección del punto de origen dentro de la ciudad).
  **El catálogo se auto-completa solo**: cada vez que se guarda una
  solicitud cuya caseta no vino de un match exacto en este catálogo
  (porque se resolvió con INEGI o con el estimado por km), esa ruta
  (Punto A, Punto B, costo, con nota de que viene de una solicitud
  guardada) se agrega automáticamente a `Casetas` como respaldo para el
  día en que INEGI no responda. Vanessa puede después ir ajustando esos
  valores con el costo real de CAPUFE.
- **Alcance-Puesto** — el Puesto Principal (y el Auxiliar) que aplica para
  cada **Alcance** (Local/Foráneo). Es independiente de la Modalidad:
  cualquier Modalidad (Dedicada, Line Haul, Media Milla, etc.) puede ser
  Local o Foránea, y el chofer se asigna solo por Alcance.
- **Config** — precios de Diesel/Gasolina ($/L) — se mantienen solos:
  con el menú **Cotizador BDB → Activar actualización semanal de
  precios de combustible (INEGI)** se crea un disparador que cada lunes
  jala el precio promedio nacional más reciente de la API de INEGI
  (la misma que resuelve casetas) y lo escribe aquí, para no quedarse
  con un precio capturado una sola vez y desactualizado; sigue siendo
  editable a mano entre una actualización y otra. También la parte
  Fiscal por periodo (Semanal/Quincenal), la política de margen del
  dashboard de Comercial (**Margen Piso 20%** / **Margen Objetivo
  30%**), el
  **Costo Casetas Estimado ($/km)** — ~$4.00/km, promedio nacional de
  camión de 2 ejes calculado con cobertura de prensa 2026 sobre tarifas
  CAPUFE en México-Querétaro, México-Puebla, México-Toluca,
  Cuernavaca-Acapulco y México-Cuernavaca (último respaldo si ni la API
  de INEGI ni el catálogo `Casetas` resuelven la ruta), y el
  **Token INEGI (API Ruteo SAKBE)** usado para consultar casetas en vivo
  como primera fuente. Editable sin tocar el script.

Referencia de tarifas de clientes (datos reales, cargados desde el
tarifario de la empresa — solo dentro del Sheet, no en git):

- **Tarifas Vigentes** — tarifas activas por cliente (Last Mile), para
  comparar contra lo calculado.
- **MELI Estaciones** — Estación (nodo) → Nivel (L1-L4), para rutas
  dedicadas de Mercado Libre.
- **MELI Tarifas** — tarifa vigente por Vehículo × Nivel × rango de km,
  para rutas dedicadas de Mercado Libre.
- **Amazon Tarifas** — Rate Card DSP de Amazon Logistics Mexico por
  Estación (DMT6, DMT4 — misma tarifa para ambas) × Ciclo (Cycle 1,
  Same Day, MCO, Late Same Day, Nursery) × Tipo de Vehículo, con
  Tarifa Final y Tarifa Final con Helper (solo donde Amazon la ofrece).
  Por ahora solo tiene la Zona 3; se agregan más zonas según se vayan
  pasando.

Estas 4 hojas de referencia se pueden editar directo (agregar tarifas,
clientes, estaciones o ciclos nuevos) sin tocar el script: las
búsquedas y los catálogos de los dashboards siempre leen lo que
realmente hay en la hoja en ese momento, no un tamaño fijo. El menú
**Cotizador BDB → Cargar/actualizar Tarifas Vigentes y MELI** sigue
existiendo pero es destructivo (borra y reescribe estas 4 hojas desde
el catálogo semilla del código), así que ya no hace falta para el
día a día — solo se usa cuando Vanessa confirma que no tiene capturas
manuales pendientes que se perderían.

Registros (se llenan solos desde los dashboards, como historial):

- **Cotizador** — cotizaciones de la vista interna.
- **Solicitudes** — solicitudes hechas por Comercial desde su dashboard.

## Cómo se calcula la tarifa

1. **Costo variable** = Costo de gasolina por km × kilómetros de la ruta.
2. **Costo fijo mensual** = Sueldo Mensual Total (del puesto, o puesto + auxiliar) + Renta Mensual + Mantenimiento Mensual (de la unidad).
3. **Costo fijo prorrateado** = Costo fijo mensual ÷ viajes al mes.
4. **Costo total de la ruta** = Costo variable + Costo fijo prorrateado + Casetas.
5. **Tarifa** = Costo total ÷ (1 − Margen).

En la hoja **Cotizador** (vista interna) el margen lo captura Vanessa por
fila. En el dashboard de **Comercial** el margen no se captura: se
calculan ambas tarifas con la política fija de `Config` — **Tarifa Piso**
(Margen Piso, 20%) y **Tarifa Objetivo** (Margen Objetivo, 30%).

Si el Tipo de Cobro no es "Por Ruta" (Por Paquete, Por Parada o Por
Palet), además se calcula cada una entre la Cantidad capturada (ya
convertida a cantidad por viaje si se capturó "Por Semana"):

6. **Tarifa por unidad (Piso/Objetivo)** = Tarifa (Piso/Objetivo) ÷ Cantidad por viaje.

### Cómo resuelve solo el dashboard de Comercial

- **¿Se capturan Punto A / Punto B?** ← lo decide la **Modalidad**: `Line Haul`, `Media Milla`, `Service Partner` y `XPT` sí (rutas de punto a punto); `Dedicada` y `Spot` no.
- **Puesto Principal / Auxiliar** ← solo por **Alcance** (Local/Foráneo), buscado en `Alcance-Puesto`, sin importar la Modalidad. Si Comercial marca que la ruta necesita auxiliar, se suma también el Sueldo Mensual Total del Puesto Auxiliar de ese Alcance.
- **Costo de Casetas** (solo Modalidad Line Haul/Media Milla) ← se consulta primero en vivo la API de Ruteo de INEGI con el Código Vehículo INEGI del Tipo de Unidad elegido; si INEGI no responde, se busca Punto A + Punto B (sin importar el orden) en `Casetas`; si tampoco está ahí, se estima con km × Costo Casetas Estimado de `Config` (la fuente exacta usada se marca en el resultado y en la hoja Solicitudes).
- **Viajes al Mes** ← Frecuencia (ej. `7x7`, `5x7`): se toma el primer número (veces por semana) × 4.33.
- **XPT** ← no pasa por nada de lo anterior: es búsqueda de tarifa estandarizada MELI (ver arriba), no cálculo de costo, sin importar si es Tarifa Nueva o Existente. **Service Partner** sí usa el motor de costo igual que Dedicada/Spot/Line Haul/Media Milla cuando es Tarifa Nueva (puede o no llevar Punto A/B según si también es Line Haul/Media Milla).

## Puesta en marcha (Google Sheets)

1. Crea un Google Sheet nuevo (o usa el archivo maestro existente).
2. Extensiones → Apps Script.
3. Copia el contenido de cada archivo en `src/` a un archivo del mismo nombre
   en el editor de Apps Script (o usa [`clasp`](https://github.com/google/clasp)
   para hacer `clasp push` apuntando este proyecto a tu spreadsheet).
4. Recarga el spreadsheet. Aparecerá el menú **Cotizador BDB**.
5. Menú **Cotizador BDB → Inicializar hojas**. Esto crea todas las hojas
   listadas arriba, con datos de ejemplo en los catálogos.
6. Reemplaza los datos de ejemplo con tus costos, puntos, casetas y
   mapeos reales en **Costos Unidad**, **Nomina**, **Puntos**, **Casetas**
   y **Alcance-Puesto**.

## Prueba rápida (definición de "quedó" de la semana 3)

Calcular una ruta sencilla en la hoja **Cotizador** en menos de 5 minutos:
elige el tipo de unidad y el puesto/categoría de sueldo de las listas,
captura km, viajes al mes y casetas, y verifica que el costo total y la
tarifa piso aparezcan solos.

## Dashboards (Web App)

Ambos dashboards viven en la misma Web App; la URL decide cuál se sirve.

- **Acceso**: restringido al dominio de Google Workspace de BDB
  (`src/appsscript.json` → `webapp.access: "DOMAIN"`). Nadie fuera de BDB
  puede abrir el link.
- **Publicar/actualizar la Web App**: `clasp deploy` (o Implementar → Nueva
  implementación → Aplicación web desde el editor de Apps Script). Cada
  `clasp push` actualiza el código; `clasp deploy` es lo que publica esa
  versión en el link ya existente.
- **Abrir los links**: menú **Cotizador BDB → Abrir dashboard Comercial**
  o **→ Abrir dashboard interno (costos)** en el spreadsheet. El dashboard
  de Comercial es la URL base (`.../exec`); el interno es esa misma URL
  con `?vista=interna`.
- **Guardar**: el botón "Guardar cotización"/"Guardar solicitud" agrega un
  renglón a la hoja correspondiente (Cotizador o Solicitudes) con los
  valores ya calculados (no fórmulas), como registro fijo de lo que se
  cotizó en ese momento.
