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
    puede ser Local o Foránea —, Tipo de Unidad, Frecuencia, Volumen
    (cantidad de paquetes), km, Tipo de Cobro y si necesita auxiliar; y
    Punto A / Punto B solo si la Modalidad es Line Haul, Media Milla,
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
  - **Tarifa existente** (buscar en el tarifario): solo elige Cliente (y,
    para Mercado Libre, opcionalmente Estación + Tipo de Unidad + km) y
    ve directo lo que ya está registrado en `Tarifas Vigentes` / `MELI
    Tarifas`, sin recalcular nada.

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
  aparte) y se dejó el dato de prensa. Esto sigue sin ser exhaustivo, pero
  ya no depende solo de este catálogo: si Comercial cotiza una ruta que
  todavía no está aquí, el sistema consulta **en vivo la API de Ruteo de
  INEGI (SAKBÉ)** con el Código Vehículo INEGI del Tipo de Unidad elegido
  (así que Auto y Camión dan un costo de caseta distinto para la misma
  ruta, como en la realidad), y solo si INEGI tampoco puede resolverla
  (lugar no encontrado, falla de red) cae al estimado por kilómetros ×
  Costo Casetas Estimado ($/km) de `Config` — para poder cotizar
  cualquier ruta del país desde el día uno, con tres niveles de
  respaldo: catálogo exacto → INEGI en vivo → estimado por km.
  El token de INEGI vive en `Config` (fila "Token INEGI (API Ruteo
  SAKBE)"), editable sin tocar el script. Nota: INEGI es una fuente
  confiable pero no perfecta — en pruebas coincidió exacto con CAPUFE en
  algunos corredores (ej. CDMX-Puebla) y se desvió en otros (ej. CDMX-
  Querétaro, por la selección del punto de origen dentro de la ciudad),
  así que las rutas frecuentes conviene seguir capturándolas a mano en
  `Casetas` con el dato oficial.
  **El catálogo se auto-completa solo**: cada vez que se guarda una
  solicitud cuya caseta fue estimada (por km o por INEGI), esa ruta
  (Punto A, Punto B, costo, con nota de que viene de una solicitud
  guardada) se agrega automáticamente a `Casetas`, así que la próxima vez
  que pidan esa misma ruta ya no es un estimado. Vanessa puede después ir
  ajustando esos valores aprendidos con el costo real de CAPUFE.
- **Alcance-Puesto** — el Puesto Principal (y el Auxiliar) que aplica para
  cada **Alcance** (Local/Foráneo). Es independiente de la Modalidad:
  cualquier Modalidad (Dedicada, Line Haul, Media Milla, etc.) puede ser
  Local o Foránea, y el chofer se asigna solo por Alcance.
- **Config** — precios de Diesel/Gasolina ($/L), la parte Fiscal por
  periodo (Semanal/Quincenal), la política de margen del dashboard de
  Comercial (**Margen Piso 20%** / **Margen Objetivo 30%**), el
  **Costo Casetas Estimado ($/km)** — ~$4.00/km, promedio nacional de
  camión de 2 ejes calculado con cobertura de prensa 2026 sobre tarifas
  CAPUFE en México-Querétaro, México-Puebla, México-Toluca,
  Cuernavaca-Acapulco y México-Cuernavaca (último respaldo si ni el
  catálogo `Casetas` ni la API de INEGI resuelven la ruta), y el
  **Token INEGI (API Ruteo SAKBE)** usado para consultar casetas en vivo.
  Editable sin tocar el script.

Referencia de tarifas de clientes (datos reales, cargados desde el
tarifario de la empresa — solo dentro del Sheet, no en git):

- **Tarifas Vigentes** — tarifas activas por cliente (Last Mile), para
  comparar contra lo calculado.
- **MELI Estaciones** — Estación (nodo) → Nivel (L1-L4), para rutas
  dedicadas de Mercado Libre.
- **MELI Tarifas** — tarifa vigente por Vehículo × Nivel × rango de km,
  para rutas dedicadas de Mercado Libre.

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
Palet), además se calcula cada una entre la Cantidad capturada:

6. **Tarifa por unidad (Piso/Objetivo)** = Tarifa (Piso/Objetivo) ÷ Cantidad (paquetes/paradas/palets).

### Cómo resuelve solo el dashboard de Comercial

- **¿Se capturan Punto A / Punto B?** ← lo decide la **Modalidad**: `Line Haul`, `Media Milla`, `Service Partner` y `XPT` sí (rutas de punto a punto); `Dedicada` y `Spot` no.
- **Puesto Principal / Auxiliar** ← solo por **Alcance** (Local/Foráneo), buscado en `Alcance-Puesto`, sin importar la Modalidad. Si Comercial marca que la ruta necesita auxiliar, se suma también el Sueldo Mensual Total del Puesto Auxiliar de ese Alcance.
- **Costo de Casetas** (solo Modalidad Line Haul/Media Milla) ← se busca Punto A + Punto B (sin importar el orden) en `Casetas`; si esa ruta exacta no está capturada, se consulta en vivo la API de Ruteo de INEGI con el Código Vehículo INEGI del Tipo de Unidad elegido; si tampoco la resuelve, se estima con km × Costo Casetas Estimado de `Config` (se marca como "estimado" en el resultado y en la hoja Solicitudes, con la fuente exacta, para que Vanessa sepa qué rutas conviene capturar con dato real).
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
