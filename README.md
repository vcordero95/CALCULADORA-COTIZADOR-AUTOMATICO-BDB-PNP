# Cotizador Automático BDB

Herramienta en Google Sheets + Apps Script para calcular el costo mínimo y la
tarifa piso de una ruta nueva, sin tener que juntar manualmente sueldos,
gasolina, renta y casetas desde varios archivos.

Origen y contexto completo del proyecto: [docs/Hoja de Proyecto IA - Cotizador Automatico TAREA.pdf](docs/Hoja%20de%20Proyecto%20IA%20-%20Cotizador%20Automatico%20TAREA.pdf).

## Roadmap (por semana)

| Semana | Qué se construye | Estado |
|---|---|---|
| 3 · El PMV | Hoja con lista desplegable de unidad, viajes al mes y km, que calcula costo y tarifa con `Costo ÷ (1 - Margen)` | 🚧 En construcción |
| 4 · Que aguante | Semáforos (20% / 23%), bloqueos si faltan datos, casetas por ruta | Pendiente |
| 5 · Que lo use otro | Vista resumen para Miguel (5 datos clave) | Pendiente |
| 6 · Que se defienda | Prueba con 3 cotizaciones reales, medición de tiempo | Pendiente |

## Dos formas de cotizar

- **Vista interna (costos)** — Vanessa elige directo Tipo de Unidad y
  Puesto/Categoría de Sueldo. Es la hoja **Cotizador** y el dashboard
  `?vista=interna`.
- **Dashboard de Comercial (solicitud de tarifa de cliente)** — Comercial
  no ve costos: captura Tipo de Ruta, Zona/Ciudad, Tipo de Unidad,
  Frecuencia, Volumen, km, Tipo de Cobro y si necesita auxiliar. El
  sistema resuelve solo el Puesto (y el Auxiliar), el costo de casetas por
  zona y los viajes al mes por frecuencia, y calcula la tarifa. Es la hoja
  **Solicitudes** y el dashboard por default (sin `?vista=`).

La tarifa para proveedores de red externa **no** está incluida todavía
(queda para después, como marca el brief original). La comparación contra
tarifas vigentes también queda pendiente hasta cargar ese listado.

## Estructura del proyecto

```
src/
  appsscript.json      Manifiesto del proyecto de Apps Script (incluye acceso de la Web App)
  Code.gs               Menú personalizado, onEdit de la hoja Cotizador
  SetupSheets.gs        Crea/formatea todas las hojas del spreadsheet
  Cotizar.gs            COTIZAR(): costo/tarifa por Tipo de Unidad + Puesto (vista interna)
  Solicitudes.gs        SOLICITAR_TARIFA(): costo/tarifa por Ruta/Zona/Frecuencia (Comercial)
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
  por KM se resuelve solo contra los precios de `Config`.
- **Nomina** — por Puesto / Categoría de Sueldo (no por unidad), incluye
  puestos principales y de auxiliar:
  - **Fiscal**: monto fijo según el periodo (Semanal/Quincenal), tomado de `Config`.
  - **Complemento** = Sueldo − Fiscal.
  - **IMSS** = 30% de Fiscal.
  - **Comisiones** = 5% del Complemento.
  - **Sueldo Mensual Total** = (Sueldo + IMSS + Comisiones) × 4.33 si es Semanal, o × 2 si es Quincenal.
- **Zonas** — Zona / Ciudad → Costo de Casetas.
- **Ruta-Zona-Puesto** — para cada combinación de Tipo de Ruta + Zona, qué
  Puesto Principal (y cuál Auxiliar) aplica. Esto es lo que permite que
  Comercial elija Tipo de Ruta + Zona sin saber nada de nómina.
- **Config** — precios de Diesel/Gasolina ($/L) y la parte Fiscal por
  periodo (Semanal/Quincenal). Editable sin tocar el script.

Registros (se llenan solos desde los dashboards, como historial):

- **Cotizador** — cotizaciones de la vista interna.
- **Solicitudes** — solicitudes hechas por Comercial desde su dashboard.

## Cómo se calcula la tarifa

1. **Costo variable** = Costo de gasolina por km × kilómetros de la ruta.
2. **Costo fijo mensual** = Sueldo Mensual Total (del puesto, o puesto + auxiliar) + Renta Mensual + Mantenimiento Mensual (de la unidad).
3. **Costo fijo prorrateado** = Costo fijo mensual ÷ viajes al mes.
4. **Costo total de la ruta** = Costo variable + Costo fijo prorrateado + Casetas.
5. **Tarifa piso** = Costo total ÷ (1 − Margen).

En el dashboard de Comercial, si el Tipo de Cobro no es "Por Ruta"
(Por Paquete, Por Parada o Por Palet), además se calcula:

6. **Tarifa por unidad** = Tarifa piso ÷ Cantidad (paquetes/paradas/palets capturados).

### Cómo resuelve solo el dashboard de Comercial

- **Puesto Principal / Auxiliar** ← Tipo de Ruta + Zona, buscado en `Ruta-Zona-Puesto`. Si Comercial marca que la ruta necesita auxiliar, se suma también el Sueldo Mensual Total del Puesto Auxiliar mapeado para esa combinación.
- **Costo de Casetas** ← Zona, buscado en `Zonas`.
- **Viajes al Mes** ← Frecuencia (ej. `7x7`, `5x7`): se toma el primer número (veces por semana) × 4.33.

## Puesta en marcha (Google Sheets)

1. Crea un Google Sheet nuevo (o usa el archivo maestro existente).
2. Extensiones → Apps Script.
3. Copia el contenido de cada archivo en `src/` a un archivo del mismo nombre
   en el editor de Apps Script (o usa [`clasp`](https://github.com/google/clasp)
   para hacer `clasp push` apuntando este proyecto a tu spreadsheet).
4. Recarga el spreadsheet. Aparecerá el menú **Cotizador BDB**.
5. Menú **Cotizador BDB → Inicializar hojas**. Esto crea todas las hojas
   listadas arriba, con datos de ejemplo en los catálogos.
6. Reemplaza los datos de ejemplo con tus costos, zonas y mapeos reales en
   **Costos Unidad**, **Nomina**, **Zonas** y **Ruta-Zona-Puesto**.

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
