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

## Estructura del proyecto

```
src/
  appsscript.json   Manifiesto del proyecto de Apps Script (incluye acceso de la Web App)
  Code.gs            Menú personalizado y arranque
  SetupSheets.gs      Crea/formatea las hojas "Hoja Maestra" y "Cotizador"
  Cotizar.gs          Función COTIZAR(): costo total y tarifa piso
  WebApp.gs           Sirve el dashboard y expone COTIZAR() a la Web App
  Dashboard.html       Interfaz del dashboard (Web App)
docs/
  Hoja de Proyecto IA - Cotizador Automatico TAREA.pdf   Brief original del proyecto
```

## Cómo se calcula la tarifa

El Tipo de Unidad y el Puesto / Categoría de Sueldo son **independientes**:
en el Cotizador se eligen por separado, cada uno de su propio catálogo en
la Hoja Maestra.

1. **Costo variable** = Costo de gasolina por km (según Rendimiento y si la unidad usa Diesel o Gasolina) × kilómetros de la ruta.
2. **Costo fijo mensual** = Sueldo Mensual Total (del puesto elegido) + Renta Mensual + Mantenimiento Mensual (de la unidad elegida).
3. **Costo fijo prorrateado** = Costo fijo mensual ÷ viajes al mes.
4. **Costo total de la ruta** = Costo variable + Costo fijo prorrateado + Casetas.
5. **Tarifa piso** = Costo total ÷ (1 − Margen).

### Las dos tablas de la Hoja Maestra

- **Costos por Tipo de Unidad** (columnas A:H): Renta Mensual, Mantenimiento
  Mensual y Rendimiento/Tipo de Combustible, con sus columnas diarias y de
  costo de gasolina calculadas solas.
- **Nómina** (columnas J:Q), por Puesto / Categoría de Sueldo, no por unidad:

  - **Fiscal**: monto fijo según el periodo (referencia editable en `Hoja Maestra!S3:T4`, no cambia con el sueldo capturado).
  - **Complemento** = Sueldo − Fiscal.
  - **IMSS** = 30% de Fiscal.
  - **Comisiones** = 5% del Complemento.
  - **Sueldo Mensual Total** = (Sueldo + IMSS + Comisiones) × 4.33 si es Semanal, o × 2 si es Quincenal.

El precio de Diesel y Gasolina ($/L) también es editable en `Hoja Maestra!S1:T2`
y aplica a todas las unidades según su Tipo de Combustible.

## Puesta en marcha (Google Sheets)

1. Crea un Google Sheet nuevo (o usa el archivo maestro existente).
2. Extensiones → Apps Script.
3. Copia el contenido de cada archivo en `src/` a un archivo del mismo nombre
   en el editor de Apps Script (o usa [`clasp`](https://github.com/google/clasp)
   para hacer `clasp push` apuntando este proyecto a tu spreadsheet).
4. Recarga el spreadsheet. Aparecerá el menú **Cotizador BDB**.
5. Menú **Cotizador BDB → Inicializar hojas**. Esto crea:
   - **Hoja Maestra**: dos catálogos independientes — Costos por Tipo de
     Unidad (renta, mantenimiento, rendimiento/tipo de combustible) y
     Nómina por Puesto / Categoría de Sueldo (sueldo, periodicidad, fiscal,
     complemento, IMSS, comisiones).
   - **Cotizador**: captura de rutas nuevas (Tipo de Unidad y Puesto por
     lista desplegable, km, viajes al mes, casetas, margen) con el cálculo
     automático de costo y tarifa piso.
6. Llena tus costos reales en **Hoja Maestra** (reemplaza los datos de ejemplo).

## Prueba rápida (definición de "quedó" de la semana 3)

Calcular una ruta sencilla en la hoja **Cotizador** en menos de 5 minutos:
elige el tipo de unidad y el puesto/categoría de sueldo de las listas,
captura km, viajes al mes y casetas, y verifica que el costo total y la
tarifa piso aparezcan solos.

## Dashboard (Web App)

Además de la hoja Cotizador, el mismo cálculo (`COTIZAR()`) está disponible
como una Web App independiente para solicitud/revisión de tarifa de
cliente — uso interno (Comercial/Sihanka, Miguel), sin exponer la hoja de
cálculo. La tarifa para proveedores de red externa **no** está incluida
todavía (queda para después, como marca el brief original).

- **Acceso**: restringido al dominio de Google Workspace de BDB
  (`src/appsscript.json` → `webapp.access: "DOMAIN"`). Nadie fuera de BDB
  puede abrir el link.
- **Publicar/actualizar la Web App**: `clasp deploy` (o Implementar → Nueva
  implementación → Aplicación web desde el editor de Apps Script). Cada
  `clasp push` actualiza el código; `clasp deploy` es lo que publica esa
  versión en el link ya existente.
- **Abrir el link**: menú **Cotizador BDB → Abrir dashboard** en el
  spreadsheet.
- **Guardar cotización**: el botón "Guardar cotización" del dashboard
  agrega un renglón a la hoja Cotizador con los valores ya calculados (no
  fórmulas), como registro fijo de la tarifa que se cotizó en ese momento.
