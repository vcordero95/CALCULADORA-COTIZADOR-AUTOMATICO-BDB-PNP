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
  appsscript.json   Manifiesto del proyecto de Apps Script
  Code.gs            Menú personalizado y arranque
  SetupSheets.gs      Crea/formatea las hojas "Hoja Maestra" y "Cotizador"
  Cotizar.gs          Función COTIZAR(): costo total y tarifa piso
docs/
  Hoja de Proyecto IA - Cotizador Automatico TAREA.pdf   Brief original del proyecto
```

## Cómo se calcula la tarifa

1. **Costo variable** = Costo de gasolina por km (según Rendimiento y si la unidad usa Diesel o Gasolina) × kilómetros de la ruta.
2. **Costo fijo mensual** = Sueldo Mensual Total + Renta Mensual + Mantenimiento Mensual de la unidad.
3. **Costo fijo prorrateado** = Costo fijo mensual ÷ viajes al mes.
4. **Costo total de la ruta** = Costo variable + Costo fijo prorrateado + Casetas.
5. **Tarifa piso** = Costo total ÷ (1 − Margen).

### Cómo se arma el Sueldo Mensual Total

En la Hoja Maestra, el sueldo se captura por periodo (Semanal o Quincenal),
y el resto se calcula solo:

- **Fiscal**: monto fijo según el periodo (referencia editable en `Hoja Maestra!Q3:R4`, no cambia con el sueldo capturado).
- **Complemento** = Sueldo − Fiscal.
- **IMSS** = 30% de Fiscal.
- **Comisiones** = 5% del Complemento.
- **Sueldo Mensual Total** = (Sueldo + IMSS + Comisiones) × 4.33 si es Semanal, o × 2 si es Quincenal.

El precio de Diesel y Gasolina ($/L) también es editable en `Hoja Maestra!Q1:R2`
y aplica a todas las unidades según su Tipo de Combustible.

## Puesta en marcha (Google Sheets)

1. Crea un Google Sheet nuevo (o usa el archivo maestro existente).
2. Extensiones → Apps Script.
3. Copia el contenido de cada archivo en `src/` a un archivo del mismo nombre
   en el editor de Apps Script (o usa [`clasp`](https://github.com/google/clasp)
   para hacer `clasp push` apuntando este proyecto a tu spreadsheet).
4. Recarga el spreadsheet. Aparecerá el menú **Cotizador BDB**.
5. Menú **Cotizador BDB → Inicializar hojas**. Esto crea:
   - **Hoja Maestra**: catálogo de tipos de unidad con renta, mantenimiento,
     rendimiento/tipo de combustible y nómina (sueldo, periodicidad, fiscal,
     complemento, IMSS, comisiones).
   - **Cotizador**: captura de rutas nuevas (unidad por lista desplegable, km,
     viajes al mes, casetas, margen) con el cálculo automático de costo y
     tarifa piso.
6. Llena tus costos reales en **Hoja Maestra** (reemplaza los datos de ejemplo).

## Prueba rápida (definición de "quedó" de la semana 3)

Calcular una ruta sencilla en la hoja **Cotizador** en menos de 5 minutos:
elige el tipo de unidad de la lista, captura km, viajes al mes y casetas, y
verifica que el costo total y la tarifa piso aparezcan solos.
