# Dashboard de dosificación SAP + Prinect · enero–julio 2026

Reporte por orden de producción (OP) basado en la hoja **«desviacion vs equivalente»** del archivo `ENE-JUL_MOVIMIENTO_SAPPRINECT.xlsx`. Los valores de cada OP son los que calcula Excel en esa hoja. Es la versión de enero a julio del [dashboard de agosto](https://alejandratajiboy.github.io/dashboard-sap-prinect/).

1. **Cantidades:** cantidad equivalente vs. cantidad recibida, con desviación absoluta y porcentual.
2. **Costos:** valor monetario, costo unidad ideal y costo unidad sistema, con variación absoluta y porcentual (en quetzales).
3. **SAP vs. Prinect:** pliegos de prensa SAP vs. Prinect, con desviación absoluta y porcentual.
4. **OP sin datos:** todas las OP que no se pueden comparar en alguna sección, con el motivo.
5. **Revisión de la lógica:** observaciones sobre las fórmulas del libro.

Arriba hay un **filtro por mes** (enero a julio). Cambia gráficas, tablas y totales. El archivo también trae OP con fecha de agosto y septiembre: no entran en «Enero–julio» y se ven en la opción «Fuera del periodo».

Cada sección tiene una gráfica por rango de desviación (0–10%, 11–20%, … 91–100% y más de 100%) y una tabla con filtro de estado:

| Sección | Estados sin datos |
|---|---|
| Cantidades | Sin rollo (equivalente 0), Sin entrada (recibida 0), Entrada en metros |
| Costos | Sin rollo, Sin entrada, Entrada en metros |
| SAP vs. Prinect | No está en Prinect (no existe en la hoja «prinect»), Sin cortes en Prinect, Sin entrada SAP, Entrada en metros |

Las OP «Con datos» son las únicas que entran en las gráficas y en los totales.

## Cómo se alimenta

El archivo Excel **no se sube** a este repositorio y **nunca se modifica**. El script `generar_datos.py` hace una copia temporal del Excel, la lee y genera `data.json`. La página `index.html` solo muestra ese `data.json`.

Para actualizar el dashboard después de cambiar el Excel:

1. Guarda y cierra el Excel.
2. Haz doble clic en `actualizar_dashboard.bat`.
3. El script vuelve a leer `C:\Users\dtajiboy\Documents\Dosificadacion\Movimiento\ENE-JUL\ENE-JUL_MOVIMIENTO_SAPPRINECT.xlsx` y publica los datos nuevos en GitHub.
4. La página se actualiza en 1 o 2 minutos.

Para leer otro archivo, arrástralo sobre el `.bat` o ejecuta:

```
actualizar_dashboard.bat "C:\...\otro_archivo.xlsx"
```

El lector busca los encabezados sin importar mayúsculas, espacios ni acentos, así que acepta el formato de agosto («Desviacion_Cantidad_vs_Equiv», «Hoja1») y el de enero–julio («desviacion vs equivalente», «prinect»). El periodo (enero–julio 2026) está en `ANIO` y `MESES_PERIODO` al inicio de `generar_datos.py`.

Requisitos: Python 3 con `openpyxl` (`pip install openpyxl`) y Git con acceso al repositorio.

## Lógica del cálculo

| Concepto | Columna de «desviacion vs equivalente» |
|---|---|
| Cantidad equivalente | Cantidad equivalente |
| Cantidad recibida | cantidad real recibida |
| Desviación absoluta / % | DESVIACION ABSOLUTA / DESVIACION % |
| Valor monetario | Valor Monetario |
| Costo ideal / sistema | Costo Unidad Ideal / Costo unidad Sistema |
| Variación absoluta / % | VARIACION / VARIACION % |
| Desviación Prinect absoluta / % | DESVIACION PRINECT/SAP / DESVIACION % PRINECT/SAP |
| Pliegos SAP / Prinect | suma por OP de «PLIEGOS PRENSA SAP» y «PLIEGOS PRENSA PRINECT» en `general` |
| Mes | mes de «Fecha de contabilización» |

Los totales usan solo las OP con datos del mes elegido. El costo ideal y el costo sistema totales son promedios ponderados por cantidad equivalente y recibida. Como validación, el script recalcula el equivalente (metros de salida × 39.3701 ÷ altura promedio del pliego en pulgadas) y la cantidad recibida desde `general`, y los compara con la hoja.
