# checks-dataone

Control de carga del **SBP 2027** para el cluster Conosur: compara lo que quedó
cargado en SAP BW contra la hoja de trabajo del negocio, y marca las diferencias
con el responsable de cada línea.

El problema que resuelve es simple de enunciar y molesto de hacer a mano: los dos
lados dicen lo mismo con vocabularios distintos. La query de BW viene a nivel SKU,
con códigos de sociedad y las deducciones abiertas en tres líneas; la hoja de
trabajo viene a nivel marca, por país y con las deducciones sumadas en una. Este
proyecto normaliza ambos a la misma granularidad y recién ahí los resta.

---

## Uso

Doble clic en `Actualizar_check.bat`. Eso corre el check, arma el tablero y lo
publica en la carpeta compartida del ciclo.

**Antes de correrlo hay que refrescar la query en `Checksibf.xlsx` y guardar el
archivo.** Los scripts leen el xlsx del disco, no la sesión de Excel: si el
refresh está sin guardar, se cruza contra los datos viejos sin avisar.

Los dos scripts también corren sueltos:

```
py check_sbp.py        # escribe Check_SBP2027.xlsx
py dashboard.py        # escribe Tablero_SBP2027.html
```

### Requisitos

- Python 3 con `openpyxl` (probado en 3.14.6 / openpyxl 3.1.5)
- `pywin32` y Excel instalado, sólo para el utilitario `_add_fullyear_sum.py`

---

## Entradas

Nada de esto está en el repo: son datos internos y los excluye el `.gitignore`.
Todos viven en la misma carpeta que los scripts.

| Archivo | Qué es |
|---|---|
| `Checksibf.xlsx` | El libro principal. Tres hojas relevantes. |
| `FactoresGly.xlsx` | Tabla `SKU → Convertor` para el volumen de glifosato. La mantiene el negocio. |
| `sku_overrides.csv` | `SKU;SBE.3;Comentario`. Reemplaza a la columna `Franchise` que BW no trae a nivel detalle. |

Hojas de `Checksibf.xlsx`:

- **`query`** — extracción de BW vía Analysis for Excel. Encabezado en la fila 2,
  datos desde la 3. Granularidad SKU. Las columnas se leen **por posición**:
  P&L, company code, CV6, Franchise, SKU, Material, Valor. Si cambia el orden,
  el script lee cualquier cosa.
- **`PreFC Y1`** — la hoja de trabajo. Encabezado en la fila 2, datos desde la 3.
  Acá las columnas se leen **por nombre**, así que el orden no importa.
- **`Pivot`** — Net Sales por país × SBE.1. No la usa el script.

---

## Salidas

**`Check_SBP2027.xlsx`** — seis hojas: `Desvios` (sólo lo que no da OK, ordenado
por impacto), `Reconciliacion` (todo, incluido lo que cierra), `Resumen`,
`Entidad legal`, `Query normalizada` (útil para auditar cómo clasificó cada SKU)
y `Parametros`.

**`Tablero_SBP2027.html`** — un archivo solo, sin dependencias, para mandarle a
los owners de cada negocio. Se abre con doble clic.

Cada combinación cae en uno de cuatro estados:

| Estado | Significa |
|---|---|
| `OK` | La diferencia está dentro de la tolerancia. |
| `DESVIO` | Los dos lados tienen el dato y no coinciden. |
| `FALTA EN QUERY` | Está en la hoja de trabajo pero no se cargó en el sistema. |
| `SOLO EN QUERY` | Está cargado en el sistema pero no figura en la hoja. |

---

## Cómo compara

**Granularidad.** Por defecto `Country × SBE.1 × SBE.2 × SBE.3 × P&L`. Se puede
subir con `--grain` a `sbe2`, `sbe1` o `country`; sirve para ubicar en qué nivel
empieza el desvío cuando el detalle está lleno de ruido.

**Escalas.** La query muestra "× 1.000". La plata en la hoja de trabajo ya viene
en miles, el volumen viene en unidades — de ahí que la plata se multiplique por 1
y el volumen por 1.000.

**Sólo filas `Del`.** La query refleja delivery. Las filas `Inv` de la hoja de
trabajo son de invoice y no se cargan, así que quedan afuera.

**Líneas de P&L que no se comparan.** `COGS IIa`, `Variances`, `Gross Profit` y
`Net Invoice` existen en la hoja de trabajo pero la query no las trae.

**Deducciones agrupadas.** Comisiones, license fees, otros ajustes y flete se
suman en una sola línea `Sales Adj`. La hoja de trabajo carga el total en una
línea y el sistema lo abre en tres: comparadas por separado dan desvíos que se
cancelan entre sí. Con `--pl-level sub` se comparan línea por línea.

**Tolerancia.** `--tol` en miles USD, por defecto 1. `--tol-pct` agrega un umbral
relativo, por defecto apagado.

### Clasificación

BW no trae tres campos que sí tiene la base del negocio, y cada uno obligó a un
reemplazo:

- **`CV Brand Family`** → el glifosato se detecta por patrones en la descripción
  del material.
- **`Franchise` a nivel detalle** → la query sólo distingue 'Non-Franchise
  Business' de 'Crop Protection', así que la marca sale de `sku_overrides.csv` y
  de los prefijos de descripción (`C DK` → Dekalb, `C LT` → La Tijereta, etc.).
- **`Month`** → la carga del SBP es anual, sin apertura mensual.

Las reglas replican el script de Power Query que el negocio usa sobre BASE D1.

### Los tres casos que no se comparan de frente

Son decisiones deliberadas, y todas responden a lo mismo: **un desvío inventado es
peor que una línea sin comparar**, porque hace perder tiempo persiguiendo un
número que nunca estuvo mal.

1. **Volumen de glifosato.** BW lo trae en otra unidad de medida que la hoja
   (Regs). La conversión es por SKU vía `FactoresGly.xlsx`. Si falta la tabla, se
   excluye todo el volumen de GLY; si falta el factor de un SKU puntual, se
   excluye sólo esa clave — de las dos puntas, porque sacar sólo el lado de la
   query dejaría la línea de la hoja sola y con un desvío del 100%.

2. **Glifosato genérico.** La query lo carga contra un material sin marca; la
   hoja lo reparte entre Roundup y La Tijereta con un criterio que el dato no
   expone. Repartirlo acá con un porcentaje inventado daría dos desvíos falsos
   que se cancelan. Esas combinaciones se colapsan a `GLY (todas las marcas)`,
   el máximo detalle donde los dos lados dicen lo mismo.

3. **Entidad legal de Paraguay.** Hay dos sociedades: `2663` Monsanto Paraguay
   (Seeds y GLY) y `0916` Bayer Paraguay (Crop Protection). Una fila cargada en la
   sociedad equivocada no se ve en el cruce por país, así que va como control
   aparte en la hoja `Entidad legal`.

---

## Qué se edita en cada ciclo

- **`Actualizar_check.bat`, variable `DEST`** — la carpeta compartida donde se
  publica el tablero. Cambia con cada cierre; es la única línea a tocar ahí.
- **`sku_overrides.csv`** — cuando aparece un SKU nuevo que clasifica mal.
- **`FactoresGly.xlsx`** — cuando el negocio agrega SKU de glifosato.

Si un SKU de GLY con volumen no tiene factor, el script lo avisa por consola y en
el tablero, con el detalle de qué quedó afuera. No inventa el factor.

---

## Trampas conocidas

**`Full Year` en `PreFC Y1` estaba pegado en valores.** El script compara contra
esa columna. La carga del SBP se hace en la celda de diciembre, así que el total
anual quedaba con el número viejo y el tablero mostraba desvíos falsos contra una
query que estaba bien. Hoy la columna es `=SUM(TablaSBP[@[Jan]:[Dec]])`.

Quedan diez filas de `Gross Profit` como valor fijo a propósito: tienen el anual
cargado y los doce meses en cero, así que el `=SUM` las pondría en cero y se
perdería el dato. Igual son líneas que el check ignora.

Si un refresh futuro vuelve a pegar la columna como valores, `_add_fullyear_sum.py`
la reconstruye. Usa Excel vía COM y no `openpyxl` a propósito: el libro tiene un
pivot table, la tabla `TablaSBP` y las partes del add-in de SAP Analysis, y
`openpyxl` las tira al reescribir el xlsx.

**El orden de columnas de la hoja `query`.** Se leen por posición. Si Analysis
devuelve las columnas en otro orden, el cruce sale mal sin dar error.

---

## Archivos

```
check_sbp.py            el cruce; escribe el Excel de reporte
dashboard.py            arma el HTML; importa check_sbp y reusa su lógica
Actualizar_check.bat    corre los dos y publica el tablero
_add_fullyear_sum.py    utilitario: repone las fórmulas de Full Year
```
