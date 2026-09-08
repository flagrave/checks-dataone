# checks-dataone

Control de carga de la planificación del cluster Conosur: compara lo que quedó
cargado en SAP BW contra la hoja de trabajo del negocio, y marca las diferencias
con el responsable de cada línea.

El problema que resuelve es simple de enunciar y molesto de hacer a mano: los dos
lados dicen lo mismo con vocabularios distintos. La query de BW viene a nivel SKU,
con códigos de sociedad y las deducciones abiertas en tres líneas; la hoja de
trabajo viene a nivel marca, por país y con las deducciones sumadas en una. Este
proyecto normaliza ambos a la misma granularidad y recién ahí los resta.

---

## Ciclos

Lo único que cambia entre un ciclo de planificación y otro está en el diccionario
`CYCLES` de `check_sbp.py`:

| Ciclo | Libro | Apertura | Salidas |
|---|---|---|---|
| `sbp` | `Checksibf.xlsx` | anual (SBP 2027, cerrado) | `Check_SBP2027.xlsx`, `Tablero_SBP2027.html` |
| `fc3` | `ChecksibfFC3.xlsx` | actual ene–ago + to go sep–dic (2026) | `Check_FC3_2026.xlsx`, `Tablero_FC3_2026.html` |

El FC3 **no** compara el año contra el año: separa el actual del to go y cruza
cada bloque por su cuenta. Un desvío de más en el actual y uno de menos en el to
go se cancelarían en el total anual y el cruce diría que todo cierra; abiertos,
los dos aparecen.

El año igual se puede mirar: hay un tercer bloque **`Full Year`** que suma los
dos, línea por línea, *después* de cruzados. Es una vista derivada, no un cruce
nuevo — por eso no aparece en `periods_in_play()` (los bloques que se comparan)
sino en `periods_shown()` (los que se muestran): si fuera un bloque más, la hoja
de trabajo repartiría los meses en tres y el año contaría dos veces.

Con el Full Year seleccionado, el tablero avisa cuántas líneas cierran en el año
teniendo desvío en algún bloque — el caso que la apertura estaba justamente para
detectar. Cada registro derivado se lleva el conteo de bloques que no cerraron,
así que el aviso es un número calculado, no una advertencia genérica.

Agregar un ciclo nuevo es agregar una entrada a `CYCLES` (libro, año, si abre por
período, hasta qué mes es actual, nombres de salida). El resto del código no se
toca. `DEFAULT_CYCLE` define cuál corre sin `--cycle`.

---

## Uso

Doble clic en `Actualizar_check.bat`. Eso corre el check, arma el tablero y lo
publica en la carpeta compartida del ciclo.

**Antes de correrlo hay que refrescar la query y guardar el archivo.** Los
scripts leen el xlsx del disco, no la sesión de Excel: si el refresh está sin
guardar, se cruza contra los datos viejos sin avisar.

Los dos scripts también corren sueltos:

```
py check_sbp.py                 # ciclo por defecto -> Check_FC3_2026.xlsx
py dashboard.py                 # ciclo por defecto -> Tablero_FC3_2026.html
py check_sbp.py --cycle sbp     # el ciclo anterior, sin apertura de periodo
py dashboard.py --cycle sbp
```

`--book` fuerza otro libro y `--actual-thru MES` mueve el corte entre actual y to
go (por defecto, el del ciclo).

### Requisitos

- Python 3 con `openpyxl` (probado en 3.14.6 / openpyxl 3.1.5)
- `pywin32` y Excel instalado, sólo para el utilitario `_add_fullyear_sum.py`

---

## Entradas

Nada de esto está en el repo: son datos internos y los excluye el `.gitignore`.
Todos viven en la misma carpeta que los scripts.

| Archivo | Qué es |
|---|---|
| `ChecksibfFC3.xlsx` | El libro del ciclo (`Checksibf.xlsx` para el SBP). Tres hojas relevantes. |
| `FactoresGly.xlsx` | Tabla `SKU → Convertor` para el volumen de glifosato. La mantiene el negocio. |
| `sku_overrides.csv` | `SKU;SBE.3;Comentario`. Reemplaza a la columna `Franchise` que BW no trae a nivel detalle. |

Hojas del libro:

- **`query`** — extracción de BW vía Analysis for Excel. Encabezado en la fila 2,
  datos desde la 3. Granularidad SKU. Las columnas se buscan **por nombre**
  (P&L, company code, CV6, Franchise, material, período contable); el valor es
  siempre la última columna y el texto del material, la que sigue al código. Si
  ningún nombre resuelve y el layout tiene las 7 columnas históricas, se leen por
  posición como antes; si no, el script corta y lista los encabezados que
  encontró en vez de cruzar cualquier cosa.
- **`PreFC Y1`** — la hoja de trabajo. Encabezado en la fila 2, datos desde la 3.
  Acá las columnas también se leen por nombre. En un ciclo con apertura hacen
  falta las doce columnas de mes (`Jan`…`Dec`), no alcanza con `Full Year`.
- **`Pivot`** — Net Sales por país × SBE.1. No la usa el script.

---

## Salidas

**El Excel del ciclo** — siete hojas: `Desvios` (sólo lo que no da OK, ordenado
por impacto), `Reconciliacion` (todo, incluido lo que cierra), `Resumen`,
`Total Conosur` (net sales y sales adj del cluster, por bloque), `Entidad legal`,
`Query normalizada` (útil para auditar cómo clasificó cada SKU) y `Parametros`.

**El tablero HTML** — un archivo solo, sin dependencias, para mandarle a los
owners de cada negocio. Se abre con doble clic. Arriba de todo muestra el total
del cluster, sistema contra hoja, un panel por bloque × línea de P&L — con
apertura son seis: actual, to go y año completo, por Net Sales y Sales Adj —, y
ese bloque no sigue el filtro de período: un total parcial no es un total. Abajo,
el desvío por país, por responsable, las mayores diferencias y la tabla completa,
que sí siguen el filtro y miran un bloque por vez.

Con apertura de período, las claves llevan el bloque adelante (`Actual | ARG | CP
| … | Net Sales`) y tanto el Excel como el tablero lo muestran como una columna
más. La hoja `Total Conosur` del Excel trae las mismas tres filas por línea de
P&L; el resto de las hojas son sólo el cruce, sin el Full Year, para que las
líneas no se cuenten dos veces.

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
empieza el desvío cuando el detalle está lleno de ruido. En un ciclo con apertura
la clave arranca con el bloque (`Actual` / `To Go`), sea cual sea el `--grain`.

**Apertura del período.** Del lado de la query sale del período contable: cada
fila se bucketea por su mes, hasta `actual_thru` es actual y de ahí en más to go.
Del lado de la hoja de trabajo, el actual es la suma de `Jan`…`Aug` y el to go la
de `Sep`…`Dec`. La columna `Full Year` de la hoja no se lee en un ciclo abierto:
el bloque Full Year del tablero sale de sumar los dos bloques ya cruzados, no de
esa columna. Así el año no puede diferir de sus propias partes.

Si la query no trae el período, el script **no reparte nada**: avisa por consola
y en el tablero, y cae al cruce anual del total contra la suma de los doce meses.
Si una línea de la hoja tiene el `Full Year` cargado y los doce meses vacíos, no
hay forma de saber cuánto es actual y cuánto to go: queda fuera del cruce y
también se avisa.

**Escalas.** La query muestra "× 1.000". La plata se multiplica por 1 y el
volumen por 1.000 para dejar los dos lados en miles de USD y en unidades.

**Escala de la hoja de trabajo (`ws_scale`).** No es la misma en todos los
libros: el del SBP carga miles de USD y unidades; el del FC3 carga millones y
miles, mil veces más chico. La columna `Unit` dice `MUSD` y `Kssus` en los dos,
así que la etiqueta no alcanza para decidir. El factor va **declarado** en el
ciclo y después se **verifica**: terminado el cruce, el script calcula la
relación típica (mediana) entre lo que dice el sistema y lo que dice la hoja en
las líneas donde los dos lados tienen dato. Si la escala está bien esa relación
ronda 1 — los desvíos reales mueven cada línea, no la mediana. Si da 1.000 o
0,001, el libro cambió de convención: el script lo avisa por consola y en el
tablero en vez de escupir cuatrocientos desvíos falsos. El valor medido queda en
la hoja `Parametros`.

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
- **`Month`** → el SBP se carga anual, sin apertura; el FC3 la saca del período
  contable de la query, que hay que agregar como característica en Analysis.

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

3. **Países con más de una entidad legal.** Paraguay, Perú y Chile facturan cada
   negocio por una sociedad distinta: Paraguay `2663` (Seeds y GLY) y `0916`
   (Crop Protection); Perú `2690` (Corn), `1321` (GLY) y `0197` (CP); Chile
   `2611` (Corn) y `0164` (GLY y CP). Una fila cargada en la sociedad equivocada
   no se ve en el cruce por país, así que va como control aparte en la hoja
   `Entidad legal`.

---

## Qué se edita en cada ciclo

- **`Actualizar_check.bat`, las cinco variables del arranque** — `CICLO`,
  `TITULO`, `LIBRO`, `TABLERO` y `DEST` (la carpeta compartida donde se publica).
  Es lo único que se toca ahí.
- **`check_sbp.py`, `CYCLES` y `DEFAULT_CYCLE`** — cuando arranca un ciclo nuevo.
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

**Los encabezados de la hoja `query`.** Se buscan por nombre justamente porque
agregar una característica de fila en Analysis corre todas las columnas de lugar.
Si el negocio renombra una columna, el script corta y muestra los encabezados que
leyó; el que queda sin red es el layout viejo de 7 columnas, que sigue leyéndose
por posición para no romper los libros anteriores.

**La query tiene que traer el período contable.** Sin esa columna no hay actual
ni to go, y el FC3 se degrada a un cruce anual. Se agrega en Analysis como
característica de fila (`Posting period` o `Calendar Year/Month`) y se refresca.
El script acepta los formatos que devuelve Analysis: `8`, `008`, `08.2026`,
`2026/08`, `202608`, `2026008` (fiscal) o `AUG 2026`. Los períodos especiales
13–16 no son un mes y quedan afuera, con aviso.

**La hoja de trabajo tiene que tener los meses cargados.** En un ciclo con
apertura, una línea con el anual puesto a mano y los meses vacíos no se puede
partir y queda fuera del cruce.

---

## Archivos

```
check_sbp.py            el cruce; escribe el Excel de reporte
dashboard.py            arma el HTML; importa check_sbp y reusa su lógica
Actualizar_check.bat    corre los dos y publica el tablero
_add_fullyear_sum.py    utilitario: repone las fórmulas de Full Year
```
