"""
Check de carga IBF - Conosur
============================

Compara lo cargado en SAP BW (hoja `query`) contra la hoja de trabajo
(hoja `PreFC Y1`) del libro del ciclo, normalizando ambos lados a la misma
granularidad: Country x SBE.1 x SBE.2 x SBE.3 x P&L.

Sirve para los dos ciclos, que se eligen con `--cycle`:

    sbp   SBP 2027, cerrado. Carga anual, sin apertura mensual: se compara
          la columna `Full Year` de la hoja contra el total de la query.
    fc3   FC3 2026, el ciclo vivo. La query trae el actual de enero a agosto
          y el forecast de septiembre a diciembre en el mismo numero, asi que
          el cruce se abre en dos bloques -Actual y To Go- y cada uno se
          compara por separado. Ver SECCION 1B.

El nombre del archivo quedo de la primera version (SBP); el script ya no es
de un ciclo solo.

Las reglas de clasificacion replican el script de Power Query de BASE D1,
adaptadas a las columnas que trae la extraccion de BW (que no tiene
`CV Brand Family` ni `Franchise` a nivel detalle).

Uso:
    py check_sbp.py                      # ciclo por defecto (fc3)
    py check_sbp.py --cycle sbp          # el ciclo cerrado
    py check_sbp.py --tol 5              # tolerancia 5 (miles USD)
    py check_sbp.py --grain sbe1         # compara a nivel SBE.1 en vez de SBE.3
    py check_sbp.py --actual-thru 9      # mueve el corte actual/to go a septiembre
"""

import argparse
import datetime as dt
import os
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# SECCION 1 - PARAMETROS Y CONFIGURACION
# ============================================================

SHEET_QUERY = "query"
SHEET_WS = "PreFC Y1"

# Fila de encabezado y primera fila de datos (1-based)
QUERY_HEADER_ROW = 2
QUERY_FIRST_DATA_ROW = 3
WS_HEADER_ROW = 2
WS_FIRST_DATA_ROW = 3

# La query muestra "* 1.000". La plata en la hoja de trabajo ya esta en miles,
# el volumen esta en unidades. De ahi los dos factores distintos.
SCALE_MONEY = 1.0
SCALE_VOLUME = 1000.0

# Escala de la hoja de trabajo, por ciclo (`ws_scale` en CYCLES).
# La columna `Unit` dice "MUSD" y "Kssus" en los dos libros, pero las cifras
# no estan en la misma escala: el libro del SBP carga miles de USD y unidades,
# el del FC3 carga millones y miles. La etiqueta no sirve para decidir, asi
# que el factor va declarado en el ciclo y despues se verifica contra el cruce
# (ver check_ws_scale): si no da, se avisa en vez de comparar mal.
WS_SCALE = 1.0
# Relacion tipica query/hoja a partir de la cual se sospecha de la escala.
WS_SCALE_MIN, WS_SCALE_MAX = 0.5, 2.0
# Debajo de esto una linea es ruido y no vota en la mediana.
WS_SCALE_FLOOR = 10.0
# La relacion medida en el cruce; la completa check_ws_scale().
WS_SCALE_MEDIAN = None

# El volumen de GLY viene en otra unidad de medida en la query que en la hoja
# de trabajo (Regs). La conversion es por SKU: volumen query / factor = Regs.
# La tabla vive afuera, en un Excel que mantiene el negocio.
GLY_FACTOR_FILE = "FactoresGly.xlsx"
GLY_FACTOR_SHEET = "Hoja1"
GLY_FACTOR_COLS = ("SKU", "Convertor")
FACTOR_MIN, FACTOR_MAX = 0.5, 20.0

# Se completa en main() con {SKU: factor}. Vacio -> no hay tabla y el volumen
# de GLY queda fuera del cruce, como cuando el factor estaba en stand by.
GLY_FACTORS = {}
# SKU de GLY con volumen en la query que no figuran en la tabla de factores.
# No se inventa un factor: esas lineas quedan fuera del cruce y se avisa.
GLY_SIN_FACTOR = {}

# Tolerancia por defecto, en las mismas unidades que la plata (miles USD).
DEFAULT_TOL = 1.0
# Tolerancia relativa: ademas del umbral absoluto, ignorar desvios por debajo de esto.
DEFAULT_TOL_PCT = 0.0

# Company code (Auth) -> Country de la hoja de trabajo.
COCD_TO_COUNTRY = {
    "2601": "ARG",
    "0164": "CHI",   # Bayer Chile (CP y GLY)
    "2611": "CHI",   # Corn
    "0197": "PER",   # Bayer Peru (CP)
    "2690": "PER",   # Corn
    "1321": "PER",   # GLY
    "1386": "BOL",
    "0194": "UGY",
    "0916": "PGY",  # Bayer Paraguay
    "2663": "PGY",  # Monsanto Paraguay
}

# Nombre de la entidad legal, para el reporte.
# Los tres codigos nuevos van con una etiqueta descriptiva hasta que se
# confirme la razon social; no cambia el cruce, solo el texto del reporte.
COCD_NAME = {
    "0916": "Bayer Paraguay",
    "2663": "Monsanto Paraguay",
    "0164": "Bayer Chile",
    "2611": "Chile - Corn",
    "0197": "Bayer Peru",
    "2690": "Peru - Corn",
    "1321": "Peru - GLY",
}

# Paises con mas de una entidad legal: que negocio se factura por cual.
# Se reporta como control aparte; una fila cargada en la entidad equivocada
# no se ve en el cruce por pais pero si es un error de carga.
ENTITY_RULES = {
    "PGY": {
        "Corn": "2663", "SOY": "2663", "GLY": "2663", "OTHERS": "2663",
        "CP": "0916",
    },
    "PER": {
        "Corn": "2690", "GLY": "1321", "CP": "0197",
    },
    "CHI": {
        "Corn": "2611", "GLY": "0164", "CP": "0164",
    },
}

# Solo se comparan las filas de la hoja de trabajo con este flag.
# La query refleja delivery; las filas `Inv` son de invoice y no se cargan.
WS_INV_DEL_KEEP = {"Del"}

# Columna de la hoja de trabajo que contiene el valor anual.
WS_VALUE_COL = "Full Year"

# Columnas mensuales de la hoja de trabajo, en orden.
WS_MONTH_COLS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


# ============================================================
# SECCION 1B - CICLOS Y APERTURA DE PERIODO
# ============================================================
# Un ciclo es un libro, un anio y una forma de partir el anio. El SBP se
# carga de una sola vez para todo el anio; el FC3 tiene el actual ya cerrado
# hasta agosto y el to go de septiembre a diciembre, y son dos cosas
# distintas: que el total cierre no dice nada si el actual esta bien y el
# forecast mal por el mismo importe con el signo cambiado.

CYCLES = {
    "sbp": {
        "label": "SBP 2027",
        "book": "Checksibf.xlsx",
        "year": 2027,
        "split": False,          # carga anual: no hay actual ni to go
        "actual_thru": None,
        "ws_scale": 1.0,         # la hoja carga miles de USD y unidades
        "out_xlsx": "Check_SBP2027.xlsx",
        "out_html": "Tablero_SBP2027.html",
    },
    "fc3": {
        "label": "FC3 2026",
        "book": "ChecksibfFC3.xlsx",
        "year": 2026,
        "split": True,
        "actual_thru": 8,        # agosto es el ultimo mes cerrado
        "ws_scale": 1000.0,      # la hoja carga millones y miles de unidades
        "out_xlsx": "Check_FC3_2026.xlsx",
        "out_html": "Tablero_FC3_2026.html",
    },
}
DEFAULT_CYCLE = "fc3"

# Nombres de los dos bloques. Cortos a proposito: son parte de la clave de
# cruce y encabezan columnas del reporte. El rango de meses de cada uno se
# arma en period_label() y va en Parametros y en el tablero.
P_ACTUAL = "Actual"
P_TOGO = "To Go"
P_FULLYEAR = "Full Year"   # modo degradado: la query no trae el mes

MONTH_ES = ("Ene", "Feb", "Mar", "Abr", "May", "Jun",
            "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")

# Estado del ciclo en curso. Se completa en setup_cycle(); el tablero lo lee
# para saber si tiene que mostrar la dimension Periodo.
CYCLE = CYCLES[DEFAULT_CYCLE]
PERIOD_SPLIT = False        # True solo si ademas la query trae el mes
ACTUAL_THRU = None
# Se completa al leer la query: como se resolvio la columna de mes y que
# quedo afuera. Todo esto se reporta; nada se resuelve en silencio.
PERIOD_COL_NAME = None
PERIOD_MISSING = False      # la query no trae mes -> se degrada a Full Year
PERIOD_OTHER_YEAR = {}      # (anio, mes) -> importe de filas de otro anio
PERIOD_UNPARSED = {}        # valor crudo -> veces que aparecio


def setup_cycle(name, actual_thru=None):
    """Fija el ciclo en curso. Devuelve su configuracion."""
    global CYCLE, PERIOD_SPLIT, ACTUAL_THRU, WS_SCALE
    CYCLE = CYCLES[name]
    ACTUAL_THRU = actual_thru or CYCLE["actual_thru"]
    PERIOD_SPLIT = bool(CYCLE["split"] and ACTUAL_THRU)
    WS_SCALE = CYCLE.get("ws_scale", 1.0)
    return CYCLE


def period_of_month(month):
    """Mes (1-12) -> bloque del ciclo."""
    if not PERIOD_SPLIT:
        return P_FULLYEAR
    return P_ACTUAL if month <= ACTUAL_THRU else P_TOGO


def period_label(period):
    """'Actual' -> 'Actual (Ene-Ago)'. Para titulos y parametros."""
    if period == P_ACTUAL:
        return f"Actual ({MONTH_ES[0]}-{MONTH_ES[ACTUAL_THRU - 1]})"
    if period == P_TOGO:
        return f"To Go ({MONTH_ES[ACTUAL_THRU]}-{MONTH_ES[11]})"
    if period == P_FULLYEAR and PERIOD_SPLIT:
        return f"Full Year ({MONTH_ES[0]}-{MONTH_ES[11]})"
    return period


def periods_in_play():
    """Bloques que se comparan en este ciclo, en orden de calendario."""
    return [P_ACTUAL, P_TOGO] if PERIOD_SPLIT else [P_FULLYEAR]


def periods_shown():
    """Bloques que se muestran: los que se cruzan, mas el ano completo.

    El Full Year no es un bloque mas del cruce: es la suma de los otros dos,
    derivada despues (ver fullyear_recs). Va aparte de periods_in_play() justo
    por eso -- si entrara ahi, read_worksheet repartiria los meses en tres
    bloques y el ano se contaria dos veces.
    """
    return periods_in_play() + ([P_FULLYEAR] if PERIOD_SPLIT else [])

# --- Tokens de clasificacion (espejo del script M) ---
CP_TOKENS = ("FUNGICIDES", "HERBICIDES", "INSECTICIDES", "SEEDGROWTH")
OTHERS_TOKENS = ("COTTON", "OTHER", "OTHERS", "OIL", "OILSEEDS")

CORN_BRANDS = {"DEKALB", "LA TIJERETA", "AGROCERES", "AGROESTE", "LT"}
GLY_BRANDS = {"ROUNDUP", "LA TIJERETA"}

ALLOWED_SBE3_CORN = {
    "AGROCERES", "AGROESTE", "DEKALB", "DOW LICENSING",
    "FULL PACK", "LA TIJERETA", "LICENSE SEED", "PRIVATE LABEL SEEDS",
}
ALLOWED_SBE3_SOY = {"CERTIFIED SEED", "SAVED SEED", "POD", "GERM"}

# SBE3 que quedan exentos del blanqueo de volumen para corn sin prefijo "C".
EXCLUDED_SBE3 = {"DOW LICENSING", "FULL PACK", "LICENSE SEED"}

# Materiales IBC que van a volumen cero.
IBC_VOL_ZERO = {
    "IBC LLENO, LTP II, 1000 L, AR",
    "IBC LLENO, SNIPER II, 1000 L, AR",
    "IBC LLENO,RUP TOP,1000 L,AR",
    "IBC VACIO, 1000 L, AR",
    "IBC VACIO, RUP TOP, 1000 L, AR",
    "IBC VACIO,1000 L,AR",
}

CREATE_SKU = "12845195"
SCALE_THRESHOLD = 1_000_000
SCALE_DIVISOR_PRIVATE_LABEL = 21328

# SKUs con clasificacion forzada (del script M).
SKU_OVERRIDES = {
    "64184626": "Dow Licensing",
    "64184456": "Dow Licensing",
    "62014979": "Dow Licensing",  # sin Franchise no se puede desambiguar Dekalb/License Seed
}

# Tabla externa SKU -> SBE.3, editable por el usuario. Reemplaza a la columna
# `Franchise` a nivel detalle, que la extraccion de BW no trae.
# Formato CSV: SKU;SBE.3;Comentario
SKU_OVERRIDE_FILE = "sku_overrides.csv"

# Prefijo de descripcion de material -> marca de Corn.
CORN_PREFIX_TO_SBE3 = {
    "C DK": "Dekalb",
    "C LT": "La Tijereta",
    "C AG": "Agroeste",
    "C AC": "Agroceres",
    "C OL": "Private Label Seeds",
}

# P&L: query -> hoja de trabajo (ya trimmeada)
PL_QUERY_TO_WS = {
    "Quantity": "Volume",
    "Net Sales": "Net Sales",
    "Commission Expenses": "Commission Expenses",
    "Other Sales Adjustments": "Other Sales Adjustments",
    "Outward Freight, Transport Ins.": "Outward Freight, Transport Ins.",
}
# Lineas de la hoja de trabajo que la query no trae: no se comparan.
WS_PL_IGNORE = {"COGS IIa", "Variances", "Gross Profit", "Net Invoice"}

# Agrupacion P&L2 del script M: las lineas de deduccion se comparan sumadas.
# La hoja de trabajo carga el total en una sola linea y el sistema lo abre en
# tres, asi que a nivel detalle el cruce da falsos desvios.
PL_GROUP = {
    "Volume": "Volume",
    "Net Sales": "Net Sales",
    "Commission Expenses": "Sales Adj",
    "License Fees": "Sales Adj",
    "Other Sales Adjustments": "Sales Adj",
    "Outward Freight, Transport Ins.": "Sales Adj",
}


# ============================================================
# SECCION 2 - FUNCIONES HELPER
# ============================================================

def clean_upper(x):
    return ("" if x is None else str(x)).strip().upper()


def clean_trim(x):
    return ("" if x is None else str(x)).strip()


def contains_any(txt, tokens):
    up = txt.upper()
    return any(t in up for t in tokens)


def to_num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace(" ", "")
    if not s:
        return 0.0
    # SAP puede escribir el signo al final
    neg = s.endswith("-")
    s = s.rstrip("-").replace(",", "")
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return -n if neg else n


def material_norm(mat):
    return re.sub(r"[ \-_./]", "", mat)


# ============================================================
# SECCION 3 - CLASIFICACION GLY
# ============================================================
# El script M decide GLY con `CV Brand Family`, que la query no trae.
# Aca se reemplaza por los patrones de descripcion de material, que es
# de donde salen igual las marcas. Validado contra ARG (253.800 exacto).

RUP_PATTERNS = ("RUP ", "RUP,", "ROUNDUP", "FAENA", "CONTROLMAX")
TIJERETA_GLY_PATTERNS = ("LA TIJERETA", "TIJERETA", "SNIPER", "ESTRELLA")


def is_rup(mat):
    up = mat.upper()
    if up.startswith("RUP"):
        return True
    return any(p in up for p in RUP_PATTERNS)


def is_tijereta_gly(mat):
    up = mat.upper()
    return any(p in up for p in TIJERETA_GLY_PATTERNS)


def is_ibc_gly(mat):
    up = mat.upper()
    if "IBC LLENO" not in up and "IBC VACIO" not in up:
        return False
    return not any(x in up for x in ("PUMA", "CROPSTAR", "CREATE"))


# Glifosato generico: la query lo carga contra un material sin marca (llegan
# como "Non-Franchise Business"). El negocio es GLY, pero la marca no viene en
# el dato: no se puede saber si es Roundup o La Tijereta.
GLY_GENERIC_PATTERNS = ("GLYPHOSATE", "GLIFOSATO")
GLY_SIN_MARCA = "GLY sin marca"
GLY_ROLLUP = "GLY (todas las marcas)"


def is_gly_generico(mat):
    up = mat.upper()
    if is_rup(up) or is_tijereta_gly(up):
        return False
    return any(p in up for p in GLY_GENERIC_PATTERNS)


def is_gly(mat, cv6):
    """GLY solo aplica a filas de herbicidas o a los IBC."""
    if is_ibc_gly(mat):
        return True
    if "HERBICIDES" not in cv6.upper():
        return False
    return is_rup(mat) or is_tijereta_gly(mat) or is_gly_generico(mat)


# ============================================================
# SECCION 4 - CLASIFICACION SBE3 (query)
# ============================================================

EXTRA_OVERRIDES = {}


def load_sku_overrides(path):
    """Carga la tabla externa SKU -> SBE.3 (opcional)."""
    out = {}
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8-sig") as fh:
        for ln, line in enumerate(fh, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.split(";")]
            if len(parts) < 2 or parts[0].upper() == "SKU":
                continue
            out[parts[0]] = parts[1]
    return out


def load_gly_factors(path):
    """Carga la tabla SKU -> factor de conversion de volumen de GLY a Regs.

    Se resuelven las columnas por nombre de encabezado, no por posicion, para
    que agregar una columna al Excel no rompa la lectura.
    """
    out = {}
    if not os.path.exists(path):
        return out
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        raise SystemExit(
            f"{path} esta abierto en Excel y bloqueado. Cerralo y volve a correr.")
    ws = wb[GLY_FACTOR_SHEET]
    it = ws.iter_rows(values_only=True)
    header = [clean_trim(c) for c in next(it)]
    idx = {name: i for i, name in enumerate(header) if name}
    col_sku, col_fac = GLY_FACTOR_COLS
    if col_sku not in idx or col_fac not in idx:
        raise SystemExit(
            f"{path}: falta la columna '{col_sku}' o '{col_fac}' "
            f"en la hoja '{GLY_FACTOR_SHEET}'. Encontradas: {header}")
    # Los factores reales viven entre 1 y 4. Un valor afuera de esa banda no es
    # un factor: es un volumen pegado en la columna equivocada. Ese error no da
    # un resultado absurdo a la vista (divide y devuelve ~1 Reg), asi que si no
    # lo cortamos aca pasa como un desvio plausible pero falso.
    fuera = []
    for r in it:
        sku = clean_trim(r[idx[col_sku]])
        fac = to_num(r[idx[col_fac]])
        if not sku or not fac:
            continue
        if not (FACTOR_MIN <= fac <= FACTOR_MAX):
            fuera.append((sku, fac))
            continue
        out[sku] = fac
    wb.close()
    if fuera:
        detalle = "\n".join(f"    SKU {s}: {f:g}" for s, f in fuera)
        raise SystemExit(
            f"{path}: hay factores fuera del rango esperado "
            f"({FACTOR_MIN}-{FACTOR_MAX}) en la columna '{col_fac}'.\n"
            f"{detalle}\n"
            "  Reviso que no se haya pegado el volumen en vez del factor.")
    return out


def classify_sbe3(cv6, mat, sku):
    """Devuelve el SBE.3 de una fila de la query.

    Espejo de la Seccion 5 del script M, sin las ramas que dependen de
    `CV Brand Family` ni de `Franchise` a nivel detalle.
    """
    if clean_trim(sku) in EXTRA_OVERRIDES:
        return EXTRA_OVERRIDES[clean_trim(sku)]
    cv6u = clean_upper(cv6)
    matu = clean_upper(mat)
    matn = material_norm(matu)
    sku = clean_trim(sku)

    is_corn = "CORN" in cv6u or "MAIZ" in cv6u
    is_soy = "SOY" in cv6u or "SOYBEAN SEED" in cv6u
    is_cp = contains_any(cv6u, CP_TOKENS)
    is_others = contains_any(cv6u, OTHERS_TOKENS)

    # --- GLY (prioridad maxima, igual que en M) ---
    if is_gly(matu, cv6u):
        if is_gly_generico(matu):
            return GLY_SIN_MARCA
        if is_ibc_gly(matu):
            # El IBC de Tijereta se separa; el resto va a Roundup.
            return "La Tijereta" if is_tijereta_gly(matu) else "Roundup"
        if is_tijereta_gly(matu):
            return "La Tijereta"
        return "Roundup"

    # --- Crop Protection: SBE.3 es el codigo entre parentesis del CV6 ---
    # "SeedGrowth (SGR)" -> SGR. El script M usa Text.Start(cv6,3), que para
    # SeedGrowth devuelve "SEE" y no matchea contra la hoja de trabajo.
    if is_cp:
        m = re.search(r"\(([A-Z]{2,4})\)", cv6u)
        return m.group(1) if m else cv6u[:3]

    # --- Overrides por SKU ---
    if sku in SKU_OVERRIDES:
        return SKU_OVERRIDES[sku]

    # --- SOY ---
    if is_soy:
        if "GERMOPLASMA" in matu:
            return "GERM"
        if "POD" in matu or "PUNTO DE ENTREGA" in matu:
            return "POD"
        if ("CERTIFICADA" in matu or "CERTIFIED" in matu
                or " SC " in matu or matu.endswith("SC")
                or "-SC" in matu or "_SC" in matu or "/SC" in matu):
            return "Certified Seed"
        if ("RESERVA" in matu or " RG " in matu or matu.endswith("RG")
                or "-RG" in matu or "_RG" in matu or "/RG" in matu):
            return "Saved Seed"
        return "UX"

    # --- CORN / Seeds por patron de material ---
    for pref, brand in CORN_PREFIX_TO_SBE3.items():
        if matu.startswith(pref):
            return brand
    if "COL" in matn:
        return "Private Label Seeds"
    if matu.startswith("TF C") or matu.startswith("ROY"):
        return "License Seed"
    if is_corn and (matu.startswith("CP") or matu.startswith("ACA")
                    or matu.startswith("GDM")):
        return "Full Pack"
    if is_others:
        return "Other Seeds"
    if is_corn:
        return "UX"
    return "UX"


def apply_sbe3_overrides(sbe3, cv6, mat):
    """Seccion 6 del script M."""
    cv6u = clean_upper(cv6)
    matu = clean_upper(mat)
    is_corn = "CORN" in cv6u or "MAIZ" in cv6u
    if is_corn and matu == "CP.R2506YTCJZ.T1B2.80M.T.US":
        return "License Seed"
    if clean_upper(sbe3) == "LA TIJERTA":
        return "La Tijereta"
    return sbe3


def finalize_sbe3(sbe3, cv6):
    """Overrides Corn/Soy + reasignacion de UX remanente (final Seccion 5)."""
    cv6u = clean_upper(cv6)
    is_corn = "CORN" in cv6u or "MAIZ" in cv6u
    is_soy = "SOY" in cv6u or "SOYBEAN SEED" in cv6u
    is_cp = contains_any(cv6u, CP_TOKENS)

    up = clean_upper(sbe3)
    if is_corn and up not in ALLOWED_SBE3_CORN:
        sbe3, up = "UX", "UX"
    if is_soy and up not in ALLOWED_SBE3_SOY:
        sbe3, up = "UX", "UX"

    if up == "UX" and is_corn:
        return "Dekalb"
    if up == "UX" and is_cp:
        return "Other HER"
    return sbe3


# ============================================================
# SECCION 5 - BUSINESS (SBE.1), SBE.2 y SBE
# ============================================================

def classify_business(cv6, mat, sbe3):
    """Seccion 11 del script M -> se corresponde con SBE.1 de la hoja."""
    cv6u = clean_upper(cv6)
    matu = clean_upper(mat)
    sbe3u = clean_upper(sbe3)

    is_corn = "CORN" in cv6u or "MAIZ" in cv6u
    tijereta_gly = sbe3u == "LA TIJERETA" and not is_corn
    full_gly = (is_gly(matu, cv6u) or sbe3u in ("IBC", "ROUNDUP")
                or tijereta_gly)

    if full_gly:
        return "GLY"
    if "SOYBEAN SEED" in cv6u:
        return "SOY"
    if "CP_DIFFERENCE (CP_DIF)" in cv6u:
        return "CP"
    if "NOT DIVIDED INTO SBE (CS)" in cv6u:
        return "OTHERS"
    if contains_any(cv6u, CP_TOKENS):
        return "CP"
    if is_corn:
        return "CORN"
    if "SOY" in cv6u:
        return "SOY"
    if contains_any(cv6u, OTHERS_TOKENS):
        return "OTHERS"
    return cv6u


# La hoja de trabajo escribe SBE.1 con otro casing que el script M.
BUSINESS_TO_SBE1 = {"CORN": "Corn", "SOY": "SOY", "CP": "CP",
                    "GLY": "GLY", "OTHERS": "OTHERS"}


def classify_sbe2(business, sbe3, cv6):
    """Seccion 13, mapeado al vocabulario de la hoja de trabajo."""
    sbe3u = clean_upper(sbe3)
    if business == "CORN":
        return "Branded" if sbe3u in CORN_BRANDS else "Licences"
    if business == "GLY":
        return "Branded" if sbe3u in GLY_BRANDS else "No-Brand"
    if business == "CP":
        return clean_trim(cv6)  # "Fungicides (FUN)"
    return "Branded"


def classify_sbe(business):
    return "SEEDS" if business in ("CORN", "SOY") else "CHEM"


# ============================================================
# SECCION 6 - AJUSTES DE VOLUMEN (Secciones 9 y 11B del script M)
# ============================================================

def adjust_volume(value, cv6, mat, sku, sbe3, business, country):
    cv6u = clean_upper(cv6)
    matu = clean_upper(mat)
    sbe3u = clean_upper(sbe3)
    is_corn = "CORN" in cv6u or "MAIZ" in cv6u

    if matu in {m.upper() for m in IBC_VOL_ZERO}:
        return 0.0
    if is_corn and sbe3u == "UX":
        return 0.0
    if is_corn and not matu.startswith("C") and sbe3u not in EXCLUDED_SBE3:
        return 0.0
    if business == "OTHERS" and country != "BOL":
        return 0.0
    if clean_trim(sku) == CREATE_SKU:
        return value / 1000.0
    if sbe3u == "PRIVATE LABEL SEEDS" and abs(value) > SCALE_THRESHOLD:
        return value / SCALE_DIVISOR_PRIVATE_LABEL

    # GLY: la query trae kg/litros, la hoja de trabajo Regs. El factor es por
    # SKU. Si el SKU no esta en la tabla no se adivina: se anota y la linea
    # queda fuera del cruce (drop_gly_volume la saca).
    if business == "GLY":
        fac = GLY_FACTORS.get(clean_trim(sku))
        if not fac:
            GLY_SIN_FACTOR[clean_trim(sku)] = {
                "Country": country, "SBE.3": sbe3,
                "Material": clean_trim(mat), "Value": value}
            return value
        return value / fac

    return value


# ============================================================
# SECCION 7 - LECTURA Y NORMALIZACION
# ============================================================

# --- Resolucion de columnas de la query ---------------------------------
# Historicamente se leian por posicion, y eso es exactamente lo que se rompe
# al agregarle el mes a la extraccion: Analysis mete las caracteristicas de
# fila a la izquierda y corre todo lo demas un lugar. Ahora se resuelven por
# encabezado, con la posicion fija como respaldo para el layout viejo.
QUERY_COL_PATTERNS = {
    "pl":        ("p&l", "p and l", "pyl"),
    "cocd":      ("company code", "sociedad", "cocd"),
    "cv6":       ("cv 6", "cv6", "strategic business"),
    "franchise": ("franchise", "franquicia"),
    "sku":       ("material", "sku"),
    "period":    ("month", "mes", "period", "periodo", "calmonth",
                  "calendar year", "fiscal"),
}
QUERY_COLS_LEGACY = {"pl": 0, "cocd": 1, "cv6": 2, "franchise": 3,
                     "sku": 4, "mat": 5, "period": None}
QUERY_NCOLS_LEGACY = 7


def resolve_query_cols(header):
    """Encabezado de la query -> indice de cada columna que usa el cruce.

    El valor es siempre la ultima columna: las caracteristicas van a la
    izquierda y el ratio a la derecha, asi que sumar una caracteristica no lo
    mueve. La descripcion del material es la columna que sigue al SKU y viene
    sin encabezado, que es como Analysis muestra clave y texto.
    """
    low = [clean_trim(h).lower() for h in header]
    cols = {}
    for field, pats in QUERY_COL_PATTERNS.items():
        for i, h in enumerate(low):
            if h and any(p in h for p in pats):
                cols[field] = i
                break

    faltan = [f for f in ("pl", "cocd", "cv6", "franchise", "sku")
              if f not in cols]
    if faltan:
        # Layout viejo (encabezados en otro idioma o celdas fusionadas): se
        # cae a las posiciones fijas solo si la grilla es la de siempre.
        if len(header) == QUERY_NCOLS_LEGACY:
            return dict(QUERY_COLS_LEGACY), True
        raise SystemExit(
            f"No se reconocieron columnas de la hoja '{SHEET_QUERY}': "
            f"falta {', '.join(faltan)}. Encabezado leido: "
            f"{[h for h in header if h]!r}. Revisa que el refresh de Analysis "
            f"haya dejado el encabezado en la fila {QUERY_HEADER_ROW}.")

    cols["mat"] = cols["sku"] + 1
    cols.setdefault("period", None)
    return cols, False


# --- Lectura del mes ----------------------------------------------------
MONTH_TOKENS = {
    "JAN": 1, "ENE": 1, "FEB": 2, "MAR": 3, "APR": 4, "ABR": 4, "MAY": 5,
    "JUN": 6, "JUL": 7, "AUG": 8, "AGO": 8, "SEP": 9, "SET": 9, "OCT": 10,
    "NOV": 11, "DEC": 12, "DIC": 12,
}


def parse_period(v, year_default):
    """Valor de la columna de mes -> (anio, mes). None si no se entiende.

    Analysis escribe el periodo de varias formas segun la caracteristica que
    se haya puesto en las filas: '08.2026', '2026/08', '202608', 'AUG 2026',
    y el periodo fiscal como '2026008'. Se aceptan todas; lo que no se
    entiende se cuenta y se avisa, nunca se adivina.
    """
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime) or isinstance(v, dt.date):
        return v.year, v.month

    s = clean_upper(str(v))
    if not s:
        return None

    # Nombre de mes, con o sin anio: 'AUG', 'AUG 2026', 'AGO-26'
    for tok, m in MONTH_TOKENS.items():
        if tok in s:
            y = re.search(r"(20\d{2})", s)
            return (int(y.group(1)) if y else year_default), m

    # Con separador se leen las partes por lo que son, no concatenadas:
    # '8.2026' es agosto de 2026 y no el numero 82026.
    partes = [p for p in re.split(r"[^0-9]+", s) if p]
    if len(partes) == 2:
        a, b = (int(p) for p in partes)
        anio, mes = (a, b) if a > 12 else (b, a)
        if anio < 100:
            anio += 2000
        return (anio, mes) if 1 <= mes <= 12 else None

    digits = re.sub(r"\D", "", s)
    if not digits:
        return None
    if len(digits) == 7:                       # periodo fiscal YYYY0PP
        y, m = int(digits[:4]), int(digits[-2:])
    elif len(digits) == 6:                     # YYYYMM o MMYYYY
        if 2000 <= int(digits[:4]) <= 2100:
            y, m = int(digits[:4]), int(digits[4:])
        else:
            y, m = int(digits[2:]), int(digits[:2])
    elif len(digits) in (1, 2, 3):             # posting period suelto: 8, 08, 008
        y, m = year_default, int(digits)
    else:
        return None
    # Los periodos especiales de cierre (13 a 16) no son un mes: caen como
    # ilegibles y se avisan, en vez de colarse en el to go.
    return (y, m) if 1 <= m <= 12 else None


def read_query(path):
    global PERIOD_COL_NAME, PERIOD_MISSING, PERIOD_OTHER_YEAR, PERIOD_UNPARSED
    PERIOD_OTHER_YEAR, PERIOD_UNPARSED = defaultdict(float), defaultdict(int)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_QUERY]
    it = ws.iter_rows(min_row=QUERY_HEADER_ROW, values_only=True)
    header = list(next(it))
    cols, legacy = resolve_query_cols(header)
    vcol = len(header) - 1
    pcol = cols["period"]
    PERIOD_COL_NAME = clean_trim(header[pcol]) if pcol is not None else None
    # Sin mes en la query no se puede saber que parte del numero es actual y
    # que parte es forecast. Antes de inventar un corte, se compara el anual:
    # una linea sin abrir es mejor que dos desvios inventados.
    PERIOD_MISSING = PERIOD_SPLIT and pcol is None

    rows = []
    unmapped_cocd = set()
    for r in it:
        pl_raw = clean_trim(r[cols["pl"]])
        if not pl_raw:
            continue
        cocd = clean_trim(r[cols["cocd"]])
        cv6 = clean_trim(r[cols["cv6"]])
        franchise = clean_trim(r[cols["franchise"]])
        sku = clean_trim(r[cols["sku"]])
        mat = clean_trim(r[cols["mat"]])
        val = to_num(r[vcol])

        # Periodo: fuera del modo split todo cae en un unico bloque.
        periodo = P_FULLYEAR
        if PERIOD_SPLIT and pcol is not None:
            ym = parse_period(r[pcol], CYCLE["year"])
            if ym is None:
                PERIOD_UNPARSED[clean_trim(r[pcol])] += 1
                continue
            anio, mes = ym
            if anio != CYCLE["year"]:
                PERIOD_OTHER_YEAR[(anio, mes)] += val
                continue
            periodo = period_of_month(mes)

        country = COCD_TO_COUNTRY.get(cocd)
        if country is None:
            unmapped_cocd.add(cocd)
            country = f"?{cocd}"

        pl = PL_QUERY_TO_WS.get(pl_raw, pl_raw)

        sbe3 = classify_sbe3(cv6, mat, sku)
        sbe3 = apply_sbe3_overrides(sbe3, cv6, mat)
        sbe3 = finalize_sbe3(sbe3, cv6)
        business = classify_business(cv6, mat, sbe3)
        sbe1 = BUSINESS_TO_SBE1.get(business, business)
        sbe2 = classify_sbe2(business, sbe3, cv6)
        sbe = classify_sbe(business)

        if pl == "Volume":
            val = val * SCALE_VOLUME
            val = adjust_volume(val, cv6, mat, sku, sbe3, business, country)
        else:
            val = val * SCALE_MONEY

        rows.append({
            "Periodo": periodo,
            "Country": country, "CoCd": cocd, "SBE": sbe, "SBE.1": sbe1,
            "SBE.2": sbe2, "SBE.3": sbe3, "P&L": pl, "SKU": sku,
            "Material": mat, "CV6": cv6, "Franchise": franchise,
            "Value": val,
        })
    wb.close()
    if legacy:
        print(f"Nota         : la hoja '{SHEET_QUERY}' no trae encabezados "
              f"reconocibles; se leyo por posicion (layout de {QUERY_NCOLS_LEGACY} "
              f"columnas)")
    return rows, unmapped_cocd


# Lineas de la hoja con los doce meses vacios pero con anual cargado: el
# valor esta ahi pero no se puede repartir entre actual y to go.
WS_SIN_MESES = {}


def read_worksheet(path):
    """Lee la hoja de trabajo.

    En el ciclo anual devuelve una fila por linea, con la columna `Full Year`.
    Con apertura de periodo devuelve una fila por linea y por bloque, sumando
    los meses de cada bloque: ene-ago para el actual, sep-dic para el to go.
    """
    global WS_SIN_MESES
    WS_SIN_MESES = {}

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_WS]
    it = ws.iter_rows(min_row=WS_HEADER_ROW, values_only=True)
    header = [clean_trim(c) for c in next(it)]
    idx = {name: i for i, name in enumerate(header) if name}

    por_mes = PERIOD_SPLIT and not PERIOD_MISSING
    if por_mes:
        faltan = [m for m in WS_MONTH_COLS if m not in idx]
        if faltan:
            raise SystemExit(
                f"La hoja '{SHEET_WS}' no tiene las columnas mensuales "
                f"{', '.join(faltan)}, que hacen falta para separar el actual "
                f"del to go.")
        mcols = [idx[m] for m in WS_MONTH_COLS]
    vcol = idx[WS_VALUE_COL]

    rows = []
    for r in it:
        if r[0] is None or clean_trim(r[0]) == "":
            continue
        inv_del = clean_trim(r[idx["Inv / Del"]])
        pl = clean_trim(r[idx["P&L"]])
        if pl in WS_PL_IGNORE:
            continue
        base = {
            "Country": clean_trim(r[idx["Country"]]),
            "Inv/Del": inv_del,
            "Responsable": clean_trim(r[idx["Responsable"]]),
            "SBE": clean_trim(r[idx["SBE"]]),
            "SBE.1": clean_trim(r[idx["SBE.1"]]),
            "SBE.2": clean_trim(r[idx["SBE.2"]]),
            "SBE.3": clean_trim(r[idx["SBE.3"]]),
            "P&L": pl,
            "Unit": clean_trim(r[idx["Unit"]]),
        }
        if not por_mes:
            rows.append(dict(base, Periodo=P_FULLYEAR,
                             Value=to_num(r[vcol]) * WS_SCALE))
            continue

        meses = [to_num(r[c]) * WS_SCALE for c in mcols]
        anual = to_num(r[vcol]) * WS_SCALE
        # La carga anual se hace en una sola celda: si eso pasa en el FC3, los
        # meses quedan en cero y el bloque compara contra nada. Se avisa en vez
        # de dar por bueno un cero.
        if not any(meses) and anual:
            k = (base["Country"], base["SBE.1"], base["SBE.3"], pl)
            WS_SIN_MESES[k] = WS_SIN_MESES.get(k, 0.0) + anual
        for p in periods_in_play():
            val = sum(v for i, v in enumerate(meses, start=1)
                      if period_of_month(i) == p)
            rows.append(dict(base, Periodo=p, Value=val))

    wb.close()
    return rows


def check_ws_scale(recs):
    """Verifica el factor `ws_scale` del ciclo contra el resultado del cruce.

    Devuelve la relacion tipica query/hoja de las lineas donde los dos lados
    tienen dato. Si la hoja esta bien escalada esa relacion ronda 1: los
    desvios reales mueven cada linea, pero no la mediana. Si da 1000 o 0,001,
    el libro cambio de escala y el cruce entero seria falso -> el que llama
    avisa. La mediana y no el promedio, justamente para que un par de lineas
    muy desviadas no la corran.
    """
    global WS_SCALE_MEDIAN
    rat = sorted(r["query"] / r["worksheet"] for r in recs
                 if abs(r["query"]) > WS_SCALE_FLOOR
                 and abs(r["worksheet"]) > WS_SCALE_FLOOR)
    if not rat:
        WS_SCALE_MEDIAN = None
    else:
        n = len(rat)
        WS_SCALE_MEDIAN = (rat[n // 2] if n % 2
                           else (rat[n // 2 - 1] + rat[n // 2]) / 2)
    return WS_SCALE_MEDIAN


def ws_scale_sospechosa():
    """True si la relacion medida no se parece a 1: la escala esta mal."""
    return (WS_SCALE_MEDIAN is not None
            and not WS_SCALE_MIN <= WS_SCALE_MEDIAN <= WS_SCALE_MAX)


# ============================================================
# SECCION 8 - RECONCILIACION
# ============================================================

GRAINS = {
    "sbe3": ("Country", "SBE.1", "SBE.2", "SBE.3", "P&L"),
    "sbe2": ("Country", "SBE.1", "SBE.2", "P&L"),
    "sbe1": ("Country", "SBE.1", "P&L"),
    "country": ("Country", "P&L"),
}
# La hoja de trabajo tiene 'LT' y 'La Tijereta' como SBE.3 distintos en Corn.
SBE3_ALIASES = {"LT": "La Tijereta"}


def grain_fields(grain):
    """Campos de la clave de cruce. El periodo va primero: es el corte de
    mayor nivel y ordena el reporte por bloque antes que por pais."""
    if PERIOD_SPLIT and not PERIOD_MISSING:
        return ("Periodo",) + GRAINS[grain]
    return GRAINS[grain]


def key_of(row, grain, pl_level):
    out = []
    for f in grain_fields(grain):
        v = row.get(f, "")
        if f == "SBE.3":
            v = SBE3_ALIASES.get(v, v)
        elif f == "P&L" and pl_level == "group":
            v = PL_GROUP.get(v, v)
        out.append(v)
    return tuple(out)


def drop_gly_volume(rows):
    """Saca del cruce el volumen de GLY que no se puede convertir a Regs.

    Sin tabla de factores se van todas las lineas. Con tabla, solo las de los
    SKU que no figuran en ella: convertir con un factor inventado seria peor
    que no comparar, porque el desvio resultante no querria decir nada.
    """
    def es_gly_vol(r):
        return r["P&L"] == "Volume" and r["SBE.1"] == "GLY"

    if not GLY_FACTORS:
        return [r for r in rows if not es_gly_vol(r)]
    if not GLY_SIN_FACTOR:
        return rows

    # Se excluye por clave de cruce, no por SKU: la hoja de trabajo no tiene
    # SKU, y sacar solo el lado de la query dejaria la linea de la hoja sola,
    # con un desvio inventado del 100%.
    tocadas = {(v["Country"], v["SBE.3"]) for v in GLY_SIN_FACTOR.values()}
    return [r for r in rows
            if not (es_gly_vol(r) and (r["Country"], r["SBE.3"]) in tocadas)]


def collapse_gly_sin_marca(qrows, wrows):
    """Colapsa la marca donde la query trae glifosato generico.

    La query carga esos importes contra un material sin marca; la hoja de
    trabajo los reparte entre Roundup y La Tijereta con un criterio de negocio
    que el dato no expone. Repartirlos aca con un porcentaje inventado daria
    dos desvios falsos que se cancelan entre si -exactamente el error que este
    check tiene que evitar-, asi que se comparan a nivel GLY, que es el maximo
    detalle en el que los dos lados dicen lo mismo.

    Devuelve (qrows, wrows, claves colapsadas).
    """
    tocadas = {(r["Country"], r["SBE.1"], r["P&L"]) for r in qrows
               if r["SBE.3"] == GLY_SIN_MARCA}
    if not tocadas:
        return qrows, wrows, []

    def collapse(rows):
        # Se colapsan las dos puntas: si solo se tocara la query, la marca de
        # la hoja quedaria sola y volveria el desvio inventado.
        return [dict(r, **{"SBE.2": "Todas las marcas", "SBE.3": GLY_ROLLUP})
                if (r["Country"], r["SBE.1"], r["P&L"]) in tocadas else r
                for r in rows]

    return collapse(qrows), collapse(wrows), sorted(tocadas)


def check_entities(qrows):
    """Filas cargadas en una entidad legal distinta de la esperada."""
    agg = defaultdict(float)
    for r in qrows:
        if r["P&L"] == "Volume":
            continue
        rules = ENTITY_RULES.get(r["Country"])
        if not rules:
            continue
        expected = rules.get(r["SBE.1"])
        if expected and expected != r["CoCd"]:
            agg[(r["Country"], r["SBE.1"], r["CoCd"], expected,
                 r["P&L"])] += r["Value"]
    return [{"key": k, "value": v} for k, v in sorted(agg.items()) if v]


def reconcile(qrows, wrows, grain, tol, tol_pct, pl_level):
    q = defaultdict(float)
    for r in qrows:
        q[key_of(r, grain, pl_level)] += r["Value"]

    w = defaultdict(float)
    resp = {}
    for r in wrows:
        if r["Inv/Del"] not in WS_INV_DEL_KEEP:
            continue
        k = key_of(r, grain, pl_level)
        w[k] += r["Value"]
        if r["Responsable"]:
            resp.setdefault(k, set()).add(r["Responsable"])

    out = []
    for k in sorted(set(q) | set(w)):
        qv, wv = q.get(k, 0.0), w.get(k, 0.0)
        delta = qv - wv
        base = max(abs(qv), abs(wv))
        pct = (delta / base * 100.0) if base else 0.0
        if abs(delta) <= tol or (tol_pct and base and abs(pct) <= tol_pct):
            status = "OK"
        elif k not in w:
            status = "SOLO EN QUERY"
        elif k not in q:
            status = "FALTA EN QUERY"
        else:
            status = "DESVIO"
        out.append({
            "key": k, "query": qv, "worksheet": wv, "delta": delta,
            "pct": pct, "status": status,
            "resp": ", ".join(sorted(resp.get(k, []))),
        })
    return out


def fullyear_recs(recs, tol, tol_pct=0.0):
    """Vista derivada: suma los bloques ya cruzados, clave por clave.

    No es un cruce nuevo -- el actual y el to go ya se compararon cada uno
    contra su parte de la hoja; esto los suma para poder leer el ano completo,
    que es como el negocio mira el forecast.

    El estado se recalcula sobre la suma con la misma tolerancia, y ahi esta la
    trampa: un desvio de mas en el actual y uno de menos en el to go se
    cancelan, y la linea cierra en el ano estando mal en los dos bloques. Por
    eso cada registro se lleva `nbad` (bloques que no cerraron): el tablero
    cuenta cuantas lineas cierran solo por compensacion y lo avisa.
    """
    if not PERIOD_SPLIT:
        return []

    agg = {}
    for r in recs:
        k = (P_FULLYEAR,) + tuple(r["key"][1:])
        a = agg.get(k)
        if a is None:
            a = agg[k] = {"query": 0.0, "worksheet": 0.0, "en_q": False,
                          "en_w": False, "resp": set(), "nbad": 0}
        a["query"] += r["query"]
        a["worksheet"] += r["worksheet"]
        # De que lado existe la clave lo dice el estado de cada bloque: una
        # clave que falta en la query en los dos sigue faltando en el ano.
        a["en_q"] = a["en_q"] or r["status"] != "FALTA EN QUERY"
        a["en_w"] = a["en_w"] or r["status"] != "SOLO EN QUERY"
        a["nbad"] += 1 if r["status"] != "OK" else 0
        if r["resp"]:
            a["resp"].update(r["resp"].split(", "))

    out = []
    for k in sorted(agg):
        a = agg[k]
        qv, wv = a["query"], a["worksheet"]
        delta = qv - wv
        base = max(abs(qv), abs(wv))
        pct = (delta / base * 100.0) if base else 0.0
        if abs(delta) <= tol or (tol_pct and base and abs(pct) <= tol_pct):
            status = "OK"
        elif not a["en_w"]:
            status = "SOLO EN QUERY"
        elif not a["en_q"]:
            status = "FALTA EN QUERY"
        else:
            status = "DESVIO"
        out.append({
            "key": k, "query": qv, "worksheet": wv, "delta": delta,
            "pct": pct, "status": status,
            "resp": ", ".join(sorted(a["resp"])), "nbad": a["nbad"],
        })
    return out


# ============================================================
# SECCION 9 - OUTPUT
# ============================================================

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(color="FFFFFF", bold=True)
FILLS = {
    "OK": PatternFill("solid", fgColor="E2EFDA"),
    "DESVIO": PatternFill("solid", fgColor="FCE4D6"),
    "SOLO EN QUERY": PatternFill("solid", fgColor="FFF2CC"),
    "FALTA EN QUERY": PatternFill("solid", fgColor="F8CBAD"),
}


def write_sheet(ws, header, rows, numfmt_from=None):
    ws.append(header)
    for c in range(1, len(header) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for r in rows:
        ws.append(r)
    if numfmt_from:
        for row in ws.iter_rows(min_row=2, min_col=numfmt_from,
                                max_col=len(header)):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
    for c in range(1, len(header) + 1):
        width = max([len(str(header[c - 1]))] +
                    [len(str(r[c - 1])) for r in rows[:400]] or [10])
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 2, 10), 42)
    ws.freeze_panes = "A2"


def build_report(recs, grain, qrows, unmapped, path, tol, ent):
    wb = openpyxl.Workbook()
    fields = list(grain_fields(grain))

    # --- Desvios ---
    ws = wb.active
    ws.title = "Desvios"
    bad = [r for r in recs if r["status"] != "OK"]
    bad.sort(key=lambda r: -abs(r["delta"]))
    write_sheet(
        ws,
        fields + ["Query", "Hoja de trabajo", "Desvio", "Desvio %",
                  "Estado", "Responsable"],
        [list(r["key"]) + [round(r["query"], 2), round(r["worksheet"], 2),
                           round(r["delta"], 2), round(r["pct"], 2),
                           r["status"], r["resp"]] for r in bad],
        numfmt_from=len(fields) + 1,
    )
    scol = len(fields) + 5
    for i, r in enumerate(bad, start=2):
        ws.cell(row=i, column=scol).fill = FILLS.get(r["status"])

    # --- Reconciliacion completa ---
    ws2 = wb.create_sheet("Reconciliacion")
    write_sheet(
        ws2,
        fields + ["Query", "Hoja de trabajo", "Desvio", "Desvio %",
                  "Estado", "Responsable"],
        [list(r["key"]) + [round(r["query"], 2), round(r["worksheet"], 2),
                           round(r["delta"], 2), round(r["pct"], 2),
                           r["status"], r["resp"]] for r in recs],
        numfmt_from=len(fields) + 1,
    )

    # --- Resumen por pais ---
    # Las columnas de la clave se ubican por nombre: con apertura de periodo
    # la primera ya no es Country.
    pos = {f: i for i, f in enumerate(fields)}
    cut = [f for f in ("Periodo", "Country", "P&L") if f in pos]
    agg = defaultdict(lambda: [0.0, 0.0, 0])
    for r in recs:
        a = agg[tuple(r["key"][pos[f]] for f in cut)]
        a[0] += r["query"]
        a[1] += r["worksheet"]
        a[2] += 1 if r["status"] != "OK" else 0
    write_sheet(
        wb.create_sheet("Resumen"),
        cut + ["Query", "Hoja de trabajo", "Desvio", "Lineas con desvio"],
        [list(k) + [round(v[0], 2), round(v[1], 2), round(v[0] - v[1], 2), v[2]]
         for k, v in sorted(agg.items())],
        numfmt_from=len(cut) + 1,
    )

    # --- Total del cluster por bloque ---
    # El numero que se mira primero. Van los dos bloques y ademas el Full Year,
    # que es la suma de los dos: el ano es lo que se reporta hacia arriba, pero
    # si el actual cierra de mas y el to go de menos el total anual da bien
    # igual y tapa las dos puntas. Por eso el ano nunca va solo.
    wst = wb.create_sheet("Total Conosur")
    tot = defaultdict(lambda: [0.0, 0.0, 0, 0])
    for r in recs + fullyear_recs(recs, tol):
        pl = r["key"][pos["P&L"]]
        if pl == "Volume":
            continue   # unidades: no se suman con los importes
        a = tot[(r["key"][pos["Periodo"]] if "Periodo" in pos else P_FULLYEAR, pl)]
        a[0] += r["query"]
        a[1] += r["worksheet"]
        a[2] += 1 if r["status"] != "OK" else 0
        a[3] += 1
    orden = {p: i for i, p in enumerate(periods_shown())}
    write_sheet(
        wst,
        ["Periodo", "P&L", "Query", "Hoja de trabajo", "Desvio", "Desvio %",
         "Lineas con desvio", "Lineas"],
        # El % va sobre el mayor de los dos lados, igual que en el resto del
        # reporte: con la hoja en cero, dividir por ella daria 0% y se leeria
        # como que cierra.
        [[period_label(p), pl, round(v[0], 2), round(v[1], 2),
          round(v[0] - v[1], 2),
          round((v[0] - v[1]) / max(abs(v[0]), abs(v[1])) * 100, 2)
          if max(abs(v[0]), abs(v[1])) else 0.0,
          v[2], v[3]]
         for (p, pl), v in sorted(tot.items(),
                                  key=lambda kv: (orden.get(kv[0][0], 9),
                                                  kv[0][1]))],
        numfmt_from=3,
    )

    # --- Control de entidad legal ---
    wse = wb.create_sheet("Entidad legal")
    write_sheet(
        wse,
        ["Country", "SBE.1", "Cargado en", "Entidad cargada",
         "Deberia ser", "Entidad esperada", "P&L", "Valor"],
        [[k[0], k[1], k[2], COCD_NAME.get(k[2], ""), k[3],
          COCD_NAME.get(k[3], ""), k[4], round(r["value"], 2)]
         for r in ent for k in [r["key"]]],
        numfmt_from=8,
    )
    if not ent:
        wse.append(["Sin desvios de entidad."])

    # --- Query normalizada (trazabilidad SKU -> SBE.3) ---
    ws4 = wb.create_sheet("Query normalizada")
    write_sheet(
        ws4,
        ["Periodo", "Country", "CoCd", "SBE", "SBE.1", "SBE.2", "SBE.3",
         "P&L", "SKU", "Material", "CV6", "Franchise", "Valor"],
        [[r["Periodo"], r["Country"], r["CoCd"], r["SBE"], r["SBE.1"],
          r["SBE.2"], r["SBE.3"], r["P&L"], r["SKU"], r["Material"],
          r["CV6"], r["Franchise"], round(r["Value"], 4)] for r in qrows],
        numfmt_from=13,
    )

    # --- Parametros ---
    ws5 = wb.create_sheet("Parametros")
    params = [
        ["Generado", dt.datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Ciclo", f"{CYCLE['label']} ({CYCLE['book']})"],
        ["Apertura de periodo",
            " / ".join(period_label(p) for p in periods_in_play())],
        ["Columna de mes en la query",
            PERIOD_COL_NAME or ("NO VIENE -> se compara el anual sin abrir "
                                "actual/to go" if PERIOD_SPLIT else "no aplica")],
        ["Columna hoja de trabajo",
            f"suma de {WS_MONTH_COLS[0]}..{WS_MONTH_COLS[-1]} por bloque"
            if PERIOD_SPLIT and not PERIOD_MISSING else WS_VALUE_COL],
        ["Lineas de la hoja sin apertura mensual",
            len(WS_SIN_MESES) or "-"],
        ["Granularidad", grain],
        # Este Excel queda en la escala de origen (miles de USD) para cotejar
        # celda a celda contra la hoja de trabajo. El tablero HTML muestra lo
        # mismo en millones de USD, que es la magnitud real del negocio.
        ["Unidad de los importes", "miles de USD (x1000 = USD reales)"],
        ["Unidad del volumen", "unidades"],
        ["Tolerancia (miles USD)", tol],
        ["Inv/Del comparados", ", ".join(sorted(WS_INV_DEL_KEEP))],
        ["Escala plata query", SCALE_MONEY],
        ["Escala volumen query", SCALE_VOLUME],
        ["Escala hoja de trabajo", WS_SCALE],
        ["Relacion tipica query/hoja",
            "-" if WS_SCALE_MEDIAN is None else round(WS_SCALE_MEDIAN, 4)],
        ["P&L ignoradas", ", ".join(sorted(WS_PL_IGNORE))],
        ["Volumen GLY", f"convertido a Regs con {len(GLY_FACTORS)} factores "
                        f"por SKU desde {GLY_FACTOR_FILE}"
            if GLY_FACTORS else "excluido (falta la tabla de factores)"],
        ["Volumen GLY sin factor",
            ", ".join(sorted(GLY_SIN_FACTOR)) or "-"],
        ["Company codes sin mapear", ", ".join(sorted(unmapped)) or "-"],
    ]
    params += [[f"CoCd {k}", v] for k, v in sorted(COCD_TO_COUNTRY.items())]
    write_sheet(ws5, ["Parametro", "Valor"], params)

    wb.save(path)


# ============================================================
# SECCION 10 - MAIN
# ============================================================

def add_common_args(ap):
    """Flags que comparten el check y el tablero."""
    ap.add_argument("--cycle", default=DEFAULT_CYCLE, choices=sorted(CYCLES),
                    help="ciclo de planificacion (define libro y apertura)")
    ap.add_argument("--book", default=None,
                    help="libro a leer; por defecto, el del ciclo")
    ap.add_argument("--actual-thru", type=int, default=None, metavar="MES",
                    help="ultimo mes cerrado (1-12); de ahi en mas es to go")
    ap.add_argument("--grain", default="sbe3", choices=sorted(GRAINS))
    ap.add_argument("--tol", type=float, default=DEFAULT_TOL)
    ap.add_argument("--pl-level", default="group", choices=("group", "sub"),
                    help="group = agrupa las deducciones en 'Sales Adj' "
                         "(como el script M); sub = linea por linea")
    return ap


def main():
    ap = add_common_args(argparse.ArgumentParser(description="Check de carga IBF"))
    ap.add_argument("--out", default=None)
    ap.add_argument("--tol-pct", type=float, default=DEFAULT_TOL_PCT)
    ap.add_argument("--no-volume", action="store_true",
                    help="excluye las lineas de Volume de la comparacion")
    args = ap.parse_args()

    if args.actual_thru is not None and not 1 <= args.actual_thru <= 12:
        raise SystemExit("--actual-thru tiene que ser un mes entre 1 y 12")
    cyc = setup_cycle(args.cycle, args.actual_thru)
    book = args.book or cyc["book"]
    base = os.path.dirname(os.path.abspath(book))
    out = args.out or os.path.join(base, cyc["out_xlsx"])
    print(f"Ciclo        : {cyc['label']}  ({os.path.basename(book)})")

    global EXTRA_OVERRIDES, GLY_FACTORS
    EXTRA_OVERRIDES = load_sku_overrides(os.path.join(base, SKU_OVERRIDE_FILE))
    if EXTRA_OVERRIDES:
        print(f"Overrides SKU : {len(EXTRA_OVERRIDES)} desde {SKU_OVERRIDE_FILE}")

    GLY_FACTORS = load_gly_factors(os.path.join(base, GLY_FACTOR_FILE))
    if GLY_FACTORS:
        print(f"Factores GLY  : {len(GLY_FACTORS)} desde {GLY_FACTOR_FILE}")
    else:
        print(f"Factores GLY  : sin tabla ({GLY_FACTOR_FILE} no encontrado), "
              f"el volumen de GLY queda fuera del cruce")

    # read_query primero: es la que descubre si la query trae el mes, y de eso
    # depende como se lee la hoja de trabajo (por meses o por anual).
    qrows, unmapped = read_query(book)
    wrows = read_worksheet(book)
    ent = check_entities(qrows)

    if args.no_volume:
        qrows = [r for r in qrows if r["P&L"] != "Volume"]
        wrows = [r for r in wrows if r["P&L"] != "Volume"]
    else:
        qrows, wrows = drop_gly_volume(qrows), drop_gly_volume(wrows)

    qrows, wrows, colapsadas = collapse_gly_sin_marca(qrows, wrows)

    recs = reconcile(qrows, wrows, args.grain, args.tol, args.tol_pct,
                     args.pl_level)
    check_ws_scale(recs)
    build_report(recs, args.grain, qrows, unmapped, out, args.tol, ent)

    bad = [r for r in recs if r["status"] != "OK"]
    print(f"Query        : {len(qrows)} filas")
    print(f"Hoja trabajo : {len(wrows)} filas")
    print(f"Granularidad : {args.grain}  |  tolerancia {args.tol}")
    print(f"Combinaciones: {len(recs)}   con desvio: {len(bad)}")

    if PERIOD_SPLIT and not PERIOD_MISSING:
        print(f"Apertura     : {' | '.join(period_label(p) for p in periods_in_play())}"
              f"  (columna '{PERIOD_COL_NAME}' de la query)")
    if WS_SCALE != 1.0:
        print(f"Escala hoja  : x{WS_SCALE:,.0f} (la hoja carga en millones; "
              f"relacion tipica query/hoja despues de escalar: "
              f"{WS_SCALE_MEDIAN:,.3f})")
    if ws_scale_sospechosa():
        print()
        print(f"!! La hoja de trabajo parece estar en otra escala: aun aplicando")
        print(f"   el factor del ciclo (x{WS_SCALE:,.0f}), la relacion tipica")
        print(f"   query/hoja da {WS_SCALE_MEDIAN:,.4f} y deberia rondar 1. Asi el")
        print(f"   cruce entero es falso. Revisa en que unidad quedo cargada la")
        print(f"   hoja y ajusta 'ws_scale' del ciclo '{args.cycle}' en CYCLES.")
    if PERIOD_MISSING:
        print()
        print("!! La hoja 'query' no trae el mes: no se puede separar el actual")
        print("   del to go y se compara el total del anio contra la suma de los")
        print("   doce meses de la hoja. Para abrirlo, refresca Analysis con")
        print("   'Calendar Year/Month' como caracteristica de fila.")
    if PERIOD_UNPARSED:
        print(f"!! Mes ilegible en {sum(PERIOD_UNPARSED.values())} fila(s) de la "
              f"query; quedaron FUERA del cruce: "
              f"{', '.join(repr(k) for k in sorted(PERIOD_UNPARSED)[:6])}")
    if PERIOD_OTHER_YEAR:
        tot = sum(PERIOD_OTHER_YEAR.values())
        print(f"!! La query trae {len(PERIOD_OTHER_YEAR)} periodo(s) de otro anio "
              f"que {CYCLE['year']} ({tot:,.2f}); quedaron fuera del cruce: "
              f"{', '.join(f'{a}.{m:02d}' for a, m in sorted(PERIOD_OTHER_YEAR))}")
    if WS_SIN_MESES:
        tot = sum(WS_SIN_MESES.values())
        print(f"!! Hoja de trabajo: {len(WS_SIN_MESES)} linea(s) con anual "
              f"cargado y los doce meses vacios ({tot:,.2f}). Sin apertura "
              f"mensual no entran en ningun bloque:")
        for k, v in sorted(WS_SIN_MESES.items(), key=lambda kv: -abs(kv[1]))[:10]:
            print(f"   {k[0]:4} {k[1]:<8} {k[2]:<14} {k[3]:<12} {v:>13,.2f}")
    if not args.no_volume and not GLY_FACTORS:
        print("Nota         : volumen de GLY excluido (falta la tabla de factores)")
    if not args.no_volume and GLY_SIN_FACTOR:
        print(f"!! Volumen GLY: {len(GLY_SIN_FACTOR)} SKU sin factor en "
              f"{GLY_FACTOR_FILE}. Esas lineas quedan FUERA del cruce:")
        for sku, v in sorted(GLY_SIN_FACTOR.items(),
                             key=lambda kv: -abs(kv[1]["Value"])):
            print(f"   {sku}  {v['Country']:4} {v['SBE.3']:<12} "
                  f"{v['Value']:>13,.0f}  {v['Material'][:44]}")
    if colapsadas:
        print(f"Nota         : {len(colapsadas)} combinacion(es) con glifosato "
              f"generico -> se compara a nivel '{GLY_ROLLUP}', sin abrir marca:")
        for pais, sbe1, pl in colapsadas:
            print(f"   {pais} {sbe1} {pl}")
    if unmapped:
        print(f"!! Company codes sin mapear: {sorted(unmapped)}")
    if ent:
        print(f"!! Entidad legal: {len(ent)} combinaciones cargadas en la "
              f"sociedad equivocada (ver el Excel)")
        for r in ent[:10]:
            k = r["key"]
            print(f"   {k[0]} {k[1]} {k[4]}: cargado en {k[2]}, "
                  f"deberia ser {k[3]}  ->  {r['value']:,.2f}")
    print()
    if bad:
        w = max(len(" | ".join(r["key"])) for r in bad)
        print(f"{'Clave'.ljust(w)} | {'Query':>14} | {'Hoja':>14} | "
              f"{'Desvio':>12} | Estado")
        print("-" * (w + 62))
        for r in sorted(bad, key=lambda r: -abs(r["delta"]))[:40]:
            print(f"{' | '.join(r['key']).ljust(w)} | {r['query']:>14,.2f} | "
                  f"{r['worksheet']:>14,.2f} | {r['delta']:>12,.2f} | "
                  f"{r['status']}{'  -> ' + r['resp'] if r['resp'] else ''}")
        if len(bad) > 40:
            print(f"... y {len(bad) - 40} mas (ver el Excel)")
    else:
        print("Sin desvios por encima de la tolerancia.")
    print(f"\nReporte: {out}")


if __name__ == "__main__":
    main()
