"""
Check SBP 2027 - IBF
====================

Compara lo cargado en SAP BW (hoja `query`) contra la hoja de trabajo
(hoja `PreFC Y1`) de Checksibf.xlsx, normalizando ambos lados a la misma
granularidad: Country x SBE.1 x SBE.2 x SBE.3 x P&L.

Las reglas de clasificacion replican el script de Power Query de BASE D1,
adaptadas a las columnas que trae la extraccion de BW (que no tiene
`CV Brand Family` ni `Franchise` a nivel detalle).

Uso:
    py check_sbp.py                      # usa Checksibf.xlsx, escribe Check_SBP2027.xlsx
    py check_sbp.py --tol 5              # tolerancia 5 (miles USD)
    py check_sbp.py --grain sbe1         # compara a nivel SBE.1 en vez de SBE.3
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

BOOK = "Checksibf.xlsx"
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
    "0164": "CHI",
    "0197": "PER",
    "1386": "BOL",
    "0194": "UGY",
    "0916": "PGY",  # Bayer Paraguay
    "2663": "PGY",  # Monsanto Paraguay
}

# Nombre de la entidad legal, para el reporte.
COCD_NAME = {
    "0916": "Bayer Paraguay",
    "2663": "Monsanto Paraguay",
}

# Paises con mas de una entidad legal: que negocio se factura por cual.
# Se reporta como control aparte; una fila cargada en la entidad equivocada
# no se ve en el cruce por pais pero si es un error de carga.
ENTITY_RULES = {
    "PGY": {
        "Corn": "2663", "SOY": "2663", "GLY": "2663", "OTHERS": "2663",
        "CP": "0916",
    },
}

# Solo se comparan las filas de la hoja de trabajo con este flag.
# La query refleja delivery; las filas `Inv` son de invoice y no se cargan.
WS_INV_DEL_KEEP = {"Del"}

# Columna de la hoja de trabajo que contiene el valor anual.
WS_VALUE_COL = "Full Year"

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

def read_query(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_QUERY]
    rows = []
    unmapped_cocd = set()
    for r in ws.iter_rows(min_row=QUERY_FIRST_DATA_ROW, values_only=True):
        pl_raw = clean_trim(r[0])
        if not pl_raw:
            continue
        cocd = clean_trim(r[1])
        cv6 = clean_trim(r[2])
        franchise = clean_trim(r[3])
        sku = clean_trim(r[4])
        mat = clean_trim(r[5])
        val = to_num(r[6])

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
            "Country": country, "CoCd": cocd, "SBE": sbe, "SBE.1": sbe1,
            "SBE.2": sbe2, "SBE.3": sbe3, "P&L": pl, "SKU": sku,
            "Material": mat, "CV6": cv6, "Franchise": franchise,
            "Value": val,
        })
    wb.close()
    return rows, unmapped_cocd


def read_worksheet(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_WS]
    it = ws.iter_rows(min_row=WS_HEADER_ROW, values_only=True)
    header = [clean_trim(c) for c in next(it)]
    idx = {name: i for i, name in enumerate(header) if name}
    vcol = idx[WS_VALUE_COL]

    rows = []
    for r in it:
        if r[0] is None or clean_trim(r[0]) == "":
            continue
        inv_del = clean_trim(r[idx["Inv / Del"]])
        pl = clean_trim(r[idx["P&L"]])
        if pl in WS_PL_IGNORE:
            continue
        rows.append({
            "Country": clean_trim(r[idx["Country"]]),
            "Inv/Del": inv_del,
            "Responsable": clean_trim(r[idx["Responsable"]]),
            "SBE": clean_trim(r[idx["SBE"]]),
            "SBE.1": clean_trim(r[idx["SBE.1"]]),
            "SBE.2": clean_trim(r[idx["SBE.2"]]),
            "SBE.3": clean_trim(r[idx["SBE.3"]]),
            "P&L": pl,
            "Unit": clean_trim(r[idx["Unit"]]),
            "Value": to_num(r[vcol]),
        })
    wb.close()
    return rows


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


def key_of(row, grain, pl_level):
    out = []
    for f in GRAINS[grain]:
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
    fields = list(GRAINS[grain])

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
    ws3 = wb.create_sheet("Resumen")
    agg = defaultdict(lambda: [0.0, 0.0, 0])
    for r in recs:
        country, pl = r["key"][0], r["key"][-1]
        a = agg[(country, pl)]
        a[0] += r["query"]
        a[1] += r["worksheet"]
        a[2] += 1 if r["status"] != "OK" else 0
    write_sheet(
        ws3,
        ["Country", "P&L", "Query", "Hoja de trabajo", "Desvio",
         "Lineas con desvio"],
        [[c, p, round(v[0], 2), round(v[1], 2), round(v[0] - v[1], 2), v[2]]
         for (c, p), v in sorted(agg.items())],
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
        ["Country", "CoCd", "SBE", "SBE.1", "SBE.2", "SBE.3", "P&L",
         "SKU", "Material", "CV6", "Franchise", "Valor"],
        [[r["Country"], r["CoCd"], r["SBE"], r["SBE.1"], r["SBE.2"],
          r["SBE.3"], r["P&L"], r["SKU"], r["Material"], r["CV6"],
          r["Franchise"], round(r["Value"], 4)] for r in qrows],
        numfmt_from=12,
    )

    # --- Parametros ---
    ws5 = wb.create_sheet("Parametros")
    params = [
        ["Generado", dt.datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Granularidad", grain],
        # Este Excel queda en la escala de origen (miles de USD) para cotejar
        # celda a celda contra la hoja de trabajo. El tablero HTML muestra lo
        # mismo en millones de USD, que es la magnitud real del negocio.
        ["Unidad de los importes", "miles de USD (x1000 = USD reales)"],
        ["Unidad del volumen", "unidades"],
        ["Tolerancia (miles USD)", tol],
        ["Inv/Del comparados", ", ".join(sorted(WS_INV_DEL_KEEP))],
        ["Columna hoja de trabajo", WS_VALUE_COL],
        ["Escala plata query", SCALE_MONEY],
        ["Escala volumen query", SCALE_VOLUME],
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

def main():
    ap = argparse.ArgumentParser(description="Check SBP 2027 - IBF")
    ap.add_argument("--book", default=BOOK)
    ap.add_argument("--out", default=None)
    ap.add_argument("--grain", default="sbe3", choices=sorted(GRAINS))
    ap.add_argument("--tol", type=float, default=DEFAULT_TOL)
    ap.add_argument("--tol-pct", type=float, default=DEFAULT_TOL_PCT)
    ap.add_argument("--pl-level", default="group", choices=("group", "sub"),
                    help="group = agrupa las deducciones en 'Sales Adj' "
                         "(como el script M); sub = linea por linea")
    ap.add_argument("--no-volume", action="store_true",
                    help="excluye las lineas de Volume de la comparacion")
    args = ap.parse_args()

    base = os.path.dirname(os.path.abspath(args.book))
    out = args.out or os.path.join(base, "Check_SBP2027.xlsx")

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

    qrows, unmapped = read_query(args.book)
    wrows = read_worksheet(args.book)
    ent = check_entities(qrows)

    if args.no_volume:
        qrows = [r for r in qrows if r["P&L"] != "Volume"]
        wrows = [r for r in wrows if r["P&L"] != "Volume"]
    else:
        qrows, wrows = drop_gly_volume(qrows), drop_gly_volume(wrows)

    qrows, wrows, colapsadas = collapse_gly_sin_marca(qrows, wrows)

    recs = reconcile(qrows, wrows, args.grain, args.tol, args.tol_pct,
                     args.pl_level)
    build_report(recs, args.grain, qrows, unmapped, out, args.tol, ent)

    bad = [r for r in recs if r["status"] != "OK"]
    print(f"Query        : {len(qrows)} filas")
    print(f"Hoja trabajo : {len(wrows)} filas")
    print(f"Granularidad : {args.grain}  |  tolerancia {args.tol}")
    print(f"Combinaciones: {len(recs)}   con desvio: {len(bad)}")
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
