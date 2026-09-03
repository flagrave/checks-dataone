"""
Tablero de control SBP 2027 - IBF
=================================

Genera un HTML autocontenido (sin dependencias, sin servidor) con el cruce
entre lo cargado en SAP BW y la hoja de trabajo, para compartir con el equipo.

    py dashboard.py
    py dashboard.py --out Tablero.html --grain sbe3
"""

import argparse
import datetime as dt
import html
import json
import os

import check_sbp as chk


# ============================================================
# DATOS
# ============================================================

def build_rows(book, grain, tol, pl_level):
    base = os.path.dirname(os.path.abspath(book))
    chk.EXTRA_OVERRIDES = chk.load_sku_overrides(
        os.path.join(base, chk.SKU_OVERRIDE_FILE))
    chk.GLY_FACTORS = chk.load_gly_factors(
        os.path.join(base, chk.GLY_FACTOR_FILE))

    qrows, unmapped = chk.read_query(book)
    wrows = chk.read_worksheet(book)
    ent = chk.check_entities(qrows)

    qrows, wrows = chk.drop_gly_volume(qrows), chk.drop_gly_volume(wrows)
    qrows, wrows, colapsadas = chk.collapse_gly_sin_marca(qrows, wrows)

    recs = chk.reconcile(qrows, wrows, grain, tol, 0.0, pl_level)
    fields = list(chk.GRAINS[grain])

    rows = []
    for r in recs:
        d = dict(zip(fields, r["key"]))
        rows.append({
            "country": d.get("Country", ""),
            "sbe1": d.get("SBE.1", ""),
            "sbe2": d.get("SBE.2", ""),
            "sbe3": d.get("SBE.3", ""),
            "pl": d.get("P&L", ""),
            "q": round(r["query"], 2),
            "w": round(r["worksheet"], 2),
            "d": round(r["delta"], 2),
            "s": r["status"],
            "r": r["resp"],
        })

    meta = {
        "generado": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "tol": tol,
        "grain": grain,
        "plLevel": pl_level,
        "glyVol": not chk.GLY_FACTORS,
        "glySinFactor": [
            {"sku": k, "country": v["Country"], "sbe3": v["SBE.3"],
             "mat": v["Material"]}
            for k, v in sorted(chk.GLY_SIN_FACTOR.items(),
                               key=lambda kv: -abs(kv[1]["Value"]))],
        "glyColapsado": [list(k) for k in colapsadas],
        "unmapped": sorted(unmapped),
        "entidad": [{"k": list(e["key"]), "v": round(e["value"], 2)}
                    for e in ent],
        "entidadNombres": chk.COCD_NAME,
    }
    return rows, meta


# ============================================================
# HTML
# ============================================================

TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tablero SBP 2027 &middot; Checks IBF</title>
<style>
  .viz-root {
    color-scheme: light;
    --surface-1:      #fcfcfb;
    --plane:          #f9f9f7;
    --text-primary:   #0b0b0b;
    --text-secondary: #52514e;
    --text-muted:     #898781;
    --grid:           #e1e0d9;
    --axis:           #c3c2b7;
    --border:         rgba(11,11,11,0.10);
    --pos:            #2a78d6;   /* de mas en sistema */
    --neg:            #e34948;   /* falta en sistema  */
    --neutral:        #f0efec;
    --good:           #0ca30c;
    --warning:        #fab219;
    --serious:        #ec835a;
    --critical:       #d03b3b;
  }
  @media (prefers-color-scheme: dark) {
    :root:where(:not([data-theme="light"])) .viz-root {
      color-scheme: dark;
      --surface-1:      #1a1a19;
      --plane:          #0d0d0d;
      --text-primary:   #ffffff;
      --text-secondary: #c3c2b7;
      --text-muted:     #898781;
      --grid:           #2c2c2a;
      --axis:           #383835;
      --border:         rgba(255,255,255,0.10);
      --pos:            #3987e5;
      --neg:            #e66767;
      --neutral:        #383835;
    }
  }
  :root[data-theme="dark"] .viz-root {
    color-scheme: dark;
    --surface-1:      #1a1a19;
    --plane:          #0d0d0d;
    --text-primary:   #ffffff;
    --text-secondary: #c3c2b7;
    --text-muted:     #898781;
    --grid:           #2c2c2a;
    --axis:           #383835;
    --border:         rgba(255,255,255,0.10);
    --pos:            #3987e5;
    --neg:            #e66767;
    --neutral:        #383835;
  }

  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    background: var(--plane);
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    color: var(--text-primary);
  }
  .viz-root { background: var(--plane); min-height: 100vh; padding: 28px 32px 56px; }
  .wrap { max-width: 1280px; margin: 0 auto; }

  header.top { display: flex; align-items: flex-start; justify-content: space-between;
               gap: 24px; margin-bottom: 24px; }
  h1 { font-size: 20px; font-weight: 600; margin: 0 0 4px; letter-spacing: -0.01em; }
  .sub { font-size: 13px; color: var(--text-secondary); margin: 0; }
  .btn {
    font: inherit; font-size: 12px; color: var(--text-secondary);
    background: var(--surface-1); border: 1px solid var(--border);
    border-radius: 8px; padding: 7px 12px; cursor: pointer; white-space: nowrap;
  }
  .btn:hover { color: var(--text-primary); }

  /* --- filtros: una sola fila, arriba de todo lo que scopean --- */
  .filters { display: flex; flex-wrap: wrap; gap: 10px 14px; align-items: center;
             padding: 12px 16px; background: var(--surface-1);
             border: 1px solid var(--border); border-radius: 12px; margin-bottom: 20px; }
  .filters label { font-size: 11px; color: var(--text-muted); text-transform: uppercase;
                   letter-spacing: 0.04em; margin-right: 6px; }
  .filters select {
    font: inherit; font-size: 13px; color: var(--text-primary);
    background: var(--plane); border: 1px solid var(--border);
    border-radius: 7px; padding: 5px 9px; cursor: pointer;
  }
  .fgroup { display: flex; align-items: center; }
  .chk { display: flex; align-items: center; gap: 6px; font-size: 13px;
         color: var(--text-secondary); cursor: pointer; }

  /* --- hero + tiles --- */
  .kpis { display: grid; grid-template-columns: 1.5fr repeat(4, 1fr); gap: 14px;
          margin-bottom: 20px; }
  .card { background: var(--surface-1); border: 1px solid var(--border);
          border-radius: 12px; padding: 16px 18px; }
  .tile-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase;
                letter-spacing: 0.04em; margin-bottom: 8px; }
  .hero-val { font-size: 48px; font-weight: 600; line-height: 1.05;
              letter-spacing: -0.02em; }
  .tile-val { font-size: 25px; font-weight: 600; line-height: 1.15;
              letter-spacing: -0.01em; }
  .tile-note { font-size: 12px; color: var(--text-secondary); margin-top: 6px; }
  .unit { font-size: 14px; color: var(--text-muted); font-weight: 400; margin-left: 5px; }
  .hero-val .unit { font-size: 18px; margin-left: 7px; }

  /* --- charts --- */
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 14px; }
  .chart-title { font-size: 14px; font-weight: 600; margin: 0 0 2px; }
  .chart-sub { font-size: 12px; color: var(--text-secondary); margin: 0 0 14px; }
  svg { display: block; width: 100%; height: auto; overflow: visible; }
  .legend { display: flex; gap: 16px; margin-top: 12px; flex-wrap: wrap; }
  .lg { display: flex; align-items: center; gap: 7px; font-size: 12px;
        color: var(--text-secondary); }
  .lg i { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }
  .empty { font-size: 13px; color: var(--text-muted); padding: 28px 0; text-align: center; }

  /* --- tabla --- */
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  th { text-align: left; font-weight: 600; font-size: 11px; color: var(--text-muted);
       text-transform: uppercase; letter-spacing: 0.04em; padding: 8px 10px;
       border-bottom: 1px solid var(--axis); white-space: nowrap;
       position: sticky; top: 0; background: var(--surface-1); z-index: 2; }
  td { padding: 7px 10px; border-bottom: 1px solid var(--grid); color: var(--text-secondary); }
  td.k { color: var(--text-primary); }
  th.num, td.num { text-align: right; font-variant-numeric: tabular-nums; }
  tbody tr:hover td { background: var(--plane); }
  .chip { display: inline-flex; align-items: center; gap: 5px; font-size: 11px;
          font-weight: 600; padding: 2px 8px; border-radius: 999px;
          border: 1px solid var(--border); white-space: nowrap; }
  .chip b { font-size: 12px; line-height: 1; }
  .scroll { max-height: 460px; overflow: auto; }

  /* --- tooltip --- */
  #tip { position: fixed; pointer-events: none; opacity: 0; transition: opacity .09s;
         background: var(--surface-1); border: 1px solid var(--border);
         border-radius: 9px; padding: 9px 11px; font-size: 12px;
         box-shadow: 0 6px 20px rgba(0,0,0,.14); z-index: 50; max-width: 280px; }
  #tip .t-cat { color: var(--text-secondary); font-size: 11px; margin-bottom: 5px; }
  #tip .t-row { display: flex; align-items: baseline; gap: 8px; margin-top: 3px; }
  #tip .t-key { width: 14px; height: 2px; border-radius: 2px; flex: none;
                position: relative; top: -3px; }
  #tip .t-val { font-weight: 600; font-variant-numeric: tabular-nums;
                color: var(--text-primary); }
  #tip .t-lbl { color: var(--text-secondary); font-size: 11px; }

  .foot { font-size: 11.5px; color: var(--text-muted); margin-top: 22px; line-height: 1.7; }
  .warnbox { background: var(--surface-1); border: 1px solid var(--border);
             border-left: 3px solid var(--warning); border-radius: 10px;
             padding: 12px 16px; font-size: 12.5px; color: var(--text-secondary);
             margin-bottom: 14px; }

  @media (max-width: 1000px) {
    .kpis { grid-template-columns: 1fr 1fr; }
    .grid2 { grid-template-columns: 1fr; }
  }
  @media print {
    .viz-root { padding: 0; }
    .btn, .filters { display: none; }
    .card { break-inside: avoid; }
    .scroll { max-height: none; overflow: visible; }
  }
</style>
</head>
<body>
<div class="viz-root"><div class="wrap">

  <header class="top">
    <div>
      <h1>Carga SBP 2027 vs hoja de trabajo</h1>
      <p class="sub" id="subtitle"></p>
    </div>
    <button class="btn" id="theme">Modo oscuro</button>
  </header>

  <div id="warn"></div>

  <div class="filters">
    <div class="fgroup"><label for="f-pl">L&iacute;nea P&amp;L</label>
      <select id="f-pl"></select></div>
    <div class="fgroup"><label for="f-country">Pa&iacute;s</label>
      <select id="f-country"></select></div>
    <div class="fgroup"><label for="f-sbe1">Negocio</label>
      <select id="f-sbe1"></select></div>
    <div class="fgroup"><label for="f-resp">Responsable</label>
      <select id="f-resp"></select></div>
    <label class="chk"><input type="checkbox" id="f-dev"> Solo l&iacute;neas con desv&iacute;o</label>
  </div>

  <div class="kpis">
    <div class="card">
      <div class="tile-label">Impacto total a revisar</div>
      <div class="hero-val" id="k-hero"></div>
      <div class="tile-note" id="k-hero-note"></div>
    </div>
    <div class="card">
      <div class="tile-label">Cargado en sistema</div>
      <div class="tile-val" id="k-q"></div>
      <div class="tile-note">Query SAP BW</div>
    </div>
    <div class="card">
      <div class="tile-label">Hoja de trabajo</div>
      <div class="tile-val" id="k-w"></div>
      <div class="tile-note">PreFC Y1</div>
    </div>
    <div class="card">
      <div class="tile-label">Desv&iacute;o neto</div>
      <div class="tile-val" id="k-d"></div>
      <div class="tile-note" id="k-d-note"></div>
    </div>
    <div class="card">
      <div class="tile-label">L&iacute;neas con desv&iacute;o</div>
      <div class="tile-val" id="k-n"></div>
      <div class="tile-note" id="k-n-note"></div>
    </div>
  </div>

  <div class="grid2">
    <div class="card">
      <p class="chart-title">Desv&iacute;o por pa&iacute;s</p>
      <p class="chart-sub" id="c-pais-sub"></p>
      <div id="c-pais"></div>
      <div class="legend" id="lg-pais"></div>
    </div>
    <div class="card">
      <p class="chart-title">Desv&iacute;o por responsable de carga</p>
      <p class="chart-sub" id="c-resp-sub"></p>
      <div id="c-resp"></div>
      <div class="legend" id="lg-resp"></div>
    </div>
  </div>

  <div class="card" style="margin-bottom:14px">
    <p class="chart-title">Mayores desv&iacute;os por l&iacute;nea</p>
    <p class="chart-sub" id="c-linea-sub"></p>
    <div id="c-linea"></div>
    <div class="legend" id="lg-linea"></div>
  </div>

  <div class="card">
    <p class="chart-title">Detalle</p>
    <p class="chart-sub" id="t-sub"></p>
    <div class="scroll">
      <table>
        <thead><tr>
          <th>Pa&iacute;s</th><th>Negocio</th><th>SBE.2</th><th>SBE.3</th><th>P&amp;L</th>
          <th class="num">Sistema</th><th class="num">Hoja</th><th class="num">Desv&iacute;o</th>
          <th>Estado</th><th>Resp.</th>
        </tr></thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
  </div>

  <p class="foot" id="foot"></p>
</div></div>
<div id="tip" role="tooltip"></div>

<script>
const DATA = __DATA__;
const META = __META__;

const NS = "http://www.w3.org/2000/svg";
const fmt = (n, d) => n.toLocaleString("es-AR",
  {minimumFractionDigits: d ?? 0, maximumFractionDigits: d ?? 0});
const sign = (n) => (n > 0 ? "+" : n < 0 ? "−" : "");

/* ---------- unidades ----------
   La fuente trae los importes divididos por 1000 (miles de USD): es la escala
   con la que se arman los reportes, no la magnitud real del negocio. El
   tablero los lleva a millones de USD. El volumen no se escala: son unidades. */
const K_TO_M = 1e-3;
const UNIT_MONEY = {k: K_TO_M, name: "millones de USD", suf: "M USD", money: true};
const UNIT_QTY   = {k: 1,      name: "unidades",        suf: "un.",   money: false};
const unit = () => (state.pl === "Volume" ? UNIT_QTY : UNIT_MONEY);

// Importe ya escalado -> texto. Decimales segun magnitud, para no perder
// las lineas chicas ni ensuciar los totales.
function amt(v, u, extra) {
  if (!u.money) return fmt(Math.round(v));
  if (extra) return fmt(v, 3);   // vista de tabla: 3 decimales = 1 kUSD, la tolerancia
  const a = Math.abs(v);
  if (a >= 1000) return fmt(v, 0);
  if (a >= 100)  return fmt(v, 1);
  return fmt(v, 2);
}
// Un valor que redondea a cero se muestra "0", nunca "−0,000".
function samt(v, u, extra) {
  const txt = amt(Math.abs(v), u, extra);
  return /[1-9]/.test(txt) ? sign(v) + txt : "0";
}

// Ticks del eje, en la unidad ya escalada. 12,5 no puede leerse "13".
function tickFmt(n, u) {
  const a = Math.abs(n);
  if (a === 0) return "0";
  if (u && u.money)
    return fmt(n, Number.isInteger(n) ? 0 : a >= 10 ? 1 : 2);
  if (a >= 1e6) return (n/1e6).toLocaleString("es-AR",{maximumFractionDigits:1}) + "M";
  if (a >= 1e3) return (n/1e3).toLocaleString("es-AR",{maximumFractionDigits:1}) + "k";
  return fmt(n);
}
const css = (v) => getComputedStyle(document.querySelector(".viz-root"))
  .getPropertyValue(v).trim();

// Valor grande + su unidad al lado, para hero y tiles.
function setVal(id, txt, u) {
  const el = document.getElementById(id);
  el.textContent = txt;
  if (!u) return;
  const s = document.createElement("span");
  s.className = "unit"; s.textContent = u.suf;
  el.appendChild(s);
}

const state = {pl: "Net Sales", country: "Todos", sbe1: "Todos",
               resp: "Todos", dev: false};

/* ---------- filtros ---------- */
function uniq(key) {
  return [...new Set(DATA.map(r => r[key]).filter(Boolean))].sort();
}
function fillSelect(id, opts, val) {
  const el = document.getElementById(id);
  el.textContent = "";
  for (const o of opts) {
    const [v, txt] = Array.isArray(o) ? o : [o, o];
    const op = document.createElement("option");
    op.value = v; op.textContent = txt;
    if (v === val) op.selected = true;
    el.appendChild(op);
  }
}
// "Todas" agrupa solo los importes: sumar unidades de volumen con millones de
// USD no significa nada, asi que Volume se mira siempre por separado.
function filtered() {
  return DATA.filter(r =>
    (state.pl === "Todas" ? r.pl !== "Volume" : r.pl === state.pl) &&
    (state.country === "Todos" || r.country === state.country) &&
    (state.sbe1 === "Todos" || r.sbe1 === state.sbe1) &&
    (state.resp === "Todos" || (r.r || "").split(", ").includes(state.resp)) &&
    (!state.dev || r.s !== "OK"));
}

/* ---------- tooltip ---------- */
const tip = document.getElementById("tip");
function showTip(ev, cat, rows) {
  tip.textContent = "";
  const c = document.createElement("div");
  c.className = "t-cat"; c.textContent = cat; tip.appendChild(c);
  for (const [label, val, color] of rows) {
    const row = document.createElement("div"); row.className = "t-row";
    if (color) {
      const k = document.createElement("span");
      k.className = "t-key"; k.style.background = color; row.appendChild(k);
    }
    const v = document.createElement("span");
    v.className = "t-val"; v.textContent = val; row.appendChild(v);
    const l = document.createElement("span");
    l.className = "t-lbl"; l.textContent = label; row.appendChild(l);
    tip.appendChild(row);
  }
  tip.style.opacity = "1";
  moveTip(ev);
}
function moveTip(ev) {
  const r = tip.getBoundingClientRect();
  let x = ev.clientX + 14, y = ev.clientY + 14;
  if (x + r.width > innerWidth - 8) x = ev.clientX - r.width - 14;
  if (y + r.height > innerHeight - 8) y = ev.clientY - r.height - 14;
  tip.style.left = x + "px"; tip.style.top = y + "px";
}
function hideTip() { tip.style.opacity = "0"; }

/* ---------- barra divergente horizontal ---------- */
// Extremo del dato redondeado 4px, escuadrado contra la linea del cero.
function barPath(x0, x1, y, h, r) {
  const w = Math.abs(x1 - x0);
  r = Math.min(r, w, h / 2);
  if (w < 0.5) return "";
  return x1 >= x0
    ? `M${x0},${y} H${x1-r} a${r},${r} 0 0 1 ${r},${r} V${y+h-r} a${r},${r} 0 0 1 ${-r},${r} H${x0} Z`
    : `M${x0},${y} H${x1+r} a${r},${r} 0 0 0 ${-r},${r} V${y+h-r} a${r},${r} 0 0 0 ${r},${r} H${x0} Z`;
}

function divergingBars(mount, items, opts) {
  const host = document.getElementById(mount);
  host.textContent = "";
  if (!items.length) {
    const p = document.createElement("p");
    p.className = "empty"; p.textContent = "Sin datos para este filtro.";
    host.appendChild(p); return;
  }

  const BAND = 30, BAR = 18, GAP_R = 4;
  // Canaleta reservada a cada lado del plot para el valor directo, para que
  // nunca invada la columna de etiquetas de categoria.
  const LABEL_W = opts.labelW, GUTTER = 76;
  const padL = LABEL_W + GUTTER, padR = GUTTER, padT = 8, padB = 26;
  // 1 unidad de SVG = 1 px: el viewBox sigue al ancho real del contenedor,
  // asi el texto no se escala segun el tamano del card.
  const W = Math.max(host.clientWidth || 620, LABEL_W + 2 * GUTTER + 120);
  const H = padT + items.length * BAND + padB;
  const plotW = W - padL - padR;

  const u = opts.u;
  const max = Math.max(...items.map(d => Math.abs(d.value)), u.money ? 0.01 : 1);
  const nice = niceMax(max);
  const zero = padL + plotW / 2;
  const x = v => zero + (v / nice) * (plotW / 2);

  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("viewBox", `0 0 ${W} ${H}`);
  svg.setAttribute("width", W); svg.setAttribute("height", H);
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", opts.aria);

  // grilla: hairline solida, recesiva
  for (const t of [-nice, -nice/2, 0, nice/2, nice]) {
    const gx = x(t);
    const ln = document.createElementNS(NS, "line");
    ln.setAttribute("x1", gx); ln.setAttribute("x2", gx);
    ln.setAttribute("y1", padT); ln.setAttribute("y2", H - padB);
    ln.setAttribute("stroke", t === 0 ? css("--axis") : css("--grid"));
    ln.setAttribute("stroke-width", "1");
    svg.appendChild(ln);
    const tx = document.createElementNS(NS, "text");
    tx.setAttribute("x", gx); tx.setAttribute("y", H - padB + 16);
    tx.setAttribute("text-anchor", "middle");
    tx.setAttribute("font-size", "10.5"); tx.setAttribute("fill", css("--text-muted"));
    tx.style.fontVariantNumeric = "tabular-nums";
    tx.textContent = tickFmt(t, u);
    svg.appendChild(tx);
  }

  items.forEach((d, i) => {
    const y = padT + i * BAND + (BAND - BAR) / 2;
    const pos = d.value >= 0;
    const color = pos ? css("--pos") : css("--neg");

    const p = document.createElementNS(NS, "path");
    p.setAttribute("d", barPath(zero, x(d.value), y, BAR, GAP_R));
    p.setAttribute("fill", color);
    svg.appendChild(p);

    // etiqueta de categoria (ink, nunca el color de la serie)
    const lb = document.createElementNS(NS, "text");
    lb.setAttribute("x", LABEL_W); lb.setAttribute("y", y + BAR / 2 + 4);
    lb.setAttribute("text-anchor", "end"); lb.setAttribute("font-size", "12");
    lb.setAttribute("fill", css("--text-primary"));
    lb.textContent = d.label;
    svg.appendChild(lb);

    // valor directo, siempre fuera de la punta y dentro de la canaleta
    const tipX = x(d.value);
    const vl = document.createElementNS(NS, "text");
    vl.setAttribute("y", y + BAR / 2 + 4);
    vl.setAttribute("text-anchor", pos ? "start" : "end");
    vl.setAttribute("font-size", "11.5");
    vl.setAttribute("fill", css("--text-secondary"));
    vl.style.fontVariantNumeric = "tabular-nums";
    vl.textContent = samt(d.value, u);
    vl.setAttribute("x", tipX + (pos ? 8 : -8));
    svg.appendChild(vl);

    // hit target: toda la banda, mas grande que la marca
    const hit = document.createElementNS(NS, "rect");
    hit.setAttribute("x", 0);
    hit.setAttribute("y", padT + i * BAND);
    hit.setAttribute("width", W); hit.setAttribute("height", BAND);
    hit.setAttribute("fill", "transparent");
    hit.setAttribute("tabindex", "0");
    hit.setAttribute("role", "listitem");
    hit.setAttribute("aria-label", `${d.label}: ${samt(d.value, u)} ${u.name}`);
    const rows = [
      [`cargado en sistema (${u.suf})`, amt(d.q, u), null],
      [`hoja de trabajo (${u.suf})`, amt(d.w, u), null],
      ["desvío", samt(d.value, u, true), color],
    ];
    if (d.extra) rows.push([d.extra, "", null]);
    const show = e => { p.setAttribute("opacity", ".72"); showTip(e, d.tip || d.label, rows); };
    const hide = () => { p.setAttribute("opacity", "1"); hideTip(); };
    hit.addEventListener("pointerenter", show);
    hit.addEventListener("pointermove", moveTip);
    hit.addEventListener("pointerleave", hide);
    hit.addEventListener("focus", e => {
      const r = hit.getBoundingClientRect();
      show({clientX: r.left + r.width / 2, clientY: r.top + r.height});
    });
    hit.addEventListener("blur", hide);
    svg.appendChild(hit);
  });

  host.appendChild(svg);
}

function niceMax(m) {
  const p = Math.pow(10, Math.floor(Math.log10(m)));
  for (const s of [1, 2, 2.5, 5, 10]) if (m <= s * p) return s * p;
  return 10 * p;
}

function legend(id) {
  const el = document.getElementById(id);
  el.textContent = "";
  const items = [["De más en sistema", "--pos"],
                 ["Falta en sistema", "--neg"]];
  for (const [txt, v] of items) {
    const d = document.createElement("span"); d.className = "lg";
    const i = document.createElement("i"); i.style.background = css(v);
    const s = document.createElement("span"); s.textContent = txt;
    d.appendChild(i); d.appendChild(s); el.appendChild(d);
  }
}

/* ---------- agregacion ---------- */
function groupBy(rows, keyFn) {
  const m = new Map();
  for (const r of rows) {
    for (const k of [].concat(keyFn(r))) {
      if (!k) continue;
      const a = m.get(k) || {q: 0, w: 0, d: 0, n: 0};
      a.q += r.q; a.w += r.w; a.d += r.d; a.n += r.s !== "OK" ? 1 : 0;
      m.set(k, a);
    }
  }
  return m;
}

/* ---------- estados ---------- */
const CHIP = {
  "OK":             ["--good",     "✓", "OK"],
  "DESVIO":         ["--serious",  "▲", "Desvío"],
  "FALTA EN QUERY": ["--critical", "✕", "Falta en sistema"],
  "SOLO EN QUERY":  ["--warning",  "!",      "Solo en sistema"],
};

/* ---------- render ---------- */
function render() {
  const rows = filtered();
  const u = unit();
  const sc = v => v * u.k;   // fuente (miles de USD) -> unidad del tablero

  const q = sc(rows.reduce((s, r) => s + r.q, 0));
  const w = sc(rows.reduce((s, r) => s + r.w, 0));
  const dev = rows.filter(r => r.s !== "OK");
  const absImpact = sc(dev.reduce((s, r) => s + Math.abs(r.d), 0));

  setVal("k-hero", amt(absImpact, u), u);
  document.getElementById("k-hero-note").textContent =
    `Suma de desvíos en valor absoluto, en ${u.name} · ` +
    `${dev.length} de ${rows.length} líneas de control`;
  setVal("k-q", amt(q, u), u);
  setVal("k-w", amt(w, u), u);
  const d = q - w;
  setVal("k-d", samt(d, u), u);
  document.getElementById("k-d-note").textContent =
    Math.abs(d) < 1e-9 ? "Sistema y hoja coinciden"
      : d > 0 ? "Sistema por encima de la hoja" : "Sistema por debajo de la hoja";
  document.getElementById("k-n").textContent = dev.length;
  const paises = new Set(dev.map(r => r.country));
  document.getElementById("k-n-note").textContent =
    paises.size ? `${paises.size} país${paises.size > 1 ? "es" : ""} afectado${paises.size > 1 ? "s" : ""}`
                : "Todo conciliado";

  // por pais
  const byC = [...groupBy(rows, r => r.country)]
    .map(([k, v]) => ({label: k, tip: k, value: sc(v.d), q: sc(v.q), w: sc(v.w),
                       extra: v.n ? `${v.n} línea(s) con desvío` : null}))
    .filter(d => Math.abs(d.value) > 1e-9)
    .sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
  divergingBars("c-pais", byC, {labelW: 60, u, aria: "Desvío por país"});
  document.getElementById("c-pais-sub").textContent =
    `Sistema menos hoja de trabajo, en ${u.name}`;
  legend("lg-pais");

  // por responsable
  const byR = [...groupBy(rows.filter(r => r.s !== "OK"), r => (r.r || "").split(", "))]
    .map(([k, v]) => ({label: k, tip: `Responsable ${k}`,
                       value: sc(v.d), q: sc(v.q), w: sc(v.w),
                       extra: `${v.n} línea(s) con desvío`}))
    .filter(d => Math.abs(d.value) > 1e-9)
    .sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
  divergingBars("c-resp", byR, {labelW: 60, u, aria: "Desvío por responsable"});
  document.getElementById("c-resp-sub").textContent =
    `A quién hay que avisar, ordenado por impacto, en ${u.name}`;
  legend("lg-resp");

  // top lineas
  const top = dev.slice().sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 12)
    .map(r => ({label: `${r.country} · ${r.sbe3 || r.sbe1}`,
                tip: `${r.country} · ${r.sbe1} · ${r.sbe3} · ${r.pl}`,
                value: sc(r.d), q: sc(r.q), w: sc(r.w),
                extra: r.r ? `Responsable: ${r.r}` : null}));
  divergingBars("c-linea", top, {labelW: 190, u, aria: "Mayores desvíos por línea"});
  document.getElementById("c-linea-sub").textContent =
    `Las ${top.length} combinaciones de mayor impacto` +
    (state.pl === "Todas" ? " (todas las líneas de importe)" : ` · ${state.pl}`) +
    `, en ${u.name}`;
  legend("lg-linea");

  // tabla
  const tb = document.getElementById("tbody");
  tb.textContent = "";
  const sorted = rows.slice().sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
  for (const r of sorted) {
    const tr = document.createElement("tr");
    for (const [v, cls] of [[r.country, "k"], [r.sbe1, ""], [r.sbe2, ""],
                            [r.sbe3, ""], [r.pl, ""]]) {
      const td = document.createElement("td");
      td.className = cls; td.textContent = v; tr.appendChild(td);
    }
    for (const v of [r.q, r.w]) {
      const td = document.createElement("td");
      td.className = "num"; td.textContent = amt(sc(v), u, true); tr.appendChild(td);
    }
    const tdd = document.createElement("td");
    tdd.className = "num k"; tdd.textContent = samt(sc(r.d), u, true);
    tr.appendChild(tdd);

    const tds = document.createElement("td");
    const [tok, icon, label] = CHIP[r.s] || CHIP["DESVIO"];
    const chip = document.createElement("span");
    chip.className = "chip";
    const b = document.createElement("b");
    b.textContent = icon; b.style.color = css(tok);
    chip.appendChild(b);
    chip.appendChild(document.createTextNode(label));
    tds.appendChild(chip); tr.appendChild(tds);

    const tdr = document.createElement("td");
    tdr.textContent = r.r; tr.appendChild(tdr);
    tb.appendChild(tr);
  }
  document.getElementById("t-sub").textContent =
    `${sorted.length} línea(s), ordenadas por impacto, en ${u.name}. ` +
    `Es la vista de tabla del tablero: todo valor de los gráficos está acá.`;
}

/* ---------- init ---------- */
function init() {
  document.getElementById("subtitle").textContent =
    `Cruce de la query de SAP BW contra PreFC Y1 · granularidad ${META.grain.toUpperCase()} · ` +
    `tolerancia ${META.tol} · generado el ${META.generado}`;

  const notes = [];
  if (META.glyVol) notes.push(
    "El volumen de GLY está excluido del cruce: la query lo trae en otra unidad de medida " +
    "y no se encontró la tabla FactoresGly.xlsx para convertirlo a Regs.");
  if (META.glySinFactor && META.glySinFactor.length) notes.push(
    `Volumen de GLY: ${META.glySinFactor.length} SKU sin factor en FactoresGly.xlsx ` +
    `(${META.glySinFactor.map(s => `${s.sku} ${s.country} ${s.sbe3}`).join("; ")}). ` +
    "Esas líneas quedan fuera del cruce hasta que se agreguen a la tabla.");
  if (META.glyColapsado && META.glyColapsado.length) notes.push(
    `Glifosato genérico: la query lo carga contra un material sin marca, así que ` +
    `${META.glyColapsado.map(k => `${k[0]} · ${k[2]}`).join("; ")} se compara a nivel GLY ` +
    "sin abrir Roundup vs La Tijereta. El total se controla igual; el reparto por marca " +
    "que hace la hoja no se puede verificar contra el sistema.");
  if (META.plLevel === "group") notes.push(
    "Las líneas de deducción se comparan agrupadas como Sales Adj: el sistema las abre " +
    "en tres y la hoja carga el total en una.");
  if (META.entidad.length) notes.push(
    `${META.entidad.length} combinación(es) cargadas en una entidad legal distinta de la esperada.`);
  if (META.unmapped.length) notes.push(
    `Company codes sin mapear a país: ${META.unmapped.join(", ")}.`);
  if (notes.length) {
    const box = document.createElement("div");
    box.className = "warnbox";
    notes.forEach((n, i) => {
      if (i) box.appendChild(document.createElement("br"));
      box.appendChild(document.createTextNode("⚠  " + n));
    });
    document.getElementById("warn").appendChild(box);
  }

  fillSelect("f-pl", [["Todas", "Todos los importes"], ...uniq("pl")], state.pl);
  fillSelect("f-country", ["Todos", ...uniq("country")], state.country);
  fillSelect("f-sbe1", ["Todos", ...uniq("sbe1")], state.sbe1);
  const resps = [...new Set(DATA.flatMap(r => (r.r || "").split(", ")).filter(Boolean))].sort();
  fillSelect("f-resp", ["Todos", ...resps], state.resp);

  const bind = (id, key) => document.getElementById(id)
    .addEventListener("change", e => { state[key] = e.target.value; render(); });
  bind("f-pl", "pl"); bind("f-country", "country");
  bind("f-sbe1", "sbe1"); bind("f-resp", "resp");
  document.getElementById("f-dev").addEventListener("change", e => {
    state.dev = e.target.checked; render();
  });

  const btn = document.getElementById("theme");
  const apply = (t) => {
    document.documentElement.setAttribute("data-theme", t);
    btn.textContent = t === "dark" ? "Modo claro" : "Modo oscuro";
    render();
  };
  btn.addEventListener("click", () => apply(
    document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark"));
  apply(matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");

  // los SVG se dibujan a 1 unidad = 1 px, asi que hay que redibujar al resize
  let rt; addEventListener("resize", () => {
    clearTimeout(rt); rt = setTimeout(render, 120);
  });

  document.getElementById("foot").textContent =
    "Importes en millones de USD: la query y la hoja de trabajo los traen divididos " +
    "por 1000 (miles de USD) para armar los reportes, acá se muestran en la magnitud " +
    "real del negocio. El volumen queda en unidades y se mira por separado. " +
    "Desvío = sistema − hoja de trabajo. Fuente: Checksibf.xlsx, hojas query y " +
    "PreFC Y1. Generado por check_sbp.py / dashboard.py.";
}
init();
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description="Tablero SBP 2027")
    ap.add_argument("--book", default=chk.BOOK)
    ap.add_argument("--out", default=None)
    ap.add_argument("--grain", default="sbe3", choices=sorted(chk.GRAINS))
    ap.add_argument("--tol", type=float, default=chk.DEFAULT_TOL)
    ap.add_argument("--pl-level", default="group", choices=("group", "sub"))
    args = ap.parse_args()

    base = os.path.dirname(os.path.abspath(args.book))
    out = args.out or os.path.join(base, "Tablero_SBP2027.html")

    rows, meta = build_rows(args.book, args.grain, args.tol, args.pl_level)
    doc = (TEMPLATE
           .replace("__DATA__", json.dumps(rows, ensure_ascii=False))
           .replace("__META__", json.dumps(meta, ensure_ascii=False)))

    with open(out, "w", encoding="utf-8") as fh:
        fh.write(doc)

    dev = [r for r in rows if r["s"] != "OK"]
    print(f"Lineas de control : {len(rows)}")
    print(f"Con desvio        : {len(dev)}")
    print(f"Tablero           : {out}")


if __name__ == "__main__":
    main()
