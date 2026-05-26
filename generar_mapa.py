# -*- coding: utf-8 -*-
"""
generar_mapa.py
Genera mapa_dgt.html: mapa interactivo municipal DGT 2020-2025 con perspectiva de género.
Uso: python generar_mapa.py
"""

import json
import os
import sys
from pathlib import Path

import geopandas as gpd
import pandas as pd

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
SHP_PATH = (
    DATA_DIR
    / "recintos_municipales_inspire_peninbal_etrs89"
    / "recintos_municipales_inspire_peninbal_etrs89.shp"
)
EXCEL_2025 = DATA_DIR / "DatosMunicipalesGeneral_2025(1).xlsx"
EXCEL_2020 = DATA_DIR / "DatosMunicipalesGeneral_2020.xlsx"
OUTPUT_HTML = BASE_DIR / "mapa_dgt.html"

SIMPLIFY_TOLERANCE = 0.005
COORD_PRECISION = 4

# ---------------------------------------------------------------------------
# Column names
# ---------------------------------------------------------------------------
C_INE        = "Código INE"
C_MUN        = "Municipio"
C_PROV       = "Provincia"
C_CA         = "Comunidad Autónoma"
C_POB        = "Población Total"
C_COND_H     = "Conductores Hombres"
C_COND_M     = "Conductoras Mujeres"
C_COND_TOT   = "Censo Conductores"
C_PARQ_CIC   = "Parque Ciclomotores"
C_PARQ_MOTO  = "Parque Motocicletas"
C_PARQ_TUR   = "Parque Turismos"
C_PARQ_FUR   = "Parque Furgonetas"
C_PARQ_CAM   = "Parque Camiones"
C_PARQ_TOT   = "Parque Total"
C_MOTO_ITV   = "Motocicletas sin ITV (<25 años)"
C_TUR_ITV    = "Turismos sin ITV (<25 años)"
C_RESTO_ITV  = "Resto de Vehículos sin ITV (<25 años)"
C_MOTO_25    = "Parque Motocicletas (<25 años)"
C_MOTO_4     = "Parque Motocicletas (<4 años)"
C_TUR_25     = "Parque Turismos (<25 años)"
C_TUR_4      = "Parque Turismos (<4 años)"
C_CAM_25     = "Parque Camiones (<25 años)"
C_CAM_4      = "Parque Camiones (<4 años)"
C_ANTIG_MOTO = "Antigüedad Media de Motocicletas"
C_ANTIG_TUR  = "Antigüedad Media de Turismos"
C_ANTIG_CAM  = "Antigüedad Media de Camiones"
C_GASOLINA   = "Gasolina"
C_DIESEL     = "Diesel"
C_ELEC       = "Electrificado"
C_GLP        = "GLP (Gas Licuado del Petroleo)"
C_GNL        = "GNL (Gas Natural Licuado)"
C_OTROS_PROP = "Otros Tipos de Propulsión"
C_DIST_B     = "Distintivo B"
C_DIST_C     = "Distintivo C"
C_DIST_ECO   = "Distintivo ECO"
C_DIST_0     = "Distintivo 0"
C_SIN_DIST   = "Sin Distintivo"

FUEL_COLS = [C_GASOLINA, C_DIESEL, C_ELEC, C_GLP, C_GNL, C_OTROS_PROP]
DIST_COLS = [C_DIST_B, C_DIST_C, C_DIST_ECO, C_DIST_0, C_SIN_DIST]


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_shapefile():
    print("  Cargando shapefile...")
    gdf = gpd.read_file(SHP_PATH)
    gdf = gdf.to_crs(epsg=4326)
    print(f"  Simplificando geometrías (tolerancia={SIMPLIFY_TOLERANCE})...")
    gdf["geometry"] = gdf.geometry.simplify(SIMPLIFY_TOLERANCE, preserve_topology=True)
    gdf["ine_code"] = gdf["NATCODE"].str[-5:]
    return gdf[["ine_code", "NAMEUNIT", "geometry"]].copy()


def load_excel(path):
    print(f"  Cargando {path.name}...")
    df = pd.read_excel(path, dtype={C_INE: str})
    mask = df[C_MUN].str.contains("sin especificar", case=False, na=False)
    df = df[~mask].copy()
    df["ine"] = df[C_INE].str.strip().str.zfill(5)
    return df


# ---------------------------------------------------------------------------
# Indicator computation
# ---------------------------------------------------------------------------

def safe_div(num, den):
    return num / den.replace(0, float("nan"))


def compute_indicators(df, s):
    """Compute all indicators. s = suffix string e.g. '_25' or '_20'."""
    d = df.copy()

    cond_tot  = d[C_COND_TOT].replace(0, float("nan"))
    pop_tot   = d[C_POB].replace(0, float("nan"))
    moto_25   = d[C_MOTO_25].replace(0, float("nan"))
    tur_25    = d[C_TUR_25].replace(0, float("nan"))
    cam_25    = d[C_CAM_25].replace(0, float("nan"))
    parq_tot  = d[C_PARQ_TOT].replace(0, float("nan"))
    fuel_sum  = d[FUEL_COLS].sum(axis=1).replace(0, float("nan"))
    dist_sum  = d[DIST_COLS].sum(axis=1).replace(0, float("nan"))

    # Tab 1 - Gender
    d[f"pct_conductoras{s}"]   = (safe_div(d[C_COND_M], cond_tot) * 100).round(2)
    d[f"ratio_hm{s}"]          = safe_div(d[C_COND_H], d[C_COND_M].replace(0, float("nan"))).round(2)
    d[f"n_conductoras{s}"]     = d[C_COND_M].round(0)

    # Tab 2 - Vehicle fleet
    d[f"veh_por_1000{s}"]      = (safe_div(d[C_PARQ_TOT], pop_tot) * 1000).round(1)
    d[f"pct_turismos{s}"]      = (safe_div(d[C_PARQ_TUR], parq_tot) * 100).round(1)
    d[f"parque_total{s}"]      = d[C_PARQ_TOT].round(0)

    # Tab 3 - Fleet age
    d[f"antig_turismos{s}"]    = d[C_ANTIG_TUR].round(2)
    d[f"pct_tur_nuevos{s}"]    = (safe_div(d[C_TUR_4], tur_25) * 100).round(1)

    # Tab 4 - Electrification & fuel
    d[f"pct_electrificado{s}"] = (safe_div(d[C_ELEC], fuel_sum) * 100).round(3)
    d[f"pct_diesel{s}"]        = (safe_div(d[C_DIESEL], fuel_sum) * 100).round(1)
    d[f"pct_gasolina{s}"]      = (safe_div(d[C_GASOLINA], fuel_sum) * 100).round(1)

    # Tab 5 - Environmental badges
    d[f"pct_sin_distintivo{s}"] = (safe_div(d[C_SIN_DIST], dist_sum) * 100).round(1)
    d[f"pct_eco_0{s}"]          = (safe_div(d[C_DIST_ECO] + d[C_DIST_0], dist_sum) * 100).round(1)

    # Tab 6 - ITV
    d[f"pct_moto_sin_itv{s}"]  = (safe_div(d[C_MOTO_ITV], moto_25) * 100).round(1)

    # Tab 7 - Motocicletas
    d[f"pct_motos{s}"]         = (safe_div(d[C_PARQ_MOTO], parq_tot) * 100).round(1)
    d[f"motos_por_1000{s}"]    = (safe_div(d[C_PARQ_MOTO], pop_tot) * 1000).round(1)
    d[f"antig_motos{s}"]       = d[C_ANTIG_MOTO].round(2)
    d[f"pct_motos_nuevas{s}"]  = (safe_div(d[C_MOTO_4], moto_25) * 100).round(1)

    # Tab 8 - Camiones
    d[f"pct_camiones{s}"]      = (safe_div(d[C_PARQ_CAM], parq_tot) * 100).round(1)
    d[f"cam_por_1000{s}"]      = (safe_div(d[C_PARQ_CAM], pop_tot) * 1000).round(1)
    d[f"antig_camiones{s}"]    = d[C_ANTIG_CAM].round(2)
    d[f"pct_cam_nuevos{s}"]    = (safe_div(d[C_CAM_4], cam_25) * 100).round(1)
    d[f"pct_resto_sin_itv{s}"] = (safe_div(d[C_RESTO_ITV], cam_25) * 100).round(1)

    return d


def compute_evolution(df25, df20):
    cols_20 = [
        "ine",
        "pct_conductoras_20",
        "veh_por_1000_20",
        "antig_turismos_20",
        "pct_electrificado_20",
        "pct_sin_distintivo_20",
        "antig_motos_20",
        "pct_motos_20",
        "antig_camiones_20",
        "pct_camiones_20",
    ]
    merged = df25.merge(df20[cols_20], on="ine", how="left")
    merged["delta_conductoras"]    = (merged["pct_conductoras_25"]    - merged["pct_conductoras_20"]).round(2)
    merged["delta_veh_por_1000"]   = (merged["veh_por_1000_25"]       - merged["veh_por_1000_20"]).round(1)
    merged["delta_antig_turismos"] = (merged["antig_turismos_25"]     - merged["antig_turismos_20"]).round(2)
    merged["delta_electrificado"]  = (merged["pct_electrificado_25"]  - merged["pct_electrificado_20"]).round(3)
    merged["delta_sin_distintivo"] = (merged["pct_sin_distintivo_25"] - merged["pct_sin_distintivo_20"]).round(1)
    merged["delta_antig_motos"]    = (merged["antig_motos_25"]        - merged["antig_motos_20"]).round(2)
    merged["delta_pct_motos"]      = (merged["pct_motos_25"]          - merged["pct_motos_20"]).round(1)
    merged["delta_antig_camiones"] = (merged["antig_camiones_25"]     - merged["antig_camiones_20"]).round(2)
    merged["delta_pct_camiones"]   = (merged["pct_camiones_25"]       - merged["pct_camiones_20"]).round(1)
    # Ratio de crecimiento eléctrico: NaN cuando base es 0
    merged["ratio_crecimiento_elec"] = safe_div(
        merged["pct_electrificado_25"],
        merged["pct_electrificado_20"].replace(0, float("nan"))
    ).round(2)
    return merged


# ---------------------------------------------------------------------------
# GeoJSON assembly
# ---------------------------------------------------------------------------

INDICATOR_COLS = [
    "ine", C_MUN, C_PROV, C_CA,
    # Tab 1
    "pct_conductoras_25", "ratio_hm_25", "n_conductoras_25", "delta_conductoras",
    "pct_conductoras_20",
    # Tab 2
    "veh_por_1000_25", "pct_turismos_25", "parque_total_25", "delta_veh_por_1000",
    # Tab 3
    "antig_turismos_25", "pct_tur_nuevos_25", "delta_antig_turismos",
    # Tab 4
    "pct_electrificado_25", "pct_electrificado_20", "delta_electrificado",
    "ratio_crecimiento_elec", "pct_diesel_25", "pct_gasolina_25",
    # Tab 5
    "pct_sin_distintivo_25", "pct_eco_0_25", "delta_sin_distintivo",
    # Tab 6
    "pct_moto_sin_itv_25",
    # Tab 7 - Motos
    "pct_motos_25", "motos_por_1000_25", "antig_motos_25", "pct_motos_nuevas_25",
    "delta_antig_motos", "delta_pct_motos",
    # Tab 8 - Camiones
    "pct_camiones_25", "cam_por_1000_25", "antig_camiones_25", "pct_cam_nuevos_25",
    "pct_resto_sin_itv_25", "delta_antig_camiones", "delta_pct_camiones",
]


def round_coords(coords):
    if not coords:
        return coords
    if isinstance(coords[0], (int, float)):
        return [round(c, COORD_PRECISION) for c in coords]
    return [round_coords(ring) for ring in coords]


def build_geojson(gdf_shp, df_indicators):
    print("  Uniendo geometrías con indicadores...")
    cols_needed = [c for c in INDICATOR_COLS if c in df_indicators.columns]
    final = gdf_shp.merge(df_indicators[cols_needed], left_on="ine_code", right_on="ine", how="left")

    keep = ["NAMEUNIT", "ine_code"] + [c for c in cols_needed if c != "ine"]
    keep = [c for c in keep if c in final.columns]

    print("  Exportando GeoJSON...")
    geojson_dict = json.loads(final[keep + ["geometry"]].to_json(drop_id=True, show_bbox=False))

    print("  Redondeando coordenadas...")
    for feature in geojson_dict["features"]:
        geom = feature.get("geometry")
        if geom and geom.get("coordinates"):
            geom["coordinates"] = round_coords(geom["coordinates"])

    return json.dumps(geojson_dict, separators=(",", ":"), ensure_ascii=False)


# ---------------------------------------------------------------------------
# Color scale metadata for JS
# ---------------------------------------------------------------------------

SCALE_SPECS = [
    # red2grn = alto es bueno (verde) | grn2red = alto es malo (rojo) | div = delta (rojo=peor, verde=mejor)
    ("pct_conductoras_25",    "red2grn"),  # más conductoras = más igualdad
    ("ratio_hm_25",           "grn2red"),  # más desigualdad H/M = peor
    ("delta_conductoras",     "div"),      # positivo = más mujeres = mejor
    ("pct_conductoras_20",    "red2grn"),
    ("veh_por_1000_25",       "grn2red"),  # más motorización = peor para sostenibilidad
    ("pct_turismos_25",       "grn2red"),  # más turismos = más vehículos privados = peor
    ("delta_veh_por_1000",    "div_inv"),  # positivo = más vehículos = peor
    ("antig_turismos_25",     "grn2red"),  # más antigüedad = más contaminante
    ("pct_tur_nuevos_25",     "red2grn"),  # más nuevos = más sostenible
    ("delta_antig_turismos",  "div_inv"),  # positivo = flota más vieja = peor
    ("pct_electrificado_25",  "red2grn"),  # más eléctrico = mejor
    ("pct_electrificado_20",  "red2grn"),
    ("delta_electrificado",   "div"),      # positivo = más eléctrico = mejor
    ("ratio_crecimiento_elec","red2grn"),  # mayor ratio = mejor
    ("pct_diesel_25",         "grn2red"),  # más diésel = más contaminante
    ("pct_sin_distintivo_25", "grn2red"),  # más sin etiqueta = más contaminante
    ("pct_eco_0_25",          "red2grn"),  # más ECO/0 = mejor
    ("delta_sin_distintivo",  "div_inv"),  # positivo = más sin etiqueta = peor
    ("pct_moto_sin_itv_25",   "grn2red"),  # más sin ITV = más riesgo
    ("pct_motos_25",          "grn2red"),  # más motos = más exposición a riesgo
    ("motos_por_1000_25",     "grn2red"),
    ("antig_motos_25",        "grn2red"),  # más antigüedad = peor
    ("pct_motos_nuevas_25",   "red2grn"),  # más nuevas = mejor
    ("delta_antig_motos",     "div_inv"),  # positivo = flota más vieja = peor
    ("delta_pct_motos",       "div_inv"),  # positivo = más motos = peor
    ("pct_camiones_25",       "grn2red"),  # más camiones = más exposición a riesgo
    ("cam_por_1000_25",       "grn2red"),
    ("antig_camiones_25",     "grn2red"),  # más antigüedad = peor
    ("pct_cam_nuevos_25",     "red2grn"),  # más nuevos = mejor
    ("pct_resto_sin_itv_25",  "grn2red"),  # más sin ITV = más riesgo
    ("delta_antig_camiones",  "div_inv"),  # positivo = flota más vieja = peor
    ("delta_pct_camiones",    "div_inv"),  # positivo = más camiones = peor
]


def compute_scales(df):
    scales = {}
    for field, ctype in SCALE_SPECS:
        if field not in df.columns:
            continue
        series = pd.to_numeric(df[field], errors="coerce").dropna()
        if series.empty:
            continue
        nonzero = series[series > 0]
        vmin = round(float(nonzero.min()), 4) if not nonzero.empty else 0
        vmax = round(float(series.quantile(0.99)), 4)
        scales[field] = {"type": ctype, "vmin": vmin, "vmax": vmax}
    return scales


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Visualizador DGT Municipal 2020–2025</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#1a1a2e;color:#eee;display:flex;flex-direction:column;height:100vh;overflow:hidden}
#header{padding:8px 16px;background:#16213e;display:flex;align-items:center;gap:16px;flex-shrink:0;border-bottom:1px solid #0f3460}
#header h1{font-size:1.05em;color:#e94560;font-weight:600;white-space:nowrap}
#header .note{color:#666;font-size:0.75em;margin-left:auto;white-space:nowrap}
#tabs{display:flex;background:#0f3460;overflow-x:auto;flex-shrink:0;scrollbar-width:thin}
#tabs::-webkit-scrollbar{height:4px}
#tabs::-webkit-scrollbar-thumb{background:#e94560;border-radius:2px}
.tab-btn{padding:9px 16px;cursor:pointer;white-space:nowrap;border:none;background:none;color:#aaa;font-size:0.82em;border-bottom:2px solid transparent;transition:all .15s}
.tab-btn.active{color:#e94560;border-bottom-color:#e94560;background:#16213e}
.tab-btn:hover{background:#16213e;color:#ddd}
#layer-bar{display:flex;padding:6px 12px;background:#16213e;gap:8px;flex-wrap:wrap;align-items:center;flex-shrink:0;border-bottom:1px solid #0f3460;min-height:40px}
.lbtn{padding:5px 12px;border-radius:16px;border:1px solid #333;cursor:pointer;background:#1a1a2e;color:#aaa;font-size:0.78em;transition:all .15s;white-space:nowrap}
.lbtn.active{background:#e94560;color:#fff;border-color:#e94560}
.lbtn:hover{border-color:#e94560;color:#fff}
#map{flex:1;min-height:0}
#legend{position:absolute;bottom:28px;right:10px;z-index:1000;background:rgba(22,33,62,0.93);padding:10px 12px;border-radius:6px;min-width:150px;border:1px solid #0f3460;pointer-events:none}
#legend h4{font-size:0.75em;margin-bottom:6px;color:#e94560;font-weight:600}
#legend-bar{height:12px;border-radius:3px;margin:4px 0}
.leg-lbls{display:flex;justify-content:space-between;font-size:0.7em;color:#888}
#tooltip{position:fixed;z-index:2000;background:rgba(22,33,62,0.97);padding:10px 14px;border-radius:6px;pointer-events:none;border-left:3px solid #e94560;max-width:290px;font-size:0.8em;display:none;box-shadow:0 4px 20px rgba(0,0,0,.5)}
#tooltip h3{color:#e94560;margin-bottom:5px;font-size:1em;font-weight:600}
#tooltip .prov{color:#666;font-size:0.85em;margin-bottom:6px}
#tooltip table{border-collapse:collapse;width:100%}
#tooltip td{padding:2px 0}
#tooltip td:first-child{color:#888;padding-right:10px;white-space:nowrap}
#tooltip td:last-child{color:#eee;font-weight:500;text-align:right}
#info-note{position:absolute;bottom:8px;left:10px;z-index:1000;color:#444;font-size:0.68em}
.leaflet-popup-content-wrapper{background:rgba(22,33,62,0.97);color:#eee;border:1px solid #0f3460;border-radius:6px}
.leaflet-popup-tip{background:rgba(22,33,62,0.97)}
.leaflet-popup-content{font-size:0.82em;margin:10px 14px}
.popup-title{color:#e94560;font-weight:600;font-size:1.05em;margin-bottom:3px}
.popup-sub{color:#666;font-size:0.85em;margin-bottom:8px}
.popup-table{border-collapse:collapse;width:100%}
.popup-table td{padding:3px 0}
.popup-table td:first-child{color:#888;padding-right:12px}
.popup-table td:last-child{color:#eee;font-weight:500;text-align:right}
.popup-table tr.highlight td{color:#e9c46a}
</style>
</head>
<body>
<div id="header">
  <h1>&#x1F6E3;&#xFE0F; Visualizador DGT — Datos Municipales 2020 / 2025</h1>
  <span class="note">Fuente: DGT &nbsp;|&nbsp; Generado: __FECHA__ &nbsp;|&nbsp; Pen&iacute;nsula + Baleares (Canarias no disponible)</span>
</div>
<div id="tabs"></div>
<div id="layer-bar"></div>
<div id="map"></div>
<div id="legend"><h4 id="leg-title">Leyenda</h4><div id="legend-bar"></div><div class="leg-lbls"><span id="leg-min"></span><span id="leg-max"></span></div><div id="legend-zero" style="margin-top:7px;font-size:0.72em;color:#aaa;align-items:center;gap:5px"><span style="display:inline-block;width:12px;height:12px;background:#000;border-radius:2px;flex-shrink:0"></span>0% sin vehículos eléctricos</div></div>
<div id="tooltip"></div>
<div id="info-note">Sin datos: gris &nbsp;|&nbsp; Hover: detalle &nbsp;|&nbsp; Click: ficha completa</div>

<script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
// ── EMBEDDED DATA ──────────────────────────────────────────────────────────
const GEOJSON = __GEOJSON__;
const SCALES  = __SCALES__;

// ── TAB / LAYER CONFIGURATION ──────────────────────────────────────────────
const TABS = [
  { name: "1. Conductores y Género", layers: [
    { key: "pct_conductoras_25",  label: "% Conductoras 2025",          unit: "%",    fmt:1 },
    { key: "pct_conductoras_20",  label: "% Conductoras 2020",          unit: "%",    fmt:1 },
    { key: "ratio_hm_25",         label: "Ratio conductores H/M 2025",  unit: "x",    fmt:2 },
    { key: "delta_conductoras",   label: "Evolución % conductoras (pp)", unit: "pp", fmt:2, isDelta:true },
  ]},
  { name: "2. Parque de Vehículos", layers: [
    { key: "veh_por_1000_25",     label: "Vehículos por 1.000 hab. 2025", unit: "",  fmt:1 },
    { key: "pct_turismos_25",     label: "% Turismos sobre parque 2025",  unit: "%",  fmt:1 },
    { key: "delta_veh_por_1000",  label: "Evolución vehículos/1.000 hab.", unit: "", fmt:1, isDelta:true },
  ]},
  { name: "3. Antigüedad del Parque", layers: [
    { key: "antig_turismos_25",    label: "Antigüedad media turismos 2025", unit: "años", fmt:1 },
    { key: "pct_tur_nuevos_25",    label: "% Turismos <4 años (2025)",       unit: "%",    fmt:1 },
    { key: "delta_antig_turismos", label: "Evolución antigüedad turismos",    unit: "años", fmt:2, isDelta:true },
  ]},
  { name: "4. Electrificación y Combustible", layers: [
    { key: "pct_electrificado_25",   label: "% Electrificados 2025",           unit: "%",  fmt:3 },
    { key: "pct_electrificado_20",   label: "% Electrificados 2020",           unit: "%",  fmt:3 },
    { key: "delta_electrificado",    label: "Evolución % electrificados (pp)",  unit: "pp", fmt:3, isDelta:true },
    { key: "ratio_crecimiento_elec", label: "Ratio crecimiento eléctrico (x veces)", unit: "x", fmt:2 },
    { key: "pct_diesel_25",          label: "% Diésel 2025",                    unit: "%",  fmt:1 },
  ]},
  { name: "5. Distintivos Ambientales", layers: [
    { key: "pct_sin_distintivo_25", label: "% Sin distintivo 2025",          unit: "%",  fmt:1 },
    { key: "pct_eco_0_25",          label: "% Distintivo ECO + 0 (2025)",    unit: "%",  fmt:1 },
    { key: "delta_sin_distintivo",  label: "Evolución % sin distintivo (pp)", unit: "pp", fmt:1, isDelta:true },
  ]},
  { name: "6. ITV", layers: [
    { key: "pct_moto_sin_itv_25", label: "% Motocicletas sin ITV (<25 años)", unit: "%", fmt:1 },
  ]},
  { name: "7. Motocicletas", layers: [
    { key: "pct_motos_25",        label: "% Motos sobre parque total 2025",   unit: "%",    fmt:1 },
    { key: "motos_por_1000_25",   label: "Motos por 1.000 hab. 2025",         unit: "",     fmt:1 },
    { key: "antig_motos_25",      label: "Antigüedad media motos 2025",       unit: "años", fmt:1 },
    { key: "pct_motos_nuevas_25", label: "% Motos <4 años (2025)",            unit: "%",    fmt:1 },
    { key: "delta_antig_motos",   label: "Evolución antigüedad motos",        unit: "años", fmt:2, isDelta:true },
    { key: "delta_pct_motos",     label: "Evolución % motos en parque",       unit: "pp",   fmt:1, isDelta:true },
  ]},
  { name: "8. Camiones", layers: [
    { key: "pct_camiones_25",      label: "% Camiones sobre parque total 2025", unit: "%",    fmt:1 },
    { key: "cam_por_1000_25",      label: "Camiones por 1.000 hab. 2025",       unit: "",     fmt:1 },
    { key: "antig_camiones_25",    label: "Antigüedad media camiones 2025",     unit: "años", fmt:1 },
    { key: "pct_cam_nuevos_25",    label: "% Camiones <4 años (2025)",          unit: "%",    fmt:1 },
    { key: "pct_resto_sin_itv_25", label: "% Resto vehículos sin ITV (<25 años)", unit: "%", fmt:1 },
    { key: "delta_antig_camiones", label: "Evolución antigüedad camiones",      unit: "años", fmt:2, isDelta:true },
    { key: "delta_pct_camiones",   label: "Evolución % camiones en parque",     unit: "pp",   fmt:1, isDelta:true },
  ]},
];

// ── COLOR PALETTES ─────────────────────────────────────────────────────────
const PAL = {
  red2grn: ["#d73027","#f46d43","#fee08b","#d9ef8b","#1a9850"],  // rojo→verde: alto=bueno
  grn2red: ["#1a9850","#d9ef8b","#fee08b","#f46d43","#d73027"],  // verde→rojo: alto=malo
  div:     ["#d73027","#f46d43","#ffffbf","#91cf60","#1a9850"],  // div: negativo=rojo, positivo=verde
  div_inv: ["#1a9850","#91cf60","#ffffbf","#f46d43","#d73027"],  // div inv: positivo=rojo (empeorar)
};

function hexToRgb(h){return[parseInt(h.slice(1,3),16),parseInt(h.slice(3,5),16),parseInt(h.slice(5,7),16)]}
function rgbToHex(r,g,b){return'#'+[r,g,b].map(v=>Math.round(v).toString(16).padStart(2,'0')).join('')}
function lerp(a,b,t){return a+(b-a)*t}

function interpColor(pal, t){
  t = Math.max(0,Math.min(1,t));
  const n=pal.length, pos=t*(n-1), lo=Math.floor(pos), hi=Math.ceil(pos), f=pos-lo;
  const [r1,g1,b1]=hexToRgb(pal[lo]), [r2,g2,b2]=hexToRgb(pal[hi]);
  return rgbToHex(lerp(r1,r2,f),lerp(g1,g2,f),lerp(b1,b2,f));
}

function getColor(val, field){
  if(val===null||val===undefined||isNaN(val)) return '#3a3a4a';
  if(field==='pct_electrificado_25' && val===0) return '#000000';
  const sc=SCALES[field];
  if(!sc) return '#888';
  const t=(val-sc.vmin)/(sc.vmax-sc.vmin);
  return interpColor(PAL[sc.type]||PAL.seq_red, t);
}

// ── MAP SETUP ──────────────────────────────────────────────────────────────
const map = L.map('map',{center:[40.2,-3.7],zoom:6,preferCanvas:true,zoomControl:true});
L.tileLayer('https://{s}.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}{r}.png',{
  attribution:'&copy; OpenStreetMap &copy; CartoDB | Datos: DGT',
  subdomains:'abcd',maxZoom:19
}).addTo(map);

// ── STATE ──────────────────────────────────────────────────────────────────
let curTabIdx=0, curLayerIdx=0;
let curField=TABS[0].layers[0].key;
let gLayer=null;

// ── GEOJSON LAYER ──────────────────────────────────────────────────────────
function featureStyle(feat){
  const v=feat.properties[curField];
  return{fillColor:getColor(v,curField),fillOpacity:0.78,color:'#fff',weight:0.25,opacity:0.4};
}

function initLayer(){
  gLayer=L.geoJSON(GEOJSON,{
    style:featureStyle,
    onEachFeature(feat,layer){
      layer.on({mouseover:onHover,mouseout:onOut,click:onClickFeat});
    },
    renderer:L.canvas()
  }).addTo(map);
}

function refreshStyle(){if(gLayer)gLayer.setStyle(featureStyle);}

// ── TOOLTIP ────────────────────────────────────────────────────────────────
const ttEl=document.getElementById('tooltip');

function fmtVal(v,layerDef){
  if(v===null||v===undefined||isNaN(v)) return '<span style="color:#555">Sin datos</span>';
  const n=v.toFixed(layerDef.fmt||1);
  const sign=layerDef.isDelta&&v>0?'+':'';
  return `<strong>${sign}${n} ${layerDef.unit}</strong>`;
}

function onHover(e){
  const p=e.target.feature.properties;
  const ld=TABS[curTabIdx].layers[curLayerIdx];
  const val=p[curField];

  // extra context per tab
  let rows='';
  if(curTabIdx===0){
    rows=`
      <tr><td>Conductoras (nº)</td><td>${fmtNum(p.n_conductoras_25)}</td></tr>
      <tr><td>Ratio H/M</td><td>${fmtNum(p.ratio_hm_25,'x',2)}</td></tr>
      <tr><td>% Cond. 2020</td><td>${fmtNum(p.pct_conductoras_20,'%',1)}</td></tr>
      <tr><td>Evol. (pp)</td><td>${fmtDelta(p.delta_conductoras,'pp')}</td></tr>`;
  } else if(curTabIdx===1){
    rows=`
      <tr><td>Parque total</td><td>${fmtNum(p.parque_total_25,'',0)}</td></tr>
      <tr><td>% Turismos</td><td>${fmtNum(p.pct_turismos_25,'%',1)}</td></tr>
      <tr><td>Evol. veh/1000</td><td>${fmtDelta(p.delta_veh_por_1000,'')}</td></tr>`;
  } else if(curTabIdx===2){
    rows=`
      <tr><td>Antig. turismos</td><td>${fmtNum(p.antig_turismos_25,'años',1)}</td></tr>
      <tr><td>% Turismos <4a</td><td>${fmtNum(p.pct_tur_nuevos_25,'%',1)}</td></tr>
      <tr><td>Evol. antig.</td><td>${fmtDelta(p.delta_antig_turismos,'años')}</td></tr>`;
  } else if(curTabIdx===3){
    rows=`
      <tr><td>Elec. 2025</td><td>${fmtNum(p.pct_electrificado_25,'%',3)}</td></tr>
      <tr><td>Elec. 2020</td><td>${fmtNum(p.pct_electrificado_20,'%',3)}</td></tr>
      <tr><td>Evol. (pp)</td><td>${fmtDelta(p.delta_electrificado,'pp',3)}</td></tr>
      <tr><td>Ratio crecim.</td><td>${fmtNum(p.ratio_crecimiento_elec,'x',2)}</td></tr>`;
  } else if(curTabIdx===4){
    rows=`
      <tr><td>% Sin dist.</td><td>${fmtNum(p.pct_sin_distintivo_25,'%',1)}</td></tr>
      <tr><td>% ECO+0</td><td>${fmtNum(p.pct_eco_0_25,'%',1)}</td></tr>
      <tr><td>Evol. (pp)</td><td>${fmtDelta(p.delta_sin_distintivo,'pp')}</td></tr>`;
  } else if(curTabIdx===5){
    rows=`<tr><td>% Motos sin ITV</td><td>${fmtNum(p.pct_moto_sin_itv_25,'%',1)}</td></tr>`;
  } else if(curTabIdx===6){
    rows=`
      <tr><td>% Motos parque</td><td>${fmtNum(p.pct_motos_25,'%',1)}</td></tr>
      <tr><td>Motos/1.000 hab.</td><td>${fmtNum(p.motos_por_1000_25,'',1)}</td></tr>
      <tr><td>Antigüedad media</td><td>${fmtNum(p.antig_motos_25,'años',1)}</td></tr>
      <tr><td>% Motos <4 años</td><td>${fmtNum(p.pct_motos_nuevas_25,'%',1)}</td></tr>
      <tr><td>Evol. antigüedad</td><td>${fmtDelta(p.delta_antig_motos,'años')}</td></tr>`;
  } else if(curTabIdx===7){
    rows=`
      <tr><td>% Camiones parque</td><td>${fmtNum(p.pct_camiones_25,'%',1)}</td></tr>
      <tr><td>Camiones/1.000 hab.</td><td>${fmtNum(p.cam_por_1000_25,'',1)}</td></tr>
      <tr><td>Antigüedad media</td><td>${fmtNum(p.antig_camiones_25,'años',1)}</td></tr>
      <tr><td>% Camiones <4 años</td><td>${fmtNum(p.pct_cam_nuevos_25,'%',1)}</td></tr>
      <tr><td>% Resto sin ITV</td><td>${fmtNum(p.pct_resto_sin_itv_25,'%',1)}</td></tr>
      <tr><td>Evol. antigüedad</td><td>${fmtDelta(p.delta_antig_camiones,'años')}</td></tr>`;
  }

  ttEl.innerHTML=`
    <h3>${p['Municipio']||p.NAMEUNIT||''}</h3>
    <div class="prov">${p['Provincia']||''} · ${p['Comunidad Autónoma']||''}</div>
    <table>
      <tr><td>${ld.label}</td><td>${fmtVal(val,ld)}</td></tr>
      ${rows}
    </table>`;
  ttEl.style.display='block';
  positionTT(e.originalEvent);
  e.target.setStyle({weight:1.5,color:'#e94560',fillOpacity:0.9});
}

function onOut(e){
  ttEl.style.display='none';
  if(gLayer) gLayer.resetStyle(e.target);
}

document.addEventListener('mousemove',ev=>{if(ttEl.style.display!=='none') positionTT(ev);});

function positionTT(ev){
  const x=ev.clientX,y=ev.clientY,pad=14;
  const tw=ttEl.offsetWidth||290, th=ttEl.offsetHeight||120;
  const vw=window.innerWidth, vh=window.innerHeight;
  let lx=x+pad, ly=y+pad;
  if(lx+tw>vw) lx=x-tw-pad;
  if(ly+th>vh) ly=y-th-pad;
  ttEl.style.left=lx+'px';
  ttEl.style.top=ly+'px';
}

function fmtNum(v,unit='',dec=1){
  if(v===null||v===undefined||isNaN(v)) return '<span style="color:#555">N/D</span>';
  return v.toFixed(dec)+(unit?' '+unit:'');
}

function fmtDelta(v,unit='',dec=1){
  if(v===null||v===undefined||isNaN(v)) return '<span style="color:#555">N/D</span>';
  const sign=v>0?'+':'';
  const col=v>0?'#91cf60':v<0?'#f46d43':'#aaa';
  return `<span style="color:${col}">${sign}${v.toFixed(dec)}${unit?' '+unit:''}</span>`;
}

// ── CLICK POPUP ────────────────────────────────────────────────────────────
function onClickFeat(e){
  const p=e.target.feature.properties;
  const allLayers=TABS[curTabIdx].layers;
  let rows='';
  allLayers.forEach((ld,i)=>{
    const v=p[ld.key];
    const active=i===curLayerIdx?' class="highlight"':'';
    rows+=`<tr${active}><td>${ld.label}</td><td>${fmtVal(v,ld)}</td></tr>`;
  });

  // gender section always shown if tab 1
  let genderBlock='';
  if(curTabIdx===0){
    genderBlock=`<hr style="border-color:#0f3460;margin:8px 0">
      <div style="color:#e94560;font-size:0.9em;font-weight:600;margin-bottom:4px">Perspectiva de género</div>
      <table class="popup-table">
        <tr><td>Conductoras 2025</td><td>${fmtNum(p.pct_conductoras_25,'%',1)}</td></tr>
        <tr><td>Conductoras 2020</td><td>${fmtNum(p.pct_conductoras_20,'%',1)}</td></tr>
        <tr><td>Conductoras (nº)</td><td>${fmtNum(p.n_conductoras_25,'',0)}</td></tr>
        <tr><td>Ratio H/M</td><td>${fmtNum(p.ratio_hm_25,'x',2)}</td></tr>
        <tr><td>Evolución 2020→2025</td><td>${fmtDelta(p.delta_conductoras,'pp')}</td></tr>
      </table>`;
  }

  const html=`
    <div class="popup-title">${p['Municipio']||p.NAMEUNIT||''}</div>
    <div class="popup-sub">${p['Provincia']||''} &middot; ${p['Comunidad Autónoma']||''}</div>
    <table class="popup-table">${rows}</table>
    ${genderBlock}`;

  L.popup({maxWidth:320,className:'dgt-popup'})
    .setLatLng(e.latlng).setContent(html).openOn(map);
}

// ── TAB + LAYER SWITCHER ───────────────────────────────────────────────────
function setTab(idx){
  curTabIdx=idx; curLayerIdx=0;
  curField=TABS[idx].layers[0].key;
  document.querySelectorAll('.tab-btn').forEach((b,i)=>b.classList.toggle('active',i===idx));
  buildLayerBar();
  refreshStyle();
  updateLegend();
}

function setLayer(idx){
  curLayerIdx=idx;
  curField=TABS[curTabIdx].layers[idx].key;
  document.querySelectorAll('.lbtn').forEach((b,i)=>b.classList.toggle('active',i===idx));
  refreshStyle();
  updateLegend();
}

function buildTabBar(){
  const bar=document.getElementById('tabs');
  bar.innerHTML='';
  TABS.forEach((t,i)=>{
    const b=document.createElement('button');
    b.className='tab-btn'+(i===0?' active':'');
    b.textContent=t.name;
    b.onclick=()=>setTab(i);
    bar.appendChild(b);
  });
}

function buildLayerBar(){
  const bar=document.getElementById('layer-bar');
  bar.innerHTML='';
  TABS[curTabIdx].layers.forEach((l,i)=>{
    const b=document.createElement('button');
    b.className='lbtn'+(i===0?' active':'');
    b.textContent=l.label;
    b.onclick=()=>setLayer(i);
    bar.appendChild(b);
  });
}

// ── LEGEND ─────────────────────────────────────────────────────────────────
function updateLegend(){
  const ld=TABS[curTabIdx].layers[curLayerIdx];
  const sc=SCALES[curField];
  document.getElementById('leg-title').textContent=ld.label;
  const zeroEl=document.getElementById('legend-zero');
  zeroEl.style.display=curField==='pct_electrificado_25'?'flex':'none';
  if(!sc){document.getElementById('legend-bar').style.background='#555';return;}
  const pal=PAL[sc.type]||PAL.seq_red;
  document.getElementById('legend-bar').style.background=
    `linear-gradient(to right,${pal.join(',')})`;
  const fmt=v=>Math.abs(v)<10?v.toFixed(ld.fmt||1):Math.round(v);
  const sign_min=ld.isDelta&&sc.vmin>0?'+':'';
  const sign_max=ld.isDelta&&sc.vmax>0?'+':'';
  document.getElementById('leg-min').textContent=sign_min+fmt(sc.vmin)+(ld.unit?' '+ld.unit:'');
  document.getElementById('leg-max').textContent=sign_max+fmt(sc.vmax)+(ld.unit?' '+ld.unit:'');
}

// ── INIT ───────────────────────────────────────────────────────────────────
buildTabBar();
buildLayerBar();
initLayer();
updateLegend();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# HTML generation
# ---------------------------------------------------------------------------

def generate_html(geojson_str, scales_json, fecha, output_path):
    html = HTML_TEMPLATE
    html = html.replace("__GEOJSON__", geojson_str)
    html = html.replace("__SCALES__",  scales_json)
    html = html.replace("__FECHA__",   fecha)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    size_mb = os.path.getsize(output_path) / 1e6
    print(f"  Archivo generado: {output_path}  ({size_mb:.1f} MB)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    from datetime import date

    print("1/6  Cargando shapefile municipal...")
    gdf_shp = load_shapefile()
    print(f"     {len(gdf_shp)} municipios en shapefile")

    print("2/6  Cargando Excel 2025...")
    df2025_raw = load_excel(EXCEL_2025)
    print(f"     {len(df2025_raw)} filas (tras filtrar agregados)")

    print("3/6  Cargando Excel 2020...")
    df2020_raw = load_excel(EXCEL_2020)
    print(f"     {len(df2020_raw)} filas (tras filtrar agregados)")

    print("4/6  Calculando indicadores 2025...")
    df2025 = compute_indicators(df2025_raw, "_25")

    print("5/6  Calculando indicadores 2020 y evolución...")
    df2020 = compute_indicators(df2020_raw, "_20")
    df_final = compute_evolution(df2025, df2020)

    scales = compute_scales(df_final)
    scales_json = json.dumps(scales, ensure_ascii=False, separators=(",", ":"))

    print("6/6  Construyendo GeoJSON y generando HTML...")
    geojson_str = build_geojson(gdf_shp, df_final)
    fecha = date.today().isoformat()
    generate_html(geojson_str, scales_json, fecha, OUTPUT_HTML)

    print("\nListo. Abre el archivo en el navegador:")
    print(f"  {OUTPUT_HTML}")


if __name__ == "__main__":
    main()
