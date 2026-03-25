#!/usr/bin/env python3
"""
Generate Fondo Adaptación Dashboard for Power BI Desktop.
Creates:
  - Portfolio_Data.xlsx  (clean data for Power BI)
  - PortfolioDashboard.pbip + PBIP folder structure
"""

import pandas as pd
import json
import os
import uuid
from datetime import date

BASE_DIR = "/home/debugluis/projects/portfolio/PowerBI_Kevinsito"
SOURCE_DIR = os.path.join(BASE_DIR, "source")
SOURCE_CSV = os.path.join(SOURCE_DIR, "Proyectos_Fondo_Adaptacion_20250513 - CD.csv")

# ── Color palette ─────────────────────────────────────────────────
AZUL_OSCURO = "#1B3A5C"
AZUL_MEDIO  = "#4A90C4"
VERDE       = "#2E7D32"
AMARILLO    = "#F9A825"
ROJO        = "#C62828"
GRIS        = "#9E9E9E"
GRIS_FONDO  = "#F5F5F5"
BLANCO      = "#FFFFFF"
GRIS_TEXTO  = "#616161"

PALETTE = [AZUL_OSCURO, AZUL_MEDIO, VERDE, AMARILLO, ROJO, GRIS,
           "#3D6B7F", "#A23B72", "#95A5A6", "#2D936C"]


def uid():
    return str(uuid.uuid4())


# ════════════════════════════════════════════════════════════════════
# 1.  DATA PROCESSING
# ════════════════════════════════════════════════════════════════════

_DEPT_LAT = {
    "AMAZONAS": -4, "ANTIOQUIA": 7, "ARAUCA": 7, "ATLANTICO": 11,
    "BOGOTA  D.C.": 5, "BOLIVAR": 9, "BOYACA": 6, "CALDAS": 5,
    "CAQUETA": 1, "CASANARE": 5, "CAUCA": 2, "CESAR": 10,
    "CHOCO": 6, "CORDOBA": 9, "CUNDINAMARCA": 5, "GUAINIA": 3,
    "GUAVIARE": 2, "HUILA": 2, "LA GUAJIRA": 11, "MAGDALENA": 10,
    "META": 3, "NARI O": 1, "NORTE DE SANTANDER": 8,
    "PUTUMAYO": 1, "QUINDIO": 4, "RISARALDA": 5, "SANTANDER": 7,
    "SUCRE": 9, "TOLIMA": 4, "TRANSVERSAL": 5, "VALLE DEL CAUCA": 4,
    "VAUPES": 1, "VICHADA": 4,
}


def _fix_lat(val, dept=None):
    s = str(val).strip()
    if s in ("nan", "", "None"):
        return None
    neg = s.startswith("-")
    raw = s[1:] if neg else s
    parts = raw.split(".")
    if len(parts) == 3:
        opt1 = float(parts[0] + "." + parts[1] + parts[2])
        if neg:
            opt1 = -opt1
        if parts[0] == "1":
            digits = raw.replace(".", "")
            opt2 = float(digits[:2] + "." + digits[2:])
            if neg:
                opt2 = -opt2
            dept_lat = _DEPT_LAT.get(dept, 5)
            if abs(opt2 - dept_lat) < abs(opt1 - dept_lat):
                return opt2 if -5 <= opt2 <= 14 else opt1
        return opt1 if -5 <= opt1 <= 14 else None
    elif len(parts) == 2:
        num = float(raw)
        test = -num if neg else num
        if -5 <= test <= 14:
            return test
        digits = raw.replace(".", "")
        num = float(digits[0] + "." + digits[1:])
        if neg:
            num = -num
        return num if -5 <= num <= 14 else None
    elif len(parts) == 1:
        num = float(raw)
        if neg:
            num = -num
        return num if -5 <= num <= 14 else None
    return None


def _fix_lon(val):
    s = str(val).strip()
    if s in ("nan", "", "None"):
        return None
    neg = s.startswith("-")
    raw = s[1:] if neg else s
    parts = raw.split(".")
    if len(parts) == 3:
        num = float(parts[0] + "." + parts[1] + parts[2])
        test = -num if neg else num
        if -83 <= test <= -66:
            return test
        digits = raw.replace(".", "")
        num = float(digits[:2] + "." + digits[2:])
        if neg:
            num = -num
        return num if -83 <= num <= -66 else None
    elif len(parts) == 2:
        num = float(raw)
        test = -num if neg else num
        if -83 <= test <= -66:
            return test
        digits = raw.replace(".", "")
        num = float(digits[:2] + "." + digits[2:])
        if neg:
            num = -num
        return num if -83 <= num <= -66 else None
    elif len(parts) == 1:
        num = float(raw)
        if neg:
            num = -num
        return num if -83 <= num <= -66 else None
    return None


def process_data():
    print("  Reading source data from CSV...")
    df = pd.read_csv(SOURCE_CSV)
    print("  Fixing coordinates...")
    df["Latitud"] = df.apply(lambda r: _fix_lat(r["Latitud"], r["Departamento"]), axis=1)
    df["Longitud"] = df["Longitud"].apply(_fix_lon)
    if "Fecha_entrega" in df.columns:
        df["Fecha_entrega"] = (
            pd.to_datetime(df["Fecha_entrega"], dayfirst=True, errors="coerce")
            .dt.strftime("%Y-%m-%d")
        )
        df["Fecha_entrega"] = df["Fecha_entrega"].where(df["Fecha_entrega"].notna(), None)
    for c in ["Departamento", "Municipio", "Sector", "Proyecto",
              "Intervencion", "Nombre_Proyecto", "Tipo_Producto",
              "Estado_Proyecto", "U_Medida", "Cod_Intervencion"]:
        if c in df.columns:
            df[c] = df[c].astype(str).replace("nan", "")
    df["Cod_Proyecto"] = df["Cod_Proyecto"].astype(str)
    # Severity order: worst (1) → best (9)
    estado_orden = {
        "Suspendido": 1,
        "Presunto Incumplimiento": 2,
        "Vencido en proceso de soluciOn": 3,
        "Despriorizado": 4,
        "Terminado sin entregar": 5,
        "En ejecuci n": 6,
        "Contratado sin iniciar": 7,
        "Por contratar": 8,
        "Entregado": 9,
    }
    df["Estado_Orden"] = df["Estado_Proyecto"].map(estado_orden).fillna(5).astype(int)
    valid_lat = df["Latitud"].notna().sum()
    valid_lon = df["Longitud"].notna().sum()
    print(f"  Coordinates fixed: {valid_lat}/{len(df)} lat, {valid_lon}/{len(df)} lon valid")
    return {"Proyectos": df}


# ════════════════════════════════════════════════════════════════════
# 2.  EXCEL OUTPUT
# ════════════════════════════════════════════════════════════════════
def create_excel(data):
    os.makedirs(SOURCE_DIR, exist_ok=True)
    out = os.path.join(SOURCE_DIR, "Portfolio_Data.xlsx")
    print(f"  Writing {out}")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for name, df in data.items():
            df.to_excel(w, sheet_name=name, index=False)
    return out


# ════════════════════════════════════════════════════════════════════
# 3.  POWER BI THEME
# ════════════════════════════════════════════════════════════════════
def build_theme():
    return {
        "name": "Fondo Adaptacion",
        "dataColors": PALETTE,
        "background": GRIS_FONDO,
        "foreground": AZUL_OSCURO,
        "tableAccent": AZUL_MEDIO,
        "maximum": VERDE,
        "center": AMARILLO,
        "minimum": ROJO,
        "good": VERDE,
        "neutral": AMARILLO,
        "bad": ROJO,
        "textClasses": {
            "callout": {"fontSize": 28, "fontFace": "Segoe UI Light", "color": AZUL_MEDIO},
            "title":   {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": AZUL_OSCURO},
            "header":  {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": AZUL_OSCURO},
            "label":   {"fontSize": 10, "fontFace": "Segoe UI", "color": GRIS_TEXTO},
        },
        "visualStyles": {
            "*": {
                "*": {
                    "background": [{"color": {"solid": {"color": BLANCO}}, "transparency": 0}],
                    "border":     [{"show": False}],
                    "title":      [{"show": True, "fontColor": {"solid": {"color": AZUL_OSCURO}},
                                    "fontSize": 12, "fontFamily": "Segoe UI Semibold"}],
                    "labels":     [{"fontSize": 10, "fontFamily": "Segoe UI",
                                    "color": {"solid": {"color": GRIS_TEXTO}}}],
                    "outspace":   [{"color": {"solid": {"color": GRIS_FONDO}}, "transparency": 0}],
                }
            },
            "page": {
                "*": {
                    "background": [{"color": {"solid": {"color": GRIS_FONDO}}, "transparency": 0}],
                }
            },
        },
    }




# ════════════════════════════════════════════════════════════════════
# 6.  DATA MODEL SCHEMA  (TOM JSON)
# ════════════════════════════════════════════════════════════════════
def build_model_schema():
    def col(name, dtype="string", summarize="none"):
        return {
            "name": name, "dataType": dtype, "sourceColumn": name,
            "lineageTag": uid(), "summarizeBy": summarize,
            "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}],
        }

    measures = [
        ("Inversion Total",       "SUM(Proyectos[Total])",        "Currency"),
        ("Total Proyectos",       "COUNTROWS(Proyectos)",         "WholeNumber"),
        ("Beneficiarios",         "SUM(Proyectos[Cantidad])",     "WholeNumber"),
        ("Pct Entrega",
         'DIVIDE(CALCULATE(COUNTROWS(Proyectos), Proyectos[Estado_Proyecto] = "Entregado"), COUNTROWS(Proyectos), 0)',
         "Percentage"),
        ("Conteo Departamentos",  "DISTINCTCOUNT(Proyectos[Departamento])", "WholeNumber"),
        ("Conteo Municipios",     "DISTINCTCOUNT(Proyectos[Municipio])",    "WholeNumber"),
    ]

    def make_measures(measure_list):
        out = []
        for name, expr, fmt in measure_list:
            m = {"name": name, "expression": expr, "lineageTag": uid()}
            if fmt == "Currency":
                m["formatString"] = "$#,0;($#,0);$#,0"
            elif fmt == "Percentage":
                m["formatString"] = "0.0%;-0.0%;0.0%"
            elif fmt == "WholeNumber":
                m["formatString"] = "#,0"
            out.append(m)
        return out

    model = {
        "name": "SemanticModel",
        "compatibilityLevel": 1600,
        "model": {
            "culture": "en-US",
            "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "sourceQueryCulture": "en-US",
            "tables": [{
                "name": "Proyectos",
                "lineageTag": uid(),
                "columns": [
                    col("Departamento"), col("Municipio"), col("Sector"),
                    col("Cod_Proyecto"), col("Proyecto"), col("Cod_Intervencion"),
                    col("Intervencion"), col("Nombre_Proyecto"), col("Tipo_Producto"),
                    col("Estado_Proyecto"), col("Fecha_entrega"),
                    col("Cantidad", "int64", "sum"), col("U_Medida"),
                    col("Latitud", "double", "none"), col("Longitud", "double", "none"),
                    col("Total", "double", "sum"), col("Estado_Orden", "int64", "none"),
                ],
                "measures": make_measures(measures),
                "partitions": [{
                    "name": "Proyectos", "mode": "import",
                    "source": {"type": "m", "expression": [
                        "let",
                        "    Source = Excel.Workbook(File.Contents(#\"ExcelFilePath\"), null, true),",
                        "    Sheet = Source{[Item=\"Proyectos\",Kind=\"Sheet\"]}[Data],",
                        "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                        "    Typed = Table.TransformColumnTypes(Headers, {",
                        "        {\"Departamento\", type text}, {\"Municipio\", type text},",
                        "        {\"Sector\", type text}, {\"Cod_Proyecto\", type text},",
                        "        {\"Proyecto\", type text}, {\"Cod_Intervencion\", type text},",
                        "        {\"Intervencion\", type text}, {\"Nombre_Proyecto\", type text},",
                        "        {\"Tipo_Producto\", type text}, {\"Estado_Proyecto\", type text},",
                        "        {\"Estado_Orden\", Int64.Type},",
                        "        {\"Fecha_entrega\", type text},",
                        "        {\"Cantidad\", Int64.Type}, {\"U_Medida\", type text},",
                        "        {\"Latitud\", type number}, {\"Longitud\", type number},",
                        "        {\"Total\", type number}",
                        "    })",
                        "in",
                        "    Typed",
                    ]},
                }],
            }],
            "relationships": [],
            "annotations": [
                {"name": "PBI_QueryOrder", "value": json.dumps(["ExcelFilePath", "Proyectos"])},
                {"name": "PBIDesktopVersion", "value": "2.127.0"},
            ],
            "expressions": [{
                "name": "ExcelFilePath", "kind": "m", "lineageTag": uid(),
                "expression": [
                    '"\\\\wsl.localhost\\Ubuntu-24.04\\home\\debugluis\\projects\\portfolio\\PowerBI_Kevinsito\\source\\Portfolio_Data.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
                ],
                "annotations": [
                    {"name": "PBI_NavigationStepName", "value": "Navigation"},
                    {"name": "PBI_ResultType", "value": "Text"},
                ],
            }],
        },
    }
    return model


# ════════════════════════════════════════════════════════════════════
# 7.  REPORT LAYOUT (visuals)
# ════════════════════════════════════════════════════════════════════

def _from_ref(alias, entity):
    return {"Name": alias, "Entity": entity, "Type": 0}

def _sel_measure(alias, table, measure):
    return {"Measure": {"Expression": {"SourceRef": {"Source": alias}}, "Property": measure},
            "Name": f"{table}.{measure}"}

def _sel_column(alias, table, column):
    return {"Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": column},
            "Name": f"{table}.{column}"}

def _sel_agg(alias, table, column, func=0):
    func_names = {0: "Sum", 1: "Avg", 2: "Min", 3: "Max", 5: "Count", 6: "CountNonNull"}
    fname = func_names.get(func, "Sum")
    return {
        "Aggregation": {
            "Expression": {"Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": column}},
            "Function": func,
        },
        "Name": f"{fname}({table}.{column})",
    }


def _visual_base(vid, x, y, w, h, config):
    return {
        "x": x, "y": y, "z": 0, "width": w, "height": h,
        "config": json.dumps(config, separators=(",", ":")),
        "filters": "[]", "tabOrder": 0,
    }


def make_textbox(vid, x, y, w, h, text, font_size=20, color=AZUL_OSCURO,
                 bold=False, bg_color=None, align="left"):
    font_family = "Segoe UI Bold" if bold else "Segoe UI Semibold"
    paragraphs = [{"textRuns": [{"value": text, "textStyle": {
        "fontFamily": font_family, "fontSize": f"{font_size}px", "color": color,
    }}], "horizontalTextAlignment": align}]
    objects = {"general": [{"properties": {
        "paragraphs": {"expr": {"Literal": {"Value": json.dumps(paragraphs)}}},
    }}]}
    if bg_color:
        objects["background"] = [{"properties": {
            "color": {"solid": {"color": bg_color}}, "transparency": 0,
        }}]
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {"visualType": "textbox", "objects": objects, "vcObjects": {}},
    }
    return _visual_base(vid, x, y, w, h, config)


def make_shape(vid, x, y, w, h, bg_color):
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "shape",
            "objects": {
                "general": [{"properties": {"maintainAspectRatio": {"expr": {"Literal": {"Value": "false"}}}}}],
                "fill": [{"properties": {"fillColor": {"solid": {"color": bg_color}}, "transparency": 0}}],
                "line": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
            },
            "vcObjects": {},
        },
    }
    return {"x": x, "y": y, "z": -1, "width": w, "height": h,
            "config": json.dumps(config, separators=(",", ":")), "filters": "[]", "tabOrder": 0}


def make_slicer(vid, x, y, w, h, table, column, title):
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "slicer",
            "projections": {"Values": [{"queryRef": f"{table}.{column}"}]},
            "prototypeQuery": {"Version": 2, "From": [_from_ref("t", table)],
                               "Select": [_sel_column("t", table, column)]},
            "objects": {"data": [{"properties": {"mode": {"expr": {"Literal": {"Value": "'Dropdown'"}}}}}]},
            "vcObjects": {"title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                "fontColor": {"expr": {"Literal": {"Value": f"'{AZUL_OSCURO}'"}}},
            }}]},
        },
    }
    return _visual_base(vid, x, y, w, h, config)


def make_card(vid, x, y, w, h, table, measure, title):
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": f"{table}.{measure}"}]},
            "prototypeQuery": {"Version": 2, "From": [_from_ref("t", table)],
                               "Select": [_sel_measure("t", table, measure)]},
            "objects": {},
            "vcObjects": {"title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "fontSize": {"expr": {"Literal": {"Value": "10D"}}},
                "fontColor": {"expr": {"Literal": {"Value": f"'{GRIS_TEXTO}'"}}},
            }}]},
        },
    }
    return _visual_base(vid, x, y, w, h, config)


def make_stacked_bar(vid, x, y, w, h, table, cat_col, series_col, measure, title):
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "clusteredBarChart",
            "projections": {
                "Category": [{"queryRef": f"{table}.{cat_col}"}],
                "Y": [{"queryRef": f"{table}.{measure}"}],
                "Series": [{"queryRef": f"{table}.{series_col}"}],
            },
            "prototypeQuery": {
                "Version": 2, "From": [_from_ref("t", table)],
                "Select": [
                    _sel_column("t", table, cat_col),
                    _sel_column("t", table, series_col),
                    _sel_measure("t", table, measure),
                ],
                "OrderBy": [{
                    "Direction": 2,
                    "Expression": {"Measure": {
                        "Expression": {"SourceRef": {"Source": "t"}},
                        "Property": measure,
                    }},
                }],
            },
            "objects": {
                "legend": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "position": {"expr": {"Literal": {"Value": "'Bottom'"}}},
                }}],
                "categoryAxis": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                }}],
                "valueAxis": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "gridlineShow": {"expr": {"Literal": {"Value": "false"}}},
                }}],
            },
            "sort": [{"Direction": 2, "Expression": {"Measure": {
                "Expression": {"SourceRef": {"Source": "t"}}, "Property": measure}}}],
            "vcObjects": {"title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "fontColor": {"expr": {"Literal": {"Value": f"'{AZUL_OSCURO}'"}}},
                "fontSize": {"expr": {"Literal": {"Value": "12D"}}},
            }}]},
        },
    }
    return _visual_base(vid, x, y, w, h, config)


def make_map(vid, x, y, w, h, table, lat_col, lon_col, size_measure, title):
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "map",
            "projections": {
                "X": [{"queryRef": f"{table}.{lon_col}"}],
                "Y": [{"queryRef": f"{table}.{lat_col}"}],
                "Size": [{"queryRef": f"{table}.{size_measure}"}],
            },
            "prototypeQuery": {
                "Version": 2, "From": [_from_ref("t", table)],
                "Select": [
                    _sel_column("t", table, lat_col),
                    _sel_column("t", table, lon_col),
                    _sel_measure("t", table, size_measure),
                ],
            },
            "objects": {},
            "vcObjects": {"title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "fontColor": {"expr": {"Literal": {"Value": f"'{AZUL_OSCURO}'"}}},
                "fontSize": {"expr": {"Literal": {"Value": "12D"}}},
            }}]},
        },
    }
    return _visual_base(vid, x, y, w, h, config)


def make_donut(vid, x, y, w, h, table, cat_col, measure, title):
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "donutChart",
            "projections": {
                "Category": [{"queryRef": f"{table}.{cat_col}"}],
                "Y": [{"queryRef": f"{table}.{measure}"}],
            },
            "prototypeQuery": {
                "Version": 2, "From": [_from_ref("t", table)],
                "Select": [
                    _sel_column("t", table, cat_col),
                    _sel_measure("t", table, measure),
                ],
            },
            "objects": {
                "legend": [{"properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "position": {"expr": {"Literal": {"Value": "'Bottom'"}}},
                }}],
            },
            "vcObjects": {"title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "fontColor": {"expr": {"Literal": {"Value": f"'{AZUL_OSCURO}'"}}},
                "fontSize": {"expr": {"Literal": {"Value": "12D"}}},
            }}]},
        },
    }
    return _visual_base(vid, x, y, w, h, config)


def make_table_filtered(vid, x, y, w, h, table, columns, measures, title,
                        filter_col=None, exclude_values=None,
                        sort_measure=None, sort_desc=True,
                        sort_column=None, sort_col_asc=True):
    projs = ([{"queryRef": f"{table}.{c}"} for c in columns]
             + [{"queryRef": f"{table}.{m}"} for m in measures])
    selects = ([_sel_column("t", table, c) for c in columns]
               + [_sel_measure("t", table, m) for m in measures])
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h, "tabOrder": 0}}],
        "singleVisual": {
            "visualType": "tableEx",
            "projections": {"Values": projs},
            "prototypeQuery": {"Version": 2, "From": [_from_ref("t", table)], "Select": selects},
            "objects": {
                "grid": [{"properties": {
                    "gridVertical": {"expr": {"Literal": {"Value": "false"}}},
                    "gridHorizontal": {"expr": {"Literal": {"Value": "false"}}},
                }}],
                "columnHeaders": [{"properties": {
                    "fontColor": {"solid": {"color": BLANCO}},
                    "backColor": {"solid": {"color": AZUL_OSCURO}},
                    "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                }}],
                "values": [{"properties": {
                    "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                    "fontColor": {"solid": {"color": GRIS_TEXTO}},
                }}],
            },
            "vcObjects": {"title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
                "fontColor": {"expr": {"Literal": {"Value": f"'{AZUL_OSCURO}'"}}},
                "fontSize": {"expr": {"Literal": {"Value": "12D"}}},
            }}]},
        },
    }
    if sort_measure:
        config["singleVisual"]["sort"] = [{"Direction": 2 if sort_desc else 1,
            "Expression": {"Measure": {"Expression": {"SourceRef": {"Source": "t"}},
                                       "Property": sort_measure}}}]
    if sort_column:
        config["singleVisual"]["sort"] = [{"Direction": 1 if sort_col_asc else 2,
            "Expression": {"Column": {"Expression": {"SourceRef": {"Source": "t"}},
                                      "Property": sort_column}}}]
    filters = []
    if filter_col and exclude_values:
        conds = [{"Comparison": {"ComparisonKind": 1,
                  "Left": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": filter_col}},
                  "Right": {"Literal": {"Value": f"'{v}'"}}}} for v in exclude_values]
        if len(conds) == 1:
            where_cond = conds[0]
        else:
            where_cond = {"And": {"Left": conds[0], "Right": conds[1]}}
            for c in conds[2:]:
                where_cond = {"And": {"Left": where_cond, "Right": c}}
        filters.append({
            "name": f"Filter_{filter_col}",
            "expression": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": filter_col}},
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": table, "Type": 0}],
                       "Where": [{"Condition": where_cond}]},
            "type": "Advanced",
            "howCreated": 0,
            "isHiddenInViewMode": True,
        })
    result = _visual_base(vid, x, y, w, h, config)
    if filters:
        result["filters"] = json.dumps(filters, separators=(",", ":"))
    return result


def build_report_layout():
    W, H = 1440, 1080
    report_id = uid()
    today_str = date.today().strftime("%d/%m/%Y")
    M = 15  # margin
    GAP = 12  # gap between elements

    # Column helpers
    half_w = (W - 2 * M - GAP) // 2  # 696
    quarter_w = (W - 2 * M - 3 * GAP) // 4  # 339

    visuals = [
        # ── Header bar (dark blue) ──
        make_shape("shape_header", 0, 0, W, 56, AZUL_OSCURO),
        make_textbox("txt_title", M, 10, 800, 38,
                     "Portafolio de Proyectos Fondo Adaptacion",
                     font_size=22, color=BLANCO, bold=True),
        make_textbox("txt_date", W - M - 250, 18, 250, 25,
                     f"Actualizado: {today_str}",
                     font_size=10, color=BLANCO, align="right"),

        # ── Filters row (y=68) ──
        make_slicer("slicer_depto",  M,                    68, quarter_w, 45, "Proyectos", "Departamento",  "Departamento"),
        make_slicer("slicer_sector", M + quarter_w + GAP,  68, quarter_w, 45, "Proyectos", "Sector",         "Sector"),
        make_slicer("slicer_estado", M + 2*(quarter_w+GAP),68, quarter_w, 45, "Proyectos", "Estado_Proyecto", "Estado"),
        make_slicer("slicer_fecha",  M + 3*(quarter_w+GAP),68, quarter_w, 45, "Proyectos", "Fecha_entrega",   "Fecha Entrega"),

        # ── KPI Cards (y=125) ──
        make_card("card_inversion",  M,                    125, quarter_w, 85, "Proyectos", "Inversion Total",  "Inversion Total (COP)"),
        make_card("card_proyectos",  M + quarter_w + GAP,  125, quarter_w, 85, "Proyectos", "Total Proyectos",  "Total Proyectos"),
        make_card("card_benefic",    M + 2*(quarter_w+GAP),125, quarter_w, 85, "Proyectos", "Beneficiarios",    "Beneficiarios"),
        make_card("card_pct",        M + 3*(quarter_w+GAP),125, quarter_w, 85, "Proyectos", "Pct Entrega",      "Porcentaje de Entrega"),

        # ── Middle row (y=222) ──
        make_stacked_bar("bar_sector", M, 222, half_w, 360, "Proyectos",
                         "Sector", "Estado_Proyecto", "Inversion Total",
                         "Inversion por Sector y Estado"),
        make_map("map_colombia", M + half_w + GAP, 222, half_w, 360, "Proyectos",
                 "Latitud", "Longitud", "Inversion Total",
                 "Distribucion Geografica de Inversion"),

        # ── Bottom row (y=594) ──
        make_donut("donut_estado", M, 594, half_w, 474, "Proyectos",
                   "Estado_Proyecto", "Total Proyectos",
                   "Estado del Portafolio"),
        make_table_filtered("tbl_atencion", M + half_w + GAP, 594, half_w, 474, "Proyectos",
                            ["Nombre_Proyecto", "Departamento", "Sector", "Estado_Proyecto"],
                            ["Inversion Total"],
                            "Proyectos que Requieren Atencion",
                            filter_col="Estado_Proyecto",
                            exclude_values=["Entregado", "Despriorizado"],
                            sort_column="Estado_Orden", sort_col_asc=True),
    ]

    report_config = {
        "version": "5.53",
        "themeCollection": {"baseTheme": {"name": "CY24SU02", "version": "5.53", "type": 2}},
        "activeSectionIndex": 0,
        "defaultDrillFilterOtherVisuals": True,
        "linguisticSchemaSyncVersion": 2,
        "settings": {"useStaleDataOnError": True, "filterPaneEnabled": True,
                      "navContentPaneEnabled": True},
        "objects": {"outspace": [{"properties": {"color": {"solid": {"color": GRIS_FONDO}}}}]},
    }

    def page(name, display_name, ordinal, page_visuals):
        page_config = {"layouts": [{"id": 0, "position": {}}], "visibility": 0}
        return {
            "name": name, "displayName": display_name, "displayOption": 1,
            "filters": "[]", "height": H, "width": W, "ordinal": ordinal,
            "config": json.dumps(page_config, separators=(",", ":")),
            "visualContainers": page_visuals,
        }

    layout = {
        "id": 0, "reportId": report_id,
        "config": json.dumps(report_config, separators=(",", ":")),
        "filters": "[]", "layoutOptimization": 0,
        "publicCustomVisuals": [], "resourcePackages": [],
        "sections": [page("ReportSection_Main", "Portafolio Fondo Adaptacion", 0, visuals)],
        "theme": json.dumps(build_theme(), separators=(",", ":")),
        "pods": [],
    }
    return layout


# ════════════════════════════════════════════════════════════════════
# 9.  STANDALONE THEME FILE
# ════════════════════════════════════════════════════════════════════
def save_theme():
    path = os.path.join(BASE_DIR, "portfolio_theme.json")
    with open(path, "w") as f:
        json.dump(build_theme(), f, indent=2)
    print(f"  -> {path}")


# ════════════════════════════════════════════════════════════════════
# 10. PBIP PROJECT  (folder-based, pure JSON, no binary format)
# ════════════════════════════════════════════════════════════════════
def create_pbip_project():
    proj_dir = BASE_DIR
    sem_dir  = os.path.join(proj_dir, "PortfolioDashboard.SemanticModel")
    rep_dir  = os.path.join(proj_dir, "PortfolioDashboard.Report")
    for d in [sem_dir, os.path.join(sem_dir, ".pbi"),
              rep_dir, os.path.join(rep_dir, ".pbi")]:
        os.makedirs(d, exist_ok=True)

    def write_json(path, obj):
        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
            json.dump(obj, f, indent=2, ensure_ascii=False)
        print(f"    {os.path.relpath(path, BASE_DIR)}")

    write_json(os.path.join(proj_dir, "PortfolioDashboard.pbip"), {
        "version": "1.0",
        "artifacts": [{"report": {"path": "PortfolioDashboard.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })
    write_json(os.path.join(sem_dir, "definition.pbism"), {"version": "1.0", "settings": {}})
    write_json(os.path.join(sem_dir, ".pbi", "localSettings.json"), {"version": "1.0"})
    write_json(os.path.join(sem_dir, "model.bim"), build_model_schema())
    write_json(os.path.join(rep_dir, "definition.pbir"), {
        "version": "1.0",
        "datasetReference": {"byPath": {"path": "../PortfolioDashboard.SemanticModel"}, "byConnection": None},
    })
    write_json(os.path.join(rep_dir, ".pbi", "localSettings.json"), {"version": "1.0"})
    write_json(os.path.join(rep_dir, "report.json"), build_report_layout())
    theme_dir = os.path.join(rep_dir, "StaticResources", "SharedResources", "BaseThemes")
    os.makedirs(theme_dir, exist_ok=True)
    write_json(os.path.join(theme_dir, "CY24SU02.json"), build_theme())
    print(f"  -> PBIP project: {proj_dir}")
    return proj_dir


# ════════════════════════════════════════════════════════════════════
#     MAIN
# ════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  Fondo Adaptacion Dashboard Generator for Power BI")
    print("=" * 60)
    print("\n[1/4] Processing data...")
    data = process_data()
    print("[2/4] Creating Excel data file...")
    create_excel(data)
    print("[3/4] Building PBIP project (folder-based)...")
    create_pbip_project()
    print("[4/4] Saving standalone theme...")
    save_theme()
    print("\n" + "=" * 60)
    print("  DONE!")
    print("=" * 60)
    print("  1. source/Portfolio_Data.xlsx")
    print("  2. PortfolioDashboard.pbip -> abrir en Power BI Desktop")
    print("  3. portfolio_theme.json")
    print()
    print("  Rechazar TMDL y PBIR upgrades al abrir.")
    print("=" * 60)


if __name__ == "__main__":
    main()
