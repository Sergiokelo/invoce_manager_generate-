
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AgriTrace Interactive Polygon QC (Excel WKT -> GeoJSON + KML)
=============================================================

Fonctions clés :
- Script INTERACTIF en ligne de commande (prompts) pour choisir : fichier Excel, feuille, colonnes, sorties, CSR, etc.
- Validation/réparation des géométries (compat Shapely 1.x/2.x : make_valid si dispo sinon buffer(0))
- Option pour SUPPRIMER les trous (par défaut = Oui)
- Export GeoJSON (schéma restreint) + KML (toutes les colonnes Excel à l’identique)

Schéma GeoJSON (fixe) :
  Area, ProductionPlace, ProducerName, ProducerCountry, geometry
Correspondances par défaut (modifiable dans le prompt) :
  Area <- HECTARE
  ProductionPlace <- VILLAGE
  ProducerName <- CODE_PLANTATION
  ProducerCountry <- 'CD' (fixe)

CRS : si vous entrez "6", le script l’interprète comme EPSG:4326 (WGS84). KML attend WGS84.
Dépendances : pandas, shapely, openpyxl

Exécution :
  python agritrace_polygon_qc_cli.py
"""

import json
import math
import os
import sys
from pathlib import Path
from typing import List, Optional, Tuple, Union

# ---- Dépendances ----
try:
    import pandas as pd
except Exception as e:
    print("[ERREUR] pandas n'est pas installé. Installez-le : pip install pandas", file=sys.stderr)
    sys.exit(2)

try:
    from shapely import wkt as shapely_wkt
    from shapely.geometry import Polygon, MultiPolygon, mapping
    from shapely.ops import unary_union
    HAS_SHAPELY = True
    try:
        # Shapely 2.x
        from shapely.validation import make_valid as _make_valid
        def MAKE_VALID(g):
            try:
                return _make_valid(g)
            except Exception:
                return g.buffer(0)
    except Exception:
        # Shapely 1.x
        def MAKE_VALID(g):
            try:
                return g.buffer(0)
            except Exception:
                return g
except Exception as e:
    HAS_SHAPELY = False
    _IMPORT_ERR = e


def ask_path(prompt: str, default: Optional[str] = None) -> Path:
    while True:
        val = input(f"{prompt}" + (f" [{default}]" if default else "") + ": ").strip()
        if not val and default:
            p = Path(default)
        else:
            p = Path(val.strip().strip('"'))
        if p.exists():
            return p
        print(f"Chemin introuvable : {p}. Réessayez.")


def ask_str(prompt: str, default: Optional[str] = None) -> str:
    val = input(f"{prompt}" + (f" [{default}]" if default else "") + ": ").strip()
    return default if (not val and default is not None) else val


def ask_int(prompt: str, default: Optional[int] = None) -> int:
    while True:
        raw = input(f"{prompt}" + (f" [{default}]" if default is not None else "") + ": ").strip()
        if not raw and default is not None:
            return default
        try:
            return int(raw)
        except Exception:
            print("Entrez un entier valide.")


def ask_yes_no(prompt: str, default_yes: bool = True) -> bool:
    default = "O" if default_yes else "N"
    while True:
        raw = input(f"{prompt} (O/N) [{default}]: ").strip().lower()
        if raw == "" and default_yes:
            return True
        if raw == "" and not default_yes:
            return False
        if raw in ("o", "oui", "y", "yes"):
            return True
        if raw in ("n", "non", "no"):
            return False
        print("Répondez par O/N.")


def list_sheets(xlsx_path: Path) -> List[str]:
    try:
        xl = pd.ExcelFile(xlsx_path)
        return xl.sheet_names
    except Exception as e:
        print(f"[ERREUR] Impossible de lire les feuilles Excel : {e}", file=sys.stderr)
        sys.exit(2)


def pick_sheet(xlsx_path: Path) -> Optional[str]:
    sheets = list_sheets(xlsx_path)
    if not sheets:
        return None
    print("\nFeuilles détectées :")
    for i, s in enumerate(sheets):
        print(f"  [{i}] {s}")
    raw = input("Choisissez un index (Enter = 0 / première feuille) : ").strip()
    if raw == "":
        return sheets[0]
    try:
        idx = int(raw)
        return sheets[idx]
    except Exception:
        # Accepter un nom saisi
        if raw in sheets:
            return raw
        print("Saisie invalide. On prend la première feuille.")
        return sheets[0]


def infer_wkt_col(df: "pd.DataFrame", explicit: Optional[str]) -> str:
    if explicit and explicit in df.columns:
        return explicit
    lower_cols = {c.lower(): c for c in df.columns}
    for candidate in ["wkt_geom", "geometry", "wkt", "geom", "GEOMETRY", "WKT_GEOM"]:
        if candidate.lower() in lower_cols:
            return lower_cols[candidate.lower()]
    print("\nColonnes disponibles :", ", ".join(df.columns))
    while True:
        chosen = input("Nom de la colonne WKT (ex: wkt_geom) : ").strip()
        if chosen in df.columns:
            return chosen
        print("Colonne introuvable. Réessayez.")


def close_ring(coords: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    if not coords:
        return coords
    if coords[0] != coords[-1]:
        coords = coords + [coords[0]]
    return coords


def polygon_to_kml_coords(poly: Polygon) -> List[str]:
    exterior = list(poly.exterior.coords)
    exterior = close_ring([(float(x), float(y)) for x, y, *rest in exterior])
    return [f"{x},{y},0" for x, y in exterior]


def strip_holes(geom) -> Optional[Union[Polygon, MultiPolygon]]:
    """Supprime les trous ; retourne Polygon/MultiPolygon sans anneaux intérieurs."""
    if geom.geom_type == "Polygon":
        return Polygon(geom.exterior)
    elif geom.geom_type == "MultiPolygon":
        parts = [Polygon(p.exterior) for p in geom.geoms if not p.is_empty]
        try:
            u = unary_union(parts)
            return u
        except Exception:
            from shapely.geometry import MultiPolygon as MP
            return MP([p for p in parts if isinstance(p, Polygon) and not p.is_empty])
    else:
        try:
            g2 = geom.buffer(0)
            if g2.geom_type in ("Polygon", "MultiPolygon"):
                return strip_holes(g2)
            return None
        except Exception:
            return None


def fix_geometry(geom, remove_holes: bool = True):
    """Répare la géométrie (MAKE_VALID/buffer(0)). Option pour supprimer les trous."""
    try:
        geom = MAKE_VALID(geom)
    except Exception:
        try:
            geom = geom.buffer(0)
        except Exception:
            return None
    if geom.is_empty:
        return None
    if remove_holes:
        geom2 = strip_holes(geom)
        if geom2 is None or geom2.is_empty:
            return None
        try:
            geom2 = MAKE_VALID(geom2)
        except Exception:
            try:
                geom2 = geom2.buffer(0)
            except Exception:
                return None
        if geom2.is_empty:
            return None
        return geom2
    return geom


def write_geojson(features: List[dict], out_path: Union[str, Path]):
    fc = {"type": "FeatureCollection", "features": features}
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(fc, f, ensure_ascii=False)


def write_kml(rows: List[dict], polygons: List[Union[Polygon, MultiPolygon]], out_path: Union[str, Path], crs_epsg: int):
    def placemark_xml(props: dict, geom) -> str:
        name_val = None
        for key in ["ProducerName", "CODE_PLANTATION", "Code_Plantation", "code_plantation"]:
            if key in props and pd.notna(props[key]):
                name_val = str(props[key])
                break
        if name_val is None and props:
            key0 = next((k for k in props.keys() if k.lower() not in ("geometry",)), None)
            if key0:
                name_val = str(props[key0])
        if name_val is None:
            name_val = "Feature"

        ed = []
        for k, v in props.items():
            if k == "geometry":
                continue
            if v is None or (isinstance(v, float) and math.isnan(v)):
                v_str = ""
            else:
                v_str = str(v)
            ed.append(f'<Data name="{k}"><value>{v_str}</value></Data>')
        ed_xml = "<ExtendedData><SchemaData><SimpleData></SimpleData></SchemaData>" + "".join(ed) + "</ExtendedData>"

        def poly_xml(poly: Polygon) -> str:
            coords_str = " ".join(polygon_to_kml_coords(poly))
            return (
                "<Polygon>"
                "  <outerBoundaryIs>"
                "    <LinearRing>"
                f"      <coordinates>{coords_str}</coordinates>"
                "    </LinearRing>"
                "  </outerBoundaryIs>"
                "</Polygon>"
            )

        if geom.geom_type == "Polygon":
            geom_xml = poly_xml(geom)
        elif geom.geom_type == "MultiPolygon":
            parts = []
            for p in geom.geoms:
                if not p.is_empty:
                    parts.append(poly_xml(p))
            geom_xml = "<MultiGeometry>" + "".join(parts) + "</MultiGeometry>"
        else:
            return ""

        return (
            "<Placemark>"
            f"  <name>{name_val}</name>"
            f"  {ed_xml}"
            f"  {geom_xml}"
            "</Placemark>"
        )

    kml_header = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
        "  <!-- CRS supposé EPSG:4326 (lon/lat). Paramètre script: EPSG:{} -->\n"
        "  <Document>\n"
    ).format(crs_epsg)

    kml_footer = "  </Document>\n</kml>\n"

    placemarks = []
    for props, geom in zip(rows, polygons):
        if geom is None:
            continue
        pm = placemark_xml(props, geom)
        if pm:
            placemarks.append(pm)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(kml_header + "\n".join(placemarks) + kml_footer)


def main():
    print("===============================================")
    print(" AgriTrace — Interactive Polygon QC (Excel WKT)")
    print("===============================================\n")

    if not HAS_SHAPELY:
        print("[ERREUR] shapely n'est pas installé.", file=sys.stderr)
        print("Installez : pip install shapely", file=sys.stderr)
        print(f"Détail import: {_IMPORT_ERR}", file=sys.stderr)
        sys.exit(2)

    excel_path = ask_path("Chemin du fichier Excel (glisser-déposer possible)")
    sheet_name = pick_sheet(excel_path)
    print(f"Feuille choisie : {sheet_name}")

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"[ERREUR] Lecture Excel : {e}", file=sys.stderr)
        sys.exit(2)

    print("\nColonnes Excel détectées :")
    print(", ".join(df.columns))

    # Détection WKT
    wkt_col = infer_wkt_col(df, explicit=None)
    print(f"Colonne WKT retenue : {wkt_col}")

    # Mapping colonnes GeoJSON (avec suggestions)
    def pick_map_col(hint: str, default_name: str) -> Optional[str]:
        lower_map = {c.lower(): c for c in df.columns}
        suggested = lower_map.get(default_name.lower())
        return ask_str(f"Colonne pour {hint} (Enter = {suggested})", suggested)

    col_Area = pick_map_col("Area", "HECTARE")
    col_ProductionPlace = pick_map_col("ProductionPlace", "VILLAGE")
    col_ProducerName = pick_map_col("ProducerName", "CODE_PLANTATION")

    # Sorties par défaut dans le même dossier que l'Excel
    base = excel_path.with_suffix("")
    out_geojson = ask_str("Chemin de sortie GeoJSON", str(base) + "_cleaned.geojson")
    out_kml     = ask_str("Chemin de sortie KML",     str(base) + "_cleaned.kml")

    # CRS
    crs_val = ask_int("CRS EPSG (6 = EPSG:4326)", 6)
    crs_epsg = 4326 if crs_val == 6 else crs_val

    # Option trous
    remove_holes = ask_yes_no("Supprimer les trous (inner rings) ?", True)

    total = len(df)
    fixed = 0
    dropped = 0
    rows_for_kml = []
    geoms = []
    features_geojson = []

    print("\nTraitement en cours...\n")
    for idx, row in df.iterrows():
        wkt_val = row.get(wkt_col, None)
        if pd.isna(wkt_val) or (isinstance(wkt_val, str) and not wkt_val.strip()):
            dropped += 1
            continue
        try:
            geom = shapely_wkt.loads(str(wkt_val))
        except Exception:
            dropped += 1
            continue

        cleaned = fix_geometry(geom, remove_holes=remove_holes)
        if cleaned is None or cleaned.is_empty:
            dropped += 1
            continue

        if cleaned.wkt != geom.wkt:
            fixed += 1

        # KML : conserver toutes les colonnes Excel à l'identique
        rows_for_kml.append(row.to_dict())
        geoms.append(cleaned)

        # GeoJSON : schéma restreint
        props = {
            "Area": row.get(col_Area, None) if col_Area in df.columns else None,
            "ProductionPlace": row.get(col_ProductionPlace, None) if col_ProductionPlace in df.columns else None,
            "ProducerName": row.get(col_ProducerName, None) if col_ProducerName in df.columns else None,
            "ProducerCountry": "CD"
        }
        try:
            if props["Area"] is not None and props["Area"] != "":
                props["Area"] = float(props["Area"])
        except Exception:
            pass

        features_geojson.append({
            "type": "Feature",
            "properties": props,
            "geometry": mapping(cleaned)
        })

    # Écriture sorties
    write_geojson(features_geojson, out_geojson)
    write_kml(rows_for_kml, geoms, out_kml, crs_epsg=crs_epsg)

    print("\n===== RÉSULTAT =====")
    print(f"Lignes totales          : {total}")
    print(f"Géométries corrigées    : {fixed}")
    print(f"Lignes ignorées (vides/invalides) : {dropped}")
    print(f"GeoJSON écrit : {Path(out_geojson).resolve()}")
    print(f"KML écrit     : {Path(out_kml).resolve()}")
    print("====================\n")
    print("Terminé.")

if __name__ == "__main__":
    main()
