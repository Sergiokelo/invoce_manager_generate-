import argparse
import json
from pathlib import Path
import sys

import pandas as pd
from shapely import wkt as shapely_wkt
from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, LinearRing, GeometryCollection
from shapely.ops import polygonize, unary_union
from shapely.validation import explain_validity
from shapely.geometry import mapping as shp_mapping
from shapely.geometry.polygon import orient as orient_polygon

def lines_to_polygons(geom):
    try:
        if isinstance(geom, LineString):
            lines = [geom]
        elif isinstance(geom, MultiLineString):
            lines = list(geom.geoms)
        else:
            return None

        polys = list(polygonize(lines))
        if polys:
            return polys[0] if len(polys) == 1 else MultiPolygon(polys)

        if isinstance(geom, LineString):
            coords = list(geom.coords)
            if coords and coords[0] != coords[-1]:
                coords.append(coords[0])
            try:
                ring = LinearRing(coords)
                poly = Polygon(ring)
                if not poly.is_empty:
                    return poly
            except Exception:
                return None
        return None
    except Exception:
        return None

def make_valid_fallback(g):
    notes = []
    if g.is_valid and isinstance(g, (Polygon, MultiPolygon)):
        return g, notes
    try:
        g1 = unary_union([g])
        if isinstance(g1, (Polygon, MultiPolygon)) and g1.is_valid:
            notes.append("Fixed via unary_union")
            return g1, notes
    except Exception:
        pass
    try:
        g2 = g.buffer(0)
        if isinstance(g2, (Polygon, MultiPolygon)) and g2.is_valid:
            notes.append("Fixed via buffer(0)")
            return g2, notes
        if isinstance(g2, GeometryCollection):
            polys = [p for p in g2.geoms if isinstance(p, (Polygon, MultiPolygon))]
            if polys:
                merged = polys[0]
                for p in polys[1:]:
                    merged = merged.union(p)
                if isinstance(merged, (Polygon, MultiPolygon)) and merged.is_valid:
                    notes.append("Extracted polygons from GeometryCollection after buffer(0)")
                    return merged, notes
    except Exception:
        pass
    try:
        boundary = g.boundary
        polys = list(polygonize(boundary))
        if polys:
            merged = polys[0]
            for p in polys[1:]:
                merged = merged.union(p)
            if isinstance(merged, (Polygon, MultiPolygon)) and merged.is_valid:
                notes.append("Fixed via polygonize(boundary)")
                return merged, notes
    except Exception:
        pass
    notes.append(f"Still invalid: {explain_validity(g)}")
    return None, notes

def remove_holes(geom):
    if isinstance(geom, Polygon):
        return Polygon(geom.exterior)
    if isinstance(geom, MultiPolygon):
        return MultiPolygon([Polygon(p.exterior) for p in geom.geoms])
    return geom

def coerce_to_multipolygon(g):
    if isinstance(g, Polygon):
        return MultiPolygon([g])
    return g

def largest_part_only(g):
    if isinstance(g, MultiPolygon):
        parts = list(g.geoms)
        if not parts:
            return g
        parts_sorted = sorted(parts, key=lambda p: p.area if isinstance(p, Polygon) else 0.0, reverse=True)
        return parts_sorted[0]
    return g

def fix_geometry(geom, drop_holes=False, force_multipolygon=False, largest_only=False, orient_ccw=False):
    notes = []
    g = geom
    if isinstance(g, (LineString, MultiLineString)):
        conv = lines_to_polygons(g)
        if conv is not None:
            notes.append("Converted from line(s) to polygon")
            g = conv
        else:
            return None, ["Could not convert lines to polygon"]
    if not isinstance(g, (Polygon, MultiPolygon)):
        return None, ["Not a Polygon/MultiPolygon"]
    if drop_holes:
        g = remove_holes(g)
        notes.append("Holes removed")
    if not g.is_valid:
        gv, more = make_valid_fallback(g)
        if gv is None:
            return None, more
        notes += more
        g = gv
    if largest_only and isinstance(g, MultiPolygon):
        g = largest_part_only(g)
        notes.append("Kept largest part only (single Polygon)")
    if force_multipolygon:
        before = g.geom_type
        g = coerce_to_multipolygon(g)
        if g.geom_type == "MultiPolygon" and before != "MultiPolygon":
            notes.append("Coerced to MultiPolygon")
    if orient_ccw:
        if isinstance(g, Polygon):
            g = orient_polygon(g, sign=1.0)
            notes.append("Oriented CCW (Polygon)")
        elif isinstance(g, MultiPolygon):
            g = MultiPolygon([orient_polygon(p, sign=1.0) for p in g.geoms])
            notes.append("Oriented CCW (MultiPolygon)")
    return g, notes

def process_excel(in_path, out_dir, wkt_col="wkt_geom", sheet_name=0, drop_holes=False, force_multipolygon=False, uppercase_type=False, largest_only=False, orient_ccw=False):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(in_path, sheet_name=sheet_name)
    rows = []
    has_holes_col = []
    part_count_col = []
    area_col = []

    for _, row in df.iterrows():
        raw = row.get(wkt_col)
        before_type = after_type = fixed_wkt = None
        is_valid = False
        notes = []
        try:
            geom = shapely_wkt.loads(str(raw)) if pd.notna(raw) else None
            before_type = geom.geom_type if geom is not None else None
            # diagnostics
            if geom is None:
                notes.append("Empty geometry")
                has_holes_col.append(None)
                part_count_col.append(None)
                area_col.append(None)
            else:
                has_holes_col.append(
                    (len(getattr(geom, "interiors", [])) > 0) if isinstance(geom, Polygon)
                    else (any(len(p.interiors) > 0 for p in geom.geoms) if isinstance(geom, MultiPolygon) else False)
                )
                part_count_col.append(len(geom.geoms) if hasattr(geom, "geoms") else 1)
                area_col.append(getattr(geom, "area", None))

                fixed, n = fix_geometry(
                    geom, drop_holes=drop_holes, force_multipolygon=force_multipolygon,
                    largest_only=largest_only, orient_ccw=orient_ccw
                )
                notes.extend(n if n else [])
                if fixed is not None:
                    is_valid = fixed.is_valid and isinstance(fixed, (Polygon, MultiPolygon))
                    after_type = fixed.geom_type
                    if uppercase_type:
                        if before_type: before_type = before_type.upper()
                        if after_type:  after_type  = after_type.upper()
                    fixed_wkt = fixed.wkt
                    if not is_valid:
                        notes.append("Geometry still invalid after fix")
                else:
                    is_valid = False
        except Exception as e:
            notes.append(f"Parse error: {e}")
            has_holes_col.append(None)
            part_count_col.append(None)
            area_col.append(None)

        rows.append((before_type, after_type, is_valid, "; ".join(notes) if notes else "", fixed_wkt))

    df_out = df.copy()
    df_out["geom_type_before"] = [r[0] for r in rows]
    df_out["geom_type_after"]  = [r[1] for r in rows]
    df_out["is_valid"]         = [r[2] for r in rows]
    df_out["notes"]            = [r[3] for r in rows]
    df_out["wkt_geom_fixed"]   = [r[4] for r in rows]
    df_out["has_holes"]        = has_holes_col
    df_out["part_count"]       = part_count_col
    df_out["area"]             = area_col

    base = Path(in_path).stem + "_fixed"
    xlsx_path = out_dir / f"{base}.xlsx"
    csv_path  = out_dir / f"{base}_report.csv"
    geojson_path = out_dir / f"{base}.geojson"

    df_out.to_excel(xlsx_path, index=False)
    df_out.loc[~df_out["is_valid"]].to_csv(csv_path, index=False)

    prop_cols = [c for c in df_out.columns if c not in (wkt_col,"wkt_geom_fixed","geom_type_before","geom_type_after","is_valid","notes","has_holes","part_count","area")]
    features = []
    for _, r in df_out.iterrows():
        w = r.get("wkt_geom_fixed") or r.get(wkt_col)
        if pd.isna(w): 
            continue
        try:
            g = shapely_wkt.loads(str(w))
            features.append({
                "type": "Feature",
                "properties": {k: (None if pd.isna(r[k]) else r[k]) for k in prop_cols},
                "geometry": shp_mapping(g)
            })
        except Exception:
            continue

    with open(geojson_path, "w", encoding="utf-8") as f:
        json.dump({"type":"FeatureCollection","features":features}, f, ensure_ascii=False)

    return xlsx_path, csv_path, geojson_path

def parse_or_prompt():
    parser = argparse.ArgumentParser(
        description="Corriger les géométries WKT (Excel) : conversion lignes→polygones, réparation, options EUDR (remove-holes, largest-part-only, orient-ccw), export Excel/CSV/GeoJSON."
    )
    parser.add_argument("--excel", help="Chemin du fichier Excel d'entrée (ex: data.xlsx)")
    parser.add_argument("--sheet", default=0, help="Nom ou index de feuille (défaut: 0)")
    parser.add_argument("--wkt-col", default="wkt_geom", help="Nom de la colonne WKT (défaut: wkt_geom)")
    parser.add_argument("--outdir", default="outputs", help="Dossier de sortie (défaut: outputs)")
    parser.add_argument("--remove-holes", action="store_true", help="Supprime les trous (anneaux intérieurs)")
    parser.add_argument("--force-multipolygon", action="store_true", help="Convertit tout Polygon en MultiPolygon")
    parser.add_argument("--uppercase-type", action="store_true", help="MAJUSCULES pour geom_type_* (EXCEL/CSV)")
    parser.add_argument("--largest-part-only", action="store_true", help="Si MultiPolygon, garde uniquement la plus grande partie (Polygon)")
    parser.add_argument("--orient-ccw", action="store_true", help="Réoriente les anneaux extérieurs en CCW")
    args, _ = parser.parse_known_args()

    if args.excel is None:
        print("Aucun argument fourni. Mode interactif :")
        excel = input("Chemin du fichier Excel (.xlsx) : ").strip().strip('"').strip("'")
        outdir = input("Dossier de sortie (par ex. outputs) : ").strip().strip('"').strip("'") or "outputs"
        sheet = input("Feuille (nom ou index, défaut 0) : ").strip() or "0"
        wkt_col = input("Nom de la colonne WKT (défaut wkt_geom) : ").strip() or "wkt_geom"
        remove_holes = input("Supprimer les trous ? (o/N) : ").strip().lower() in ("o", "oui", "y", "yes")
        force_multi = input("Forcer MultiPolygon ? (o/N) : ").strip().lower() in ("o", "oui", "y", "yes")
        upper = input("Mettre les types en MAJUSCULES ? (o/N) : ").strip().lower() in ("o", "oui", "y", "yes")
        largest = input("Garder seulement la plus grande partie si MultiPolygon ? (o/N) : ").strip().lower() in ("o", "oui", "y", "yes")
        orient = input("Réorienter exterieurs CCW ? (o/N) : ").strip().lower() in ("o", "oui", "y", "yes")
        return excel, sheet if sheet.isdigit() else sheet, wkt_col, outdir, remove_holes, force_multi, upper, largest, orient

    return args.excel, args.sheet, args.wkt_col, args.outdir, args.remove_holes, args.force_multipolygon, args.uppercase_type, args.largest_part_only, args.orient_ccw

def main():
    excel, sheet, wkt_col, outdir, remove_holes, force_multi, upper, largest, orient = parse_or_prompt()
    xlsx_path, csv_path, geojson_path = process_excel(
        excel, outdir, wkt_col=wkt_col, sheet_name=sheet,
        drop_holes=remove_holes, force_multipolygon=force_multi,
        uppercase_type=upper, largest_only=largest, orient_ccw=orient
    )
    print("Terminé ✅")
    print(f"Excel corrigé : {xlsx_path}")
    print(f"Rapport CSV    : {csv_path}")
    print(f"GeoJSON        : {geojson_path}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Erreur :", e)
        sys.exit(1)
