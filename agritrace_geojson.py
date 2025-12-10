#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AgriTrace GeoJSON Generator – INTERACTIF (dedupe + area compute)
----------------------------------------------------------------
- Pose des questions si les options ne sont pas passées en arguments
- Génère un GeoJSON avec les champs: Area, ProductionPlace, ProducerName, ProducerCountry, geometry
- Déduplication stricte par géométrie (coords arrondies -> empreinte WKB)
- Option pour calculer Area (ha) à partir des géométries (géodésique WGS84)

Usage simple (tout interactif):
    python agritrace_geojson.py

Usage avec quelques options (le reste sera demandé):
    python agritrace_geojson.py --excel "chemin/fichier.xlsx"
"""
import os
import sys
import json
import argparse
import unicodedata
from typing import Optional, Tuple, List

import pandas as pd
from shapely import wkt
from shapely.geometry import mapping, shape
from shapely.geometry.base import BaseGeometry
from pyproj import Geod, CRS, Transformer

# ----------------- Defaults & fallbacks -----------------
DEF_COUNTRY = "CD"
DEF_WKT_COL = "geometry"
DEF_PRODUCER_NAME_COL = "CODE_PLANTATION"
DEF_INPUT_CRS = "EPSG:4326"  # WGS84 lon/lat

FALLBACK_AREA = ["HECTARE", "Area", "Superficie", "Surface", "Hectares", "HA", "ha"]
FALLBACK_PLACE = [
    "CODE_PLANTEUR", "Code_Planteur", "code_planteur",
    "CODE PLANTEUR", "Code Planteur", "code planteur",
    "CODE-PLANTEUR", "Code-Planteur", "code-planteur",
    "Village", "ProductionPlace", "Localite", "Localité", "Site", "Place"
]
FALLBACK_WKT = ["geometry", "wkt_geom", "wkt", "GEOMETRY", "WKT_GEOM", "WKT"]


# ----------------- Utils -----------------
def normalize_path(path_str: str) -> str:
    try:
        return unicodedata.normalize("NFC", path_str)
    except Exception:
        return path_str


def ensure_outfile_path(out_path: str) -> str:
    out_path = normalize_path(out_path)
    if os.path.isdir(out_path):
        out_path = os.path.join(out_path, "agritrace_export.geojson")
    root, ext = os.path.splitext(out_path)
    if ext.lower() not in (".geojson", ".json"):
        out_path = out_path + ".geojson"
    parent = os.path.dirname(out_path) or "."
    try:
        if parent and not os.path.exists(parent):
            os.makedirs(parent, exist_ok=True)
    except Exception:
        print(f"⚠️  Impossible de créer le dossier '{parent}'. Export dans le dossier courant.")
        out_path = os.path.join(os.getcwd(), "agritrace_export.geojson")
    return out_path


def write_geojson_safely(feature_collection: dict, out_path: str) -> Tuple[bool, str]:
    out_path = ensure_outfile_path(out_path)
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(feature_collection, f, ensure_ascii=False)
        print(f"\n✅ GeoJSON écrit: {out_path}")
        return True, out_path
    except PermissionError:
        print(f"⚠️  Permission refusée pour écrire: {out_path}")
        fallback = os.path.join(os.getcwd(), "agritrace_export.geojson")
        try:
            with open(fallback, "w", encoding="utf-8") as f:
                json.dump(feature_collection, f, ensure_ascii=False)
            print(f"↪️  Sauvegarde de secours: {fallback}")
            return True, fallback
        except Exception as e2:
            print(f"❌ Échec d'écriture même dans le dossier courant: {e2}")
            return False, out_path
    except Exception as e:
        print(f"❌ ERREUR: Impossible d'écrire le GeoJSON ({e})")
        return False, out_path


def pick_first_present(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    low_map = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in low_map:
            return low_map[c.lower()]
    return None


def ask(prompt: str, default: Optional[str] = None) -> str:
    suf = f" [{default}]" if default is not None else ""
    ans = input(f"{prompt}{suf}: ").strip()
    return ans if ans else (default or "")


def yesno(prompt: str, default: bool = True) -> bool:
    suf = " [O/n]" if default else " [o/N]"
    ans = input(f"{prompt}{suf}: ").strip().lower()
    if ans == "":
        return default
    return ans in ("o", "oui", "y", "yes")


def geom_fingerprint_geojson(geom_dict: dict, precision: int = 6) -> str:
    """Empreinte stable d'une géométrie en arrondissant les coords puis WKB hex."""
    g: BaseGeometry = shape(geom_dict)
    m = mapping(g)

    def round_coords(c):
        if isinstance(c, (list, tuple)):
            return [round_coords(x) for x in c]
        if isinstance(c, float):
            return round(c, precision)
        return c

    m["coordinates"] = round_coords(m["coordinates"])
    return shape(m).wkb_hex


# ----------------- Area (ha) -----------------
def geodesic_area_ha(geom: BaseGeometry) -> float:
    """Aire géodésique WGS84 (m² -> ha) pour Polygon/MultiPolygon (coords en lon/lat)."""
    geod = Geod(ellps="WGS84")
    area_m2 = 0.0

    def ring_area(coords):
        lons, lats = zip(*coords)
        a, _ = geod.polygon_area_perimeter(lons, lats)
        return abs(a)

    if geom.geom_type == "Polygon":
        area_m2 += ring_area(list(geom.exterior.coords))
        for ring in geom.interiors:
            area_m2 -= ring_area(list(ring.coords))
    elif geom.geom_type == "MultiPolygon":
        for p in geom.geoms:
            area_m2 += ring_area(list(p.exterior.coords))
            for ring in p.interiors:
                area_m2 -= ring_area(list(ring.coords))
    else:
        return 0.0
    return area_m2 / 10_000.0


def to_wgs84(geom: BaseGeometry, input_crs: str) -> BaseGeometry:
    """Reprojette depuis input_crs vers WGS84 si nécessaire (EPSG:4326)."""
    if input_crs.upper() in ("EPSG:4326", "WGS84"):
        return geom
    transformer = Transformer.from_crs(CRS.from_string(input_crs), CRS.from_epsg(4326), always_xy=True)

    def _reproj_coords(coords):
        return [transformer.transform(x, y) for (x, y) in coords]

    from shapely.geometry import Polygon, MultiPolygon
    if geom.geom_type == "Polygon":
        ext = _reproj_coords(list(geom.exterior.coords))
        ints = [_reproj_coords(list(r.coords)) for r in geom.interiors]
        return Polygon(ext, ints)
    if geom.geom_type == "MultiPolygon":
        polys = []
        for p in geom.geoms:
            ext = _reproj_coords(list(p.exterior.coords))
            ints = [_reproj_coords(list(r.coords)) for r in p.interiors]
            polys.append(Polygon(ext, ints))
        return MultiPolygon(polys)
    return geom


# ----------------- Core build/dedupe -----------------
def build_geojson(
    df: pd.DataFrame,
    area_col: Optional[str],
    place_col: Optional[str],
    producer_name_col: str,
    wkt_col: str,
    producer_country: str,
    compute_area: bool,
    input_crs: str,
    errors_list: List[tuple],
) -> dict:
    features = []
    for idx, row in df.iterrows():
        w = row.get(wkt_col, None)
        if not isinstance(w, str) or not w.strip():
            errors_list.append((idx, "geometry vide ou non string"))
            continue
        try:
            geom = wkt.loads(w)  # CRS: input_crs
        except Exception as e:
            errors_list.append((idx, f"WKT invalide: {e}"))
            continue

        def _safe_float(x):
            if x is None:
                return None
            try:
                return float(str(x).replace(",", "."))
            except Exception:
                return None

        area_val = _safe_float(row.get(area_col, None) if area_col else None)

        props = {
            "Area": area_val,
            "ProductionPlace": row.get(place_col, None) if place_col else None,
            "ProducerName": row.get(producer_name_col, None) if producer_name_col else None,
            "ProducerCountry": producer_country or DEF_COUNTRY,
        }

        if compute_area and (props["Area"] is None or props["Area"] <= 0):
            try:
                g_wgs84 = to_wgs84(geom, input_crs)
                props["Area"] = round(geodesic_area_ha(g_wgs84), 6)
            except Exception:
                pass

        features.append({"type": "Feature", "properties": props, "geometry": mapping(geom)})

    return {"type": "FeatureCollection", "features": features}


def dedupe_by_geometry(feature_collection: dict, precision: int = 6) -> Tuple[dict, List[dict]]:
    seen = set()
    kept, removed = [], []
    for i, ft in enumerate(feature_collection.get("features", []), start=1):
        geom = ft.get("geometry")
        if not geom:
            kept.append(ft)
            continue
        try:
            fp = geom_fingerprint_geojson(geom, precision=precision)
        except Exception:
            kept.append(ft)
            continue
        if fp in seen:
            props = ft.get("properties", {}) or {}
            removed.append({
                "feature_index": i,
                "ProducerName": props.get("ProducerName"),
                "ProductionPlace": props.get("ProductionPlace"),
                "Area": props.get("Area"),
                "reason": "Duplicate geometry"
            })
            continue
        seen.add(fp)
        kept.append(ft)
    return {"type": "FeatureCollection", "features": kept}, removed


# ----------------- CLI + INTERACTIF -----------------
def main():
    p = argparse.ArgumentParser(description="AgriTrace: générateur GeoJSON (interactif)")
    p.add_argument("--excel", help="Chemin du fichier Excel (.xlsx)")
    p.add_argument("--sheet", help="Nom de la feuille (facultatif)")
    p.add_argument("--wkt-col", default=DEF_WKT_COL, help=f"Colonne WKT d'entrée (défaut: {DEF_WKT_COL})")
    p.add_argument("--producer-name-col", default=DEF_PRODUCER_NAME_COL, help=f"Colonne pour ProducerName (défaut: {DEF_PRODUCER_NAME_COL})")
    p.add_argument("--area-col", help="Colonne pour Area (priorité à 'HECTARE' si présent)")
    p.add_argument("--place-col", help="Colonne pour ProductionPlace (priorité à 'CODE_PLANTEUR' si présent)")
    p.add_argument("--country", default=DEF_COUNTRY, help=f"ProducerCountry (défaut: {DEF_COUNTRY})")
    p.add_argument("--out", default="agritrace_export.geojson", help="Chemin du GeoJSON de sortie")
    p.add_argument("--errors-csv", help="Chemin CSV pour erreurs WKT et doublons")
    p.add_argument("--precision", type=int, default=6, help="Décimales d'arrondi pour la déduplication")
    p.add_argument("--compute-area", action="store_true", help="Calculer Area (ha) si manquante/invalide")
    p.add_argument("--input-crs", default=DEF_INPUT_CRS, help=f"CRS d'entrée des WKT (défaut: {DEF_INPUT_CRS})")
    args = p.parse_args()

    # --- Mode INTERACTIF (si paramètres manquants) ---
    if not args.excel:
        print("\n=== AgriTrace GeoJSON – Mode interactif ===")
        args.excel = ask("Chemin du fichier Excel (.xlsx)")
        if not args.excel or not os.path.exists(args.excel):
            print("ERREUR: Fichier Excel introuvable.")
            sys.exit(1)

    if not args.sheet:
        args.sheet = ask("Nom de la feuille (laisser vide pour la 1ère)", "")

    # Lecture Excel
    try:
        df = pd.read_excel(args.excel, sheet_name=args.sheet) if args.sheet else pd.read_excel(args.excel)
    except Exception as e:
        print(f"ERREUR: Impossible de lire l'Excel ({e})")
        sys.exit(1)

    print(f"Colonnes détectées: {list(df.columns)}")

    # WKT
    if args.wkt_col not in df.columns:
        fb = pick_first_present(df, FALLBACK_WKT)
        if not fb:
            args.wkt_col = ask("Nom de la colonne WKT (ex: geometry)")
            if args.wkt_col not in df.columns:
                print("ERREUR: colonne WKT introuvable.")
                sys.exit(1)
        else:
            if yesno(f"Remplacer '{fb}' par 'geometry' automatiquement ?", True):
                df = df.rename(columns={fb: "geometry"})
                args.wkt_col = "geometry"
            else:
                args.wkt_col = ask("Nom de la colonne WKT à utiliser", fb)

    # ProducerName
    if args.producer_name_col not in df.columns:
        cand = pick_first_present(df, [args.producer_name_col, "CODE_PLANTATION", "Code_Plantation", "code_plantation", "CODE", "ID"])
        if not cand:
            args.producer_name_col = ask("Colonne pour ProducerName (ex: CODE_PLANTATION)")
            if args.producer_name_col not in df.columns:
                print("ERREUR: colonne ProducerName introuvable.")
                sys.exit(1)
        else:
            args.producer_name_col = cand

    # Area & ProductionPlace
    if not args.area_col:
        auto_area = pick_first_present(df, FALLBACK_AREA)
        args.area_col = ask("Colonne Area (laisser vide pour auto)", auto_area or "")
        if args.area_col and args.area_col not in df.columns:
            print(f"⚠️  '{args.area_col}' introuvable, Area sera calculée si activé.")
            args.area_col = None

    if not args.place_col:
        auto_place = pick_first_present(df, FALLBACK_PLACE)
        args.place_col = ask("Colonne ProductionPlace (laisser vide pour auto)", auto_place or "")
        if args.place_col and args.place_col not in df.columns:
            print(f"⚠️  '{args.place_col}' introuvable, ProductionPlace sera vide.")
            args.place_col = None

    # Country / CRS / compute-area / precision / out / errors
    args.country = ask("ProducerCountry (code ISO-2)", args.country or DEF_COUNTRY).upper() or DEF_COUNTRY
    args.input_crs = ask("CRS d'entrée (ex: EPSG:4326, EPSG:32735)", args.input_crs or DEF_INPUT_CRS)
    if not args.compute_area:
        args.compute_area = yesno("Calculer Area (hectares) si manquante/invalide ?", True)
    try:
        args.precision = int(ask("Précision d'arrondi des coords pour déduplication (2-8)", str(args.precision)))
        args.precision = max(2, min(8, args.precision))
    except Exception:
        pass
    args.out = ask("Chemin du GeoJSON de sortie", args.out or "agritrace_export.geojson")
    if not args.errors_csv:
        ec = ask("Chemin CSV pour erreurs & doublons (optionnel)", "")
        args.errors_csv = ec if ec else None

    # Génération
    errors = []
    fc = build_geojson(
        df=df,
        area_col=args.area_col,
        place_col=args.place_col,
        producer_name_col=args.producer_name_col,
        wkt_col=args.wkt_col,
        producer_country=args.country,
        compute_area=args.compute_area,
        input_crs=args.input_crs,
        errors_list=errors,
    )

    # Déduplication
    fc, dupes = dedupe_by_geometry(fc, precision=args.precision)
    print(f"  - Entités exportées (après déduplication): {len(fc['features'])}")
    print(f"  - Doublons de géométrie supprimés: {len(dupes)}")

    # Écriture GeoJSON
    ok, final_out = write_geojson_safely(fc, args.out)
    if not ok:
        sys.exit(1)

    # CSV erreurs/dupes
    if args.errors_csv:
        path = normalize_path(args.errors_csv)
        parent = os.path.dirname(path) or "."
        try:
            if parent and not os.path.exists(parent):
                os.makedirs(parent, exist_ok=True)
        except Exception:
            print(f"  (Impossible de créer le dossier '{parent}' pour les CSV)")
        try:
            if errors:
                pd.DataFrame(errors, columns=["RowIndex", "Issue"]).to_csv(path, index=False, encoding="utf-8")
                print(f"  → Détails WKT invalides: {path}")
            if dupes:
                dup_path = path.replace(".csv", "_dupes.csv") if path.endswith(".csv") else (path + "_dupes.csv")
                pd.DataFrame(dupes).to_csv(dup_path, index=False, encoding="utf-8")
                print(f"  → Détails des doublons: {dup_path}")
        except Exception as e:
            print(f"  (Impossible d'écrire les CSV: {e})")
    else:
        if errors:
            print(f"\n⚠️ {len(errors)} ligne(s) ignorée(s) (WKT vide/invalide). Donne un chemin --errors-csv pour exporter les détails.")

    print("\n✅ Terminé.")


if __name__ == "__main__":
    main()
