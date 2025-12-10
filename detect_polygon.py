#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
detect_polygon.py  —  Mode interactif

Détecte les polygones potentiellement "fabriqués" (vs collectés sur terrain)
à partir d'un GeoJSON ou KML, en posant une série de questions dans le terminal.

Sorties :
  - CSV : métriques + risk_score + risk_level + risk_reasons
  - GeoJSON annoté (mêmes colonnes)

Dépendances (installer une seule fois) :
    pip install geopandas shapely fiona pyproj pandas numpy
"""

import os, sys, math, hashlib
from typing import Tuple, Optional, List

import numpy as np
import pandas as pd
import geopandas as gpd
from shapely.geometry import Polygon, MultiPolygon, Point
from shapely.ops import unary_union


# ==========================
# ===== Prompts I/O =========
# ==========================

def prompt_path(msg: str, must_exist: bool = True, default: Optional[str] = None) -> Optional[str]:
    while True:
        s = input(f"{msg}{f' [{default}]' if default else ''} : ").strip().strip('"')
        if not s and default:
            s = default
        if s == "":
            return None
        # Nettoyage d’éventuels caractères invisibles copiés-collés (ex: LRM)
        s = s.replace("\u200e", "").replace("\u202a", "").replace("\u202c", "")
        if must_exist and not os.path.exists(s):
            print("  -> Chemin introuvable. Réessaie.")
            continue
        return s

def prompt_yn(msg: str, default: Optional[bool] = None) -> bool:
    suffix = " (o/n)"
    if default is True: suffix = " (O/n)"
    if default is False: suffix = " (o/N)"
    while True:
        s = input(f"{msg}{suffix} : ").strip().lower()
        if s == "" and default is not None:
            return default
        if s in ("o", "oui", "y", "yes"): return True
        if s in ("n", "non", "no"): return False
        print("  -> Répondre o/n.")

def prompt_strictness() -> str:
    print("Niveau de sévérité pour le scoring :")
    print("  1) Strict  (plus de HIGH)")
    print("  2) Moyen   (équilibré)")
    print("  3) Souple  (moins de HIGH)")
    while True:
        s = input("Choix [1/2/3] : ").strip()
        if s in ("1","2","3"):
            return {"1":"strict","2":"medium","3":"lenient"}[s]
        print("  -> Choisir 1, 2 ou 3.")

def prompt_float(msg: str, default: Optional[float] = None) -> Optional[float]:
    while True:
        s = input(f"{msg}{f' [{default}]' if default is not None else ''} : ").strip()
        if s == "":
            return default
        try:
            return float(s)
        except ValueError:
            print("  -> Saisir un nombre (ex: 0.05).")

def parse_office(s: Optional[str]) -> Optional[Tuple[float,float]]:
    if not s: return None
    try:
        lon, lat = s.split(",")
        return (float(lon), float(lat))
    except Exception:
        print('  -> Format attendu "lon,lat" (ex: 29.70,0.30). Ignoré.')
        return None


# ==================================================
# ===== Métriques géométriques & utilitaires =======
# ==================================================

def choose_utm_epsg(lat: float, lon: float) -> int:
    zone = int((lon + 180) // 6) + 1
    return (32600 if lat >= 0 else 32700) + zone

def to_metric_crs(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Reprojette en UTM adapté au centroïde global pour calculs métriques."""
    if gdf.crs is None:
        gdf = gdf.set_crs(epsg=4326, allow_override=True)
    # Evite le warning centroid en WGS84 en centrant sur l’union
    try:
        c = gdf.geometry.unary_union.centroid
    except Exception:
        c = gdf.geometry.unary_union().centroid  # compat
    epsg = choose_utm_epsg(c.y, c.x)
    return gdf.to_crs(epsg=epsg)

def feature_hash(geom, ndigits: int = 5) -> str:
    """Hash approx. (coords arrondies) pour détecter doublons/quasi-doublons."""
    if geom is None or geom.is_empty:
        return "EMPTY"
    coords = []
    if isinstance(geom, Polygon):
        coords.append([(round(x, ndigits), round(y, ndigits)) for x,y in geom.exterior.coords])
        for ring in geom.interiors:
            coords.append([(round(x, ndigits), round(y, ndigits)) for x,y in ring.coords])
    elif isinstance(geom, MultiPolygon):
        for poly in geom.geoms:
            coords.append([(round(x, ndigits), round(y, ndigits)) for x,y in poly.exterior.coords])
    else:
        try:
            coords.append([(round(x, ndigits), round(y, ndigits)) for x,y in geom.coords])
        except Exception:
            coords.append([(round(geom.centroid.x, ndigits), round(geom.centroid.y, ndigits))])
    return hashlib.sha256(str(coords).encode("utf-8")).hexdigest()

def angle_degrees(p1, p2) -> float:
    dx, dy = p2[0] - p1[0], p2[1] - p1[1]
    return math.degrees(math.atan2(dy, dx)) % 180.0

def edge_angles(poly: Polygon) -> List[float]:
    coords = list(poly.exterior.coords)
    return [angle_degrees(coords[i], coords[i+1]) for i in range(len(coords)-1)]

def frac_near_right_angles(angles: List[float], tol: float = 4.0) -> float:
    if not angles: return 0.0
    hits = 0
    for a in angles:
        if min(abs((a % 90) - 0), abs((a % 90) - 90)) <= tol or abs(a) <= tol or abs(a-90) <= tol:
            hits += 1
    return hits / len(angles)

def rectangularity(poly: Polygon) -> float:
    """Aire / aire du rectangle minimum tourné (≈1 si rectangle parfait)."""
    try:
        mrr = poly.minimum_rotated_rectangle
        if not isinstance(mrr, Polygon): return 0.0
        a = poly.area; b = mrr.area if mrr.area > 0 else 1e-9
        return a / b
    except Exception:
        return 0.0

def polsby_popper(a: float, p: float) -> float:
    if p <= 0: return 0.0
    return (4 * math.pi * a) / (p * p)

def coord_decimal_stats(poly: Polygon):
    """Stats nb de décimales des coordonnées (arrondis → suspects)."""
    def count_decimals(v: float) -> int:
        s = f"{v:.10f}".rstrip("0")
        return len(s.split(".")[1]) if "." in s else 0
    vals = []
    for x,y in poly.exterior.coords:
        vals += [count_decimals(x), count_decimals(y)]
    if not vals: return (0.0, 0.0, 0.0)
    arr = np.array(vals)
    return float(arr.mean()), float(arr.std()), float(arr.min())

def edge_length_cv(poly: Polygon) -> float:
    coords = list(poly.exterior.coords)
    lengths = []
    for i in range(len(coords)-1):
        dx = coords[i+1][0] - coords[i][0]
        dy = coords[i+1][1] - coords[i][1]
        lengths.append((dx*dx + dy*dy) ** 0.5)
    if not lengths: return 0.0
    arr = np.array(lengths)
    return float(0.0 if arr.mean() == 0 else arr.std() / arr.mean())

def compute_metrics(gdf_wgs84: gpd.GeoDataFrame) -> pd.DataFrame:
    """Calcule toutes les métriques sur une reprojection métrique robuste."""
    gdf_m = to_metric_crs(gdf_wgs84)
    recs = []
    for idx, (geom_w, geom_m) in enumerate(zip(gdf_wgs84.geometry, gdf_m.geometry)):
        if geom_w is None or geom_w.is_empty:
            recs.append(dict(idx=idx, valid=False, reason_empty=True)); continue
        if isinstance(geom_m, Polygon):
            poly = geom_m
        elif isinstance(geom_m, MultiPolygon):
            try:
                poly = unary_union(geom_m)
                if isinstance(poly, MultiPolygon):
                    poly = max(geom_m.geoms, key=lambda p: p.area)
            except Exception:
                poly = max(geom_m.geoms, key=lambda p: p.area)
        else:
            recs.append(dict(idx=idx, valid=False, non_polygon=True)); continue

        n_vertices = max(len(list(poly.exterior.coords)) - 1, 0)
        n_holes = len(poly.interiors)
        area_m2 = float(poly.area)
        perim_m = float(poly.length)
        pp = polsby_popper(area_m2, perim_m)
        rect = rectangularity(poly)
        angs = edge_angles(poly)
        frac90 = frac_near_right_angles(angs, tol=4.0)
        cvlen = edge_length_cv(poly)
        mean_dec, std_dec, min_dec = coord_decimal_stats(poly)

        try:
            mrr = poly.minimum_rotated_rectangle
            bbox_ratio = area_m2 / max(mrr.area, 1e-9)
        except Exception:
            bbox_ratio = rect

        recs.append(dict(
            idx=idx, valid=True, n_vertices=n_vertices, n_holes=n_holes,
            area_m2=area_m2, perimeter_m=perim_m, compactness_pp=pp,
            rectangularity=bbox_ratio, frac_right_angles=frac90, edge_len_cv=cvlen,
            mean_decimals=mean_dec, std_decimals=std_dec, min_decimals=min_dec,
        ))
    return pd.DataFrame.from_records(recs)


# ============================================
# ===== Contexte (AOI, bureau, bornes) =======
# ============================================

def add_context_flags(df: pd.DataFrame,
                      gdf_wgs84: gpd.GeoDataFrame,
                      office: Optional[Tuple[float,float]] = None,
                      aoi_path: Optional[str] = None,
                      area_min_ha: Optional[float] = None,
                      area_max_ha: Optional[float] = None) -> pd.DataFrame:
    df = df.copy()

    # Doublons géométriques
    hashes = [feature_hash(g) for g in gdf_wgs84.geometry]
    df["geom_hash"] = hashes
    df["dup_count"] = pd.Series(hashes).map(pd.Series(hashes).value_counts()).values

    # Centroïdes : calcule en métrique puis reprojette en WGS84 (évite le warning)
    gdf_m   = to_metric_crs(gdf_wgs84)
    cents_m = gdf_m.geometry.centroid
    cents   = gpd.GeoSeries(cents_m, crs=gdf_m.crs).to_crs(4326)
    df["centroid_lon"] = cents.x
    df["centroid_lat"] = cents.y

    # Distance au bureau (m)
    if office is not None:
        office_pt = gpd.GeoSeries([Point(office)], crs=4326).to_crs(gdf_m.crs).iloc[0]
        df["dist_office_m"] = cents_m.distance(office_pt)

    # AOI in/out (sur centroïde)
    if aoi_path:
        try:
            aoi = gpd.read_file(aoi_path)
            if aoi.crs is None:
                aoi = aoi.set_crs(epsg=4326, allow_override=True)
            aoi = aoi.to_crs(4326)
            try:
                union_aoi = aoi.union_all()  # GeoPandas récent
            except Exception:
                union_aoi = aoi.unary_union    # fallback
            df["in_aoi"] = gpd.GeoSeries(cents, crs=4326).within(union_aoi).values
        except Exception as e:
            print(f"AVERTISSEMENT: AOI non lue: {e}")

    # Bornes d’aire (ha)
    if area_min_ha is not None:
        df["flag_area_min"] = (df.get("area_m2", 0) < area_min_ha * 10000)
    if area_max_ha is not None:
        df["flag_area_max"] = (df.get("area_m2", 0) > area_max_ha * 10000)

    return df


# ============================================
# ===== Scoring & classification =============
# ============================================

def risk_score_row(row: pd.Series, strictness: str = "medium", office: bool = False):
    if strictness == "strict":
        tol_right, tol_rect, tol_vertices, tol_cv, tol_dec, near_office = 0.80, 0.98, 5, 0.12, 3.5, 300.0
    elif strictness == "lenient":
        tol_right, tol_rect, tol_vertices, tol_cv, tol_dec, near_office = 0.92, 0.995, 4, 0.08, 3.0, 200.0
    else:
        tol_right, tol_rect, tol_vertices, tol_cv, tol_dec, near_office = 0.85, 0.99, 5, 0.10, 3.2, 250.0

    score, reasons = 0, []
    if not row.get("valid", False): score += 2; reasons.append("geom_invalide_ou_non_polygone")
    if row.get("dup_count", 1) > 1: score += 4; reasons.append("dupli_geom_hash")
    if row.get("rectangularity", 0) >= tol_rect and row.get("n_vertices", 0) <= tol_vertices:
        score += 3; reasons.append("rectangle_parfait_peu_de_vertices")
    if row.get("frac_right_angles", 0) >= tol_right:
        score += 2; reasons.append("trop_d_angles_90deg")
    if row.get("edge_len_cv", 1) <= tol_cv and row.get("n_vertices", 0) >= 4:
        score += 1; reasons.append("aretes_trop_regulieres")
    if row.get("min_decimals", 10) <= tol_dec:
        score += 1; reasons.append("coords_arrondies")
    if row.get("flag_area_min", False):
        score += 2; reasons.append("aire_trop_petite")
    if row.get("flag_area_max", False):
        score += 1; reasons.append("aire_trop_grande")
    if office and row.get("dist_office_m", 1e9) <= near_office:
        score += 3; reasons.append("proche_du_bureau")
    if "in_aoi" in row and row["in_aoi"] is False:
        score += 2; reasons.append("hors_AOI")
    return score, reasons

def classify_level(score: int) -> str:
    return "HIGH" if score >= 6 else ("MEDIUM" if score >= 3 else "LOW")


# ===========================
# ===== Lecture entrée ======
# ===========================

def read_any(path: str) -> gpd.GeoDataFrame:
    try:
        gdf = gpd.read_file(path)
    except Exception as e:
        if path.lower().endswith((".kml", ".kmz")):
            try:
                gdf = gpd.read_file(path, driver="KML")
            except Exception as e2:
                raise SystemExit(f"Impossible de lire KML. Convertissez en GeoJSON (QGIS/ogr2ogr). Détails: {e2}")
        else:
            raise
    if gdf.crs is None:
        gdf = gdf.set_crs(epsg=4326, allow_override=True)
    else:
        gdf = gdf.to_crs(4326)
    return gdf


# ==============================================
# ===== Aide chemins fichiers de sortie =========
# ==============================================

def ensure_file_path(path: Optional[str], default_filename: str) -> str:
    """
    - Si path est un dossier OU sans extension : crée le dossier et renvoie path/<default_filename>.
    - Sinon : crée le dossier parent et renvoie path tel quel.
    """
    if not path or path.strip() == "":
        path = default_filename
    root, ext = os.path.splitext(path)
    # Dossier donné (ou pas d'extension) -> compose un vrai chemin fichier
    if (ext == "") or os.path.isdir(path):
        # S'il n'y a pas d'extension, 'path' est traité comme dossier
        out_dir = path if ext == "" else path
        os.makedirs(out_dir, exist_ok=True)
        return os.path.join(out_dir if ext == "" else out_dir, default_filename)
    # Chemin fichier avec extension
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    return path


# ===========================
# ===== Programme main ======
# ===========================

def main():
    print("=== Détecteur de polygones fabriqués — Mode interactif ===")

    in_path = prompt_path('Chemin du fichier d\'entrée (GeoJSON ou KML)', must_exist=True)
    if not in_path:
        print("Aucun fichier fourni. Fin."); sys.exit(1)

    office_txt = input('Coordonnées du bureau "lon,lat" (laisser vide si aucun) : ').strip()
    office = parse_office(office_txt) if office_txt else None

    aoi_path = prompt_path("Chemin de l'AOI (GeoJSON/KML) (laisser vide si aucun)", must_exist=True, default=None)

    area_min_ha = prompt_float("Aire minimale attendue en hectares (laisser vide si aucun)", default=None)
    area_max_ha = prompt_float("Aire maximale attendue en hectares (laisser vide si aucun)", default=None)

    strictness = prompt_strictness()

    base_no_ext = os.path.splitext(in_path)[0]
    out_csv_in  = prompt_path("Chemin CSV de sortie", must_exist=False, default=base_no_ext + "_qc.csv")
    out_gj_in   = prompt_path("Chemin GeoJSON annoté de sortie", must_exist=False, default=base_no_ext + "_scored.geojson")

    # S'assure d'avoir de vrais chemins "fichiers"
    base_name = os.path.basename(base_no_ext)
    out_csv = ensure_file_path(out_csv_in,     f"{base_name}_qc.csv")
    out_gj  = ensure_file_path(out_gj_in,      f"{base_name}_scored.geojson")

    add_wkt = prompt_yn("Ajouter une colonne WKT dans le CSV ?", default=False)

    # ---- Traitement
    print("\nLecture des données…")
    gdf = read_any(in_path)

    print("Calcul des métriques…")
    metrics = compute_metrics(gdf)
    metrics = add_context_flags(metrics, gdf, office=office, aoi_path=aoi_path,
                                area_min_ha=area_min_ha, area_max_ha=area_max_ha)

    print("Scoring…")
    scores, reasons_list, levels = [], [], []
    for _, row in metrics.iterrows():
        s, rs = risk_score_row(row, strictness=strictness, office=(office is not None))
        scores.append(s)
        reasons_list.append(",".join(rs) if rs else "")
        levels.append(classify_level(s))
    metrics["risk_score"]   = scores
    metrics["risk_reasons"] = reasons_list
    metrics["risk_level"]   = levels

    if add_wkt:
        try:
            metrics["geometry_wkt"] = gdf.geometry.apply(lambda g: g.wkt if g is not None else None)
        except Exception:
            pass

    print(f"Écriture CSV → {out_csv}")
    os.makedirs(os.path.dirname(out_csv) or ".", exist_ok=True)
    metrics.to_csv(out_csv, index=False, encoding="utf-8-sig")

    try:
        print(f"Écriture GeoJSON annoté → {out_gj}")
        gdf_out = gdf.copy()
        for col in metrics.columns:
            if col == "idx":
                continue
            gdf_out[col] = metrics[col].values
        os.makedirs(os.path.dirname(out_gj) or ".", exist_ok=True)
        gdf_out.to_file(out_gj, driver="GeoJSON")
    except Exception as e:
        print(f"AVERTISSEMENT: GeoJSON de sortie non écrit ({e})")

    # Résumé console
    summary = metrics["risk_level"].value_counts().to_dict() if "risk_level" in metrics else {}
    print("\nRésumé des niveaux :", summary)
    print("\nTop 5 suspects :")
    cols = [c for c in ["risk_score","risk_reasons","area_m2","n_vertices","rectangularity","frac_right_angles"] if c in metrics.columns]
    print(metrics.sort_values("risk_score", ascending=False).head(5)[cols])
    print("\nTerminé ✅")


if __name__ == "__main__":
    main()
