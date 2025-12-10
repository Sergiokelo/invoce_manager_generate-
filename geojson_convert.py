#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
geojson_convert.py
Convert a GeoJSON file to KML and Excel (XLSX).
- Preserves attributes (as ExtendedData in KML, columns in Excel)
- Supports Point, LineString, Polygon, and Multi* geometries
- Auto-reprojects to WGS84 (EPSG:4326) for KML if needed
- Adds lon/lat columns (Point) or centroid_lon/centroid_lat for non-point in Excel
- Attempts to fix invalid geometries (buffer(0)) when exporting KML

Usage:
    python geojson_convert.py input.geojson --kml out.kml --xlsx out.xlsx [--name-field name]

Dependencies (install once):
    pip install geopandas pandas simplekml shapely fiona pyproj openpyxl

"""
import argparse
import os
import sys
from typing import Optional, Dict, Any

import geopandas as gpd
import pandas as pd
from shapely.geometry import Point, LineString, Polygon, MultiPoint, MultiLineString, MultiPolygon, GeometryCollection
import shapely
import simplekml


def html_escape(s: Any) -> str:
    if s is None:
        return ""
    t = str(s)
    return (
        t.replace("&", "&amp;")
         .replace("<", "&lt;")
         .replace(">", "&gt;")
         .replace('"', "&quot;")
         .replace("'", "&#39;")
    )


def props_to_html(props: Dict[str, Any]) -> str:
    if not props:
        return ""
    rows = []
    for k, v in props.items():
        rows.append(f"<tr><th style='text-align:left;padding-right:8px'>{html_escape(k)}</th>"
                    f"<td>{html_escape(v)}</td></tr>")
    return "<table>{}</table>".format("".join(rows))


def ensure_wgs84(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """Reproject to EPSG:4326 if needed (for KML)."""
    if gdf.crs is None:
        # assume already WGS84 if unknown; user can set CRS earlier if needed
        return gdf.set_crs(epsg=4326, allow_override=True)
    try:
        epsg = gdf.crs.to_epsg()
    except Exception:
        epsg = None
    if epsg == 4326:
        return gdf
    return gdf.to_crs(epsg=4326)


def flatten_geometries(geom):
    """Yield individual simple geometries from possibly multi/collection geometries."""
    if geom is None:
        return
    if isinstance(geom, (Point, LineString, Polygon)):
        yield geom
    elif isinstance(geom, (MultiPoint, MultiLineString, MultiPolygon)):
        for part in geom.geoms:
            yield part
    elif isinstance(geom, GeometryCollection):
        for part in geom.geoms:
            for g in flatten_geometries(part):
                yield g
    else:
        # Unknown or custom types; try to yield as-is
        yield geom


def fix_invalid(geom):
    """Try to fix invalid polygons/lines using buffer(0)."""
    try:
        if geom and hasattr(geom, "is_valid") and not geom.is_valid:
            return geom.buffer(0)
    except Exception:
        return geom
    return geom


def add_excel_coords(df: gpd.GeoDataFrame, decimals: int = 6) -> pd.DataFrame:
    """Add lon/lat columns for Point or centroid for others; drop geometry for Excel."""
    df_wgs84 = ensure_wgs84(df)
    def get_lon_lat(g):
        if g is None:
            return (None, None)
        if isinstance(g, Point):
            return (round(g.x, decimals), round(g.y, decimals))
        else:
            c = g.centroid
            return (round(c.x, decimals), round(c.y, decimals))

    lon_vals, lat_vals = [], []
    for g in df_wgs84.geometry:
        lon, lat = get_lon_lat(g)
        lon_vals.append(lon)
        lat_vals.append(lat)

    df_tab = df_wgs84.drop(columns=["geometry"]).copy()
    is_all_points = all(isinstance(g, Point) for g in df_wgs84.geometry if g is not None)
    if is_all_points:
        df_tab.insert(0, "lat", lat_vals)
        df_tab.insert(0, "lon", lon_vals)
    else:
        df_tab.insert(0, "centroid_lat", lat_vals)
        df_tab.insert(0, "centroid_lon", lon_vals)
    return df_tab


def export_to_excel(gdf: gpd.GeoDataFrame, out_xlsx: str, sheet_name: str = "data", decimals: int = 6):
    df = add_excel_coords(gdf, decimals=decimals)
    os.makedirs(os.path.dirname(out_xlsx) or ".", exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def _add_kml_geometry(container, geom, name: str, description_html: str):
    """Add a Shapely geometry to a simplekml container as a Placemark."""
    if isinstance(geom, Point):
        pm = container.newpoint(name=name, coords=[(geom.x, geom.y)])
    elif isinstance(geom, LineString):
        pm = container.newlinestring(name=name, coords=list(geom.coords))
    elif isinstance(geom, Polygon):
        ext = list(geom.exterior.coords) if geom.exterior is not None else []
        pm = container.newpolygon(name=name, outerboundaryis=ext)
        if geom.interiors:
            pm.innerboundaryis = [list(r.coords) for r in geom.interiors]
    else:
        try:
            coords = list(getattr(geom, "coords", []))
            if coords:
                pm = container.newlinestring(name=name, coords=coords)
            else:
                c = geom.centroid
                pm = container.newpoint(name=name, coords=[(c.x, c.y)])
        except Exception:
            c = geom.centroid
            pm = container.newpoint(name=name, coords=[(c.x, c.y)])

    if description_html:
        pm.description = description_html


def export_to_kml(gdf: gpd.GeoDataFrame, out_kml: str, name_field: str = None):
    gdf_wgs84 = ensure_wgs84(gdf)
    kml = simplekml.Kml()
    fol = kml.newfolder(name=os.path.basename(out_kml))

    for idx, row in gdf_wgs84.iterrows():
        geom = row.geometry
        if geom is None or geom.is_empty:
            continue

        # Placemark name
        placename = None
        if name_field and name_field in row and pd.notna(row[name_field]):
            placename = str(row[name_field])
        else:
            placename = f"Feature {idx}"

        props = row.drop(labels=[col for col in row.index if col == "geometry"]).to_dict()
        descr_html = props_to_html(props)

        parts = list(flatten_geometries(geom))
        if len(parts) > 1:
            feat_folder = fol.newfolder(name=placename)
            for i, part in enumerate(parts, start=1):
                part = fix_invalid(part)
                _add_kml_geometry(feat_folder, part, f"{placename} (part {i})", descr_html)
        else:
            part = fix_invalid(parts[0])
            _add_kml_geometry(fol, part, placename, descr_html)

    os.makedirs(os.path.dirname(out_kml) or ".", exist_ok=True)
    kml.save(out_kml)


def run_cli_or_interactive():
    parser = argparse.ArgumentParser(description="Convert GeoJSON to KML and/or Excel (interactive if no args).")
    parser.add_argument("input", nargs="?", help="Path to input GeoJSON")
    parser.add_argument("--kml", help="Path to output KML")
    parser.add_argument("--xlsx", help="Path to output Excel (XLSX)")
    parser.add_argument("--name-field", default=None, help="Attribute field for KML placemark names")
    parser.add_argument("--sheet-name", default="data", help="Excel sheet name")
    parser.add_argument("--round", type=int, default=6, help="Decimals for lon/lat rounding in Excel")
    args = parser.parse_args()

    # Interactive mode if no input provided
    if not args.input:
        print("=== Mode interactif ===")
        input_path = input("Chemin du fichier GeoJSON : ").strip().strip('"')
        while not input_path:
            input_path = input("Chemin du fichier GeoJSON (obligatoire) : ").strip().strip('"')

        base_no_ext = os.path.splitext(input_path)[0]

        mode = ""
        while mode not in ("1", "2", "3"):
            mode = input("Exporter: (1) KML, (2) Excel, (3) Les deux ? [1/2/3] : ").strip()

        # Suggest default outputs
        out_kml = ""
        out_xlsx = ""
        if mode in ("1", "3"):
            out_kml = input(f"Chemin KML de sortie [{base_no_ext}.kml] : ").strip().strip('"')
            if not out_kml:
                out_kml = base_no_ext + ".kml"

        if mode in ("2", "3"):
            out_xlsx = input(f"Chemin Excel de sortie [{base_no_ext}.xlsx] : ").strip().strip('"')
            if not out_xlsx:
                out_xlsx = base_no_ext + ".xlsx"

        name_field = input("Nom du champ pour nommer les entités KML (laisser vide si aucun) : ").strip()
        sheet_name = input("Nom de feuille Excel [data] : ").strip() or "data"
        try:
            decimals = int(input("Décimales pour lon/lat [6] : ").strip() or "6")
        except ValueError:
            decimals = 6

        # Load and export
        try:
            gdf = gpd.read_file(input_path)
        except Exception as e:
            print(f"ERREUR lecture GeoJSON: {e}", file=sys.stderr)
            sys.exit(1)

        if out_kml:
            try:
                export_to_kml(gdf, out_kml, name_field=name_field or None)
                print(f"✅ KML sauvegardé : {out_kml}")
            except Exception as e:
                print(f"ERREUR export KML: {e}", file=sys.stderr)

        if out_xlsx:
            try:
                export_to_excel(gdf, out_xlsx, sheet_name=sheet_name, decimals=decimals)
                print(f"✅ Excel sauvegardé : {out_xlsx}")
            except Exception as e:
                print(f"ERREUR export Excel: {e}", file=sys.stderr)
        return

    # CLI path
    input_path = args.input
    try:
        gdf = gpd.read_file(input_path)
    except Exception as e:
        print(f"ERROR: Unable to read GeoJSON '{input_path}': {e}", file=sys.stderr)
        sys.exit(1)

    if not args.kml and not args.xlsx:
        print("ERROR: Specify at least one output: --kml and/or --xlsx", file=sys.stderr)
        sys.exit(2)

    if args.kml:
        try:
            export_to_kml(gdf, args.kml, name_field=args.name_field)
            print(f"✅ KML saved to: {args.kml}")
        except Exception as e:
            print(f"ERROR exporting KML: {e}", file=sys.stderr)

    if args.xlsx:
        try:
            export_to_excel(gdf, args.xlsx, sheet_name=args.sheet_name, decimals=args.round)
            print(f"✅ Excel saved to: {args.xlsx}")
        except Exception as e:
            print(f"ERROR exporting Excel: {e}", file=sys.stderr)


if __name__ == "__main__":
    run_cli_or_interactive()
