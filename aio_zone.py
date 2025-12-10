import geopandas as gpd
from shapely.geometry import Polygon
import sys

URLS = [
    "https://naciscdn.org/naturalearth/110m/cultural/ne_110m_admin_0_countries.zip",
    "https://www.naturalearthdata.com/http//www.naturalearthdata.com/download/110m/cultural/ne_110m_admin_0_countries.zip",
]

def load_world():
    last_err = None
    for url in URLS:
        try:
            print(f"-> Tente : {url}")
            return gpd.read_file(url)
        except Exception as e:
            print(f"   échoué: {e}")
            last_err = e
    raise last_err

def pick_drc(world):
    cols = {c.lower(): c for c in world.columns}
    # clés candidates (selon versions Natural Earth)
    def col(name): return cols.get(name, None)

    # 1) Par code ISO (le plus robuste)
    for key in ("adm0_a3", "iso_a3"):
        k = col(key)
        if k is not None:
            drc = world[world[k].astype(str).str.upper() == "COD"]
            if len(drc):
                return drc.to_crs(4326)

    # 2) Par nom
    for key in ("admin", "name", "sovereignt"):
        k = col(key)
        if k is not None:
            m = world[k].astype(str).str.contains(
                "Democratic Republic of the Congo|Congo \\(Kinshasa\\)",
                case=False, regex=True
            )
            drc = world[m]
            if len(drc):
                return drc.to_crs(4326)

    raise RuntimeError("Impossible de trouver la RDC dans le jeu de données.")

def fallback_bbox():
    # BBox englobante (approx) de la RDC
    # longitudes ~12.0–31.5 E, latitudes ~-13.8–5.7
    coords = [(12.0, -13.8), (12.0, 5.7), (31.5, 5.7), (31.5, -13.8), (12.0, -13.8)]
    return gpd.GeoDataFrame({"name": ["RDC_bbox"]},
                            geometry=[Polygon(coords)], crs=4326)

def main(out_path="aoi_rdc.geojson"):
    try:
        world = load_world()
        drc = pick_drc(world)
        print(f"OK: {len(drc)} entité(s) RDC trouvée(s).")
    except Exception as e:
        print(f"[AVERTISSEMENT] Source Natural Earth indisponible ({e}).")
        print("-> Utilisation d'un polygone BBOX de secours.")
        drc = fallback_bbox()

    drc.to_file(out_path, driver="GeoJSON")
    print(f"✅ AOI écrit : {out_path}")

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "aoi_rdc.geojson"
    main(out)
