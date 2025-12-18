import pandas as pd
import folium
from shapely.geometry import Polygon, Point, LineString
from shapely.ops import split
from folium.features import DivIcon

# ===============================
# FILE PATHS
# ===============================
ZONE_FILE = "Data_Zone-2.xlsx"
SITE_FILE = "Property_la_lo.xlsx"
OFFICER_FILE = "officer.xlsx"

TARGET_ZONE = "Z7"   # zone id as string

# ===============================
# READ DATA
# ===============================
zones_df = pd.read_excel(ZONE_FILE)
sites_df = pd.read_excel(SITE_FILE)
officers_df = pd.read_excel(OFFICER_FILE)

zones_df.columns = zones_df.columns.str.strip().str.lower()
sites_df.columns = sites_df.columns.str.strip().str.lower()
officers_df.columns = officers_df.columns.str.strip().str.lower()

# ===============================
# DETECT LAT/LONG COLUMNS
# ===============================
lat_cols = sorted([c for c in zones_df.columns if c.startswith("lat")],
                  key=lambda x: int(x.replace("lat", "")))
lon_cols = sorted([c for c in zones_df.columns if c.startswith("long")],
                  key=lambda x: int(x.replace("long", "")))

# ===============================
# GET ZONE 7 ROW
# ===============================
zones_df["zone"] = zones_df["zone"].astype(str).str.strip()
zone_df = zones_df[zones_df["zone"] == TARGET_ZONE]

if zone_df.empty:
    raise Exception(f"Zone {TARGET_ZONE} not found")

row = zone_df.iloc[0]

# ===============================
# BUILD ZONE POLYGON
# ===============================
coords = []
for lat_c, lon_c in zip(lat_cols, lon_cols):
    if pd.notna(row[lat_c]) and pd.notna(row[lon_c]):
        coords.append((row[lon_c], row[lat_c]))

zone_polygon = Polygon(coords)

# ===============================
# SPLIT LINE: P3 → SPLIT → P6
# ===============================
p3 = (row["long3"], row["lat3"])
split_pt = (row["split_long"], row["split_lat"])
p6 = (row["long6"], row["lat6"])

cut_line = LineString([p3, split_pt, p6])

# ===============================
# SPLIT POLYGON
# ===============================
result = split(zone_polygon, cut_line)

if len(result.geoms) != 2:
    raise Exception("Zone split failed – check split point location")

poly1, poly2 = result.geoms

# ===============================
# INNER / OUTER DECISION
# ===============================
if poly1.area < poly2.area:
    inner_poly, outer_poly = poly1, poly2
else:
    inner_poly, outer_poly = poly2, poly1

# ===============================
# MAP CENTER
# ===============================
center_lat = sum([p[1] for p in coords]) / len(coords)
center_lon = sum([p[0] for p in coords]) / len(coords)

m = folium.Map(location=[center_lat, center_lon], zoom_start=13)

zone_layer = folium.FeatureGroup(name="Zone 7 (Inner / Outer)")
site_layer = folium.FeatureGroup(name="Sites")
officer_layer = folium.FeatureGroup(name="Officers")

# ===============================
# DRAW INNER ZONE
# ===============================
folium.Polygon(
    locations=[[lat, lon] for lon, lat in inner_poly.exterior.coords],
    color="green",
    fill=True,
    fill_opacity=0.5,
    tooltip="Zone 7 - INNER"
).add_to(zone_layer)

# ===============================
# DRAW OUTER ZONE
# ===============================
folium.Polygon(
    locations=[[lat, lon] for lon, lat in outer_poly.exterior.coords],
    color="orange",
    fill=True,
    fill_opacity=0.4,
    tooltip="Zone 7 - OUTER"
).add_to(zone_layer)

# ===============================
# DRAW SPLIT LINE (NO POINT MARKER)
# ===============================
folium.PolyLine(
    locations=[
        [row["lat3"], row["long3"]],
        [row["split_lat"], row["split_long"]],
        [row["lat6"], row["long6"]],
    ],
    color="red",
    weight=3,
    dash_array="5,5",
    tooltip="Zone Split Line"
).add_to(zone_layer)

# ===============================
# ADD SITES
# ===============================
for _, s in sites_df.iterrows():
    pt = Point(s["property_longitude"], s["property_latitude"])

    if inner_poly.contains(pt) or inner_poly.touches(pt):
        color = "blue"
    elif outer_poly.contains(pt) or outer_poly.touches(pt):
        color = "black"
    else:
        continue

    folium.CircleMarker(
        location=[s["property_latitude"], s["property_longitude"]],
        radius=6,
        color=color,
        fill=True,
        fill_color=color
    ).add_to(site_layer)

    folium.Marker(
        location=[s["property_latitude"], s["property_longitude"]],
        icon=DivIcon(
            html=f"""
            <div style="font-size:10px;font-weight:bold;color:{color}">
            {s['property_id']}
            </div>
            """
        )
    ).add_to(site_layer)

# ===============================
# ADD OFFICERS
# ===============================
for _, o in officers_df.iterrows():
    folium.Marker(
        location=[o["lat"], o["long"]],
        icon=folium.Icon(color="red", icon="user", prefix="fa"),
        popup=o["off_id"]
    ).add_to(officer_layer)

    folium.Marker(
        location=[o["lat"], o["long"]],
        icon=DivIcon(
            html=f"""
            <div style="font-size:11px;font-weight:bold;color:red">
            {o['off_id']}
            </div>
            """
        )
    ).add_to(officer_layer)

# ===============================
# FINALIZE MAP
# ===============================
zone_layer.add_to(m)
site_layer.add_to(m)
officer_layer.add_to(m)

folium.LayerControl(collapsed=False).add_to(m)

m.save("zone7_inner_outer_with_split_point.html")

print("✅ Zone 7 split using split point completed")
