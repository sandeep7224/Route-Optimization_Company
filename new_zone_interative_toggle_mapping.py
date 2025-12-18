import pandas as pd
import folium
import matplotlib.cm as cm
import matplotlib.colors as mcolors
import re
from shapely.geometry import Point, Polygon
from folium.features import DivIcon

# =====================================================
# 1. FILE PATHS
# =====================================================
ZONE_FILE = "Data_Zone-2.xlsx"
SITE_FILE = "Property_la_lo.xlsx"
OFFICER_FILE = "officer.xlsx"

# =====================================================
# 2. READ EXCEL FILES
# =====================================================
zones_df = pd.read_excel(ZONE_FILE)
sites_df = pd.read_excel(SITE_FILE)
officers_df = pd.read_excel(OFFICER_FILE)

# =====================================================
# 3. CLEAN & NORMALIZE ZONE COLUMN NAMES (EXCEL SAFE)
# =====================================================
zones_df.columns = (
    zones_df.columns
    .astype(str)
    .str.lower()
    .str.replace(r"_x000d_", "", regex=True)
    .str.replace(r"\s+", "", regex=True)
)

# =====================================================
# 4. FORCE LAT/LONG COLUMNS TO NUMERIC
# =====================================================
for col in zones_df.columns:
    if col.startswith("lat") or col.startswith("long"):
        zones_df[col] = pd.to_numeric(zones_df[col], errors="coerce")

# =====================================================
# 5. DETECT & SORT LAT/LONG COLUMNS (REGEX SAFE)
# =====================================================
lat_cols = [c for c in zones_df.columns if c.startswith("lat")]
lon_cols = [c for c in zones_df.columns if c.startswith("long")]

def extract_index(col):
    m = re.search(r"\d+", col)
    return int(m.group()) if m else 0

lat_cols.sort(key=extract_index)
lon_cols.sort(key=extract_index)

if not lat_cols or not lon_cols:
    raise ValueError(f"❌ No lat/long columns found: {zones_df.columns.tolist()}")

# =====================================================
# 6. MAP CENTER (100% TYPE SAFE)
# =====================================================
all_lats, all_lons = [], []

for _, row in zones_df.iterrows():
    for lat_c, lon_c in zip(lat_cols, lon_cols):
        lat = row[lat_c]
        lon = row[lon_c]
        if pd.notna(lat) and pd.notna(lon):
            try:
                all_lats.append(float(lat))
                all_lons.append(float(lon))
            except ValueError:
                continue

if not all_lats or not all_lons:
    raise ValueError("❌ No valid zone coordinates found")

center_lat = sum(all_lats) / len(all_lats)
center_lon = sum(all_lons) / len(all_lons)

m = folium.Map(
    location=[center_lat, center_lon],
    zoom_start=12,
    tiles="OpenStreetMap"
)

# =====================================================
# 7. FEATURE GROUPS
# =====================================================
zone_layer = folium.FeatureGroup(name="Zones", show=True)
site_layer = folium.FeatureGroup(name="Sites", show=True)
officer_layer = folium.FeatureGroup(name="Field Officers", show=True)

# =====================================================
# 8. COLOR MAP FOR ZONES
# =====================================================
cmap = cm.get_cmap("tab20", len(zones_df))
def to_hex(c): return mcolors.to_hex(c)

# =====================================================
# 9. BUILD ZONE POLYGONS
# =====================================================
zone_polygons = {}
zone_site_count = {}

for idx, row in zones_df.iterrows():
    zone_id = row["zone"]
    coords = []

    for lat_c, lon_c in zip(lat_cols, lon_cols):
        lat = row[lat_c]
        lon = row[lon_c]
        if pd.notna(lat) and pd.notna(lon):
            try:
                coords.append((float(lon), float(lat)))
            except ValueError:
                continue

    if len(coords) < 3:
        continue

    zone_polygons[zone_id] = {
        "polygon": Polygon(coords),
        "coords": coords,
        "color": to_hex(cmap(idx))
    }

    zone_site_count[zone_id] = 0

# =====================================================
# 10. COUNT SITES PER ZONE
# =====================================================
for _, site in sites_df.iterrows():
    try:
        point = Point(float(site["property_longitude"]), float(site["property_latitude"]))
    except Exception:
        continue

    for zid, data in zone_polygons.items():
        if data["polygon"].contains(point) or data["polygon"].touches(point):
            zone_site_count[zid] += 1
            break

# =====================================================
# 11. ADD ZONES TO MAP
# =====================================================
for zid, data in zone_polygons.items():
    folium_coords = [[lat, lon] for lon, lat in data["coords"]]

    folium.Polygon(
        locations=folium_coords,
        color=data["color"],
        fill=True,
        fill_color=data["color"],
        fill_opacity=0.35,
        tooltip=f"Zone {zid} | Sites: {zone_site_count[zid]}",
        popup=f"<b>Zone:</b> {zid}<br><b>Total Sites:</b> {zone_site_count[zid]}"
    ).add_to(zone_layer)

# =====================================================
# 12. ADD SITES
# =====================================================
for _, site in sites_df.iterrows():
    try:
        site_id = site["property_id"]
        lat = float(site["property_latitude"])
        lon = float(site["property_longitude"])
    except Exception:
        continue

    point = Point(lon, lat)
    inside = False
    zone_name = "Outside"

    for zid, data in zone_polygons.items():
        if data["polygon"].contains(point) or data["polygon"].touches(point):
            inside = True
            zone_name = zid
            break

    color = "blue" if inside else "black"

    folium.CircleMarker(
        location=[lat, lon],
        radius=6,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=1,
        popup=f"<b>Site ID:</b> {site_id}<br><b>Zone:</b> {zone_name}"
    ).add_to(site_layer)

    folium.Marker(
        location=[lat, lon],
        icon=DivIcon(
            icon_size=(30, 30),
            icon_anchor=(0, 0),
            html=f"""
            <div style="font-size:10px;font-weight:bold;color:{color};
                        background:white;padding:1px 4px;
                        border:1px solid #888;border-radius:3px;">
                {site_id}
            </div>
            """
        )
    ).add_to(site_layer)

# =====================================================
# 13. ADD FIELD OFFICERS
# =====================================================
for _, off in officers_df.iterrows():
    try:
        off_id = off["off_id"]
        lat = float(off["lat"])
        lon = float(off["long"])
    except Exception:
        continue

    folium.Marker(
        location=[lat, lon],
        icon=folium.Icon(color="red", icon="user", prefix="fa"),
        popup=f"<b>Officer ID:</b> {off_id}"
    ).add_to(officer_layer)

    folium.Marker(
        location=[lat, lon],
        icon=DivIcon(
            icon_size=(30, 30),
            icon_anchor=(0, 0),
            html=f"""
            <div style="font-size:11px;font-weight:bold;color:red;
                        background:white;padding:2px 4px;
                        border:1px solid red;border-radius:4px;">
                {off_id}
            </div>
            """
        )
    ).add_to(officer_layer)

# =====================================================
# 14. FINALIZE MAP
# =====================================================
zone_layer.add_to(m)
site_layer.add_to(m)
officer_layer.add_to(m)

folium.LayerControl(collapsed=False).add_to(m)

m.save("yagyank_interactive_zone_site_officer_map.html")

print("✅ Map generated successfully: updates_interactive_zone_site_officer_map.html")