import pandas as pd
import folium
import matplotlib.cm as cm
import matplotlib.colors as mcolors
from shapely.geometry import Point, Polygon

# ===============================
# 1. Read Excel files
# ===============================
zone_file = "ZONE_INFO.xlsx"
site_file = "Property_la_lo.xlsx"
officer_file = "officer.xlsx"

zones_df = pd.read_excel(zone_file)
sites_df = pd.read_excel(site_file)
officers_df = pd.read_excel(officer_file)

# ===============================
# 2. Prepare colors
# ===============================
num_zones = len(zones_df)
cmap = cm.get_cmap("tab20", num_zones)

def to_hex(color):
    return mcolors.to_hex(color)

# ===============================
# 3. Create base map
# ===============================
center_lat = zones_df[["lat1", "lat2", "lat3", "lat4"]].values.mean()
center_lon = zones_df[["long1", "long2", "long3", "long4"]].values.mean()

m = folium.Map(
    location=[center_lat, center_lon],
    zoom_start=13,
    tiles="OpenStreetMap"
)

zone_polygons = {}

# ===============================
# 4. Add zones
# ===============================
for idx, row in zones_df.iterrows():
    zone_id = row["zone_id"]
    color = to_hex(cmap(idx))

    folium_coords = [
        [row["lat1"], row["long1"]],
        [row["lat2"], row["long2"]],
        [row["lat3"], row["long3"]],
        [row["lat4"], row["long4"]],
        [row["lat1"], row["long1"]]
    ]

    polygon = Polygon([
        (row["long1"], row["lat1"]),
        (row["long2"], row["lat2"]),
        (row["long3"], row["lat3"]),
        (row["long4"], row["lat4"])
    ])

    zone_polygons[zone_id] = polygon

    folium.Polygon(
        locations=folium_coords,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.35,
        tooltip=f"Zone {zone_id}",
        popup=f"<b>Zone ID:</b> {zone_id}"
    ).add_to(m)

# ===============================
# 5. Add sites (Blue / Black)
# ===============================
for _, site in sites_df.iterrows():
    site_id = site["property_id"]
    lat = site["property_latitude"]
    lon = site["property_longitude"]

    point = Point(lon, lat)

    inside_zone = False
    zone_name = "Outside all zones"

    for zid, polygon in zone_polygons.items():
        if polygon.contains(point) or polygon.touches(point):
            inside_zone = True
            zone_name = zid
            break

    site_color = "blue" if inside_zone else "black"

    folium.CircleMarker(
        location=[lat, lon],
        radius=6,
        color=site_color,
        fill=True,
        fill_color=site_color,
        fill_opacity=1,
        tooltip=f"Site {site_id}",
        popup=f"""
        <b>Site ID:</b> {site_id}<br>
        <b>Zone:</b> {zone_name}<br>
        <b>Latitude:</b> {lat}<br>
        <b>Longitude:</b> {lon}
        """
    ).add_to(m)

# ===============================
# 6. Add field officers
# ===============================
for _, officer in officers_df.iterrows():
    off_id = officer["off_id"]
    lat = officer["lat"]
    lon = officer["long"]

    folium.Marker(
        location=[lat, lon],
        tooltip=f"Officer {off_id}",
        popup=f"""
        <b>Officer ID:</b> {off_id}<br>
        <b>Latitude:</b> {lat}<br>
        <b>Longitude:</b> {lon}
        """,
        icon=folium.Icon(color="red", icon="user", prefix="fa")
    ).add_to(m)

# ===============================
# 7. Save map
# ===============================
m.save("zone_site_officer_map.html")

print("✅ Interactive map saved as zone_site_officer_map.html")
