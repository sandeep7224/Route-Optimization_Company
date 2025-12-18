import pandas as pd
import folium
import matplotlib.cm as cm
import matplotlib.colors as mcolors
from shapely.geometry import Point, Polygon
from folium.features import DivIcon

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

# ===============================
# 4. Feature Groups (LayerControl)
# ===============================
zone_layer = folium.FeatureGroup(name="Zones", show=True)
site_layer = folium.FeatureGroup(name="Sites", show=True)
officer_layer = folium.FeatureGroup(name="Field Officers", show=True)

# ===============================
# 5. Build zone polygons
# ===============================
zone_polygons = {}
zone_site_count = {}

for idx, row in zones_df.iterrows():
    zone_id = row["zone_id"]
    color = to_hex(cmap(idx))

    polygon = Polygon([
        (row["long1"], row["lat1"]),
        (row["long2"], row["lat2"]),
        (row["long3"], row["lat3"]),
        (row["long4"], row["lat4"])
    ])

    zone_polygons[zone_id] = {
        "polygon": polygon,
        "color": color
    }
    zone_site_count[zone_id] = 0

# ===============================
# 6. Count sites per zone
# ===============================
for _, site in sites_df.iterrows():
    point = Point(site["property_longitude"], site["property_latitude"])

    for zone_id, data in zone_polygons.items():
        polygon = data["polygon"]
        if polygon.contains(point) or polygon.touches(point):
            zone_site_count[zone_id] += 1
            break

# ===============================
# 7. Add zones to map
# ===============================
for idx, row in zones_df.iterrows():
    zone_id = row["zone_id"]
    color = zone_polygons[zone_id]["color"]
    site_count = zone_site_count[zone_id]

    folium_coords = [
        [row["lat1"], row["long1"]],
        [row["lat2"], row["long2"]],
        [row["lat3"], row["long3"]],
        [row["lat4"], row["long4"]],
        [row["lat1"], row["long1"]]
    ]

    folium.Polygon(
        locations=folium_coords,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.35,
        tooltip=f"Zone {zone_id} | Sites: {site_count}",
        popup=f"""
        <b>Zone ID:</b> {zone_id}<br>
        <b>Total Sites:</b> {site_count}
        """
    ).add_to(zone_layer)

# ===============================
# 8. Add sites (blue / black + visible labels)
# ===============================
for _, site in sites_df.iterrows():
    site_id = site["property_id"]
    lat = site["property_latitude"]
    lon = site["property_longitude"]

    point = Point(lon, lat)

    inside_zone = False
    zone_name = "Outside all zones"

    for zid, data in zone_polygons.items():
        polygon = data["polygon"]
        if polygon.contains(point) or polygon.touches(point):
            inside_zone = True
            zone_name = zid
            break

    site_color = "blue" if inside_zone else "black"

    # Site marker
    folium.CircleMarker(
        location=[lat, lon],
        radius=6,
        color=site_color,
        fill=True,
        fill_color=site_color,
        fill_opacity=1,
        popup=f"""
        <b>Site ID:</b> {site_id}<br>
        <b>Zone:</b> {zone_name}<br>
        <b>Latitude:</b> {lat}<br>
        <b>Longitude:</b> {lon}
        """
    ).add_to(site_layer)

    # Always-visible Site ID
    folium.Marker(
        location=[lat, lon],
        icon=DivIcon(
            icon_size=(120, 30),
            icon_anchor=(0, 0),
            html=f"""
            <div style="
                font-size:10px;
                font-weight:bold;
                color:{site_color};
                background:white;
                padding:1px 3px;
                border-radius:3px;
                border:1px solid #999;
            ">
                {site_id}
            </div>
            """
        )
    ).add_to(site_layer)

# ===============================
# 9. Add field officers (with visible labels)
# ===============================
for _, officer in officers_df.iterrows():
    off_id = officer["off_id"]
    lat = officer["lat"]
    lon = officer["long"]

    # Officer marker
    folium.Marker(
        location=[lat, lon],
        popup=f"""
        <b>Officer ID:</b> {off_id}<br>
        <b>Latitude:</b> {lat}<br>
        <b>Longitude:</b> {lon}
        """,
        icon=folium.Icon(color="red", icon="user", prefix="fa")
    ).add_to(officer_layer)

    # Always-visible Officer ID
    folium.Marker(
        location=[lat, lon],
        icon=DivIcon(
            icon_size=(120, 30),
            icon_anchor=(0, 0),
            html=f"""
            <div style="
                font-size:11px;
                font-weight:bold;
                color:red;
                background:white;
                padding:2px 4px;
                border-radius:4px;
                border:1px solid red;
            ">
                {off_id}
            </div>
            """
        )
    ).add_to(officer_layer)

# ===============================
# 10. Add layers & controls
# ===============================
zone_layer.add_to(m)
site_layer.add_to(m)
officer_layer.add_to(m)

folium.LayerControl(collapsed=False).add_to(m)

# ===============================
# 11. Save map
# ===============================
m.save("zone_site_officer_map.html")

print("✅ Interactive map saved as zone_site_officer_map.html")
