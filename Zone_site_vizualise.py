import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from shapely.geometry import Point, Polygon

# ===============================
# 1. Read Excel files
# ===============================
zone_file = "ZONE_INFO.xlsx"
site_file = "Property_la_lo.xlsx"

zones_df = pd.read_excel(zone_file)
sites_df = pd.read_excel(site_file)

# ===============================
# 2. Prepare colors
# ===============================
num_zones = len(zones_df)
colors = cm.get_cmap("tab20", num_zones)

# Store polygons for point-in-polygon check
zone_polygons = {}

# ===============================
# 3. Plot zones
# ===============================
plt.figure(figsize=(10, 8))

for idx, row in zones_df.iterrows():
    zone_id = row["zone_id"]
    color = colors(idx)

    # Polygon coordinates (lon, lat)
    coords = [
        (row["long1"], row["lat1"]),
        (row["long2"], row["lat2"]),
        (row["long3"], row["lat3"]),
        (row["long4"], row["lat4"])
    ]

    polygon = Polygon(coords)
    zone_polygons[zone_id] = {
        "polygon": polygon,
        "color": color
    }

    lons, lats = zip(*(coords + [coords[0]]))

    # Draw polygon
    plt.plot(lons, lats, color=color, linewidth=2)
    plt.fill(lons, lats, color=color, alpha=0.25)

    # Label zone
    centroid = polygon.centroid
    plt.text(
        centroid.x,
        centroid.y,
        f"{zone_id}",
        fontsize=9,
        fontweight="bold",
        ha="center",
        va="center"
    )

# ===============================
# 4. Plot sites & detect zones
# ===============================
for _, site in sites_df.iterrows():
    site_id = site["property_id"]
    lat = site["property_latitude"]
    lon = site["property_longitude"]

    point = Point(lon, lat)

    site_color = "black"
    assigned_zone = None

    for zone_id, data in zone_polygons.items():
        polygon = data["polygon"]

        # inside OR on boundary
        if polygon.contains(point) or polygon.touches(point):
            site_color = data["color"]
            assigned_zone = zone_id
            break

    # Plot site
    plt.scatter(lon, lat, color=site_color, marker="o", zorder=5)
    plt.text(
        lon,
        lat,
        f"{site_id}",
        fontsize=8,
        ha="left",
        va="bottom"
    )

# ===============================
# 5. Styling
# ===============================
plt.xlabel("Longitude")
plt.ylabel("Latitude")
plt.title("Zone & Site Mapping (Point-in-Polygon)")
plt.grid(True)
plt.axis("equal")

plt.show()
