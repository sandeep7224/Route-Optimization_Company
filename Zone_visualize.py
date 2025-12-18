import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm

# ===============================
# 1. Read Excel file
# ===============================
file_path = "ZONE_INFO.xlsx"   # change path if needed
df = pd.read_excel(file_path)

# ===============================
# 2. Prepare color map
# ===============================
num_zones = len(df)
colors = cm.get_cmap("tab20", num_zones)  # good for distinct colors

# ===============================
# 3. Create plot
# ===============================
plt.figure(figsize=(10, 8))

for idx, row in df.iterrows():
    zone_id = row["zone_id"]
    color = colors(idx)

    # Longitude = X, Latitude = Y
    lons = [row["long1"], row["long2"], row["long3"], row["long4"], row["long1"]]
    lats = [row["lat1"], row["lat2"], row["lat3"], row["lat4"], row["lat1"]]

    # Plot polygon
    plt.plot(
        lons,
        lats,
        marker="o",
        linewidth=2,
        color=color,
        label=f"Zone {zone_id}"
    )

    # Fill polygon with transparent color
    plt.fill(
        lons,
        lats,
        color=color,
        alpha=0.25
    )

    # Label zone at centroid
    centroid_lon = sum(lons[:-1]) / 4
    centroid_lat = sum(lats[:-1]) / 4
    plt.text(
        centroid_lon,
        centroid_lat,
        f"{zone_id}",
        fontsize=9,
        ha="center",
        va="center",
        fontweight="bold"
    )

# ===============================
# 4. Plot styling
# ===============================
plt.xlabel("Longitude")
plt.ylabel("Latitude")
plt.title("Zone Visualization with Unique Colors")
plt.grid(True)
plt.axis("equal")   # critical for geo accuracy
plt.legend(loc="best")

# ===============================
# 5. Show plot
# ===============================
plt.show()
