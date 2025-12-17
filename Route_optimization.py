# ============================================================
# Dynamic Zone-Based Field Officer Allocation (FINAL)
# ============================================================

import pandas as pd
import math
from shapely.geometry import Point, Polygon
from geopy.distance import geodesic

# ============================================================
# Helper Functions
# ============================================================

def calculate_distance(lat1, lon1, lat2, lon2):
    return geodesic((lat1, lon1), (lat2, lon2)).km


def build_zone_polygon(row):
    """
    Build polygon from zone table:
    lat1,long1 ... lat4,long4
    """
    coords = []
    for i in range(1, 5):
        lat_col = f"lat{i}"
        lon_col = f"long{i}"
        if not pd.isna(row[lat_col]) and not pd.isna(row[lon_col]):
            coords.append((row[lon_col], row[lat_col]))
    return Polygon(coords)


def find_current_zone(point, zones_df):
    """
    Find zone where officer is currently present
    """
    for _, zone in zones_df.iterrows():
        if zone["polygon"].contains(point):
            return zone["zone"], zone["polygon"]
    return None, None


def distance_after_current_zone_exit(current_zone_polygon, site_point):
    """
    Rule 4:
    Distance traveled after exiting CURRENT zone
    """
    if current_zone_polygon is None:
        return 0.0

    if current_zone_polygon.contains(site_point):
        return 0.0

    boundary_point = current_zone_polygon.exterior.interpolate(
        current_zone_polygon.exterior.project(site_point)
    )

    return geodesic(
        (boundary_point.y, boundary_point.x),
        (site_point.y, site_point.x)
    ).km


# ============================================================
# Scoring Logic
# ============================================================

def calculate_officer_score(officer, site, zones_df):
    score = 0.0

    officer_point = Point(officer["long"], officer["lat"])
    site_point = Point(site["property_longitude"], site["property_latitude"])

    # Rule A: Idle / Active
    if officer["Active (Y/N)"] == "Y":
        score += 0.3

    # Find officer current zone
    _, current_zone_polygon = find_current_zone(officer_point, zones_df)

    # Rule B: Site inside officer current zone
    if current_zone_polygon and current_zone_polygon.contains(site_point):
        score += 0.4

    # Rule C: Distance from officer to site
    dist_to_site = calculate_distance(
        officer["lat"], officer["long"],
        site["property_latitude"], site["property_longitude"]
    )

    if dist_to_site <= 10:
        score += (0.4 - dist_to_site * 0.04)

    # Rule D: Distance after crossing CURRENT zone
    outside_distance = distance_after_current_zone_exit(
        current_zone_polygon, site_point
    )

    if outside_distance <= 10:
        score += (0.2 - outside_distance * 0.02)

    return score, dist_to_site


# ============================================================
# Allocation Engine
# ============================================================

def allocate_sites(officers_df, sites_df, zones_df):
    allocations = []

    for _, site in sites_df.iterrows():

        best_score = -math.inf
        best_idx = None
        best_distance = math.inf

        for idx, officer in officers_df.iterrows():
            score, dist = calculate_officer_score(officer, site, zones_df)

            # Tie-breaker: nearest officer
            if score > best_score or (score == best_score and dist < best_distance):
                best_score = score
                best_idx = idx
                best_distance = dist

        chosen_officer = officers_df.loc[best_idx]

        allocations.append({
            "request_id": site["request_id"],
            "customer_name": site["customer_name"],
            "assigned_FO_Id": chosen_officer["FO Id"],
            "assigned_FO_Name": chosen_officer["Field officer Name"],
            "site_lat": site["property_latitude"],
            "site_lon": site["property_longitude"],
            "final_score": round(best_score, 3)
        })

        # Sequential update of officer location & status
        officers_df.at[best_idx, "lat"] = site["property_latitude"]
        officers_df.at[best_idx, "long"] = site["property_longitude"]
        officers_df.at[best_idx, "Active (Y/N)"] = "N"

    return pd.DataFrame(allocations), officers_df


# ============================================================
# Main Execution
# ============================================================

if __name__ == "__main__":
    
    SITE_FILE =    r"C:\Users\Dell\Pictures\sites.xlsx"
    ZONE_FILE =    r"C:\Users\Dell\Pictures\zone.xlsx"
    OFFICER_FILE = r"C:\Users\Dell\Pictures\off.xlsx"

    OUTPUT_ALLOC = "final_site_allocation.xlsx"
    OUTPUT_UPDATED_OFFICERS = "updated_field_officers.xlsx"

    # Load Excel files
    officers_df = pd.read_excel(OFFICER_FILE)
    sites_df = pd.read_excel(SITE_FILE)
    zones_df = pd.read_excel(ZONE_FILE)

    # Build zone polygons
    zones_df["polygon"] = zones_df.apply(build_zone_polygon, axis=1)

    # Run allocation
    allocation_df, updated_officers_df = allocate_sites(
        officers_df, sites_df, zones_df
    )

    # Save output
    allocation_df.to_excel(OUTPUT_ALLOC, index=False)
    updated_officers_df.to_excel(OUTPUT_UPDATED_OFFICERS, index=False)

    print("✅ Allocation completed successfully")
