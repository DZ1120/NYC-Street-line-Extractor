import geopandas as gpd
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from shapely.geometry import Point, LineString
import os
from datetime import datetime
import folium
from folium import plugins
import xlsxwriter
import math
import matplotlib.pyplot as plt

def get_user_input():
    """Get user input for address and radius"""
    print("\n=== NYC Street Centerline Extractor ===")
    address = input("\nPlease enter a NYC address (e.g., '350 5th Ave, New York, NY'): ")
    while True:
        try:
            radius = float(input("\nPlease enter search radius (meters): "))
            if radius <= 0:
                print("Radius must be greater than 0.")
                continue
            break
        except ValueError:
            print("Please enter a valid number.")
    return address, radius

def geocode_address(address):
    """Geocode address to coordinates"""
    geolocator = Nominatim(user_agent="nyc_street_extractor")
    try:
        location = geolocator.geocode(address)
        if location:
            return Point(location.longitude, location.latitude)
        else:
            raise ValueError("Address not found.")
    except Exception as e:
        raise Exception(f"Geocoding error: {str(e)}")

def extract_streets(center_point, radius, shapefile_path):
    """Extract street centerlines within a given radius"""
    try:
        print(f"Reading Shapefile: {shapefile_path}")
        streets = gpd.read_file(shapefile_path)
        if streets.crs is None:
            streets.set_crs(epsg=4326, inplace=True)
        buffer = center_point.buffer(radius / 111000)
        streets_in_radius = streets[streets.geometry.intersects(buffer)]
        return streets_in_radius
    except Exception as e:
        raise Exception(f"Error extracting street data: {str(e)}")

def convert_coords_to_excel_coords(coords, center_lon, center_lat, scale=100):
    """Convert longitude/latitude coordinates to Excel cell coordinates"""
    excel_coords = []
    for lon, lat in coords:
        # Calculate offset relative to center point
        x = (lon - center_lon) * scale
        y = (lat - center_lat) * scale
        # Convert to Excel cell coordinates
        excel_x = int(x + 50)  # 50 is the center X
        excel_y = int(50 - y)  # 50 is the center Y, Y axis is reversed
        excel_coords.append((excel_x, excel_y))
    return excel_coords

def draw_line(worksheet, start_x, start_y, end_x, end_y, format):
    """Draw a line using cell fill"""
    dx = abs(end_x - start_x)
    dy = abs(end_y - start_y)
    sx = 1 if start_x < end_x else -1
    sy = 1 if start_y < end_y else -1
    err = dx - dy
    x, y = start_x, start_y
    while True:
        worksheet.write(y, x, '', format)
        if x == end_x and y == end_y:
            break
        e2 = 2 * err
        if e2 > -dy:
            err -= dy
            x += sx
        if e2 < dx:
            err += dx
            y += sy

def export_streets_to_excel(streets_data, center_point, output_path):
    """Export street lines as graphics to Excel (using cell fill)"""
    try:
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet("Street Map")
        worksheet.set_column(0, 100, 2)
        for i in range(100):
            worksheet.set_row(i, 15)
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center'
        })
        worksheet.merge_range('A1:Z1', "NYC Street Map", title_format)
        street_format = workbook.add_format({
            'bg_color': '#0000FF',  # Blue background
            'border': 0
        })
        center_lon, center_lat = center_point.x, center_point.y
        for idx, row in streets_data.iterrows():
            if isinstance(row.geometry, LineString):
                excel_coords = convert_coords_to_excel_coords(
                    row.geometry.coords, 
                    center_lon, 
                    center_lat
                )
                for i in range(len(excel_coords) - 1):
                    start_x, start_y = excel_coords[i]
                    end_x, end_y = excel_coords[i + 1]
                    draw_line(worksheet, start_x, start_y, end_x, end_y, street_format)
        workbook.close()
        return True
    except Exception as e:
        raise Exception(f"Error exporting to Excel: {str(e)}")

def get_shapefile_path():
    """Get absolute path of the Shapefile"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    shapefile_path = os.path.join(current_dir, "Centerline_20250522", "geo_export_afeb978c-bd19-430e-9999-6417824a9aae.shp")
    print(f"\nCurrent working directory: {os.getcwd()}")
    print(f"Shapefile path: {shapefile_path}")
    return shapefile_path

def check_shapefile_files(shapefile_path):
    """Check if all required Shapefile files exist and are accessible"""
    base_path = os.path.splitext(shapefile_path)[0]
    required_extensions = ['.shp', '.shx', '.dbf', '.prj']
    print("\nChecking the following files:")
    for ext in required_extensions:
        file_path = base_path + ext
        print(f"- {file_path}")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Required file not found: {file_path}")
        if not os.access(file_path, os.R_OK):
            raise PermissionError(f"Cannot read file: {file_path}")
    return True

def visualize_streets(streets_data, center_point, radius, output_path):
    """Visualize street data as an interactive map"""
    try:
        m = folium.Map(location=[center_point.y, center_point.x], zoom_start=15)
        folium.Circle(
            location=[center_point.y, center_point.x],
            radius=radius,
            color='red',
            fill=True,
            fill_opacity=0.2
        ).add_to(m)
        for idx, row in streets_data.iterrows():
            if isinstance(row.geometry, LineString):
                coords = [[y, x] for x, y in row.geometry.coords]
                folium.PolyLine(
                    coords,
                    color='blue',
                    weight=2,
                    opacity=0.8,
                    popup=row.get('STNAME', 'Unknown Street')
                ).add_to(m)
        legend_html = '''
            <div style="position: fixed; bottom: 50px; left: 50px; z-index: 1000; background-color: white; padding: 10px; border: 2px solid grey; border-radius: 5px;">
                <p><strong>Legend</strong></p>
                <p><span style="color: red;">●</span> Search Radius</p>
                <p><span style="color: blue;">━</span> Street</p>
            </div>
        '''
        m.get_root().html.add_child(folium.Element(legend_html))
        m.save(output_path)
        return True
    except Exception as e:
        raise Exception(f"Error generating map: {str(e)}")

def export_streets_to_svg(streets_data, output_path):
    """Export street lines as SVG vector image (no basemap, no axes)"""
    try:
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.axis('off')
        for idx, row in streets_data.iterrows():
            if isinstance(row.geometry, LineString):
                x, y = row.geometry.xy
                ax.plot(x, y, color='blue', linewidth=1)
        plt.savefig(output_path, format='svg', bbox_inches='tight', pad_inches=0, transparent=True)
        plt.close(fig)
        return True
    except Exception as e:
        raise Exception(f"Error exporting SVG: {str(e)}")

def main():
    try:
        address, radius = get_user_input()
        print("\nGeocoding address...")
        center_point = geocode_address(address)
        shapefile_path = get_shapefile_path()
        print("\nChecking Shapefile files...")
        check_shapefile_files(shapefile_path)
        print("\nExtracting street data...")
        streets_data = extract_streets(center_point, radius, shapefile_path)
        if len(streets_data) == 0:
            print("\nNo streets found within the specified radius.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        map_file = f"street_map_{timestamp}.html"
        svg_file = f"street_lines_{timestamp}.svg"
        print("\nGenerating interactive map...")
        visualize_streets(streets_data, center_point, radius, map_file)
        print("\nExporting SVG street lines...")
        export_streets_to_svg(streets_data, svg_file)
        print(f"\nSuccess!")
        print(f"Map saved to: {map_file}")
        print(f"SVG street lines saved to: {svg_file}")
        print(f"Total {len(streets_data)} streets found.")
        print("\nInsert the SVG file into Excel and ungroup to get editable street line vectors.")
    except Exception as e:
        print(f"\nError: {str(e)}")

if __name__ == "__main__":
    main() 