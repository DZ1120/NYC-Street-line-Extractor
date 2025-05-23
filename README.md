# NYC Street Centerline Extractor

This Python project allows you to extract all street centerlines within a specified radius of any address in New York City, and export the results as an interactive map (HTML) and an editable SVG vector file (for Excel ungrouping and editing).

## Features

- Interactive address input
- Customizable search radius (in meters)
- Automatic geocoding
- Spatial query and buffer analysis
- Export to:
  - Interactive HTML map (with Folium)
  - Editable SVG vector file (for Excel ungrouping)
  - (Optional) Visual Excel file (cell fill, see code)

## Requirements

- Python 3.7 or higher
- Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```

## Data

NYC street centerline Shapefile data is required for this project.

You can download the latest official data from NYC Open Data:
[NYC Street Centerline (CSCL) Shapefile - NYC Open Data](https://data.cityofnewyork.us/Transportation/NYC-Street-Centerline-CSCL-/i4gi-tjb9)

After downloading, extract the files into the Centerline_20250522/ directory (or update the script to point to your data folder).

## Usage

1. Make sure all dependencies are installed.
2. Run the script:
   ```bash
   python nyc_street_extractor.py
   ```
3. Enter a NYC address (e.g., `350 5th Ave, New York, NY`)
4. Enter the search radius (in meters)
5. Wait for processing. Results will be saved as:
   - `street_map_*.html` (interactive map)
   - `street_lines_*.svg` (editable vector for Excel)
   - (Optional) `street_lines_*.xlsx` (visual cell fill)

## Output

- **HTML Map:** Interactive, viewable in browser.
- **SVG:** Insert into Excel, right-click and ungroup to get editable street lines.
- **Excel:** (Optional) Visual representation using cell fill.

## Notes

- Requires internet connection for geocoding.
- Address input should be as accurate as possible.
- Radius is in meters.
- Output filenames include a timestamp to avoid overwriting.

## License

MIT License

---

**Enjoy mapping NYC streets!** 
