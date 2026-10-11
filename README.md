# Adani Total Gas CNG Station Scraper

Python automation that queries the Adani Gas area and CNG-station APIs, consolidates station records, deduplicates them by location number, and exports an Excel workbook.

## How it works

1. Iterates through the city/municipality search values configured in the script.
2. Sends a GET request to the area endpoint for each configured value.
3. Sends a POST request for each returned area to retrieve station records.
4. Combines the station responses, removes duplicates using LocNo, and derives standardized address and map-link fields.
5. Sorts the final dataset and writes an Excel workbook.

## Output data

The workbook columns are:

- LocNo — source location identifier
- Station Name
- Address — assembled from the source address fields
- NearArea
- City and State
- PinCode and MobileNo
- Latitude and Longitude
- Google Maps — map query constructed from the coordinates

The script reports raw record count and unique station count in the terminal.

## Files

- Adani_Total_Gas_Stations_scraper.py — scraper entry point
- `Adani_Total_Gas_Stations.xlsx` — existing scraper output workbook
- requirements.txt — base Python dependencies

## Setup

Python 3.10 or a compatible version is recommended. Create a virtual environment and install the project dependencies plus the Excel writer engine:

    python -m venv .venv
    .venv/Scripts/Activate.ps1
    pip install -r requirements.txt openpyxl

## Configuration and run

The source endpoints, request headers, and cities search list are declared near the top of Adani_Total_Gas_Stations_scraper.py. Update the list to adjust the search coverage.

Run:

    python Adani_Total_Gas_Stations_scraper.py

The workbook Adani_Total_Gas_Stations.xlsx is written to the current working directory.

## Notes

- Excel export requires openpyxl; the repository's requirements list may need to be supplemented as shown above.
- Coverage is limited to the configured search values and the results returned by the source endpoints.
- Source APIs and returned fields can change. Follow source terms and applicable data-use policies.

## Synthetic sample data

[`sample_synthetic_12_records.csv`](sample_synthetic_12_records.csv) contains **12 synthetic test records** using the scraper's output schema. These records were generated for schema inspection and portfolio/demo purposes; they were not collected from the source website or API. Names and addresses are fictional, contact details and example.com links are placeholders, coordinates are illustrative, and ratings/prices are not real or current. Do not use them for real-world navigation or pricing decisions.
