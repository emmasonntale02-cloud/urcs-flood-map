# URCS Community Flood Early Warning Network – Interactive Map

Live map: https://emmasonntale02-cloud.github.io/urcs-flood-map/

Interactive map of the community river gauges, automatic flood monitoring stations and
community radios installed by the Uganda Red Cross Society.

## How to update the map

1. Edit your copy of the spreadsheet (keep the same columns).
2. On GitHub, open the **`data`** folder → **Add file → Upload files**.
3. Drag in the new spreadsheet. If its file name is different from the old one,
   delete the old file too: the `data` folder must contain **only one** `.xlsx` file.
4. Click **Commit changes**.

About 2 minutes later the live map shows the new data. Progress is shown on the
**Actions** tab: a green tick means the map is updated.

If the spreadsheet has a problem, the run shows a red cross, the live map stays as it was,
and GitHub emails you. Open the failed run → **build** → **Check the spreadsheet and build the map**
to see which row to fix, for example:

    Row 3 (Site ID 2): coordinates 34.0, 0.87 are outside Uganda. Latitude and longitude look swapped.

## Spreadsheet rules

| Column | Notes |
|---|---|
| Site ID | Whole number, unique |
| District / Area | e.g. `Kasese` |
| Monitoring location / river reach | Shown as the site name |
| Sensor technology / type | Must contain one of: `Community Radio`, `Graduated river level gauge`, `OTT` / `Radar` / `telemetric` (automatic station) |
| GPS coordinates (as supplied) | Optional; worked out from the decimals if empty |
| Latitude / Longitude (decimal, WGS84) | Required; must be inside Uganda |
| Short name (map label) | Optional label next to the icon; made from the location if empty |

The columns are found by their header text, so their order can change.

## Offline copy

`index.html` in this repository is always the latest map as a single file. Download it
(open it → **Download raw file**) and double-click to use it without internet.

## For maintainers

- `build/build.py` checks the spreadsheet and builds the page; run `python build/build.py data site` locally.
- `build/template.html` is the page design; `build/basemap.json` holds the boundaries, lakes and rivers (Natural Earth).
- `.github/workflows/build-map.yml` runs the build and publishes to GitHub Pages.
