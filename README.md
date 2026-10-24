# Drug Tariff Dashboard

A static dashboard for exploring historical UK Drug Tariff prices and the June 2026 snapshot.

## Run locally

Open `Drug_Tariff_Dashboard.html` in a modern browser with an internet connection. The page loads Plotly from its CDN and reads the bundled `dashboard-data.js` file from the same folder.

## Data

The published `dashboard-data.js` bundle contains 606,233 historical and June 2026 tariff observations. The source spreadsheets and consolidated CSV are excluded from this repository. The dashboard's data note describes how the June 2026 Category M extract was interpreted.

The included `build_dashboard.ps1` can regenerate the data bundle and CSV when run on Windows with Excel installed and the five source workbooks available beside the script. Those source workbooks are not included here.
