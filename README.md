# Hotel BI Template

A Power BI reporting template for hotel operations, built around **Opera PMS** exports. Includes a pre-built report, Power Query data-cleaning templates, branded themes, and sample data files.

## Contents

```
Hotel BI/
├── Power BI/
│   └── Hotel_Report.pbix          # Main Power BI report template
└── Raw Exports/
    ├── Daily Flash/               # Daily Flash Excel exports (sample: March 2024)
    ├── Market Segments/           # Market Segment Summary reports
    ├── Package Forecast/          # Package Forecast reports
    ├── Profile Production Statistics/  # Opera XML profile stats
    ├── Reservations/              # Reservation Activity XML exports (Q1 2024–2026)
    └── Room Status/               # Monthly Statistics reports
Highgate_PowerBI_Theme.json        # Highgate Hotels & Resorts colour theme
IHG_PowerBI_Theme.json             # IHG colour theme
Opera_PowerQuery_Template.md       # M code snippets for cleaning Opera exports
```

## Getting Started

### 1. Set Up Your Data Folders

Create local folders where Opera will drop its exports, then update the Power BI parameters to point to them:

| Parameter | Description |
|---|---|
| `HotelName` | Displayed in report titles and headers |
| `TotalRooms` | Total bookable rooms in the property |
| `FiscalYearStartMonth` | `1` = Jan, `4` = Apr, etc. |
| `CurrencySymbol` | e.g. `€`, `£`, `$` |
| `DateFormat` | `"DMY"` or `"MDY"` — match your Opera locale |
| `FolderPath_Flash` | Folder where Daily Flash exports are saved |
| `FolderPath_Reservations` | Folder for Reservation Activity exports |
| `FolderPath_Segments` | Folder for Market Segment exports |

### 2. Apply the Power Query Template

The `Opera_PowerQuery_Template.md` file contains ready-to-use M code for ingesting and cleaning each Opera report type. Copy each snippet into **Power BI Desktop → Advanced Editor** for the corresponding query.

### 3. Apply a Theme (Optional)

In Power BI Desktop, go to **View → Themes → Browse for themes** and select either:
- `Highgate_PowerBI_Theme.json` — dark navy and gold palette
- `IHG_PowerBI_Theme.json` — IHG brand colours

## Data Sources

The report is designed to work with standard Opera PMS exports:

- **Daily Flash** — daily revenue and occupancy summary (Excel)
- **Reservation Activity** — history and forecast (XML)
- **Profile Production Statistics** — source/agent production (XML)
- **Market Segment Summary** — segment breakdown (Excel)
- **Package Forecast** — 8-day package revenue forecast (Excel)
- **Monthly Statistics** — room status summary (Excel)

## Requirements

- Power BI Desktop (latest recommended)
- Oracle Opera PMS (reports available as standard exports)
