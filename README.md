# Staging Log Sheet — Google Sheets Auto-Sync & API Endpoints

Live dashboard and JSON/CSV endpoints for **Log Sheet for Staging Dispensed Material and Intermediate Product** (Renata PLC, Rajendrapur Potent Product Facility).

- **Dashboard & Endpoint Explorer**: https://arif0850.github.io/staging-log-sheet/
- **Apps Script**: [`apps-script/Code.gs`](apps-script/Code.gs)

## Available Static Endpoints

| Endpoint | Description |
| --- | --- |
| `data/all.json` | Complete dataset (meta + room & product summary + all entries) |
| `data/not-exited.json` | All Not Exited entries (Rooms 115, 119 & 126) |
| `data/rooms.json` | Entries grouped by Staging Room (`115`, `119` & `126`) |
| `data/room-115.json` | Only Staging Room 115 Not Exited entries |
| `data/room-119.json` | Only Staging Room 119 Not Exited entries |
| `data/room-126.json` | Only Staging Room 126 Not Exited entries |
| `data/summary.json` | Room-wise & product-wise summary + validity breakdown |
| `data/expired.json` | Expired entries only |
| `data/valid.json` | Valid entries only |
| `data/na.json` | N/A validity entries only |
| `data/products.json` | Entries grouped by product |
| `data/meta.json` | Sync metadata & endpoint directory |
| `data/not-exited.csv` | CSV export of Not Exited sheet |
| `data/summary.csv` | CSV export of Summary sheet |

## Google Apps Script Setup

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste `apps-script/Code.gs` and save.
3. Reload the Google Sheet and use **🔄 GitHub Sync → Configure GitHub Settings** to save your GitHub PAT (`repo` scope).
4. Click **🔄 GitHub Sync → Enable Auto-Sync Triggers** so every new entry or edit automatically commits to `main`.
