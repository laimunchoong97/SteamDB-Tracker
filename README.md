# SteamDB Upcoming Games Tracker

A Node.js tracker that collects upcoming Steam releases from SteamDB, enriches qualifying games with Steam Store metadata, and maintains an Excel workbook with permanent month-by-month snapshots.

## Features

- Scrapes SteamDB's upcoming releases sorted by follower count.
- Covers the current week and the next five weeks.
- Uses a manually verified Edge session to get through Cloudflare Turnstile.
- Enriches games at or above the follower threshold with publisher, developer, free-to-play status, genres, categories, and supported platforms.
- Stores daily follower observations in a local master JSON database.
- Creates a live `Current Tracker` sheet and a permanent sheet for every month.
- Preserves the last complete crawl of each month instead of deleting prior-month data.
- Keeps month-only release dates as text and writes exact release dates as sortable Excel dates.

## Requirements

- Windows with Microsoft Edge
- Node.js 18 or newer
- npm

## Installation

```bash
git clone https://github.com/laimunchoong97/SteamDB-Tracker.git
cd SteamDB-Tracker
npm install
```

Copy `.env.example` to `.env`. The default configuration uses Microsoft Edge:

```env
BROWSER_CHANNEL=msedge
BROWSER_EXECUTABLE_PATH=
```

## Run The Tracker

1. Close `steamdb_upcoming_tracker.xlsx` if it is open.
2. Double-click `start_tracker.bat`.
3. Complete the Cloudflare check in the Edge window if one appears.
4. Wait until the SteamDB table is visible, then return to the command window and press any key.
5. The batch file runs the tracker and creates or updates the Excel workbook.

The equivalent terminal command, after Edge has been started on debugging port `9222`, is:

```bash
npm run remote
```

## Monthly Workbook Behavior

The workbook contains sheets such as:

```text
Current Tracker | Oct 2026 | Sep 2026
```

- During a month, each complete crawl replaces that month's sheet with the newest successful snapshot.
- When a new month begins, the previous month's final snapshot is left unchanged and a new monthly sheet is created.
- Incomplete crawls do not replace a valid monthly snapshot.
- Monthly sheets contain only games meeting the configured follower threshold at the time of that snapshot.
- The newest month appears first after `Current Tracker`.

Every tracker sheet uses this column order:

```text
Release Date, Game Title, Followers, Publisher, Developer,
Is Free, Genres, Categories, Platforms, AppID, Snapshot Date
```

Game titles link to SteamDB, and AppIDs link to the corresponding Steam Store pages.

## Local Data Files

The tracker creates these files locally:

- `steamdb_master_data.json`: daily observations and cached Steam metadata.
- `steamdb_monthly_snapshots.json`: frozen monthly snapshots used to rebuild the workbook.
- `steamdb_upcoming_tracker.xlsx`: the generated Excel tracker.
- `debug_cloudflare_*.png`: troubleshooting screenshots created only when a weekly table fails to load.

These generated files are intentionally ignored by Git. Someone cloning the repository starts with a clean tracker and creates their own local history.

To regenerate Excel from the existing local JSON data without scraping SteamDB:

```bash
npm run rebuild
```

## Configuration

Edit the constants near the top of `steamdb_tracker.js`:

```javascript
const FOCUSED_THRESHOLD = 1000;
const WEEKS_TO_SCRAPE = 6;
```

The Steam Store API is queried only for qualifying games whose metadata is not already cached.

## Disclaimer

This project is for personal and educational tracking. Run it no more than once per day and respect SteamDB's server load and terms.
