# SteamDB Upcoming Games Tracker

A Node.js tracker that collects upcoming Steam releases from SteamDB, retains follower history, enriches qualifying games with Steam Store metadata, and builds a formatted Excel release-planning workbook.

## Features

- Scrapes SteamDB's upcoming releases sorted by follower count.
- Covers the current week and the next five weeks.
- Uses a manually verified Edge session to get through Cloudflare Turnstile.
- Keeps every scraped game in the JSON database while showing only threshold-qualified games in Excel.
- Groups visible games by their announced **release month**, not the month in which they were crawled.
- Moves a game to its new release-month sheet when SteamDB reports a changed release date.
- Places Coming Soon, year-only, quarter-only, and other non-month-specific releases in `Unscheduled`.
- Retains dated follower history across crawls.
- Uses the same configurable follower threshold for Excel visibility and Steam Store API enrichment.
- Produces a styled Excel dashboard with filters, frozen panes, conditional formatting, useful column widths, and clickable links.

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
FOCUSED_THRESHOLD=1000
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

## Workbook Layout

The workbook is rebuilt from the permanent JSON history after every complete crawl:

```text
Summary | Qualified Upcoming | Sep 2026 | Oct 2026 | ... | Unscheduled
```

- `Summary` contains headline metrics, release-month totals, top movers since the previous crawl, and releases due within 30 days.
- `Qualified Upcoming` contains threshold-qualified games seen in the latest complete crawl.
- Month sheets contain threshold-qualified games currently scheduled to launch in that month.
- `Unscheduled` contains threshold-qualified games that cannot be assigned to a specific release month.
- An incomplete crawl does not modify the master history.

Games below the threshold are absent from every Excel analysis sheet and Summary calculation. They remain in `steamdb_master_data.json`, so a game automatically appears with its complete earlier follower history when its latest count reaches the threshold. Qualifying games are enriched with Publisher, Developer, Is Free, Genres, Categories, and Platforms.

Each tracker sheet begins with this column order:

```text
Release Date, Game Title, Followers, Publisher, Developer,
Is Free, Genres, Categories, Platforms, AppID,
Status, Days to Release, Previous Crawl Followers,
Change Since Previous Crawl, First Seen, Last Updated
```

Dated follower columns are appended after these fields. Exact release dates are sortable Excel dates, while month-only values such as `Oct 2026` remain text. Game titles link to SteamDB and AppIDs link to Steam.

## Release-Month Behavior

- A qualifying game with a September release date appears in `Sep YYYY`, even if it was first crawled in August.
- Subsequent crawls update its latest follower count and append dated follower history in the same sheet.
- If its release changes from September to October, it moves to `Oct YYYY` while keeping its complete follower history.
- Released games remain in their release-month sheet because the JSON master database is never replaced by only the latest crawl.

## Local Data Files

The tracker creates these files locally:

- `steamdb_master_data.json`: permanent follower observations, latest release information, and cached Steam metadata.
- `steamdb_upcoming_tracker.xlsx`: the generated Excel workbook.
- `debug_cloudflare_*.png`: troubleshooting screenshots created only when a weekly table fails to load.

Generated data is intentionally ignored by Git. Someone cloning the repository starts with a clean tracker and creates their own local history.

To rebuild Excel from the existing local JSON data without crawling SteamDB or calling the Steam API:

```bash
npm run rebuild
```

## Configuration

Set the Excel visibility and Steam metadata threshold in `.env`:

```env
FOCUSED_THRESHOLD=1000
```

The value defaults to `1000` when the setting is missing or invalid. To change the number of release weeks crawled, edit the constant near the top of `steamdb_tracker.js`:

```javascript
const WEEKS_TO_SCRAPE = 6;
```

The Steam Store API is called at a conservative rate and only for qualifying games whose metadata is not already cached. The first enrichment pass can take several minutes; subsequent runs reuse the cache.

## Disclaimer

This project is for personal and educational tracking. Run it no more than once per day and respect SteamDB's server load and terms.
