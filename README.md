# SteamDB Upcoming Games Tracker

A Windows-based Node.js tracker that collects upcoming Steam releases from SteamDB, retains follower history, enriches qualifying games with Steam Store metadata, and generates a formatted Excel workbook for release analysis.

## What The Tracker Does

1. Opens SteamDB's upcoming-release pages for the current week and the next five weeks.
2. Sorts each SteamDB page by descending follower count.
3. Stores every successfully scraped game in `steamdb_master_data.json`.
4. Fetches Steam Store metadata only for games meeting the configured follower threshold.
5. Builds `steamdb_upcoming_tracker.xlsx` using only threshold-qualified games.
6. Groups games into Excel tabs based on **their release month**, not the month they were crawled.
7. Interprets each visible follower count using a Tier 2–5 commercial-significance framework.

This separation keeps the Excel analysis focused without discarding lower-follower source data.

## Visibility Rules

The latest follower count controls whether a game appears in Excel:

- A game is visible when `Latest Followers >= FOCUSED_THRESHOLD`.
- A game below the threshold is hidden from every Excel sheet and Summary calculation.
- Hidden-game counts are not displayed in the workbook.
- All games still remain in the JSON master database.
- When a game crosses the threshold, it appears automatically with its complete earlier follower history.
- If a game's latest count falls below the threshold, it is hidden again without losing its stored history.

The default threshold is 1,000 followers and can be changed in `.env`.

## Release-Month Rules

Excel month tabs represent when games are scheduled to launch:

- A game releasing in September appears in `Sep YYYY`, even if it was crawled in August or September.
- A game with a month-only date such as `Oct 2026` appears in `Oct 2026`, while the cell remains text.
- Exact dates are normalized to UTC midnight before being written as real Excel dates, preventing one-day shifts in GMT+8 and other non-UTC time zones.
- If SteamDB moves a release from September to October, the row moves to `Oct YYYY` on the next complete crawl.
- Moving a game does not remove its previous follower observations.
- Coming Soon, year-only, quarter-only, and other dates without a specific month go to `Unscheduled`.
- Released games remain in their release-month tab because historical games are retained in the master database.

## Excel Workbook

The generated workbook is organized as follows:

```text
Summary | Qualified Upcoming | Sep 2026 | Oct 2026 | ... | Unscheduled
```

### Summary

The dashboard contains:

- Last Complete Crawl
- Qualified Upcoming
- New This Crawl
- Releases in 30 Days
- Release Months
- Unscheduled
- Total Qualified Tracked
- Follower Threshold
- Release Month Overview
- Top 10 Movers Since Previous Crawl
- Next 25 Releases Within 30 Days

All Summary metrics and tables use only games meeting the current threshold.
`Days to Release` is intentionally omitted from the detailed tracker sheets and appears only in the Summary's 30-day release-planning table.

### Qualified Upcoming

Contains threshold-qualified games found in the latest complete six-week crawl window. Older games remain available in their release-month tabs but do not stay in this live sheet unless they are present in the latest crawl.

### Release-Month Tabs

Each tab contains all currently qualified games assigned to that release month. Tabs are generated chronologically from the permanent master data.

### Unscheduled

Contains qualified games whose latest release information cannot be assigned to a specific month.

## Tracker Columns

The requested core fields always appear first. API and calculated fields follow, then one follower column for every captured crawl date.

| Column | Source | Description |
|---|---|---|
| Release Date | SteamDB | Exact Excel date or month-only/unscheduled text |
| Game Title | SteamDB | Game name linked to its SteamDB page |
| Followers | SteamDB | Latest captured follower count |
| Publisher | Steam Store API | Publisher list for qualifying games |
| Developer | Steam Store API | Developer list for qualifying games |
| Follower Tier | Calculated | Tier 2, 3, 4, or 5 based on the latest follower count |
| Commercial Standing | Calculated | Concise interpretation associated with the follower tier |
| Is Free | Steam Store API | Whether the game is free to play |
| Genres | Steam Store API | Steam genre list |
| Categories | Steam Store API | Features such as single-player, co-op, or controller support |
| Platforms | Steam Store API | Windows, macOS, and/or Linux |
| AppID | SteamDB | Steam AppID linked to the Steam Store page |
| Status | Calculated | Upcoming, Releases Today, Released, Launch Month, or Unscheduled |
| Previous Crawl Followers | Calculated | Previous available observation for that game |
| Change Since Previous Crawl | Calculated | Latest followers minus the previous observation; this is not a 7-day metric |
| First Seen | Calculated | First date the tracker captured the game |
| Last Updated | Calculated | Most recent complete crawl containing the game |
| `YYYY-MM-DD` columns | History | Follower count captured on each crawl date |

## Follower Interpretation Framework

The tracker interprets the latest SteamDB follower count using four mutually exclusive tiers. It does not generate wishlist, unit-sales, or revenue estimates.

| Tier | Exact Follower Range | Commercial Standing | Interpretation |
|---|---:|---|---|
| Tier 2 | 1,000–2,999 | Barely Viable / Solo Indie Floor | Baseline level for a title to enter focused commercial monitoring. |
| Tier 3 | 3,000–9,999 | Commercial Hit / Sustainable Indie | Stronger organic-interest signal associated with a potentially sustainable indie release. |
| Tier 4 | 10,000–29,999 | AA / Mid-Tier Blockbuster | High-confidence signal of substantial pre-launch demand and broad market visibility. |
| Tier 5 | 30,000+ | Major Commercial Hit / Megahit | Strongest tracked signal, indicating exceptional organic attention before launch. |

Tier assignment is recalculated from the latest follower count whenever the workbook is built. It is a prioritization aid rather than a guaranteed commercial outcome.

## Excel Quality-Of-Life Features

- Filter controls on every tracker sheet
- Frozen header row and first two columns
- Release-date-first sorting, then follower count
- Wrapped game titles and metadata text
- Purposeful column widths and number/date formats
- Timezone-safe crawl and release dates
- SteamDB links on game titles
- Steam Store links on AppIDs
- Follower color scale for relative popularity
- Tier-colored Follower Tier and Commercial Standing cells
- Green/red formatting for positive and negative follower movement
- Highlighting for releases within 14 days
- Styled Summary cards and tables

## Installation

### Requirements

- Windows with Microsoft Edge
- Node.js 18 or newer
- npm

### Set Up

```powershell
git clone https://github.com/laimunchoong97/SteamDB-Tracker.git
Set-Location SteamDB-Tracker
npm install
Copy-Item .env.example .env
```

Default `.env` configuration:

```env
BROWSER_CHANNEL=msedge
BROWSER_EXECUTABLE_PATH=
FOCUSED_THRESHOLD=1000
```

## Running A Crawl

The recommended workflow is:

1. Close `steamdb_upcoming_tracker.xlsx` if it is open.
2. Double-click `start_tracker.bat`.
3. Microsoft Edge opens with remote debugging enabled.
4. Complete the Cloudflare verification if it appears.
5. Wait until the SteamDB table is visible.
6. Return to the command window and press any key.
7. The tracker crawls all configured weeks, enriches qualifying games, saves JSON history, and rebuilds Excel.

To attach manually to an Edge instance already running on debugging port `9222`:

```bash
npm run remote
```

## Commands

| Command | Description |
|---|---|
| `npm start` | Launch with the configured Playwright browser profile |
| `npm run remote` | Attach to the manually verified Edge session on port 9222 |
| `npm run rebuild` | Rebuild Excel from local JSON without crawling SteamDB or calling the Steam API |
| `npm run check` | Check JavaScript syntax |
| `npm test` | Run threshold, tier-boundary, date, and workbook tests |

## Configuration

### Follower Threshold

Set the threshold in `.env`:

```env
FOCUSED_THRESHOLD=1000
```

The setting controls both:

- Which games are visible anywhere in Excel
- Which games receive Steam Store API enrichment

The comparison is inclusive, so a game with exactly 1,000 followers qualifies. The Tier 2 framework enforces a hard floor of 1,000: higher settings narrow the workbook further, while lower, missing, invalid, zero, or negative values resolve to 1,000.

After changing the threshold, close Excel and run the following command to refresh the workbook without another crawl:

```bash
npm run rebuild
```

### Crawl Window

Change the number of release weeks near the top of `steamdb_tracker.js`:

```javascript
const WEEKS_TO_SCRAPE = 6;
```

## Data And Update Behavior

### Complete Crawls

A crawl is committed only when every configured weekly SteamDB table loads successfully. The master JSON is saved first, then the Excel workbook is rebuilt from that history.

### Incomplete Crawls

If any weekly table fails to load:

- The partial crawl is not written into the master history.
- Existing follower observations are not overwritten.
- The workbook is regenerated only from the last complete stored data.
- A `debug_cloudflare_YYYYWNN.png` screenshot is created for troubleshooting.

### Steam Metadata Cache

Steam metadata is saved in the master JSON and reused on later runs. Requests are sent conservatively to avoid Steam API rate limits, so the first enrichment of many qualifying games can take several minutes. Failed or unavailable records are retried later.

### Date And Time-Zone Handling

Crawl history keys use the computer's local calendar date. Before a date is written to Excel, the tracker creates it at UTC midnight so ExcelJS cannot shift it into the previous day during serialization. For example, an underlying `01 Oct 2026` release remains `01 Oct 2026` in both the October tab and the visible Release Date cell when the tracker runs in GMT+8.

After upgrading from an earlier version that displayed dates one day early, close Excel and run `npm run rebuild` to regenerate all visible dates from the unchanged JSON history.

## Local Files

| File | Purpose | Committed to Git? |
|---|---|---|
| `steamdb_tracker.js` | Scraper, persistence, analysis, and Excel generation | Yes |
| `start_tracker.bat` | Recommended Windows/Edge launch workflow | Yes |
| `.env.example` | Shareable configuration template | Yes |
| `steamdb_master_data.json` | Permanent local game history and metadata cache | No |
| `steamdb_upcoming_tracker.xlsx` | Generated analysis workbook | No |
| `edge_profile/` | Local Edge profile used for Cloudflare verification | No |
| `debug_cloudflare_*.png` | Failure screenshots | No |

Versions before the release-month redesign may have created `steamdb_monthly_snapshots.json`. The current tracker does not use that file; `steamdb_master_data.json` is the source of truth.

## Troubleshooting

### Excel file is locked

If the tracker says to close `steamdb_upcoming_tracker.xlsx`, close the workbook in Excel and rerun the command. Excel prevents the script from replacing an open workbook.

### Cloudflare or a weekly table fails

Run `start_tracker.bat`, complete the browser verification, wait for the table to appear, and then press a key in the command window. An incomplete crawl is deliberately rejected to protect the history.

### Publisher or developer is blank

The game may be below the threshold, unavailable through the Steam Store API, or waiting for a retry after a failed API response.

### A game is missing from Excel

Check its latest follower count in `steamdb_master_data.json`. A game appears only when its latest value meets `FOCUSED_THRESHOLD`. Its history is retained even while hidden.

### A date still appears one day early

Close every open copy of `steamdb_upcoming_tracker.xlsx` and run `npm run rebuild`. Existing workbooks are not changed until they can be successfully regenerated.

## Sharing The Repository

Another person can clone the repository, run `npm install`, copy `.env.example` to `.env`, and follow the batch-file workflow. Generated history and Excel files are intentionally local, so every user starts with an independent tracker database.

## License

ISC

## Disclaimer

This project is for personal and educational tracking. Run it no more than once per day and respect SteamDB's server load and terms.
