# SteamDB Upcoming Games Tracker

[English](README.md) | [简体中文](README.zh-CN.md)

A Windows-based Node.js tracker that collects upcoming Steam releases from SteamDB, retains follower history, enriches qualifying games with Steam Store metadata, and generates a formatted Excel workbook for release analysis. The workbook can be produced in English, Simplified Chinese, or both at once.

## What The Tracker Does

1. Computes every ISO week that intersects the **current calendar month and the next calendar month** from the local crawl date.
2. Opens SteamDB's upcoming-release pages for each of those weeks.
3. Sorts each SteamDB page by descending follower count.
4. Stores every successfully scraped game in `steamdb_master_data.json`.
5. Fetches Steam Store metadata only for games meeting the configured follower threshold.
6. Builds the localized Excel workbook (or workbooks) using only threshold-qualified games.
7. Groups games into Excel tabs based on **their release month**, not the month they were crawled.
8. Assigns a dynamic `Current Outlook` from P0 to P3 using the latest follower count.

This separation keeps the Excel analysis focused without discarding lower-follower source data.

## Calendar-Month Crawl

The crawl no longer uses a fixed six-week window. Instead it derives the crawl window from the calendar:

- The tracker takes the **local crawl date**.
- It selects the **current calendar month** and the **next calendar month**.
- It collects every **ISO week** that intersects any day of those two months.
- Because ISO weeks span month boundaries, the set can include weeks that spill into the previous or following month. SteamDB pages are scraped for all of those weeks, and every row is stored.
- A crawl is **complete** only when every expected weekly table loads successfully. If any expected week fails, the partial crawl is **not committed** to the master data and the previous history is preserved.

Example: a crawl on 30 Sep 2026 covers September and October 2026. September contributes weeks 36–40 and October contributes weeks 40–44, so the expected set is `2026W36` through `2026W44` (9 weeks).

## Release-Month Visibility Model

The tracker uses the **latest complete crawl date** to decide which release months are live, frozen, or hidden.

- **Active months** — the release month of the latest complete crawl plus the following release month. Only these two months are live. Both tabs are always generated, even when a month has zero qualified games.
- **Hidden future months** — any release month later than the next active month. Their games stay in `steamdb_master_data.json` and become visible automatically once that month becomes active.
- **Archived months** — qualified release months earlier than the active current month. Their tabs are frozen forever in `steamdb_release_archives.json` and are regenerated from that immutable snapshot, never from the live master. Past months that are **not** archived are never shown.
- **Unscheduled** — qualified games without a specific release month, refreshed from the latest complete crawl.

Archives represent only observations that existed **before the active current month**. When the crawl date moves to a new month, prior months are frozen from the pre-crawl state and their histories are trimmed at the current month's start, so a spillover row scraped in the new month can never fill or alter an already completed month. A game first seen only in the new month's spillover is never added to the previous month's archive.

Behavior examples:

- With the current local master (latest complete crawl `2026-09-30`), a rebuild shows the active **Sep 2026** and **Oct 2026** tabs (Sep empty if no September games are stored) and hides **Nov 2026**.
- A future complete crawl in October 2026 freezes **Sep 2026** at its latest September observation, then makes **Oct 2026** and **Nov 2026** active. October spillover rows for September releases are stored in the master but do not touch the frozen September tab.
- If the very first crawl happens in October and only sees September spillover with no pre-October history, no September tab or September archive is created.
- Once a month exists in the archive, it never changes on later crawls or threshold changes. Missing past months are added automatically the next time a complete crawl or rebuild establishes that they should exist.

The Summary sheet and the Qualified Upcoming sheet include **only** qualified games from the **latest complete crawl** that are assigned to the two active release months, so stale games absent from the newest crawl do not linger there. The active month tabs may still retain their stored qualified games.

## Output Languages

Set the workbook language in `.env`:

```env
OUTPUT_LANGUAGE=both
```

| Value | Output file(s) |
|---|---|
| `en` | `steamdb_upcoming_tracker.xlsx` |
| `zh-CN` | `steamdb_upcoming_tracker_zh-CN.xlsx` |
| `both` (default) | both files above |

When `both` is selected, the two workbooks are generated from the **same in-memory model** in a single pass — no extra crawling and no extra Steam API calls.

Only fixed workbook text is translated: sheet names, column headers, statuses, Summary titles/cards/tables, Current Outlook labels, the fixed `Yes`/`No` values, the fixed `N/A` missing-data marker used for publisher/developer/genres/categories/platforms, and link tooltips. Crawled values such as real game names, publishers, developers, genres, and categories are never translated. Chinese uses **Microsoft YaHei** and the Excel-safe `yyyy-mm-dd` date format; English uses **Aptos / Aptos Display**.

If a test or script passes an explicit `excelFile`, a single English workbook is written to that path unless a `language` option is also given.

## Visibility Rules

The latest follower count controls whether a game appears in Excel:

- A game is visible when `Latest Followers >= FOCUSED_THRESHOLD`.
- A game below the threshold is hidden from every Excel sheet and Summary calculation.
- Hidden-game counts are not displayed in the workbook.
- All games still remain in the JSON master database.
- When a game crosses the threshold, it appears automatically with its complete earlier follower history.
- If a game's latest count falls below the threshold, it is hidden again without losing its stored history.

The default threshold is 1,000 followers and can be changed in `.env`, but the Current Outlook framework enforces a hard floor of 1,000.

## Release-Month Rules

Excel month tabs represent when games are scheduled to launch:

- A game releasing in September appears in `Sep 2026` (or `2026年9月`), even if it was crawled in August or September.
- A game with a month-only date such as `Oct 2026` appears in `Oct 2026`, while the cell remains text.
- Exact dates are normalized to UTC midnight before being written as real Excel dates, preventing one-day shifts in GMT+8 and other non-UTC time zones.
- If SteamDB moves a release from September to October, the row moves to `Oct 2026` on the next complete crawl while frozen archive tabs stay untouched.
- Coming Soon, year-only, quarter-only, and other dates without a specific month go to `Unscheduled`.
- Released games remain in their release-month tab because historical games are retained in the master database.

## Excel Workbook

The generated workbook is organized as follows:

```text
Summary | Qualified Upcoming | <archived months> | <active months> | Unscheduled
```

Archived month tabs are ordered chronologically before the active months. The two active month tabs are always present, even if empty. Past months appear only from the frozen archive; unarchived past months and future months are never rendered as tabs.

### Summary

The dashboard contains:

- Last Complete Crawl
- Qualified Upcoming
- New This Crawl
- Releases in 30 Days
- Current Release Month
- Next Release Month
- Unscheduled
- Follower Threshold
- Release Month Overview
- Top 10 Movers Since Previous Crawl
- Next 25 Releases Within 30 Days

Summary game metrics and tables use only active-month games meeting the current threshold. `Days to Release` is intentionally omitted from the detailed tracker sheets and appears only in the Summary's 30-day release-planning table.

### Qualified Upcoming

Contains threshold-qualified games assigned to the two active release months and seen in the latest complete crawl. Archived and future months do not appear here.

### Release-Month Tabs

Active month tabs are always present (even when empty) and contain all currently qualified games stored for that month, which may be more than the latest crawl returned. Archived month tabs are frozen snapshots rendered from `steamdb_release_archives.json`, built only from observations before the active current month, and remain stable across later crawls and threshold changes. Past months that are not archived are not displayed.

### Unscheduled

Contains qualified games whose latest release information cannot be assigned to a specific month, refreshed from the latest complete crawl.

## Tracker Columns

The requested core fields always appear first. API and calculated fields follow, then one follower column for every captured crawl date.

| Column (en / zh-CN) | Source | Description |
|---|---|---|
| Release Date / 发售日期 | SteamDB | Exact Excel date or month-only/unscheduled text |
| Game Title / 游戏名称 | SteamDB | Game name linked to its SteamDB page |
| Followers / 关注人数 | SteamDB | Latest captured follower count |
| Publisher / 发行商 | Steam Store API | Publisher list for qualifying games |
| Developer / 开发商 | Steam Store API | Developer list for qualifying games |
| Current Outlook / 当前评级 | Calculated | P0–P3 outlook based on the latest follower count |
| Is Free / 是否免费 | Steam Store API | Whether the game is free to play |
| Genres / 游戏类型 | Steam Store API | Steam genre list |
| Categories / 功能标签 | Steam Store API | Features such as single-player, co-op, or controller support |
| Platforms / 支持平台 | Steam Store API | Windows, macOS, and/or Linux |
| AppID | SteamDB | Steam AppID linked to the Steam Store page |
| Status / 状态 | Calculated | Upcoming, Releases Today, Released, Released (Month), Launch Month, Upcoming (Month), or Unscheduled |
| Previous Crawl Followers / 上次抓取关注人数 | Calculated | Previous available observation for that game |
| Change Since Previous Crawl / 较上次抓取变化 | Calculated | Latest followers minus the previous observation; this is not a 7-day metric |
| First Seen / 首次发现 | Calculated | First date the tracker captured the game |
| Last Updated / 最近更新 | Calculated | Most recent complete crawl containing the game |
| `YYYY-MM-DD` columns | History | Follower count captured on each crawl date |

## Current Outlook Framework

The tracker assigns one concise outlook using four mutually exclusive follower ranges. It does not generate wishlist, unit-sales, or revenue estimates.

| Priority | Exact Follower Range | Current Outlook (en) | Current Outlook (zh-CN) |
|---|---:|---|---|
| P0 | 30,000+ | AAA | AAA |
| P1 | 10,000–29,999 | AA | AA |
| P2 | 3,000–9,999 | Indie | 独立游戏 |
| P3 | 1,000–2,999 | Barely Viable | 勉强可行 |

The displayed value combines both parts, such as `P1 - AA`. Current Outlook is recalculated whenever the workbook is built, so it can improve as followers grow closer to launch. It is a prioritization aid rather than a guaranteed commercial outcome.

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
- Priority-colored Current Outlook cells
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
OUTPUT_LANGUAGE=both
```

## Running A Crawl

The recommended workflow is:

1. Close `steamdb_upcoming_tracker.xlsx` and `steamdb_upcoming_tracker_zh-CN.xlsx` if either is open.
2. Double-click `start_tracker.bat`.
3. Microsoft Edge opens with remote debugging enabled.
4. Complete the Cloudflare verification if it appears.
5. Wait until the SteamDB table is visible.
6. Return to the command window and press any key.
7. The tracker crawls the two calendar months, enriches qualifying games, saves JSON history, updates archives, and rebuilds the workbook(s).

To attach manually to an Edge instance already running on debugging port `9222`:

```bash
npm run remote
```

## Commands

| Command | Description |
|---|---|
| `npm start` | Launch with the configured Playwright browser profile |
| `npm run remote` | Attach to the manually verified Edge session on port 9222 |
| `npm run rebuild` | Rebuild the workbook(s) and archives from local JSON without crawling SteamDB or calling the Steam API |
| `npm run check` | Check JavaScript syntax |
| `npm test` | Run calendar, archive, localization, threshold, and workbook tests |

## Configuration

### Follower Threshold

Set the threshold in `.env`:

```env
FOCUSED_THRESHOLD=1000
```

The setting controls both:

- Which games are visible anywhere in Excel
- Which games receive Steam Store API enrichment

The comparison is inclusive, so a game with exactly 1,000 followers qualifies. The P3 framework enforces a hard floor of 1,000: higher settings narrow the workbook further, while lower, missing, invalid, zero, or negative values resolve to 1,000.

After changing the threshold, close Excel and run the following command to refresh the active workbook tabs without another crawl:

```bash
npm run rebuild
```

Frozen archive tabs do not change when the threshold changes.

### Language

```env
OUTPUT_LANGUAGE=both
```

Accepted values are `en`, `zh-CN`, and `both`. See [Output Languages](#output-languages).

### Crawl Window

The crawl window is derived from the calendar and is not configured manually. The tracker always targets the current calendar month and the next calendar month and computes the intersecting ISO weeks at runtime.

## Data And Update Behavior

### Complete Crawls

A crawl is committed only when every expected weekly SteamDB table loads successfully. Past months are frozen from the pre-crawl state first, then the new crawl is merged into the master JSON, and finally the Excel workbook(s) are rebuilt.

### Incomplete Crawls

If any expected weekly table fails to load:

- The partial crawl is not written into the master history.
- Existing follower observations are not overwritten.
- The workbook is regenerated only from the last complete stored data.
- The latest complete crawl date does not advance.
- A `debug_cloudflare_YYYYWNN.png` screenshot is created for troubleshooting.

### Release Archives

`steamdb_release_archives.json` stores immutable, versioned snapshots of past qualified release months. Each archive entry keeps the full game data, follower history, computed status, and outlook needed to regenerate that tab exactly. Snapshots are captured from the pre-crawl state and their histories are trimmed at the start of the active current month, so newer spillover observations can never contaminate them. Archived entries are never rewritten.

### Steam Metadata Cache

Steam metadata is saved in the master JSON and reused on later runs. Requests are sent conservatively to avoid Steam API rate limits, so the first enrichment of many qualifying games can take several minutes. Failed or unavailable records are retried later.

### Date And Time-Zone Handling

Crawl history keys use the computer's local calendar date. Before a date is written to Excel, the tracker creates it at UTC midnight so ExcelJS cannot shift it into the previous day during serialization. For example, an underlying `01 Oct 2026` release remains `01 Oct 2026` in both the October tab and the visible Release Date cell when the tracker runs in GMT+8.

After upgrading from an earlier version that displayed dates one day early, close Excel and run `npm run rebuild` to regenerate all visible dates from the unchanged JSON history.

## Local Files

| File | Purpose | Committed to Git? |
|---|---|---|
| `steamdb_tracker.js` | Scraper, persistence, archives, localization, and Excel generation | Yes |
| `start_tracker.bat` | Recommended Windows/Edge launch workflow | Yes |
| `.env.example` | Shareable configuration template | Yes |
| `README.zh-CN.md` | Simplified Chinese documentation | Yes |
| `steamdb_master_data.json` | Permanent local game history and metadata cache | No |
| `steamdb_release_archives.json` | Immutable frozen past-month snapshots | No |
| `steamdb_upcoming_tracker.xlsx` | Generated English workbook | No |
| `steamdb_upcoming_tracker_zh-CN.xlsx` | Generated Simplified Chinese workbook | No |
| `edge_profile/` | Local Edge profile used for Cloudflare verification | No |
| `debug_cloudflare_*.png` | Failure screenshots | No |

Versions before the release-month redesign may have created `steamdb_monthly_snapshots.json`. The current tracker does not use that file; `steamdb_master_data.json` is the source of truth.

## Troubleshooting

### Excel file is locked

If the tracker says to close a workbook, close `steamdb_upcoming_tracker.xlsx` and `steamdb_upcoming_tracker_zh-CN.xlsx` in Excel and rerun the command. Excel prevents the script from replacing an open workbook.

### Excel reports a problem with the Chinese workbook

An earlier Chinese workbook format used localized date literals that some Excel versions attempted to repair in `styles.xml`. The current version uses the Excel-safe `yyyy-mm-dd` format. Close the workbook and run `npm run rebuild` to replace the older file.

### Cloudflare or a weekly table fails

Run `start_tracker.bat`, complete the browser verification, wait for the table to appear, and then press a key in the command window. An incomplete crawl is deliberately rejected to protect the history.

### Publisher or developer is blank

The game may be below the threshold, unavailable through the Steam Store API, or waiting for a retry after a failed API response.

### A game is missing from Excel

Check its latest follower count in `steamdb_master_data.json`. A game appears only when its latest value meets `FOCUSED_THRESHOLD`. Its history is retained even while hidden.

### A future release month is missing

Months later than the next active month are intentionally hidden. They return automatically once the calendar advances and they become active.

### A date still appears one day early

Close every open copy of the workbooks and run `npm run rebuild`. Existing workbooks are not changed until they can be successfully regenerated.

### Only one workbook is produced

Check `OUTPUT_LANGUAGE` in `.env`. Use `both` to generate the English and Simplified Chinese files together.

## Sharing The Repository

Another person can clone the repository, run `npm install`, copy `.env.example` to `.env`, and follow the batch-file workflow. Generated history, archives, and Excel files are intentionally local, so every user starts with an independent tracker database.

## License

ISC

## Disclaimer

This project is for personal and educational tracking. Run it no more than once per day and respect SteamDB's server load and terms.
