# SteamDB Upcoming Games Tracker

An automated scraper built with Node.js and Playwright that bypasses Cloudflare protections to track the follower counts of upcoming Steam games week over week, cross-referencing data with the official Steam Store API.

## Features
- **Cloudflare Bypass:** Includes a `start_tracker.bat` workflow to manually solve Turnstile captchas in a real Edge browser, then connects Playwright via CDP to completely bypass bot protection.
- **Automated Lookahead:** Scrapes the current week + the next 5 weeks of top upcoming games.
- **Steam API Integration:** Automatically fetches Developer, Publisher, Platform, Category, and Genre metadata directly from Steam for games that reach a certain follower threshold.
- **Intelligent Excel Generation:** 
  - Outputs a 2-tab Excel file (`steamdb_upcoming_tracker.xlsx`).
  - Automatically calculates 7-day rolling follower gains.
  - Correctly formats dates for Excel sorting.

## Prerequisites
- **Node.js** (v18 or higher)
- **Windows** (Because it uses a `.bat` file to open MS Edge)

## Installation

1. Clone or download this repository.
2. Open your terminal/command prompt and navigate to the folder.
3. Install the dependencies:
   ```bash
   npm install
   ```

## Setup Environment Variables
1. Copy the `.env.example` file and rename it to `.env`:
   ```bash
   cp .env.example .env
   ```
2. Open the `.env` file and ensure `BROWSER_CHANNEL=msedge` (or provide an absolute path to a specific executable if needed).

## Usage

Because SteamDB uses extremely aggressive Cloudflare Turnstile protection, standard headless scrapers will be blocked. You must use the following workflow:

1. Double-click the **`start_tracker.bat`** file.
2. A new Microsoft Edge window will open to `https://steamdb.info/upcoming/`. 
3. **If you see a Cloudflare "Verify you are human" checkbox, click it manually.**
4. Once the actual SteamDB page loads, leave the browser open.
5. Open a terminal in the project folder and run:
   ```bash
   node steamdb_tracker.js --remote
   ```
6. The script will attach to the Edge window you just verified, scrape the data, fetch Steam API metadata, and generate `steamdb_upcoming_tracker.xlsx`.

## Configuration
You can change the follower threshold for when games trigger the Steam API metadata fetch.
Open `steamdb_tracker.js` and edit this line at the top of the file:
```javascript
const FOCUSED_THRESHOLD = 1000;
```

## Disclaimer
This tool is for educational and personal data tracking purposes. Please respect SteamDB's server load by running this script a maximum of once per day.
