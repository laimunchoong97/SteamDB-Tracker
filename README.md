# SteamDB Upcoming Games Tracker

An automated scraper built with Node.js and Playwright that bypasses Cloudflare protections to track the follower counts of upcoming Steam games week over week. 

This tracker pulls upcoming releases from [SteamDB](https://steamdb.info/upcoming/), bypasses bot-protection using Playwright Stealth, and dynamically generates an Excel (`.xlsx`) report.

## Features
- **Cloudflare Bypass:** Uses a persistent browser context and `playwright-extra` stealth plugins to reliably navigate SteamDB without getting blocked.
- **Automated Lookahead:** Automatically calculates and scrapes the current week + the next 5 weeks (approx. 1.5 months).
- **Intelligent Excel Generation:** 
  - Generates a **Master Database** (`steamdb_master_data.json`) so no history is ever lost.
  - Outputs a 2-tab Excel file (`steamdb_upcoming_tracker.xlsx`).
  - **Tab 1 ("All Games"):** Complete historical tracking of every game.
  - **Tab 2 ("> 1000 Followers"):** A focused, noise-free view. Games automatically graduate into this tab once they hit your configured threshold, carrying their past follower history with them!

## Prerequisites
- **Node.js** (v16 or higher)
- **Windows / macOS / Linux**

## Installation

1. Clone or download this repository.
2. Open your terminal/command prompt and navigate to the folder.
3. Install the dependencies:
   ```bash
   npm install
   ```
4. Install the Playwright Chromium browser:
   ```bash
   npx playwright install chromium
   ```

## Usage

To run the tracker, execute the following command:
```bash
node steamdb_tracker.js
```
*Note: The first time you run the script, a visible browser will appear. If Cloudflare prompts you with a "Verify you are human" checkbox, click it. Playwright will save a clearance cookie in the `playwright_data` folder, allowing future runs to proceed automatically.*

## Configuration
You can easily change the follower threshold for the focused Excel tab. 
Open `steamdb_tracker.js` and edit this line at the top of the file:
```javascript
const FOCUSED_THRESHOLD = 1000;
```

## Automating Daily Execution (Windows Task Scheduler)
1. Open the Start menu and search for **Task Scheduler**.
2. Click **Create Basic Task...** on the right pane.
3. **Name:** "SteamDB Daily Tracker"
4. **Trigger:** Select **Daily**, pick a time (e.g., 9:00 AM).
5. **Action:** Start a program.
6. **Program/script:** Type `node`
7. **Add arguments:** Type `"C:\path\to\your\folder\steamdb_tracker.js"`
8. **Start in:** Type the folder path: `C:\path\to\your\folder`
9. Finish and save. The script will now quietly run in the background every day!

## Disclaimer
This tool is for educational and personal data tracking purposes. Please respect SteamDB's server load by running this script a maximum of once per day.
