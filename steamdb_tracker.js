const { chromium } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();
chromium.use(stealth);
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// === CONFIGURATION ===
// You can change this number anytime. Games with followers >= this number go to the "Focused" tab.
const FOCUSED_THRESHOLD = 1000; 

async function scrapeSteamDB() {
    console.log("Starting SteamDB Tracker...");
    
    const userDataDir = path.join(__dirname, 'playwright_data');
    const context = await chromium.launchPersistentContext(userDataDir, { 
        headless: false,
        userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        viewport: { width: 1280, height: 720 },
        args: ['--disable-blink-features=AutomationControlled']
    });
    
    await context.addInitScript(() => {
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    });

    const page = await context.newPage();
    const allGames = [];
    const weeksToScrape = getUpcomingWeeks(6); 
    
    for (const week of weeksToScrape) {
        const url = `https://steamdb.info/upcoming/?week=${week}`;
        console.log(`Navigating to ${url}`);
        
        await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

        try {
            console.log(`Waiting for table to load... (🚨 IF YOU SEE A CLOUDFLARE CHECKBOX IN THE BROWSER, PLEASE CLICK IT! 🚨)`);
            await page.waitForSelector('table.table-sales tbody tr.app', { timeout: 90000 });
            console.log(`Table loaded for ${week}. Extracting...`);
        } catch (e) {
            console.log(`Skipping ${week} - Table didn't load. This usually means Cloudflare blocked us or you ran out of time to click the checkbox.`);
            continue;
        }

        const games = await page.evaluate(async () => {
            const results = [];
            document.querySelectorAll('tr.app').forEach(row => {
                const appId = row.getAttribute('data-appid');
                const cells = Array.from(row.querySelectorAll('td')).map(td => td.innerText.trim());
                if (cells.length >= 8) {
                    const followersStr = cells[7].replace(/,/g, '');
                    results.push({
                        appId: appId,
                        name: cells[2].replace(/\n/g, ' '),
                        releaseDate: cells[6],
                        followers: parseInt(followersStr) || 0
                    });
                }
            });
            return results;
        });
        
        allGames.push(...games);
        await page.waitForTimeout(3000); 
    }

    await context.close();
    console.log(`Extracted a total of ${allGames.length} upcoming games.`);

    if (allGames.length > 0) {
        saveData(allGames);
    }
}

function getUpcomingWeeks(numWeeks) {
    const weeks = [];
    const now = new Date();
    
    for (let i = 0; i < numWeeks; i++) {
        const d = new Date(now.getTime() + (i * 7 * 24 * 60 * 60 * 1000));
        const day = d.getDay() || 7;
        d.setDate(d.getDate() + 4 - day);
        const yearStart = new Date(d.getFullYear(), 0, 1);
        const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
        weeks.push(`${d.getFullYear()}W${weekNo.toString().padStart(2, '0')}`);
    }
    return [...new Set(weeks)];
}

function getLatestFollowers(gameData) {
    const dates = Object.keys(gameData.history).sort();
    if (dates.length === 0) return 0;
    return parseInt(gameData.history[dates[dates.length - 1]]) || 0;
}

function saveData(games) {
    const jsonFile = path.join(__dirname, 'steamdb_master_data.json');
    const oldCsvFile = path.join(__dirname, 'steamdb_upcoming_tracker.csv');
    const excelFile = path.join(__dirname, 'steamdb_upcoming_tracker.xlsx');
    const today = new Date().toISOString().split('T')[0];

    let existingData = {};

    // 1. Load existing data (migrate from old CSV if JSON doesn't exist yet)
    if (fs.existsSync(jsonFile)) {
        existingData = JSON.parse(fs.readFileSync(jsonFile, 'utf-8'));
    } else if (fs.existsSync(oldCsvFile)) {
        console.log("Migrating old CSV data into new master database...");
        const rows = fs.readFileSync(oldCsvFile, 'utf-8').trim().split('\n');
        if (rows.length > 0) {
            const headers = rows[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
            const dateColumns = headers.slice(3).filter(h => /^\d{4}-\d{2}-\d{2}$/.test(h));
            
            for (let i = 1; i < rows.length; i++) {
                const rowRegex = /(".*?"|[^",\s]+)(?=\s*,|\s*$)/g;
                let match, row = [];
                while ((match = rowRegex.exec(rows[i])) !== null) {
                    row.push(match[1].replace(/^"|"$/g, ''));
                }
                if (row.length >= 3) {
                    const appId = row[0];
                    existingData[appId] = { name: row[1], releaseDate: row[2], history: {} };
                    dateColumns.forEach((date, j) => {
                        const colIdx = 3 + j;
                        if (colIdx < row.length) existingData[appId].history[date] = row[colIdx];
                    });
                }
            }
        }
    }

    // 2. Merge today's scraped data
    if (games && games.length > 0) {
        games.forEach(game => {
            if (!existingData[game.appId]) {
                existingData[game.appId] = {
                    name: game.name,
                    releaseDate: game.releaseDate,
                    history: {}
                };
            }
            existingData[game.appId].releaseDate = game.releaseDate; 
            existingData[game.appId].history[today] = game.followers; 
        });
    }

    // Save JSON database
    fs.writeFileSync(jsonFile, JSON.stringify(existingData, null, 2), 'utf-8');

    // 3. Generate Excel file
    const allUniqueDates = new Set();
    Object.values(existingData).forEach(game => {
        Object.keys(game.history).forEach(date => allUniqueDates.add(date));
    });
    const sortedDates = Array.from(allUniqueDates).sort(); // Chronological

    const allApps = Object.keys(existingData).sort((a, b) => {
        return getLatestFollowers(existingData[b]) - getLatestFollowers(existingData[a]);
    });

    const focusedApps = allApps.filter(appId => getLatestFollowers(existingData[appId]) >= FOCUSED_THRESHOLD);

    const createSheetData = (appsList) => {
        return appsList.map(appId => {
            const data = existingData[appId];
            const row = {
                'AppID': appId,
                'Name': data.name,
                'Release Date': data.releaseDate
            };
            sortedDates.forEach(date => {
                row[date] = data.history[date] !== undefined ? data.history[date] : '';
            });
            return row;
        });
    };

    const wb = XLSX.utils.book_new();
    
    const wsFocused = XLSX.utils.json_to_sheet(createSheetData(focusedApps));
    XLSX.utils.book_append_sheet(wb, wsFocused, `> ${FOCUSED_THRESHOLD} Followers`);
    
    const wsAll = XLSX.utils.json_to_sheet(createSheetData(allApps));
    XLSX.utils.book_append_sheet(wb, wsAll, 'All Games');

    XLSX.writeFile(wb, excelFile);
    console.log(`[SUCCESS] Excel file generated at ${excelFile}`);
}

// If run directly via node, scrape SteamDB. 
// (If you just want to test Excel generation without scraping, you can call saveData([]) )
scrapeSteamDB().catch(console.error);