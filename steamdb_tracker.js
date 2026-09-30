require('dotenv').config();

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();

chromium.use(stealth);

const FOCUSED_THRESHOLD = 1000;
const WEEKS_TO_SCRAPE = 6;
const STEAM_API_BATCH_SIZE = 4;
const STEAM_API_BATCH_DELAY_MS = 750;

const TRACKER_HEADERS = [
    'Release Date',
    'Game Title',
    'Followers',
    'Publisher',
    'Developer',
    'Is Free',
    'Genres',
    'Categories',
    'Platforms',
    'AppID',
    'Snapshot Date'
];

const MONTH_NAMES = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

async function scrapeSteamDB() {
    console.log('Starting SteamDB Tracker...');

    let browser;
    let context;
    let page;
    const isRemote = process.argv.includes('--remote');

    if (isRemote) {
        console.log('Connecting to the pre-launched Edge browser...');
        browser = await chromium.connectOverCDP('http://localhost:9222');
        context = browser.contexts()[0];
        page = context.pages().find(p => p.url().includes('steamdb.info'));
        if (!page) page = await context.newPage();
    } else {
        const userDataDir = path.join(__dirname, 'playwright_data');
        const launchOptions = {
            headless: false,
            userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
            viewport: { width: 1280, height: 720 },
            args: ['--disable-blink-features=AutomationControlled']
        };

        if (process.env.BROWSER_EXECUTABLE_PATH) {
            launchOptions.executablePath = process.env.BROWSER_EXECUTABLE_PATH;
            console.log(`Booting custom browser from: ${launchOptions.executablePath}`);
        } else if (process.env.BROWSER_CHANNEL) {
            launchOptions.channel = process.env.BROWSER_CHANNEL;
            console.log(`Booting browser channel: ${launchOptions.channel}`);
        } else {
            console.log('Booting default Playwright Chromium browser...');
        }

        context = await chromium.launchPersistentContext(userDataDir, launchOptions);
        await context.addInitScript(() => {
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        });
        page = await context.newPage();
    }

    const allGames = [];
    const weeksToScrape = getUpcomingWeeks(WEEKS_TO_SCRAPE);
    let successfulWeeks = 0;

    for (const week of weeksToScrape) {
        const url = `https://steamdb.info/upcoming/?sort=followers_desc&week=${week}`;
        console.log(`Navigating to ${url}`);

        await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

        try {
            console.log('Waiting for the table. Complete any Cloudflare check in the browser.');
            await page.waitForSelector('table.table-sales', { timeout: 90000 });
        } catch (error) {
            console.log(`Skipping ${week}: the table did not load.`);
            await page.screenshot({ path: path.join(__dirname, `debug_cloudflare_${week}.png`) });
            continue;
        }

        const games = await page.evaluate(() => {
            const results = [];

            document.querySelectorAll('table.table-sales tbody tr.app').forEach(row => {
                const appId = row.getAttribute('data-appid');
                const cells = Array.from(row.querySelectorAll('td'));
                if (!appId || cells.length < 8) return;

                const releaseTimestamp = Number(cells[6].getAttribute('data-sort'));
                results.push({
                    appId,
                    name: cells[2].innerText.trim().replace(/\n/g, ' '),
                    releaseDate: cells[6].innerText.trim(),
                    releaseTimestamp: Number.isFinite(releaseTimestamp) && releaseTimestamp > 0
                        ? releaseTimestamp
                        : null,
                    followers: Number(cells[7].getAttribute('data-sort'))
                        || parseInt(cells[7].innerText.replace(/,/g, ''), 10)
                        || 0
                });
            });

            return results;
        });

        allGames.push(...games);
        successfulWeeks += 1;
        console.log(`Extracted ${games.length} games from ${week}.`);
        await page.waitForTimeout(3000);
    }

    if (isRemote && browser) {
        await browser.close();
    } else {
        await context.close();
    }

    const crawlComplete = successfulWeeks === weeksToScrape.length;
    console.log(`Extracted ${allGames.length} rows across ${successfulWeeks}/${weeksToScrape.length} weeks.`);

    if (allGames.length > 0) {
        await saveData(allGames, {
            crawlDate: getLocalDateKey(),
            crawlComplete,
            expectedWeeks: weeksToScrape.length,
            successfulWeeks
        });
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

function getLocalDateKey(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function readJson(filePath, fallback) {
    if (!fs.existsSync(filePath)) return fallback;

    try {
        return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    } catch (error) {
        throw new Error(`Could not read ${path.basename(filePath)}: ${error.message}`);
    }
}

function writeJson(filePath, value) {
    fs.writeFileSync(filePath, JSON.stringify(value, null, 2), 'utf-8');
}

function migrateCsv(oldCsvFile) {
    const migrated = {};
    if (!fs.existsSync(oldCsvFile)) return migrated;

    console.log('Migrating the old CSV into the master database...');
    const rows = fs.readFileSync(oldCsvFile, 'utf-8').trim().split('\n');
    if (rows.length === 0) return migrated;

    const headers = rows[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const dateColumns = headers.slice(3).filter(h => /^\d{4}-\d{2}-\d{2}$/.test(h));

    for (let i = 1; i < rows.length; i++) {
        const rowRegex = /(".*?"|[^",\s]+)(?=\s*,|\s*$)/g;
        let match;
        const row = [];

        while ((match = rowRegex.exec(rows[i])) !== null) {
            row.push(match[1].replace(/^"|"$/g, ''));
        }

        if (row.length < 3) continue;
        const appId = row[0];
        migrated[appId] = { name: row[1], releaseDate: row[2], history: {} };
        dateColumns.forEach((date, index) => {
            const columnIndex = 3 + index;
            if (columnIndex < row.length) migrated[appId].history[date] = row[columnIndex];
        });
    }

    return migrated;
}

function deduplicateGames(games) {
    const uniqueGames = new Map();

    for (const game of games || []) {
        if (!game || !game.appId) continue;
        const existing = uniqueGames.get(String(game.appId));
        if (!existing || Number(game.followers) >= Number(existing.followers)) {
            uniqueGames.set(String(game.appId), {
                ...game,
                appId: String(game.appId),
                followers: Number(game.followers) || 0
            });
        }
    }

    return Array.from(uniqueGames.values());
}

function mergeScrapedGames(existingData, games, crawlDate) {
    for (const game of games) {
        const appId = String(game.appId);
        if (!existingData[appId]) {
            existingData[appId] = {
                name: game.name,
                releaseDate: game.releaseDate,
                history: {},
                firstSeen: crawlDate
            };
        }

        const record = existingData[appId];
        record.history = record.history || {};
        record.name = game.name;
        record.releaseDate = game.releaseDate;
        if (game.releaseTimestamp) record.releaseTimestamp = game.releaseTimestamp;
        record.firstSeen = record.firstSeen || crawlDate;
        record.lastSeen = crawlDate;
        record.history[crawlDate] = Number(game.followers) || 0;
    }
}

function daysBetween(startDateKey, endDateKey) {
    const start = parseDateKey(startDateKey);
    const end = parseDateKey(endDateKey);
    if (!start || !end) return Number.POSITIVE_INFINITY;
    return Math.floor((end.getTime() - start.getTime()) / 86400000);
}

function needsSteamMetadata(game, crawlDate) {
    if (game.metadataStatus === 'unavailable'
        && daysBetween(game.metadataLastChecked, crawlDate) < 30) {
        return false;
    }
    if (game.metadataStatus === 'error'
        && daysBetween(game.metadataLastChecked, crawlDate) < 1) {
        return false;
    }

    const requiredFields = ['developer', 'publisher', 'isFree', 'genres', 'categories', 'platforms'];
    return requiredFields.some(field => game[field] === undefined || game[field] === '' || game[field] === 'Error');
}

async function fetchSteamAppDetails(appId) {
    const response = await fetch(`https://store.steampowered.com/api/appdetails?appids=${appId}`, {
        headers: { Accept: 'application/json' }
    });

    if (!response.ok) {
        throw new Error(`Steam API returned HTTP ${response.status}`);
    }

    const payload = await response.json();
    const result = payload && payload[appId];
    return result && result.success ? result.data : null;
}

function applySteamMetadata(game, appData, crawlDate) {
    if (!appData) {
        game.metadataStatus = 'unavailable';
        game.metadataLastChecked = crawlDate;
        return;
    }

    const platformLabels = {
        windows: 'Windows',
        mac: 'macOS',
        linux: 'Linux'
    };

    game.developer = appData.developers ? appData.developers.join(', ') : 'N/A';
    game.publisher = appData.publishers ? appData.publishers.join(', ') : 'N/A';
    game.isFree = appData.is_free ? 'Yes' : 'No';
    game.genres = appData.genres ? appData.genres.map(item => item.description).join(', ') : 'N/A';
    game.categories = appData.categories ? appData.categories.map(item => item.description).join(', ') : 'N/A';
    game.platforms = appData.platforms
        ? Object.keys(appData.platforms)
            .filter(key => appData.platforms[key])
            .map(key => platformLabels[key] || key)
            .join(', ')
        : 'N/A';
    game.metadataStatus = 'complete';
    game.metadataLastChecked = crawlDate;
}

async function enrichSteamMetadata(existingData, appIds, crawlDate, masterFile, options = {}) {
    if (options.fetchMetadata === false) return;

    const fetcher = options.fetchAppDetails || fetchSteamAppDetails;
    const pendingIds = appIds.filter(appId => needsSteamMetadata(existingData[appId], crawlDate));
    console.log(`Steam metadata required for ${pendingIds.length}/${appIds.length} focused games.`);

    for (let index = 0; index < pendingIds.length; index += STEAM_API_BATCH_SIZE) {
        const batch = pendingIds.slice(index, index + STEAM_API_BATCH_SIZE);

        await Promise.all(batch.map(async appId => {
            try {
                const appData = await fetcher(appId);
                applySteamMetadata(existingData[appId], appData, crawlDate);
            } catch (error) {
                console.error(`Steam metadata failed for AppID ${appId}: ${error.message}`);
                existingData[appId].metadataStatus = 'error';
                existingData[appId].metadataLastChecked = crawlDate;
            }
        }));

        writeJson(masterFile, existingData);
        if (index + STEAM_API_BATCH_SIZE < pendingIds.length) {
            await new Promise(resolve => setTimeout(resolve, STEAM_API_BATCH_DELAY_MS));
        }
    }
}

function createSnapshotGame(appId, game, followers) {
    const normalizedPlatforms = String(game.platforms || '')
        .split(',')
        .map(platform => platform.trim())
        .filter(Boolean)
        .map(platform => ({ windows: 'Windows', mac: 'macOS', macos: 'macOS', linux: 'Linux' }[platform.toLowerCase()] || platform))
        .join(', ');

    return {
        appId: String(appId),
        releaseDate: game.releaseDate || '',
        releaseTimestamp: game.releaseTimestamp || null,
        name: game.name || '',
        followers: Number(followers) || 0,
        publisher: game.publisher || '',
        developer: game.developer || '',
        isFree: game.isFree || '',
        genres: game.genres || '',
        categories: game.categories || '',
        platforms: normalizedPlatforms
    };
}

function buildCurrentMonthlySnapshot(existingData, currentGames, crawlDate) {
    const games = currentGames
        .filter(game => Number(game.followers) >= FOCUSED_THRESHOLD)
        .map(game => createSnapshotGame(game.appId, existingData[game.appId], game.followers));

    return {
        snapshotDate: crawlDate,
        threshold: FOCUSED_THRESHOLD,
        gameCount: games.length,
        games
    };
}

function bootstrapMonthlySnapshots(existingData) {
    const datesByMonth = new Map();

    Object.values(existingData).forEach(game => {
        Object.keys(game.history || {}).forEach(date => {
            if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
            const monthKey = date.slice(0, 7);
            if (!datesByMonth.has(monthKey)) datesByMonth.set(monthKey, new Set());
            datesByMonth.get(monthKey).add(date);
        });
    });

    const months = {};
    for (const [monthKey, dates] of datesByMonth.entries()) {
        const snapshotDate = Array.from(dates).sort().pop();
        const games = Object.entries(existingData)
            .filter(([, game]) => game.history && game.history[snapshotDate] !== undefined)
            .filter(([, game]) => Number(game.history[snapshotDate]) >= FOCUSED_THRESHOLD)
            .map(([appId, game]) => createSnapshotGame(appId, game, game.history[snapshotDate]));

        months[monthKey] = {
            snapshotDate,
            threshold: FOCUSED_THRESHOLD,
            gameCount: games.length,
            migratedFromDailyHistory: true,
            games
        };
    }

    return { version: 1, months };
}

function parseDateKey(dateKey) {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateKey || '');
    if (!match) return '';
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function parseReleaseDate(game, snapshotDate) {
    const rawDate = String(game.releaseDate || '').trim();
    const exactDateMatch = /^(\d{1,2})\s+([A-Za-z]{3,9})(?:,?\s+(\d{4}))?$/.exec(rawDate);
    if (!exactDateMatch) return rawDate;

    if (game.releaseTimestamp) {
        const timestampDate = new Date(Number(game.releaseTimestamp) * 1000);
        if (!Number.isNaN(timestampDate.getTime())) {
            return new Date(
                timestampDate.getUTCFullYear(),
                timestampDate.getUTCMonth(),
                timestampDate.getUTCDate()
            );
        }
    }

    const monthIndex = MONTH_NAMES.findIndex(month => (
        exactDateMatch[2].slice(0, 3).toLowerCase() === month.toLowerCase()
    ));
    if (monthIndex === -1) return rawDate;

    const referenceDate = parseDateKey(snapshotDate) || new Date();
    let year = exactDateMatch[3] ? Number(exactDateMatch[3]) : referenceDate.getFullYear();
    if (!exactDateMatch[3] && referenceDate.getMonth() - monthIndex > 6) year += 1;

    return new Date(year, monthIndex, Number(exactDateMatch[1]));
}

function getMonthSheetName(monthKey) {
    const match = /^(\d{4})-(\d{2})$/.exec(monthKey);
    if (!match) return monthKey.slice(0, 31);
    return `${MONTH_NAMES[Number(match[2]) - 1]} ${match[1]}`;
}

function sortSnapshotGames(games) {
    return [...games].sort((a, b) => {
        const aDate = a.releaseTimestamp ? Number(a.releaseTimestamp) : Number.MAX_SAFE_INTEGER;
        const bDate = b.releaseTimestamp ? Number(b.releaseTimestamp) : Number.MAX_SAFE_INTEGER;
        if (aDate !== bDate) return aDate - bDate;
        if (Number(a.followers) !== Number(b.followers)) return Number(b.followers) - Number(a.followers);
        return String(a.name).localeCompare(String(b.name));
    });
}

function createTrackerSheet(snapshot) {
    const rows = sortSnapshotGames(snapshot.games || []).map(game => ({
        'Release Date': parseReleaseDate(game, snapshot.snapshotDate),
        'Game Title': game.name,
        'Followers': Number(game.followers) || 0,
        'Publisher': game.publisher || '',
        'Developer': game.developer || '',
        'Is Free': game.isFree || '',
        'Genres': game.genres || '',
        'Categories': game.categories || '',
        'Platforms': game.platforms || '',
        'AppID': Number(game.appId),
        'Snapshot Date': parseDateKey(snapshot.snapshotDate)
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows, {
        header: TRACKER_HEADERS,
        cellDates: true,
        dateNF: 'dd mmm yyyy'
    });

    worksheet['!autofilter'] = { ref: worksheet['!ref'] || 'A1:K1' };
    worksheet['!cols'] = [
        { wch: 14 }, { wch: 42 }, { wch: 12 }, { wch: 28 }, { wch: 28 },
        { wch: 10 }, { wch: 32 }, { wch: 50 }, { wch: 20 }, { wch: 12 }, { wch: 14 }
    ];
    for (let rowIndex = 2; rowIndex <= rows.length + 1; rowIndex++) {
        if (worksheet[`A${rowIndex}`] && worksheet[`A${rowIndex}`].t === 'd') {
            worksheet[`A${rowIndex}`].z = 'dd mmm yyyy';
        }
        if (worksheet[`C${rowIndex}`]) worksheet[`C${rowIndex}`].z = '#,##0';
        if (worksheet[`J${rowIndex}`]) {
            const appId = worksheet[`J${rowIndex}`].v;
            worksheet[`J${rowIndex}`].l = {
                Target: `https://store.steampowered.com/app/${appId}/`,
                Tooltip: 'Open Steam store page'
            };
        }
        if (worksheet[`B${rowIndex}`]) {
            const appId = worksheet[`J${rowIndex}`].v;
            worksheet[`B${rowIndex}`].l = {
                Target: `https://steamdb.info/app/${appId}/`,
                Tooltip: 'Open SteamDB page'
            };
        }
        if (worksheet[`K${rowIndex}`] && worksheet[`K${rowIndex}`].t === 'd') {
            worksheet[`K${rowIndex}`].z = 'dd mmm yyyy';
        }
    }

    return worksheet;
}

function generateWorkbook(snapshotStore, excelFile) {
    const monthKeys = Object.keys(snapshotStore.months || {}).sort().reverse();
    if (monthKeys.length === 0) {
        console.log('No complete monthly snapshot is available yet; Excel was not changed.');
        return false;
    }

    const workbook = XLSX.utils.book_new();
    const latestSnapshot = snapshotStore.months[monthKeys[0]];
    XLSX.utils.book_append_sheet(workbook, createTrackerSheet(latestSnapshot), 'Current Tracker');

    for (const monthKey of monthKeys) {
        XLSX.utils.book_append_sheet(
            workbook,
            createTrackerSheet(snapshotStore.months[monthKey]),
            getMonthSheetName(monthKey)
        );
    }

    try {
        XLSX.writeFile(workbook, excelFile, { cellDates: true });
    } catch (error) {
        if (error.code === 'EBUSY') {
            throw new Error(`Close ${path.basename(excelFile)} in Excel, then run the tracker again.`);
        }
        throw error;
    }

    console.log(`Excel tracker generated with ${monthKeys.length} monthly tab(s): ${excelFile}`);
    return true;
}

async function saveData(games, options = {}) {
    const dataDir = options.dataDir || __dirname;
    const masterFile = path.join(dataDir, 'steamdb_master_data.json');
    const snapshotsFile = path.join(dataDir, 'steamdb_monthly_snapshots.json');
    const oldCsvFile = path.join(dataDir, 'steamdb_upcoming_tracker.csv');
    const excelFile = options.excelFile || path.join(dataDir, 'steamdb_upcoming_tracker.xlsx');
    const crawlDate = options.crawlDate || getLocalDateKey();
    const currentGames = deduplicateGames(games);

    const existingData = fs.existsSync(masterFile)
        ? readJson(masterFile, {})
        : migrateCsv(oldCsvFile);

    const snapshotsFileExists = fs.existsSync(snapshotsFile);
    const canBootstrap = currentGames.length === 0 || options.crawlComplete !== false;
    const bootstrappedSnapshots = !snapshotsFileExists && canBootstrap;
    const snapshotStore = snapshotsFileExists
        ? readJson(snapshotsFile, { version: 1, months: {} })
        : canBootstrap
            ? bootstrapMonthlySnapshots(existingData)
            : { version: 1, months: {} };

    snapshotStore.version = 1;
    snapshotStore.months = snapshotStore.months || {};

    const crawlCanBeCommitted = currentGames.length > 0 && options.crawlComplete !== false;

    if (crawlCanBeCommitted) {
        mergeScrapedGames(existingData, currentGames, crawlDate);

        const focusedAppIds = currentGames
            .filter(game => game.followers >= FOCUSED_THRESHOLD)
            .map(game => game.appId);

        await enrichSteamMetadata(existingData, focusedAppIds, crawlDate, masterFile, options);
        writeJson(masterFile, existingData);

        const currentSnapshot = buildCurrentMonthlySnapshot(existingData, currentGames, crawlDate);
        const monthKey = crawlDate.slice(0, 7);
        if (currentSnapshot.games.length > 0) {
            snapshotStore.months[monthKey] = currentSnapshot;
            snapshotStore.updatedAt = crawlDate;
            console.log(`Updated ${getMonthSheetName(monthKey)} with the ${crawlDate} snapshot.`);
        } else {
            console.log(`No games met the ${FOCUSED_THRESHOLD}-follower threshold; the monthly snapshot was not replaced.`);
        }
    } else if (currentGames.length > 0) {
        console.log(
            `Crawl data was not committed because the crawl was incomplete `
            + `(${options.successfulWeeks || 0}/${options.expectedWeeks || WEEKS_TO_SCRAPE} weeks).`
        );
    } else {
        writeJson(masterFile, existingData);
    }

    if (snapshotsFileExists || bootstrappedSnapshots) {
        writeJson(snapshotsFile, snapshotStore);
    }
    generateWorkbook(snapshotStore, excelFile);
}

if (require.main === module) {
    const command = process.argv.includes('--rebuild')
        ? saveData([], { crawlComplete: false })
        : scrapeSteamDB();

    command.catch(error => {
        console.error(error.message || error);
        process.exitCode = 1;
    });
}

module.exports = {
    bootstrapMonthlySnapshots,
    createTrackerSheet,
    generateWorkbook,
    getLocalDateKey,
    parseReleaseDate,
    saveData
};
