require('dotenv').config();

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { chromium } = require('playwright-extra');
const stealth = require('puppeteer-extra-plugin-stealth')();

chromium.use(stealth);

const configuredThreshold = Number.parseInt(process.env.FOCUSED_THRESHOLD || '', 10);
const MINIMUM_FOLLOWER_THRESHOLD = 1000;
const FOCUSED_THRESHOLD = Number.isFinite(configuredThreshold)
    ? Math.max(configuredThreshold, MINIMUM_FOLLOWER_THRESHOLD)
    : MINIMUM_FOLLOWER_THRESHOLD;
const WEEKS_TO_SCRAPE = 6;
const STEAM_API_BATCH_SIZE = 1;
const STEAM_API_BATCH_DELAY_MS = 1600;
const STEAM_API_MAX_RETRIES = 2;
const STEAM_METADATA_RECHECK_DAYS = 7;
const EXCEL_DATE_FORMAT = 'dd mmm yyyy';

const MONTH_NAMES = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

const COLORS = {
    navy: 'FF17324D',
    teal: 'FF0F766E',
    lightTeal: 'FFDFF3F0',
    blue: 'FF2563A6',
    lightBlue: 'FFEAF2F8',
    orange: 'FFF59E0B',
    lightOrange: 'FFFFF1D6',
    green: 'FF2E7D32',
    lightGreen: 'FFE4F2E5',
    red: 'FFB42318',
    lightRed: 'FFFDE7E5',
    grey: 'FF64748B',
    lightGrey: 'FFF1F5F9',
    white: 'FFFFFFFF',
    border: 'FFD7E0E8'
};

const TIER_STYLES = {
    'Tier 2': { fill: COLORS.lightOrange, font: 'FF8A4B08' },
    'Tier 3': { fill: COLORS.lightGreen, font: COLORS.green },
    'Tier 4': { fill: COLORS.lightBlue, font: COLORS.blue },
    'Tier 5': { fill: 'FFFFE7A3', font: 'FF7A4B00' }
};

const FIXED_COLUMNS = [
    { header: 'Release Date', key: 'releaseDate', width: 15 },
    { header: 'Game Title', key: 'gameTitle', width: 42 },
    { header: 'Followers', key: 'followers', width: 13 },
    { header: 'Publisher', key: 'publisher', width: 28 },
    { header: 'Developer', key: 'developer', width: 28 },
    { header: 'Follower Tier', key: 'followerTier', width: 14 },
    { header: 'Commercial Standing', key: 'commercialStanding', width: 34 },
    { header: 'Is Free', key: 'isFree', width: 10 },
    { header: 'Genres', key: 'genres', width: 30 },
    { header: 'Categories', key: 'categories', width: 48 },
    { header: 'Platforms', key: 'platforms', width: 20 },
    { header: 'AppID', key: 'appId', width: 13 },
    { header: 'Status', key: 'status', width: 18 },
    { header: 'Previous Crawl Followers', key: 'previousFollowers', width: 23 },
    { header: 'Change Since Previous Crawl', key: 'changeSincePrevious', width: 27 },
    { header: 'First Seen', key: 'firstSeen', width: 14 },
    { header: 'Last Updated', key: 'lastUpdated', width: 14 }
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
        page = context.pages().find(candidate => candidate.url().includes('steamdb.info'));
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

    for (let index = 0; index < numWeeks; index++) {
        const date = new Date(now.getTime() + (index * 7 * 24 * 60 * 60 * 1000));
        const day = date.getDay() || 7;
        date.setDate(date.getDate() + 4 - day);
        const yearStart = new Date(date.getFullYear(), 0, 1);
        const weekNumber = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
        weeks.push(`${date.getFullYear()}W${weekNumber.toString().padStart(2, '0')}`);
    }

    return [...new Set(weeks)];
}

function getLocalDateKey(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function parseDateKey(dateKey) {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateKey || '');
    if (!match) return null;
    return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
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
    const temporaryFile = `${filePath}.tmp`;
    fs.writeFileSync(temporaryFile, JSON.stringify(value, null, 2), 'utf-8');
    fs.renameSync(temporaryFile, filePath);
}

function migrateCsv(oldCsvFile) {
    const migrated = {};
    if (!fs.existsSync(oldCsvFile)) return migrated;

    console.log('Migrating the old CSV into the master database...');
    const rows = fs.readFileSync(oldCsvFile, 'utf-8').trim().split('\n');
    if (rows.length === 0) return migrated;

    const headers = rows[0].split(',').map(header => header.trim().replace(/^"|"$/g, ''));
    const dateColumns = headers.slice(3).filter(header => /^\d{4}-\d{2}-\d{2}$/.test(header));

    for (let index = 1; index < rows.length; index++) {
        const rowRegex = /(".*?"|[^",\s]+)(?=\s*,|\s*$)/g;
        let match;
        const row = [];

        while ((match = rowRegex.exec(rows[index])) !== null) {
            row.push(match[1].replace(/^"|"$/g, ''));
        }

        if (row.length < 3) continue;
        const appId = row[0];
        migrated[appId] = { name: row[1], releaseDate: row[2], history: {} };
        dateColumns.forEach((date, dateIndex) => {
            const columnIndex = 3 + dateIndex;
            if (columnIndex < row.length) migrated[appId].history[date] = row[columnIndex];
        });
    }

    return migrated;
}

function deduplicateGames(games) {
    const uniqueGames = new Map();

    for (const game of games || []) {
        if (!game || !game.appId) continue;
        const appId = String(game.appId);
        uniqueGames.set(appId, {
            ...game,
            appId,
            followers: Number(game.followers) || 0
        });
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
                releaseTimestamp: game.releaseTimestamp || null,
                history: {},
                firstSeen: crawlDate
            };
        }

        const record = existingData[appId];
        record.history = record.history || {};
        record.name = game.name;
        if (String(game.releaseDate || '').trim()) {
            record.releaseDate = game.releaseDate;
            record.releaseTimestamp = game.releaseTimestamp || null;
        }
        record.firstSeen = record.firstSeen || crawlDate;
        record.lastSeen = crawlDate;
        record.history[crawlDate] = Number(game.followers) || 0;
    }
}

function getHistoryDates(game) {
    return Object.keys((game && game.history) || {})
        .filter(date => /^\d{4}-\d{2}-\d{2}$/.test(date))
        .sort();
}

function getLatestFollowers(game) {
    const dates = getHistoryDates(game);
    if (dates.length === 0) return 0;
    return Number(game.history[dates[dates.length - 1]]) || 0;
}

function getLatestCrawlDate(existingData) {
    const dates = new Set();
    Object.values(existingData).forEach(game => {
        getHistoryDates(game).forEach(date => dates.add(date));
    });
    return Array.from(dates).sort().pop() || null;
}

function daysBetween(startDateKey, endDateKey) {
    const start = parseDateKey(startDateKey);
    const end = parseDateKey(endDateKey);
    if (!start || !end) return Number.POSITIVE_INFINITY;
    return Math.floor((toUtcDay(end) - toUtcDay(start)) / 86400000);
}

function needsSteamMetadata(game, crawlDate) {
    if (game.metadataStatus === 'unavailable'
        && daysBetween(game.metadataLastChecked, crawlDate) < STEAM_METADATA_RECHECK_DAYS) {
        return false;
    }
    if (game.metadataStatus === 'error'
        && daysBetween(game.metadataLastChecked, crawlDate) < 1) {
        return false;
    }

    const requiredFields = ['developer', 'publisher', 'isFree', 'genres', 'categories', 'platforms'];
    return requiredFields.some(field => game[field] === undefined || game[field] === '' || game[field] === 'Error');
}

async function fetchSteamAppDetails(appId, attempt = 0) {
    const response = await fetch(`https://store.steampowered.com/api/appdetails?appids=${appId}`, {
        headers: { Accept: 'application/json' }
    });

    if (response.status === 429 && attempt < STEAM_API_MAX_RETRIES) {
        const retryAfter = Number(response.headers.get('retry-after')) || 5;
        console.log(`Steam API rate-limited AppID ${appId}; retrying in ${retryAfter} seconds.`);
        await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
        return fetchSteamAppDetails(appId, attempt + 1);
    }

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

    const platformLabels = { windows: 'Windows', mac: 'macOS', linux: 'Linux' };
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
    console.log(`Steam metadata required for ${pendingIds.length}/${appIds.length} qualifying games.`);

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

function normalizePlatforms(value) {
    const labels = { windows: 'Windows', mac: 'macOS', macos: 'macOS', linux: 'Linux' };
    return String(value || '')
        .split(',')
        .map(platform => platform.trim())
        .filter(Boolean)
        .map(platform => labels[platform.toLowerCase()] || platform)
        .join(', ');
}

function findMonthIndex(monthName) {
    const abbreviation = String(monthName || '').slice(0, 3).toLowerCase();
    return MONTH_NAMES.findIndex(month => month.toLowerCase() === abbreviation);
}

function inferReleaseYear(monthIndex, referenceDateKey) {
    const referenceDate = parseDateKey(referenceDateKey) || new Date();
    let year = referenceDate.getUTCFullYear();
    if (referenceDate.getUTCMonth() - monthIndex > 6) year += 1;
    return year;
}

function getReleaseInfo(game, referenceDateKey) {
    const rawDate = String(game.releaseDate || '').trim();
    const exactDateMatch = /^(\d{1,2})\s+([A-Za-z]{3,9})(?:,?\s+(\d{4}))?$/.exec(rawDate);

    if (exactDateMatch) {
        let exactDate = null;
        if (game.releaseTimestamp) {
            const timestampDate = new Date(Number(game.releaseTimestamp) * 1000);
            if (!Number.isNaN(timestampDate.getTime())) {
                exactDate = new Date(Date.UTC(
                    timestampDate.getUTCFullYear(),
                    timestampDate.getUTCMonth(),
                    timestampDate.getUTCDate()
                ));
            }
        }

        if (!exactDate) {
            const monthIndex = findMonthIndex(exactDateMatch[2]);
            if (monthIndex !== -1) {
                const year = exactDateMatch[3]
                    ? Number(exactDateMatch[3])
                    : inferReleaseYear(monthIndex, referenceDateKey);
                exactDate = new Date(Date.UTC(year, monthIndex, Number(exactDateMatch[1])));
            }
        }

        if (exactDate && !Number.isNaN(exactDate.getTime())) {
            return {
                monthKey: `${exactDate.getUTCFullYear()}-${String(exactDate.getUTCMonth() + 1).padStart(2, '0')}`,
                excelValue: exactDate,
                exactDate,
                sortValue: exactDate.getTime()
            };
        }
    }

    const monthYearMatch = /^([A-Za-z]{3,9})\s+(\d{4})$/.exec(rawDate);
    if (monthYearMatch) {
        const monthIndex = findMonthIndex(monthYearMatch[1]);
        if (monthIndex !== -1) {
            return {
                monthKey: `${monthYearMatch[2]}-${String(monthIndex + 1).padStart(2, '0')}`,
                excelValue: rawDate,
                exactDate: null,
                sortValue: Date.UTC(Number(monthYearMatch[2]), monthIndex, 1)
            };
        }
    }

    return {
        monthKey: null,
        excelValue: rawDate || 'TBA',
        exactDate: null,
        sortValue: Number.MAX_SAFE_INTEGER
    };
}

function toUtcDay(date) {
    return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function getDaysToRelease(exactDate, referenceDateKey) {
    const referenceDate = parseDateKey(referenceDateKey);
    if (!exactDate || !referenceDate) return null;
    return Math.round((toUtcDay(exactDate) - toUtcDay(referenceDate)) / 86400000);
}

function getReleaseStatus(releaseInfo, referenceDateKey) {
    if (!releaseInfo.monthKey) return 'Unscheduled';

    if (releaseInfo.exactDate) {
        const days = getDaysToRelease(releaseInfo.exactDate, referenceDateKey);
        if (days < 0) return 'Released';
        if (days === 0) return 'Releases Today';
        return 'Upcoming';
    }

    const referenceMonth = String(referenceDateKey || '').slice(0, 7);
    if (releaseInfo.monthKey < referenceMonth) return 'Released (Month)';
    if (releaseInfo.monthKey === referenceMonth) return 'Launch Month';
    return 'Upcoming (Month)';
}

function getFollowerInterpretation(followers) {
    if (followers < MINIMUM_FOLLOWER_THRESHOLD) {
        return { tier: '', standing: '' };
    }
    if (followers >= 30000) {
        return {
            tier: 'Tier 5',
            standing: 'Major Commercial Hit / Megahit'
        };
    }
    if (followers >= 10000) {
        return {
            tier: 'Tier 4',
            standing: 'AA / Mid-Tier Blockbuster'
        };
    }
    if (followers >= 3000) {
        return {
            tier: 'Tier 3',
            standing: 'Commercial Hit / Sustainable Indie'
        };
    }
    return {
        tier: 'Tier 2',
        standing: 'Barely Viable / Solo Indie Floor'
    };
}

function buildGameView(appId, game, referenceDateKey) {
    const historyDates = getHistoryDates(game);
    const latestDate = historyDates[historyDates.length - 1] || null;
    const previousDate = historyDates.length > 1 ? historyDates[historyDates.length - 2] : null;
    const latestFollowers = latestDate ? Number(game.history[latestDate]) || 0 : 0;
    const previousFollowers = previousDate ? Number(game.history[previousDate]) || 0 : null;
    const releaseInfo = getReleaseInfo(game, latestDate || referenceDateKey);
    const followerInterpretation = getFollowerInterpretation(latestFollowers);

    return {
        appId: String(appId),
        game,
        historyDates,
        latestDate,
        latestFollowers,
        previousFollowers,
        changeSincePrevious: previousFollowers === null ? null : latestFollowers - previousFollowers,
        followerTier: followerInterpretation.tier,
        commercialStanding: followerInterpretation.standing,
        firstSeen: game.firstSeen || historyDates[0] || null,
        releaseInfo,
        status: getReleaseStatus(releaseInfo, referenceDateKey),
        daysToRelease: getDaysToRelease(releaseInfo.exactDate, referenceDateKey)
    };
}

function buildWorkbookModel(existingData) {
    const latestCrawlDate = getLatestCrawlDate(existingData);
    const views = Object.entries(existingData)
        .filter(([, game]) => game && typeof game === 'object' && getHistoryDates(game).length > 0)
        .map(([appId, game]) => buildGameView(appId, game, latestCrawlDate))
        .filter(view => view.latestFollowers >= FOCUSED_THRESHOLD);

    const releaseMonths = new Map();
    const unscheduled = [];

    for (const view of views) {
        if (!view.releaseInfo.monthKey) {
            unscheduled.push(view);
            continue;
        }
        if (!releaseMonths.has(view.releaseInfo.monthKey)) {
            releaseMonths.set(view.releaseInfo.monthKey, []);
        }
        releaseMonths.get(view.releaseInfo.monthKey).push(view);
    }

    return {
        latestCrawlDate,
        views,
        allUpcoming: views.filter(view => view.latestDate === latestCrawlDate),
        releaseMonths,
        unscheduled
    };
}

function getMonthSheetName(monthKey) {
    const match = /^(\d{4})-(\d{2})$/.exec(monthKey);
    if (!match) return monthKey.slice(0, 31);
    return `${MONTH_NAMES[Number(match[2]) - 1]} ${match[1]}`;
}

function sortGameViews(views) {
    return [...views].sort((left, right) => {
        if (left.releaseInfo.sortValue !== right.releaseInfo.sortValue) {
            return left.releaseInfo.sortValue - right.releaseInfo.sortValue;
        }
        if (left.latestFollowers !== right.latestFollowers) {
            return right.latestFollowers - left.latestFollowers;
        }
        return String(left.game.name).localeCompare(String(right.game.name));
    });
}

function historyKey(date) {
    return `history_${date.replace(/-/g, '_')}`;
}

function getSheetHistoryDates(views) {
    const dates = new Set();
    views.forEach(view => view.historyDates.forEach(date => dates.add(date)));
    return Array.from(dates).sort();
}

function buildTrackerRow(view, historyDates) {
    const row = {
        releaseDate: view.releaseInfo.excelValue,
        gameTitle: view.game.name || '',
        followers: view.latestFollowers,
        publisher: view.game.publisher || '',
        developer: view.game.developer || '',
        followerTier: view.followerTier,
        commercialStanding: view.commercialStanding,
        isFree: view.game.isFree || '',
        genres: view.game.genres || '',
        categories: view.game.categories || '',
        platforms: normalizePlatforms(view.game.platforms),
        appId: Number(view.appId),
        status: view.status,
        previousFollowers: view.previousFollowers,
        changeSincePrevious: view.changeSincePrevious,
        firstSeen: view.firstSeen ? parseDateKey(view.firstSeen) : null,
        lastUpdated: view.latestDate ? parseDateKey(view.latestDate) : null
    };

    historyDates.forEach(date => {
        row[historyKey(date)] = view.game.history[date] === undefined
            ? null
            : Number(view.game.history[date]) || 0;
    });
    return row;
}

function styleTrackerWorksheet(worksheet, rowCount, historyDates) {
    worksheet.views = [{
        state: 'frozen',
        xSplit: 2,
        ySplit: 1,
        topLeftCell: 'C2',
        activeCell: 'C2',
        showGridLines: false
    }];
    worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: FIXED_COLUMNS.length + historyDates.length }
    };
    worksheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    worksheet.properties.defaultRowHeight = 22;

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell(cell => {
        cell.font = { name: 'Aptos Display', size: 11, bold: true, color: { argb: COLORS.white } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = { bottom: { style: 'medium', color: { argb: COLORS.teal } } };
    });

    for (let rowIndex = 2; rowIndex <= rowCount + 1; rowIndex++) {
        const row = worksheet.getRow(rowIndex);
        row.height = 45;
        row.eachCell({ includeEmpty: true }, cell => {
            cell.font = { name: 'Aptos', size: 10, color: { argb: COLORS.navy } };
            cell.alignment = { vertical: 'top' };
            cell.border = { bottom: { style: 'hair', color: { argb: COLORS.border } } };
            if (rowIndex % 2 === 0) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
            }
        });

        ['B', 'D', 'E', 'G', 'I', 'J', 'K'].forEach(column => {
            row.getCell(column).alignment = { vertical: 'top', wrapText: true };
        });
        row.getCell('A').numFmt = EXCEL_DATE_FORMAT;
        row.getCell('C').numFmt = '#,##0';
        row.getCell('N').numFmt = '#,##0';
        row.getCell('O').numFmt = '+#,##0;-#,##0;0';
        row.getCell('P').numFmt = EXCEL_DATE_FORMAT;
        row.getCell('Q').numFmt = EXCEL_DATE_FORMAT;

        for (let columnIndex = FIXED_COLUMNS.length + 1;
            columnIndex <= FIXED_COLUMNS.length + historyDates.length;
            columnIndex++) {
            row.getCell(columnIndex).numFmt = '#,##0';
        }
    }

    if (rowCount === 0) return;
    const lastRow = rowCount + 1;

    worksheet.addConditionalFormatting({
        ref: `C2:C${lastRow}`,
        rules: [{
            type: 'colorScale',
            cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
            color: [
                { argb: 'FFFFF1D6' },
                { argb: 'FFDFF3F0' },
                { argb: 'FF9BD5C8' }
            ]
        }]
    });
    worksheet.addConditionalFormatting({
        ref: `O2:O${lastRow}`,
        rules: [
            {
                type: 'cellIs',
                operator: 'greaterThan',
                formulae: [0],
                style: {
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGreen } },
                    font: { color: { argb: COLORS.green } }
                }
            },
            {
                type: 'cellIs',
                operator: 'lessThan',
                formulae: [0],
                style: {
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightRed } },
                    font: { color: { argb: COLORS.red } }
                }
            }
        ]
    });
}

function addTrackerWorksheet(workbook, name, views, tabColor) {
    const sortedViews = sortGameViews(views);
    const historyDates = getSheetHistoryDates(sortedViews);
    const worksheet = workbook.addWorksheet(name, {
        properties: { tabColor: { argb: tabColor } }
    });

    worksheet.columns = [
        ...FIXED_COLUMNS,
        ...historyDates.map(date => ({ header: date, key: historyKey(date), width: 13 }))
    ];

    sortedViews.forEach(view => {
        worksheet.addRow(buildTrackerRow(view, historyDates));
    });

    styleTrackerWorksheet(worksheet, sortedViews.length, historyDates);

    sortedViews.forEach((view, index) => {
        const row = worksheet.getRow(index + 2);
        const tierColor = TIER_STYLES[view.followerTier];
        if (tierColor) {
            ['followerTier', 'commercialStanding'].forEach(column => {
                row.getCell(column).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: tierColor.fill }
                };
                row.getCell(column).font = {
                    name: 'Aptos',
                    size: 10,
                    bold: true,
                    color: { argb: tierColor.font }
                };
            });
        }
        if (view.daysToRelease !== null && view.daysToRelease >= 0 && view.daysToRelease <= 14) {
            row.getCell('releaseDate').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: COLORS.lightOrange }
            };
            row.getCell('releaseDate').font = { color: { argb: 'FF8A4B08' }, bold: true };
        }
        const titleCell = row.getCell('gameTitle');
        titleCell.value = {
            text: view.game.name || '',
            hyperlink: `https://steamdb.info/app/${view.appId}/`,
            tooltip: 'Open SteamDB page'
        };
        titleCell.font = { name: 'Aptos', size: 10, color: { argb: COLORS.blue }, underline: true };

        const appIdCell = row.getCell('appId');
        appIdCell.value = {
            text: view.appId,
            hyperlink: `https://store.steampowered.com/app/${view.appId}/`,
            tooltip: 'Open Steam store page'
        };
        appIdCell.font = { name: 'Aptos', size: 10, color: { argb: COLORS.blue }, underline: true };
    });
    return worksheet;
}

function styleSectionTitle(worksheet, range, title) {
    worksheet.mergeCells(range);
    const cell = worksheet.getCell(range.split(':')[0]);
    cell.value = title;
    cell.font = { name: 'Aptos Display', size: 12, bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.teal } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
}

function styleSummaryHeader(row, startColumn, endColumn) {
    for (let column = startColumn; column <= endColumn; column++) {
        const cell = row.getCell(column);
        cell.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
}

function addMetricCard(worksheet, startColumn, labelRow, label, value, fillColor) {
    worksheet.mergeCells(labelRow, startColumn, labelRow, startColumn + 1);
    worksheet.mergeCells(labelRow + 1, startColumn, labelRow + 2, startColumn + 1);

    const labelCell = worksheet.getCell(labelRow, startColumn);
    labelCell.value = label;
    labelCell.font = { name: 'Aptos', size: 9, bold: true, color: { argb: COLORS.grey } };
    labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGrey } };
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };

    const valueCell = worksheet.getCell(labelRow + 1, startColumn);
    valueCell.value = value;
    valueCell.font = { name: 'Aptos Display', size: 18, bold: true, color: { argb: COLORS.navy } };
    valueCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
    valueCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    if (value instanceof Date) valueCell.numFmt = EXCEL_DATE_FORMAT;
    if (typeof value === 'number') valueCell.numFmt = '#,##0';
}

function addSummarySheet(workbook, model) {
    const worksheet = workbook.addWorksheet('Summary', {
        properties: { tabColor: { argb: COLORS.orange } },
        views: [{ state: 'frozen', ySplit: 2, showGridLines: false }]
    });
    worksheet.getColumn('A').width = 26;
    worksheet.getColumn('B').width = 16;
    worksheet.getColumn('C').width = 18;
    worksheet.getColumn('D').width = 18;
    worksheet.getColumn('E').width = 18;
    worksheet.getColumn('F').width = 4;
    worksheet.getColumn('G').width = 34;
    worksheet.getColumn('H').width = 15;
    worksheet.getColumn('I').width = 14;
    worksheet.getColumn('J').width = 14;
    worksheet.getColumn('K').width = 14;
    worksheet.getColumn('L').width = 16;

    worksheet.mergeCells('A1:L1');
    worksheet.getCell('A1').value = 'SteamDB Release Tracker';
    worksheet.getCell('A1').font = { name: 'Aptos Display', size: 24, bold: true, color: { argb: COLORS.white } };
    worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getRow(1).height = 40;

    worksheet.mergeCells('A2:L2');
    worksheet.getCell('A2').value = 'Games are grouped by launch month; follower history is retained across crawls.';
    worksheet.getCell('A2').font = { name: 'Aptos', size: 10, italic: true, color: { argb: COLORS.grey } };
    worksheet.getCell('A2').alignment = { vertical: 'middle' };

    const newThisCrawl = model.allUpcoming.filter(view => view.firstSeen === model.latestCrawlDate).length;
    const nextThirtyDays = model.allUpcoming.filter(view => (
        view.daysToRelease !== null && view.daysToRelease >= 0 && view.daysToRelease <= 30
    ));
    const highlightedReleases = sortGameViews(nextThirtyDays).slice(0, 25);
    const metrics = [
        ['Last Complete Crawl', parseDateKey(model.latestCrawlDate), COLORS.lightBlue],
        ['Qualified Upcoming', model.allUpcoming.length, COLORS.lightTeal],
        ['New This Crawl', newThisCrawl, COLORS.lightOrange],
        ['Releases in 30 Days', nextThirtyDays.length, COLORS.lightGreen],
        ['Release Months', model.releaseMonths.size, COLORS.lightBlue],
        ['Unscheduled', model.unscheduled.length, COLORS.lightOrange],
        ['Total Qualified Tracked', model.views.length, COLORS.lightTeal],
        ['Follower Threshold', FOCUSED_THRESHOLD, COLORS.lightGreen]
    ];

    metrics.slice(0, 4).forEach((metric, index) => {
        addMetricCard(worksheet, 1 + (index * 3), 4, metric[0], metric[1], metric[2]);
    });
    metrics.slice(4).forEach((metric, index) => {
        addMetricCard(worksheet, 1 + (index * 3), 8, metric[0], metric[1], metric[2]);
    });

    const monthStartRow = 13;
    styleSectionTitle(worksheet, `A${monthStartRow}:E${monthStartRow}`, 'Release Month Overview');
    const monthHeaderRow = worksheet.getRow(monthStartRow + 1);
    ['Release Month', 'Qualified Games', 'Total Followers', 'Average Followers', 'Top Game']
        .forEach((header, index) => monthHeaderRow.getCell(index + 1).value = header);
    styleSummaryHeader(monthHeaderRow, 1, 5);

    const monthKeys = Array.from(model.releaseMonths.keys()).sort();
    monthKeys.forEach((monthKey, index) => {
        const views = model.releaseMonths.get(monthKey);
        const totalFollowers = views.reduce((sum, view) => sum + view.latestFollowers, 0);
        const topGame = [...views].sort((left, right) => right.latestFollowers - left.latestFollowers)[0];
        const row = worksheet.getRow(monthStartRow + 2 + index);
        row.values = [
            getMonthSheetName(monthKey),
            views.length,
            totalFollowers,
            views.length ? Math.round(totalFollowers / views.length) : 0,
            topGame ? topGame.game.name : ''
        ];
        row.getCell(3).numFmt = '#,##0';
        row.getCell(4).numFmt = '#,##0';
        if (topGame) {
            row.getCell(5).value = {
                text: topGame.game.name,
                hyperlink: `https://steamdb.info/app/${topGame.appId}/`
            };
            row.getCell(5).font = { color: { argb: COLORS.blue }, underline: true };
        }
    });

    styleSectionTitle(worksheet, `G${monthStartRow}:L${monthStartRow}`, 'Top Movers Since Previous Crawl');
    const moverHeaderRow = worksheet.getRow(monthStartRow + 1);
    ['Game Title', 'Release Date', 'Followers', 'Previous', 'Change', 'Release Month']
        .forEach((header, index) => moverHeaderRow.getCell(index + 7).value = header);
    styleSummaryHeader(moverHeaderRow, 7, 12);

    const topMovers = [...model.allUpcoming]
        .filter(view => view.changeSincePrevious !== null)
        .sort((left, right) => right.changeSincePrevious - left.changeSincePrevious)
        .slice(0, 10);

    topMovers.forEach((view, index) => {
        const row = worksheet.getRow(monthStartRow + 2 + index);
        row.getCell(7).value = {
            text: view.game.name,
            hyperlink: `https://steamdb.info/app/${view.appId}/`
        };
        row.getCell(7).font = { color: { argb: COLORS.blue }, underline: true };
        row.getCell(8).value = view.releaseInfo.excelValue;
        row.getCell(8).numFmt = EXCEL_DATE_FORMAT;
        row.getCell(9).value = view.latestFollowers;
        row.getCell(10).value = view.previousFollowers;
        row.getCell(11).value = view.changeSincePrevious;
        row.getCell(12).value = view.releaseInfo.monthKey
            ? getMonthSheetName(view.releaseInfo.monthKey)
            : 'Unscheduled';
        [9, 10].forEach(column => row.getCell(column).numFmt = '#,##0');
        row.getCell(11).numFmt = '+#,##0;-#,##0;0';
    });

    if (topMovers.length > 0) {
        worksheet.addConditionalFormatting({
            ref: `K${monthStartRow + 2}:K${monthStartRow + 1 + topMovers.length}`,
            rules: [
                {
                    type: 'cellIs',
                    operator: 'greaterThan',
                    formulae: [0],
                    style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGreen } } }
                },
                {
                    type: 'cellIs',
                    operator: 'lessThan',
                    formulae: [0],
                    style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightRed } } }
                }
            ]
        });
    }

    const detailStartRow = monthStartRow + Math.max(monthKeys.length, topMovers.length, 1) + 4;
    styleSectionTitle(worksheet, `A${detailStartRow}:E${detailStartRow}`, 'Next 25 Releases Within 30 Days');
    const releaseHeaderRow = worksheet.getRow(detailStartRow + 1);
    ['Game Title', 'Release Date', 'Followers', 'Publisher', 'Days to Release']
        .forEach((header, index) => releaseHeaderRow.getCell(index + 1).value = header);
    styleSummaryHeader(releaseHeaderRow, 1, 5);

    highlightedReleases.forEach((view, index) => {
        const row = worksheet.getRow(detailStartRow + 2 + index);
        row.getCell(1).value = {
            text: view.game.name,
            hyperlink: `https://steamdb.info/app/${view.appId}/`
        };
        row.getCell(1).font = { color: { argb: COLORS.blue }, underline: true };
        row.getCell(2).value = view.releaseInfo.excelValue;
        row.getCell(2).numFmt = EXCEL_DATE_FORMAT;
        row.getCell(3).value = view.latestFollowers;
        row.getCell(3).numFmt = '#,##0';
        row.getCell(4).value = view.game.publisher || '';
        row.getCell(5).value = view.daysToRelease;
    });

    worksheet.eachRow({ includeEmpty: false }, row => {
        row.eachCell({ includeEmpty: true }, cell => {
            cell.alignment = { ...cell.alignment, vertical: 'middle', wrapText: true };
        });
    });
    return worksheet;
}

async function generateWorkbook(existingData, excelFile) {
    const model = buildWorkbookModel(existingData);
    if (!model.latestCrawlDate) {
        console.log('No crawl history is available yet; Excel was not changed.');
        return false;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'SteamDB Upcoming Games Tracker';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.calcProperties.fullCalcOnLoad = true;

    addSummarySheet(workbook, model);
    addTrackerWorksheet(workbook, 'Qualified Upcoming', model.allUpcoming, COLORS.teal);

    const monthKeys = Array.from(model.releaseMonths.keys()).sort();
    monthKeys.forEach(monthKey => {
        addTrackerWorksheet(
            workbook,
            getMonthSheetName(monthKey),
            model.releaseMonths.get(monthKey),
            COLORS.blue
        );
    });
    addTrackerWorksheet(workbook, 'Unscheduled', model.unscheduled, COLORS.orange);

    try {
        await workbook.xlsx.writeFile(excelFile);
    } catch (error) {
        if (error.code === 'EBUSY' || error.code === 'EPERM') {
            throw new Error(`Close ${path.basename(excelFile)} in Excel, then run the tracker again.`);
        }
        throw error;
    }

    console.log(`Excel tracker generated with ${monthKeys.length} release-month tab(s): ${excelFile}`);
    return true;
}

async function saveData(games, options = {}) {
    const dataDir = options.dataDir || __dirname;
    const masterFile = path.join(dataDir, 'steamdb_master_data.json');
    const oldCsvFile = path.join(dataDir, 'steamdb_upcoming_tracker.csv');
    const excelFile = options.excelFile || path.join(dataDir, 'steamdb_upcoming_tracker.xlsx');
    const crawlDate = options.crawlDate || getLocalDateKey();
    const currentGames = deduplicateGames(games);
    const existingData = fs.existsSync(masterFile)
        ? readJson(masterFile, {})
        : migrateCsv(oldCsvFile);

    const crawlCanBeCommitted = currentGames.length > 0 && options.crawlComplete !== false;
    if (crawlCanBeCommitted) {
        mergeScrapedGames(existingData, currentGames, crawlDate);

        const qualifyingAppIds = Object.entries(existingData)
            .filter(([, game]) => getLatestFollowers(game) >= FOCUSED_THRESHOLD)
            .map(([appId]) => appId);

        await enrichSteamMetadata(existingData, qualifyingAppIds, crawlDate, masterFile, options);
        writeJson(masterFile, existingData);
    } else if (currentGames.length > 0) {
        console.log(
            `Crawl data was not committed because the crawl was incomplete `
            + `(${options.successfulWeeks || 0}/${options.expectedWeeks || WEEKS_TO_SCRAPE} weeks).`
        );
    } else if (!fs.existsSync(masterFile)) {
        writeJson(masterFile, existingData);
    }

    await generateWorkbook(existingData, excelFile);
}

if (require.main === module) {
    const command = process.argv.includes('--rebuild')
        ? saveData([], { crawlComplete: false, fetchMetadata: false })
        : scrapeSteamDB();

    command.catch(error => {
        console.error(error.message || error);
        process.exitCode = 1;
    });
}

module.exports = {
    buildWorkbookModel,
    generateWorkbook,
    getFollowerInterpretation,
    getLocalDateKey,
    getReleaseInfo,
    saveData
};
