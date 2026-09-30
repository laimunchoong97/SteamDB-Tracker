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

const OUTPUT_LANGUAGE = normalizeLanguage(process.env.OUTPUT_LANGUAGE);

const CRAWL_META_KEY = '_crawlMeta';

const ARCHIVE_FILE_NAME = 'steamdb_release_archives.json';
const ARCHIVE_SCHEMA_VERSION = 1;

const STEAM_API_BATCH_SIZE = 1;
const STEAM_API_BATCH_DELAY_MS = 1600;
const STEAM_API_MAX_RETRIES = 2;
const STEAM_METADATA_RECHECK_DAYS = 7;

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
    P0: { fill: 'FFFFE7A3', font: 'FF7A4B00' },
    P1: { fill: COLORS.lightBlue, font: COLORS.blue },
    P2: { fill: COLORS.lightGreen, font: COLORS.green },
    P3: { fill: COLORS.lightOrange, font: 'FF8A4B08' }
};

const COLUMN_SPECS = [
    { key: 'releaseDate', width: 15 },
    { key: 'gameTitle', width: 42 },
    { key: 'followers', width: 13 },
    { key: 'publisher', width: 28 },
    { key: 'developer', width: 28 },
    { key: 'currentOutlook', width: 22 },
    { key: 'isFree', width: 10 },
    { key: 'genres', width: 30 },
    { key: 'categories', width: 48 },
    { key: 'platforms', width: 20 },
    { key: 'appId', width: 13 },
    { key: 'status', width: 18 },
    { key: 'previousFollowers', width: 23 },
    { key: 'changeSincePrevious', width: 27 },
    { key: 'firstSeen', width: 14 },
    { key: 'lastUpdated', width: 14 }
];

const LOCALES = {
    en: {
        fontDisplay: 'Aptos Display',
        fontBody: 'Aptos',
        dateFormat: 'dd mmm yyyy',
        sheet: {
            summary: 'Summary',
            qualifiedUpcoming: 'Qualified Upcoming',
            unscheduled: 'Unscheduled'
        },
        title: 'SteamDB Release Tracker',
        subtitle: 'Games are grouped by launch month; follower history is retained across crawls.',
        columns: {
            releaseDate: 'Release Date',
            gameTitle: 'Game Title',
            followers: 'Followers',
            publisher: 'Publisher',
            developer: 'Developer',
            currentOutlook: 'Current Outlook',
            isFree: 'Is Free',
            genres: 'Genres',
            categories: 'Categories',
            platforms: 'Platforms',
            appId: 'AppID',
            status: 'Status',
            previousFollowers: 'Previous Crawl Followers',
            changeSincePrevious: 'Change Since Previous Crawl',
            firstSeen: 'First Seen',
            lastUpdated: 'Last Updated'
        },
        status: {
            upcoming: 'Upcoming',
            releasesToday: 'Releases Today',
            released: 'Released',
            releasedMonth: 'Released (Month)',
            launchMonth: 'Launch Month',
            upcomingMonth: 'Upcoming (Month)',
            unscheduled: 'Unscheduled'
        },
        outlook: {
            P0: 'P0 - AAA',
            P1: 'P1 - AA',
            P2: 'P2 - Indie',
            P3: 'P3 - Barely Viable'
        },
        summary: {
            lastCompleteCrawl: 'Last Complete Crawl',
            qualifiedUpcoming: 'Qualified Upcoming',
            newThisCrawl: 'New This Crawl',
            releasesIn30Days: 'Releases in 30 Days',
            currentReleaseMonth: 'Current Release Month',
            nextReleaseMonth: 'Next Release Month',
            unscheduled: 'Unscheduled',
            followerThreshold: 'Follower Threshold',
            releaseMonthOverview: 'Release Month Overview',
            topMovers: 'Top Movers Since Previous Crawl',
            nextReleases: 'Next 25 Releases Within 30 Days',
            releaseMonth: 'Release Month',
            qualifiedGames: 'Qualified Games',
            totalFollowers: 'Total Followers',
            averageFollowers: 'Average Followers',
            topGame: 'Top Game',
            gameTitle: 'Game Title',
            releaseDate: 'Release Date',
            followers: 'Followers',
            previous: 'Previous',
            change: 'Change',
            publisher: 'Publisher',
            daysToRelease: 'Days to Release'
        },
        yes: 'Yes',
        no: 'No',
        notAvailable: 'N/A',
        tooltips: {
            steamdb: 'Open SteamDB page',
            steam: 'Open Steam store page'
        }
    },
    'zh-CN': {
        fontDisplay: 'Microsoft YaHei',
        fontBody: 'Microsoft YaHei',
        dateFormat: 'yyyy-mm-dd',
        sheet: {
            summary: '汇总',
            qualifiedUpcoming: '达标即将发售',
            unscheduled: '未定档期'
        },
        title: 'SteamDB 发售追踪',
        subtitle: '游戏按发售月份分组，关注人数历史会跨抓取保留。',
        columns: {
            releaseDate: '发售日期',
            gameTitle: '游戏名称',
            followers: '关注人数',
            publisher: '发行商',
            developer: '开发商',
            currentOutlook: '当前评级',
            isFree: '是否免费',
            genres: '游戏类型',
            categories: '功能标签',
            platforms: '支持平台',
            appId: 'AppID',
            status: '状态',
            previousFollowers: '上次抓取关注人数',
            changeSincePrevious: '较上次抓取变化',
            firstSeen: '首次发现',
            lastUpdated: '最近更新'
        },
        status: {
            upcoming: '即将发售',
            releasesToday: '今日发售',
            released: '已发售',
            releasedMonth: '已发售（按月份）',
            launchMonth: '本月发售',
            upcomingMonth: '即将发售（按月份）',
            unscheduled: '未定档期'
        },
        outlook: {
            P0: 'P0 - AAA',
            P1: 'P1 - AA',
            P2: 'P2 - 独立游戏',
            P3: 'P3 - 勉强可行'
        },
        summary: {
            lastCompleteCrawl: '最近完整抓取',
            qualifiedUpcoming: '达标即将发售',
            newThisCrawl: '本次新增',
            releasesIn30Days: '30天内发售',
            currentReleaseMonth: '当前发售月份',
            nextReleaseMonth: '下一发售月份',
            unscheduled: '未定档期',
            followerThreshold: '关注人数阈值',
            releaseMonthOverview: '发售月份概览',
            topMovers: '较上次抓取增长榜',
            nextReleases: '30天内即将发售（前25）',
            releaseMonth: '发售月份',
            qualifiedGames: '达标游戏',
            totalFollowers: '关注人数合计',
            averageFollowers: '平均关注人数',
            topGame: '最高关注游戏',
            gameTitle: '游戏名称',
            releaseDate: '发售日期',
            followers: '关注人数',
            previous: '上次',
            change: '变化',
            publisher: '发行商',
            daysToRelease: '距发售天数'
        },
        yes: '是',
        no: '否',
        notAvailable: '暂无',
        tooltips: {
            steamdb: '打开 SteamDB 页面',
            steam: '打开 Steam 商店页面'
        }
    }
};

function normalizeLanguage(value) {
    const raw = String(value == null ? '' : value).trim().toLowerCase();
    if (raw === 'en') return 'en';
    if (raw === 'zh-cn' || raw === 'zh' || raw === 'zh_cn' || raw === 'zh-hans' || raw === 'zh-hans-cn') {
        return 'zh-CN';
    }
    return 'both';
}

function getLocale(language) {
    return LOCALES[language] || LOCALES.en;
}

function localizeOutlook(locale, canonicalOutlook) {
    if (!canonicalOutlook) return '';
    const tier = canonicalOutlook.slice(0, 2);
    return locale.outlook[tier] || canonicalOutlook;
}

function localizeStatus(locale, canonicalStatus) {
    const map = {
        'Upcoming': locale.status.upcoming,
        'Releases Today': locale.status.releasesToday,
        'Released': locale.status.released,
        'Released (Month)': locale.status.releasedMonth,
        'Launch Month': locale.status.launchMonth,
        'Upcoming (Month)': locale.status.upcomingMonth,
        'Unscheduled': locale.status.unscheduled
    };
    return map[canonicalStatus] || canonicalStatus;
}

function localizeFixedValue(locale, value) {
    if (value === undefined || value === null) return '';
    const raw = String(value);
    if (raw === 'Yes') return locale.yes;
    if (raw === 'No') return locale.no;
    if (raw === 'N/A') return locale.notAvailable;
    return value;
}

function getIsoWeek(date) {
    const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNumber = target.getUTCDay() || 7;
    target.setUTCDate(target.getUTCDate() + 4 - dayNumber);
    const isoYear = target.getUTCFullYear();
    const yearStart = new Date(Date.UTC(isoYear, 0, 1));
    const week = Math.ceil((((target - yearStart) / 86400000) + 1) / 7);
    return { isoYear, week };
}

function getCrawlMonthKeys(crawlDate = new Date()) {
    const year = crawlDate.getFullYear();
    const month = crawlDate.getMonth();
    const current = `${year}-${String(month + 1).padStart(2, '0')}`;
    const next = new Date(year, month + 1, 1);
    const nextKey = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, '0')}`;
    return [current, nextKey];
}

function getUpcomingWeeksForMonths(crawlDate = new Date()) {
    const weeks = new Set();

    for (const monthKey of getCrawlMonthKeys(crawlDate)) {
        const [year, month] = monthKey.split('-').map(Number);
        const daysInMonth = new Date(year, month, 0).getDate();
        for (let day = 1; day <= daysInMonth; day++) {
            const { isoYear, week } = getIsoWeek(new Date(year, month - 1, day));
            weeks.add(`${isoYear}W${String(week).padStart(2, '0')}`);
        }
    }

    return Array.from(weeks).sort();
}

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
    const weeksToScrape = getUpcomingWeeksForMonths();
    console.log(`Calendar-month crawl covers ${weeksToScrape.length} ISO week(s): ${weeksToScrape.join(', ')}`);
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
    Object.values(existingData || {}).forEach(game => {
        getHistoryDates(game).forEach(date => dates.add(date));
    });
    return Array.from(dates).sort().pop() || null;
}

function getLatestCompleteCrawlDate(existingData) {
    const meta = existingData && existingData[CRAWL_META_KEY];
    if (meta && /^\d{4}-\d{2}-\d{2}$/.test(meta.latestCompleteCrawlDate || '')) {
        return meta.latestCompleteCrawlDate;
    }

    const counts = new Map();
    Object.keys(existingData || {}).forEach(key => {
        const game = existingData[key];
        if (!game || typeof game !== 'object') return;
        getHistoryDates(game).forEach(date => counts.set(date, (counts.get(date) || 0) + 1));
    });

    let bestDate = null;
    let bestCount = -1;
    for (const [date, count] of counts) {
        if (count > bestCount || (count === bestCount && date > bestDate)) {
            bestDate = date;
            bestCount = count;
        }
    }
    return bestDate;
}

function getActiveMonthKeys(crawlDateKey) {
    const date = parseDateKey(crawlDateKey);
    if (!date) return [];
    const year = date.getUTCFullYear();
    const month = date.getUTCMonth();
    const current = `${year}-${String(month + 1).padStart(2, '0')}`;
    const next = new Date(Date.UTC(year, month + 1, 1));
    const nextKey = `${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, '0')}`;
    return [current, nextKey];
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

function getCurrentOutlook(followers) {
    if (followers < MINIMUM_FOLLOWER_THRESHOLD) {
        return '';
    }
    if (followers >= 30000) {
        return 'P0 - AAA';
    }
    if (followers >= 10000) {
        return 'P1 - AA';
    }
    if (followers >= 3000) {
        return 'P2 - Indie';
    }
    return 'P3 - Barely Viable';
}

function buildGameView(appId, game, referenceDateKey) {
    const historyDates = getHistoryDates(game);
    const latestDate = historyDates[historyDates.length - 1] || null;
    const previousDate = historyDates.length > 1 ? historyDates[historyDates.length - 2] : null;
    const latestFollowers = latestDate ? Number(game.history[latestDate]) || 0 : 0;
    const previousFollowers = previousDate ? Number(game.history[previousDate]) || 0 : null;
    const releaseInfo = getReleaseInfo(game, referenceDateKey || latestDate);
    const currentOutlook = getCurrentOutlook(latestFollowers);

    return {
        appId: String(appId),
        game,
        historyDates,
        latestDate,
        latestFollowers,
        previousFollowers,
        changeSincePrevious: previousFollowers === null ? null : latestFollowers - previousFollowers,
        currentOutlook,
        firstSeen: game.firstSeen || historyDates[0] || null,
        releaseInfo,
        status: getReleaseStatus(releaseInfo, referenceDateKey),
        daysToRelease: getDaysToRelease(releaseInfo.exactDate, referenceDateKey)
    };
}

// --- Release archives -------------------------------------------------------

function emptyArchives() {
    return { schemaVersion: ARCHIVE_SCHEMA_VERSION, updatedAt: null, crawlDate: null, months: {} };
}

function loadReleaseArchives(dataDir) {
    const file = path.join(dataDir || __dirname, ARCHIVE_FILE_NAME);
    const parsed = readJson(file, null);
    if (!parsed || typeof parsed !== 'object' || typeof parsed.months !== 'object' || parsed.months === null) {
        return emptyArchives();
    }
    return {
        schemaVersion: Number(parsed.schemaVersion) || ARCHIVE_SCHEMA_VERSION,
        updatedAt: parsed.updatedAt || null,
        crawlDate: parsed.crawlDate || null,
        months: parsed.months
    };
}

function saveReleaseArchives(dataDir, archives) {
    const file = path.join(dataDir || __dirname, ARCHIVE_FILE_NAME);
    writeJson(file, archives);
}

function serializeGameView(view) {
    const excelValueIsDate = view.releaseInfo.excelValue instanceof Date;
    return {
        appId: view.appId,
        name: view.game.name || '',
        releaseDate: view.game.releaseDate || '',
        releaseTimestamp: view.game.releaseTimestamp || null,
        publisher: view.game.publisher || '',
        developer: view.game.developer || '',
        isFree: view.game.isFree || '',
        genres: view.game.genres || '',
        categories: view.game.categories || '',
        platforms: view.game.platforms || '',
        history: { ...(view.game.history || {}) },
        releaseInfoMonthKey: view.releaseInfo.monthKey,
        releaseInfoExcelValue: excelValueIsDate ? view.releaseInfo.excelValue.toISOString() : view.releaseInfo.excelValue,
        releaseInfoExcelValueIsDate: excelValueIsDate,
        releaseInfoExactDate: view.releaseInfo.exactDate ? view.releaseInfo.exactDate.toISOString() : null,
        releaseInfoSortValue: view.releaseInfo.sortValue,
        latestDate: view.latestDate,
        latestFollowers: view.latestFollowers,
        previousFollowers: view.previousFollowers,
        changeSincePrevious: view.changeSincePrevious,
        currentOutlook: view.currentOutlook,
        status: view.status,
        firstSeen: view.firstSeen,
        daysToRelease: view.daysToRelease
    };
}

function deserializeGameView(record) {
    const game = {
        name: record.name || '',
        releaseDate: record.releaseDate || '',
        releaseTimestamp: record.releaseTimestamp || null,
        publisher: record.publisher || '',
        developer: record.developer || '',
        isFree: record.isFree || '',
        genres: record.genres || '',
        categories: record.categories || '',
        platforms: record.platforms || '',
        history: { ...(record.history || {}) }
    };
    if (record.firstSeen) game.firstSeen = record.firstSeen;

    const excelValue = record.releaseInfoExcelValueIsDate && record.releaseInfoExcelValue
        ? new Date(record.releaseInfoExcelValue)
        : record.releaseInfoExcelValue;

    return {
        appId: record.appId,
        game,
        historyDates: getHistoryDates(game),
        latestDate: record.latestDate || null,
        latestFollowers: Number(record.latestFollowers) || 0,
        previousFollowers: record.previousFollowers === undefined ? null : record.previousFollowers,
        changeSincePrevious: record.changeSincePrevious === undefined ? null : record.changeSincePrevious,
        currentOutlook: record.currentOutlook || '',
        firstSeen: record.firstSeen || null,
        releaseInfo: {
            monthKey: record.releaseInfoMonthKey || null,
            excelValue,
            exactDate: record.releaseInfoExactDate ? new Date(record.releaseInfoExactDate) : null,
            sortValue: record.releaseInfoSortValue === undefined ? Number.MAX_SAFE_INTEGER : record.releaseInfoSortValue
        },
        status: record.status || 'Unscheduled',
        daysToRelease: record.daysToRelease === undefined ? null : record.daysToRelease
    };
}

function serializeArchiveMonth(monthKey, views, crawlDate) {
    const sorted = sortGameViews(views);
    return {
        monthKey,
        archivedAt: crawlDate,
        historyDates: getSheetHistoryDates(sorted),
        games: sorted.map(serializeGameView)
    };
}

function deserializeArchiveMonth(entry) {
    const historyDates = Array.isArray(entry.historyDates) ? [...entry.historyDates].sort() : [];
    return {
        monthKey: entry.monthKey,
        archivedAt: entry.archivedAt || null,
        historyDates,
        views: (entry.games || []).map(deserializeGameView)
    };
}

function trimGameHistory(game, beforeMonthKey) {
    const history = {};
    const raw = (game && game.history) || {};
    for (const [date, value] of Object.entries(raw)) {
        if (/^\d{4}-\d{2}-\d{2}$/.test(date) && date.slice(0, 7) < beforeMonthKey) {
            history[date] = value;
        }
    }
    return { ...game, history };
}

function buildWorkbookModel(existingData, options = {}) {
    const derivedCrawlDate = options.crawlDate || getLatestCompleteCrawlDate(existingData);
    const referenceDateKey = derivedCrawlDate || getLatestCrawlDate(existingData);
    const crawlDate = derivedCrawlDate || referenceDateKey;
    const archives = options.archives || emptyArchives();
    const archiveMonths = archives.months || {};
    const historyBeforeMonthKey = options.historyBeforeMonthKey || null;

    const views = Object.entries(existingData || {})
        .filter(([key, game]) => key !== CRAWL_META_KEY && game && typeof game === 'object')
        .map(([appId, game]) => [appId, historyBeforeMonthKey ? trimGameHistory(game, historyBeforeMonthKey) : game])
        .filter(([, game]) => getHistoryDates(game).length > 0)
        .map(([appId, game]) => buildGameView(appId, game, referenceDateKey))
        .filter(view => view.latestFollowers >= FOCUSED_THRESHOLD);

    const activeMonths = getActiveMonthKeys(crawlDate);
    const currentMonthKey = activeMonths[0] || null;
    const nextMonthKey = activeMonths[1] || null;

    const activeMonthViews = [];
    const unscheduledCandidates = [];
    const livePastMonths = new Map();
    const hiddenFuture = new Set();

    for (const view of views) {
        const monthKey = view.releaseInfo.monthKey;
        if (!monthKey) {
            unscheduledCandidates.push(view);
            continue;
        }
        if (currentMonthKey && monthKey < currentMonthKey) {
            if (!livePastMonths.has(monthKey)) livePastMonths.set(monthKey, []);
            livePastMonths.get(monthKey).push(view);
        } else if (activeMonths.includes(monthKey)) {
            activeMonthViews.push(view);
        } else if (nextMonthKey && monthKey > nextMonthKey) {
            hiddenFuture.add(monthKey);
        }
    }

    const isFromLatestCrawl = view => Boolean(crawlDate) && view.historyDates.includes(crawlDate);
    const activeViews = activeMonthViews.filter(isFromLatestCrawl);
    const unscheduled = unscheduledCandidates.filter(isFromLatestCrawl);

    const activeReleaseMonths = new Map();
    activeMonths.forEach(monthKey => activeReleaseMonths.set(monthKey, []));
    activeViews.forEach(view => {
        if (activeReleaseMonths.has(view.releaseInfo.monthKey)) {
            activeReleaseMonths.get(view.releaseInfo.monthKey).push(view);
        }
    });

    const activeByMonth = new Map();
    activeMonthViews.forEach(view => {
        const monthKey = view.releaseInfo.monthKey;
        if (!activeByMonth.has(monthKey)) activeByMonth.set(monthKey, []);
        activeByMonth.get(monthKey).push(view);
    });

    const sectionKeys = new Set();
    Object.keys(archiveMonths).forEach(monthKey => {
        if (currentMonthKey && monthKey < currentMonthKey) sectionKeys.add(monthKey);
    });
    activeMonths.forEach(monthKey => {
        if (monthKey) sectionKeys.add(monthKey);
    });

    const sections = Array.from(sectionKeys).sort().map(monthKey => {
        if (archiveMonths[monthKey]) {
            const archived = deserializeArchiveMonth(archiveMonths[monthKey]);
            return {
                monthKey,
                views: archived.views,
                historyDates: archived.historyDates,
                archived: true,
                archivedAt: archived.archivedAt
            };
        }
        return {
            monthKey,
            views: sortGameViews(activeByMonth.get(monthKey) || []),
            historyDates: null,
            archived: false,
            archivedAt: null
        };
    });

    return {
        latestCrawlDate: crawlDate,
        currentMonthKey,
        activeMonths,
        views,
        activeViews,
        allUpcoming: activeViews,
        activeMonthViews,
        releaseMonths: activeReleaseMonths,
        unscheduled,
        livePastMonths,
        hiddenFutureMonths: Array.from(hiddenFuture).sort(),
        sections
    };
}

function updateReleaseArchives(existingData, options = {}) {
    const dataDir = options.dataDir || __dirname;
    const archives = options.archives || loadReleaseArchives(dataDir);
    const crawlDate = options.crawlDate || getLatestCompleteCrawlDate(existingData);
    if (!crawlDate) return archives;

    const currentMonthKey = getActiveMonthKeys(crawlDate)[0] || null;
    const model = buildWorkbookModel(existingData, {
        crawlDate,
        archives: emptyArchives(),
        historyBeforeMonthKey: currentMonthKey
    });
    const months = { ...(archives.months || {}) };
    let changed = false;

    for (const [monthKey, monthViews] of model.livePastMonths) {
        if (months[monthKey]) continue;
        if (!monthViews || monthViews.length === 0) continue;
        months[monthKey] = serializeArchiveMonth(monthKey, monthViews, crawlDate);
        changed = true;
    }

    if (!changed) return archives;

    const updated = {
        schemaVersion: ARCHIVE_SCHEMA_VERSION,
        updatedAt: crawlDate,
        crawlDate,
        months
    };
    saveReleaseArchives(dataDir, updated);
    return updated;
}

function getMonthSheetName(monthKey, language = 'en') {
    const match = /^(\d{4})-(\d{2})$/.exec(monthKey || '');
    if (!match) return String(monthKey || '').slice(0, 31);
    if (language === 'zh-CN') {
        return `${match[1]}年${Number(match[2])}月`;
    }
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

function getFixedColumns(locale) {
    return COLUMN_SPECS.map(spec => ({
        header: locale.columns[spec.key],
        key: spec.key,
        width: spec.width
    }));
}

function buildTrackerRow(view, historyDates, locale) {
    const row = {
        releaseDate: view.releaseInfo.excelValue,
        gameTitle: view.game.name || '',
        followers: view.latestFollowers,
        publisher: localizeFixedValue(locale, view.game.publisher),
        developer: localizeFixedValue(locale, view.game.developer),
        currentOutlook: localizeOutlook(locale, view.currentOutlook),
        isFree: localizeFixedValue(locale, view.game.isFree),
        genres: localizeFixedValue(locale, view.game.genres),
        categories: localizeFixedValue(locale, view.game.categories),
        platforms: localizeFixedValue(locale, normalizePlatforms(view.game.platforms)),
        appId: Number(view.appId),
        status: localizeStatus(locale, view.status),
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

function styleTrackerWorksheet(worksheet, rowCount, historyDates, locale) {
    const fixedColumns = getFixedColumns(locale);
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
        to: { row: 1, column: fixedColumns.length + historyDates.length }
    };
    worksheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    worksheet.properties.defaultRowHeight = 22;

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell(cell => {
        cell.font = { name: locale.fontDisplay, size: 11, bold: true, color: { argb: COLORS.white } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = { bottom: { style: 'medium', color: { argb: COLORS.teal } } };
    });

    for (let rowIndex = 2; rowIndex <= rowCount + 1; rowIndex++) {
        const row = worksheet.getRow(rowIndex);
        row.height = 45;
        row.eachCell({ includeEmpty: true }, cell => {
            cell.font = { name: locale.fontBody, size: 10, color: { argb: COLORS.navy } };
            cell.alignment = { vertical: 'top' };
            cell.border = { bottom: { style: 'hair', color: { argb: COLORS.border } } };
            if (rowIndex % 2 === 0) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
            }
        });

        ['B', 'D', 'E', 'F', 'H', 'I', 'J'].forEach(column => {
            row.getCell(column).alignment = { vertical: 'top', wrapText: true };
        });
        row.getCell('A').numFmt = locale.dateFormat;
        row.getCell('C').numFmt = '#,##0';
        row.getCell('M').numFmt = '#,##0';
        row.getCell('N').numFmt = '+#,##0;-#,##0;0';
        row.getCell('O').numFmt = locale.dateFormat;
        row.getCell('P').numFmt = locale.dateFormat;

        for (let columnIndex = fixedColumns.length + 1;
            columnIndex <= fixedColumns.length + historyDates.length;
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
        ref: `N2:N${lastRow}`,
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

function addTrackerWorksheet(workbook, name, views, tabColor, locale, historyDatesOverride = null) {
    const sortedViews = sortGameViews(views);
    const historyDates = historyDatesOverride
        ? [...historyDatesOverride].sort()
        : getSheetHistoryDates(sortedViews);
    const worksheet = workbook.addWorksheet(name, {
        properties: { tabColor: { argb: tabColor } }
    });

    worksheet.columns = [
        ...getFixedColumns(locale),
        ...historyDates.map(date => ({ header: date, key: historyKey(date), width: 13 }))
    ];

    sortedViews.forEach(view => {
        worksheet.addRow(buildTrackerRow(view, historyDates, locale));
    });

    styleTrackerWorksheet(worksheet, sortedViews.length, historyDates, locale);

    sortedViews.forEach((view, index) => {
        const row = worksheet.getRow(index + 2);
        const tierColor = TIER_STYLES[view.currentOutlook.slice(0, 2)];
        if (tierColor) {
            row.getCell('currentOutlook').fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: tierColor.fill }
            };
            row.getCell('currentOutlook').font = {
                name: locale.fontBody,
                size: 10,
                bold: true,
                color: { argb: tierColor.font }
            };
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
            tooltip: locale.tooltips.steamdb
        };
        titleCell.font = { name: locale.fontBody, size: 10, color: { argb: COLORS.blue }, underline: true };

        const appIdCell = row.getCell('appId');
        appIdCell.value = {
            text: view.appId,
            hyperlink: `https://store.steampowered.com/app/${view.appId}/`,
            tooltip: locale.tooltips.steam
        };
        appIdCell.font = { name: locale.fontBody, size: 10, color: { argb: COLORS.blue }, underline: true };
    });
    return worksheet;
}

function styleSectionTitle(worksheet, range, title, locale) {
    worksheet.mergeCells(range);
    const cell = worksheet.getCell(range.split(':')[0]);
    cell.value = title;
    cell.font = { name: locale.fontDisplay, size: 12, bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.teal } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
}

function styleSummaryHeader(row, startColumn, endColumn, locale) {
    for (let column = startColumn; column <= endColumn; column++) {
        const cell = row.getCell(column);
        cell.font = { name: locale.fontBody, size: 10, bold: true, color: { argb: COLORS.white } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
}

function addMetricCard(worksheet, startColumn, labelRow, label, value, fillColor, locale) {
    worksheet.mergeCells(labelRow, startColumn, labelRow, startColumn + 1);
    worksheet.mergeCells(labelRow + 1, startColumn, labelRow + 2, startColumn + 1);

    const labelCell = worksheet.getCell(labelRow, startColumn);
    labelCell.value = label;
    labelCell.font = { name: locale.fontBody, size: 9, bold: true, color: { argb: COLORS.grey } };
    labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGrey } };
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };

    const valueCell = worksheet.getCell(labelRow + 1, startColumn);
    valueCell.value = value;
    valueCell.font = { name: locale.fontDisplay, size: 18, bold: true, color: { argb: COLORS.navy } };
    valueCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
    valueCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    if (value instanceof Date) valueCell.numFmt = locale.dateFormat;
    if (typeof value === 'number') valueCell.numFmt = '#,##0';
}

function addSummarySheet(workbook, model, language) {
    const locale = getLocale(language);
    const worksheet = workbook.addWorksheet(locale.sheet.summary, {
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
    worksheet.getCell('A1').value = locale.title;
    worksheet.getCell('A1').font = { name: locale.fontDisplay, size: 24, bold: true, color: { argb: COLORS.white } };
    worksheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getRow(1).height = 40;

    worksheet.mergeCells('A2:L2');
    worksheet.getCell('A2').value = locale.subtitle;
    worksheet.getCell('A2').font = { name: locale.fontBody, size: 10, italic: true, color: { argb: COLORS.grey } };
    worksheet.getCell('A2').alignment = { vertical: 'middle' };

    const newThisCrawl = model.activeViews.filter(view => view.firstSeen === model.latestCrawlDate).length;
    const nextThirtyDays = model.activeViews.filter(view => (
        view.daysToRelease !== null && view.daysToRelease >= 0 && view.daysToRelease <= 30
    ));
    const highlightedReleases = sortGameViews(nextThirtyDays).slice(0, 25);
    const metrics = [
        [locale.summary.lastCompleteCrawl, parseDateKey(model.latestCrawlDate), COLORS.lightBlue],
        [locale.summary.qualifiedUpcoming, model.activeViews.length, COLORS.lightTeal],
        [locale.summary.newThisCrawl, newThisCrawl, COLORS.lightOrange],
        [locale.summary.releasesIn30Days, nextThirtyDays.length, COLORS.lightGreen],
        [locale.summary.currentReleaseMonth, getMonthSheetName(model.activeMonths[0], language), COLORS.lightBlue],
        [locale.summary.nextReleaseMonth, getMonthSheetName(model.activeMonths[1], language), COLORS.lightBlue],
        [locale.summary.unscheduled, model.unscheduled.length, COLORS.lightOrange],
        [locale.summary.followerThreshold, FOCUSED_THRESHOLD, COLORS.lightGreen]
    ];

    metrics.slice(0, 4).forEach((metric, index) => {
        addMetricCard(worksheet, 1 + (index * 3), 4, metric[0], metric[1], metric[2], locale);
    });
    metrics.slice(4).forEach((metric, index) => {
        addMetricCard(worksheet, 1 + (index * 3), 8, metric[0], metric[1], metric[2], locale);
    });

    const monthStartRow = 13;
    styleSectionTitle(worksheet, `A${monthStartRow}:E${monthStartRow}`, locale.summary.releaseMonthOverview, locale);
    const monthHeaderRow = worksheet.getRow(monthStartRow + 1);
    [
        locale.summary.releaseMonth,
        locale.summary.qualifiedGames,
        locale.summary.totalFollowers,
        locale.summary.averageFollowers,
        locale.summary.topGame
    ].forEach((header, index) => monthHeaderRow.getCell(index + 1).value = header);
    styleSummaryHeader(monthHeaderRow, 1, 5, locale);

    const monthKeys = Array.from(model.releaseMonths.keys()).sort();
    monthKeys.forEach((monthKey, index) => {
        const views = model.releaseMonths.get(monthKey);
        const totalFollowers = views.reduce((sum, view) => sum + view.latestFollowers, 0);
        const topGame = [...views].sort((left, right) => right.latestFollowers - left.latestFollowers)[0];
        const row = worksheet.getRow(monthStartRow + 2 + index);
        row.values = [
            getMonthSheetName(monthKey, language),
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

    styleSectionTitle(worksheet, `G${monthStartRow}:L${monthStartRow}`, locale.summary.topMovers, locale);
    const moverHeaderRow = worksheet.getRow(monthStartRow + 1);
    [
        locale.summary.gameTitle,
        locale.summary.releaseDate,
        locale.summary.followers,
        locale.summary.previous,
        locale.summary.change,
        locale.summary.releaseMonth
    ].forEach((header, index) => moverHeaderRow.getCell(index + 7).value = header);
    styleSummaryHeader(moverHeaderRow, 7, 12, locale);

    const topMovers = [...model.activeViews]
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
        row.getCell(8).numFmt = locale.dateFormat;
        row.getCell(9).value = view.latestFollowers;
        row.getCell(10).value = view.previousFollowers;
        row.getCell(11).value = view.changeSincePrevious;
        row.getCell(12).value = view.releaseInfo.monthKey
            ? getMonthSheetName(view.releaseInfo.monthKey, language)
            : locale.sheet.unscheduled;
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
    styleSectionTitle(worksheet, `A${detailStartRow}:E${detailStartRow}`, locale.summary.nextReleases, locale);
    const releaseHeaderRow = worksheet.getRow(detailStartRow + 1);
    [
        locale.summary.gameTitle,
        locale.summary.releaseDate,
        locale.summary.followers,
        locale.summary.publisher,
        locale.summary.daysToRelease
    ].forEach((header, index) => releaseHeaderRow.getCell(index + 1).value = header);
    styleSummaryHeader(releaseHeaderRow, 1, 5, locale);

    highlightedReleases.forEach((view, index) => {
        const row = worksheet.getRow(detailStartRow + 2 + index);
        row.getCell(1).value = {
            text: view.game.name,
            hyperlink: `https://steamdb.info/app/${view.appId}/`
        };
        row.getCell(1).font = { color: { argb: COLORS.blue }, underline: true };
        row.getCell(2).value = view.releaseInfo.excelValue;
        row.getCell(2).numFmt = locale.dateFormat;
        row.getCell(3).value = view.latestFollowers;
        row.getCell(3).numFmt = '#,##0';
        row.getCell(4).value = localizeFixedValue(locale, view.game.publisher);
        row.getCell(5).value = view.daysToRelease;
    });

    worksheet.eachRow({ includeEmpty: false }, row => {
        row.eachCell({ includeEmpty: true }, cell => {
            cell.alignment = { ...cell.alignment, vertical: 'middle', wrapText: true };
        });
    });
    return worksheet;
}

function buildWorkbook(model, language) {
    const locale = getLocale(language);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'SteamDB Upcoming Games Tracker';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.calcProperties.fullCalcOnLoad = true;

    addSummarySheet(workbook, model, language);
    addTrackerWorksheet(workbook, locale.sheet.qualifiedUpcoming, model.activeViews, COLORS.teal, locale);

    model.sections.forEach(section => {
        addTrackerWorksheet(
            workbook,
            getMonthSheetName(section.monthKey, language),
            section.views,
            COLORS.blue,
            locale,
            section.historyDates
        );
    });
    addTrackerWorksheet(workbook, locale.sheet.unscheduled, model.unscheduled, COLORS.orange, locale);

    return workbook;
}

function getOutputFilename(language, dataDir) {
    const base = 'steamdb_upcoming_tracker';
    const name = language === 'en' ? `${base}.xlsx` : `${base}_${language}.xlsx`;
    return path.join(dataDir || __dirname, name);
}

function resolveOutputTargets(options = {}) {
    const explicitFile = options.excelFile;
    const hasLanguage = Object.prototype.hasOwnProperty.call(options, 'language')
        && options.language !== null
        && options.language !== undefined
        && options.language !== '';

    if (explicitFile) {
        if (!hasLanguage) {
            return [{ language: 'en', file: explicitFile }];
        }
        const language = normalizeLanguage(options.language);
        if (language === 'both') {
            return [
                { language: 'en', file: explicitFile },
                { language: 'zh-CN', file: explicitFile.replace(/\.xlsx$/i, '_zh-CN.xlsx') }
            ];
        }
        return [{ language, file: explicitFile }];
    }

    const language = hasLanguage ? normalizeLanguage(options.language) : OUTPUT_LANGUAGE;
    const dataDir = options.dataDir || __dirname;
    if (language === 'both') {
        return [
            { language: 'en', file: getOutputFilename('en', dataDir) },
            { language: 'zh-CN', file: getOutputFilename('zh-CN', dataDir) }
        ];
    }
    return [{ language, file: getOutputFilename(language, dataDir) }];
}

async function writeWorkbookFile(workbook, excelFile) {
    try {
        await workbook.xlsx.writeFile(excelFile);
    } catch (error) {
        if (error.code === 'EBUSY' || error.code === 'EPERM') {
            throw new Error(`Close ${path.basename(excelFile)} in Excel, then run the tracker again.`);
        }
        throw error;
    }
}

async function generateWorkbook(existingData, optionsOrFile) {
    const options = typeof optionsOrFile === 'string'
        ? { excelFile: optionsOrFile }
        : (optionsOrFile || {});
    const dataDir = options.dataDir || __dirname;
    const archives = options.archives || loadReleaseArchives(dataDir);
    const model = buildWorkbookModel(existingData, {
        crawlDate: options.crawlDate,
        archives
    });

    if (!model.latestCrawlDate) {
        console.log('No crawl history is available yet; Excel was not changed.');
        return false;
    }

    const targets = resolveOutputTargets(options);
    const workbooks = new Map();

    for (const target of targets) {
        if (!workbooks.has(target.language)) {
            workbooks.set(target.language, buildWorkbook(model, target.language));
        }
        await writeWorkbookFile(workbooks.get(target.language), target.file);
    }

    const languages = targets.map(target => target.language).join(', ');
    console.log(
        `Excel tracker generated for ${languages} with ${model.sections.length} release-month tab(s): `
        + targets.map(target => target.file).join(', ')
    );
    return true;
}

async function saveData(games, options = {}) {
    const dataDir = options.dataDir || __dirname;
    const masterFile = path.join(dataDir, 'steamdb_master_data.json');
    const oldCsvFile = path.join(dataDir, 'steamdb_upcoming_tracker.csv');
    const crawlDate = options.crawlDate || getLocalDateKey();
    const currentGames = deduplicateGames(games);
    const existingData = fs.existsSync(masterFile)
        ? readJson(masterFile, {})
        : migrateCsv(oldCsvFile);

    const crawlCanBeCommitted = currentGames.length > 0 && options.crawlComplete !== false;
    let archives = loadReleaseArchives(dataDir);

    if (crawlCanBeCommitted) {
        // Freeze past months from the pre-merge state so a new (possibly
        // spillover) crawl can never contaminate an already completed month.
        archives = updateReleaseArchives(existingData, { dataDir, archives, crawlDate });

        mergeScrapedGames(existingData, currentGames, crawlDate);
        existingData[CRAWL_META_KEY] = {
            latestCompleteCrawlDate: crawlDate,
            expectedWeeks: options.expectedWeeks || null,
            successfulWeeks: options.successfulWeeks || null,
            updatedAt: crawlDate
        };

        const qualifyingAppIds = Object.entries(existingData)
            .filter(([appId, game]) => appId !== CRAWL_META_KEY && getLatestFollowers(game) >= FOCUSED_THRESHOLD)
            .map(([appId]) => appId);

        await enrichSteamMetadata(existingData, qualifyingAppIds, crawlDate, masterFile, options);
        writeJson(masterFile, existingData);
    } else {
        if (currentGames.length > 0) {
            console.log(
                `Crawl data was not committed because the crawl was incomplete `
                + `(${options.successfulWeeks || 0}/${options.expectedWeeks || '?'} weeks).`
            );
        } else if (!fs.existsSync(masterFile)) {
            writeJson(masterFile, existingData);
        }
        archives = updateReleaseArchives(existingData, { dataDir, archives });
    }

    const committedCrawlDate = crawlCanBeCommitted ? crawlDate : undefined;
    await generateWorkbook(existingData, {
        ...options,
        crawlDate: committedCrawlDate,
        dataDir,
        archives
    });
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
    COLUMN_SPECS,
    LOCALES,
    buildWorkbookModel,
    generateWorkbook,
    getActiveMonthKeys,
    getCurrentOutlook,
    getLatestCompleteCrawlDate,
    getLocalDateKey,
    getReleaseInfo,
    getUpcomingWeeksForMonths,
    loadReleaseArchives,
    normalizeLanguage,
    resolveOutputTargets,
    saveData,
    updateReleaseArchives
};
