const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

process.env.FOCUSED_THRESHOLD = '500';

const {
    assertCrawlProducedData,
    buildWorkbookModel,
    enrichSteamMetadata,
    formatDuration,
    generateWorkbook,
    getActiveMonthKeys,
    getCurrentOutlook,
    getLatestCompleteCrawlDate,
    getLocalDateKey,
    getReleaseInfo,
    getUpcomingWeeksForMonths,
    normalizeLanguage,
    resolveOutputTargets,
    saveData
} = require('../steamdb_tracker');

test('rejects crawls where no week loaded or no rows were extracted', () => {
    assert.throws(
        () => assertCrawlProducedData(0, 9, 0),
        /none of the 9 week\(s\) loaded/
    );
    assert.throws(
        () => assertCrawlProducedData(9, 9, 0),
        /contained no game rows/
    );
    assert.doesNotThrow(() => assertCrawlProducedData(8, 9, 120));
});

test('reports metadata progress and completion totals', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-metadata-test-'));
    const masterFile = path.join(tempDirectory, 'master.json');
    const messages = [];
    const errors = [];
    const data = {
        10: { name: 'Alpha' },
        20: { name: 'Beta' }
    };
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));

    await enrichSteamMetadata(data, ['10', '20'], '2026-10-02', masterFile, {
        fetchAppDetails: async appId => appId === '10'
            ? {
                developers: ['Developer'],
                publishers: ['Publisher'],
                is_free: false,
                genres: [],
                categories: [],
                platforms: { windows: true }
            }
            : null,
        metadataConcurrency: 2,
        metadataRequestSpacingMs: 0,
        log: message => messages.push(message),
        logError: message => errors.push(message)
    });

    assert.equal(data[10].metadataStatus, 'complete');
    assert.equal(data[20].metadataStatus, 'unavailable');
    assert.equal(errors.length, 0);
    assert(messages.some(message => message.includes('[Metadata 1/2] Fetching Alpha')));
    assert(messages.some(message => message.includes('[Metadata 2/2] Finished')));
    assert(messages.some(message => message.includes('1 complete, 1 unavailable, 0 failed')));
    assert.equal(formatDuration(15000), '15s');
    assert.equal(formatDuration(65000), '1m 5s');
});

test('enriches metadata with the configured worker concurrency', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-concurrency-test-'));
    const masterFile = path.join(tempDirectory, 'master.json');
    const data = Object.fromEntries(
        Array.from({ length: 6 }, (_, index) => [String(index + 1), { name: `Game ${index + 1}` }])
    );
    let inFlight = 0;
    let maxInFlight = 0;
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));

    await enrichSteamMetadata(data, Object.keys(data), '2026-10-02', masterFile, {
        fetchAppDetails: async () => {
            inFlight += 1;
            maxInFlight = Math.max(maxInFlight, inFlight);
            await new Promise(resolve => setTimeout(resolve, 10));
            inFlight -= 1;
            return null;
        },
        metadataConcurrency: 3,
        metadataRequestSpacingMs: 0,
        log: () => {},
        logError: () => {}
    });

    assert.equal(maxInFlight, 3);
});

test('enforces the 1,000 follower floor and retains earlier history', () => {
    const model = buildWorkbookModel({
        100: {
            name: 'Hidden Game',
            releaseDate: 'Oct 2026',
            history: { '2026-09-29': 500, '2026-09-30': 999 }
        },
        200: {
            name: 'Qualified Game',
            releaseDate: 'Oct 2026',
            history: { '2026-09-28': 500, '2026-09-29': 900, '2026-09-30': 1000 }
        }
    });

    assert.deepEqual(model.views.map(view => view.appId), ['200']);
    assert.deepEqual(model.views[0].historyDates, ['2026-09-28', '2026-09-29', '2026-09-30']);
});

test('uses exact Current Outlook boundaries', () => {
    assert.equal(getCurrentOutlook(999), '');
    assert.equal(getCurrentOutlook(1000), 'P3 - Barely Viable');
    assert.equal(getCurrentOutlook(2999), 'P3 - Barely Viable');
    assert.equal(getCurrentOutlook(3000), 'P2 - Indie');
    assert.equal(getCurrentOutlook(9999), 'P2 - Indie');
    assert.equal(getCurrentOutlook(10000), 'P1 - AA');
    assert.equal(getCurrentOutlook(29999), 'P1 - AA');
    assert.equal(getCurrentOutlook(30000), 'P0 - AAA');
});

test('keeps month-boundary dates at UTC midnight', () => {
    const release = getReleaseInfo({
        releaseDate: '01 Oct',
        releaseTimestamp: Date.UTC(2026, 9, 1) / 1000
    }, '2026-09-30');

    assert.equal(release.monthKey, '2026-10');
    assert.equal(release.exactDate.toISOString(), '2026-10-01T00:00:00.000Z');
});

test('writes Current Outlook without a Days to Release detail column', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-tracker-test-'));
    const outputFile = path.join(tempDirectory, 'tracker.xlsx');
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));

    await generateWorkbook({
        300: {
            name: 'Tier Three Game',
            releaseDate: '01 Oct',
            releaseTimestamp: Date.UTC(2026, 9, 1) / 1000,
            publisher: 'Publisher',
            developer: 'Developer',
            history: { '2026-09-29': 2500, '2026-09-30': 3000 }
        }
    }, { excelFile: outputFile, dataDir: tempDirectory });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(outputFile);
    assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), [
        'Summary', 'Qualified Upcoming', 'Sep 2026', 'Oct 2026', 'Unscheduled'
    ]);

    const worksheet = workbook.getWorksheet('Oct 2026');
    const headers = worksheet.getRow(1).values.slice(1);
    assert.equal(headers.includes('Days to Release'), false);
    assert.deepEqual(headers.slice(0, 6), [
        'Release Date', 'Game Title', 'Followers', 'Publisher', 'Developer',
        'Current Outlook'
    ]);

    assert.equal(worksheet.getCell('A2').value.toISOString(), '2026-10-01T00:00:00.000Z');
    assert.equal(worksheet.getCell('F2').value, 'P2 - Indie');
    assert.equal(worksheet.views[0].state, 'frozen');
    assert(worksheet.autoFilter);
});

function makeMaster(metaDate, games) {
    const master = { _crawlMeta: { latestCompleteCrawlDate: metaDate } };
    for (const game of games) {
        const { appId, followers, ...fields } = game;
        master[String(appId)] = {
            ...fields,
            releaseTimestamp: game.releaseTimestamp || null,
            history: { [metaDate]: followers }
        };
    }
    return master;
}

test('generates every ISO week for two full calendar months including spillover', () => {
    assert.deepEqual(getUpcomingWeeksForMonths(new Date(2026, 8, 15)), [
        '2026W36', '2026W37', '2026W38', '2026W39', '2026W40',
        '2026W41', '2026W42', '2026W43', '2026W44'
    ]);

    assert.deepEqual(getUpcomingWeeksForMonths(new Date(2026, 11, 10)), [
        '2026W49', '2026W50', '2026W51', '2026W52', '2026W53',
        '2027W01', '2027W02', '2027W03', '2027W04'
    ]);
});

test('resolves the two active release months from the crawl date', () => {
    assert.deepEqual(getActiveMonthKeys('2026-09-30'), ['2026-09', '2026-10']);
    assert.deepEqual(getActiveMonthKeys('2026-12-31'), ['2026-12', '2027-01']);
});

test('only current and next release months are active; future hidden; unscheduled visible', () => {
    const model = buildWorkbookModel(makeMaster('2026-09-30', [
        { appId: 1, name: 'Sep Game', releaseDate: '15 Sep 2026', followers: 5000 },
        { appId: 2, name: 'Oct Game', releaseDate: '10 Oct 2026', followers: 4000 },
        { appId: 3, name: 'Nov Game', releaseDate: '10 Nov 2026', followers: 3500 },
        { appId: 4, name: 'Unscheduled Game', releaseDate: 'TBA', followers: 3000 }
    ]));

    assert.equal(model.currentMonthKey, '2026-09');
    assert.deepEqual(model.activeMonths, ['2026-09', '2026-10']);
    assert.deepEqual(model.activeViews.map(view => view.appId).sort(), ['1', '2']);
    assert.deepEqual(model.allUpcoming.map(view => view.appId).sort(), ['1', '2']);
    assert.deepEqual(model.hiddenFutureMonths, ['2026-11']);
    assert.deepEqual(model.unscheduled.map(view => view.appId), ['4']);
    assert.deepEqual(model.sections.map(section => section.monthKey), ['2026-09', '2026-10']);
});

test('falls back to the broadest crawl date when crawl metadata is absent', () => {
    const legacy = {
        1: { name: 'A', releaseDate: 'Oct 2026', history: { '2026-09-30': 2000 } },
        2: { name: 'B', releaseDate: 'Oct 2026', history: { '2026-09-30': 2000, '2026-10-01': 2000 } }
    };
    assert.equal(getLatestCompleteCrawlDate(legacy), '2026-09-30');
});

test('freezes past release months in a versioned immutable archive', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-archive-test-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const excelFile = path.join(tempDirectory, 'tracker.xlsx');

    const sepGame = releaseDate => ({ appId: 10, name: 'Archived Sep', releaseDate, followers: 0 });
    const octGame = releaseDate => ({ appId: 20, name: 'Fall Oct', releaseDate, followers: 0 });

    await saveData([{ ...sepGame('15 Sep 2026'), followers: 5000 }], {
        dataDir: tempDirectory,
        excelFile,
        crawlDate: '2026-09-30',
        crawlComplete: true,
        fetchMetadata: false
    });
    const archiveFile = path.join(tempDirectory, 'steamdb_release_archives.json');
    assert.equal(fs.existsSync(archiveFile), false);
    let archives;

    await saveData([
        { ...sepGame('15 Sep 2026'), followers: 6000 },
        { ...octGame('10 Oct 2026'), followers: 4500 }
    ], {
        dataDir: tempDirectory,
        excelFile,
        crawlDate: '2026-10-15',
        crawlComplete: true,
        fetchMetadata: false
    });
    archives = JSON.parse(fs.readFileSync(archiveFile, 'utf8'));
    assert.equal(archives.schemaVersion, 1);
    assert.deepEqual(Object.keys(archives.months), ['2026-09']);
    // September is frozen at its latest September observation, not the newer October spillover value.
    assert.equal(archives.months['2026-09'].games[0].latestFollowers, 5000);
    assert.deepEqual(Object.keys(archives.months['2026-09'].games[0].history), ['2026-09-30']);

    await saveData([
        { ...sepGame('15 Sep 2026'), followers: 9000 },
        { ...octGame('10 Oct 2026'), followers: 8000 },
        { appId: 30, name: 'Winter Nov', releaseDate: '12 Nov 2026', followers: 3000 }
    ], {
        dataDir: tempDirectory,
        excelFile,
        crawlDate: '2026-11-15',
        crawlComplete: true,
        fetchMetadata: false
    });

    archives = JSON.parse(fs.readFileSync(archiveFile, 'utf8'));
    assert.equal(archives.months['2026-09'].games[0].latestFollowers, 5000);
    assert.deepEqual(Object.keys(archives.months).sort(), ['2026-09', '2026-10']);
    // October froze at its latest October observation, not the newer November value.
    assert.equal(archives.months['2026-10'].games[0].latestFollowers, 4500);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFile);
    assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), [
        'Summary', 'Qualified Upcoming', 'Sep 2026', 'Oct 2026', 'Nov 2026', 'Dec 2026', 'Unscheduled'
    ]);
    assert.equal(workbook.getWorksheet('Sep 2026').getCell('C2').value, 5000);
    assert.equal(workbook.getWorksheet('Nov 2026').getCell('C2').value, 3000);
});

test('October spillover does not create a September month tab or archive', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-spillover-test-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const excelFile = path.join(tempDirectory, 'tracker.xlsx');
    const archiveFile = path.join(tempDirectory, 'steamdb_release_archives.json');
    const masterFile = path.join(tempDirectory, 'steamdb_master_data.json');

    await saveData([{ appId: 99, name: 'Sep Spillover', releaseDate: '15 Sep 2026', followers: 5000 }], {
        dataDir: tempDirectory,
        excelFile,
        crawlDate: '2026-10-15',
        crawlComplete: true,
        fetchMetadata: false
    });

    assert.equal(fs.existsSync(archiveFile), false);
    const master = JSON.parse(fs.readFileSync(masterFile, 'utf8'));
    assert.ok(master['99']);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFile);
    assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), [
        'Summary', 'Qualified Upcoming', 'Oct 2026', 'Nov 2026', 'Unscheduled'
    ]);
});

test('does not commit or advance the calendar on an incomplete crawl', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-partial-test-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const excelFile = path.join(tempDirectory, 'tracker.xlsx');
    const masterFile = path.join(tempDirectory, 'steamdb_master_data.json');
    const archiveFile = path.join(tempDirectory, 'steamdb_release_archives.json');

    await saveData([{ appId: 1, name: 'Sep Game', releaseDate: '15 Sep 2026', followers: 5000 }], {
        dataDir: tempDirectory,
        excelFile,
        crawlDate: '2026-09-30',
        crawlComplete: true,
        fetchMetadata: false
    });

    await saveData([{ appId: 1, name: 'Sep Game', releaseDate: '15 Sep 2026', followers: 9000 }], {
        dataDir: tempDirectory,
        excelFile,
        crawlDate: '2026-10-15',
        crawlComplete: false,
        expectedWeeks: 9,
        successfulWeeks: 3,
        fetchMetadata: false
    });

    const master = JSON.parse(fs.readFileSync(masterFile, 'utf8'));
    assert.equal(master._crawlMeta.latestCompleteCrawlDate, '2026-09-30');
    assert.equal(Object.prototype.hasOwnProperty.call(master['1'].history, '2026-10-15'), false);
    assert.equal(fs.existsSync(archiveFile), false);
});

test('renders English and Simplified Chinese workbooks from the same model', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-i18n-test-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const master = makeMaster('2026-10-01', [
        {
            appId: 300,
            name: 'Indie Game',
            releaseDate: '20 Oct 2026',
            releaseTimestamp: Date.UTC(2026, 9, 20) / 1000,
            followers: 3000,
            publisher: 'N/A',
            developer: 'N/A',
            isFree: 'No'
        },
        { appId: 400, name: 'No Date', releaseDate: 'TBA', followers: 2000 }
    ]);

    const enFile = path.join(tempDirectory, 'en.xlsx');
    const zhFile = path.join(tempDirectory, 'zh.xlsx');
    await generateWorkbook(master, { excelFile: enFile, language: 'en', dataDir: tempDirectory });
    await generateWorkbook(master, { excelFile: zhFile, language: 'zh-CN', dataDir: tempDirectory });

    const enWorkbook = new ExcelJS.Workbook();
    await enWorkbook.xlsx.readFile(enFile);
    assert.deepEqual(enWorkbook.worksheets.map(sheet => sheet.name), [
        'Summary', 'Qualified Upcoming', 'Oct 2026', 'Nov 2026', 'Unscheduled'
    ]);
    const enSheet = enWorkbook.getWorksheet('Oct 2026');
    assert.deepEqual(enSheet.getRow(1).values.slice(1, 7), [
        'Release Date', 'Game Title', 'Followers', 'Publisher', 'Developer', 'Current Outlook'
    ]);
    assert.equal(enSheet.getCell('F2').value, 'P2 - Indie');
    assert.equal(enSheet.getCell('L2').value, 'Upcoming');
    assert.equal(enSheet.getCell('D2').value, 'N/A');
    assert.equal(enSheet.getCell('E2').value, 'N/A');
    assert.equal(enSheet.getCell('G2').value, 'No');

    const zhWorkbook = new ExcelJS.Workbook();
    await zhWorkbook.xlsx.readFile(zhFile);
    assert.deepEqual(zhWorkbook.worksheets.map(sheet => sheet.name), [
        '汇总', '达标即将发售', '2026年10月', '2026年11月', '未定档期'
    ]);
    const zhSheet = zhWorkbook.getWorksheet('2026年10月');
    assert.deepEqual(zhSheet.getRow(1).values.slice(1, 7), [
        '发售日期', '游戏名称', '关注人数', '发行商', '开发商', '当前评级'
    ]);
    assert.equal(zhSheet.getCell('F2').value, 'P2 - 独立游戏');
    assert.equal(zhSheet.getCell('L2').value, '即将发售');
    assert.equal(zhSheet.getCell('D2').value, '暂无');
    assert.equal(zhSheet.getCell('E2').value, '暂无');
    assert.equal(zhSheet.getCell('G2').value, '否');
    assert.equal(zhSheet.getCell('A2').numFmt, 'yyyy-mm-dd');

    const zip = await JSZip.loadAsync(fs.readFileSync(zhFile));
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    assert.equal(stylesXml.includes('&quot;年&quot;'), false);

    assert.deepEqual(zhWorkbook.getWorksheet('汇总').getRow(4).values.slice(1, 5)[0], '最近完整抓取');
});

test('keeps both active month tabs even when one has no qualified games', () => {
    const model = buildWorkbookModel(makeMaster('2026-10-01', [
        { appId: 1, name: 'Oct Only', releaseDate: '10 Oct 2026', followers: 2500 }
    ]), { archives: { schemaVersion: 1, months: {} } });

    assert.deepEqual(model.sections.map(section => section.monthKey), ['2026-10', '2026-11']);
    assert.equal(model.sections.find(section => section.monthKey === '2026-11').views.length, 0);
});

test('live sheets use only the latest crawl while month tabs retain stored games', async context => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'steamdb-live-test-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const excelFile = path.join(tempDirectory, 'tracker.xlsx');

    const master = {
        _crawlMeta: { latestCompleteCrawlDate: '2026-10-01' },
        1: { name: 'Stale Oct', releaseDate: '05 Oct 2026', history: { '2026-09-30': 5000 } },
        2: {
            name: 'Fresh Oct',
            releaseDate: '12 Oct 2026',
            history: { '2026-09-30': 4000, '2026-10-01': 4500 }
        },
        3: {
            name: 'Carried Oct',
            releaseDate: '18 Oct 2026',
            history: { '2026-09-30': 2000, '2026-10-01': 2500 }
        }
    };

    const model = buildWorkbookModel(master, { archives: { schemaVersion: 1, months: {} } });
    // A game is live when it appears in the latest complete crawl, even if a newer
    // observation exists from a later (partial) crawl.
    assert.deepEqual(model.activeViews.map(view => view.appId).sort(), ['2', '3']);
    assert.deepEqual(model.activeMonthViews.map(view => view.appId).sort(), ['1', '2', '3']);
    assert.deepEqual(model.unscheduled, []);

    await generateWorkbook(master, { excelFile, language: 'en', dataDir: tempDirectory });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFile);
    assert.equal(workbook.getWorksheet('Qualified Upcoming').rowCount - 1, 2);
    assert.equal(workbook.getWorksheet('Oct 2026').rowCount - 1, 3);
    assert.equal(workbook.getWorksheet('Summary').getCell('D5').value, 2);
});

test('resolves output languages and file targets', () => {
    assert.equal(normalizeLanguage('en'), 'en');
    assert.equal(normalizeLanguage('ZH-CN'), 'zh-CN');
    assert.equal(normalizeLanguage('zh'), 'zh-CN');
    assert.equal(normalizeLanguage(''), 'both');
    assert.equal(normalizeLanguage('nonsense'), 'both');

    assert.deepEqual(resolveOutputTargets({ excelFile: 'out.xlsx' }), [
        { language: 'en', file: 'out.xlsx' }
    ]);
    assert.deepEqual(resolveOutputTargets({ excelFile: 'out.xlsx', language: 'zh-CN' }), [
        { language: 'zh-CN', file: 'out.xlsx' }
    ]);
    assert.deepEqual(resolveOutputTargets({ language: 'both', dataDir: 'dir' }), [
        { language: 'en', file: path.join('dir', 'steamdb_upcoming_tracker.xlsx') },
        { language: 'zh-CN', file: path.join('dir', 'steamdb_upcoming_tracker_zh-CN.xlsx') }
    ]);
    assert.deepEqual(resolveOutputTargets({ language: 'zh-CN', dataDir: 'dir' }), [
        { language: 'zh-CN', file: path.join('dir', 'steamdb_upcoming_tracker_zh-CN.xlsx') }
    ]);
});

test('keeps the local date key in the machine time zone', () => {
    assert.match(getLocalDateKey(), /^\d{4}-\d{2}-\d{2}$/);
});
