const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const ExcelJS = require('exceljs');

process.env.FOCUSED_THRESHOLD = '500';

const {
    buildWorkbookModel,
    generateWorkbook,
    getFollowerInterpretation,
    getReleaseInfo
} = require('../steamdb_tracker');

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

test('uses exact MECE follower tier boundaries', () => {
    assert.equal(getFollowerInterpretation(999).tier, '');
    assert.equal(getFollowerInterpretation(1000).tier, 'Tier 2');
    assert.equal(getFollowerInterpretation(2999).tier, 'Tier 2');
    assert.equal(getFollowerInterpretation(3000).tier, 'Tier 3');
    assert.equal(getFollowerInterpretation(9999).tier, 'Tier 3');
    assert.equal(getFollowerInterpretation(10000).tier, 'Tier 4');
    assert.equal(getFollowerInterpretation(29999).tier, 'Tier 4');
    assert.equal(getFollowerInterpretation(30000).tier, 'Tier 5');
});

test('keeps month-boundary dates at UTC midnight', () => {
    const release = getReleaseInfo({
        releaseDate: '01 Oct',
        releaseTimestamp: Date.UTC(2026, 9, 1) / 1000
    }, '2026-09-30');

    assert.equal(release.monthKey, '2026-10');
    assert.equal(release.exactDate.toISOString(), '2026-10-01T00:00:00.000Z');
});

test('writes interpreted fields without a Days to Release detail column', async context => {
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
    }, outputFile);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(outputFile);
    assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), [
        'Summary', 'Qualified Upcoming', 'Oct 2026', 'Unscheduled'
    ]);

    const worksheet = workbook.getWorksheet('Oct 2026');
    const headers = worksheet.getRow(1).values.slice(1);
    assert.equal(headers.includes('Days to Release'), false);
    assert.deepEqual(headers.slice(0, 7), [
        'Release Date', 'Game Title', 'Followers', 'Publisher', 'Developer',
        'Follower Tier', 'Commercial Standing'
    ]);

    assert.equal(worksheet.getCell('A2').value.toISOString(), '2026-10-01T00:00:00.000Z');
    assert.equal(worksheet.getCell('F2').value, 'Tier 3');
    assert.equal(worksheet.getCell('G2').value, 'Commercial Hit / Sustainable Indie');
    assert.equal(worksheet.views[0].state, 'frozen');
    assert(worksheet.autoFilter);
});
