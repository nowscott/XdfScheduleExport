import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

function columnName(index) {
    let result = '';
    let value = index + 1;
    while (value) {
        const remainder = (value - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        value = Math.floor((value - 1) / 26);
    }
    return result;
}

function createXlsxSpy() {
    const writes = [];
    return {
        writes,
        utils: {
            encode_col: columnName,
            encode_cell: ({ r, c }) => `${columnName(c)}${r + 1}`,
            aoa_to_sheet(rows) {
                const sheet = {};
                rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
                    if (value !== '') sheet[`${columnName(columnIndex)}${rowIndex + 1}`] = { t: 's', v: value };
                }));
                return sheet;
            },
            book_new: () => ({ SheetNames: [], Sheets: {} }),
            book_append_sheet(workbook, sheet, name) {
                workbook.SheetNames.push(name);
                workbook.Sheets[name] = sheet;
            },
        },
        writeFile(workbook, filename) {
            writes.push({ workbook, filename });
        },
    };
}

const sourcePath = new URL('../userscripts/xdf-schedule-export.user.js', import.meta.url);
const userscriptSource = fs.readFileSync(sourcePath, 'utf8');
assert.match(userscriptSource, /@version\s+1\.3\.6/);
assert.match(userscriptSource, /backdrop-filter: blur\(36px\) saturate\(180%\)/);
assert.match(userscriptSource, /@media \(max-width: 600px\)/);
assert.match(userscriptSource, /@media \(prefers-reduced-motion: reduce\)/);
assert.match(userscriptSource, /aria-labelledby="xdf-export-range-label"/);
assert.match(userscriptSource, /event\.key === 'Escape'/);
assert.match(userscriptSource, /item\.setAttribute\('aria-pressed', String\(active\)\)/);
assert.match(userscriptSource, /FLOATING_BUTTON_KEY/);
assert.match(userscriptSource, /pointerdown/);
assert.match(userscriptSource, /is-docked/);
assert.match(userscriptSource, /is-docked-left/);
assert.match(userscriptSource, /is-docked-right/);
assert.match(userscriptSource, /dockButton/);
assert.match(userscriptSource, /transition: left \.46s/);
assert.match(userscriptSource, /aria-label', '导出课表'/);
assert.doesNotMatch(userscriptSource, /<span>导出课表<\/span>/);
assert.match(userscriptSource, /window\.innerHeight \* \.25/);

const source = userscriptSource.replace(
    '    addExportButton();',
    '    globalThis.__userscriptTestHooks = { exportWorkbook, fetchLessonDetails, loadPreferences, savePreferences, loadFloatingPosition, saveFloatingPosition, rangeForPreset };',
);
const xlsx = createXlsxSpy();
const stored = new Map();
let failSecondDay = true;
const context = {
    XLSX: xlsx,
    console,
    Date,
    Map,
    Set,
    JSON,
    String,
    Number,
    Array,
    Math,
    RegExp,
    performance: { now: () => 0 },
    window: { setTimeout },
    localStorage: {
        getItem: (key) => stored.get(key) ?? null,
        setItem: (key, value) => stored.set(key, value),
    },
    fetch: async (url, options = {}) => {
        const day = JSON.parse(options.body || '{}').date;
        if (day === '2026-07-14' && failSecondDay) {
            return { ok: false, status: 503, json: async () => ({ msg: 'temporary failure' }) };
        }
        return {
            ok: true,
            status: 200,
            json: async () => ({ code: '1', data: { lessonList: [{ lessonName: `${day} 学员`, lessonStartTime: `${day} 10:20:00`, lessonEndTime: `${day} 12:20:00` }] } }),
        };
    },
};
context.globalThis = context;
vm.runInNewContext(source, context, { filename: 'xdf-schedule-export.user.js' });

const { exportWorkbook, fetchLessonDetails, loadPreferences, savePreferences, loadFloatingPosition, saveFloatingPosition, rangeForPreset } = context.__userscriptTestHooks;
const schedules = [
    { _date: '2026-07-13', lessonName: '陈同学', courseName: '数学', teacherName: '牛老师', campus: '广州', lessonStartTime: '2026-07-13 08:00:00', lessonEndTime: '2026-07-13 10:00:00', roomName: '个性化V228' },
    { _date: '2026-07-13', lessonName: '王同学', courseName: '数学', teacherName: '牛老师', campus: '广州', lessonStartTime: '2026-07-13 10:20:00', lessonEndTime: '2026-07-13 12:20:00', roomName: '个性化V229' },
    { _date: '2026-07-14', lessonName: '赵同学', courseName: '数学', teacherName: '牛老师', campus: '广州', lessonStartTime: '2026-07-14 10:20:00', lessonEndTime: '2026-07-14 12:20:00', roomName: '个性化V231' },
    { _date: '2026-07-31', lessonName: '跨月前', courseName: '数学', teacherName: '牛老师', campus: '广州', lessonStartTime: '2026-07-31 10:20:00', lessonEndTime: '2026-07-31 12:20:00', roomName: '个性化V232' },
    { _date: '2026-08-01', lessonName: '跨月后', courseName: '数学', teacherName: '牛老师', campus: '广州', lessonStartTime: '2026-08-01 10:20:00', lessonEndTime: '2026-08-01 12:20:00', roomName: '个性化V233' },
    { _date: '2026-08-03', lessonName: '李同学', courseName: '数学', teacherName: '牛老师', campus: '广州', lessonStartTime: '2026-08-03 16:00:00', lessonEndTime: '2026-08-03 18:00:00', roomName: '个性化V230' },
];

exportWorkbook(schedules, '2026-07-13', '2026-08-31', { combineMonthViews: true });
let write = xlsx.writes.at(-1);
assert.deepEqual(write.workbook.SheetNames, ['月视图', '统计', '详细课表']);
assert.equal(write.workbook.Sheets.月视图.A1.v, '2026-07-13 至 2026-08-31 课表月视图');
assert.ok(!Object.values(write.workbook.Sheets.月视图).some((cell) => cell?.v?.includes('课程月视图') && cell.v !== '2026-07-13 至 2026-08-31 课表月视图'));
assert.ok(Object.values(write.workbook.Sheets.月视图).some((cell) => cell?.v?.includes('08:00–10:00  陈同学')));
assert.ok(Object.values(write.workbook.Sheets.月视图).some((cell) => cell?.v?.includes('10:20–12:20  王同学')));
assert.ok(!Object.values(write.workbook.Sheets.月视图).some((cell) => cell?.v?.includes('08:00–10:00  无课')));
const julyMonthSheet = write.workbook.Sheets.月视图;
const cellAddressFor = (text) => Object.entries(julyMonthSheet).find(([, cell]) => cell?.v?.includes(text))?.[0];
const rowFor = (text) => Number(cellAddressFor(text).match(/\d+$/)[0]);
assert.notEqual(rowFor('08:00–10:00  陈同学'), rowFor('10:20–12:20  赵同学'));
assert.equal(rowFor('10:20–12:20  王同学'), rowFor('10:20–12:20  赵同学'));
assert.equal(rowFor('10:20–12:20  跨月前'), rowFor('10:20–12:20  跨月后'));
assert.equal(julyMonthSheet.A2.v, '时间段');
assert.equal(julyMonthSheet[`A${rowFor('10:20–12:20  王同学')}`].v, '10:20–12:20');
assert.equal(Object.values(julyMonthSheet).filter((cell) => cell?.v === '时间段').length, 1);
assert.ok(Object.values(julyMonthSheet).some((cell) => cell?.v === '8 月 1 日 · 1 节'));
assert.equal(write.workbook.Sheets.统计.A1.v, '课表导出统计');
assert.equal(write.workbook.Sheets.统计['!ref'], 'A1:J20');
assert.equal(write.workbook.Sheets.详细课表.A1.v, '日期');

exportWorkbook(schedules, '2026-07-13', '2026-08-31', { combineMonthViews: false });
write = xlsx.writes.at(-1);
assert.deepEqual(write.workbook.SheetNames, ['2026年7月月视图', '2026年8月月视图', '统计', '详细课表']);
assert.equal(write.workbook.Sheets['2026年7月月视图'].A1.v, '2026 年 7 月课程月视图');

savePreferences({ startDate: '2026-07-13', endDate: '2026-08-31', combineMonthViews: false });
assert.equal(JSON.stringify(loadPreferences()), JSON.stringify({ startDate: '2026-07-13', endDate: '2026-08-31', combineMonthViews: false }));
saveFloatingPosition({ x: 120, y: 240 });
assert.equal(JSON.stringify(loadFloatingPosition()), JSON.stringify({ x: 120, y: 240 }));
stored.set('xdf-schedule-export-floating-button-v1', '{bad json');
assert.equal(loadFloatingPosition(), null);
const [termStart, termEnd] = rangeForPreset('this-term');
assert.match(termStart, /^\d{4}-\d{2}-\d{2}$/);
assert.match(termEnd, /^\d{4}-\d{2}-\d{2}$/);

const initial = await fetchLessonDetails([{ day: '2026-07-13' }, { day: '2026-07-14' }], () => {});
assert.equal(initial.schedules.length, 1);
assert.equal(JSON.stringify(initial.failedDays.map(({ day }) => day)), '["2026-07-14"]');
failSecondDay = false;
const retried = await fetchLessonDetails(initial.failedDays, () => {});
assert.equal(retried.schedules.length, 1);
assert.equal(retried.failedDays.length, 0);

console.log('userscript tests passed');
