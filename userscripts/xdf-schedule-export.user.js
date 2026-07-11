// ==UserScript==
// @name         XDF 课表导出
// @namespace    https://github.com/nowscott/XdfScheduleCrawler
// @version      1.1.1
// @description  在已登录的课表页面中导出课表明细和五时段月视图。
// @author       nowscott
// @match        https://we.xdf.cn/*
// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
// @downloadURL  https://raw.githubusercontent.com/nowscott/XdfScheduleCrawler/main/userscripts/xdf-schedule-export.user.js
// @updateURL    https://raw.githubusercontent.com/nowscott/XdfScheduleCrawler/main/userscripts/xdf-schedule-export.user.js
// @supportURL   https://github.com/nowscott/XdfScheduleCrawler/issues
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    const API_BASE = 'https://gw-xeasy.xdf.cn/xeasy-srv-teachinghub';
    const MAX_CONCURRENT_REQUESTS = 3;
    const WEEKDAYS = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
    const EMPTY_SLOT_LABELS = [
        '08:00–10:00  无课',
        '10:00/10:20 时段  无课',
        '13:40–15:40  无课',
        '16:00–18:00  无课',
        '18:00/18:30 时段  无课',
    ];
    const FIELD_NAME_MAP = {
        courseName: '课程名称', coursename: '课程名称', name: '课程名称',
        className: '课程名称', classname: '课程名称', course_name: '课程名称',
        lessonName: '学生', date: '日期', courseDate: '日期', scheduleDate: '日期',
        startDate: '日期', day: '日期', startTime: '开始时间', endTime: '结束时间',
        time: '时间', courseTime: '时间', beginTime: '开始时间', finishTime: '结束时间',
        startTimeStr: '开始时间', endTimeStr: '结束时间',
        lessonStartTime: '开始时间', lessonEndTime: '结束时间', room: '教室',
        classRoom: '教室', classroom: '教室', roomName: '教室', address: '地点',
        teacher: '老师', teacherName: '老师', instructor: '老师', lecturer: '老师',
        student: '学生', studentName: '学生', studentList: '学生名单', status: '状态',
        courseStatus: '状态', lessonStatus: '状态', campus: '校区', schoolName: '校区',
        school: '校区', branch: '校区', remark: '备注', note: '备注', memo: '备注',
        description: '备注', desc: '备注', grade: '年级', subject: '科目',
        lessonTypeDesc: '课次类型', lessonType: '课次类型码', classCode: '班级代码',
        lessonId: '课次ID', stuCount: '学生数', schoolId: '校区ID', resourceCount: '资源数',
        videoType: '视频类型', feedbackAllFinished: '反馈完成', bindStatus: '绑定状态',
        teachingChannelCodeList: '教学渠道', _date: '日期',
    };
    const COLUMN_PRIORITY = ['日期', '开始时间', '结束时间', '学生', '课程名称', '老师', '教室', '校区', '课次类型', '状态', '备注'];
    const STATUS_MAP = { 0: '未开始', 1: '进行中', 2: '已结束', '0': '未开始', '1': '进行中', '2': '已结束' };
    const COLORS = {
        header: '4472C4', weekday: '5B9BD5', date: 'D9E7F5', activeDate: '9DC3E6',
        course: 'E2F0D9', weekend: 'F7F7F7', outside: 'F5F6F8', evenRow: 'D6E4F0',
        muted: '7F8C9A', titleText: 'FFFFFF', summaryText: '365F91', courseText: '375623',
        border: 'B4B4B4', white: 'FFFFFF', weekendDate: 'E7E6E6',
    };
    const THIN_BORDER = ['left', 'right', 'top', 'bottom'].reduce((border, side) => {
        border[side] = { style: 'thin', color: { rgb: COLORS.border } };
        return border;
    }, {});

    function font(size, options = {}) {
        return { name: 'Microsoft YaHei', sz: size, ...options };
    }

    function fill(color) {
        return { patternType: 'solid', fgColor: { rgb: color } };
    }

    function alignment(horizontal = 'left') {
        return { horizontal, vertical: 'center', wrapText: true };
    }

    function monthRanges(startDate, endDate) {
        const ranges = [];
        let cursor = new Date(`${startDate}T00:00:00`);
        const end = new Date(`${endDate}T00:00:00`);
        while (cursor <= end) {
            const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0);
            const rangeEnd = monthEnd < end ? monthEnd : end;
            ranges.push([formatDate(cursor), formatDate(rangeEnd)]);
            cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
        }
        return ranges;
    }

    function formatDate(date) {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    }

    function wait(ms) {
        return new Promise((resolve) => window.setTimeout(resolve, ms));
    }

    async function requestJson(url, options) {
        const response = await fetch(url, { credentials: 'include', ...options });
        const body = await response.json().catch(() => null);
        if (!response.ok || body?.code !== '1') {
            throw new Error(body?.msg || `请求失败（HTTP ${response.status}）`);
        }
        return body.data;
    }

    async function fetchSchedules(startDate, endDate, onProgress) {
        const calendar = [];
        for (const [rangeStart, rangeEnd] of monthRanges(startDate, endDate)) {
            const data = await requestJson(`${API_BASE}/lesson/calendar/v2?startDate=${rangeStart}&endDate=${rangeEnd}`);
            calendar.push(...(Array.isArray(data) ? data : []));
        }
        const daysWithLessons = calendar
            .filter((item) => Number(item.lessonCount) > 0)
            .sort((a, b) => String(a.day).localeCompare(String(b.day)));
        const schedules = [];
        let nextIndex = 0;
        let completedDays = 0;
        const workerCount = Math.min(MAX_CONCURRENT_REQUESTS, daysWithLessons.length);

        async function worker() {
            while (nextIndex < daysWithLessons.length) {
                const index = nextIndex;
                nextIndex += 1;
                const { day } = daysWithLessons[index];
                onProgress(`正在读取 ${index + 1}/${daysWithLessons.length} 天 · 已完成 ${completedDays} 天…`);
                const data = await requestJson(`${API_BASE}/lesson/list-by-date/v2`, {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ date: day, lessonStatus: -1 }),
                });
                schedules.push(...(data?.lessonList || []).map((lesson) => ({ ...lesson, _date: day })));
                completedDays += 1;
                onProgress(`已完成 ${completedDays}/${daysWithLessons.length} 天 · 已获取 ${schedules.length} 节课…`);
            }
        }

        await Promise.all(Array.from({ length: workerCount }, worker));
        return schedules.sort((a, b) => String(a._date).localeCompare(String(b._date))
            || String(a.lessonStartTime || '').localeCompare(String(b.lessonStartTime || '')));
    }

    function cell(sheet, row, column, value, style) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        sheet[address] = { t: typeof value === 'number' ? 'n' : 's', v: value ?? '', s: style };
    }

    function columnLetter(column) {
        return XLSX.utils.encode_col(column);
    }

    function displayValue(value, header) {
        if (Array.isArray(value)) {
            if (value[0] && typeof value[0] === 'object' && 'studentName' in value[0]) {
                return value.map((item) => item.studentName || '').filter(Boolean).join(', ');
            }
            return value.map((item) => String(item)).join(', ');
        }
        if (value && typeof value === 'object') return value.name || String(value);
        if (header === '状态') return STATUS_MAP[value] || value || '';
        return value ?? '';
    }

    function visualLength(value) {
        return Array.from(String(value || '')).reduce((length, character) => length + (character.charCodeAt(0) > 127 ? 2 : 1), 0);
    }

    function createDetailSheet(schedules) {
        const keys = [...new Set(schedules.flatMap((lesson) => Object.keys(lesson)))];
        const rawToHeader = new Map(keys.map((key) => [key, FIELD_NAME_MAP[key] || key]));
        const headers = [...new Set(keys.map((key) => rawToHeader.get(key)))].sort((a, b) => {
            const aIndex = COLUMN_PRIORITY.indexOf(a);
            const bIndex = COLUMN_PRIORITY.indexOf(b);
            return (aIndex < 0 ? 999 : aIndex) - (bIndex < 0 ? 999 : bIndex);
        });
        const headerToRaw = new Map();
        keys.forEach((key) => headerToRaw.set(rawToHeader.get(key), key));
        const sheet = XLSX.utils.aoa_to_sheet([headers]);
        const widths = headers.map((header) => visualLength(header));

        headers.forEach((header, index) => {
            cell(sheet, 0, index, header, { font: font(11, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.header), alignment: alignment('center'), border: THIN_BORDER });
        });
        schedules.forEach((lesson, rowIndex) => {
            headers.forEach((header, columnIndex) => {
                const value = displayValue(lesson[headerToRaw.get(header)], header);
                widths[columnIndex] = Math.max(widths[columnIndex], visualLength(value));
                cell(sheet, rowIndex + 1, columnIndex, value, {
                    font: font(10), alignment: alignment('center'), border: THIN_BORDER,
                    ...(rowIndex % 2 === 0 ? { fill: fill(COLORS.evenRow) } : {}),
                });
            });
        });
        sheet['!cols'] = widths.map((width) => ({ wch: Math.min(Math.max(width + 2, 10), 40) }));
        sheet['!rows'] = [{ hpt: 24 }];
        sheet['!autofilter'] = { ref: `A1:${columnLetter(headers.length - 1)}1` };
        sheet['!ref'] = `A1:${columnLetter(headers.length - 1)}${schedules.length + 1}`;
        return sheet;
    }

    function timePart(value) {
        const match = String(value || '').match(/(\d{2}:\d{2})/);
        return match ? match[1] : '';
    }

    function studentName(lesson) {
        if (lesson.lessonName) return String(lesson.lessonName);
        return (lesson.studentList || []).map((item) => item.studentName || '').filter(Boolean).join('、');
    }

    function shortRoom(value) {
        const room = String(value || '');
        const match = room.match(/个性化([^（(]+)/);
        return match ? match[1].trim() : room.replace('万博敏捷广场', '');
    }

    function slotIndex(startTime) {
        const [hour, minute] = startTime.split(':').map(Number);
        const total = hour * 60 + minute;
        if (total < 9 * 60 + 30) return 0;
        if (total < 12 * 60 + 30) return 1;
        if (total < 15 * 60) return 2;
        if (total < 17 * 60 + 30) return 3;
        return 4;
    }

    function monthWeeks(year, month) {
        const firstWeekday = (new Date(year, month - 1, 1).getDay() + 6) % 7;
        const days = new Date(year, month, 0).getDate();
        const weeks = [];
        let day = 1 - firstWeekday;
        while (day <= days) {
            const week = [];
            for (let index = 0; index < 7; index += 1, day += 1) week.push(day > 0 && day <= days ? day : 0);
            weeks.push(week);
        }
        return weeks;
    }

    function createMonthSheet(year, month, lessons) {
        const sheet = XLSX.utils.aoa_to_sheet(Array.from({ length: 3 + monthWeeks(year, month).length * 6 }, () => Array(7).fill('')));
        const byDate = new Map();
        lessons.forEach((lesson) => {
            const date = String(lesson._date || '').slice(0, 10);
            const entries = byDate.get(date) || [];
            entries.push(lesson);
            byDate.set(date, entries);
        });
        byDate.forEach((entries) => entries.sort((a, b) => String(a.lessonStartTime || '').localeCompare(String(b.lessonStartTime || ''))));
        const weeks = monthWeeks(year, month);
        sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } }];
        sheet['!cols'] = Array.from({ length: 7 }, () => ({ wch: 27 }));
        sheet['!rows'] = [{ hpt: 40 }, { hpt: 24 }, { hpt: 28 }, ...weeks.flatMap(() => [{ hpt: 22 }, ...Array.from({ length: 5 }, () => ({ hpt: 23 }))])];

        cell(sheet, 0, 0, `${year} 年 ${month} 月课程月视图`, { font: font(20, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.header), alignment: alignment('center') });
        cell(sheet, 1, 0, `共 ${lessons.length} 节课 · ${byDate.size} 个有课日期 · 每个日期固定 5 个时段，一节课占一行`, { font: font(10, { color: { rgb: COLORS.summaryText } }), fill: fill(COLORS.date), alignment: alignment('center') });
        WEEKDAYS.forEach((weekday, column) => {
            cell(sheet, 2, column, weekday, { font: font(11, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.weekday), alignment: alignment('center'), border: THIN_BORDER });
        });

        weeks.forEach((week, weekIndex) => {
            const dateRow = 3 + weekIndex * 6;
            week.forEach((day, weekdayIndex) => {
                const isWeekend = weekdayIndex >= 5;
                if (!day) {
                    for (let slot = 0; slot < 6; slot += 1) cell(sheet, dateRow + slot, weekdayIndex, '', { fill: fill(COLORS.outside), border: THIN_BORDER });
                    return;
                }
                const dateKey = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                const dayLessons = byDate.get(dateKey) || [];
                cell(sheet, dateRow, weekdayIndex, `${day} 日${dayLessons.length ? ` · ${dayLessons.length} 节` : ''}`, {
                    font: font(10, { bold: true, color: { rgb: COLORS.summaryText } }),
                    fill: fill(dayLessons.length ? COLORS.activeDate : (isWeekend ? COLORS.weekendDate : COLORS.date)), alignment: alignment('left'), border: THIN_BORDER,
                });
                const slots = Array(5).fill(null);
                dayLessons.forEach((lesson) => {
                    const start = timePart(lesson.lessonStartTime);
                    if (!start) return;
                    const index = slotIndex(start);
                    if (slots[index]) throw new Error(`${dateKey} 的第 ${index + 1} 时段存在多节课程，无法放入固定五行月视图。`);
                    slots[index] = lesson;
                });
                slots.forEach((lesson, slot) => {
                    const row = dateRow + slot + 1;
                    if (!lesson) {
                        cell(sheet, row, weekdayIndex, EMPTY_SLOT_LABELS[slot], { font: font(8, { color: { rgb: COLORS.muted } }), fill: fill(isWeekend ? COLORS.weekend : COLORS.white), alignment: alignment('left'), border: THIN_BORDER });
                        return;
                    }
                    const start = timePart(lesson.lessonStartTime);
                    const end = timePart(lesson.lessonEndTime);
                    cell(sheet, row, weekdayIndex, `${start}–${end}  ${studentName(lesson)} · ${shortRoom(lesson.roomName)}`, { font: font(8, { bold: true, color: { rgb: COLORS.courseText } }), fill: fill(COLORS.course), alignment: alignment('left'), border: THIN_BORDER });
                });
            });
        });
        return sheet;
    }

    function exportWorkbook(schedules, startDate, endDate) {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, createDetailSheet(schedules), '课表');
        const grouped = new Map();
        schedules.forEach((lesson) => {
            const match = String(lesson._date || '').match(/^(\d{4})-(\d{2})/);
            if (!match) return;
            const key = `${match[1]}-${match[2]}`;
            const entries = grouped.get(key) || [];
            entries.push(lesson);
            grouped.set(key, entries);
        });
        [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([key, lessons]) => {
            const [year, month] = key.split('-').map(Number);
            XLSX.utils.book_append_sheet(workbook, createMonthSheet(year, month, lessons), `${year}年${month}月月视图`);
        });
        XLSX.writeFile(workbook, `课表_${startDate}_至_${endDate}.xlsx`);
    }

    function showExportDialog() {
        if (document.getElementById('xdf-schedule-export-dialog')) return;
        const dialog = document.createElement('div');
        dialog.id = 'xdf-schedule-export-dialog';
        dialog.innerHTML = `
            <div class="xdf-export-card" role="dialog" aria-modal="true" aria-labelledby="xdf-export-title">
                <div class="xdf-export-heading"><div><h2 id="xdf-export-title">导出课表 Excel</h2><p>三路并发读取，完整导出明细表和月视图</p></div><button type="button" class="xdf-export-close" aria-label="关闭">×</button></div>
                <div class="xdf-export-fields"><label>开始日期<input id="xdf-export-start" type="date" required></label><label>结束日期<input id="xdf-export-end" type="date" required></label></div>
                <p id="xdf-export-status" class="xdf-export-status" aria-live="polite"></p>
                <div class="xdf-export-actions"><button type="button" class="xdf-export-cancel">取消</button><button type="button" class="xdf-export-submit">开始导出</button></div>
            </div>`;
        document.body.appendChild(dialog);
        const startInput = dialog.querySelector('#xdf-export-start');
        const endInput = dialog.querySelector('#xdf-export-end');
        const status = dialog.querySelector('#xdf-export-status');
        const submit = dialog.querySelector('.xdf-export-submit');
        const close = () => dialog.remove();
        dialog.querySelector('.xdf-export-close').addEventListener('click', close);
        dialog.querySelector('.xdf-export-cancel').addEventListener('click', close);
        dialog.addEventListener('click', (event) => { if (event.target === dialog) close(); });
        startInput.focus();
        submit.addEventListener('click', async () => {
            const rangeStart = startInput.value;
            const rangeEnd = endInput.value;
            if (!rangeStart || !rangeEnd) return void (status.textContent = '请选择开始日期和结束日期。');
            if (rangeStart > rangeEnd) return void (status.textContent = '结束日期不能早于开始日期。');
            submit.disabled = true;
            status.textContent = '正在查询有课日期…';
            const startedAt = performance.now();
            try {
                const schedules = await fetchSchedules(rangeStart, rangeEnd, (message) => { status.textContent = message; });
                if (!schedules.length) return void (status.textContent = '所选日期范围没有课程，未生成文件。');
                status.textContent = '正在生成明细表和月视图…';
                exportWorkbook(schedules, rangeStart, rangeEnd);
                const elapsedSeconds = ((performance.now() - startedAt) / 1000).toFixed(1);
                status.textContent = `已导出 ${schedules.length} 条课程记录，耗时 ${elapsedSeconds} 秒，文件已开始下载。`;
            } catch (error) {
                console.error('[XDF 课表导出]', error);
                status.textContent = `导出失败：${error.message || error}`;
            } finally {
                submit.disabled = false;
            }
        });
    }

    function addExportButton() {
        if (document.getElementById('xdf-schedule-export-button')) return;
        const button = document.createElement('button');
        button.id = 'xdf-schedule-export-button';
        button.type = 'button';
        button.textContent = '导出课表 Excel';
        button.title = '导出指定日期范围的课表明细和月视图';
        Object.assign(button.style, { position: 'fixed', right: '24px', bottom: '24px', zIndex: '99999', border: '0', borderRadius: '8px', padding: '10px 14px', cursor: 'pointer', background: '#1677ff', color: '#fff', boxShadow: '0 4px 14px rgb(0 0 0 / 18%)' });
        button.addEventListener('click', showExportDialog);
        document.body.appendChild(button);
        const style = document.createElement('style');
        style.textContent = `
            #xdf-schedule-export-dialog { position: fixed; inset: 0; z-index: 100000; display: grid; place-items: center; padding: 20px; background: rgb(15 23 42 / 45%); }
            #xdf-schedule-export-dialog * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
            .xdf-export-card { width: min(420px, 100%); padding: 24px; border-radius: 16px; background: #fff; color: #172033; box-shadow: 0 18px 48px rgb(15 23 42 / 25%); }
            .xdf-export-heading { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; margin-bottom: 22px; }.xdf-export-heading h2 { margin: 0; font-size: 20px; line-height: 1.3; }.xdf-export-heading p { margin: 6px 0 0; color: #667085; font-size: 14px; }
            .xdf-export-close { border: 0; padding: 0 4px; background: transparent; color: #667085; cursor: pointer; font-size: 28px; line-height: 1; }.xdf-export-fields { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }.xdf-export-fields label { display: grid; gap: 7px; color: #344054; font-size: 14px; font-weight: 600; }.xdf-export-fields input { width: 100%; min-height: 40px; padding: 8px 10px; border: 1px solid #d0d5dd; border-radius: 8px; color: #172033; font: inherit; }.xdf-export-fields input:focus { outline: 3px solid rgb(22 119 255 / 18%); border-color: #1677ff; }
            .xdf-export-status { min-height: 20px; margin: 16px 0 0; color: #667085; font-size: 13px; line-height: 1.5; }.xdf-export-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; }.xdf-export-actions button { min-height: 38px; padding: 0 14px; border: 1px solid #d0d5dd; border-radius: 8px; background: #fff; color: #344054; cursor: pointer; font: inherit; }.xdf-export-actions .xdf-export-submit { border-color: #1677ff; background: #1677ff; color: #fff; }.xdf-export-actions button:disabled { cursor: wait; opacity: .7; }
            @media (max-width: 480px) { .xdf-export-fields { grid-template-columns: 1fr; } }
        `;
        document.head.appendChild(style);
    }

    addExportButton();
}());
