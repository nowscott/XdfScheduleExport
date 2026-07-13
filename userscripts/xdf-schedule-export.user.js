// ==UserScript==
// @name         XDF 课表导出
// @namespace    https://github.com/nowscott/XdfScheduleCrawler
// @version      1.3.0
// @description  在已登录的课表页面中导出月视图、统计和课表明细。
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
    const PREFERENCES_KEY = 'xdf-schedule-export-preferences-v1';
    const WEEKDAYS = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'];
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

    function dateFromParts(year, month, day) {
        return new Date(year, month - 1, day);
    }

    function currentMonthRange() {
        const today = new Date();
        const start = dateFromParts(today.getFullYear(), today.getMonth() + 1, 1);
        const end = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        return [formatDate(start), formatDate(end)];
    }

    function rangeForPreset(preset) {
        const today = new Date();
        const weekday = (today.getDay() + 6) % 7;
        if (preset === 'this-week' || preset === 'next-week') {
            const offset = preset === 'next-week' ? 7 : 0;
            const start = new Date(today.getFullYear(), today.getMonth(), today.getDate() - weekday + offset);
            const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6);
            return [formatDate(start), formatDate(end)];
        }
        if (preset === 'this-month' || preset === 'next-month') {
            const offset = preset === 'next-month' ? 1 : 0;
            const start = new Date(today.getFullYear(), today.getMonth() + offset, 1);
            const end = new Date(today.getFullYear(), today.getMonth() + offset + 1, 0);
            return [formatDate(start), formatDate(end)];
        }
        const autumnTerm = today.getMonth() >= 7;
        const start = new Date(today.getFullYear(), autumnTerm ? 7 : 1, 1);
        const end = new Date(today.getFullYear() + (autumnTerm ? 1 : 0), autumnTerm ? 0 : 6, autumnTerm ? 31 : 31);
        return [formatDate(start), formatDate(end)];
    }

    function loadPreferences() {
        try {
            const value = JSON.parse(localStorage.getItem(PREFERENCES_KEY) || '{}');
            return {
                startDate: /^\d{4}-\d{2}-\d{2}$/.test(value.startDate || '') ? value.startDate : '',
                endDate: /^\d{4}-\d{2}-\d{2}$/.test(value.endDate || '') ? value.endDate : '',
                combineMonthViews: value.combineMonthViews !== false,
            };
        } catch {
            return { startDate: '', endDate: '', combineMonthViews: true };
        }
    }

    function savePreferences({ startDate, endDate, combineMonthViews }) {
        localStorage.setItem(PREFERENCES_KEY, JSON.stringify({ startDate, endDate, combineMonthViews }));
    }

    async function requestJson(url, options) {
        const response = await fetch(url, { credentials: 'include', ...options });
        const body = await response.json().catch(() => null);
        if (!response.ok || body?.code !== '1') {
            throw new Error(body?.msg || `请求失败（HTTP ${response.status}）`);
        }
        return body.data;
    }

    async function fetchLessonDetails(daysWithLessons, onProgress) {
        const schedules = [];
        const failedDays = [];
        let nextIndex = 0;
        let completedDays = 0;
        const workerCount = Math.min(MAX_CONCURRENT_REQUESTS, daysWithLessons.length);

        async function worker() {
            while (nextIndex < daysWithLessons.length) {
                const index = nextIndex;
                nextIndex += 1;
                const { day } = daysWithLessons[index];
                onProgress(`正在读取 ${index + 1}/${daysWithLessons.length} 天 · 已完成 ${completedDays} 天…`);
                try {
                    const data = await requestJson(`${API_BASE}/lesson/list-by-date/v2`, {
                        method: 'POST', headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ date: day, lessonStatus: -1 }),
                    });
                    schedules.push(...(data?.lessonList || []).map((lesson) => ({ ...lesson, _date: day })));
                } catch (error) {
                    failedDays.push({ day, message: error.message || String(error) });
                } finally {
                    completedDays += 1;
                    onProgress(`已完成 ${completedDays}/${daysWithLessons.length} 天 · 已获取 ${schedules.length} 节课${failedDays.length ? ` · ${failedDays.length} 天待重试` : ''}…`);
                }
            }
        }

        await Promise.all(Array.from({ length: workerCount }, worker));
        return { schedules, failedDays };
    }

    function sortSchedules(schedules) {
        return schedules.sort((a, b) => String(a._date).localeCompare(String(b._date))
            || String(a.lessonStartTime || '').localeCompare(String(b.lessonStartTime || '')));
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
        const result = await fetchLessonDetails(daysWithLessons, onProgress);
        result.schedules = sortSchedules(result.schedules);
        return result;
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

    function monthsInRange(startDate, endDate) {
        const months = [];
        const cursor = new Date(`${startDate}T00:00:00`);
        const end = new Date(`${endDate}T00:00:00`);
        cursor.setDate(1);
        while (cursor <= end) {
            months.push({ year: cursor.getFullYear(), month: cursor.getMonth() + 1 });
            cursor.setMonth(cursor.getMonth() + 1, 1);
        }
        return months;
    }

    function createMonthSheet(year, month, lessons) {
        const byDate = new Map();
        lessons.forEach((lesson) => {
            const date = String(lesson._date || '').slice(0, 10);
            const entries = byDate.get(date) || [];
            entries.push(lesson);
            byDate.set(date, entries);
        });
        byDate.forEach((entries) => entries.sort((a, b) => String(a.lessonStartTime || '').localeCompare(String(b.lessonStartTime || ''))));
        const weeks = monthWeeks(year, month);
        const lessonRowsByWeek = weeks.map((week) => Math.max(1, ...week.map((day) => {
            if (!day) return 0;
            const dateKey = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            return (byDate.get(dateKey) || []).length;
        })));
        const totalRows = 3 + lessonRowsByWeek.reduce((sum, lessonRows) => sum + 1 + lessonRows, 0);
        const sheet = XLSX.utils.aoa_to_sheet(Array.from({ length: totalRows }, () => Array(7).fill('')));
        sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } }];
        sheet['!cols'] = Array.from({ length: 7 }, () => ({ wch: 27 }));
        sheet['!rows'] = [{ hpt: 40 }, { hpt: 24 }, { hpt: 28 }, ...lessonRowsByWeek.flatMap((lessonRows) => [{ hpt: 22 }, ...Array.from({ length: lessonRows }, () => ({ hpt: 23 }))])];
        sheet['!ref'] = `A1:G${totalRows}`;

        cell(sheet, 0, 0, `${year} 年 ${month} 月课程月视图`, { font: font(20, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.header), alignment: alignment('center') });
        cell(sheet, 1, 0, `共 ${lessons.length} 节课 · ${byDate.size} 个有课日期 · 按每天实际课程逐行显示`, { font: font(10, { color: { rgb: COLORS.summaryText } }), fill: fill(COLORS.date), alignment: alignment('center') });
        WEEKDAYS.forEach((weekday, column) => {
            cell(sheet, 2, column, weekday, { font: font(11, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.weekday), alignment: alignment('center'), border: THIN_BORDER });
        });

        let dateRow = 3;
        weeks.forEach((week, weekIndex) => {
            const lessonRows = lessonRowsByWeek[weekIndex];
            week.forEach((day, weekdayIndex) => {
                const isWeekend = weekdayIndex >= 5;
                if (!day) {
                    for (let rowOffset = 0; rowOffset <= lessonRows; rowOffset += 1) cell(sheet, dateRow + rowOffset, weekdayIndex, '', { fill: fill(COLORS.outside), border: THIN_BORDER });
                    return;
                }
                const dateKey = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                const dayLessons = byDate.get(dateKey) || [];
                cell(sheet, dateRow, weekdayIndex, `${day} 日${dayLessons.length ? ` · ${dayLessons.length} 节` : ''}`, {
                    font: font(10, { bold: true, color: { rgb: COLORS.summaryText } }),
                    fill: fill(dayLessons.length ? COLORS.activeDate : (isWeekend ? COLORS.weekendDate : COLORS.date)), alignment: alignment('left'), border: THIN_BORDER,
                });
                for (let lessonIndex = 0; lessonIndex < lessonRows; lessonIndex += 1) {
                    const lesson = dayLessons[lessonIndex];
                    const row = dateRow + lessonIndex + 1;
                    if (!lesson) {
                        const noCourse = lessonIndex === 0 && !dayLessons.length ? '无课' : '';
                        cell(sheet, row, weekdayIndex, noCourse, { font: font(8, { color: { rgb: COLORS.muted } }), fill: fill(isWeekend ? COLORS.weekend : COLORS.white), alignment: alignment('left'), border: THIN_BORDER });
                        continue;
                    }
                    const start = timePart(lesson.lessonStartTime);
                    const end = timePart(lesson.lessonEndTime);
                    cell(sheet, row, weekdayIndex, `${start}–${end}  ${studentName(lesson)} · ${shortRoom(lesson.roomName)}`, { font: font(8, { bold: true, color: { rgb: COLORS.courseText } }), fill: fill(COLORS.course), alignment: alignment('left'), border: THIN_BORDER });
                }
            });
            dateRow += lessonRows + 1;
        });
        return sheet;
    }

    function appendSheetAtRow(target, source, rowOffset) {
        Object.entries(source).forEach(([address, value]) => {
            const match = address.match(/^([A-Z]+)(\d+)$/);
            if (!match) return;
            target[`${match[1]}${Number(match[2]) + rowOffset}`] = value;
        });
        (source['!merges'] || []).forEach((merge) => {
            target['!merges'].push({
                s: { r: merge.s.r + rowOffset, c: merge.s.c },
                e: { r: merge.e.r + rowOffset, c: merge.e.c },
            });
        });
        target['!rows'].push(...(source['!rows'] || []));
    }

    function createCombinedMonthSheet(months) {
        const sheet = XLSX.utils.aoa_to_sheet([[]]);
        sheet['!merges'] = [];
        sheet['!cols'] = Array.from({ length: 7 }, () => ({ wch: 27 }));
        sheet['!rows'] = [];
        let rowOffset = 0;
        months.forEach(({ year, month, lessons }, index) => {
            const monthSheet = createMonthSheet(year, month, lessons);
            appendSheetAtRow(sheet, monthSheet, rowOffset);
            rowOffset += monthSheet['!rows'].length;
            if (index < months.length - 1) {
                sheet['!rows'].push({ hpt: 10 }, { hpt: 10 });
                rowOffset += 2;
            }
        });
        sheet['!ref'] = `A1:G${Math.max(rowOffset, 1)}`;
        return sheet;
    }

    function valueFromLesson(lesson, fields, fallback = '未填写') {
        for (const field of fields) {
            const value = lesson[field];
            if (value !== undefined && value !== null && value !== '') {
                const displayed = displayValue(value, FIELD_NAME_MAP[field] || '');
                if (displayed) return String(displayed);
            }
        }
        return fallback;
    }

    function lessonDurationMinutes(lesson) {
        const start = timePart(lesson.lessonStartTime || lesson.startTime);
        const end = timePart(lesson.lessonEndTime || lesson.endTime);
        if (!start || !end) return 0;
        const [startHour, startMinute] = start.split(':').map(Number);
        const [endHour, endMinute] = end.split(':').map(Number);
        const minutes = endHour * 60 + endMinute - startHour * 60 - startMinute;
        return minutes > 0 ? minutes : 0;
    }

    function aggregateLessons(schedules, label) {
        const counts = new Map();
        schedules.forEach((lesson) => {
            const key = label(lesson);
            counts.set(key, (counts.get(key) || 0) + 1);
        });
        return [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'zh-CN'));
    }

    function addSummaryTable(sheet, row, column, title, header, entries) {
        sheet['!merges'].push({ s: { r: row, c: column }, e: { r: row, c: column + 1 } });
        cell(sheet, row, column, title, { font: font(11, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.header), alignment: alignment('center'), border: THIN_BORDER });
        [header, '节数'].forEach((value, index) => cell(sheet, row + 1, column + index, value, { font: font(10, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.weekday), alignment: alignment('center'), border: THIN_BORDER }));
        entries.forEach(([name, count], index) => {
            const style = { font: font(10), alignment: alignment(index === 0 ? 'left' : 'center'), border: THIN_BORDER, ...(index % 2 === 0 ? { fill: fill(COLORS.evenRow) } : {}) };
            cell(sheet, row + 2 + index, column, name, style);
            cell(sheet, row + 2 + index, column + 1, count, style);
        });
    }

    function createSummarySheet(schedules, startDate, endDate, failedDays) {
        const sheet = XLSX.utils.aoa_to_sheet(Array.from({ length: 1 }, () => Array(10).fill('')));
        sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }];
        sheet['!cols'] = [
            { wch: 26 }, { wch: 10 }, { wch: 4 }, { wch: 24 }, { wch: 10 },
            { wch: 4 }, { wch: 22 }, { wch: 10 }, { wch: 4 }, { wch: 22 },
        ];
        const activeDays = new Set(schedules.map((lesson) => lesson._date)).size;
        const studentCount = new Set(schedules.map((lesson) => studentName(lesson) || '未填写').filter(Boolean)).size;
        const totalMinutes = schedules.reduce((sum, lesson) => sum + lessonDurationMinutes(lesson), 0);
        const metrics = [
            ['导出范围', `${startDate} 至 ${endDate}`], ['课程记录', `${schedules.length} 节`],
            ['有课日期', `${activeDays} 天`], ['涉及学员', `${studentCount} 人`],
            ['累计时长', totalMinutes ? `${(totalMinutes / 60).toFixed(1)} 小时` : '未能计算'],
            ['读取状态', failedDays.length ? `有 ${failedDays.length} 天待重试` : '全部读取完成'],
        ];
        cell(sheet, 0, 0, '课表导出统计', { font: font(20, { bold: true, color: { rgb: COLORS.titleText } }), fill: fill(COLORS.header), alignment: alignment('center') });
        metrics.forEach(([label, value], index) => {
            const row = 2 + index;
            cell(sheet, row, 0, label, { font: font(10, { bold: true }), fill: fill(COLORS.date), alignment: alignment('center'), border: THIN_BORDER });
            sheet['!merges'].push({ s: { r: row, c: 1 }, e: { r: row, c: 4 } });
            cell(sheet, row, 1, value, { font: font(10), alignment: alignment('left'), border: THIN_BORDER });
        });
        if (failedDays.length) {
            const row = 9;
            sheet['!merges'].push({ s: { r: row, c: 0 }, e: { r: row, c: 9 } });
            cell(sheet, row, 0, `待重试日期：${failedDays.map(({ day }) => day).join('、')}`, { font: font(9, { color: { rgb: '9C5700' } }), fill: fill('FFEB9C'), alignment: alignment('left'), border: THIN_BORDER });
        }
        const tableRow = 12;
        const tables = [
            ['按学员统计', '学员', aggregateLessons(schedules, (lesson) => studentName(lesson) || '未填写')],
            ['按课程统计', '课程', aggregateLessons(schedules, (lesson) => valueFromLesson(lesson, ['courseName', 'className', 'name']))],
            ['按老师统计', '老师', aggregateLessons(schedules, (lesson) => valueFromLesson(lesson, ['teacherName', 'teacher', 'instructor']))],
            ['按校区统计', '校区', aggregateLessons(schedules, (lesson) => valueFromLesson(lesson, ['campus', 'schoolName', 'school', 'branch']))],
        ];
        tables.forEach(([title, header, entries], index) => addSummaryTable(sheet, tableRow, index * 3, title, header, entries));
        const maxEntries = Math.max(...tables.map(([, , entries]) => entries.length));
        sheet['!ref'] = `A1:J${tableRow + maxEntries + 2}`;
        sheet['!rows'] = [{ hpt: 36 }];
        return sheet;
    }

    function exportWorkbook(schedules, startDate, endDate, { combineMonthViews, failedDays = [], suffix = '' }) {
        const workbook = XLSX.utils.book_new();
        const grouped = new Map();
        schedules.forEach((lesson) => {
            const match = String(lesson._date || '').match(/^(\d{4})-(\d{2})/);
            if (!match) return;
            const key = `${match[1]}-${match[2]}`;
            const entries = grouped.get(key) || [];
            entries.push(lesson);
            grouped.set(key, entries);
        });
        const months = monthsInRange(startDate, endDate).map(({ year, month }) => ({ year, month, lessons: grouped.get(`${year}-${String(month).padStart(2, '0')}`) || [] }));
        if (combineMonthViews) {
            XLSX.utils.book_append_sheet(workbook, createCombinedMonthSheet(months), '月视图');
        } else {
            months.forEach(({ year, month, lessons }) => XLSX.utils.book_append_sheet(workbook, createMonthSheet(year, month, lessons), `${year}年${month}月月视图`));
        }
        XLSX.utils.book_append_sheet(workbook, createSummarySheet(schedules, startDate, endDate, failedDays), '统计');
        XLSX.utils.book_append_sheet(workbook, createDetailSheet(schedules), '详细课表');
        XLSX.writeFile(workbook, `课表_${startDate}_至_${endDate}${suffix}.xlsx`);
    }

    function showExportDialog() {
        if (document.getElementById('xdf-schedule-export-dialog')) return;
        const preferences = loadPreferences();
        const [defaultStart, defaultEnd] = preferences.startDate && preferences.endDate ? [preferences.startDate, preferences.endDate] : currentMonthRange();
        const dialog = document.createElement('div');
        dialog.id = 'xdf-schedule-export-dialog';
        dialog.innerHTML = `
            <div class="xdf-export-card" role="dialog" aria-modal="true" aria-labelledby="xdf-export-title">
                <div class="xdf-export-glow xdf-export-glow-one"></div><div class="xdf-export-glow xdf-export-glow-two"></div>
                <div class="xdf-export-content">
                    <div class="xdf-export-heading">
                        <div class="xdf-export-title-group"><span class="xdf-export-app-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M7 3.75h7.2L18.5 8v11.75A1.25 1.25 0 0 1 17.25 21H7a1.5 1.5 0 0 1-1.5-1.5V5.25A1.5 1.5 0 0 1 7 3.75Z"/><path d="M14 3.9V8h4.15M8.5 12h7M8.5 15h7M8.5 18h4"/></svg></span><div><span class="xdf-export-eyebrow">XDF SCHEDULE</span><h2 id="xdf-export-title">导出课表</h2><p>选择日期范围，即可生成完整的 Excel 课表。</p></div></div>
                        <button type="button" class="xdf-export-close" aria-label="关闭"><svg viewBox="0 0 20 20" aria-hidden="true"><path d="m5.5 5.5 9 9m0-9-9 9"/></svg></button>
                    </div>
                    <section class="xdf-export-section" aria-labelledby="xdf-export-range-label">
                        <div class="xdf-export-section-heading"><span id="xdf-export-range-label">日期范围</span><span>快速选择</span></div>
                        <div class="xdf-export-presets" aria-label="日期快捷选择"><button type="button" data-preset="this-week">本周</button><button type="button" data-preset="next-week">下周</button><button type="button" data-preset="this-month">本月</button><button type="button" data-preset="next-month">下月</button><button type="button" data-preset="this-term">本学期</button></div>
                        <div class="xdf-export-fields"><label><span>开始日期</span><input id="xdf-export-start" type="date" required value="${defaultStart}"></label><span class="xdf-export-range-arrow" aria-hidden="true">→</span><label><span>结束日期</span><input id="xdf-export-end" type="date" required value="${defaultEnd}"></label></div>
                    </section>
                    <label class="xdf-export-option"><input id="xdf-export-combine-months" type="checkbox" ${preferences.combineMonthViews ? 'checked' : ''}><span class="xdf-export-check"><svg viewBox="0 0 16 16"><path d="m3.5 8.2 2.8 2.8 6.2-6"/></svg></span><span><strong>合并多个月份</strong><small>连续放在同一个工作表中，查看更方便</small></span></label>
                    <p id="xdf-export-status" class="xdf-export-status" aria-live="polite"></p>
                    <div class="xdf-export-actions"><button type="button" class="xdf-export-cancel">取消</button><button type="button" class="xdf-export-retry" hidden>重试失败日期</button><button type="button" class="xdf-export-submit"><span>开始导出</span><svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 3v9m0 0 3.5-3.5M10 12 6.5 8.5M4 14.5v1A1.5 1.5 0 0 0 5.5 17h9a1.5 1.5 0 0 0 1.5-1.5v-1"/></svg></button></div>
                </div>
            </div>`;
        document.body.appendChild(dialog);
        const startInput = dialog.querySelector('#xdf-export-start');
        const endInput = dialog.querySelector('#xdf-export-end');
        const combineInput = dialog.querySelector('#xdf-export-combine-months');
        const status = dialog.querySelector('#xdf-export-status');
        const submit = dialog.querySelector('.xdf-export-submit');
        const retry = dialog.querySelector('.xdf-export-retry');
        const presetButtons = [...dialog.querySelectorAll('[data-preset]')];
        let lastExport = null;
        const close = () => dialog.remove();
        dialog.querySelector('.xdf-export-close').addEventListener('click', close);
        dialog.querySelector('.xdf-export-cancel').addEventListener('click', close);
        dialog.addEventListener('click', (event) => { if (event.target === dialog) close(); });
        dialog.addEventListener('keydown', (event) => { if (event.key === 'Escape') close(); });
        presetButtons.forEach((button) => button.addEventListener('click', () => {
            const [startDate, endDate] = rangeForPreset(button.dataset.preset);
            startInput.value = startDate;
            endInput.value = endDate;
            presetButtons.forEach((item) => {
                const active = item === button;
                item.classList.toggle('is-active', active);
                item.setAttribute('aria-pressed', String(active));
            });
        }));
        [startInput, endInput].forEach((input) => input.addEventListener('input', () => presetButtons.forEach((item) => {
            item.classList.remove('is-active');
            item.setAttribute('aria-pressed', 'false');
        })));
        dialog.querySelector('.xdf-export-close').focus();

        function setBusy(busy) {
            [startInput, endInput, combineInput, submit, retry, ...presetButtons].forEach((control) => { control.disabled = busy; });
        }

        function exportResult(result, suffix = '') {
            const { schedules, failedDays, startDate, endDate, combineMonthViews } = result;
            if (!schedules.length) return false;
            status.textContent = '正在生成月视图、统计和详细课表…';
            exportWorkbook(schedules, startDate, endDate, { combineMonthViews, failedDays, suffix });
            const elapsedSeconds = ((performance.now() - result.startedAt) / 1000).toFixed(1);
            status.textContent = failedDays.length
                ? `已导出 ${schedules.length} 条课程记录，${failedDays.length} 天未读取成功，可点击“重试失败日期”。耗时 ${elapsedSeconds} 秒。`
                : `已导出 ${schedules.length} 条课程记录，耗时 ${elapsedSeconds} 秒，文件已开始下载。`;
            retry.hidden = !failedDays.length;
            return true;
        }

        submit.addEventListener('click', async () => {
            const rangeStart = startInput.value;
            const rangeEnd = endInput.value;
            if (!rangeStart || !rangeEnd) return void (status.textContent = '请选择开始日期和结束日期。');
            if (rangeStart > rangeEnd) return void (status.textContent = '结束日期不能早于开始日期。');
            const combineMonthViews = combineInput.checked;
            savePreferences({ startDate: rangeStart, endDate: rangeEnd, combineMonthViews });
            setBusy(true);
            retry.hidden = true;
            status.textContent = '正在查询有课日期…';
            const startedAt = performance.now();
            try {
                const result = await fetchSchedules(rangeStart, rangeEnd, (message) => { status.textContent = message; });
                lastExport = { ...result, startDate: rangeStart, endDate: rangeEnd, combineMonthViews, startedAt };
                if (!exportResult(lastExport, result.failedDays.length ? '（部分导出）' : '')) {
                    status.textContent = result.failedDays.length ? `没有课程成功读取；${result.failedDays.length} 天可重试。` : '所选日期范围没有课程，未生成文件。';
                    retry.hidden = !result.failedDays.length;
                }
            } catch (error) {
                console.error('[XDF 课表导出]', error);
                status.textContent = `导出失败：${error.message || error}`;
            } finally {
                setBusy(false);
            }
        });

        retry.addEventListener('click', async () => {
            if (!lastExport?.failedDays.length) return;
            setBusy(true);
            status.textContent = `正在重试 ${lastExport.failedDays.length} 个失败日期…`;
            const startedAt = performance.now();
            try {
                const result = await fetchLessonDetails(lastExport.failedDays, (message) => { status.textContent = `重试中：${message}`; });
                lastExport.schedules = sortSchedules([...lastExport.schedules, ...result.schedules]);
                lastExport.failedDays = result.failedDays;
                lastExport.startedAt = startedAt;
                exportResult(lastExport, result.failedDays.length ? '（部分补全）' : '（已补全）');
            } catch (error) {
                console.error('[XDF 课表导出]', error);
                status.textContent = `重试失败：${error.message || error}`;
            } finally {
                setBusy(false);
            }
        });
    }

    function addExportButton() {
        if (document.getElementById('xdf-schedule-export-button')) return;
        const button = document.createElement('button');
        button.id = 'xdf-schedule-export-button';
        button.type = 'button';
        button.innerHTML = '<span class="xdf-export-button-icon"><svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 2.75v9m0 0 3.5-3.5M10 11.75l-3.5-3.5M3.75 14.25v1A1.75 1.75 0 0 0 5.5 17h9a1.75 1.75 0 0 0 1.75-1.75v-1"/></svg></span><span>导出课表</span>';
        button.title = '导出月视图、统计和详细课表';
        button.addEventListener('click', showExportDialog);
        document.body.appendChild(button);
        const style = document.createElement('style');
        style.textContent = `
            #xdf-schedule-export-button { position: fixed; right: 24px; bottom: 24px; z-index: 99999; display: inline-flex; align-items: center; gap: 9px; min-height: 48px; padding: 7px 17px 7px 8px; border: 1px solid rgb(255 255 255 / 72%); border-radius: 999px; background: linear-gradient(135deg, rgb(255 255 255 / 86%), rgb(239 246 255 / 68%)); color: #17426d; box-shadow: 0 12px 32px rgb(15 54 92 / 18%), inset 0 1px 0 rgb(255 255 255 / 86%); backdrop-filter: blur(22px) saturate(180%); -webkit-backdrop-filter: blur(22px) saturate(180%); cursor: pointer; font: 600 14px/1 -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang SC", sans-serif; letter-spacing: -.1px; transition: transform .2s ease, box-shadow .2s ease, background .2s ease; }
            #xdf-schedule-export-button:hover { transform: translateY(-2px) scale(1.015); background: linear-gradient(135deg, rgb(255 255 255 / 96%), rgb(232 243 255 / 82%)); box-shadow: 0 16px 38px rgb(15 54 92 / 23%), inset 0 1px 0 #fff; }
            #xdf-schedule-export-button:active { transform: translateY(0) scale(.98); }
            #xdf-schedule-export-button:focus-visible { outline: 3px solid rgb(20 122 255 / 28%); outline-offset: 3px; }
            .xdf-export-button-icon { display: grid; width: 34px; height: 34px; place-items: center; border-radius: 50%; background: linear-gradient(145deg, #3aa0ff, #0878f9 70%); color: #fff; box-shadow: 0 5px 14px rgb(5 116 246 / 28%), inset 0 1px 1px rgb(255 255 255 / 45%); }
            .xdf-export-button-icon svg { width: 18px; fill: none; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.7; }
            #xdf-schedule-export-dialog { position: fixed; inset: 0; z-index: 100000; display: grid; place-items: center; overflow-y: auto; padding: 24px; background: radial-gradient(circle at 15% 10%, rgb(86 175 255 / 25%), transparent 35%), radial-gradient(circle at 88% 90%, rgb(131 100 255 / 20%), transparent 35%), rgb(12 24 43 / 42%); backdrop-filter: blur(14px) saturate(135%); -webkit-backdrop-filter: blur(14px) saturate(135%); animation: xdf-export-fade-in .22s ease both; }
            #xdf-schedule-export-dialog * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang SC", "Segoe UI", sans-serif; }
            .xdf-export-card { position: relative; isolation: isolate; width: min(560px, 100%); overflow: hidden; border: 1px solid rgb(255 255 255 / 70%); border-radius: 28px; background: linear-gradient(145deg, rgb(255 255 255 / 84%), rgb(244 249 255 / 66%)); color: #172b43; box-shadow: 0 32px 90px rgb(7 24 48 / 32%), inset 0 1px 0 rgb(255 255 255 / 90%); backdrop-filter: blur(36px) saturate(180%); -webkit-backdrop-filter: blur(36px) saturate(180%); animation: xdf-export-card-in .28s cubic-bezier(.2,.8,.2,1) both; }
            .xdf-export-content { position: relative; z-index: 2; padding: 28px; }
            .xdf-export-glow { position: absolute; z-index: 0; width: 220px; height: 220px; border-radius: 50%; filter: blur(8px); opacity: .5; pointer-events: none; }.xdf-export-glow-one { top: -140px; right: -75px; background: radial-gradient(circle, rgb(101 188 255 / 62%), transparent 68%); }.xdf-export-glow-two { bottom: -170px; left: -80px; background: radial-gradient(circle, rgb(143 114 255 / 38%), transparent 68%); }
            .xdf-export-heading { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; margin-bottom: 24px; }.xdf-export-title-group { display: flex; gap: 14px; align-items: center; }.xdf-export-app-icon { display: grid; flex: 0 0 auto; width: 48px; height: 48px; place-items: center; border: 1px solid rgb(255 255 255 / 72%); border-radius: 15px; background: linear-gradient(145deg, rgb(69 169 255 / 92%), rgb(29 107 238 / 88%)); color: #fff; box-shadow: 0 10px 24px rgb(18 108 220 / 25%), inset 0 1px 1px rgb(255 255 255 / 44%); }.xdf-export-app-icon svg { width: 27px; fill: rgb(255 255 255 / 12%); stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.55; }.xdf-export-eyebrow { display: block; margin-bottom: 3px; color: #5180ad; font-size: 9px; font-weight: 750; letter-spacing: 1.6px; }.xdf-export-heading h2 { margin: 0; font-size: 23px; font-weight: 700; line-height: 1.2; letter-spacing: -.6px; }.xdf-export-heading p { margin: 5px 0 0; color: #667a90; font-size: 13px; line-height: 1.45; }
            .xdf-export-close { display: grid; flex: 0 0 auto; width: 34px; height: 34px; place-items: center; border: 1px solid rgb(255 255 255 / 64%); border-radius: 50%; background: rgb(255 255 255 / 42%); color: #62758a; box-shadow: inset 0 1px 0 rgb(255 255 255 / 70%); cursor: pointer; transition: .18s ease; }.xdf-export-close:hover { background: rgb(255 255 255 / 76%); color: #273c53; transform: rotate(3deg); }.xdf-export-close svg { width: 18px; fill: none; stroke: currentColor; stroke-linecap: round; stroke-width: 1.7; }
            .xdf-export-section { padding: 18px; border: 1px solid rgb(255 255 255 / 62%); border-radius: 20px; background: linear-gradient(145deg, rgb(255 255 255 / 48%), rgb(255 255 255 / 24%)); box-shadow: inset 0 1px 0 rgb(255 255 255 / 68%), 0 8px 28px rgb(44 89 134 / 7%); }.xdf-export-section-heading { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 11px; color: #233c57; font-size: 13px; font-weight: 700; }.xdf-export-section-heading span:last-child { color: #8291a2; font-size: 11px; font-weight: 500; }
            .xdf-export-presets { display: grid; grid-template-columns: repeat(5, 1fr); gap: 7px; margin-bottom: 16px; }.xdf-export-presets button { min-height: 34px; border: 1px solid rgb(132 158 185 / 20%); border-radius: 10px; background: rgb(255 255 255 / 42%); color: #54708d; cursor: pointer; font-size: 12px; font-weight: 550; transition: .18s ease; }.xdf-export-presets button:hover { border-color: rgb(47 135 236 / 30%); background: rgb(255 255 255 / 74%); color: #176fcb; transform: translateY(-1px); }.xdf-export-presets button.is-active { border-color: rgb(25 125 245 / 28%); background: linear-gradient(145deg, #3899f5, #1674e8); color: #fff; box-shadow: 0 5px 15px rgb(20 116 227 / 22%), inset 0 1px 0 rgb(255 255 255 / 35%); }
            .xdf-export-fields { display: grid; grid-template-columns: 1fr 22px 1fr; gap: 8px; align-items: end; }.xdf-export-fields label { display: grid; gap: 7px; color: #51687f; font-size: 11px; font-weight: 600; }.xdf-export-fields input { width: 100%; min-height: 44px; padding: 9px 11px; border: 1px solid rgb(111 137 164 / 23%); border-radius: 12px; background: rgb(255 255 255 / 55%); color: #18324e; box-shadow: inset 0 1px 2px rgb(31 66 103 / 5%); font: 500 13px/1 -apple-system, BlinkMacSystemFont, "SF Pro Text", sans-serif; transition: .18s ease; color-scheme: light; }.xdf-export-fields input:hover { background: rgb(255 255 255 / 78%); }.xdf-export-fields input:focus { border-color: rgb(26 124 242 / 48%); outline: 3px solid rgb(26 124 242 / 13%); background: rgb(255 255 255 / 88%); }.xdf-export-range-arrow { display: grid; height: 44px; place-items: center; color: #91a2b4; font-size: 15px; }
            .xdf-export-option { display: flex; align-items: center; gap: 12px; margin-top: 14px; padding: 14px 15px; border: 1px solid transparent; border-radius: 17px; color: #344e69; cursor: pointer; transition: .18s ease; }.xdf-export-option:hover { border-color: rgb(255 255 255 / 58%); background: rgb(255 255 255 / 30%); }.xdf-export-option input { position: absolute; opacity: 0; pointer-events: none; }.xdf-export-check { display: grid; flex: 0 0 auto; width: 22px; height: 22px; place-items: center; border: 1px solid rgb(106 133 159 / 38%); border-radius: 7px; background: rgb(255 255 255 / 45%); box-shadow: inset 0 1px 0 rgb(255 255 255 / 60%); color: transparent; transition: .18s ease; }.xdf-export-check svg { width: 15px; fill: none; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 2; }.xdf-export-option input:checked + .xdf-export-check { border-color: #2380ec; background: linear-gradient(145deg, #429eff, #1375e8); color: #fff; box-shadow: 0 5px 12px rgb(16 111 221 / 20%), inset 0 1px 0 rgb(255 255 255 / 35%); }.xdf-export-option input:focus-visible + .xdf-export-check { outline: 3px solid rgb(25 124 241 / 18%); outline-offset: 2px; }.xdf-export-option strong, .xdf-export-option small { display: block; }.xdf-export-option strong { margin-bottom: 2px; font-size: 13px; }.xdf-export-option small { color: #788b9e; font-size: 11px; font-weight: 400; }
            .xdf-export-status { min-height: 20px; margin: 12px 5px 0; color: #657b91; font-size: 12px; line-height: 1.55; }.xdf-export-status:not(:empty) { padding: 8px 11px; border: 1px solid rgb(255 255 255 / 55%); border-radius: 11px; background: rgb(255 255 255 / 30%); }
            .xdf-export-actions { display: flex; justify-content: flex-end; gap: 9px; margin-top: 17px; }.xdf-export-actions button { min-height: 42px; padding: 0 16px; border: 1px solid rgb(112 137 162 / 20%); border-radius: 13px; background: rgb(255 255 255 / 40%); color: #49627b; cursor: pointer; font: 600 13px/1 -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang SC", sans-serif; transition: .18s ease; }.xdf-export-actions button:hover { background: rgb(255 255 255 / 72%); transform: translateY(-1px); }.xdf-export-actions .xdf-export-submit { display: inline-flex; min-width: 124px; align-items: center; justify-content: center; gap: 7px; border-color: rgb(42 130 235 / 42%); background: linear-gradient(145deg, #409dfb, #1175e8); color: #fff; box-shadow: 0 9px 22px rgb(13 105 214 / 25%), inset 0 1px 0 rgb(255 255 255 / 36%); }.xdf-export-submit svg { width: 17px; fill: none; stroke: currentColor; stroke-linecap: round; stroke-linejoin: round; stroke-width: 1.6; }.xdf-export-actions .xdf-export-submit:hover { background: linear-gradient(145deg, #53a9ff, #167cf0); box-shadow: 0 11px 25px rgb(13 105 214 / 31%), inset 0 1px 0 rgb(255 255 255 / 40%); }.xdf-export-actions .xdf-export-retry { border-color: rgb(224 155 30 / 30%); background: rgb(255 247 221 / 62%); color: #92540b; }.xdf-export-actions button:disabled, .xdf-export-presets button:disabled { cursor: wait; opacity: .56; transform: none; }
            @keyframes xdf-export-fade-in { from { opacity: 0; } to { opacity: 1; } } @keyframes xdf-export-card-in { from { opacity: 0; transform: translateY(12px) scale(.975); } to { opacity: 1; transform: translateY(0) scale(1); } }
            @media (max-width: 600px) { #xdf-schedule-export-dialog { align-items: end; padding: 12px; }.xdf-export-card { border-radius: 25px; }.xdf-export-content { padding: 22px 18px 18px; }.xdf-export-heading p { max-width: 250px; }.xdf-export-presets { grid-template-columns: repeat(3, 1fr); }.xdf-export-fields { grid-template-columns: 1fr; }.xdf-export-range-arrow { display: none; }.xdf-export-actions { display: grid; grid-template: "cancel submit" auto "retry retry" auto / 1fr 1fr; }.xdf-export-actions .xdf-export-cancel { grid-area: cancel; }.xdf-export-actions .xdf-export-retry { grid-area: retry; }.xdf-export-actions .xdf-export-submit { grid-area: submit; min-width: 0; } }
            @media (max-width: 380px) { #xdf-schedule-export-button { right: 14px; bottom: 14px; }.xdf-export-app-icon { display: none; }.xdf-export-title-group { gap: 0; }.xdf-export-heading p { font-size: 12px; }.xdf-export-presets { grid-template-columns: repeat(2, 1fr); } }
            @media (prefers-reduced-motion: reduce) { #xdf-schedule-export-dialog, .xdf-export-card { animation: none; } #xdf-schedule-export-button, .xdf-export-actions button, .xdf-export-presets button { transition: none; } }
        `;
        document.head.appendChild(style);
    }

    addExportButton();
}());
