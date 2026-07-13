// ==UserScript==
// @name         XDF 课表导出
// @namespace    https://github.com/nowscott/XdfScheduleCrawler
// @version      1.2.0
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
                <div class="xdf-export-heading"><div><h2 id="xdf-export-title">导出课表 Excel</h2><p>月视图优先、统计居中、详细课表最后；详情最多三路并发读取。</p></div><button type="button" class="xdf-export-close" aria-label="关闭">×</button></div>
                <div class="xdf-export-presets" aria-label="日期快捷选择"><button type="button" data-preset="this-week">本周</button><button type="button" data-preset="next-week">下周</button><button type="button" data-preset="this-month">本月</button><button type="button" data-preset="next-month">下月</button><button type="button" data-preset="this-term">本学期</button></div>
                <div class="xdf-export-fields"><label>开始日期<input id="xdf-export-start" type="date" required value="${defaultStart}"></label><label>结束日期<input id="xdf-export-end" type="date" required value="${defaultEnd}"></label></div>
                <label class="xdf-export-option"><input id="xdf-export-combine-months" type="checkbox" ${preferences.combineMonthViews ? 'checked' : ''}> 将多个月视图连续放在同一个工作表（默认）</label>
                <p id="xdf-export-status" class="xdf-export-status" aria-live="polite"></p>
                <div class="xdf-export-actions"><button type="button" class="xdf-export-cancel">取消</button><button type="button" class="xdf-export-retry" hidden>重试失败日期</button><button type="button" class="xdf-export-submit">开始导出</button></div>
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
        presetButtons.forEach((button) => button.addEventListener('click', () => {
            const [startDate, endDate] = rangeForPreset(button.dataset.preset);
            startInput.value = startDate;
            endInput.value = endDate;
        }));

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
        button.textContent = '导出课表 Excel';
        button.title = '导出月视图、统计和详细课表';
        Object.assign(button.style, { position: 'fixed', right: '24px', bottom: '24px', zIndex: '99999', border: '0', borderRadius: '8px', padding: '10px 14px', cursor: 'pointer', background: '#1677ff', color: '#fff', boxShadow: '0 4px 14px rgb(0 0 0 / 18%)' });
        button.addEventListener('click', showExportDialog);
        document.body.appendChild(button);
        const style = document.createElement('style');
        style.textContent = `
            #xdf-schedule-export-dialog { position: fixed; inset: 0; z-index: 100000; display: grid; place-items: center; padding: 20px; background: rgb(15 23 42 / 45%); }
            #xdf-schedule-export-dialog * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
            .xdf-export-card { width: min(500px, 100%); padding: 24px; border-radius: 16px; background: #fff; color: #172033; box-shadow: 0 18px 48px rgb(15 23 42 / 25%); }
            .xdf-export-heading { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; margin-bottom: 22px; }.xdf-export-heading h2 { margin: 0; font-size: 20px; line-height: 1.3; }.xdf-export-heading p { margin: 6px 0 0; color: #667085; font-size: 14px; }
            .xdf-export-close { border: 0; padding: 0 4px; background: transparent; color: #667085; cursor: pointer; font-size: 28px; line-height: 1; }.xdf-export-presets { display: flex; flex-wrap: wrap; gap: 8px; margin: -5px 0 16px; }.xdf-export-presets button { border: 1px solid #bfd7ff; border-radius: 999px; padding: 5px 10px; background: #f4f8ff; color: #175cd3; cursor: pointer; font-size: 13px; }.xdf-export-fields { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }.xdf-export-fields label { display: grid; gap: 7px; color: #344054; font-size: 14px; font-weight: 600; }.xdf-export-fields input { width: 100%; min-height: 40px; padding: 8px 10px; border: 1px solid #d0d5dd; border-radius: 8px; color: #172033; font: inherit; }.xdf-export-fields input:focus { outline: 3px solid rgb(22 119 255 / 18%); border-color: #1677ff; }.xdf-export-option { display: flex; align-items: center; gap: 8px; margin-top: 14px; color: #344054; font-size: 13px; cursor: pointer; }
            .xdf-export-status { min-height: 20px; margin: 16px 0 0; color: #667085; font-size: 13px; line-height: 1.5; }.xdf-export-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; }.xdf-export-actions button { min-height: 38px; padding: 0 14px; border: 1px solid #d0d5dd; border-radius: 8px; background: #fff; color: #344054; cursor: pointer; font: inherit; }.xdf-export-actions .xdf-export-submit { border-color: #1677ff; background: #1677ff; color: #fff; }.xdf-export-actions .xdf-export-retry { border-color: #fdb022; background: #fffaeb; color: #93370d; }.xdf-export-actions button:disabled, .xdf-export-presets button:disabled { cursor: wait; opacity: .7; }
            @media (max-width: 480px) { .xdf-export-fields { grid-template-columns: 1fr; } }
        `;
        document.head.appendChild(style);
    }

    addExportButton();
}());
