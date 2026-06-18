'use strict';

const ExcelJS = require('exceljs');

// 配色 (ARGB)
const COLOR_HOLIDAY_FILL = 'FFE0E0E0'; // 土日祝の列の薄灰色
const COLOR_HEADER_FILL = 'FF2F5496'; // 濃い青 (ヘッダー)
const COLOR_HEADER_FONT = 'FFFFFFFF';
const COLOR_TODAY_BORDER = 'FFFF0000';
const COLOR_GRID = 'FFD9D9D9';

// 担当者ごとに割り当てる色パレット (バー用)
const BAR_COLOR_PALETTE = [
  'FF4472C4', // 青
  'FFED7D31', // オレンジ
  'FF70AD47', // 緑
  'FFFFC000', // 黄
  'FF9E480E', // 茶
  'FF636EFA', // 紫青
  'FFC00000', // 赤
  'FF7030A0', // 紫
  'FF00B0F0', // 水色
  'FF548235', // 深緑
];

function colorForAssignee(assigneeColorMap, assignee) {
  if (!assigneeColorMap.has(assignee)) {
    const idx = assigneeColorMap.size % BAR_COLOR_PALETTE.length;
    assigneeColorMap.set(assignee, BAR_COLOR_PALETTE[idx]);
  }
  return assigneeColorMap.get(assignee);
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function formatYmdSlash(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

const WEEKDAY_JP = ['日', '月', '火', '水', '木', '金', '土'];

/**
 * 週の開始日(月曜)を取得する
 */
function startOfWeekMonday(date) {
  const d = new Date(date);
  const dow = d.getDay(); // 0=日
  const diff = dow === 0 ? -6 : 1 - dow;
  d.setDate(d.getDate() + diff);
  return d;
}

/**
 * ガントチャート付きExcel(C)を生成して保存する。
 *
 * @param {Array} scheduledTasks - scheduler.scheduleTasks() の戻り値
 * @param {object} opts
 *   opts.granularity: 'day' | 'week'
 *   opts.isHoliday: function(date) => boolean
 *   opts.outputPath: string
 *   opts.projectStart: Date
 */
async function buildGanttExcel(scheduledTasks, opts) {
  const { granularity, isHoliday, outputPath, projectStart } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'gantt-tool';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('ガントチャート', {
    views: [{ state: 'frozen', xSplit: 5, ySplit: 2 }],
    pageSetup: {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9, // A4
      margins: { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2 },
    },
  });

  // 列構成: タスク名, 担当者, 工数(日), 開始日, 終了日, [タイムライン...]
  const fixedHeaders = ['タスク名', '担当者', '工数(日)', '開始日', '終了日'];
  const FIXED_COL_COUNT = fixedHeaders.length;

  // タイムラインの範囲を決定
  const minStart = scheduledTasks.reduce(
    (min, t) => (t.startDate < min ? t.startDate : min),
    scheduledTasks[0].startDate
  );
  const maxEnd = scheduledTasks.reduce(
    (max, t) => (t.endDate > max ? t.endDate : max),
    scheduledTasks[0].endDate
  );

  const timelineStart = granularity === 'week' ? startOfWeekMonday(minStart) : new Date(minStart);
  let timelineEnd = new Date(maxEnd);

  // タイムラインの列見出しリストを作る
  const timelineCols = []; // { date(or weekStart), label, isHolidayCol(day only) }
  if (granularity === 'day') {
    let cur = new Date(timelineStart);
    while (cur <= timelineEnd) {
      timelineCols.push({
        date: new Date(cur),
        label: `${cur.getMonth() + 1}/${cur.getDate()}\n(${WEEKDAY_JP[cur.getDay()]})`,
        isHolidayCol: isHoliday(cur),
      });
      cur = addDays(cur, 1);
    }
  } else {
    let cur = new Date(timelineStart);
    while (cur <= timelineEnd) {
      const weekEnd = addDays(cur, 6);
      timelineCols.push({
        date: new Date(cur),
        weekEnd: new Date(weekEnd),
        label: `${cur.getMonth() + 1}/${cur.getDate()}〜`,
      });
      cur = addDays(cur, 7);
    }
  }

  // --- ヘッダー行 (2行: 月情報 + 日/週情報) ---
  // 1行目: 固定列見出し + タイムライン上の月ラベル(連続マージ)
  // 2行目: (固定列は空) + タイムライン上の日/週ラベル
  const headerRow1 = sheet.getRow(1);
  const headerRow2 = sheet.getRow(2);

  fixedHeaders.forEach((h, i) => {
    const cell = headerRow1.getCell(i + 1);
    cell.value = h;
    sheet.mergeCells(1, i + 1, 2, i + 1);
  });

  // 月ラベル用にグルーピング
  let monthGroupStartCol = FIXED_COL_COUNT + 1;
  let prevMonthKey = null;
  timelineCols.forEach((tc, idx) => {
    const col = FIXED_COL_COUNT + 1 + idx;
    const monthKey = `${tc.date.getFullYear()}-${tc.date.getMonth()}`;
    if (prevMonthKey === null) {
      prevMonthKey = monthKey;
      monthGroupStartCol = col;
    } else if (monthKey !== prevMonthKey) {
      // 直前グループを確定 -> マージ
      const startCol = monthGroupStartCol;
      const endCol = col - 1;
      const [y, m] = prevMonthKey.split('-');
      headerRow1.getCell(startCol).value = `${y}年${Number(m) + 1}月`;
      if (endCol > startCol) sheet.mergeCells(1, startCol, 1, endCol);
      prevMonthKey = monthKey;
      monthGroupStartCol = col;
    }
    // 2行目: 日/週ラベル
    headerRow2.getCell(col).value = tc.label;
  });
  // 最後のグループを確定
  if (timelineCols.length > 0) {
    const startCol = monthGroupStartCol;
    const endCol = FIXED_COL_COUNT + timelineCols.length;
    const [y, m] = prevMonthKey.split('-');
    headerRow1.getCell(startCol).value = `${y}年${Number(m) + 1}月`;
    if (endCol > startCol) sheet.mergeCells(1, startCol, 1, endCol);
  }

  // ヘッダー装飾
  [headerRow1, headerRow2].forEach((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_HEADER_FILL } };
      cell.font = { color: { argb: COLOR_HEADER_FONT }, bold: true, size: 10 };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = borderAll(COLOR_GRID);
    });
  });
  headerRow1.height = 18;
  headerRow2.height = 30;

  // --- 列幅 ---
  for (let i = 0; i < FIXED_COL_COUNT; i++) {
    const widths = [22, 10, 9, 11, 11];
    sheet.getColumn(i + 1).width = widths[i];
  }
  const timelineColWidth = granularity === 'day' ? 4.2 : 8.5;
  for (let i = 0; i < timelineCols.length; i++) {
    sheet.getColumn(FIXED_COL_COUNT + 1 + i).width = timelineColWidth;
  }

  // --- データ行 ---
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const assigneeColorMap = new Map();
  scheduledTasks.forEach((task, idx) => {
    const rowIdx = 3 + idx;
    const row = sheet.getRow(rowIdx);
    row.getCell(1).value = task.name;
    row.getCell(2).value = task.assignee;
    row.getCell(3).value = task.days;
    row.getCell(4).value = formatYmdSlash(task.startDate);
    row.getCell(5).value = formatYmdSlash(task.endDate);

    for (let c = 1; c <= FIXED_COL_COUNT; c++) {
      const cell = row.getCell(c);
      cell.border = borderAll(COLOR_GRID);
      cell.alignment = { vertical: 'middle', horizontal: c === 1 ? 'left' : 'center' };
      cell.font = { size: 10 };
    }

    const barColor = colorForAssignee(assigneeColorMap, task.assignee);

    timelineCols.forEach((tc, ci) => {
      const col = FIXED_COL_COUNT + 1 + ci;
      const cell = row.getCell(col);

      const isTodayCol =
        granularity === 'day'
          ? sameDate(tc.date, today)
          : today >= tc.date && today <= tc.weekEnd;

      cell.border = isTodayCol
        ? borderAllColored(COLOR_TODAY_BORDER, 'medium')
        : borderAll(COLOR_GRID);

      let covered = false;
      if (granularity === 'day') {
        covered = tc.date >= task.startDate && tc.date <= task.endDate;
        if (!covered && tc.isHolidayCol) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_HOLIDAY_FILL } };
        }
      } else {
        // 週単位: タスク期間とその週が重なっていれば塗る
        covered = task.startDate <= tc.weekEnd && task.endDate >= tc.date;
      }

      if (covered) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: barColor } };
      }
    });

    row.height = 16;
  });

  // 今日の列をヘッダーにも示す(該当列があれば見出しに▼マーカーを追加)
  timelineCols.forEach((tc, ci) => {
    const isTodayCol =
      granularity === 'day' ? sameDate(tc.date, today) : today >= tc.date && today <= tc.weekEnd;
    if (isTodayCol) {
      const col = FIXED_COL_COUNT + 1 + ci;
      const cell = headerRow2.getCell(col);
      cell.value = `▼今日\n${cell.value}`;
    }
  });

  // --- 凡例シート的に、担当者と色の対応を右下に簡易表示 ---
  const legendStartRow = 3 + scheduledTasks.length + 2;
  sheet.getCell(legendStartRow, 1).value = '担当者カラー凡例';
  sheet.getCell(legendStartRow, 1).font = { bold: true, size: 10 };
  let li = 0;
  for (const [assignee, color] of assigneeColorMap.entries()) {
    const r = legendStartRow + 1 + li;
    const swatch = sheet.getCell(r, 1);
    swatch.value = '　　';
    swatch.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
    sheet.getCell(r, 2).value = assignee;
    li++;
  }
  if (granularity === 'day') {
    const r = legendStartRow + 1 + li + 1;
    const swatch = sheet.getCell(r, 1);
    swatch.value = '　　';
    swatch.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_HOLIDAY_FILL } };
    sheet.getCell(r, 2).value = '土日・祝日・年末年始休';
  }

  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

function borderAll(argb) {
  const style = { style: 'thin', color: { argb } };
  return { top: style, left: style, bottom: style, right: style };
}

function borderAllColored(argb, weight) {
  const style = { style: weight, color: { argb } };
  return { top: style, left: style, bottom: style, right: style };
}

function sameDate(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

module.exports = { buildGanttExcel };
