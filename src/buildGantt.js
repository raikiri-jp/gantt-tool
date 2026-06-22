'use strict';

const ExcelJS = require('exceljs');

// 配色 (ARGB)
const COLOR_HOLIDAY_FILL = 'FFE0E0E0'; // 土日祝の列の薄灰色
const COLOR_HEADER_FILL = 'FF2F5496'; // 濃い青 (ヘッダー)
const COLOR_HEADER_FONT = 'FFFFFFFF';
const COLOR_TODAY_BORDER = 'FFFF0000';
const COLOR_GRID = 'FFD9D9D9';
const COLOR_BAR_NORMAL = 'FF4472C4'; // バーの色(青)

// 日本語ロケールの曜日省略形式 (Excel: 月,火,水,...)
const NUMFMT_WEEKDAY_JP = '[$-411]aaa';

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

/**
 * ローカルの年/月/日を保ったまま、UTC基準のDateオブジェクトに変換する。
 *
 * 背景: `new Date(y, m, d)` はローカルタイムゾーンで0時を表すDateを作るが、
 * 日本(UTC+9)などでこれをそのままExcelの日付セルに書き込むと、内部的には
 * UTCに変換されるため「前日の15:00」のような表示になってしまう
 * (exceljsはDateのUTC値をそのままExcelシリアル値として扱うため)。
 * Excelに渡す直前に必ずこの関数を通すことで、見た目の年/月/日のズレを防ぐ。
 */
function toExcelDate(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}

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
 * 指定日が属する月の1日を返す。
 */
function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

/**
 * 指定日が属する月の末日を返す。
 */
function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

/**
 * 指定日から指定ヶ月後の同日を返す(日付計算用、月末調整はしない単純加算)。
 */
function addMonths(date, n) {
  return new Date(date.getFullYear(), date.getMonth() + n, date.getDate());
}

/**
 * 1-indexed の列番号をExcelの列文字(A, B, ..., Z, AA, ...)に変換する。
 */
function colLetter(colNum) {
  let n = colNum;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * ガントチャート付きExcel(C)を生成して保存する。
 *
 * 列構成: Bから読み込んだ全列(columnOrderの順序通り) + 開始日 + 完了日
 *         (進捗列がBになければ完了日の隣に追加) + タイムライン
 *
 * タイムラインの色塗りは条件付き書式で実現する(開始日・完了日のセルを参照するため、
 * Excel上で開始日・完了日・進捗を編集すればバーの表示も自動的に追従する)。
 *
 * @param {Array} scheduledTasks - scheduler.scheduleTasks() の戻り値
 * @param {object} opts
 *   opts.granularity: 'day' | 'week'
 *   opts.isHoliday: function(date) => boolean
 *   opts.outputPath: string
 *   opts.today: Date
 *   opts.columnOrder: string[]  Bから読み込んだ全列名(開始日・完了日は含まない)
 *   opts.hasProgressColumn: boolean  Bに進捗列が既にあったか
 */
async function buildGanttExcel(scheduledTasks, opts) {
  const { granularity, isHoliday, outputPath, today, columnOrder, hasProgressColumn } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'gantt-tool';
  workbook.created = new Date();

  // ヘッダー行数: 日単位は3行(月/日/曜日)、週単位は2行(月/週開始日)
  const HEADER_ROW_COUNT = granularity === 'day' ? 3 : 2;
  const DATA_START_ROW = HEADER_ROW_COUNT + 1;

  const sheet = workbook.addWorksheet('ガントチャート', {
    views: [{ state: 'frozen', xSplit: columnOrder.length + 2, ySplit: HEADER_ROW_COUNT }],
    pageSetup: {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9, // A4
      margins: { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2 },
    },
  });

  // 固定列見出し = Bの全列(進捗列があれば除く) + 開始日 + 完了日 + 進捗(%)
  // 要件: 「完了日の隣に進捗」なので、進捗列は常に完了日の直後(最後尾)に配置する。
  let orderedHeaders;
  if (hasProgressColumn) {
    const progressName = columnOrder.find((h) => h === '進捗' || h === '進捗(%)');
    const withoutProgress = columnOrder.filter((h) => h !== progressName);
    orderedHeaders = [...withoutProgress, '開始日', '完了日', progressName];
  } else {
    orderedHeaders = [...columnOrder, '開始日', '完了日', '進捗(%)'];
  }
  const FIXED_COL_COUNT = orderedHeaders.length;
  const startDateCol = orderedHeaders.indexOf('開始日') + 1; // 1-indexed
  const endDateCol = orderedHeaders.indexOf('完了日') + 1;

  // タイムラインの範囲を決定
  const minStart = scheduledTasks.reduce(
    (min, t) => (t.startDate < min ? t.startDate : min),
    scheduledTasks[0].startDate
  );
  const maxEnd = scheduledTasks.reduce(
    (max, t) => (t.endDate > max ? t.endDate : max),
    scheduledTasks[0].endDate
  );

  // カレンダーは月初〜月末で揺れなく表示し、終了予定の1ヶ月後まで余裕を持たせる。
  // 週単位の場合は、月初を含む週の月曜日から開始する。
  const calendarStart = startOfMonth(minStart);
  const calendarEndTarget = addMonths(maxEnd, 1);
  const calendarEnd = endOfMonth(calendarEndTarget);

  const timelineStart = granularity === 'week' ? startOfWeekMonday(calendarStart) : calendarStart;
  const timelineEnd = calendarEnd;

  // タイムラインの列見出しリストを作る
  const timelineCols = []; // { date(or weekStart), weekEnd(週単位のみ), isHolidayCol(日単位のみ) }
  if (granularity === 'day') {
    let cur = new Date(timelineStart);
    while (cur <= timelineEnd) {
      timelineCols.push({ date: new Date(cur), isHolidayCol: isHoliday(cur) });
      cur = addDays(cur, 1);
    }
  } else {
    let cur = new Date(timelineStart);
    while (cur <= timelineEnd) {
      const weekEnd = addDays(cur, 6);
      timelineCols.push({ date: new Date(cur), weekEnd: new Date(weekEnd) });
      cur = addDays(cur, 7);
    }
  }

  // --- ヘッダー行 ---
  const headerRow1 = sheet.getRow(1); // 月
  const headerRow2 = sheet.getRow(2); // 日付(day) or 週開始日(week)
  const headerRow3 = granularity === 'day' ? sheet.getRow(3) : null; // 曜日(dayのみ)

  // 固定列見出し(タスク名・担当者など)は縦方向にHEADER_ROW_COUNT行分マージする
  orderedHeaders.forEach((h, i) => {
    const cell = headerRow1.getCell(i + 1);
    cell.value = h;
    sheet.mergeCells(1, i + 1, HEADER_ROW_COUNT, i + 1);
  });

  // 月ラベル用にグルーピング(1行目)
  let monthGroupStartCol = FIXED_COL_COUNT + 1;
  let prevMonthKey = null;
  timelineCols.forEach((tc, idx) => {
    const col = FIXED_COL_COUNT + 1 + idx;
    const monthKey = `${tc.date.getFullYear()}-${tc.date.getMonth()}`;
    if (prevMonthKey === null) {
      prevMonthKey = monthKey;
      monthGroupStartCol = col;
    } else if (monthKey !== prevMonthKey) {
      const startCol = monthGroupStartCol;
      const endCol = col - 1;
      const [y, m] = prevMonthKey.split('-');
      headerRow1.getCell(startCol).value = `${y}年${Number(m) + 1}月`;
      if (endCol > startCol) sheet.mergeCells(1, startCol, 1, endCol);
      prevMonthKey = monthKey;
      monthGroupStartCol = col;
    }

    // 2行目: 日付(日付型のセルにする。日単位は「日」のみ、週単位もこの行は「日」のみ表示し、
    // 月情報は1行目にまとめる)
    const dateCell = headerRow2.getCell(col);
    dateCell.value = toExcelDate(tc.date);
    dateCell.numFmt = 'd';

    // 3行目: 曜日(日単位のみ) - 2行目を参照する数式 + aaa書式
    if (granularity === 'day') {
      const wdCell = headerRow3.getCell(col);
      wdCell.value = { formula: `${colLetter(col)}2` };
      wdCell.numFmt = NUMFMT_WEEKDAY_JP;
    }
  });
  if (timelineCols.length > 0) {
    const startCol = monthGroupStartCol;
    const endCol = FIXED_COL_COUNT + timelineCols.length;
    const [y, m] = prevMonthKey.split('-');
    headerRow1.getCell(startCol).value = `${y}年${Number(m) + 1}月`;
    if (endCol > startCol) sheet.mergeCells(1, startCol, 1, endCol);
  }

  // ヘッダー装飾
  const headerRows = granularity === 'day' ? [headerRow1, headerRow2, headerRow3] : [headerRow1, headerRow2];
  headerRows.forEach((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR_HEADER_FILL } };
      cell.font = { color: { argb: COLOR_HEADER_FONT }, bold: true, size: 10 };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = borderAll(COLOR_GRID);
    });
  });
  headerRow1.height = 18;
  headerRow2.height = granularity === 'day' ? 16 : 18;
  if (headerRow3) headerRow3.height = 16;

  // --- 列幅 ---
  const defaultWidthFor = (headerName) => {
    if (headerName === 'タスク名') return 22;
    if (headerName === '分類' || headerName === '作業Lv1' || headerName === '作業Lv2') return 14;
    if (headerName === '担当者') return 10;
    if (headerName === '工数(日)') return 9;
    if (headerName === '優先順位') return 9;
    if (headerName === '開始日' || headerName === '完了日') return 11;
    if (headerName === '進捗(%)' || headerName === '進捗') return 9;
    return 14;
  };
  orderedHeaders.forEach((h, i) => {
    sheet.getColumn(i + 1).width = defaultWidthFor(h);
  });
  const timelineColWidth = granularity === 'day' ? 4.2 : 6.5;
  for (let i = 0; i < timelineCols.length; i++) {
    sheet.getColumn(FIXED_COL_COUNT + 1 + i).width = timelineColWidth;
  }

  // --- データ行 ---
  const todayNormalized = today ? new Date(today.getFullYear(), today.getMonth(), today.getDate()) : null;

  scheduledTasks.forEach((task, idx) => {
    const rowIdx = DATA_START_ROW + idx;
    const row = sheet.getRow(rowIdx);

    orderedHeaders.forEach((h, ci) => {
      const col = ci + 1;
      const cell = row.getCell(col);
      if (h === '開始日') {
        cell.value = toExcelDate(task.startDate);
        cell.numFmt = 'yyyy/mm/dd';
      } else if (h === '完了日') {
        cell.value = toExcelDate(task.endDate);
        cell.numFmt = 'yyyy/mm/dd';
      } else if (h === '進捗(%)' || h === '進捗') {
        cell.value = task.progress; // 数値(0-100)
        cell.numFmt = '0.0"%"';
      } else if (h === 'タスク名') {
        cell.value = task.name;
      } else if (h === '担当者') {
        cell.value = task.assignee;
      } else if (h === '工数(日)') {
        cell.value = task.days;
      } else if (h === '優先順位') {
        cell.value = task.priority === null || task.priority === undefined ? '' : task.priority;
      } else {
        // Bのその他の列(分類・作業Lv1・作業Lv2など)はそのままコピー
        const v = task.rawValues ? task.rawValues[h] : undefined;
        cell.value = v === undefined || v === null ? '' : v;
      }

      cell.border = borderAll(COLOR_GRID);
      cell.alignment = { vertical: 'middle', horizontal: h === 'タスク名' ? 'left' : 'center' };
      cell.font = { size: 10 };
    });

    // タイムライン部分: 値は入れない。
    // 休日のグレー塗り・今日の右側縦線・バー(開始日〜完了日)の色は
    // すべて後段で条件付き書式として一括設定する
    // (休日を優先させるため、休日ルールをバーのルールより先に評価させる)。
    timelineCols.forEach((tc, ci) => {
      const col = FIXED_COL_COUNT + 1 + ci;
      const cell = row.getCell(col);
      cell.border = borderAll(COLOR_GRID);
    });

    row.height = 16;
  });

  const dataEndRow = DATA_START_ROW + scheduledTasks.length - 1;

  if (timelineCols.length > 0 && scheduledTasks.length > 0) {
    const firstTimelineCol = FIXED_COL_COUNT + 1;
    const lastTimelineCol = FIXED_COL_COUNT + timelineCols.length;
    const rangeRef =
      `${colLetter(firstTimelineCol)}${DATA_START_ROW}:${colLetter(lastTimelineCol)}${dataEndRow}`;

    const startColLetter = colLetter(startDateCol);
    const endColLetter = colLetter(endDateCol);
    const firstColLetterTimeline = colLetter(firstTimelineCol);

    // --- 条件付き書式 1: 休日(土日・祝日・年末年始)のグレー塗り ---
    // 優先度を1(最優先)にし、stopIfTrueでバーのルールより先に評価・確定させる。
    // これにより「休日に作業しているように見える」ことを防ぐ(休日表示が常に優先される)。
    // 日単位のみ適用(週単位は1セルが1週間を表すため、休日の概念をそのまま塗りには使わない)。
    if (granularity === 'day') {
      timelineCols.forEach((tc, ci) => {
        if (!tc.isHolidayCol) return;
        const col = FIXED_COL_COUNT + 1 + ci;
        const colLet = colLetter(col);
        const holidayRangeRef = `${colLet}${DATA_START_ROW}:${colLet}${dataEndRow}`;
        sheet.addConditionalFormatting({
          ref: holidayRangeRef,
          rules: [
            {
              type: 'expression',
              priority: 1,
              stopIfTrue: true,
              formulae: ['TRUE'],
              style: {
                fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLOR_HOLIDAY_FILL } },
              },
            },
          ],
        });
      });
    }

    // --- 条件付き書式 2: タイムラインのバー(開始日〜完了日)を塗る ---
    // 数式は「タイムラインのこの列の日付」が「この行の開始日以上、完了日以下」なら塗る。
    // 日単位: タイムラインの日付セルは headerRow2 (2行目)。
    // 週単位: タイムラインの週開始日セルは headerRow2 (2行目)。週の場合は
    //         「週の範囲(開始日〜開始日+6)」とタスク期間が重なっていれば塗る。
    let formula;
    if (granularity === 'day') {
      // ${firstColLetterTimeline}$2 はタイムライン2行目(日付)。列は相対、行は絶対。
      formula = `AND(${firstColLetterTimeline}$2>=$${startColLetter}${DATA_START_ROW},${firstColLetterTimeline}$2<=$${endColLetter}${DATA_START_ROW})`;
    } else {
      // 週単位: その週の開始日(2行目) <= 完了日 AND その週の開始日+6(週末) >= 開始日
      formula =
        `AND(${firstColLetterTimeline}$2<=$${endColLetter}${DATA_START_ROW},(${firstColLetterTimeline}$2+6)>=$${startColLetter}${DATA_START_ROW})`;
    }

    // exceljsの条件付き書式は「ref」全体に対して、行頭セル基準の相対参照で評価される。
    // ここでは1行目(DATA_START_ROW)を基準にした数式を使い、ref全体に適用する
    // (Excelの仕様上、相対参照は範囲内で自動的にスライドする)。
    // priorityを2にして休日ルールより後に評価させる(休日ルールがstopIfTrueのため、
    // 休日列ではこのルールまで到達しない)。
    sheet.addConditionalFormatting({
      ref: rangeRef,
      rules: [
        {
          type: 'expression',
          priority: 2,
          formulae: [formula],
          style: {
            fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLOR_BAR_NORMAL } },
          },
        },
      ],
    });

    // --- 条件付き書式 3: 今日の列の右側に縦線を引く ---
    // セル全体を枠で囲むのではなく、右側の縦線のみを引く(列の境界が「今日」であることを示す)。
    // この列が「今日」かどうかは、タイムラインヘッダーの日付セル(2行目)を基準に判定する。
    // 日単位: その列の日付が今日と一致する場合。
    // 週単位: 今日がその週の範囲(開始日〜開始日+6)に含まれる場合。
    if (todayNormalized) {
      const todayExcelSerial = toExcelDate(todayNormalized);
      let todayFormula;
      if (granularity === 'day') {
        todayFormula = `${firstColLetterTimeline}$2=DATE(${todayExcelSerial.getUTCFullYear()},${todayExcelSerial.getUTCMonth() + 1},${todayExcelSerial.getUTCDate()})`;
      } else {
        todayFormula =
          `AND(${firstColLetterTimeline}$2<=DATE(${todayExcelSerial.getUTCFullYear()},${todayExcelSerial.getUTCMonth() + 1},${todayExcelSerial.getUTCDate()}),(${firstColLetterTimeline}$2+6)>=DATE(${todayExcelSerial.getUTCFullYear()},${todayExcelSerial.getUTCMonth() + 1},${todayExcelSerial.getUTCDate()}))`;
      }

      sheet.addConditionalFormatting({
        ref: rangeRef,
        rules: [
          {
            type: 'expression',
            priority: 3,
            formulae: [todayFormula],
            style: {
              border: { right: { style: 'medium', color: { argb: COLOR_TODAY_BORDER } } },
            },
          },
        ],
      });
    }
  }

  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

function borderAll(argb) {
  const style = { style: 'thin', color: { argb } };
  return { top: style, left: style, bottom: style, right: style };
}

module.exports = { buildGanttExcel };
