'use strict';

const ExcelJS = require('exceljs');

// 列名のゆらぎを吸収するためのエイリアス定義 (計算に必要な列の特定に使う)
const HEADER_ALIASES = {
  name: ['タスク名', 'タスク', '作業名', '作業', 'name', 'task'],
  assignee: ['担当者', '担当', 'assignee', '担当者名'],
  days: ['工数(日)', '工数（日）', '工数', '日数', 'days', 'workdays', '工数(日数)'],
  priority: ['優先順位', '優先度', 'priority'],
  progress: ['進捗', '進捗(%)', '進捗（％）', '進捗率', 'progress'],
  // 前回生成したC(=今回のB)を再利用する場合、これらの列は再計算するので
  // 読み込み元の列としては無視し、出力時に作り直す。
  startDate: ['開始日', 'start', 'start date'],
  endDate: ['完了日', '終了日', 'end', 'end date', 'due'],
};

function normalizeHeader(h) {
  return (h || '').toString().trim();
}

function findColumnIndex(headerRow, aliases) {
  for (let c = 1; c <= headerRow.cellCount; c++) {
    const val = normalizeHeader(headerRow.getCell(c).value);
    if (aliases.includes(val)) return c;
  }
  return -1;
}

function cellPlainValue(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'object' && 'result' in v) return v.result; // 数式セル
  if (typeof v === 'object' && v instanceof Date) return v;
  return v;
}

/**
 * 行のすべてのセルが空かどうかを判定する。
 */
function isRowBlank(row) {
  if (row.cellCount === 0) return true;
  for (let c = 1; c <= row.cellCount; c++) {
    const v = row.getCell(c).value;
    if (v !== null && v !== undefined && v !== '') return false;
  }
  return true;
}

/**
 * 2つの行が「同じヘッダーの一部(マージセル由来で同じ文字列が入っている)」に
 * 見えるかどうかを判定する。
 * 固定列部分(分類・タスク名・…・進捗など)は完全一致するが、タイムライン部分は
 * 行ごとに異なる値(月・日・曜日)になるため、「先頭から連続して一致している列数」
 * を数え、ある程度(4列以上)連続一致していれば真とみなす。
 */
function looksLikeTwoRowHeader(rowA, rowB) {
  const maxCount = Math.max(rowA.cellCount, rowB.cellCount);
  let consecutiveMatch = 0;
  for (let c = 1; c <= maxCount; c++) {
    const a = normalizeHeader(rowA.getCell(c).value);
    const b = normalizeHeader(rowB.getCell(c).value);
    if (a && a === b) {
      consecutiveMatch++;
    } else {
      break;
    }
  }
  return consecutiveMatch >= 4;
}

/**
 * 1行目を基準に、2行目・3行目が「複数行ヘッダーの一部」かどうかを判定し、
 * データが実際に始まる行番号を返す。
 *
 * Cをそのまま次回のBとして使う場合、固定列部分(分類・タスク名・…・進捗など)は
 * 縦方向にマージされているため、1行目と同じ文字列が2行目・3行目にも入っている
 * (マージセルの結合範囲内であれば、exceljsはどのセルを読んでも同じ値を返す)。
 * 一方、タイムライン部分は行ごとに異なる値(月・日・曜日)になるため、
 * 「先頭から連続して一致している列数」で複数行ヘッダーかどうかを判定する。
 *
 * @returns {number} データが開始する行番号(1-indexed)
 */
function detectDataStartRow(sheet) {
  const row1 = sheet.getRow(1);
  const row2 = sheet.getRow(2);
  const row3 = sheet.getRow(3);

  if (looksLikeTwoRowHeader(row1, row2)) {
    // 2行目もヘッダーの一部。さらに3行目もヘッダーの一部(日単位Cの曜日行)かを判定する。
    if (looksLikeTwoRowHeader(row1, row3)) {
      return 4;
    }
    return 3;
  }
  return 2;
}

/**
 * B(入力Excel)を読み込む。
 * 想定: 1枚目のシート、1行目がヘッダー。
 *
 * 設計方針:
 * - Bの列構成は自由(タスク名, 担当者, 工数(日), 優先順位は必須。それ以外は何でも追加可能)。
 * - 「開始日」「完了日」列がもし存在していたら(=前回生成したCを今回のBとして使っている場合)、
 *   読み込み時には無視する(再計算するため)。
 * - 計算に使わない列も全て保持し、出力時に同じ順序・同じ見出しでコピーする。
 *
 * @param {string} filePath
 * @returns {Promise<{ tasks: Array, columnOrder: Array<string>, hasProgressColumn: boolean }>}
 *   tasks: [{ name, assignee, days, priority, progress, rawValues: {colName: value} }]
 *   columnOrder: 出力時にそのままコピーする列名の並び (開始日・完了日は除く)
 */
async function readTasksFromExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('Bファイルにシートが見つかりません。');
  }

  // ヘッダー行の検出: 1行目をヘッダーとするが、Cを再利用する場合などマージセルにより
  // 2行目・3行目にも同じヘッダー文字列が入っていることがある(日単位Cは3行ヘッダー)。
  // その場合はそれらもヘッダーとみなし、データはその直後から始まるものとして扱う。
  const headerRow = sheet.getRow(1);
  const dataStartRow = detectDataStartRow(sheet);

  const colName = findColumnIndex(headerRow, HEADER_ALIASES.name);
  const colAssignee = findColumnIndex(headerRow, HEADER_ALIASES.assignee);
  const colDays = findColumnIndex(headerRow, HEADER_ALIASES.days);
  const colPriority = findColumnIndex(headerRow, HEADER_ALIASES.priority);
  const colProgress = findColumnIndex(headerRow, HEADER_ALIASES.progress);
  const colStartDate = findColumnIndex(headerRow, HEADER_ALIASES.startDate);
  const colEndDate = findColumnIndex(headerRow, HEADER_ALIASES.endDate);

  const missing = [];
  if (colName === -1) missing.push('タスク名');
  if (colAssignee === -1) missing.push('担当者');
  if (colDays === -1) missing.push('工数(日)');
  if (colPriority === -1) missing.push('優先順位');
  if (missing.length > 0) {
    throw new Error(
      `Bファイルに必須列が見つかりません: ${missing.join(', ')}\n` +
      `1行目をヘッダー行とし、列名に「タスク名」「担当者」「工数(日)」「優先順位」を含めてください。`
    );
  }

  // 出力時にそのままコピーする列の並び (開始日・完了日列は除外して再生成する)
  //
  // 注意: headerRow.cellCount は、Cをそのまま次回のBとして使った場合、
  // タイムライン部分の月見出し(マージセル)まで含んでしまうことがある。
  // そこで「データ行に実際に値が入りうる列数」でヘッダー列を制限する。
  // (タイムラインのバー部分はセルの値を持たない色塗りのみのため、これで安全に切り分けられる)
  // 1行だけでなく数行分チェックして、最大の列数を採用する(取りこぼし防止)。
  let effectiveColCount = headerRow.cellCount;
  let maxDataColWithValue = 0;
  const sampleRowCount = Math.min(10, sheet.rowCount - dataStartRow + 1);
  for (let i = 0; i < sampleRowCount; i++) {
    const r = sheet.getRow(dataStartRow + i);
    if (!r || r.cellCount === 0) continue;
    for (let c = 1; c <= r.cellCount; c++) {
      const v = r.getCell(c).value;
      if (v !== null && v !== undefined && v !== '' && c > maxDataColWithValue) {
        maxDataColWithValue = c;
      }
    }
  }
  if (maxDataColWithValue > 0) {
    effectiveColCount = Math.min(effectiveColCount, maxDataColWithValue);
  }

  const excludedCols = new Set([colStartDate, colEndDate].filter((c) => c !== -1));
  const columnOrder = [];
  for (let c = 1; c <= effectiveColCount; c++) {
    if (excludedCols.has(c)) continue;
    const h = normalizeHeader(headerRow.getCell(c).value);
    if (h) columnOrder.push({ col: c, name: h });
  }

  const tasks = [];
  let stopReading = false;
  sheet.eachRow((row, rowNumber) => {
    if (stopReading) return;
    if (rowNumber < dataStartRow) return; // ヘッダー行はスキップ

    // 行全体が空(=空行)なら、そこでデータは終わったとみなす。
    // (Cを次回のBとして使う場合、データ行の後に凡例などの説明行が続くため、
    //  そこまで誤って読み込まないようにする)
    if (isRowBlank(row)) {
      stopReading = true;
      return;
    }

    const nameVal = cellPlainValue(row.getCell(colName));
    const name = nameVal === null ? '' : nameVal.toString().trim();
    if (!name) return; // タスク名だけ空の行はスキップ(他の列にメモ等がある場合を想定)

    const assigneeVal = cellPlainValue(row.getCell(colAssignee));
    const assignee = assigneeVal === null ? '' : assigneeVal.toString().trim();

    const days = cellPlainValue(row.getCell(colDays));
    const priority = cellPlainValue(row.getCell(colPriority));
    const progress = colProgress !== -1 ? cellPlainValue(row.getCell(colProgress)) : null;

    // 出力時にそのままコピーするための全列の値(開始日・完了日は除く)
    const rawValues = {};
    for (const oc of columnOrder) {
      rawValues[oc.name] = cellPlainValue(row.getCell(oc.col));
    }

    tasks.push({ name, assignee, days, priority, progress, rawValues });
  });

  if (tasks.length === 0) {
    throw new Error('Bファイルに有効なタスク行が見つかりませんでした。');
  }

  return {
    tasks,
    columnOrder: columnOrder.map((c) => c.name),
    hasProgressColumn: colProgress !== -1,
  };
}

module.exports = { readTasksFromExcel, HEADER_ALIASES };
