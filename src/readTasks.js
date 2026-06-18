'use strict';

const ExcelJS = require('exceljs');

// 列名のゆらぎを吸収するためのエイリアス定義
const HEADER_ALIASES = {
  name: ['タスク名', 'タスク', '作業名', '作業', 'name', 'task'],
  assignee: ['担当者', '担当', 'assignee', '担当者名'],
  days: ['工数(日)', '工数（日）', '工数', '日数', 'days', 'workdays', '工数(日数)'],
  priority: ['優先順位', '優先度', 'priority'],
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

/**
 * B(入力Excel)を読み込み、タスク配列を返す。
 * 想定: 1枚目のシート、1行目がヘッダー。
 * 列名は HEADER_ALIASES に従って柔軟に検出する。
 * タスク名, 担当者, 工数(日), 優先順位 の4列が必須。
 * それ以外の列は other に保持し、無視されるが落とさない。
 *
 * @param {string} filePath
 * @returns {Promise<Array>} [{ name, assignee, days, priority, other: {colName: value} }]
 */
async function readTasksFromExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('Bファイルにシートが見つかりません。');
  }

  const headerRow = sheet.getRow(1);
  const colName = findColumnIndex(headerRow, HEADER_ALIASES.name);
  const colAssignee = findColumnIndex(headerRow, HEADER_ALIASES.assignee);
  const colDays = findColumnIndex(headerRow, HEADER_ALIASES.days);
  const colPriority = findColumnIndex(headerRow, HEADER_ALIASES.priority);

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

  // その他の列名一覧 (既知列を除く)
  const knownCols = new Set([colName, colAssignee, colDays, colPriority]);
  const otherCols = [];
  for (let c = 1; c <= headerRow.cellCount; c++) {
    if (!knownCols.has(c)) {
      const h = normalizeHeader(headerRow.getCell(c).value);
      if (h) otherCols.push({ col: c, name: h });
    }
  }

  const tasks = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // ヘッダー行はスキップ
    const nameVal = row.getCell(colName).value;
    const name = nameVal === null || nameVal === undefined ? '' : nameVal.toString().trim();
    if (!name) return; // タスク名が空の行はスキップ

    const assigneeVal = row.getCell(colAssignee).value;
    const assignee = assigneeVal === null || assigneeVal === undefined ? '' : assigneeVal.toString().trim();

    const daysVal = row.getCell(colDays).value;
    const days = typeof daysVal === 'object' && daysVal !== null && 'result' in daysVal ? daysVal.result : daysVal;

    const priorityVal = row.getCell(colPriority).value;
    const priority = typeof priorityVal === 'object' && priorityVal !== null && 'result' in priorityVal ? priorityVal.result : priorityVal;

    const other = {};
    for (const oc of otherCols) {
      other[oc.name] = row.getCell(oc.col).value;
    }

    tasks.push({ name, assignee, days, priority, other });
  });

  if (tasks.length === 0) {
    throw new Error('Bファイルに有効なタスク行が見つかりませんでした。');
  }

  return tasks;
}

module.exports = { readTasksFromExcel, HEADER_ALIASES };
