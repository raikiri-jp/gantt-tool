'use strict';

/**
 * 日付を1営業日進める (休日ならスキップする)
 */
function addBusinessDays(startDate, numDays, isHoliday) {
  // startDate から数えて、休日を除いた numDays 日分を消費した最終日を返す。
  // 例: numDays=1 で startDate が営業日なら、終了日は startDate そのもの。
  let current = new Date(startDate);
  // startDate 自体が休日の場合は、最初の営業日まで進める
  while (isHoliday(current)) {
    current = nextDay(current);
  }
  let remaining = numDays - 1; // 開始日で1日分消費済み
  while (remaining > 0) {
    current = nextDay(current);
    if (!isHoliday(current)) {
      remaining--;
    }
  }
  return current;
}

function nextDay(date) {
  const d = new Date(date);
  d.setDate(d.getDate() + 1);
  return d;
}

/**
 * 指定日以降で最初の営業日を返す(指定日が営業日ならそのまま返す)
 */
function firstBusinessDayOnOrAfter(date, isHoliday) {
  let current = new Date(date);
  while (isHoliday(current)) {
    current = nextDay(current);
  }
  return current;
}

/**
 * タスク一覧をスケジューリングする。
 *
 * @param {Array} tasks - [{ name, assignee, days, priority, ...other }]
 * @param {Date} projectStart - プロジェクト全体の開始日
 * @param {function} isHoliday - 日付を受け取り休日かどうかを返す関数
 * @returns {Array} スケジュール結果 [{ ...task, startDate, endDate }]
 *
 * ルール:
 * - 依存関係はなし。
 * - 優先順位(数値が小さいほど優先、同値は入力順)でソートして順に配置する。
 * - 各タスクは「projectStart」と「その担当者が空く最初の営業日」の
 *   うち遅い方から開始する。
 * - 同一担当者のタスクは重複しない(直列)。別担当者は並行可能。
 */
function scheduleTasks(tasks, projectStart, isHoliday) {
  // 優先順位でソート (小さい数字が先。未指定/NaNは最後、入力順を保持する安定ソート)
  const indexed = tasks.map((t, i) => ({ t, i }));
  indexed.sort((a, b) => {
    const pa = normalizedPriority(a.t.priority);
    const pb = normalizedPriority(b.t.priority);
    if (pa !== pb) return pa - pb;
    return a.i - b.i; // 同順位は入力順
  });

  const projectStartBusiness = firstBusinessDayOnOrAfter(projectStart, isHoliday);

  // 担当者ごとの「次に空く日」を管理
  const assigneeNextAvailable = new Map();

  const results = [];
  for (const { t } of indexed) {
    const assignee = (t.assignee || '(未割当)').toString();
    const days = Number(t.days);
    if (!Number.isFinite(days) || days <= 0) {
      throw new Error(`タスク「${t.name}」の工数(日)が不正です: ${t.days}`);
    }

    const assigneeFree = assigneeNextAvailable.has(assignee)
      ? assigneeNextAvailable.get(assignee)
      : projectStartBusiness;

    // プロジェクト開始日と担当者の空き日、どちらか遅い方
    let candidateStart = assigneeFree > projectStartBusiness ? assigneeFree : projectStartBusiness;
    candidateStart = firstBusinessDayOnOrAfter(candidateStart, isHoliday);

    const endDate = addBusinessDays(candidateStart, days, isHoliday);

    results.push({
      ...t,
      assignee,
      days,
      startDate: candidateStart,
      endDate,
    });

    // 担当者の次の空き日は終了日の翌営業日
    const nextFree = firstBusinessDayOnOrAfter(nextDay(endDate), isHoliday);
    assigneeNextAvailable.set(assignee, nextFree);
  }

  // 表示用に元の入力順(または優先順位順)に並べ替えたい場合は呼び出し側で調整可能。
  // ここでは「開始日順」に並べてガントチャートで見やすくする。
  results.sort((a, b) => {
    if (a.startDate.getTime() !== b.startDate.getTime()) {
      return a.startDate.getTime() - b.startDate.getTime();
    }
    return normalizedPriority(a.priority) - normalizedPriority(b.priority);
  });

  return results;
}

function normalizedPriority(p) {
  if (p === null || p === undefined || p === '') return Number.MAX_SAFE_INTEGER;
  const n = Number(p);
  if (!Number.isFinite(n)) return Number.MAX_SAFE_INTEGER;
  return n;
}

module.exports = {
  scheduleTasks,
  addBusinessDays,
  firstBusinessDayOnOrAfter,
  nextDay,
};
