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
 * @param {Array} tasks - [{ name, assignee, days, priority, progress, ...other }]
 * @param {Map<string, Date>} assigneeStartDates - 担当者名 -> 参画開始日
 * @param {function} isHoliday - 日付を受け取り休日かどうかを返す関数
 * @returns {Array} スケジュール結果 [{ ...task, startDate, endDate, progress }]
 *   (Bファイルの行順をそのまま保持して返す)
 *
 * ルール:
 * - 依存関係はなし。
 * - 日程計算は優先順位(数値が小さいほど優先、同値は入力順)の順に行う。
 * - 各タスクは「その担当者の参画開始日」と「その担当者が空く最初の営業日」の
 *   うち遅い方から開始する(参画開始日より前には着手しない)。
 * - 同一担当者のタスクは重複しない(直列)。別担当者は並行可能。
 * - 進捗は常に保持する(未入力は0%扱い)。
 * - 戻り値の並び順は、計算順(優先順位順)ではなく、常にBファイルの行順(入力順)。
 */
function scheduleTasks(tasks, assigneeStartDates, isHoliday) {
  // 優先順位でソート (小さい数字が先。未指定/NaNは最後、入力順を保持する安定ソート)
  const indexed = tasks.map((t, i) => ({ t, i }));
  indexed.sort((a, b) => {
    const pa = normalizedPriority(a.t.priority);
    const pb = normalizedPriority(b.t.priority);
    if (pa !== pb) return pa - pb;
    return a.i - b.i; // 同順位は入力順
  });

  // 担当者ごとの「次に空く日」を管理。初期値は各担当者の参画開始日。
  const assigneeNextAvailable = new Map();

  const results = [];
  for (const { t, i } of indexed) {
    const assignee = (t.assignee || '').toString().trim();
    const days = Number(t.days);
    if (!Number.isFinite(days) || days <= 0) {
      throw new Error(`タスク「${t.name}」の工数(日)が不正です: ${t.days}`);
    }

    const assigneeStart = assigneeStartDates.get(assignee);
    if (!assigneeStart) {
      // readBFile側で事前チェックしているはずだが、直接scheduleTasksを呼ぶ
      // ケースのために二重に防御する。
      throw new Error(`担当者「${assignee}」の参画開始日が「担当者」シートに見つかりません。`);
    }
    const assigneeStartBusiness = firstBusinessDayOnOrAfter(assigneeStart, isHoliday);

    const assigneeFree = assigneeNextAvailable.has(assignee)
      ? assigneeNextAvailable.get(assignee)
      : assigneeStartBusiness;

    // 担当者の参画開始日と、その担当者の直近の空き日、どちらか遅い方
    let candidateStart = assigneeFree > assigneeStartBusiness ? assigneeFree : assigneeStartBusiness;
    candidateStart = firstBusinessDayOnOrAfter(candidateStart, isHoliday);

    const endDate = addBusinessDays(candidateStart, days, isHoliday);

    const progress = normalizedProgress(t.progress);

    results.push({
      ...t,
      assignee,
      days,
      progress,
      startDate: candidateStart,
      endDate,
      __inputIndex: i,
    });

    // 担当者の次の空き日は終了日の翌営業日
    const nextFree = firstBusinessDayOnOrAfter(nextDay(endDate), isHoliday);
    assigneeNextAvailable.set(assignee, nextFree);
  }

  // 表示順は常にBファイルの行順(入力順)をそのまま保持する。
  // スケジューリングの計算(優先順位・担当者の空き状況)は内部処理のみに使い、
  // 出力時の並び順には影響させない。
  results.sort((a, b) => a.__inputIndex - b.__inputIndex);
  results.forEach((r) => delete r.__inputIndex);

  return results;
}

function normalizedPriority(p) {
  if (p === null || p === undefined || p === '') return Number.MAX_SAFE_INTEGER;
  const n = Number(p);
  if (!Number.isFinite(n)) return Number.MAX_SAFE_INTEGER;
  return n;
}

/**
 * 進捗(%)を0-100の数値に正規化する。
 * 未入力・不正値は0%とみなす。0.5のような小数(=50%の意図)は100倍して解釈する。
 */
function normalizedProgress(p) {
  if (p === null || p === undefined || p === '') return 0;
  let n = Number(p);
  if (!Number.isFinite(n)) return 0;
  if (n > 0 && n <= 1) n = n * 100; // 0.5 -> 50% とみなす(Excelのパーセント書式セル対策)
  if (n < 0) n = 0;
  if (n > 100) n = 100;
  return Math.round(n * 10) / 10; // 小数第1位までに丸める
}

module.exports = {
  scheduleTasks,
  addBusinessDays,
  firstBusinessDayOnOrAfter,
  nextDay,
};
