'use strict';

const path = require('path');
const fs = require('fs');
const readline = require('readline');

const { loadHolidays, makeIsHoliday, buildYearEndHolidayList } = require('./src/holidays');
const { readBFile } = require('./src/readTasks');
const { scheduleTasks } = require('./src/scheduler');
const { buildGanttExcel } = require('./src/buildGantt');

/**
 * 標準入出力で1行の質問をして回答を受け取る。
 *
 * readline-syncのような同期入力ライブラリは、Windows環境で内部的に
 * 別プロセス/別I/O経路を介してプロンプト文字列を出力することがあり、
 * ターミナルのコードページ設定と噛み合わずに日本語が文字化けすることがある。
 * Node標準のreadlineモジュールはconsole.logと同じstdout経路を使うため、
 * 文字化けの問題を避けられる。
 */
function ask(rl, promptText) {
  return new Promise((resolve) => {
    rl.question(promptText, (answer) => resolve(answer));
  });
}

function todayDate() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function timestampForFilename() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return (
    `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_` +
    `${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`
  );
}

function parseArgs(argv) {
  // 使い方:
  //   node make-gantt.js <Bファイル.xlsx> [--granularity day|week]
  const positional = [];
  const opts = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--granularity' || a === '-g') {
      opts.granularity = argv[++i];
    } else if (a === '--help' || a === '-h') {
      opts.help = true;
    } else {
      positional.push(a);
    }
  }
  return { positional, opts };
}

async function main() {
  const { positional, opts } = parseArgs(process.argv.slice(2));

  if (opts.help || positional.length < 1) {
    console.log('使い方: node make-gantt.js <Bファイルのパス(.xlsx)> [--granularity day|week]');
    console.log('  --granularity, -g   day(日単位) または week(週単位) (省略時は対話入力。さらに省略するとweek)');
    console.log('');
    console.log('注: プロジェクト全体の開始日は指定しません。各タスクの開始日は');
    console.log('    Bファイル内の「担当者」シートに記載された各担当者の参画開始日から計算されます。');
    process.exit(positional.length < 1 ? 1 : 0);
  }
  const inputPath = positional[0];
  if (!fs.existsSync(inputPath)) {
    console.error(`ファイルが見つかりません: ${inputPath}`);
    process.exit(1);
  }

  console.log('=== ガントチャート自動生成ツール ===');
  console.log(`入力ファイル(B): ${inputPath}`);

  const isInteractive = process.stdin.isTTY === true;
  const rl = isInteractive
    ? readline.createInterface({ input: process.stdin, output: process.stdout })
    : null;

  // --- 粒度の決定 ---
  let granularity = null;
  if (opts.granularity !== undefined) {
    const v = opts.granularity.trim().toLowerCase();
    if (v === 'day' || v === 'd') granularity = 'day';
    else if (v === 'week' || v === 'w') granularity = 'week';
    else {
      console.error(`--granularity の値が正しくありません: ${opts.granularity} ("day" または "week" を指定してください)`);
      process.exit(1);
    }
  } else if (isInteractive) {
    while (!granularity) {
      const input = await ask(
        rl,
        'ガントチャートの粒度を選んでください [day=日単位 / week=週単位] (デフォルト: week): '
      );
      const v = input.trim().toLowerCase();
      if (v === '' || v === 'week' || v === 'w') {
        granularity = 'week';
      } else if (v === 'day' || v === 'd') {
        granularity = 'day';
      } else {
        console.log('"day" または "week" を入力してください。');
      }
    }
  } else {
    granularity = 'week';
    console.log('(非対話実行のため粒度は week を使用します。--granularity で指定できます)');
  }
  console.log(`粒度: ${granularity === 'day' ? '日単位' : '週単位'}`);

  if (rl) rl.close();

  // --- 祝日データのロード(内閣府祝日+年末年始。会社独自の休日はBファイルの「休日」シートから) ---
  let holidayInfo;
  try {
    holidayInfo = await loadHolidays();
  } catch (err) {
    console.error('\nエラー: ' + err.message);
    process.exit(1);
  }
  if (holidayInfo.usedCache) {
    console.log(`祝日キャッシュを使用しました: ${holidayInfo.cachePath}`);
  }
  const isHoliday = makeIsHoliday(holidayInfo.dateSet);

  // --- Bファイルの読み込み(タスク一覧・担当者シート・休日シート) ---
  let bData;
  try {
    bData = await readBFile(inputPath);
  } catch (err) {
    console.error('\nエラー: ' + err.message);
    process.exit(1);
  }
  const { tasks, columnOrder, hasProgressColumn, assigneeStartDates, holidaySheetRows } = bData;
  console.log(`タスクを${tasks.length}件読み込みました。`);
  console.log(`担当者を${assigneeStartDates.size}名読み込みました。`);
  if (!hasProgressColumn) {
    console.log('(「進捗」列が見つからなかったため、全タスク進捗0%として出力します)');
  }

  // --- スケジューリング ---
  let scheduled;
  try {
    scheduled = scheduleTasks(tasks, assigneeStartDates, isHoliday);
  } catch (err) {
    console.error('\nエラー: ' + err.message);
    process.exit(1);
  }

  // --- 休日シートに書き出す一覧を組み立てる ---
  // 内閣府の祝日キャッシュ(名称付き) + 年末年始休暇(タイムライン表示期間をカバーする範囲)
  // + Bファイルに既に「休日」シートがあった場合はその内容(会社独自の休日など)を引き継ぐ。
  const minStart = scheduled.reduce((min, t) => (t.startDate < min ? t.startDate : min), scheduled[0].startDate);
  const maxEnd = scheduled.reduce((max, t) => (t.endDate > max ? t.endDate : max), scheduled[0].endDate);
  // 年末年始は表示範囲より広めに(前後1年)生成しておけば、カレンダー拡張後の範囲もカバーできる。
  const yearEndRangeStart = new Date(minStart.getFullYear() - 1, 0, 1);
  const yearEndRangeEnd = new Date(maxEnd.getFullYear() + 1, 11, 31);
  const yearEndList = buildYearEndHolidayList(yearEndRangeStart, yearEndRangeEnd);

  const holidayRows = [
    ...holidayInfo.holidayList.map((h) => ({ date: parseYmd(h.date), name: h.name })),
    ...yearEndList.map((h) => ({ date: parseYmd(h.date), name: h.name })),
    ...(holidaySheetRows || []),
  ];
  // 同日重複を除去する(祝日キャッシュと年末年始が重なるケースなどがあるため)。
  // 会社独自の休日(Bの休日シートに元からあったもの)は名称が異なる可能性があるため、
  // 「同じ日付・同じ名称」の完全重複のみ除去する。
  const seenKeys = new Set();
  const dedupedHolidayRows = [];
  for (const h of holidayRows) {
    const key = `${h.date.getFullYear()}-${h.date.getMonth()}-${h.date.getDate()}|${h.name}`;
    if (seenKeys.has(key)) continue;
    seenKeys.add(key);
    dedupedHolidayRows.push(h);
  }

  // --- 出力 ---
  const inputDir = path.dirname(path.resolve(inputPath));
  let outputPath = path.join(inputDir, `gantt_${timestampForFilename()}.xlsx`);
  let suffix = 1;
  while (fs.existsSync(outputPath)) {
    suffix++;
    outputPath = path.join(inputDir, `gantt_${timestampForFilename()}_${suffix}.xlsx`);
  }

  await buildGanttExcel(scheduled, {
    granularity,
    outputPath,
    columnOrder,
    hasProgressColumn,
    holidayRows: dedupedHolidayRows,
    assigneeStartDates,
  });

  console.log('\n=== 完了 ===');
  console.log(`ガントチャートを作成しました: ${outputPath}`);
  console.log('\n--- スケジュール概要 ---');
  for (const t of scheduled) {
    const s = `${t.startDate.getFullYear()}/${t.startDate.getMonth() + 1}/${t.startDate.getDate()}`;
    const e = `${t.endDate.getFullYear()}/${t.endDate.getMonth() + 1}/${t.endDate.getDate()}`;
    console.log(`  [${t.assignee}] ${t.name} (${t.days}日, 進捗${t.progress}%): ${s} 〜 ${e}`);
  }
}

function parseYmd(ymd) {
  const [y, m, d] = ymd.split('-').map(Number);
  return new Date(y, m - 1, d);
}

main().catch((err) => {
  console.error('予期しないエラーが発生しました:', err);
  process.exit(1);
});
