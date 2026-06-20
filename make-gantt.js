'use strict';

const path = require('path');
const fs = require('fs');
const readlineSync = require('readline-sync');

const { loadHolidays, makeIsHoliday } = require('./src/holidays');
const { readTasksFromExcel } = require('./src/readTasks');
const { scheduleTasks } = require('./src/scheduler');
const { buildGanttExcel } = require('./src/buildGantt');

function parseDateInput(str) {
  const s = str.trim();
  // YYYY-MM-DD or YYYY/MM/DD or YYYYMMDD
  let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
  return date;
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
  //   node make-gantt.js <Bファイル.xlsx> [--start YYYY-MM-DD] [--granularity day|week]
  const positional = [];
  const opts = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--start' || a === '-s') {
      opts.start = argv[++i];
    } else if (a === '--granularity' || a === '-g') {
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
    console.log('使い方: node make-gantt.js <Bファイルのパス(.xlsx)> [--start YYYY-MM-DD] [--granularity day|week]');
    console.log('  --start, -s         プロジェクト開始日 (省略時は対話入力。さらに省略すると今日)');
    console.log('  --granularity, -g   day(日単位) または week(週単位) (省略時は対話入力。さらに省略するとweek)');
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

  // --- 開始日の決定 ---
  let projectStart = null;
  if (opts.start !== undefined) {
    projectStart = parseDateInput(opts.start);
    if (!projectStart) {
      console.error(`--start の日付形式が正しくありません: ${opts.start} (例: 2026-06-20)`);
      process.exit(1);
    }
  } else if (isInteractive) {
    while (!projectStart) {
      const input = readlineSync.question(
        'プロジェクト全体の開始日を入力してください (例: 2026-06-20、何も入力しないと今日): '
      );
      if (input.trim() === '') {
        projectStart = todayDate();
        break;
      }
      projectStart = parseDateInput(input);
      if (!projectStart) {
        console.log('日付の形式が正しくありません。YYYY-MM-DD の形式で入力してください。');
      }
    }
  } else {
    projectStart = todayDate();
    console.log('(非対話実行のため開始日は今日を使用します。--start で指定できます)');
  }
  console.log(`開始日: ${projectStart.getFullYear()}/${projectStart.getMonth() + 1}/${projectStart.getDate()}`);

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
      const input = readlineSync.question('ガントチャートの粒度を選んでください [day=日単位 / week=週単位] (デフォルト: week): ');
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

  // --- 祝日データのロード ---
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

  // --- タスク読み込み ---
  let taskData;
  try {
    taskData = await readTasksFromExcel(inputPath);
  } catch (err) {
    console.error('\nエラー: ' + err.message);
    process.exit(1);
  }
  const { tasks, columnOrder, hasProgressColumn } = taskData;
  console.log(`タスクを${tasks.length}件読み込みました。`);
  if (!hasProgressColumn) {
    console.log('(「進捗」列が見つからなかったため、全タスク進捗0%として出力します)');
  }

  // --- スケジューリング ---
  const today = todayDate();
  let scheduled;
  try {
    scheduled = scheduleTasks(tasks, projectStart, isHoliday);
  } catch (err) {
    console.error('\nエラー: ' + err.message);
    process.exit(1);
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
    isHoliday,
    outputPath,
    projectStart,
    today,
    columnOrder,
    hasProgressColumn,
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

main().catch((err) => {
  console.error('予期しないエラーが発生しました:', err);
  process.exit(1);
});
