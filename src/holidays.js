'use strict';

const fs = require('fs');
const path = require('path');
const https = require('https');

const HOLIDAY_CSV_URL = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
// 内閣府サイトは過去に一時的にファイル名が syukujitsu.csv <-> shukujitsu.csv で
// 変更されたことがあるため、メインURLが失敗した場合のフォールバックを用意する。
const HOLIDAY_CSV_URL_FALLBACK = 'https://www8.cao.go.jp/chosei/shukujitsu/shukujitsu.csv';
const CACHE_FILENAME = 'holidays_cache.csv';

/**
 * 年末年始の固定休業期間 (12/27 ~ 1/3)
 * 月日のみで判定する (年をまたぐ)
 */
function isYearEndHoliday(date) {
  const m = date.getMonth() + 1; // 1-12
  const d = date.getDate();
  if (m === 12 && d >= 27) return true;
  if (m === 1 && d <= 3) return true;
  return false;
}

/**
 * キャッシュファイルのパスを決定する。
 * スクリプト自身のディレクトリ(src の一つ上)に置く。
 */
function getCachePath() {
  return path.join(__dirname, '..', CACHE_FILENAME);
}

/**
 * Shift-JIS バイト列を簡易的に UTF-8 文字列に変換する。
 * Node.js 標準では Shift-JIS デコードができないため、
 * iconv-lite 等のライブラリを使う。
 */
function decodeShiftJIS(buffer) {
  const iconv = require('iconv-lite');
  return iconv.decode(buffer, 'Shift_JIS');
}

/**
 * 内閣府CSVのテキストをパースして {YYYY-MM-DD: 祝日名} の配列にする。
 * 内閣府CSVの想定フォーマット:
 *   国民の祝日・休日月日,国民の祝日・休日名称
 *   2025/1/1,元日
 *   2025/1/13,成人の日
 *   ...
 * 過去にはハイフン区切り(2017-01-09)の年もあったため、スラッシュ・ハイフン両方に対応する。
 */
function parseCaoCsv(text) {
  const lines = text.split(/\r\n|\n|\r/).filter((l) => l.trim().length > 0);
  const result = [];
  // 先頭行はヘッダーなのでスキップ。日付として解釈できない行はスキップする。
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    const cols = line.split(',');
    if (cols.length < 2) continue;
    const rawDate = cols[0].trim();
    const name = cols.slice(1).join(',').trim();
    const m = rawDate.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (!m) continue;
    const yyyy = m[1];
    const mm = String(m[2]).padStart(2, '0');
    const dd = String(m[3]).padStart(2, '0');
    result.push({ date: `${yyyy}-${mm}-${dd}`, name });
  }
  return result;
}

/**
 * 指定したURLからファイルをダウンロードする (https モジュールのみで実装、追加依存なし)。
 * 戻り値: Buffer
 */
function downloadUrl(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode !== 200) {
        reject(new Error(`HTTP ${res.statusCode}`));
        return;
      }
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

/**
 * 内閣府サイトから祝日CSVをダウンロードする。
 * メインURLが失敗した場合はフォールバックURL(ファイル名違い)も試す。
 * 戻り値: Buffer (Shift-JIS のバイト列)
 */
async function downloadCaoCsv() {
  try {
    return await downloadUrl(HOLIDAY_CSV_URL);
  } catch (err1) {
    console.log(`  メインURLの取得に失敗しました (${err1.message})。フォールバックURLを試します...`);
    console.log(`  ${HOLIDAY_CSV_URL_FALLBACK}`);
    try {
      return await downloadUrl(HOLIDAY_CSV_URL_FALLBACK);
    } catch (err2) {
      throw new Error(
        `祝日CSVの取得に失敗しました。内閣府サイトのURLが変更されている可能性があります。\n` +
        `  試したURL: ${HOLIDAY_CSV_URL} (${err1.message})\n` +
        `  試したURL: ${HOLIDAY_CSV_URL_FALLBACK} (${err2.message})`
      );
    }
  }
}

/**
 * キャッシュCSVを書き出す。フォーマット: date,name (UTF-8, ヘッダー行あり)
 */
function writeCache(holidayList) {
  const cachePath = getCachePath();
  const lines = ['date,name'];
  for (const h of holidayList) {
    // 名称にカンマが含まれる可能性は低いが、念のためダブルクオートで囲む
    const safeName = /[",\n]/.test(h.name) ? `"${h.name.replace(/"/g, '""')}"` : h.name;
    lines.push(`${h.date},${safeName}`);
  }
  fs.writeFileSync(cachePath, lines.join('\n'), 'utf8');
  return cachePath;
}

/**
 * キャッシュCSVを読み込む。
 * 戻り値: [{ date: "YYYY-MM-DD", name: string }]
 */
function readCache() {
  const cachePath = getCachePath();
  const text = fs.readFileSync(cachePath, 'utf8');
  const lines = text.split(/\r\n|\n|\r/).filter((l) => l.trim().length > 0);
  const list = [];
  for (let i = 1; i < lines.length; i++) {
    const idx = lines[i].indexOf(',');
    if (idx === -1) continue;
    const date = lines[i].slice(0, idx).trim();
    let name = lines[i].slice(idx + 1).trim();
    // ダブルクオートで囲まれている場合は外す
    if (name.startsWith('"') && name.endsWith('"')) {
      name = name.slice(1, -1).replace(/""/g, '"');
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) list.push({ date, name });
  }
  return list;
}

/**
 * 指定した期間(start〜end、両端を含む)に含まれる年末年始休暇(12/27〜1/3)の
 * 日付一覧を名称付きで生成する。
 *
 * @returns {Array<{date: string, name: string}>}
 */
function buildYearEndHolidayList(startDate, endDate) {
  const list = [];
  let cur = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (cur <= end) {
    if (isYearEndHoliday(cur)) {
      list.push({ date: formatYmd(cur), name: '年末年始休暇' });
    }
    cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
  }
  return list;
}

/**
 * 祝日セットを取得する。
 * キャッシュファイルが存在すればそれを使用し、ネットワークアクセスは行わない。
 * 存在しない場合は内閣府サイトから取得してキャッシュを作成する。
 *
 * 戻り値: { dateSet: Set<"YYYY-MM-DD">, usedCache: boolean, cachePath: string }
 */
async function loadHolidays() {
  const cachePath = getCachePath();
  if (fs.existsSync(cachePath)) {
    const holidayList = readCache();
    const dateSet = new Set(holidayList.map((h) => h.date));
    return { dateSet, holidayList, usedCache: true, cachePath };
  }

  console.log('祝日キャッシュが見つかりません。内閣府サイトから祝日データを取得します...');
  console.log(`  URL: ${HOLIDAY_CSV_URL}`);
  let buffer;
  try {
    buffer = await downloadCaoCsv();
  } catch (err) {
    throw new Error(
      `祝日データの取得に失敗しました。ネットワークに接続できる環境で一度実行し、` +
      `生成された ${CACHE_FILENAME} をこのフォルダに置いてから再実行してください。\n詳細: ${err.message}`
    );
  }
  const text = decodeShiftJIS(buffer);
  const holidayList = parseCaoCsv(text);
  if (holidayList.length === 0) {
    throw new Error('祝日CSVの解析結果が0件でした。内閣府サイトのCSV形式が変更されている可能性があります。');
  }
  writeCache(holidayList);
  console.log(`祝日データを取得しました (${holidayList.length}件)。キャッシュを作成しました: ${cachePath}`);

  const dateSet = new Set(holidayList.map((h) => h.date));
  return { dateSet, holidayList, usedCache: false, cachePath };
}

function formatYmd(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * 指定日が休日(土日・祝日・年末年始)かどうかを判定する関数を生成する。
 */
function makeIsHoliday(holidayDateSet) {
  return function isHoliday(date) {
    const dow = date.getDay(); // 0:日 6:土
    if (dow === 0 || dow === 6) return true;
    if (isYearEndHoliday(date)) return true;
    if (holidayDateSet.has(formatYmd(date))) return true;
    return false;
  };
}

module.exports = {
  loadHolidays,
  makeIsHoliday,
  formatYmd,
  isYearEndHoliday,
  buildYearEndHolidayList,
  getCachePath,
  CACHE_FILENAME,
  HOLIDAY_CSV_URL,
  parseCaoCsv,
};
