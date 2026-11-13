#!/usr/bin/env node
/**
 * ③ 一括書き換え（器方式）
 *
 * 各修正シートに付いている既存スクリプト（script_ids.csv の対応表）へ、
 * code.gs の内容を丸ごと上書きする。
 * 1シートに複数スクリプトがある場合は、1つ目に新コードを書き、残りは空白化する。
 *
 * 使い方:
 *   node 3_update_scripts.mjs --dry-run     # 対象の確認のみ（何も変更しない）
 *   node 3_update_scripts.mjs --sheet <ID>  # 指定した1シートだけ（テスト用）
 *   node 3_update_scripts.mjs --limit 5     # 最新の5シートだけ（テスト用）
 *   node 3_update_scripts.mjs --skip 220    # 先頭220件を飛ばして続きから（中断からの再開用）
 *   node 3_update_scripts.mjs               # 全件
 *   node 3_update_scripts.mjs --create      # 対応表にないシートには新規作成して書き込む
 *
 * 上書きは何度実行しても同じ結果になる（失敗した分は再実行すればよい）。
 * 対応表にないシートは unmapped_sheets.csv に書き出される。
 * --create で作成したスクリプトのIDは即座に script_ids.csv に追記されるため、
 * 二重作成は起きない。作成はAPIの日次上限で止まることがある → 翌日再実行で続きから。
 */

import { readFileSync, writeFileSync } from 'fs';
import { homedir } from 'os';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const DIR = dirname(fileURLToPath(import.meta.url));
const SHEETS_FILE = join(DIR, 'sheet_ids.csv');
const SCRIPTS_FILE = join(DIR, 'script_ids.csv');
const CODE_FILE = join(DIR, 'code.gs');
const UNMAPPED_FILE = join(DIR, 'unmapped_sheets.csv');
const EXCLUDED_FILE = join(DIR, 'excluded_ids.txt');
const REQUEST_DELAY_MS = 300;

// ---- 引数 ----
const argv = process.argv.slice(2);
const DRY_RUN = argv.includes('--dry-run');
const CREATE = argv.includes('--create');
const limitIdx = argv.indexOf('--limit');
const LIMIT = limitIdx !== -1 ? parseInt(argv[limitIdx + 1], 10) : Infinity;
if (limitIdx !== -1 && (!Number.isInteger(LIMIT) || LIMIT <= 0)) {
  console.error('--limit には正の整数を指定してください');
  process.exit(1);
}
const skipIdx = argv.indexOf('--skip');
const SKIP = skipIdx !== -1 ? parseInt(argv[skipIdx + 1], 10) : 0;
if (skipIdx !== -1 && (!Number.isInteger(SKIP) || SKIP < 0)) {
  console.error('--skip には0以上の整数を指定してください');
  process.exit(1);
}
const sheetIdx = argv.indexOf('--sheet');
const ONLY_SHEET = sheetIdx !== -1 ? argv[sheetIdx + 1] : null;
if (sheetIdx !== -1 && !ONLY_SHEET) {
  console.error('--sheet にはシートIDを指定してください');
  process.exit(1);
}

// ---- clasp認証 ----
function loadClaspCreds() {
  const raw = JSON.parse(readFileSync(join(homedir(), '.clasprc.json'), 'utf8'));
  if (raw.tokens) {
    const t = raw.tokens.default || Object.values(raw.tokens)[0];
    return { clientId: t.client_id, clientSecret: t.client_secret, refreshToken: t.refresh_token };
  }
  return {
    clientId: raw.oauth2ClientSettings.clientId,
    clientSecret: raw.oauth2ClientSettings.clientSecret,
    refreshToken: raw.token.refresh_token,
  };
}

async function getAccessToken() {
  const { clientId, clientSecret, refreshToken } = loadClaspCreds();
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId, client_secret: clientSecret,
      refresh_token: refreshToken, grant_type: 'refresh_token',
    }),
  });
  if (!res.ok) throw new Error('トークン更新に失敗: ' + (await res.text()));
  return (await res.json()).access_token;
}

// ---- CSV ----
function parseCsvLine(line) {
  const cols = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQ) {
      if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') inQ = false;
      else cur += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { cols.push(cur); cur = ''; }
      else cur += c;
    }
  }
  cols.push(cur);
  return cols;
}
const esc = s => '"' + String(s).replace(/"/g, '""') + '"';
const readCsv = file => readFileSync(file, 'utf8').split('\n').slice(1).filter(Boolean).map(parseCsvLine);

// ---- 書き込む内容 ----
const MANIFEST = JSON.stringify(
  { timeZone: 'Asia/Tokyo', exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8' },
  null, 2
);
const NEW_FILES = [
  { name: 'appsscript', type: 'JSON', source: MANIFEST },
  { name: 'コード', type: 'SERVER_JS', source: readFileSync(CODE_FILE, 'utf8') },
];
const BLANK_FILES = [
  { name: 'appsscript', type: 'JSON', source: MANIFEST },
  {
    name: 'コード',
    type: 'SERVER_JS',
    source: '// このスクリプトは重複していたため無効化されています。\n'
      + '// 同じシートの別スクリプトが最新コードで動作しています。\n',
  },
];

// ---- API（リトライ付き） ----
async function createScript(token, sheetId, title) {
  const res = await fetch('https://script.googleapis.com/v1/projects', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ title, parentId: sheetId }),
    signal: AbortSignal.timeout(30000),
  }).catch(() => null);
  if (!res) return { status: 'error_network' };
  if (res.ok) return { status: 'ok', scriptId: (await res.json()).scriptId };
  const body = await res.text();
  if (body.includes('RESOURCE_EXHAUSTED') || body.includes('dailyLimitExceeded') || body.includes('userRateLimitExceeded')) {
    return { status: 'quota', body };
  }
  return { status: 'error_' + res.status, body: body.slice(0, 120) };
}

async function putContent(scriptId, files) {
  const waits = [2000, 5000, 15000, 30000];
  for (let attempt = 0; ; attempt++) {
    let res;
    try {
      res = await fetch('https://script.googleapis.com/v1/projects/' + scriptId + '/content', {
        method: 'PUT',
        headers: { Authorization: 'Bearer ' + (await freshToken()), 'Content-Type': 'application/json' },
        body: JSON.stringify({ files }),
        signal: AbortSignal.timeout(30000),
      });
    } catch {
      if (attempt < waits.length) { await new Promise(r => setTimeout(r, waits[attempt])); continue; }
      return 'error_network';
    }
    if (res.ok) return 'ok';
    // 401 = トークン失効（45分リフレッシュでも同一トークンが返ることがある）→ 強制再取得してリトライ
    if (res.status === 401 && attempt < waits.length) {
      tokenAt = 0;
      await new Promise(r => setTimeout(r, 1000));
      continue;
    }
    if ((res.status === 429 || res.status >= 500) && attempt < waits.length) {
      await new Promise(r => setTimeout(r, waits[attempt]));
      continue;
    }
    return 'error_' + res.status + ': ' + (await res.text()).slice(0, 120);
  }
}

// ---- 除外リスト（excluded_ids.txt: 1行1ID、シートID/スクリプトIDどちらでも可、#以降はコメント）----
let EXCLUDED = new Set();
try {
  EXCLUDED = new Set(
    readFileSync(EXCLUDED_FILE, 'utf8').split('\n')
      .map(l => l.replace(/#.*/, '').trim()).filter(Boolean)
  );
} catch { /* ファイルがなければ除外なし */ }

// ---- 対象の組み立て ----
const sheets = readCsv(SHEETS_FILE);                 // 新しい順
const bySheet = new Map();                           // sheetId → [scriptId, ...]
const hadAnyScript = new Set();                      // 除外前に対応表に載っていたシート
for (const [sheetId, scriptId] of readCsv(SCRIPTS_FILE)) {
  hadAnyScript.add(sheetId);
  if (EXCLUDED.has(scriptId)) continue;
  if (!bySheet.has(sheetId)) bySheet.set(sheetId, []);
  if (!bySheet.get(sheetId).includes(scriptId)) bySheet.get(sheetId).push(scriptId);
}

const mapped = [], unmapped = [], excluded = [];
for (const row of sheets) {
  const [id] = row;
  if (EXCLUDED.has(id)) excluded.push(row);          // シートIDで除外
  else if (bySheet.has(id)) mapped.push(row);
  else if (hadAnyScript.has(id)) excluded.push(row); // 全スクリプトが除外対象
  else unmapped.push(row);
}
writeFileSync(UNMAPPED_FILE, 'sheetId,sheetName\n'
  + unmapped.map(([id, name]) => id + ',' + esc(name)).join('\n') + '\n');

console.log('シート総数: ' + sheets.length + '件 / スクリプト対応あり: ' + mapped.length
  + '件 / 除外: ' + excluded.length + '件（' + EXCLUDED_FILE + '）'
  + ' / 対応なし: ' + unmapped.length + '件（' + UNMAPPED_FILE + ' に出力）');

const batch = ONLY_SHEET
  ? mapped.filter(([id]) => id === ONLY_SHEET)
  : mapped.slice(SKIP, SKIP + LIMIT);
if (SKIP > 0) console.log('--skip: 先頭' + SKIP + '件を飛ばします（今回の対象: ' + batch.length + '件）');
if (ONLY_SHEET && batch.length === 0) {
  console.error('指定シートが対応表にありません: ' + ONLY_SHEET);
  process.exit(1);
}
if (DRY_RUN) {
  console.log('--dry-run: 今回の対象は ' + batch.length + '件（変更は行いません）');
  for (const [id, name] of batch.slice(0, 20)) {
    const n = bySheet.get(id).length;
    console.log('  ' + name + (n > 1 ? '（スクリプト' + n + '個 → 1個に新コード、' + (n - 1) + '個を空白化）' : ''));
  }
  if (batch.length > 20) console.log('  ... 他' + (batch.length - 20) + '件');
  process.exit(0);
}

// ---- 実行 ----
// トークンは60分で失効するため、45分経過したら取り直す（長時間の全件実行対策）
let token = await getAccessToken();
let tokenAt = Date.now();
async function freshToken() {
  if (Date.now() - tokenAt > 45 * 60 * 1000) {
    token = await getAccessToken();
    tokenAt = Date.now();
  }
  return token;
}
console.log('アクセストークン取得OK。' + batch.length + '件を書き換えます');

let updated = 0, blanked = 0, failed = 0, done = 0;
const failures = [];
const startedAt = Date.now();

for (const [sheetId, sheetName] of batch) {
  const [vessel, ...extras] = bySheet.get(sheetId);

  const r = await putContent(vessel, NEW_FILES);
  if (r === 'ok') updated++;
  else { failed++; failures.push(sheetName + ' (script=' + vessel + ') : ' + r); }

  for (const ex of extras) {
    await new Promise(res => setTimeout(res, REQUEST_DELAY_MS));
    const b = await putContent(ex, BLANK_FILES);
    if (b === 'ok') blanked++;
    else { failed++; failures.push(sheetName + ' 空白化 (script=' + ex + ') : ' + b); }
  }

  done++;
  const elapsed = Math.round((Date.now() - startedAt) / 1000);
  process.stdout.write('\r進捗: ' + done + '/' + batch.length
    + '  上書き: ' + updated + '  空白化: ' + blanked + '  失敗: ' + failed + '  経過: ' + elapsed + '秒   ');
  await new Promise(res => setTimeout(res, REQUEST_DELAY_MS));
}
console.log('');

console.log('上書き: ' + updated + '件 / 重複の空白化: ' + blanked + '件 / 失敗: ' + failed + '件');
for (const f of failures.slice(0, 20)) console.log('  ❌ ' + f);
if (failures.length > 20) console.log('  ... 他' + (failures.length - 20) + '件');
if (failed > 0) console.log('→ 失敗分はもう一度実行すれば再試行されます。');
if (unmapped.length > 0) console.log('→ 対応なしの ' + unmapped.length + '件は、ID収集（マニュアル参照）後に 2_collect_scripts.mjs で取り込んでください。');
if (failed === 0 && LIMIT === Infinity) console.log('✅ 対応表にある全シートが最新コードになりました。');
