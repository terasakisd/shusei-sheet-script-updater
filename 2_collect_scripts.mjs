#!/usr/bin/env node
/**
 * ② スクリプトIDの収集（ダッシュボードから集めたIDの逆引き）
 *
 * ブラウザで script.google.com/home/search?q=無題のプロジェクト を開き、
 * マニュアル.md の収集スニペットでID一覧をコピーしたあと:
 *
 *   pbpaste > collected_ids.txt
 *   node 2_collect_scripts.mjs
 *
 * やること:
 *   - collected_ids.txt からスクリプトIDを抽出（余計な行が混ざっていてもOK）
 *   - 既に script_ids.csv にあるIDは除外
 *   - 残りをAPIで逆引きし、対象フォルダの修正シートに付いているものだけを
 *     script_ids.csv に追記する
 *
 * 何度実行しても二重追記はされない。
 */

import { readFileSync, appendFileSync, existsSync } from 'fs';
import { homedir } from 'os';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const DIR = dirname(fileURLToPath(import.meta.url));
const INPUT_FILE = join(DIR, 'collected_ids.txt');
const SHEETS_FILE = join(DIR, 'sheet_ids.csv');
const SCRIPTS_FILE = join(DIR, 'script_ids.csv');
const REQUEST_DELAY_MS = 250;

// 触ってはいけないスクリプト（所要時間列ツールなど）
const EXCLUDE = new Set([
  '1do4FMWOiL-tHrBdenBQh0FGTW2C6-tRj_JQ8-MTyyV_TdTcBesyi-ek2', // 所要時間列 一括追加ツール
]);

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

// ---- 入力 ----
if (!existsSync(INPUT_FILE)) {
  console.error('collected_ids.txt がありません。');
  console.error('ブラウザの収集スニペット実行後に  pbpaste > collected_ids.txt  を実行してください。');
  process.exit(1);
}
const candidates = [...new Set(readFileSync(INPUT_FILE, 'utf8').match(/\b1[A-Za-z0-9_-]{50,70}\b/g) || [])];
console.log('入力ID: ' + candidates.length + '件');

const known = new Set(readCsv(SCRIPTS_FILE).map(r => r[1]));
const sheetIds = new Set(readCsv(SHEETS_FILE).map(r => r[0]));
const sheetNames = new Map(readCsv(SHEETS_FILE).map(r => [r[0], r[1]]));

const targets = candidates.filter(id => !known.has(id) && !EXCLUDE.has(id));
console.log('照会対象: ' + targets.length + '件（既知 ' + (candidates.length - targets.length) + '件を除外）');
if (targets.length === 0) {
  console.log('新しく照会するIDはありません。');
  process.exit(0);
}

// ---- 逆引き ----
const token = await getAccessToken();
console.log('逆引きを開始します（約' + Math.ceil(targets.length * REQUEST_DELAY_MS / 60000) + '分）');

let added = 0, outside = 0, inaccessible = 0, done = 0;
const newRows = [];
for (const id of targets) {
  let res = null;
  for (let attempt = 0; ; attempt++) {
    res = await fetch('https://script.googleapis.com/v1/projects/' + id, {
      headers: { Authorization: 'Bearer ' + token },
      signal: AbortSignal.timeout(30000),
    }).catch(() => null);
    if (res && res.status === 429 && attempt < 3) {
      await new Promise(r => setTimeout(r, 5000 * (attempt + 1)));
      continue;
    }
    break;
  }
  if (!res || !res.ok) {
    inaccessible++;
  } else {
    const p = await res.json();
    if (p.parentId && sheetIds.has(p.parentId)) {
      newRows.push(p.parentId + ',' + id + ',' + esc(sheetNames.get(p.parentId) || p.title || ''));
      added++;
    } else {
      outside++;
    }
  }
  done++;
  if (done % 20 === 0 || done === targets.length) {
    process.stdout.write('\r進捗: ' + done + '/' + targets.length
      + '  修正シートの旧スクリプト: ' + added + '  対象外: ' + outside + '  アクセス不可: ' + inaccessible + '   ');
  }
  await new Promise(r => setTimeout(r, REQUEST_DELAY_MS));
}
console.log('');

if (newRows.length > 0) {
  const cur = readFileSync(SCRIPTS_FILE, 'utf8');
  appendFileSync(SCRIPTS_FILE, (cur.endsWith('\n') ? '' : '\n') + newRows.join('\n') + '\n');
}

const covered = new Set(readCsv(SCRIPTS_FILE).map(r => r[0]));
const coveredCount = [...sheetIds].filter(id => covered.has(id)).length;
console.log('✅ script_ids.csv に ' + added + '件を追記しました');
console.log('対応表カバー率: ' + coveredCount + ' / ' + sheetIds.size + ' シート');
console.log('→ 反映は node 3_update_scripts.mjs で行います。');
