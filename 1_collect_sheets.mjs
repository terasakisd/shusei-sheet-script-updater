#!/usr/bin/env node
/**
 * ① 全修正シートの収集
 *
 * 対象フォルダ（サブフォルダ含む）の全スプレッドシートを取得し、
 * 作成日の新しい順で sheet_ids.csv に書き出す。
 *
 * 使い方: node 1_collect_sheets.mjs
 */

import { readFileSync, writeFileSync } from 'fs';
import { homedir } from 'os';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const DIR = dirname(fileURLToPath(import.meta.url));
const TARGET_FOLDER_ID = '1n1WZZSYDzQRVL2V1Dn08OqNPtkR9m2Ie';
const OUT_FILE = join(DIR, 'sheet_ids.csv');

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

const esc = s => '"' + String(s).replace(/"/g, '""') + '"';

// ---- フォルダを再帰的に検索 ----
const token = await getAccessToken();
const sheets = [];
const queue = [TARGET_FOLDER_ID];
const seen = new Set();
while (queue.length > 0) {
  const fid = queue.shift();
  if (seen.has(fid)) continue;
  seen.add(fid);
  let pageToken = '';
  do {
    const q = `'${fid}' in parents and trashed=false and (mimeType='application/vnd.google-apps.spreadsheet' or mimeType='application/vnd.google-apps.folder')`;
    const url = 'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q)
      + '&pageSize=1000&fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType,createdTime)')
      + '&supportsAllDrives=true&includeItemsFromAllDrives=true'
      + (pageToken ? '&pageToken=' + pageToken : '');
    const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    if (!res.ok) throw new Error('Drive API エラー: ' + (await res.text()));
    const json = await res.json();
    for (const f of json.files || []) {
      if (f.mimeType === 'application/vnd.google-apps.folder') queue.push(f.id);
      else sheets.push(f);
    }
    pageToken = json.nextPageToken || '';
  } while (pageToken);
}
sheets.sort((a, b) => String(b.createdTime).localeCompare(String(a.createdTime))); // 新しい順

writeFileSync(OUT_FILE, 'sheetId,sheetName,createdTime\n'
  + sheets.map(s => s.id + ',' + esc(s.name) + ',' + (s.createdTime || '').slice(0, 10)).join('\n') + '\n');
console.log('✅ ' + sheets.length + '件を ' + OUT_FILE + ' に書き出しました（作成日の新しい順）');
console.log('最新5件:');
for (const s of sheets.slice(0, 5)) console.log('  ' + (s.createdTime || '').slice(0, 10) + '  ' + s.name);
