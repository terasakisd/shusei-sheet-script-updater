// Chatwork API設定
// ROOM_IDはA2セルから、API_TOKENは担当者の苗字から管理シートより取得します

// APIトークン管理シートのID
const TOKEN_SHEET_ID = '1v4NWkZXrlGcSrPDC0YgZC1G1OpQK5PNr0G2DSQcINFY';

// 管理サイト一覧シートのID
const SITE_MANAGEMENT_SHEET_ID = '1B5Vs1MKETRsZ3yjwkU-yLDSPqIFb9wc2R500mfpDW_Q';

/**
 * カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📤 Chatwork')
    .addItem('サイト一覧を更新', 'updateSiteNameDropdown')
    .addSeparator()
    .addItem('ridを設定', 'setRidFromA1')
    .addItem('サイトを検索してridを設定', 'showSiteSearchDialog')
    .addSeparator()
    .addItem('選択行を送信', 'sendCurrentRowToChatwork')
    .addToUi();
  
  // スプレッドシートを開いたときに自動でドロップダウンを設定
  setupSiteNameDropdown();

  // I列に「所要時間」列がなければ自動で挿入
  ensureDurationColumn();

  // J列に「記事全体確認」列がなければ自動で挿入（チェックボックス付き）
  ensureArticleCheckColumn();
}

/**
 * I列に「所要時間」列を追加する
 * 対象: A列に「日付」があるヘッダー行を持つタブ
 * すでにI列ヘッダーが「所要時間」のタブはスキップするので、二重には増えない
 */
function ensureDurationColumn() {
  try {
    const tabs = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (let i = 0; i < tabs.length; i++) {
      const tab = tabs[i];
      const scanRows = Math.min(10, tab.getLastRow());
      if (scanRows < 1) continue;
      const colA = tab.getRange(1, 1, scanRows, 1).getDisplayValues();
      let headerRow = 0;
      for (let r = 0; r < colA.length; r++) {
        if (String(colA[r][0]).trim() === '日付') { headerRow = r + 1; break; }
      }
      if (headerRow === 0) continue; // 作業タブではない
      const max = tab.getMaxColumns();
      if (max >= 9) {
        const current = String(tab.getRange(headerRow, 9).getDisplayValue()).trim();
        if (current === '所要時間') continue; // 追加済み
        tab.insertColumnBefore(9);
      } else {
        tab.insertColumnsAfter(max, 9 - max); // 列が9未満しかないタブは末尾まで拡張
      }
      tab.getRange(headerRow, 9).setValue('所要時間');
    }
  } catch (e) {
    Logger.log('所要時間列の追加エラー: ' + e.toString());
  }
}

/**
 * J列に「記事全体確認」列を追加する（データ行はチェックボックス）
 * 対象: A列に「日付」があるヘッダー行を持つタブ
 * すでにJ列ヘッダーが「記事全体確認」のタブはスキップするので、二重には増えない
 * ※ ensureDurationColumn の後に呼ぶこと（I列=所要時間が入っている前提）
 */
function ensureArticleCheckColumn() {
  try {
    const tabs = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (let i = 0; i < tabs.length; i++) {
      const tab = tabs[i];
      const scanRows = Math.min(10, tab.getLastRow());
      if (scanRows < 1) continue;
      const colA = tab.getRange(1, 1, scanRows, 1).getDisplayValues();
      let headerRow = 0;
      for (let r = 0; r < colA.length; r++) {
        if (String(colA[r][0]).trim() === '日付') { headerRow = r + 1; break; }
      }
      if (headerRow === 0) continue; // 作業タブではない
      // I列が「所要時間」になっていないタブは列構成が想定外なので触らない
      if (String(tab.getRange(headerRow, 9).getDisplayValue()).trim() !== '所要時間') continue;
      const max = tab.getMaxColumns();
      if (max >= 10) {
        const current = String(tab.getRange(headerRow, 10).getDisplayValue()).trim();
        if (current === '記事全体確認') continue; // 追加済み
        tab.insertColumnAfter(9);
      } else {
        tab.insertColumnsAfter(max, 10 - max); // 列が10未満しかないタブは末尾まで拡張
      }
      tab.getRange(headerRow, 10).setValue('記事全体確認');
      const maxRows = tab.getMaxRows();
      if (maxRows > headerRow) {
        tab.getRange(headerRow + 1, 10, maxRows - headerRow, 1).insertCheckboxes();
      }
    }
  } catch (e) {
    Logger.log('記事全体確認列の追加エラー: ' + e.toString());
  }
}

/**
 * サイト検索ダイアログを表示
 */
function showSiteSearchDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          #searchBox { width: 100%; padding: 10px; font-size: 16px; margin-bottom: 10px; box-sizing: border-box; }
          #siteList { max-height: 400px; overflow-y: auto; border: 1px solid #ccc; }
          .site-item { padding: 10px; cursor: pointer; border-bottom: 1px solid #eee; }
          .site-item:hover { background-color: #f0f0f0; }
          .loading { text-align: center; padding: 20px; color: #666; }
          .hidden { display: none; }
        </style>
      </head>
      <body>
        <h3>サイト名を検索</h3>
        <input type="text" id="searchBox" placeholder="読み込み中..." disabled autofocus>
        <div id="siteList">
          <div class="loading">サイト一覧を読み込んでいます...</div>
        </div>
        
        <script>
          let allSites = [];
          let isLoaded = false;
          
          // サイト一覧を取得
          google.script.run.withSuccessHandler(function(sites) {
            allSites = sites;
            isLoaded = true;
            
            // 検索ボックスを有効化
            const searchBox = document.getElementById('searchBox');
            searchBox.disabled = false;
            searchBox.placeholder = 'サイト名を入力...';
            searchBox.focus();
            
            // 全サイトを表示
            displaySites(sites);
          }).withFailureHandler(function(error) {
            document.getElementById('siteList').innerHTML = '<div class="loading">エラー: サイト一覧の読み込みに失敗しました</div>';
          }).getAllSites();
          
          // 検索機能
          document.getElementById('searchBox').addEventListener('input', function(e) {
            if (!isLoaded) return; // 読み込み完了まで無視
            
            const query = e.target.value.toLowerCase();
            const filtered = allSites.filter(site => site.toLowerCase().includes(query));
            displaySites(filtered);
          });
          
          // サイト一覧を表示
          function displaySites(sites) {
            const list = document.getElementById('siteList');
            list.innerHTML = '';
            
            if (sites.length === 0) {
              list.innerHTML = '<div class="loading">該当するサイトが見つかりません</div>';
              return;
            }
            
            sites.forEach(site => {
              const div = document.createElement('div');
              div.className = 'site-item';
              div.textContent = site;
              div.onclick = function() {
                google.script.host.close();
                google.script.run.setSiteNameAndRid(site);
              };
              list.appendChild(div);
            });
          }
        </script>
      </body>
    </html>
  `).setWidth(500).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'サイトを選択');
}

/**
 * 全サイト名を取得（キャッシュ対応）
 */
function getAllSites() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('site_list');
    
    // キャッシュがあれば返す（有効期限: 10分）
    if (cached) {
      return JSON.parse(cached);
    }
    
    // キャッシュがなければ取得
    SpreadsheetApp.flush();
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    
    if (lastRow < 2) {
      return [];
    }
    
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues()
      .map(row => row[0])
      .filter(name => name !== '' && name !== null);
    
    // キャッシュに保存（10分間）
    cache.put('site_list', JSON.stringify(siteNames), 600);
    
    return siteNames;
  } catch (e) {
    Logger.log('サイト一覧取得エラー: ' + e.toString());
    return [];
  }
}

/**
 * サイト名を設定してridを自動取得
 */
function setSiteNameAndRid(siteName) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('A1').setValue(siteName);
  
  // ridを自動設定
  try {
    SpreadsheetApp.flush();
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    
    if (lastRow < 2) {
      return;
    }
    
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const rids = siteSheet.getRange(2, 11, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < siteNames.length; i++) {
      if (siteNames[i][0] === siteName) {
        const rid = rids[i][0];
        if (rid !== null && rid !== '') {
          sheet.getRange('A2').setValue(String(rid));
          return;
        }
      }
    }
  } catch (e) {
    Logger.log('rid設定エラー: ' + e.toString());
  }
}

/**
 * 最新5件のサイトを確認
 */
function checkLastSites() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    let info = `スプレッドシートの構造:\n`;
    info += `シート数: ${sheets.length}枚\n\n`;
    
    for (let i = 0; i < sheets.length; i++) {
      info += `${i}番目: ${sheets[i].getName()}\n`;
    }
    
    info += `\n現在使用中: ${sheets[1].getName()}\n\n`;
    
    const siteSheet = sheets[1];
    const lastRow = siteSheet.getLastRow();
    
    const startRow = Math.max(2, lastRow - 4);
    const data = siteSheet.getRange(startRow, 2, lastRow - startRow + 1, 10).getValues();
    
    info += `\n最新のサイト（${startRow}行目〜${lastRow}行目）:\n`;
    for (let i = 0; i < data.length; i++) {
      const siteName = data[i][0]; // B列
      const rid = data[i][9]; // K列
      info += `行${startRow + i}: [${siteName}] → rid: [${rid}]\n`;
    }
    
    SpreadsheetApp.getUi().alert(info);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.toString());
  }
}

/**
 * サイト一覧を手動で更新
 */
function updateSiteNameDropdown() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // キャッシュをクリアして再読み込み
    SpreadsheetApp.flush();
    
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    
    // 最新データを取得
    const allData = siteSheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues(); // getValues → getDisplayValues
    const siteNames = allData.map(row => row[0]).filter(name => name !== '' && name !== null);
    
    setupSiteNameDropdown();
    
    SpreadsheetApp.getUi().alert(`✅ サイト一覧を更新しました！`);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ エラー: ' + e.toString());
  }
}

/**
 * A1のサイト名からridを取得してA2に設定
 */
function setRidFromA1() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const siteName = sheet.getRange('A1').getValue();
  
  if (!siteName || siteName === '') {
    SpreadsheetApp.getUi().alert('❌ A1セルにサイト名が入力されていません。');
    return;
  }
  
  try {
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('❌ 管理サイト一覧にデータがありません。');
      return;
    }
    
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const rids = siteSheet.getRange(2, 11, lastRow - 1, 1).getValues();
    
    let foundRid = null;
    for (let i = 0; i < siteNames.length; i++) {
      if (siteNames[i][0] === siteName) {
        foundRid = rids[i][0];
        break;
      }
    }
    
    if (foundRid !== null && foundRid !== '') {
      sheet.getRange('A2').setValue(String(foundRid));
      SpreadsheetApp.getUi().alert(`✅ ridを設定しました: ${foundRid}`);
    } else {
      sheet.getRange('A2').setValue('');
      SpreadsheetApp.getUi().alert(`❌ 「${siteName}」のridが見つかりませんでした。`);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ エラー: ' + e.toString());
  }
}

/**
 * デバッグ用：A1の値と一致確認
 */
function debugA1Match() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const a1Value = sheet.getRange('A1').getValue();
    
    if (!a1Value || a1Value === '') {
      SpreadsheetApp.getUi().alert('A1セルが空です。');
      return;
    }
    
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const rids = siteSheet.getRange(2, 11, lastRow - 1, 1).getValues();
    
    let debugInfo = `A1の値: [${a1Value}] (文字数: ${a1Value.length})\n\n`;
    debugInfo += '一致チェック:\n';
    
    for (let i = 0; i < Math.min(25, siteNames.length); i++) {
      const siteName = siteNames[i][0];
      const rid = rids[i][0];
      const match = siteName === a1Value ? '✅ 一致' : '';
      debugInfo += `[${siteName}] (文字数: ${String(siteName).length}) → ${rid} ${match}\n`;
    }
    
    SpreadsheetApp.getUi().alert(debugInfo);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.toString());
  }
}

/**
 * デバッグ用：サイト一覧とridを表示
 */
function debugSiteList() {
  try {
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    const data = siteSheet.getRange(2, 1, Math.min(25, lastRow - 1), 11).getValues();
    
    let debugInfo = 'サイト一覧（最初の25件）:\n\n';
    for (let i = 0; i < data.length; i++) {
      const siteName = data[i][1]; // B列
      const rid = data[i][10]; // K列
      debugInfo += `行${i + 2}: [${siteName}] → rid: [${rid}]\n`;
    }
    
    SpreadsheetApp.getUi().alert(debugInfo);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.toString());
  }
}

/**
 * A1セルにサイト名のドロップダウンを設定
 */
function setupSiteNameDropdown() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // 管理サイト一覧シートを開く（2番目のタブ = index 1）
    SpreadsheetApp.flush(); // キャッシュをクリア
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    
    // B列のサイト名を取得（2行目から最終行まで）
    const lastRow = siteSheet.getLastRow();
    if (lastRow < 2) {
      return; // データがない場合は何もしない
    }
    
    // データを強制的に再読み込み
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues()
      .map(row => row[0])
      .filter(name => name !== '' && name !== null);
    
    // A1セルにデータ検証（ドロップダウン）を設定
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(siteNames, true)
      .setAllowInvalid(false)
      .build();
    
    sheet.getRange('A1').setDataValidation(rule);
  } catch (e) {
    // エラーが出てもスプレッドシートの読み込みは継続
    Logger.log('ドロップダウン設定エラー: ' + e.toString());
  }
}

/**
 * A1セルが変更されたときに自動実行
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    // A1セル以外が編集された場合は何もしない
    if (range.getA1Notation() !== 'A1') {
      return;
    }
    
    const siteName = range.getValue();
    
    // 空の場合は何もしない
    if (!siteName || siteName === '') {
      return;
    }
    
    // サイト名からridを取得してA2に設定
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    
    if (lastRow < 2) {
      sheet.getRange('A2').setValue('');
      return;
    }
    
    // B列とK列のデータを取得
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列
    const rids = siteSheet.getRange(2, 11, lastRow - 1, 1).getValues(); // K列
    
    // サイト名に一致する行を検索
    let foundRid = null;
    for (let i = 0; i < siteNames.length; i++) {
      if (siteNames[i][0] === siteName) {
        foundRid = rids[i][0];
        break;
      }
    }
    
    if (foundRid !== null && foundRid !== '') {
      sheet.getRange('A2').setValue(String(foundRid));
    } else {
      sheet.getRange('A2').setValue('');
      SpreadsheetApp.getUi().alert(`❌ 「${siteName}」のridが見つかりませんでした。`);
    }
  } catch (error) {
    Logger.log('onEditエラー: ' + error.toString());
  }
}

/**
 * 選択中の行をChatworkに送信
 */
function sendCurrentRowToChatwork() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  // ヘッダー行（3行目以前）は送信しない
  if (row <= 3) {
    SpreadsheetApp.getUi().alert('データ行を選択してください（4行目以降）。');
    return;
  }
  
  // ROOM_IDをA2セルから取得
  const roomId = sheet.getRange('A2').getValue();
  if (!roomId || roomId === '') {
    SpreadsheetApp.getUi().alert('❌ A2セルにルームIDが入力されていません。');
    return;
  }
  
  // 担当者の苗字を選択行のD列から取得
  const staffName = sheet.getRange(row, 4).getValue(); // D列 = 4列目
  if (!staffName || staffName === '') {
    SpreadsheetApp.getUi().alert('❌ この行のD列（担当者）に苗字が入力されていません。');
    return;
  }
  
  // APIトークンを管理シートから取得
  const apiToken = getApiTokenByName(staffName);
  if (!apiToken) {
    SpreadsheetApp.getUi().alert(`❌ 「${staffName}」のAPIトークンが管理シートに登録されていません。\n\n管理シートを確認してください。`);
    return;
  }
  
  // 行データを取得
  const sheetTitle = sheet.getRange(row, 2).getValue(); // B列（実施項目）例：★★修正
  const resultValue = sheet.getRange(row, 5).getValue(); // E列（結果）
  const keyword = sheet.getRange('B2').getValue(); // B2セル（KW）
  const pageUrl = sheet.getRange(`C2:D2`).getValues()[0].join(''); // C2:D2結合（ページURL）
  const content = sheet.getRange(row, 6).getValue(); // F列（内容）
  const rank = sheet.getRange(row, 7).getValue(); // G列（順位）
  const duration = sheet.getRange(row, 9).getDisplayValue(); // I列（所要時間）
  const articleCheck = sheet.getRange(row, 10).getValue(); // J列（記事全体確認・チェックボックス）
  const urlInspection = sheet.getRange(row, 13).getValue(); // M列（URL検査）※I列に所要時間・J列に記事全体確認を挿入したためK→Mに移動
  
  // スプレッドシートのURL（選択行へのリンク）
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheetId = sheet.getSheetId();
  const historyUrl = `${spreadsheetUrl}#gid=${sheetId}&range=${row}:${row}`;
  
  // F列が空の場合はスキップ
  if (!content || content === '') {
    SpreadsheetApp.getUi().alert('❌ この行のF列（内容）が入力されていません。');
    return;
  }
  
  // メッセージを作成
  let message = '';
  
  // タイトル（B列 - 実施項目）
  if (sheetTitle !== '' && sheetTitle !== null) {
    message += `★★${sheetTitle}\n`;
  }
  
  // KW（B2セル）
  if (keyword !== '' && keyword !== null) {
    message += `「${keyword}」\n`;
  }
  
  // ページURL（C2:D2）
  if (pageUrl !== '' && pageUrl !== null) {
    message += `${pageUrl}\n\n`;
  }
  
  // 内容（F列）
  message += `${content}\n`;
  
  // 結果（E列）- B列が「振り返り」の場合のみ追加（URL検査の前）
  if (sheetTitle === '振り返り' && resultValue !== '' && resultValue !== null) {
    message += `\n【結果】\n${resultValue}\n`;
  }
  
  // URL検査（M列）- B列が「記事修正」の場合のみ表示
  if (sheetTitle === '記事修正' && urlInspection !== '' && urlInspection !== null) {
    const inspectionStatus = urlInspection === true ? '検査済み' : '未検査';
    message += `\n【URL検査】\n${inspectionStatus}\n`;
  } else if (sheetTitle === '修正' && urlInspection !== '' && urlInspection !== null) {
    const inspectionStatus = urlInspection === true ? '検査済み' : '未検査';
    message += `\n【URL検査】\n${inspectionStatus}\n`;
  }

  // 所要時間（I列）
  if (duration !== '' && duration !== null) {
    message += `\n【所要時間】\n${duration}\n`;
  }

  // 記事全体確認（J列）- B列が「記事修正」「修正」の場合のみ表示
  if (sheetTitle === '記事修正' || sheetTitle === '修正') {
    message += `\n【記事全体確認】\n${articleCheck === true ? 'チェック済み' : '未チェック'}\n`;
  }

  // 履歴（選択行のスプレッドシートリンク）
  message += `\n【履歴】\n${historyUrl}\n`;
  
  // 順位（G列）
  if (rank !== '' && rank !== null) {
    message += `\n${rank}`;
  }
  
  // Chatworkに送信
  const result = sendToChatwork(message, apiToken, roomId);
  
  if (result) {
    SpreadsheetApp.getUi().alert('✅ Chatworkに送信しました！');
    
    // 送信済みフラグを立てる（Q列にチェック）※I列に所要時間・J列に記事全体確認を挿入したためO→Qに移動
    sheet.getRange(row, 17).setValue('✓ ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm'));
  } else {
    SpreadsheetApp.getUi().alert('❌ 送信に失敗しました。\n\nAPIトークンとルームIDを確認してください。');
  }
}

/**
 * サイト名からridを取得（送信時に使用）
 */
function getRidBySiteName(siteName) {
  try {
    const siteSheet = SpreadsheetApp.openById(SITE_MANAGEMENT_SHEET_ID).getSheets()[1];
    const lastRow = siteSheet.getLastRow();
    
    if (lastRow < 2) {
      return null;
    }
    
    // B列とK列のデータを取得
    const siteNames = siteSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列
    const rids = siteSheet.getRange(2, 11, lastRow - 1, 1).getValues(); // K列
    
    // サイト名に一致する行を検索
    for (let i = 0; i < siteNames.length; i++) {
      if (siteNames[i][0] === siteName) {
        const rid = rids[i][0];
        if (rid !== null && rid !== '') {
          return String(rid);
        }
      }
    }
    
    return null;
  } catch (e) {
    Logger.log('rid取得エラー: ' + e.toString());
    return null;
  }
}

/**
 * 苗字からAPIトークンを取得
 */
function getApiTokenByName(name) {
  try {
    // APIトークン管理シートを開く
    const tokenSheet = SpreadsheetApp.openById(TOKEN_SHEET_ID).getSheets()[0];
    
    // A列とB列のデータを取得（2行目から最終行まで）
    const lastRow = tokenSheet.getLastRow();
    if (lastRow < 2) {
      return null;
    }
    
    const data = tokenSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // 苗字に一致する行を検索
    for (let i = 0; i < data.length; i++) {
      const sheetName = data[i][0]; // A列
      const token = data[i][1]; // B列
      
      if (sheetName === name && token !== '' && token !== null) {
        return token;
      }
    }
    
    return null;
  } catch (e) {
    Logger.log('APIトークン取得エラー: ' + e.toString());
    return null;
  }
}

/**
 * Chatwork APIでメッセージを送信
 */
function sendToChatwork(message, apiToken, roomId) {
  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  
  const options = {
    'method': 'post',
    'headers': {
      'X-ChatWorkToken': apiToken
    },
    'payload': {
      'body': message
    },
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      return true;
    } else {
      Logger.log('エラー: ' + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log('例外エラー: ' + e.toString());
    return false;
  }
}