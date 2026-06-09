// ─────────────────────────────────────────────────
//  設定管理
// ─────────────────────────────────────────────────

// 内部用：全 Script Properties を返す（PIN_HASH 含む）
function getSettingsInternal() {
  return PropertiesService.getScriptProperties().getProperties();
}

// クライアント向け：PIN_HASH を除いた設定を返す
function getSettings() {
  try {
    const props = getSettingsInternal();
    const safe = Object.assign({}, props);
    delete safe.PIN_HASH;
    // セッションキーも除外
    Object.keys(safe).forEach(k => { if (k.startsWith('SESSION_')) delete safe[k]; });
    return { success: true, settings: safe };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 設定を保存する
function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();

    // PIN 変更は別キーで受け取る
    if (settings.PIN_NEW_HASH) {
      const pinResult = changePin(settings.PIN_CURRENT_HASH, settings.PIN_NEW_HASH);
      if (!pinResult.success) return pinResult;
      delete settings.PIN_NEW_HASH;
      delete settings.PIN_CURRENT_HASH;
    }

    // チェックリスト項目は JSON 文字列化
    if (Array.isArray(settings.CHECKLIST_ITEMS)) {
      settings.CHECKLIST_ITEMS = JSON.stringify(settings.CHECKLIST_ITEMS);
    }

    if (settings.SPREADSHEET_ID) {
      settings.SPREADSHEET_ID = normalizeSpreadsheetId(settings.SPREADSHEET_ID);
    }

    Object.keys(settings).forEach(k => {
      if (k !== 'PIN_HASH' && k !== 'PIN_CURRENT_HASH') { // ハッシュは上記パスのみで書き換え可能
        props.setProperty(k, String(settings[k] == null ? '' : settings[k]));
      }
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// スプレッドシート接続テスト
function testConnection(spreadsheetId, sheetName) {
  try {
    const cleanId = normalizeSpreadsheetId(spreadsheetId);
    if (!cleanId) return { success: false, error: 'スプレッドシートIDが空です。' };

    const ss = SpreadsheetApp.openById(cleanId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      const sheets = ss.getSheets().map(s => s.getName());
      return {
        success: false,
        error: `シート「${sheetName}」が見つかりません。`,
        availableSheets: sheets
      };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const rowCount = Math.max(0, lastRow - 1);

    let headers = [];
    let preview = [];
    if (lastRow >= 1) {
      // Date オブジェクトは google.script.run でシリアライズできないため文字列化する
      headers = sheet.getRange(1, 1, 1, Math.min(lastCol, 20)).getValues()[0]
                  .map(v => String(v == null ? '' : v));
    }
    if (lastRow >= 2) {
      const rows = Math.min(3, lastRow - 1);
      preview = sheet.getRange(2, 1, rows, Math.min(lastCol, 10)).getValues()
                  .map(r => r.map(v => String(v == null ? '' : v)));
    }

    return {
      success: true,
      spreadsheetName: ss.getName(),
      rowCount: rowCount,
      columnCount: lastCol,
      headers: headers,
      preview: preview
    };
  } catch (e) {
    return { success: false, error: `接続エラー: ${e.message}` };
  }
}

// スプレッドシートのシート一覧を返す
function getSheetNames(spreadsheetId) {
  try {
    const cleanId = normalizeSpreadsheetId(spreadsheetId);
    const ss = SpreadsheetApp.openById(cleanId);
    return { success: true, sheets: ss.getSheets().map(s => s.getName()) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 指定列のユニーク値一覧を返す（フィルター値選択用）
function getFilterValues(spreadsheetId, sheetName, colNum, dataStartRow) {
  try {
    const cleanId = normalizeSpreadsheetId(spreadsheetId);
    if (!cleanId) return { success: false, error: 'スプレッドシートIDが空です。' };

    const ss    = SpreadsheetApp.openById(cleanId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: `シート「${sheetName}」が見つかりません。` };

    const col   = parseInt(colNum) || 7;
    const start = parseInt(dataStartRow) || 2;
    const lastRow = sheet.getLastRow();
    if (lastRow < start) return { success: true, values: [] };

    const numRows = lastRow - start + 1;
    const vals = sheet.getRange(start, col, numRows, 1).getValues().flat();

    const unique = [...new Set(vals.map(v => String(v || '').trim()).filter(Boolean))].sort();
    return { success: true, values: unique };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// URL または ID を正規化してIDのみを返す
function normalizeSpreadsheetId(urlOrId) {
  if (!urlOrId) return '';
  const m = String(urlOrId).match(/\/spreadsheets\/d\/([a-zA-Z0-9\-_]+)/);
  return m ? m[1] : String(urlOrId).trim();
}
