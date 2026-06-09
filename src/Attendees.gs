// ─────────────────────────────────────────────────
//  参加者データ管理
// ─────────────────────────────────────────────────

// 参加者リストを全件取得
function getAttendees() {
  try {
    const cfg = getSettingsInternal();

    if (!cfg.SPREADSHEET_ID || !cfg.SHEET_NAME) {
      return {
        success: false,
        error: 'スプレッドシートが設定されていません。設定画面から設定してください。',
        needsSetup: true
      };
    }

    const ss = SpreadsheetApp.openById(normalizeSpreadsheetId(cfg.SPREADSHEET_ID));
    const sheet = ss.getSheetByName(cfg.SHEET_NAME);
    if (!sheet) {
      return { success: false, error: `シート「${cfg.SHEET_NAME}」が見つかりません。` };
    }

    const dataStart = parseInt(cfg.DATA_START_ROW) || 2;
    const lastRow   = sheet.getLastRow();
    const lastCol   = sheet.getLastColumn();

    if (lastRow < dataStart) {
      return _buildResult(cfg, []);
    }

    const numRows = lastRow - dataStart + 1;
    const data    = sheet.getRange(dataStart, 1, numRows, lastCol).getValues();

    const colName    = _col(cfg.COL_NAME,           2);
    const colEmail   = _col(cfg.COL_EMAIL,          3);
    const colPhone   = _col(cfg.COL_PHONE,          0);
    const colType    = _col(cfg.COL_TICKET_TYPE,    0);
    const colNotes   = _col(cfg.COL_NOTES,          0);
    const colStatus  = _col(cfg.COL_CHECKIN_STATUS, 10);
    const colTime    = _col(cfg.COL_CHECKIN_TIME,   11);
    const colStaff   = _col(cfg.COL_CHECKIN_STAFF,  12);

    // 日付・拠点フィルター（G列など）
    const colFilter  = _col(cfg.COL_FILTER,   7); // デフォルトG列
    const filterVal  = (cfg.FILTER_VALUE || '').trim();

    const clItems = _parseChecklist(cfg.CHECKLIST_ITEMS);

    const attendees = data.map((row, i) => {
      const name = row[colName] ? String(row[colName]).trim() : '';
      if (!name) return null;

      // フィルター値が設定されていれば一致しない行をスキップ
      if (filterVal && colFilter >= 0) {
        const cellVal = String(row[colFilter] || '').trim();
        if (cellVal !== filterVal) return null;
      }

      const statusVal = colStatus >= 0 ? String(row[colStatus] || '') : '';
      const status = statusVal === '受付済' ? 'checked_in'
                   : statusVal === '欠席'   ? 'absent'
                   : 'pending';

      const clValues = {};
      clItems.forEach(item => {
        const c = _col(item.col, 0);
        if (c >= 0) clValues[item.id] = !!(row[c] && row[c] !== '');
      });

      return {
        rowIndex:    dataStart + i,
        name:        name,
        email:       colEmail >= 0 ? String(row[colEmail] || '') : '',
        phone:       colPhone >= 0 ? String(row[colPhone] || '') : '',
        ticketType:  colType  >= 0 ? String(row[colType]  || '') : '',
        notes:       colNotes >= 0 ? String(row[colNotes] || '') : '',
        status:      status,
        checkInTime: colTime  >= 0 && row[colTime] ? String(row[colTime]) : null,
        staffName:   colStaff >= 0 ? String(row[colStaff] || '') : '',
        clValues:    clValues
      };
    }).filter(Boolean);

    return _buildResult(cfg, attendees);

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 受付処理（受付済 / 欠席 / 取消）
function checkIn(rowIndex, action, clValues, staffName, sessionToken) {
  try {
    const cfg = getSettingsInternal();

    if (cfg.PIN_ENABLED === 'true' && !validateSession(sessionToken)) {
      return { success: false, error: '認証が必要です。再度ログインしてください。', authRequired: true };
    }

    const ss    = SpreadsheetApp.openById(normalizeSpreadsheetId(cfg.SPREADSHEET_ID));
    const sheet = ss.getSheetByName(cfg.SHEET_NAME);

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
      return { success: false, error: '他の端末が処理中です。しばらくしてから再試行してください。' };
    }

    try {
      const colStatus = parseInt(cfg.COL_CHECKIN_STATUS) || 10;
      const colTime   = parseInt(cfg.COL_CHECKIN_TIME)   || 11;
      const colStaff  = parseInt(cfg.COL_CHECKIN_STAFF)  || 12;

      let statusVal = '';
      let timeVal   = '';
      let staffVal  = staffName || cfg.STAFF_NAME || '';

      if (action === 'check_in') {
        statusVal = '受付済';
        timeVal   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss');
      } else if (action === 'absent') {
        statusVal = '欠席';
        timeVal   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss');
      } else if (action === 'cancel') {
        statusVal = '';
        timeVal   = '';
        staffVal  = '';
      }

      sheet.getRange(rowIndex, colStatus).setValue(statusVal);
      sheet.getRange(rowIndex, colTime).setValue(timeVal);
      sheet.getRange(rowIndex, colStaff).setValue(staffVal);

      // チェックリスト書き込み
      if (clValues) {
        const clItems = _parseChecklist(cfg.CHECKLIST_ITEMS);
        clItems.forEach(item => {
          const c = parseInt(item.col);
          if (c > 0 && Object.prototype.hasOwnProperty.call(clValues, item.id)) {
            sheet.getRange(rowIndex, c).setValue(clValues[item.id] ? '✓' : '');
          }
        });
      }

      SpreadsheetApp.flush();

      return {
        success:       true,
        updatedStatus: action === 'check_in' ? 'checked_in' : action === 'absent' ? 'absent' : 'pending',
        checkInTime:   timeVal
      };

    } finally {
      lock.releaseLock();
    }

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// チェックリストのみ更新（受付済みカードの後から更新）
function updateChecklist(rowIndex, clValues, sessionToken) {
  try {
    const cfg = getSettingsInternal();

    if (cfg.PIN_ENABLED === 'true' && !validateSession(sessionToken)) {
      return { success: false, error: '認証が必要です。', authRequired: true };
    }

    const ss    = SpreadsheetApp.openById(normalizeSpreadsheetId(cfg.SPREADSHEET_ID));
    const sheet = ss.getSheetByName(cfg.SHEET_NAME);
    const clItems = _parseChecklist(cfg.CHECKLIST_ITEMS);

    clItems.forEach(item => {
      const c = parseInt(item.col);
      if (c > 0 && Object.prototype.hasOwnProperty.call(clValues, item.id)) {
        sheet.getRange(rowIndex, c).setValue(clValues[item.id] ? '✓' : '');
      }
    });

    SpreadsheetApp.flush();
    return { success: true };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 飛び込み参加者をスプレッドシートに追加する
function addAttendee(name, sessionToken) {
  try {
    const cfg = getSettingsInternal();

    if (cfg.PIN_ENABLED === 'true' && !validateSession(sessionToken)) {
      return { success: false, error: '認証が必要です。再度ログインしてください。', authRequired: true };
    }

    const cleanName = String(name || '').trim();
    if (!cleanName) return { success: false, error: '氏名を入力してください。' };

    if (!cfg.SPREADSHEET_ID || !cfg.SHEET_NAME) {
      return { success: false, error: 'スプレッドシートが設定されていません。' };
    }

    const ss    = SpreadsheetApp.openById(normalizeSpreadsheetId(cfg.SPREADSHEET_ID));
    const sheet = ss.getSheetByName(cfg.SHEET_NAME);
    if (!sheet) return { success: false, error: `シート「${cfg.SHEET_NAME}」が見つかりません。` };

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
      return { success: false, error: '他の端末が処理中です。しばらくしてから再試行してください。' };
    }

    try {
      const colName = parseInt(cfg.COL_NAME) || 2;
      const newRow  = sheet.getLastRow() + 1;

      sheet.getRange(newRow, colName).setValue(cleanName);

      // 当日フィルターが設定されていれば、その値も書いて当日対象に含める
      const colFilter = parseInt(cfg.COL_FILTER) || 0;
      const filterVal = (cfg.FILTER_VALUE || '').trim();
      if (colFilter > 0 && filterVal) {
        sheet.getRange(newRow, colFilter).setValue(filterVal);
      }

      SpreadsheetApp.flush();
      return { success: true, rowIndex: newRow };

    } finally {
      lock.releaseLock();
    }

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─── ヘルパー関数 ────────────────────────────────

// 列番号を 0-indexed で返す（未設定や0はマイナス1）
function _col(cfgVal, defaultNum) {
  const n = parseInt(cfgVal) || defaultNum;
  return n > 0 ? n - 1 : -1;
}

// チェックリスト JSON を安全にパース
function _parseChecklist(json) {
  try { return JSON.parse(json || '[]'); } catch (e) { return []; }
}

// 結果オブジェクトを組み立てる
function _buildResult(cfg, attendees) {
  return {
    success:       true,
    attendees:     attendees,
    eventName:     cfg.EVENT_NAME  || 'セミナー受付',
    eventDate:     cfg.EVENT_DATE  || '',
    clItems:       _parseChecklist(cfg.CHECKLIST_ITEMS),
    pinEnabled:    cfg.PIN_ENABLED === 'true'
  };
}
