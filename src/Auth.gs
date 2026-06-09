// ─────────────────────────────────────────────────
//  認証（PINコード + セッション管理）
// ─────────────────────────────────────────────────

// PIN を検証してセッショントークンを返す
function verifyPin(hashedPin) {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty('PIN_HASH');

    if (!stored) {
      return { success: false, error: 'PINが設定されていません。設定画面から設定してください。' };
    }
    if (stored !== hashedPin) {
      return { success: false, error: 'PINが正しくありません。' };
    }

    const token = Utilities.getUuid();
    const expiry = new Date().getTime() + (8 * 60 * 60 * 1000); // 8時間
    props.setProperty('SESSION_' + token, String(expiry));

    _cleanupSessions(props);

    return { success: true, sessionToken: token };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// セッショントークンの有効性を確認する（各 API 内部で呼ぶ）
function validateSession(token) {
  if (!token) return false;
  try {
    const props = PropertiesService.getScriptProperties();
    const expiryStr = props.getProperty('SESSION_' + token);
    if (!expiryStr) return false;

    if (new Date().getTime() > parseInt(expiryStr)) {
      props.deleteProperty('SESSION_' + token);
      return false;
    }
    return true;
  } catch (e) {
    return false;
  }
}

// PIN を変更する
function changePin(currentHashedPin, newHashedPin) {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty('PIN_HASH');

    // 既存 PIN がある場合は現在の PIN を確認
    if (stored && stored !== currentHashedPin) {
      return { success: false, error: '現在のPINが正しくありません。' };
    }

    props.setProperty('PIN_HASH', newHashedPin);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// 期限切れセッションを削除（内部用）
function _cleanupSessions(props) {
  const now = new Date().getTime();
  const all = props.getProperties();
  Object.keys(all).forEach(k => {
    if (k.startsWith('SESSION_') && now > parseInt(all[k])) {
      props.deleteProperty(k);
    }
  });
}
