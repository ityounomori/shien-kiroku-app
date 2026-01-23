/**
 * AuthService: Authentication & Master Data Logic
 * Phase 1: Modularization
 */
const AuthService = {

  /**
   * Verify user by PIN and Office permissions
   * @param {string} officeSelected
   * @param {string} staffName
   * @param {string} pin
   * @returns {Object} whoami object or throws Error
   */
  verifyPin: function (officeSelected, staffName, pin) {
    try {
      const ss = SpreadsheetApp.openById(getMasterFileId());
      const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
      const data = sheet.getDataRange().getValues();

      // 1行目はヘッダーなのでスキップ
      for (let i = 1; i < data.length; i++) {
        const rowName = String(data[i][0]).trim();
        const rowOffices = String(data[i][1] || '').split(',').map(s => s.trim());
        const rowRole = String(data[i][2] || 'staff').trim();
        const rowPin = String(data[i][3] || '').trim();

        // 名前とPINの一致確認
        if (rowName === staffName && rowPin === pin) {
          // 事業所アクセス権の確認
          // 空欄の場合は「全事業所OK」とみなす
          const isAllowed = rowOffices.length === 0 || rowOffices[0] === '' || rowOffices.includes(officeSelected);

          if (isAllowed) {
            // 認証成功
            const whoami = {
              success: true,
              name: rowName,
              role: rowRole,
              officeSelected: officeSelected,
              officesAuth: rowOffices.join(','),
              expiresAt: Date.now() + (60 * 60 * 1000) // 1時間有効
            };

            // ログは呼び出し元(Server)で行うか、ここでやるか。一旦ここで完結させるが、
            // ログ機能(logEvent)への依存が発生する。
            // 今回は server.js 側に logEvent があるため、ここではログ記録を行わず、結果だけ返す設計にする。
            // (分離を見越して、AuthServiceは純粋なロジックのみにする)
            return whoami;

          } else {
            // PINは合っているが、事業所権限がない
            throw new Error(`事業所「${officeSelected}」へのアクセス権限がありません。`);
          }
        }
      }
      // ループ終了しても見つからない -> 認証失敗
      throw new Error('認証に失敗しました。PINが正しくありません。');

    } catch (e) {
      console.error('AuthService.verifyPin error:', e);
      throw e; // 再スローして呼び出し元で処理
    }
  },

  /**
   * Get staff list for a specific office
   * @param {string} officeName
   * @returns {Array<{name:string, role:string}>}
   */
  getStaffList: function (officeName) {
    const ss = SpreadsheetApp.openById(getMasterFileId());
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    // A: Name, B: Office(comma sep), C: Role
    const staffs = [];
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][0]).trim();
      const offices = String(data[i][1] || '').split(',').map(s => s.trim());
      const role = String(data[i][2] || 'staff').trim();

      if (name && (offices.length === 0 || offices.includes('') || offices.includes(officeName))) {
        staffs.push({ name: name, role: role });
      }
    }
    return staffs;
  },

  /**
   * Get all staff list (Raw)
   * @returns {Array<{name:string, office:string}>}
   */
  getStaffListRaw: function () {
    const ss = SpreadsheetApp.openById(getMasterFileId());
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1)
      .filter(r => String(r[0] || '').trim() !== '')
      .map(r => ({
        name: String(r[0]).trim(),
        office: String(r[1] || '').split(',')[0].trim() || ''
      }));
  }
};

/**
 * 管理者PINコードを入力させて検証する (Legacy Wrapper for Menu)
 * This logic depends on UI interaction, keeping it slightly separate or wrapping it?
 * Keeping as is for now, but verifyAdminPin could also use a PropertyService helper if needed.
 */
function verifyAdminPin() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '管理者認証が必要です',
    '設定済みのPINコードを入力してください:',
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() !== ui.Button.OK) {
    return false;
  }

  const enteredPin = result.getResponseText();
  const SCRIPT_PROPS = PropertiesService.getScriptProperties();
  const correctPin = SCRIPT_PROPS.getProperty('ADMIN_PIN_CODE');

  if (enteredPin === correctPin) {
    return true;
  } else {
    ui.alert('エラー: 入力されたPINコードが正しくありません。');
    return false;
  }
}

/**
 * カスタムメニューから呼ばれる関数
 */
function showAuthDialogWithPinCheck() {
  if (!verifyAdminPin()) return;
  showAuthDialog();
}

/**
 * Dropbox認証設定ダイアログを表示
 */
function showAuthDialog() {
  const html = HtmlService.createHtmlOutputFromFile('auth_dialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dropboxリフレッシュトークン設定');
}

/**
 * Dropbox OAuth同意画面へのURLを生成
 */
function getDropboxAuthUrl() {
  const appId = getDropboxClientId();
  const SCOPES_QUERY = '&scope=account_info.read files.metadata.read files.content.read files.content.write sharing.read sharing.write';
  const url = `https://www.dropbox.com/oauth2/authorize?response_type=code&client_id=${appId}${SCOPES_QUERY}&token_access_type=offline`;
  Logger.log("デバッグ認証URL: " + url);
  return url;
}

/**
 * HTMLダイアログから認証コードを受け取り、トークン交換を実行
 */
function handleAuthCodeSubmission(authCode) {
  return exchangeCodeForToken(authCode);
}

/**
 * 認証コードをトークンに交換
 */
function exchangeCodeForToken(authCode) {
  const props = PropertiesService.getScriptProperties();
  const CLIENT_ID = props.getProperty('DROPBOX_CLIENT_ID');
  const CLIENT_SECRET = props.getProperty('DROPBOX_CLIENT_SECRET');
  const TOKEN_URL = 'https://api.dropboxapi.com/oauth2/token';

  const payload = {
    'code': authCode,
    'grant_type': 'authorization_code',
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET
  };

  const options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true,
  };

  Logger.log('Dropboxトークン交換APIを呼び出し中...');
  try {
    const response = UrlFetchApp.fetch(TOKEN_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      const refreshToken = json.refresh_token;
      const accessToken = json.access_token;

      setSettingValue('DROPBOX_REFRESH_TOKEN', refreshToken);
      setSettingValue('DROPBOX_ACCESS_TOKEN', accessToken);

      Logger.log('デバッグ: 取得したリフレッシュトークン: ' + refreshToken.substring(0, 10) + '...');
      return 'リフレッシュトークン、アクセストークンの取得と保存が完了しました。';
    } else {
      Logger.log('エラーコード: ' + responseCode);
      Logger.log('エラーレスポンス: ' + responseText);
      throw new Error(`トークン交換失敗 (コード: ${responseCode})。`);
    }
  } catch (e) {
    Logger.log('UrlFetchApp エラー: ' + e.message);
    throw new Error('トークン交換処理中に通信エラーが発生しました: ' + e.message);
  }
}