// auth.gs

/**
 * 管理者PINコードを入力させて検証する
 * スクリプトプロパティ 'ADMIN_PIN_CODE' に保存されている値を使用します。
 * @returns {boolean} PINが正しければ true
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
  const correctPin = SCRIPT_PROPS.getProperty('ADMIN_PIN_CODE'); // スクリプトプロパティから取得

  if (enteredPin === correctPin) {
    return true;
  } else {
    ui.alert('エラー: 入力されたPINコードが正しくありません。');
    return false;
  }
}

/**
 * カスタムメニューから呼ばれる関数。PINチェックを行う。
 * PIN認証成功後、ダイアログを表示します。
 */
function showAuthDialogWithPinCheck() {
  // 処理の冒頭でPINチェックを実行
  if (!verifyAdminPin()) {
    return; // 認証失敗: verifyAdminPin() がアラートを表示
  }

  // PIN認証成功後、ダイアログを表示
  showAuthDialog();
}

/**
 * Dropbox認証設定ダイアログを表示する。
 */
function showAuthDialog() {
  const html = HtmlService.createHtmlOutputFromFile('auth_dialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dropboxリフレッシュトークン設定');
}

/**
 * Dropbox OAuth同意画面へのURLを生成する。
 * @returns {string} 認証URL
 */
function getDropboxAuthUrl() {
  // getDropboxClientId()は別途定義されていることを想定
  const appId = getDropboxClientId();

  // PowerShellで成功したredirect_uriなしのフローを再現
  const SCOPES_QUERY = '&scope=account_info.read files.metadata.read files.content.read files.content.write sharing.read sharing.write';

  // URLを組み立てる。redirect_uriを入れない
  const url = `https://www.dropbox.com/oauth2/authorize?response_type=code&client_id=${appId}${SCOPES_QUERY}&token_access_type=offline`;

  Logger.log("デバッグ認証URL (PowerShell再現): " + url);
  return url;
}

/**
 * HTMLダイアログから認証コードを受け取り、トークン交換を実行するラッパー関数。
 * @param {string} authCode HTMLから入力された認証コード
 * @returns {string} 処理結果メッセージ
 */
function handleAuthCodeSubmission(authCode) {
  // PINチェックはダイアログ表示前に行われているため、ここでは直接トークン交換を実行
  return exchangeCodeForToken(authCode);
}

/**
 * 認証コードをリフレッシュトークンとアクセストークンに交換する
 * @param {string} authCode ステップ1で取得した認証コード
 */
function exchangeCodeForToken(authCode) {
  const props = PropertiesService.getScriptProperties();
  // CLIENT_ID, CLIENT_SECRET は別途プロパティに設定されていることを想定
  const CLIENT_ID = props.getProperty('DROPBOX_CLIENT_ID');
  const CLIENT_SECRET = props.getProperty('DROPBOX_CLIENT_SECRET');

  const TOKEN_URL = 'https://api.dropboxapi.com/oauth2/token';

  // redirect_uri パラメータを削除したペイロード
  const payload = {
    'code': authCode,
    'grant_type': 'authorization_code',
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET
  };

  const options = {
    'method': 'post',
    'payload': payload, // GASはこれを x-www-form-urlencoded に自動変換
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
      // アクセストークンも保存するが、refreshTokenが本命
      const accessToken = json.access_token;

      // 取得したトークンをプロパティに保存
      setSettingValue('DROPBOX_REFRESH_TOKEN', refreshToken);
      setSettingValue('DROPBOX_ACCESS_TOKEN', accessToken);

      Logger.log('デバッグ: 取得したリフレッシュトークン: ' + refreshToken.substring(0, 10) + '...');
      return 'リフレッシュトークン、アクセストークンの取得と保存が完了しました。';
    } else {
      Logger.log('エラーコード: ' + responseCode);
      Logger.log('エラーレスポンス: ' + responseText);
      throw new Error(`トークン交換失敗 (コード: ${responseCode})。認証コードが無効か、クライアントID/シークレットに誤りがある可能性があります。`);
    }
  } catch (e) {
    Logger.log('UrlFetchApp エラー: ' + e.message);
    throw new Error('トークン交換処理中に通信エラーが発生しました: ' + e.message);
  }
}