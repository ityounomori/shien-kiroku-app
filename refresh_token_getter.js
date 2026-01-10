// refresh_token_getter.gs 

/**
 * リフレッシュトークンを使用して新しいアクセストークンを取得する関数
 * 🚨【最終修正点】リフレッシュトークンをシートではなく、スクリプトプロパティから読み込みます。
 * @returns {string} 新しいアクセストークン
 */
function getNewAccessToken() {
  // ----------------------------------------------------
  // 🚨【修正箇所】リフレッシュトークンをシートからではなく、スクリプトプロパティから直接読み込む
  // ----------------------------------------------------
  const refreshToken = getSettingValue('DROPBOX_REFRESH_TOKEN');
  // ----------------------------------------------------

  const appId = getDropboxClientId();
  const appSecret = getDropboxClientSecret();

  if (!refreshToken) {
    throw new Error('Dropboxリフレッシュトークンがスクリプトプロパティに設定されていません。');
  }
  if (!appId || !appSecret) {
    throw new Error('DropboxアプリIDまたはSecretがスクリプトプロパティに設定されていません。');
  }

  // APIコール URL (api.dropbox.com ではなく api.dropboxapi.com を使用して安定性を向上)
  const url = 'https://api.dropboxapi.com/oauth2/token';

  // ペイロード
  const payload = {
    grant_type: 'refresh_token',
    refresh_token: String(refreshToken),
    client_id: String(appId),
    client_secret: String(appSecret)
  };

  // オプション設定
  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  // APIコール実行
  Logger.log(`Dropboxリフレッシュ要求送信先: ${url}`);
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode === 200) {
    const json = JSON.parse(responseText);
    const newToken = json.access_token;
    if (newToken) {
      Logger.log('新しいアクセストークンを正常に取得しました。');
      return newToken;
    }
    throw new Error('API応答にアクセストークンが含まれていません。');
  } else {
    // 400 Bad Request (リフレッシュトークンが無効など)
    Logger.log(`トークンリフレッシュAPI失敗 (HTTP ${responseCode}). 詳細: ${responseText}`);
    throw new Error(`Dropboxリフレッシュ失敗 (コード: ${responseCode}, 詳細: ${responseText.substring(0, 100)})`);
  }
}