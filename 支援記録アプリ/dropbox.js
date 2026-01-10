// dropbox.gs - v33 (Fix: Download Payload Issue & Full Logic Restoration)

// --- ヘルパー関数 ---

function getDropboxToken() {
  const token = getSettingValue('DROPBOX_ACCESS_TOKEN');
  if (!token) throw new Error('Dropboxアクセストークンが未設定です');
  return token;
}

function getExpiresAt30Minutes() {
  const now = new Date();
  const future = new Date(now.getTime() + 30 * 60 * 1000);
  return Utilities.formatDate(future, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

// トークンリフレッシュ処理（新しいトークンを返す）
function handleTokenRefresh() {
  try {
    Logger.log('トークンリフレッシュを実行します...');
    const newToken = getNewAccessToken(); // refresh_token_getter.gs
    setSettingValue('DROPBOX_ACCESS_TOKEN', newToken);
    Logger.log('トークン更新完了。');
    Utilities.sleep(1000); // 安定化のための待機
    return newToken;
  } catch (e) {
    Logger.log(`致命的エラー: リフレッシュ失敗 ${e.message}`);
    throw new Error('Dropboxアクセストークンの自動更新に失敗しました。');
  }
}

// --- API呼び出し汎用関数 ---

function callDropboxApi(url, method, payload, token, isBinary = false) {
  const maxRetries = 2;
  let currentToken = token;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const headers = {
      Authorization: 'Bearer ' + currentToken
    };

    // ダウンロード以外はContent-Typeを設定
    if (!url.includes('content.dropboxapi.com/2/files/download')) {
      headers['Content-Type'] = isBinary ? 'application/octet-stream' : 'application/json';
    }

    if (isBinary) {
      headers['Dropbox-API-Arg'] = payload;
    }

    const options = {
      method: method,
      headers: headers,
      muteHttpExceptions: true
    };

    // ダウンロード時はpayloadをセットしない
    if (!url.includes('content.dropboxapi.com/2/files/download')) {
      if (!isBinary) {
        options.payload = payload;
      } else {
        options.payload = payload.data;
      }
    }

    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 200) {
        return response; // 成功
      } else if (code === 401) {
        if (attempt < maxRetries) {
          Logger.log('401エラー検知。トークンリフレッシュを試みます。');
          currentToken = handleTokenRefresh();
          continue;
        } else {
          throw new Error('認証エラー(401)。リフレッシュ後も失敗しました。');
        }
      } else if (code === 409) {
        return response;
      } else {
        Logger.log(`APIエラー: ${code}, Body: ${response.getContentText()}`);
        throw new Error(`APIエラー: ${code} ${response.getContentText()}`);
      }
    } catch (e) {
      if (attempt === maxRetries) throw e;
    }
  }
}

// --- 個別機能関数 ---

function createSingleFolder(path, token) {
  const url = 'https://api.dropboxapi.com/2/files/create_folder_v2';
  const payload = JSON.stringify({ path: path, autorename: false });
  const headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
  const response = UrlFetchApp.fetch(url, { method: 'post', headers, payload, muteHttpExceptions: true });
  return response.getResponseCode();
}

function ensureDropboxFolder(fullPath, initialToken) {
  const parts = fullPath.split('/').filter(p => p.length > 0);
  let currentPath = "";
  let token = initialToken;

  for (let i = 0; i < parts.length; i++) {
    currentPath += "/" + parts[i];
    let code = createSingleFolder(currentPath, token);

    if (code === 401) {
      token = handleTokenRefresh();
      code = createSingleFolder(currentPath, token);
    }

    if (code !== 200 && code !== 409) {
      throw new Error(`フォルダ作成失敗: ${currentPath} (Code: ${code})`);
    }
  }
  return 200;
}

function uploadToDropbox(pdfBlob, userName, yearFolder, folderPath, fileName, initialToken, officeName, executor) {
  const url = 'https://content.dropboxapi.com/2/files/upload';
  const fullPath = `${folderPath}/${fileName}`;
  const arg = JSON.stringify({ path: fullPath, mode: 'add', autorename: true });

  let token = initialToken;
  for (let i = 0; i < 2; i++) {
    const headers = {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/octet-stream',
      'Dropbox-API-Arg': arg
    };
    const res = UrlFetchApp.fetch(url, { method: 'post', headers, payload: pdfBlob.getBytes(), muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      // 共通ログ機能(logEvent)への一本化
      if (typeof logEvent === 'function') {
        logEvent({
          executor: executor,
          action: 'DROPBOX_SAVE',
          officeSelected: officeName,
          targetType: 'PDF',
          targetId: fileName,
          targetDate: yearFolder,
          status: 'SUCCESS',
          message: 'Dropboxへの保存に成功しました',
          detail: {
            targetUser: userName,
            path: folderPath
          }
        });
      }
      return 200;
    }
    if (res.getResponseCode() === 401) {
      token = handleTokenRefresh();
    } else {
      throw new Error("Upload Failed: " + res.getResponseCode() + " " + res.getContentText());
    }
  }
  throw new Error("アップロードに失敗しました");
}

// --- [v33 Fix] ファイルダウンロード関数 ---
function downloadDropboxFile(path) {
  const token = getDropboxToken();
  const url = 'https://content.dropboxapi.com/2/files/download';
  const arg = JSON.stringify({ path: path });

  let lastError = "";
  let currentToken = token;

  try {
    for (let i = 0; i < 2; i++) {
      const options = {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + currentToken,
          'Dropbox-API-Arg': arg
          // Content-Typeを一切指定しない（GASの自動付与を抑える）
        },
        // 空のBlobをPayloadに指定。これでGASはデフォルトのContent-Typeを付与せず、
        // Dropboxも「Bodyなし」として正しく受理する。
        payload: Utilities.newBlob(''),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 200) {
        return { success: true, blob: response.getBlob() };
      } else if (code === 401) {
        lastError = `401 Unauthorized: ${response.getContentText()}`;
        currentToken = handleTokenRefresh();
        continue;
      } else {
        lastError = `${code} ${response.getContentText()}`;
      }
    }
    return { success: false, message: `Download Failed: ${lastError}` };

  } catch (e) {
    return { success: false, message: `Exception: ${e.message}` };
  }
}

// --- ファイル一覧取得 ---
function listDropboxFiles(path) {
  const token = getDropboxToken();
  const url = 'https://api.dropboxapi.com/2/files/list_folder';
  // 修正: path が null/undefined の場合は空文字（ルート）とする
  let listPath = (path === '/' || !path) ? '' : path;
  if (listPath && listPath.length > 1 && listPath.endsWith('/')) {
    listPath = listPath.slice(0, -1);
  }

  const payload = JSON.stringify({
    path: listPath, recursive: false, include_media_info: false,
    include_deleted: false, include_has_explicit_shared_members: false
  });

  try {
    const response = callDropboxApi(url, 'post', payload, token);
    const json = JSON.parse(response.getContentText());

    if (json.error) {
      if (json.error['.tag'] === 'path' && json.error.path['.tag'] === 'not_found') {
        return { success: false, message: 'フォルダが見つかりません。', debugPath: listPath };
      }
      throw new Error(JSON.stringify(json.error));
    }

    const entries = json.entries;
    const folders = entries.filter(e => e['.tag'] === 'folder').map(e => ({
      name: e.name, path: e.path_display, type: 'folder'
    })).sort((a, b) => a.name.localeCompare(b.name));

    const files = entries.filter(e => e['.tag'] === 'file').map(e => ({
      name: e.name, path: e.path_display, type: 'file',
      size: (e.size / 1024).toFixed(1) + ' KB',
      updated: new Date(e.server_modified).toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })
    })).sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, data: { folders, files, currentPath: path || '/' } };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- 一時リンク取得 ---
function getDropboxTempLink(path) {
  const token = getDropboxToken();
  const url = 'https://api.dropboxapi.com/2/files/get_temporary_link';
  const payload = JSON.stringify({ path: path });

  try {
    const response = callDropboxApi(url, 'post', payload, token);
    const json = JSON.parse(response.getContentText());
    if (json.link) return { success: true, link: json.link };
    throw new Error('Link not found in response');
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- その他必須関数 ---
// --- 保存パス生成ロジックの修正 ---
function getDropboxSavePath(userName, officeName) {
  // 修正: ファイル名からではなく、引数で渡された officeName を使う
  if (!officeName) throw new Error("事業所名が指定されていません");

  const year = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  // パス: /アプリ/Googleフォーム支援記録/[事業所名]/[年]/[利用者名]
  return `/アプリ/Googleフォーム支援記録/${officeName}/${year}年/${userName}`;
}

// --- Webアプリからの保存エントリーポイント ---
function saveToDropboxFromWebapp(officeName, userName, base64Data, fileName, executor) {
  try {
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, MimeType.PDF, fileName);

    savePDFToDropbox(blob, userName, officeName, executor);

    return { success: true, message: 'Dropboxへの保存に成功しました' };
  } catch (e) {
    Logger.log('Dropbox保存エラー: ' + e.message);
    return { success: false, message: '保存失敗: ' + e.message };
  }
}

function getUserDefaultPath(userName) {
  try { return getDropboxSavePath(userName); } catch (e) { return ''; }
}

// writePDFHistory は廃止し logEvent に統合しました

function savePDFToDropbox(pdfBlob, userName, officeName, executor) {
  let token = getDropboxToken();
  const folderPath = getDropboxSavePath(userName, officeName);
  const yearFolder = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  const fileName = pdfBlob.getName();

  try {
    ensureDropboxFolder(folderPath, token);
  } catch (e) {
    token = handleTokenRefresh();
    ensureDropboxFolder(folderPath, token);
  }
  uploadToDropbox(pdfBlob, userName, yearFolder, folderPath, fileName, token, officeName, executor);
}