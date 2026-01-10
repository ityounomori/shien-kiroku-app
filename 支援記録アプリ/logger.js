// logger.gs

function writeLog(logObj) {
  // 古い形式の呼び出しを新しい logEvent に変換して転送
  logEvent({
    executor: logObj.担当者名,
    officeSelected: logObj.事業所名,
    action: 'LEGACY_LOG',
    targetType: 'USER',
    targetId: logObj.利用者名,
    status: logObj.結果 === '失敗' ? 'ERROR' : 'SUCCESS',
    message: logObj.操作内容,
    detail: logObj.詳細
  });
}

// PDF履歴などはそのまま残しても良いが、できれば logEvent に統合推奨
function writePDFHistory(historyObj) {
  // 既存の処理を維持、またはここも統合
  const logFileId = getLogFileId();
  const sheet = SpreadsheetApp.openById(logFileId).getSheetByName('PDF保存履歴');
  sheet.appendRow([
    historyObj.保存日時,
    historyObj.利用者名,
    historyObj.期間,
    historyObj.ファイル名,
    historyObj.Dropbox保存先,
    historyObj.結果,
    historyObj.備考
  ]);
}