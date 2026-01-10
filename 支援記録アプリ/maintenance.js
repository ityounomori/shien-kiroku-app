// maintenance.gs

/**
 * システムのヘルスチェック（トークンリフレッシュ）と管理者への通知を行う。
 * 日次で実行されることを想定。
 */
function dailyHealthCheck() {
    const SCRIPT_NAME = 'dailyHealthCheck';
    const ADMIN_EMAIL = getAdminEmail();
    const SUBJECT = '【GAS通知】Dropbox連携 ヘルスチェック結果';
    let mailBody = `GAS Dropbox連携システムのヘルスチェック結果をご報告します。\n\n`;

    Logger.log(`${SCRIPT_NAME}: 処理を開始します。 (v35.5)`);

    try {
        // 1. トークンリフレッシュの試行
        handleTokenRefresh(); // dropbox.js の現行関数
        mailBody += `✅ トークンリフレッシュ: 成功（または不要）。Dropbox認証は正常です。\n`;
    } catch (e) {
        // 失敗した場合
        mailBody += `❌ トークンリフレッシュ: 失敗。\nエラー内容: ${e.message}\n\n`;
        mailBody += `システムが停止している可能性があります。リフレッシュトークンが失効していないか、手動でご確認ください。\n`;
        Logger.log(`${SCRIPT_NAME}: 致命的なトークンエラーを検知: ${e.message}`);

        // 管理者へ緊急メール送信
        if (ADMIN_EMAIL) {
            MailApp.sendEmail(ADMIN_EMAIL, SUBJECT, mailBody);
            Logger.log(`${SCRIPT_NAME}: 管理者にエラーメールを送信しました。`);
        }
        return; // 致命的なエラーなのでここで終了
    }

    // 成功した場合の通知
    mailBody += `\n現在、システムに異常はありません。`;
    if (ADMIN_EMAIL) {
        MailApp.sendEmail(ADMIN_EMAIL, SUBJECT, mailBody);
        Logger.log(`${SCRIPT_NAME}: 管理者に成功メールを送信しました。`);
    }
}

/**
 * 日次で実行する全てのメンテナンス作業を統合するメイン関数。
 * この関数を日次トリガーに設定します。
 */
function dailyMaintenance() {
    const SCRIPT_NAME = 'dailyMaintenance';
    Logger.log(`${SCRIPT_NAME}: 日次メンテナンスを開始します。`);

    // 1. 各事業所の記録アーカイブ処理
    try {
        Logger.log(`${SCRIPT_NAME}: アーカイブ移動処理を開始...`);
        archiveSupportRecordsTrigger();
        Logger.log(`${SCRIPT_NAME}: アーカイブ移動処理が完了しました。`);
    } catch (e) {
        Logger.log(`${SCRIPT_NAME}: アーカイブ処理エラー: ${e.message}`);
        Logger.log(`${SCRIPT_NAME}: スタックトレース: ${e.stack}`);
    }

    // 2. ログのクリーンアップ (365日保持)
    try {
        Logger.log(`${SCRIPT_NAME}: ログクリーンアップ処理を開始...`);
        cleanupLogsTrigger();
        Logger.log(`${SCRIPT_NAME}: クリーンアップ処理が完了しました。`);
    } catch (e) {
        Logger.log(`${SCRIPT_NAME}: クリーンアップ処理エラー: ${e.message}`);
    }

    // 2.5 ゴミ箱(支援記録)のクリーンアップ (30日保持) [Fix: Missing Call]
    try {
        Logger.log(`${SCRIPT_NAME}: 支援記録ゴミ箱のクリーンアップ処理を開始...`);
        cleanupTrashTrigger();
        Logger.log(`${SCRIPT_NAME}: 支援記録ゴミ箱のクリーンアップ処理が完了しました。`);
    } catch (e) {
        Logger.log(`${SCRIPT_NAME}: 支援記録ゴミ箱クリーンアップエラー: ${e.message}`);
    }

    // 3. Dropbox認証のヘルスチェックと通知を実行
    try {
        Logger.log(`${SCRIPT_NAME}: ヘルスチェック処理を開始...`);
        dailyHealthCheck();
        Logger.log(`${SCRIPT_NAME}: ヘルスチェック処理が完了しました。`);
    } catch (e) {
        Logger.log(`${SCRIPT_NAME}: ヘルスチェック処理エラー: ${e.message}`);
    }

    // 4. 古いアーカイブの削除（設定された保持年数を超えるもの）
    try {
        Logger.log(`${SCRIPT_NAME}: 古いアーカイブの削除処理を開始...`);
        cleanupOldArchives();
        Logger.log(`${SCRIPT_NAME}: 古いアーカイブの削除処理が完了しました。`);
    } catch (e) {
        Logger.log(`${SCRIPT_NAME}: アーカイブ消去エラー: ${e.message}`);
    }

    // 5. 事故報告ゴミ箱のクリーンアップ (30日保持)
    try {
        Logger.log(`${SCRIPT_NAME}: 事故報告ゴミ箱のクリーンアップ処理を開始...`);
        cleanupIncidentTrashTrigger();
        Logger.log(`${SCRIPT_NAME}: 事故報告ゴミ箱のクリーンアップ処理が完了しました。`);
    } catch (e) {
        Logger.log(`${SCRIPT_NAME}: 事故報告ゴミ箱クリーンアップエラー: ${e.message}`);
    }

    Logger.log(`${SCRIPT_NAME}: 日次メンテナンスを完了しました。`);
}

/**
 * 180日経過した支援記録をアーカイブシートへ移動する日次トリガー
 * 1000行バッチ処理、移動整合性検証、ロールバック、ログ記録を完備。
 */
function archiveSupportRecordsTrigger() {
    const SCRIPT_NAME = 'archiveSupportRecordsTrigger';
    const mappingSs = SpreadsheetApp.openById(getMasterFileId());
    const mappingSheet = mappingSs.getSheetByName('OfficeMapping');
    if (!mappingSheet) {
        Logger.log(`${SCRIPT_NAME}: OfficeMappingシートが見つかりません`);
        logEvent({ action: 'ARCHIVE_MOVE', status: 'ERROR', message: 'OfficeMappingシートが見つかりません' });
        return;
    }
    const mappingData = mappingSheet.getDataRange().getValues();
    const threshold = new Date();
    threshold.setDate(threshold.getDate() - 180);
    Logger.log(`${SCRIPT_NAME}: 基準日: ${Utilities.formatDate(threshold, 'Asia/Tokyo', 'yyyy-MM-dd')} (180日前)`);

    for (let i = 1; i < mappingData.length; i++) {
        const officeName = String(mappingData[i][0]);
        const recordFileId = String(mappingData[i][1]);
        if (!recordFileId) continue;

        Logger.log(`${SCRIPT_NAME}: [${officeName}] アーカイブ処理を開始...`);

        try {
            const ss = SpreadsheetApp.openById(recordFileId);
            const sheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);
            if (!sheet || sheet.getLastRow() < 2) {
                Logger.log(`${SCRIPT_NAME}: [${officeName}] データなし（スキップ）`);
                continue;
            }

            const data = sheet.getDataRange().getValues();
            const headers = data[0];
            const rowsToMove = [];
            const rowsToKeep = [headers];
            let invalidDateCount = 0;

            for (let j = 1; j < data.length; j++) {
                // COL_INDEX.DATE は config.js では 2 (1ベース) なので、0ベースでは data[j][1]
                const rowDate = new Date(data[j][1]); // 日付列（B列、インデックス1）
                if (isNaN(rowDate.getTime())) {
                    invalidDateCount++;
                    rowsToKeep.push(data[j]); // 無効な日付は保持
                    continue;
                }
                if (rowDate < threshold) rowsToMove.push(data[j]);
                else rowsToKeep.push(data[j]);
            }

            Logger.log(`${SCRIPT_NAME}: [${officeName}] 総データ数: ${data.length - 1}, 移動対象: ${rowsToMove.length}, 保持: ${rowsToKeep.length - 1}, 無効日付: ${invalidDateCount}`);

            if (rowsToMove.length > 0) {
                // バッチ処理（1000行単位での書き込み）
                const year = threshold.getFullYear();
                const archName = `記録_Archive_${year}`;
                let archSheet = ss.getSheetByName(archName) || ss.insertSheet(archName);
                if (archSheet.getLastRow() === 0) {
                    archSheet.appendRow(headers);
                }

                const currentOriginCount = sheet.getLastRow();
                const currentArchiveCount = archSheet.getLastRow();

                try {
                    // アーカイブへの追記
                    const batchSize = 1000;
                    for (let k = 0; k < rowsToMove.length; k += batchSize) {
                        const chunk = rowsToMove.slice(k, k + batchSize);
                        archSheet.getRange(archSheet.getLastRow() + 1, 1, chunk.length, headers.length).setValues(chunk);
                    }

                    // 元シートの更新（全削除後、保持分のみ再セット）
                    sheet.clearContents();
                    sheet.getRange(1, 1, rowsToKeep.length, headers.length).setValues(rowsToKeep);

                    // データ書き込み（移動完了）後に正式なテーブル変換を実行
                    try {
                        applyOfficialTable(recordFileId, archSheet);
                    } catch (tableError) {
                        Logger.log(`${SCRIPT_NAME}: [${officeName}] テーブル変換エラー: ${tableError.message}`);
                    }

                    // 検証：総行数が移動前後で一致するか (Source残 + Archive増 == 元の総数)
                    if (sheet.getLastRow() + (archSheet.getLastRow() - currentArchiveCount) !== currentOriginCount) {
                        // ロールバック（元に戻す）
                        sheet.clearContents().getRange(1, 1, data.length, headers.length).setValues(data);
                        archSheet.deleteRows(currentArchiveCount + 1, rowsToMove.length);
                        throw new Error('移動後の行数整合性チェックに失敗しました');
                    }

                    Logger.log(`${SCRIPT_NAME}: [${officeName}] アーカイブ移動完了: ${rowsToMove.length}件`);
                    logEvent({
                        action: 'ARCHIVE_MOVE', officeSelected: officeName, status: 'SUCCESS',
                        message: `${rowsToMove.length}件をアーカイブ(${archName})に移動、テーブルを更新しました`,
                        detail: { moved: rowsToMove.length, archive: archName }
                    });
                } catch (innerError) {
                    Logger.log(`${SCRIPT_NAME}: [${officeName}] 移動失敗: ${innerError.message}`);
                    logEvent({ action: 'ARCHIVE_MOVE', officeSelected: officeName, status: 'ERROR', message: `整合性エラー: ${innerError.message}` });
                }
            } else {
                Logger.log(`${SCRIPT_NAME}: [${officeName}] 移動対象データなし`);
            }
        } catch (e) {
            Logger.log(`${SCRIPT_NAME}: [${officeName}] エラー: ${e.message}`);
            Logger.log(`${SCRIPT_NAME}: [${officeName}] スタックトレース: ${e.stack}`);
            logEvent({ action: 'ARCHIVE_MOVE', status: 'ERROR', message: `事業所 ${officeName} アーカイブ処理失敗: ${e.message}` });
        }
    }
    Logger.log(`${SCRIPT_NAME}: 全事業所のアーカイブ処理が完了しました。`);
}

/**
 * ログの365日保守（日次トリガー）
 */
function cleanupLogsTrigger() {
    try {
        const ss = SpreadsheetApp.openById(getMasterFileId());
        const sheet = ss.getSheetByName('ログ');
        if (!sheet || sheet.getLastRow() < 2) return;
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const threshold = new Date();
        const rowsToKeep = [headers];
        let deleted = 0;

        for (let i = 1; i < data.length; i++) {
            const expiry = new Date(data[i][16]); // COL 17: 保持期限
            if (expiry > threshold) {
                rowsToKeep.push(data[i]);
            } else {
                deleted++;
            }
        }

        if (deleted > 0) {
            // 内容をクリアして、残す行だけを一括で書き戻す (deleteRowループより高速)
            sheet.clearContents();
            sheet.getRange(1, 1, rowsToKeep.length, headers[0] ? headers.length : rowsToKeep[0].length).setValues(rowsToKeep);
            logEvent({ action: 'LOG_CLEANUP', status: 'SUCCESS', message: `${deleted}件の期限切れログを消去（一括処理）` });
        }
    } catch (e) {
        console.error('Log cleanup error:', e);
    }
}

/**
 * ゴミ箱の30日保守（日次トリガー）
 * 各事業所の「ゴミ箱」シートから30日経過したデータを完全削除
 */
function cleanupTrashTrigger() {
    try {
        const retentionDays = 30;
        const threshold = new Date();
        threshold.setDate(threshold.getDate() - retentionDays);

        const mappingSs = SpreadsheetApp.openById(getMasterFileId());
        const mappingSheet = mappingSs.getSheetByName('OfficeMapping');
        if (!mappingSheet) return;

        const mappingData = mappingSheet.getDataRange().getValues();

        for (let i = 1; i < mappingData.length; i++) {
            const officeName = String(mappingData[i][0]);
            const recordFileId = String(mappingData[i][1]);
            if (!recordFileId) continue;

            try {
                const ss = SpreadsheetApp.openById(recordFileId);
                const sheet = ss.getSheetByName('ゴミ箱');
                if (!sheet || sheet.getLastRow() < 2) continue;

                const data = sheet.getDataRange().getValues();
                const headers = data[0];
                const rowsToKeep = [headers];
                let deletedCount = 0;

                // 2行目以降をチェック
                for (let r = 1; r < data.length; r++) {
                    // A列: 削除日時
                    const delDate = new Date(data[r][0]);
                    if (delDate > threshold) {
                        rowsToKeep.push(data[r]);
                    } else {
                        deletedCount++;
                    }
                }

                if (deletedCount > 0) {
                    sheet.clearContents();
                    // ヘッダー行のみの場合もあるため判定
                    if (rowsToKeep.length > 0) {
                        sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
                    }
                    logEvent({
                        action: 'TRASH_CLEANUP',
                        officeSelected: officeName,
                        status: 'SUCCESS',
                        message: `${officeName}: ${deletedCount}件の古いゴミ箱データを完全削除しました`
                    });
                }

            } catch (innerE) {
                console.error(`Trash cleanup failed for ${officeName}:`, innerE);
            }
        }

    } catch (e) {
        console.error('Trash cleanup trigger error:', e);
    }
}

/**
 * 事故報告のゴミ箱の30日保守（日次トリガー）
 * 事故報告ファイルの「incident_trash」シートから30日経過したデータを完全削除
 */
function cleanupIncidentTrashTrigger() {
    try {
        const retentionDays = 30;
        const threshold = new Date();
        threshold.setDate(threshold.getDate() - retentionDays);

        const mappingSs = SpreadsheetApp.openById(getMasterFileId());
        const mappingSheet = mappingSs.getSheetByName('OfficeMapping');
        if (!mappingSheet) return;

        const mappingData = mappingSheet.getDataRange().getValues();

        for (let i = 1; i < mappingData.length; i++) {
            const officeName = String(mappingData[i][0]);
            // Column C (Index 2) is Incident File ID
            const incidentFileId = String(mappingData[i][2]);

            if (!incidentFileId) continue;

            try {
                const ss = SpreadsheetApp.openById(incidentFileId);
                const trashName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_TRASH) ? SHEET_NAMES.INCIDENT_TRASH : 'incident_trash';
                const sheet = ss.getSheetByName(trashName);

                if (!sheet || sheet.getLastRow() < 2) continue;

                const data = sheet.getDataRange().getValues();
                const headers = data[0];
                const rowsToKeep = [headers];
                let deletedCount = 0;

                // 2行目以降をチェック
                for (let r = 1; r < data.length; r++) {
                    const delDate = new Date(data[r][0]);

                    if (isNaN(delDate.getTime())) {
                        rowsToKeep.push(data[r]);
                        continue;
                    }

                    if (delDate > threshold) {
                        rowsToKeep.push(data[r]);
                    } else {
                        deletedCount++;
                    }
                }

                if (deletedCount > 0) {
                    sheet.clearContents();
                    if (rowsToKeep.length > 0) {
                        sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
                    }
                    logEvent({
                        action: 'INCIDENT_TRASH_CLEANUP',
                        officeSelected: officeName,
                        status: 'SUCCESS',
                        message: `${officeName}: ${deletedCount}件の古い事故報告ゴミ箱データを完全削除しました`
                    });
                }
            } catch (innerE) {
                console.error(`Incident Trash cleanup failed for ${officeName}:`, innerE);
            }
        }

    } catch (e) {
        console.error('Incident Trash cleanup trigger error:', e);
    }
}

/**
 * Sheets API を使用してアーカイブシートを正式な「テーブル（Tableオブジェクト）」に変換する
 * (2025年4月アップデートの最新仕様に準拠)
 */
function applyOfficialTable(spreadsheetId, sheet) {
    const sheetId = sheet.getSheetId();
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // テーブルには少なくともヘッダー + 1行のデータが必要
    if (lastRow < 2) return;

    // --- 競合回避処理 (GAS) ---
    try {
        const filter = sheet.getFilter();
        if (filter) {
            filter.remove();
            SpreadsheetApp.flush(); // 即時反映
        }
        const bandings = sheet.getBandings();
        bandings.forEach(b => b.remove());
    } catch (cleanError) {
        Logger.log('Cleanup error (GAS): ' + cleanError.message);
    }

    // 1行目固定
    try {
        if (sheet.getFrozenRows() === 0) sheet.setFrozenRows(1);
    } catch (f) { /* NOP */ }

    const requests = [];

    // --- 競合回避処理 (Sheets API) ---
    // 既存の「テーブルオブジェクト」を検出して削除キューに入れる
    try {
        const ssMetadata = Sheets.Spreadsheets.get(spreadsheetId, {
            ranges: [sheetName],
            fields: "sheets(properties,tables)"
        });
        const currentSheetData = ssMetadata.sheets.find(s => s.properties.sheetId === sheetId);
        if (currentSheetData && currentSheetData.tables && currentSheetData.tables.length > 0) {
            currentSheetData.tables.forEach(t => {
                requests.push({ deleteTable: { tableId: t.tableId } });
            });
            Logger.log(`${currentSheetData.tables.length}個の既存テーブル定義を削除します`);
        }
    } catch (apiMetaError) {
        Logger.log('Metadata check failed (ignoring): ' + apiMetaError.message);
    }

    // 有効なテーブル名を作成
    const tableName = "Tbl_Arch_" + sheetId + "_" + new Date().getTime();

    // 正式なテーブル追加リクエスト
    requests.push({
        addTable: {
            table: {
                name: tableName,
                range: {
                    sheetId: sheetId,
                    startRowIndex: 0,
                    endRowIndex: lastRow,
                    startColumnIndex: 0,
                    endColumnIndex: lastCol
                }
            }
        }
    });

    try {
        Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
        Logger.log(`Successfully converted to Official Table: ${tableName}`);
    } catch (e) {
        // 重複エラー等の場合はログに留める
        if (e.message.indexOf('filter') !== -1 || e.message.indexOf('table') !== -1) {
            Logger.log('Note: Overlap detected. Falling back to basic filter.');
        } else {
            console.error('Official Table Conversion failed: ' + e.message);
        }

        // 最終フォールバック
        try {
            if (sheet.getFilter() === null) {
                sheet.getRange(1, 1, 1, lastCol).createFilter();
            }
        } catch (filterError) { /* NOP */ }
    }
}

/**
 * 保持期限を過ぎた古いアーカイブシートをすべての事業所ファイルから削除する
 */
function cleanupOldArchives() {
    const retentionYears = parseInt(getSettingValue('ARCHIVE_RETENTION_YEARS') || '5');
    const currentYear = new Date().getFullYear();
    const limitYear = currentYear - retentionYears;

    const mappingSs = SpreadsheetApp.openById(getMasterFileId());
    const mappingSheet = mappingSs.getSheetByName('OfficeMapping');
    if (!mappingSheet) return;

    const mappingData = mappingSheet.getDataRange().getValues();

    for (let i = 1; i < mappingData.length; i++) {
        const officeName = String(mappingData[i][0]);
        const recordFileId = String(mappingData[i][1]);
        if (!recordFileId) continue;

        try {
            const ss = SpreadsheetApp.openById(recordFileId);
            const sheets = ss.getSheets();

            for (const sheet of sheets) {
                const sheetName = sheet.getName();
                // 名前が "記録_Archive_YYYY" 形式かチェック
                const match = sheetName.match(/^記録_Archive_(\d{4})$/);
                if (match) {
                    const year = parseInt(match[1]);
                    if (year < limitYear) {
                        ss.deleteSheet(sheet);
                        logEvent({
                            action: 'ARCHIVE_CLEANUP', officeSelected: officeName, status: 'SUCCESS',
                            message: `${retentionYears}年以上の保持期限切れのためアーカイブシートを削除しました: ${sheetName}`
                        });
                    }
                }
            }
        } catch (e) {
            console.error(`Cleanup archives error for ${officeName}:`, e);
        }
    }
}

