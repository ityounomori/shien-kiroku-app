
// 承認処理
function approveIncidentByOffice(officeName, rowId, whoami, pin) {
    // 再認証 & 権限チェック
    const auth = verifyUserByPin(officeName, whoami.name, pin);
    if (auth.success === false) throw new Error(auth.message || 'PIN認証に失敗しました。');
    if (auth.role !== 'manager') throw new Error('承認権限がありません(Managerのみ)。');

    const files = getFilesByOffice(officeName);
    const ss = SpreadsheetApp.openById(files.incidentFileId);
    const sheet = ss.getSheetByName(SHEET_NAMES.INCIDENT_SHEET);

    // 12列目(ステータス), 13列目(承認者), 14列目(日時) を更新
    // getRange(row, col) は 1-based index
    sheet.getRange(rowId, 12).setValue('承認済');
    sheet.getRange(rowId, 13).setValue(whoami.name);
    sheet.getRange(rowId, 14).setValue(new Date());

    logEvent({
        executor: whoami.name,
        officeSelected: officeName,
        action: 'INCIDENT_APPROVE',
        targetType: 'INCIDENT',
        targetId: String(rowId),
        status: 'SUCCESS'
    });

    return "承認しました。";
}

// 差戻し処理
function returnIncidentByOffice(officeName, rowId, whoami, pin, reason) {
    // 再認証 & 権限チェック
    const auth = verifyUserByPin(officeName, whoami.name, pin);
    if (auth.success === false) throw new Error(auth.message || 'PIN認証に失敗しました。');
    if (auth.role !== 'manager') throw new Error('差戻し権限がありません(Managerのみ)。');

    // 理由必須
    if (!reason || reason.trim() === "") throw new Error('差戻し理由を入力してください。');

    const files = getFilesByOffice(officeName);
    const ss = SpreadsheetApp.openById(files.incidentFileId);
    const sheet = ss.getSheetByName(SHEET_NAMES.INCIDENT_SHEET);

    sheet.getRange(rowId, 12).setValue('差戻し'); // ステータス
    sheet.getRange(rowId, 13).setValue('');     // 承認者クリア
    sheet.getRange(rowId, 14).setValue('');     // 日時クリア
    sheet.getRange(rowId, 15).setValue(reason); // 差戻し理由
    sheet.getRange(rowId, 16).setValue(new Date()); // 差戻し日時

    // 差戻し理由はログに残す (シートに列がないため)
    // もしシートに備考列があればそこへ追記するが、今回は14列固定仕様なのでログへ。
    logEvent({
        executor: whoami.name,
        officeSelected: officeName,
        action: 'INCIDENT_RETURN',
        targetType: 'INCIDENT',
        targetId: String(rowId),
        status: 'SUCCESS',
        detail: { reason: reason }
    });

    return "差戻しました。";
}

// -----------------------------------------------------------
// Master Data Fetchers (ByOffice)
// -----------------------------------------------------------

function getUserListDirectByOffice(officeName) {
    const sheet = SpreadsheetApp.openById(getMasterFileId()).getSheetByName(SHEET_NAMES.MASTER_USER_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

    // A列(名前)が実質的な文字を含まない行を除外
    const targetOffice = String(officeName || '').trim();
    return data
        .map(r => ({ name: String(r[0] || '').trim(), office: String(r[1] || '').trim() }))
        .filter(r => {
            if (!r.name || r.name.length === 0) return false; // Empty check only
            return r.office === '' || r.office === targetOffice;
        })
        .map(r => r.name); // Return simple string array
}

function getStaffListDirectByOffice(officeName) {
    const sheet = SpreadsheetApp.openById(getMasterFileId()).getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

    const targetOffice = String(officeName || '').trim();
    const staffList = data
        .map(r => ({ name: String(r[0] || '').trim(), auth: String(r[1] || '').trim() }))
        .filter(r => {
            // 名前が空、または空白文字のみの場合は除外
            if (r.name.length === 0 || !/[^\s\u3000]/.test(r.name)) return false;
            if (r.auth === '') return true; // 全許可
            return r.auth.split(',').map(s => s.trim()).includes(targetOffice);
        })
        .map(r => r.name);

    // 重複排除して返す
    return [...new Set(staffList)];
}

function getPhraseListDirectByOffice(officeName) {
    const sheet = SpreadsheetApp.openById(getMasterFileId()).getSheetByName(SHEET_NAMES.MASTER_PHRASES);
    if (!sheet || sheet.getLastRow() < 2) return [];
    // A:ID, B:区分, C:内容, D:事業所
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

    // C列(内容)が空でない かつ D列(事業所) が 空欄 or officeName と一致
    return data
        .filter(r => {
            const content = String(r[2] || '').trim();
            if (content === '') return false;
            const pOffice = String(r[3] || '').trim();
            return pOffice === '' || pOffice === officeName;
        })
        .map(r => ({
            区分: String(r[1]).trim(),
            内容: String(r[2]).trim()
        }));
}

// -----------------------------------------------------------
// Triggers: Archiving & Cleanup
// -----------------------------------------------------------



function generatePdfByOffice(officeName, userName, yearMonth, whoami) {
    try {
        const d = yearMonth.split('-');
        const y = parseInt(d[0]); const m = parseInt(d[1]);
        const s = new Date(y, m - 1, 1); const e = new Date(y, m, 0);
        const ssStr = Utilities.formatDate(s, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const esStr = Utilities.formatDate(e, Session.getScriptTimeZone(), 'yyyy-MM-dd');

        // 変更点: outputPdf は { base64: "...", fileName: "..." } を返すようになっています
        const result = outputPdf(userName, ssStr, esStr, officeName);

        // データがない場合のチェック
        if (!result || !result.base64) {
            return { success: false, message: "記録が見つかりません。" };
        }

        logEvent({ executor: whoami.name, officeSelected: officeName, action: 'PDF_GEN', targetType: 'USER', targetId: userName, status: 'SUCCESS' });

        // 変更点: すでにBase64化されているので、そのままクライアントに返します（再エンコード不要）
        return { success: true, data: result.base64, fileName: result.fileName };

    } catch (e) {
        return { success: false, message: e.message };
    }
}

// -----------------------------------------------------------
// 2. Logging System (18列固定)
// -----------------------------------------------------------

function logEvent(params) {
    try {
        const ss = SpreadsheetApp.openById(getMasterFileId());
        let sheet = ss.getSheetByName(SHEET_NAMES.LOG_HISTORY); // 'ログ'
        if (!sheet) return;

        const now = new Date();
        const expiry = new Date();
        expiry.setDate(now.getDate() + 365); // 1年保存

        const logId = Utilities.getUuid();

        // 18列のデータ作成
        const row = [
            now,                                        // 2. 記録日時
            'v35.5',                                    // 3. バージョン
            params.executor || 'System',                // 4. 実行者
            params.role || '',                          // 5. 権限
            params.officesAuth || '',                   // 6. 所属事業所
            params.officeSelected || '',                // 7. 選択事業所
            params.action || '',                        // 8. アクション
            params.targetType || '',                    // 9. 対象種別
            params.targetId || '',                      // 10. 対象ID
            params.targetDate || '',                    // 11. 対象日付
            params.status || 'SUCCESS',                 // 12. ステータス
            params.message || '',                       // 13. メッセージ
            params.detail ? JSON.stringify(params.detail) : '', // 14. 詳細
            '',                                         // 15. クライアント情報(空欄)
            params.requestId || '',                     // 16. 要求ID
            expiry,                                     // 17. 保持期限
            `${params.officeSelected}|${params.action}|${params.executor}` // 18. 検索インデックス
        ];

        sheet.appendRow(row);
    } catch (e) {
        console.error('Critical Logging Error:', e);
    }
}

/**
 * PDFダウンロード操作をログに記録する
 */
function logPdfDownload(officeName, userName, fileName, executor) {
    try {
        logEvent({
            executor: executor,
            officeSelected: officeName,
            action: 'PDF_DOWNLOAD',
            targetType: 'PDF',
            targetId: fileName,
            status: 'SUCCESS',
            message: 'ブラウザでのPDFダウンロードを実行',
            detail: {
                targetUser: userName
            }
        });
    } catch (e) {
        console.error('Failed to log PDF download:', e);
    }
}

/**
 * Dropboxブラウザ用: ファイル一覧取得
 */
function getDropboxFilesForWeb(path, userName, officeName) {
    if (!officeName) return { success: false, message: "事業所情報が不足しています。" };

    // 事業所ルート
    const officeRoot = `/アプリ/Googleフォーム支援記録/${officeName}`;

    // [Fix] 初期表示パスの決定
    // パス指定がない場合、「事業所/年/利用者」フォルダをデフォルトとする
    let targetPath = path;

    if (!targetPath) {
        if (userName) {
            const year = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
            targetPath = `${officeRoot}/${year}年/${userName}`;
        } else {
            targetPath = officeRoot;
        }
    }

    // 範囲外チェック: officeRoot で始まらないパスは強制リセット
    // これにより officeRoot 内部での移動は許可される
    if (!targetPath.startsWith(officeRoot)) {
        console.warn(`[Dropbox] Access Denied or Reset: ${targetPath} -> ${officeRoot}`);
        targetPath = officeRoot;
    }

    const result = listDropboxFiles(targetPath);

    // フロント制御用の rootPath は常に「事業所のルート」
    if (result.success && result.data) {
        result.data.rootPath = officeRoot;
    }
    return result;
}

/**
 * Dropboxブラウザ用: プレビュー用Base64取得
 */
function getDropboxFileForPreviewWeb(path) {
    const res = downloadDropboxFile(path); // dropbox.js
    if (res.success) {
        return { success: true, data: Utilities.base64Encode(res.blob.getBytes()) };
    }
    return res;
}

/**
 * ゴミ箱用: 削除済み記録の一覧取得
 */
function getTrashRecords(officeName) {
    try {
        if (!officeName) return [];
        console.log(`[Trash] Fetching for office: ${officeName}`);

        // 事業所ごとのファイルIDを取得
        const files = getFilesByOffice(officeName);
        if (!files || !files.recordFileId) {
            console.error(`[Trash] Record file not found for ${officeName}`);
            return [];
        }

        const ss = SpreadsheetApp.openById(files.recordFileId);
        const sheet = ss.getSheetByName('ゴミ箱');
        if (!sheet) {
            // シートがない場合は空リスト
            console.warn(`[Trash] 'ゴミ箱' sheet not found in ${files.recordFileId}`);
            return [];
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return []; // ヘッダーのみ

        // データ取得
        const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

        // マッピング処理
        return data.map((r, idx) => {
            try {
                // 全要素が空の場合はスキップ
                if (r.every(c => c === '')) return null;

                return {
                    trashRowIndex: idx + 2,
                    deleteDate: r[0] ? Utilities.formatDate(new Date(r[0]), "Asia/Tokyo", "yyyy/MM/dd HH:mm") : '(日付不明)',
                    recordDate: r[2] ? Utilities.formatDate(new Date(r[2]), "Asia/Tokyo", "yyyy/MM/dd HH:mm") : '', // 日付変換を追加
                    user: r[3] || '',
                    recorder: r[4] || '',
                    item: r[5] || '',
                    content: r[14] || ''
                };
            } catch (ex) {
                console.warn('[Trash] Row Parse Error:', ex);
                return null;
            }
        }).filter(item => item !== null)
            .sort((a, b) => new Date(b.deleteDate) - new Date(a.deleteDate)); // 新しい順

    } catch (e) {
        console.error('getTrashRecords Error:', e);
        return [];
    }
}

/**
 * ゴミ箱用: 記録の復元
 */
function restoreFromTrash(officeName, trashRowIndex) {
    try {
        if (!officeName) throw new Error("事業所名が指定されていません。");
        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.recordFileId);
        const trashSheet = ss.getSheetByName('ゴミ箱');
        const targetSheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);

        const values = trashSheet.getRange(trashRowIndex, 2, 1, 15).getValues()[0];
        targetSheet.appendRow(values);
        trashSheet.deleteRow(trashRowIndex);
        return "記録を復元しました。";
    } catch (e) {
        console.error('restoreFromTrash Error:', e);
        throw new Error("復元失敗: " + e.message);
    }
}

/**
 * インシデント削除 (ByOffice)
 */
/**
 * インシデント削除 (ゴミ箱へ移動)
 */
function deleteIncidentByOffice(officeName, rowId, whoami) {
    try {
        const files = getFilesByOffice(officeName);
        const sheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET) ? SHEET_NAMES.INCIDENT_SHEET : 'incidents';
        const trashName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_TRASH) ? SHEET_NAMES.INCIDENT_TRASH : 'incident_trash';

        const ss = SpreadsheetApp.openById(files.incidentFileId);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error('Incident sheet not found');

        // ゴミ箱シート取得（なければ作成）
        let trashSheet = ss.getSheetByName(trashName);
        if (!trashSheet) {
            trashSheet = ss.insertSheet(trashName);
            // ヘッダーコピー（初回のみ）
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            // ゴミ箱用管理列を追加（削除日時、削除者、元シートRowIDなどが必要だが、
            // 既存のゴミ箱ロジックに合わせて、元のデータを保持しつつ、
            // 削除日時(Col1), 削除者(Col2) を先頭に追加挿入する形式にするか、
            // あるいは末尾に追加するか。
            // 支援記録ゴミ箱は [DeleteDate, Deleter, ...OriginalCols] の形式ではない。
            // OriginalCols そのまま + DeleteDate?
            // 支援記録ゴミ箱ロジック(deleteRecordByOffice)を見ると：
            // targetSheet のデータを取得 -> trashSheet.appendRow([new Date(), whoami.name, ...values])

            // 定義済みヘッダー + "DeletedAt", "DeletedBy"
            trashSheet.appendRow(["DeletedAt", "DeletedBy", ...headers]);
        }

        // データの取得
        const lastCol = sheet.getLastColumn();
        // rowId is 1-based index
        const dataRange = sheet.getRange(Number(rowId), 1, 1, lastCol);
        const values = dataRange.getValues()[0];

        // ゴミ箱へ移動 (DeletedAt, DeletedBy, ...Data)
        const now = new Date();
        trashSheet.appendRow([now, whoami.name, ...values]);

        // 元データを削除
        sheet.deleteRow(Number(rowId));

        logEvent({
            executor: whoami.name, officeSelected: officeName,
            action: 'INCIDENT_TRASH', targetType: 'INCIDENT', targetId: String(rowId), status: 'SUCCESS'
        });
        return "ゴミ箱へ移動しました";
    } catch (e) {
        throw new Error('削除に失敗: ' + e.message);
    }
}

/**
 * 事故報告ゴミ箱リスト取得
 */
function getIncidentTrashRecords(officeName) {
    try {
        const files = getFilesByOffice(officeName);
        const trashName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_TRASH) ? SHEET_NAMES.INCIDENT_TRASH : 'incident_trash';
        const ss = SpreadsheetApp.openById(files.incidentFileId);
        const trashSheet = ss.getSheetByName(trashName);
        if (!trashSheet) return [];

        const lastRow = trashSheet.getLastRow();
        if (lastRow < 2) return [];

        const data = trashSheet.getRange(2, 1, lastRow - 1, trashSheet.getLastColumn()).getValues();

        // Data Format: [DeletedAt, DeletedBy, ID, CreatedAt, OccurDate, Recorder, User, Type, Place, Situation, ... ]
        // Index mapping:
        // 0: DeletedAt
        // 1: DeletedBy
        // 2: ID (UUID)
        // 3: CreatedAt
        // 4: OccurDate
        // 5: Recorder
        // 6: User
        // 7: Type
        // 8: Place
        // 9: Situation
        // 10: Cause
        // 11: Response
        // 12: Prevention
        // 13: Status
        // 14: Approver
        // 15: ApprovedAt
        // 16: ReturnReason
        // 17: ReturnedAt

        return data.map((r, idx) => {
            try {
                return {
                    trashRowIndex: idx + 2, // 1-based row index in Trash Sheet
                    deleteDate: r[0] ? Utilities.formatDate(new Date(r[0]), "Asia/Tokyo", "yyyy/MM/dd HH:mm") : '',
                    deleteBy: r[1],
                    occurDate: r[4] ? Utilities.formatDate(new Date(r[4]), "Asia/Tokyo", "yyyy/MM/dd HH:mm") : '',
                    recorder: r[5],
                    user: r[6],
                    type: r[7],
                    item: r[7], // Front-end uses 'item' for display compatibility
                    place: r[8],
                    content: r[9] // Situation as content summary
                };
            } catch (ex) {
                return null;
            }
        }).filter(i => i !== null)
            .sort((a, b) => new Date(b.deleteDate) - new Date(a.deleteDate));

    } catch (e) {
        console.error('getIncidentTrashRecords Error:', e);
        return [];
    }
}

/**
 * 事故報告の復元
 */
function restoreIncidentFromTrash(officeName, trashRowIndex) {
    try {
        const files = getFilesByOffice(officeName);
        const sheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET) ? SHEET_NAMES.INCIDENT_SHEET : 'incidents';
        const trashName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_TRASH) ? SHEET_NAMES.INCIDENT_TRASH : 'incident_trash';

        const ss = SpreadsheetApp.openById(files.incidentFileId);
        const sheet = ss.getSheetByName(sheetName);
        const trashSheet = ss.getSheetByName(trashName);

        if (!sheet || !trashSheet) throw new Error('Sheet not found');

        // Get from Trash
        const lastCol = trashSheet.getLastColumn();
        const dataRange = trashSheet.getRange(trashRowIndex, 1, 1, lastCol);
        const values = dataRange.getValues()[0];

        // Restore: Skip first 2 cols (DeletedAt, DeletedBy)
        // Format: [DeletedAt, DeletedBy, ...OriginalData]
        const originalData = values.slice(2);

        // Append to Active Sheet
        sheet.appendRow(originalData);

        // Remove from Trash
        trashSheet.deleteRow(trashRowIndex);

        return "報告書を復元しました";
    } catch (e) {
        throw new Error('復元失敗: ' + e.message);
    }
}

// [v52 Feature] Server-side Pagination for Incident History
function getIncidentHistory(officeName, limit = 50, offset = 0, filters = {}) {
    try {
        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.incidentFileId);
        // Use targetSheetName logic same as getIncidentsByOfficeV2
        const targetSheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET)
            ? SHEET_NAMES.INCIDENT_SHEET
            : 'incidents';
        const sheet = ss.getSheetByName(targetSheetName);

        if (!sheet) return { data: [], total: 0, debug: { error: 'Sheet not found' } };

        // --- [v55 Performance] Chunked Reverse Search ---
        const CHUNK_SIZE = 3000;
        const lastRow = sheet.getLastRow();

        // Normalized Filters
        const normFilters = {};
        if (filters) {
            if (filters.user) normFilters.user = (filters.user).toString().replace(/[\s\u3000]+/g, '').toLowerCase();
            if (filters.recorder) normFilters.recorder = (filters.recorder).toString().replace(/[\s\u3000]+/g, '').toLowerCase();
            if (filters.from) {
                const d = new Date(filters.from);
                d.setHours(0, 0, 0, 0);
                normFilters.from = d;
            }
            if (filters.to) {
                const d = new Date(filters.to);
                d.setHours(23, 59, 59, 999);
                normFilters.to = d;
            }
        }

        let results = [];
        let currentRow = lastRow;
        const targetCount = offset + limit + 1;

        while (currentRow >= 2 && results.length < targetCount) {
            const startRow = Math.max(2, currentRow - CHUNK_SIZE + 1);
            const numRows = currentRow - startRow + 1;
            if (numRows <= 0) break;

            const data = sheet.getRange(startRow, 1, numRows, 16).getValues();

            // Scan reverse (newest first)
            for (let i = data.length - 1; i >= 0; i--) {
                const r = data[i];

                // 1. Basic Checks
                if (!r[0]) continue; // ID required

                // Status Check (Strict Whitelist: Approved Only)
                // Fixes discrepancy where 'Pending' or other statuses might sneak in.
                const status = (r[11] || '未承認').toString().trim();
                // Allow '承認済' or '承認済み' (Handling variation)
                if (status !== '承認済' && status !== '承認済み') continue;

                // 2. Filter Checks
                const occurDate = r[2] instanceof Date ? r[2] : new Date(r[2]);

                if (normFilters.from && occurDate < normFilters.from) continue;
                if (normFilters.to && occurDate > normFilters.to) continue;
                if (filters.type && r[5] !== filters.type) continue;

                if (normFilters.user) {
                    const target = (String(r[4] || '')).replace(/[\s\u3000]+/g, '').toLowerCase();
                    if (target !== normFilters.user) continue;
                }
                if (normFilters.recorder) {
                    const target = (String(r[3] || '')).replace(/[\s\u3000]+/g, '').toLowerCase();
                    if (target !== normFilters.recorder) continue;
                }

                // 3. Mapping & Add
                results.push({
                    rowId: startRow + i,
                    id: String(r[0] || ''),
                    createdAt: r[1] instanceof Date ? Utilities.formatDate(r[1], "JST", "yyyy/MM/dd HH:mm") : String(r[1] || ''),
                    occurDate: occurDate instanceof Date ? Utilities.formatDate(occurDate, "JST", "yyyy/MM/dd HH:mm") : String(occurDate || ''),
                    recorder: String(r[3] || ''),
                    user: String(r[4] || ''),
                    userName: String(r[4] || ''),
                    type: String(r[5] || ''),
                    place: String(r[6] || ''),
                    situation: String(r[7] || ''),
                    cause: String(r[8] || ''),
                    response: String(r[9] || ''),
                    prevention: String(r[10] || ''),
                    status: status,
                    approver: String(r[12] || ''),
                    approvedAt: r[13] instanceof Date ? Utilities.formatDate(r[13], "JST", "yyyy/MM/dd HH:mm") : String(r[13] || ''),
                    returnReason: String(r[14] || ''),
                    returnedAt: r[15] instanceof Date ? Utilities.formatDate(r[15], "JST", "yyyy/MM/dd HH:mm") : String(r[15] || '')
                });

                if (results.length >= targetCount) break;
            }

            currentRow = startRow - 1;
        }

        const hasMore = results.length > offset + limit;
        const sliced = results.slice(offset, offset + limit);

        return {
            data: sliced,
            total: 9999, // Unknown total with this scan method
            hasMore: hasMore,
            debug: { scannedRows: lastRow - currentRow, found: results.length }
        };

    } catch (e) {
    }
}

function getIncidentCsvData(officeName, filters) {
    try {
        const files = getFilesByOffice(officeName);
        const targetSheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET)
            ? SHEET_NAMES.INCIDENT_SHEET
            : 'incidents';
        const ss = SpreadsheetApp.openById(files.incidentFileId);
        const sheet = ss.getSheetByName(targetSheetName);

        if (!sheet) return 'Error: Sheet not found';

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return 'ID,発生日時,記録者,利用者,種別,状況,原因,対応,再発防止策,ステータス';

        // Normalized Filters (Same as getIncidentHistory)
        const normFilters = {};
        if (filters) {
            if (filters.user) normFilters.user = (filters.user).toString().replace(/[\s\u3000]+/g, '').toLowerCase();
            if (filters.recorder) normFilters.recorder = (filters.recorder).toString().replace(/[\s\u3000]+/g, '').toLowerCase();
            if (filters.from) {
                const d = new Date(filters.from);
                d.setHours(0, 0, 0, 0);
                normFilters.from = d;
            }
            if (filters.to) {
                const d = new Date(filters.to);
                d.setHours(23, 59, 59, 999);
                normFilters.to = d;
            }
        }

        // Scan All Rows (Reverse Order)
        // Note: No iteration limit here. Might hit timeout if >100k rows.
        // Assuming manageable size for now.
        const CHUNK_SIZE = 5000;
        let currentRow = lastRow;
        let csvRows = [];

        // Header
        csvRows.push(['ID', '発生日時', '記録者', '利用者', '種別', '場所', '状況', '原因', '対応', '再発防止策', 'ステータス', '承認者', '承認日時'].join(','));

        while (currentRow >= 2) {
            const startRow = Math.max(2, currentRow - CHUNK_SIZE + 1);
            const numRows = currentRow - startRow + 1;
            if (numRows <= 0) break;

            const data = sheet.getRange(startRow, 1, numRows, 16).getValues();

            for (let i = data.length - 1; i >= 0; i--) {
                const r = data[i];
                if (!r[0]) continue;

                // Status Check (Strict Whitelist)
                const status = (r[11] || '未承認').toString().trim();
                if (status !== '承認済' && status !== '承認済み') continue;

                const occurDate = r[2] instanceof Date ? r[2] : new Date(r[2]);

                // Filter Checks
                if (normFilters.from && occurDate < normFilters.from) continue;
                if (normFilters.to && occurDate > normFilters.to) continue;
                if (filters && filters.type && r[5] !== filters.type) continue;

                if (normFilters.user) {
                    const target = (String(r[4] || '')).replace(/[\s\u3000]+/g, '').toLowerCase();
                    if (target !== normFilters.user) continue;
                }
                if (normFilters.recorder) {
                    const target = (String(r[3] || '')).replace(/[\s\u3000]+/g, '').toLowerCase();
                    if (target !== normFilters.recorder) continue;
                }

                // CSV Row Construction
                // Escape quotes
                const escape = (val) => `"${(String(val || '')).replace(/"/g, '""')}"`;

                const row = [
                    r[0], // ID
                    occurDate instanceof Date ? Utilities.formatDate(occurDate, "JST", "yyyy/MM/dd HH:mm") : r[2],
                    escape(r[3]), // Recorder
                    escape(r[4]), // User
                    escape(r[5]), // Type
                    escape(r[6]), // Place
                    escape(r[7]), // Situation
                    escape(r[8]), // Cause
                    escape(r[9]), // Response
                    escape(r[10]), // Prevention
                    status,
                    escape(r[12]), // Approver
                    r[13] instanceof Date ? Utilities.formatDate(r[13], "JST", "yyyy/MM/dd HH:mm") : r[13]
                ];
                csvRows.push(row.join(','));
            }
            currentRow = startRow - 1;
        }

        return csvRows.join('\r\n');

    } catch (e) {
        return 'Error: ' + e.toString();
    }
} // End of file
