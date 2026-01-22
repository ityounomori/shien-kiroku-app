// server.gs - v35 (ByOffice Unified & Archiving)
const SCRIPT_NAME = '支援記録システム';

function doGet(e) {
    return HtmlService.createHtmlOutputFromFile('record_form')
        .setTitle('支援記録・事故報告 統合システム')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// -----------------------------------------------------------
// ByOffice Mapping & File Resolvers
// -----------------------------------------------------------

/**
 * OfficeMappingシートから事業所名に対応するファイルIDを取得
 * A列: 事業所名 (完全一致) | B列: 記録ID | C列: 事故ID
 */
function getFilesByOffice(officeName) {
    if (!officeName) throw new Error('システムエラー: 事業所名が指定されていません');

    const ss = SpreadsheetApp.openById(getMasterFileId()); // config.jsの関数を使用
    const sheet = ss.getSheetByName('OfficeMapping');
    if (!sheet) throw new Error('初期設定エラー: OfficeMappingシートが見つかりません');

    const data = sheet.getDataRange().getValues();

    // 1行目はヘッダーとみなし、2行目から検索
    for (let i = 1; i < data.length; i++) {
        // String() と trim() で揺らぎを吸収
        if (String(data[i][0]).trim() === String(officeName).trim()) {
            const recordId = String(data[i][1] || '').trim();
            const incidentId = String(data[i][2] || '').trim();

            if (!recordId) throw new Error(`事業所「${officeName}」の記録用ファイルIDが設定されていません。`);

            // [Fix] 事故報告IDが未設定の場合は、事故報告機能が利用できないことを明示する
            if (!incidentId) {
                console.warn(`[Warning] Incident ID for ${officeName} is empty. Incident reports will not be available.`);
            }

            console.log(`[Debug] getFilesByOffice: Matched "${officeName}" at row ${i + 1}. incidentFileId: ${incidentId}`);
            return {
                recordFileId: recordId,
                incidentFileId: incidentId,
                officeRow: i + 1
            };
        }
    }
    console.error(`[Error] getFilesByOffice: Office "${officeName}" not found. Available: ${getOfficeList().join(', ')}`);
    throw new Error(`エラー: 事業所「${officeName}」はシステムに登録されていません。`);
}

/**
 * フロントエンドのプルダウン用リスト取得
 */
function getOfficeList() {
    try {
        const ss = SpreadsheetApp.openById(getMasterFileId());
        const sheet = ss.getSheetByName('OfficeMapping');
        if (!sheet) return [];

        const data = sheet.getDataRange().getValues();
        // A列のみ抽出（空欄とヘッダー「事業所名」を除外）
        return data
            .map(r => String(r[0]).trim())
            .filter(v => v !== '' && v !== '事業所名');
    } catch (e) {
        console.error('getOfficeList error:', e);
        return [];
    }
}

// -----------------------------------------------------------
// Logging System (Master "ログ" sheet)
// -----------------------------------------------------------

// logEvent は後半の18列版に統合済み

// -----------------------------------------------------------
// Auth & Data Fetching
// -----------------------------------------------------------

/**
 * アプリ起動時の初期データを返す（認証前）
 * 戻り値: { offices: string[] }
 */
function getInitialData() {
    try {
        return {
            offices: getOfficeList(),
            authMode: getAuthMode()
        };
    } catch (e) {
        console.error('getInitialData Critical Error:', e);
        // エラーでもクライアントが死なないように最低限のオブジェクトを返す
        // クライアント側で offices が空ならエラー表示するなどのハンドリングが必要だが
        // 少なくともロード画面でスタックすることは防げる
        return {
            offices: [],
            authMode: 'soft', // フォールバック
            error: e.toString()
        };
    }
}



/**
 * 認証成功後、選択された事業所のデータを一括取得する
 */
function getInitialDataByOffice(officeName) {
    try {
        // 1. ファイルID解決 (エラーならここで止まる)
        const files = getFilesByOffice(officeName);

        // 2. 利用者マスタ取得 (マスタシートから、その事業所の利用者のみフィルタ)
        const users = getUserListDirectByOffice(officeName);

        // 3. 職員リスト取得 (その事業所にアクセス権がある職員のみ)
        const staffs = getStaffListDirectByOffice(officeName);

        // 4. 定型文取得 (共通 + その事業所専用)
        const phrases = getPhraseListDirectByOffice(officeName);

        return {
            success: true,
            officeName: officeName,
            users: users,
            staffs: staffs,
            phrases: phrases,
            // クライアント側でファイルIDは不要だが、デバッグ用に返すことも可
            // fileIds: files 
        };

    } catch (e) {
        logEvent({
            officeSelected: officeName,
            action: 'INIT_DATA_FAIL',
            status: 'ERROR',
            message: e.message
        });
        throw new Error(`データ取得エラー: ${e.message}`);
    }
}

/**
 * Client-side script.run helper to get users for the currently active office on client
 */
function getUserListByOffice(officeName) {
    return getUserListDirectByOffice(officeName);
}

// (Duplicate definition removed)

/**
 * サインイン画面用: 指定事業所に所属する職員名のリストを返す
 */
function getStaffListByOffice(officeName) {
    return getStaffListDirectByOffice(officeName);
}

// (Duplicate definition removed)


/**
 * 職員マスタから指定事業所の職員を取得
 */
// (Duplicate removed - kept implementation at line ~891)


/**
 * PIN認証 & 事業所アクセス権チェック
 * @param {string} officeSelected ユーザーが選択した事業所名
 * @param {string} staffName 選択した職員名
 * @param {string} pin 入力されたPIN
 */
function verifyUserByPin(officeSelected, staffName, pin) {
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
                // 空欄の場合は「全事業所OK」とみなす、または「所属なし」とする（要件次第だが今回は全許可or指定のみ）
                // ここでは「空欄なら全許可」または「指定があればその事業所のみ」とします
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

                    logEvent({
                        executor: rowName,
                        role: rowRole,
                        officeSelected: officeSelected,
                        action: 'SIGNIN_SUCCESS',
                        targetType: 'AUTH',
                        status: 'SUCCESS',
                        message: 'サインイン成功'
                    });

                    return whoami;
                } else {
                    // PINは合っているが、事業所権限がない
                    logEvent({
                        executor: staffName,
                        officeSelected: officeSelected,
                        action: 'SIGNIN_DENIED',
                        targetType: 'AUTH',
                        status: 'ERROR',
                        message: '事業所権限なし'
                    });
                    throw new Error(`事業所「${officeSelected}」へのアクセス権限がありません。`);
                }
            }
        }

        // ループ終了しても見つからない -> 認証失敗
        logEvent({
            executor: staffName,
            officeSelected: officeSelected,
            action: 'PIN_FAIL',
            targetType: 'AUTH',
            status: 'ERROR',
            message: 'PINまたは名前の不一致'
        });
        throw new Error('認証に失敗しました。PINが正しくありません。');

    } catch (e) {
        console.error(e);
        // クライアントには詳細なエラーメッセージを返す
        return { success: false, message: e.message };
    }
}

function getStaffListRaw() {
    const ss = SpreadsheetApp.openById(getMasterFileId());
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1)
        .filter(r => String(r[0] || '').trim() !== '') // 名前が空の行を除外
        .map(r => ({
            name: String(r[0]).trim(),
            office: String(r[1] || '').split(',')[0].trim() || ''
        }));
}

function getUserListByOfficeAuto() {
    // This is a placeholder that might be called with context if we had a session on server
    // But since we pass office from client, we usually use getUserListDirectByOffice(office)
    // I will add a proxy that uses the last selected office if possible, or just expect officeName in args.
    // Actually, I'll just add the missing function that the client expects.
    return []; // Placeholder - the client should really call getUserListDirectByOffice(office)
}

// -----------------------------------------------------------
// Support Records CRUD (ByOffice)
// -----------------------------------------------------------

/**
 * 支援記録を取得する
 * @param {boolean} includeArchive trueならアーカイブも検索（重くなる可能性あり）
            */
/**
 * 支援記録を取得する (バッチ処理・自動アーカイブ対応)
 * @param {string} officeName 事業所名
            * @param {string} userName 対象利用者名
            * @param {Object} options {startDate, endDate, limit, continuationToken}
            */
function getRecordsByOffice(officeName, userName, options = {}) {
    try {
        const { startDate, endDate, limit = 500, continuationToken = null } = options;
        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.recordFileId);

        // 期間から対象となる年を特定
        const startDt = startDate ? new Date(startDate) : null;
        const endDt = endDate ? new Date(endDate) : null;

        // 全シートを取得し、メインシートおよびアーカイブシートを特定
        const allSheets = ss.getSheets();
        const sheetsToScan = [];

        // 1. メインシート「記録」
        const mainSheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);
        if (mainSheet) sheetsToScan.push(SHEET_NAMES.RECORD_INPUT);

        // 2. アーカイブシート「記録_Archive_YYYY」を名前降順（新しい順）で追加
        const archiveSheetNames = allSheets
            .map(s => s.getName())
            .filter(name => name.startsWith('記録_Archive_'))
            .sort((a, b) => b.localeCompare(a));

        sheetsToScan.push(...archiveSheetNames);

        console.log(`[getRecordsByOffice] Start. range: ${startDate} ~ ${endDate}, sheets: ${sheetsToScan.join(', ')}`);


        let records = [];
        let hasMore = false;
        let nextToken = null;

        // トークン解析 (前回の続きから)
        let startSheetIndex = 0;
        let startMatchOffset = 0;
        if (typeof options === 'boolean') {
            // 旧仕様 (includeArchive) への後方互換性。基本は options オブジェクトを想定。
            // ただし、この場合は詳細なバッチ制御はできない。
            return getRecordsByOfficeOld(officeName, userName, options);
        }

        if (continuationToken) {
            const parts = continuationToken.split(':');
            const savedSheetName = parts[0];
            startMatchOffset = parseInt(parts[1], 10);
            startSheetIndex = sheetsToScan.indexOf(savedSheetName);
            if (startSheetIndex === -1) startSheetIndex = 0;
        }

        for (let i = startSheetIndex; i < sheetsToScan.length; i++) {
            const sheetName = sheetsToScan[i];
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) continue;

            const remainingLimit = limit - records.length;
            const result = fetchRecordsFromSheetOptimized(sheet, userName, startDt, endDt, remainingLimit, startMatchOffset);

            records = records.concat(result.records);
            startMatchOffset = 0; // 次のシートからはオフセットなし

            if (result.hasMore) {
                hasMore = true;
                nextToken = `${sheetName}:${result.nextOffset}`;
                break;
            }
        }

        return {
            success: true,
            records: records,
            hasMore: hasMore,
            continuationToken: nextToken
        };

    } catch (e) {
        console.error(`getRecordsByOffice error: ${e.message}`);
        throw new Error(`記録取得エラー: ${e.message}`);
    }
}

/**
 * [互換用] 旧仕様の getRecordsByOffice
 */
function getRecordsByOfficeOld(officeName, userName, includeArchive) {
    const files = getFilesByOffice(officeName);
    const ss = SpreadsheetApp.openById(files.recordFileId);
    let allRecords = [];
    const sheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);
    if (sheet) allRecords = allRecords.concat(processRecordSheetRaw(sheet, userName));
    if (includeArchive) {
        const sheets = ss.getSheets();
        const archiveSheets = sheets.filter(s => s.getName().startsWith('記録_Archive_'));
        archiveSheets.forEach(ash => { allRecords = allRecords.concat(processRecordSheetRaw(ash, userName)); });
    }
    return allRecords.sort((a, b) => new Date(b['日付']) - new Date(a['日付']));
}

/**
 * シートの内容を高速にスキャンし、必要な範囲のレコードのみを取得する
 */
function fetchRecordsFromSheetOptimized(sheet, userName, startDate, endDate, limit, skip = 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { records: [], hasMore: false };

    // パフォーマンス向上のため、シート全体を一度に読み込む（row-by-rowのAPI呼び出しを回避）
    // 15列分 (A~O) を取得
    const allData = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    const matchedRows = []; // メモリ節約のため、マッチした行データそのものを保持

    const startTs = startDate ? startDate.getTime() : 0;
    const endTs = endDate ? endDate.getTime() : Infinity;
    const sheetName = sheet.getName();

    // 最新から取得するため、逆順(下から上)でスキャン
    for (let i = allData.length - 1; i >= 0; i--) {
        const row = allData[i];
        const dateVal = row[COL_INDEX.DATE - 1];
        const userVal = String(row[COL_INDEX.USER - 1] || '').trim();

        // 利用者チェックを先に行う（高速化）
        if (userName && userVal !== userName) continue;

        // 日付チェック
        const dt = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
        if (isNaN(dt.getTime())) continue; // 無効な日付は無視

        const ts = dt.getTime();
        if (ts < startTs || ts > endTs) continue;

        // この時点では重い変換をせず、行データと行番号だけを保持
        matchedRows.push({ row, rowNum: i + 2 });
    }

    console.log(`[fetchRecordsFromSheetOptimized] Sheet: ${sheetName}, total matched: ${matchedRows.length}`);

    // バッチ切り出し
    const pagedRows = matchedRows.slice(skip, skip + limit);
    const hasMore = matchedRows.length > (skip + limit);

    // 返却分だけ重い変換（日付フォーマット等）を行う
    const records = pagedRows.map(item => transformRowToRecord(item.row, item.rowNum, sheetName));

    return {
        records: records,
        hasMore: hasMore,
        nextOffset: skip + limit
    };
}

/**
 * 支援記録の1行(配列)をオブジェクト形式に変換するヘルパー
 */
function transformRowToRecord(row, rowNumber, sheetName) {
    const dateVal = row[COL_INDEX.DATE - 1];
    const dateStr = (dateVal instanceof Date)
        ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")
        : String(dateVal);

    const item = String(row[COL_INDEX.ITEM - 1] || '');
    let detailDisplay = "";
    const d1 = row[COL_INDEX.DETAIL_1 - 1];
    const d2 = row[COL_INDEX.DETAIL_2 - 1];

    if (item === '排泄' || item === '服薬') detailDisplay = d1;
    else if (item === '食事') detailDisplay = `摂取:${d1}% / 水分:${d2}ml`;
    else if (item === 'バイタル') {
        const v = [];
        if (row[COL_INDEX.V_TEMP - 1]) v.push(`熱:${row[COL_INDEX.V_TEMP - 1]}`);
        if (row[COL_INDEX.V_BP_HIGH - 1]) v.push(`BP:${row[COL_INDEX.V_BP_HIGH - 1]}/${row[COL_INDEX.V_BP_LOW - 1]}`);
        if (row[COL_INDEX.V_PULSE - 1]) v.push(`脈:${row[COL_INDEX.V_PULSE - 1]}`);
        if (row[COL_INDEX.V_SPO2 - 1]) v.push(`SpO2:${row[COL_INDEX.V_SPO2 - 1]}`);
        if (row[COL_INDEX.V_WEIGHT - 1]) v.push(`重:${row[COL_INDEX.V_WEIGHT - 1]}`);
        detailDisplay = v.join(', ');
    } else {
        detailDisplay = d1;
    }

    return {
        rowNumber: rowNumber,
        sheetName: sheetName,
        '日付': dateStr,
        '項目': item,
        '詳細': detailDisplay,
        '経過内容・様子': String(row[COL_INDEX.CONTENT - 1] || ''),
        '記録者': String(row[COL_INDEX.RECORDER - 1] || ''),
        'raw': {
            detail1: d1, detail2: d2,
            temp: row[COL_INDEX.V_TEMP - 1],
            bph: row[COL_INDEX.V_BP_HIGH - 1], bpl: row[COL_INDEX.V_BP_LOW - 1],
            pulse: row[COL_INDEX.V_PULSE - 1], spo2: row[COL_INDEX.V_SPO2 - 1],
            weight: row[COL_INDEX.V_WEIGHT - 1]
        }
    };
}

/**
 * シートからデータを抽出するヘルパー関数 (レガシー)
 */
function processRecordSheetRaw(sheet, targetUserName) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    const records = [];
    const sheetName = sheet.getName();
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const user = String(row[COL_INDEX.USER - 1] || '');
        if (targetUserName && user !== targetUserName) continue;
        records.push(transformRowToRecord(row, i + 2, sheetName));
    }
    return records;
}

/**
 * 新規記録の追加 (ByOffice)
 */
function addRecordByOffice(officeName, recordData, targetUsers, whoami) {
    const files = getFilesByOffice(officeName);
    const ss = SpreadsheetApp.openById(files.recordFileId);
    const sheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);

    if (!targetUsers || targetUsers.length === 0) throw new Error("利用者が選択されていません");

    // targetUsers配列の人数分ループして追加
    targetUsers.forEach(user => {
        const newRow = new Array(15).fill(""); // 15列確保

        const now = new Date();
        const inputDate = new Date(recordData.日付);

        newRow[COL_INDEX.TIMESTAMP_AUTO - 1] = now;
        newRow[COL_INDEX.DATE - 1] = inputDate;
        newRow[COL_INDEX.USER - 1] = user;
        newRow[COL_INDEX.RECORDER - 1] = recordData.記録者;
        newRow[COL_INDEX.ITEM - 1] = recordData.項目;
        newRow[COL_INDEX.CONTENT - 1] = recordData['経過内容・様子'];

        // 詳細データ (raw) の展開
        const rd = recordData.raw || {};
        newRow[COL_INDEX.DETAIL_1 - 1] = rd.detail1 || '';
        newRow[COL_INDEX.DETAIL_2 - 1] = rd.detail2 || '';
        newRow[COL_INDEX.V_TEMP - 1] = rd.temp || '';
        newRow[COL_INDEX.V_BP_HIGH - 1] = rd.bph || '';
        newRow[COL_INDEX.V_BP_LOW - 1] = rd.bpl || '';
        newRow[COL_INDEX.V_PULSE - 1] = rd.pulse || '';
        newRow[COL_INDEX.V_SPO2 - 1] = rd.spo2 || '';
        newRow[COL_INDEX.V_WEIGHT - 1] = rd.weight || '';

        // 検索インデックス作成 (日付 氏名 項目 内容 詳細)
        newRow[COL_INDEX.SEARCH_INDEX - 1] = [
            Utilities.formatDate(inputDate, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
            user, recordData.項目, recordData['経過内容・様子'], rd.detail1
        ].join(' ');

        sheet.appendRow(newRow);
    });

    // ログ記録
    logEvent({
        executor: whoami.name,
        role: whoami.role,
        officeSelected: officeName,
        action: 'RECORD_SAVE',
        targetType: 'RECORD',
        targetId: targetUsers.join(','),
        status: 'SUCCESS',
        message: `${targetUsers.length}件の記録を保存`
    });

    return `${targetUsers.length}件の記録を保存しました。`;
}

/**
 * 記録の編集 (ByOffice)
 * 行番号を指定して上書き更新します。ログに RECORD_EDIT を記録します。
 */
function editRecordByOffice(officeName, rowNumber, recordData, whoami) {
    try {
        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.recordFileId);
        const sheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);

        // 行番号の妥当性チェック
        const lastRow = sheet.getLastRow();
        if (rowNumber < 2 || rowNumber > lastRow) {
            throw new Error("指定された行番号が無効です。既に削除された可能性があります。");
        }

        // 更新対象の範囲を取得 (1行15列)
        const range = sheet.getRange(rowNumber, 1, 1, 15);
        const currentValues = range.getValues()[0];

        // ユーザー名は変更不可（整合性維持のため）
        // 日付、項目、内容、記録者、詳細データ(raw) を更新
        const inputDate = new Date(recordData.日付);

        currentValues[COL_INDEX.DATE - 1] = inputDate;
        currentValues[COL_INDEX.ITEM - 1] = recordData.項目;
        currentValues[COL_INDEX.RECORDER - 1] = recordData.記録者;
        currentValues[COL_INDEX.CONTENT - 1] = recordData['経過内容・様子'];

        // 詳細データ(raw)の展開
        const rd = recordData.raw || {};
        currentValues[COL_INDEX.DETAIL_1 - 1] = rd.detail1 || '';
        currentValues[COL_INDEX.DETAIL_2 - 1] = rd.detail2 || '';
        currentValues[COL_INDEX.V_TEMP - 1] = rd.temp || '';
        currentValues[COL_INDEX.V_BP_HIGH - 1] = rd.bph || '';
        currentValues[COL_INDEX.V_BP_LOW - 1] = rd.bpl || '';
        currentValues[COL_INDEX.V_PULSE - 1] = rd.pulse || '';
        currentValues[COL_INDEX.V_SPO2 - 1] = rd.spo2 || '';
        currentValues[COL_INDEX.V_WEIGHT - 1] = rd.weight || '';

        // 検索インデックスの再生成
        currentValues[COL_INDEX.SEARCH_INDEX - 1] = [
            Utilities.formatDate(inputDate, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
            currentValues[COL_INDEX.USER - 1], // 既存のユーザー名
            recordData.項目,
            recordData['経過内容・様子'],
            rd.detail1
        ].join(' ');

        // シートに書き戻し
        range.setValues([currentValues]);

        // ログ記録
        logEvent({
            executor: whoami.name,
            role: whoami.role,
            officeSelected: officeName,
            action: 'RECORD_EDIT',
            targetType: 'RECORD',
            targetId: String(rowNumber), // 行番号をID代わりに使用
            status: 'SUCCESS',
            message: `行${rowNumber}の記録を更新`
        });

        return "記録を更新しました。";

    } catch (e) {
        logEvent({
            officeSelected: officeName,
            action: 'RECORD_EDIT_FAIL',
            status: 'ERROR',
            message: e.message
        });
        throw e;
    }
}

/**
 * 記録の削除 (ByOffice)
 * 該当行を削除し、「ゴミ箱」シートへ移動します。
 */
function deleteRecordByOffice(officeName, rowNumber, whoami) {
    try {
        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.recordFileId);
        const sheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);

        // 行番号チェック
        const lastRow = sheet.getLastRow();
        if (rowNumber < 2 || rowNumber > lastRow) {
            throw new Error("削除対象の行が見つかりません。");
        }

        // 削除するデータを取得
        const range = sheet.getRange(rowNumber, 1, 1, 15);
        const values = range.getValues()[0];

        // ゴミ箱シートへの退避処理
        let trashSheet = ss.getSheetByName('ゴミ箱');
        if (!trashSheet) {
            trashSheet = ss.insertSheet('ゴミ箱');
            // ヘッダー作成 (削除日時 + 元の列)
            trashSheet.appendRow(['削除日時', 'タイムスタンプ', '日付', '利用者名', '記録者', '項目', '詳細1', '詳細2', '体温', '血圧上', '血圧下', '脈拍', 'SPO2', '体重', '経過内容', 'Index']);
        }

        // 先頭に削除日時を追加してゴミ箱へ
        trashSheet.appendRow([new Date(), ...values]);

        // 元シートから削除
        sheet.deleteRow(rowNumber);

        // ログ記録
        logEvent({
            executor: whoami.name,
            role: whoami.role,
            officeSelected: officeName,
            action: 'RECORD_DELETE',
            targetType: 'RECORD',
            targetId: String(rowNumber),
            status: 'SUCCESS',
            message: `行${rowNumber}を削除・ゴミ箱へ移動`
        });

        return "記録を削除し、ゴミ箱へ移動しました。";

    } catch (e) {
        logEvent({
            officeSelected: officeName,
            action: 'RECORD_DELETE_FAIL',
            status: 'ERROR',
            message: e.message
        });
        throw e;
    }
}

// -----------------------------------------------------------
// Incident Management (ByOffice & 14-column)
// -----------------------------------------------------------

function addIncidentByOffice(officeName, data, whoami) {
    const files = getFilesByOffice(officeName);
    if (!files.incidentFileId) {
        throw new Error(`事業所「${officeName}」には事故報告用ファイルIDが設定されていません。マスタを確認してください。`);
    }
    const ss = SpreadsheetApp.openById(files.incidentFileId);

    // シート取得または作成
    let sheet = ss.getSheetByName(SHEET_NAMES.INCIDENT_SHEET);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.INCIDENT_SHEET);
        sheet.appendRow(INCIDENT_COLS); // ヘッダー追加
    } else {
        // 列数チェック
        if (sheet.getLastColumn() < 14) {
            // 厳密にエラーにするか、足りない列を足すか。今回はエラー推奨。
            throw new Error(`エラー: 事故報告シートの列数が不足しています(現在${sheet.getLastColumn()}列)。14列必要です。`);
        }
    }

    const id = Utilities.getUuid();
    // 14列データ作成
    const row = [
        id,                                 // 1. ID
        new Date(),                         // 2. 作成日時
        new Date(data.occurDate),           // 3. 発生日時
        data.recorder,                      // 4. 記録者 (whoami.name または 指定)
        data.user,                          // 5. 対象利用者
        data.type,                          // 6. 種別
        data.place,                         // 7. 場所
        data.situation,                     // 8. 発生状況
        data.cause,                         // 9. 原因
        data.response,                      // 10. 対応
        data.prevention,                    // 11. 再発防止策
        '未承認',                           // 12. ステータス (初期値)
        '',                                 // 13. 承認者
        '',                                 // 14. 承認日時
        '',                                 // 15. 差戻し理由
        ''                                  // 16. 差戻し日時
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush(); // 即座に反映を確実にする

    console.log(`[Debug] Incident saved by ${whoami.name} to SS: ${files.incidentFileId}, Row: ${sheet.getLastRow()}`);

    logEvent({
        executor: whoami.name,
        officeSelected: officeName,
        action: 'ADD_INCIDENT',
        targetType: 'incident',
        targetId: id,
        status: 'SUCCESS',
        detail: data
    });

    return `報告を保存しました(行:${sheet.getLastRow()} / ID末尾:${files.incidentFileId.slice(-6)})`;
}


// 統合された updateIncident (互換性維持)
function updateIncident(rowId, data, editorName, editorPin, officeName) {
    // 古い呼び出し(4引数)の場合の対応
    if (!officeName) {
        // 1. PIN認証だけで一旦ユーザーを特定し、所属事業所を取得
        // 第一引数を空にしてverifyUserByPinを呼ぶとエラーになる可能性があるため、ダミーで呼ぶか、専用ロジックが必要。
        // ここでは verifyUserByPin の仕様により officeSelected が必須。
        // 仕方がないので、MasterStaffListを全検索する verifyPinOnly を利用する。

        const auth = verifyPinOnly(editorName, editorPin); // この関数も実装が必要
        if (!auth || !auth.success) throw new Error('PIN認証に失敗しました(Legacy Update)');

        // 権限のある事業所のうち、どれか1つを特定する必要がある。
        // Auth情報から推測
        const permissibleOffices = auth.officesAuth ? auth.officesAuth.split(',') : [];
        if (permissibleOffices.length === 0) {
            // 全許可の場合は、データから探す...のは重いので、
            // エラーを投げるか、デフォルト事業所があればそれを使う。
            throw new Error('事業所が特定できません。クライアントを更新してください。');
        }
        officeName = permissibleOffices[0]; // 暫定: 最初の事業所を使用
    }

    if (!officeName) {
        throw new Error('officeName が指定されていません（事故報告の更新に必要）');
    }

    // PIN再認証で whoami を取得
    const auth = verifyUserByPin(officeName, editorName, editorPin);
    if (!auth || !auth.success) {
        throw new Error('PIN認証に失敗しました');
    }
    // 本処理へ
    return updateIncidentByOffice(officeName, rowId, data, auth);
}


function updateIncidentByOffice(officeName, rowId, data, whoami) {
    const files = getFilesByOffice(officeName);
    const sheet = SpreadsheetApp.openById(files.incidentFileId).getSheetByName('incidents');

    const range = sheet.getRange(rowId, 1, 1, 16);
    const current = range.getValues()[0];
    const originalRecorder = current[3];

    if (whoami.role !== 'manager' && whoami.name !== originalRecorder) {
        throw new Error('編集権限がありません（記録者本人または管理者が可能です）');
    }

    const row = [...current];
    row[2] = new Date(data.occurDate);
    row[4] = data.user;
    row[5] = data.type;
    row[6] = data.place;
    row[7] = data.situation;
    row[8] = data.cause;
    row[9] = data.response;
    row[10] = data.prevention;
    row[11] = '未承認';
    row[12] = '';
    row[13] = '';
    // 修正されたら差戻し理由と日時もクリア
    row[14] = '';
    row[15] = '';

    range.setValues([row]);

    logEvent({
        executor: whoami.name, officeSelected: officeName,
        action: 'INCIDENT_UPDATE', targetType: 'INCIDENT', targetId: String(rowId), status: 'SUCCESS'
    });
    // 更新後の最新データを返却して、クライアント側で即時反映できるようにする
    return getIncidentByIdV3(officeName, rowId);
}

// (Duplicate definition removed)


/**
 * IDによる単一インシデントの取得
 */
function getIncidentByIdV3(officeName, rowId) {
    try {
        console.log(`[Debug] getIncidentByIdV3 called. Office: ${officeName}, RowId: ${rowId} (${typeof rowId})`);
        const files = getFilesByOffice(officeName);
        const sheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET) ? SHEET_NAMES.INCIDENT_SHEET : 'incidents';
        const sheet = SpreadsheetApp.openById(files.incidentFileId).getSheetByName(sheetName);

        if (!sheet) {
            console.log(`[Debug] Sheet not found: ${sheetName} in File ${files.incidentFileId}`);
            return { _error: 'SheetNotFound', sheetName: sheetName, fileId: files.incidentFileId };
        }

        const data = sheet.getDataRange().getValues();
        console.log(`[Debug] Data length: ${data.length}, Requesting Index: ${rowId - 1}`);

        const r = data[rowId - 1]; // rowId is 1-indexed

        if (!r) {
            console.log(`[Debug] Row not found or undefined at index ${rowId - 1}`);
            return { _error: 'RowNotFound', dataLength: data.length, requestIndex: rowId - 1, rowId: rowId };
        }

        console.log(`[Debug] Row found. r[0]=${r[0]}, r[2]=${r[2]}`);

        let formattedOccur = "";
        try {
            const d = new Date(r[2]);
            if (!isNaN(d.getTime())) {
                formattedOccur = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
            } else {
                console.log(`[Debug] Invalid Date for occurDate: ${r[2]}`);
                formattedOccur = String(r[2]);
            }
        } catch (e) {
            console.log(`[Debug] Date formatting error: ${e.message}`);
            formattedOccur = "";
        }

        console.log(`[Debug] Returning object...`);

        return {
            _v: 3,
            rowId: rowId,
            id: String(r[0]),
            createdAt: r[1] ? String(r[1]) : "", // Avoid Object types if possible for debug
            occurDate: formattedOccur,
            recorder: r[3],
            user: r[4],
            userName: r[4], // Client-side compatibility
            status: r[11],
            type: r[5],
            place: r[6],
            situation: r[7],
            cause: r[8],
            response: r[9],
            prevention: r[10]
        };
    } catch (e) {
        console.error('[Error] getIncidentByIdV3 failed:', e);
        return { _error: 'Exception', msg: e.toString(), stack: e.stack };
    }
}

/**
 * PINのみの簡易検証
 */

function verifyPinOnly(staffName, pin) {
    try {
        const ss = SpreadsheetApp.openById(getMasterFileId());
        const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            const rowName = String(data[i][0]).trim();
            const rowPin = String(data[i][3] || '').trim();
            if (rowName === staffName && rowPin === pin) {
                const rowOffices = String(data[i][1] || '').split(',').map(s => s.trim()).filter(s => s);
                return {
                    success: true,
                    name: rowName,
                    role: String(data[i][2] || 'staff').trim(),
                    officesAuth: rowOffices.join(','), // CSV
                    officeSelected: '' // 特定不可
                };
            }
        }
        return { success: false, message: 'PIN認証失敗' };
    } catch (e) {
        console.error(e);
        return { success: false, message: e.message };
    }
}


function getIncidentsByOfficeV2(officeName) {
    try {
        const files = getFilesByOffice(officeName);
        if (!files.incidentFileId) {
            console.warn(`[Warning] getIncidentsByOffice: No incidentFileId for ${officeName}`);
            return { data: [], debug: { error: 'incidentFileId is empty' } };
        }

        const ss = SpreadsheetApp.openById(files.incidentFileId);

        // 【修正箇所1】config.jsの読み込み順序に左右されないよう、安全にシート名を取得する
        const targetSheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET)
            ? SHEET_NAMES.INCIDENT_SHEET
            : 'incidents'; // config.jsの定義

        const sheet = ss.getSheetByName(targetSheetName);

        if (!sheet) {
            return { data: [], debug: { error: 'Sheet not found: ' + targetSheetName } };
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
            return { data: [], debug: { lastRow: lastRow, info: 'No data rows found' } };
        }

        const data = sheet.getDataRange().getValues();
        // const rows = data.slice(1).filter(r => r[0]); 
        // ↑ これだと行番号がずれるため、ループで処理する

        const norm = s => (s || '').toString().replace(/[\s\u3000]+/g, '').toLowerCase();
        const mapped = [];

        // ヘッダー(i=0)はスキップ、i=1から開始
        for (let i = 1; i < data.length; i++) {
            const r = data[i];
            if (!r[0]) continue; // IDが空ならスキップ

            mapped.push({
                rowId: i + 1, // 実際の行番号 (data[0]は1行目 => data[i]はi+1行目)
                id: String(r[0] || ''),
                createdAt: r[1] instanceof Date ? Utilities.formatDate(r[1], "JST", "yyyy/MM/dd HH:mm") : String(r[1] || ''),
                occurDate: r[2] instanceof Date ? Utilities.formatDate(r[2], "JST", "yyyy/MM/dd HH:mm") : String(r[2] || ''),
                recorder: String(r[3] || ''),
                recorderNorm: norm(r[3]),
                user: String(r[4] || ''),
                userName: String(r[4] || ''),
                type: String(r[5] || ''),
                place: String(r[6] || ''),
                situation: String(r[7] || ''),
                cause: String(r[8] || ''),
                response: String(r[9] || ''),
                prevention: String(r[10] || ''),
                status: (r[11] || '未承認').toString().trim(),
                approver: String(r[12] || ''),
                approvedAt: r[13] instanceof Date ? Utilities.formatDate(r[13], "JST", "yyyy/MM/dd HH:mm") : String(r[13] || ''),
                returnReason: String(r[14] || ''), // 15列目: 差戻し理由
                returnedAt: r[15] instanceof Date ? Utilities.formatDate(r[15], "JST", "yyyy/MM/dd HH:mm") : String(r[15] || '') // 16列目: 差戻し日時
            });
        }

        // 新しい順に表示したい場合
        mapped.reverse();

        console.log("SERVER_MAPPED_COUNT: " + mapped.length);
        return {
            data: mapped,
            debug: {
                success: true,
                rowCount: data.length, // data is the raw array from sheet
                mappedCount: mapped.length,
                officeName: officeName,
                firstRecorder: mapped.length > 0 ? mapped[0].recorder : "none"
            }
        };

    } catch (e) {
        console.error('getIncidentsByOfficeV2 Error:', e);
        return {
            data: [],
            debug: { error: e.toString(), stack: e.stack }
        };
    }
}

/**
 * [v55] Optimized Pending List Fetcher
 * Server-side filtering & Pagination for "Pending" tab.
 * Mirrors the logic previously done in client-side 'loadIncidentPendingList'.
 */
function getPendingIncidentsByOffice(officeName, limit = 50, offset = 0, whoami) {
    const debugLogs = [];
    const log = (msg) => {
        console.log(msg);
        debugLogs.push(msg);
    };

    try {
        log(`[getPendingIncidentsByOffice] Start. Office: ${officeName}, Limit: ${limit}, Offset: ${offset}, User: ${whoami ? whoami.name : 'Unknown'}`);

        if (!officeName) throw new Error('Office not specified');

        // Validate limit/offset
        limit = Math.max(1, Math.min(limit, 100)); // Cap at 100
        offset = Math.max(0, offset);

        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.incidentFileId);

        // [Fix] Safe Sheet Name Resolution
        const sheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET) ? SHEET_NAMES.INCIDENT_SHEET : 'incidents';
        const sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            log(`[Warn] Sheet "${sheetName}" not found.`);
            return { data: [], hasMore: false, debugLogs: debugLogs };
        }

        // [v55.3 Performance] Column-Based Scan (Index & Fetch)
        // Scans ONLY Status (Col 12) and Recorder (Col 4) columns first.
        // Much faster than fetching all columns for 100k+ rows.

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { data: [], hasMore: false, debugLogs: debugLogs };

        // 1. Fetch Index Columns (Status & Recorder)
        // Range: Row 2 to LastRow
        const numRows = lastRow - 1;
        const statusValues = sheet.getRange(2, 12, numRows, 1).getValues(); // Col L
        const recorderValues = sheet.getRange(2, 4, numRows, 1).getValues(); // Col D

        const norm = s => (String(s || '')).replace(/[\s\u3000]+/g, '').toLowerCase();
        const myNameNorm = norm(whoami.name);
        const isManager = (whoami.role === 'manager');

        const matchedIndices = []; // Array of { rowIndex, id (we don't have ID yet, wait) } -> Just row indices
        // Actually we need ID for checking? No, ID is Col 1. We verify ID existence later or assume valid if Status exists.
        // Let's just track row indices.

        log(`[Index Scan] Start. Rows: ${numRows}, Manager: ${isManager}, User: ${whoami.name}`);

        // 2. In-Memory Filter (Reverse: Newest First)
        for (let i = numRows - 1; i >= 0; i--) {
            // Check Offset/Limit
            if (matchedIndices.length >= (offset + limit + 10)) { // Fetch a bit more to ensure enough for pagination
                // We fetch a few more than (offset + limit) to correctly determine 'hasMore'
                // without needing to scan the entire sheet if we hit the limit early.
                // The exact number (e.g., +10) is a heuristic to balance performance and accuracy.
                break;
            }

            const rawStatus = (statusValues[i][0] || '未承認').toString();
            const status = rawStatus.replace(/[\s\u3000]+/g, '');
            const recorder = recorderValues[i][0];

            let isMatch = false;

            if (isManager) {
                // Manager: sees "Return" items
                if (status === '差戻し' || status === '差戻' || status === '差し戻し') {
                    isMatch = true;
                }
            } else {
                // Staff: sees Own "Pending" or "Return"
                const rNorm = norm(recorder);
                if (rNorm === myNameNorm) {
                    if (status === '未承認' || status === '差戻し' || status === '差戻' || status === '差し戻し') {
                        isMatch = true;
                    }
                }
            }

            if (isMatch) {
                matchedIndices.push(i + 2); // Convert to 1-based Row Index
            }
        }

        log(`[Index Scan] Done. Matches found: ${matchedIndices.length}`);

        // 3. Fetch Full Data for Matches
        // Apply Offset & Limit
        const pageIndices = matchedIndices.slice(offset, offset + limit);
        const results = [];

        if (pageIndices.length > 0) {
            // Construct A1 Notations (e.g., "A100:P100")
            const ranges = pageIndices.map(r => `A${r}:P${r}`);
            const rangeList = sheet.getRangeList(ranges);
            const rangeValues = rangeList.getRanges().map(r => r.getValues()[0]);

            // Map to Objects
            for (let i = 0; i < rangeValues.length; i++) {
                const row = rangeValues[i];
                const rowIndex = pageIndices[i];

                // ID Check
                if (!row[0]) continue;

                // Re-normalize status for display if needed, but we trust the index
                const rawStatus = (row[11] || '未承認').toString();
                // We use the raw value from the full fetch for consistency

                results.push({
                    rowId: rowIndex,
                    id: row[0],
                    createdAt: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss") : '',
                    occurDate: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : '',
                    recorder: row[3],
                    user: row[4],
                    userName: row[4],
                    type: row[5],
                    place: row[6],
                    situation: row[7],
                    cause: row[8],
                    response: row[9],
                    prevention: row[10],
                    status: rawStatus, // Use the raw status from row
                    approver: row[12],
                    approvedAt: row[13] ? Utilities.formatDate(new Date(row[13]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : '',
                    returnReason: row[14],
                    returnedAt: row[15] ? Utilities.formatDate(new Date(row[15]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : ''
                });
            }
        }

        const hasMore = matchedIndices.length > (offset + limit);

        log(`[Fetch] Done. Returning ${results.length} items.`);

        return {
            data: results,
            hasMore: hasMore,
            debugLogs: debugLogs
        };

    } catch (e) {
        console.error('getPendingIncidentsByOffice Error:', e);
        return { data: [], hasMore: false, error: e.message, debugLogs: debugLogs };
    }
}

/**
 * [v55.4] Optimized Approval List Fetcher (Manager Only)
 * Scans ALL rows (Cols Status) efficiently to find 'Unapproved' items.
 * Replaces the slow getIncidentsByOfficeV2 for the Approval tab.
 */
function getApprovalIncidentsByOffice(officeName) {
    try {
        const files = getFilesByOffice(officeName);
        const ss = SpreadsheetApp.openById(files.incidentFileId);

        const targetSheetName = (typeof SHEET_NAMES !== 'undefined' && SHEET_NAMES.INCIDENT_SHEET)
            ? SHEET_NAMES.INCIDENT_SHEET
            : 'incidents';
        const sheet = ss.getSheetByName(targetSheetName);

        if (!sheet) return { data: [], debug: { error: 'Sheet not found' } };

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { data: [] };

        // 1. Column Scan (Status is Col 12)
        // Fetching only the relevant column for filtering
        const numRows = lastRow - 1;
        const statusValues = sheet.getRange(2, 12, numRows, 1).getValues(); // Col L

        const matchedIndices = [];

        // 2. In-Memory Filter (Scan All)
        for (let i = 0; i < numRows; i++) {
            const rawStatus = (statusValues[i][0] || '未承認').toString();
            // Simplify status check: purely "未承認"
            if (rawStatus.trim() === '未承認') {
                matchedIndices.push(i + 2); // 1-based row index
            }
        }

        // 3. Fetch Full Data for Matches
        // If matches are too many, this might be heavy, but usually "Unapproved" are few.
        // We fetch all matches to replicate original behavior (no pagination logic requested).
        const results = [];

        if (matchedIndices.length > 0) {
            // Optimization: If many matches are contiguous, we could group ranges, 
            // but getRangeList handles non-contiguous well enough for < 100 items.
            // If matches > 1000, we might want to cap it, but let's assume Managers keep this low.

            const ranges = matchedIndices.map(r => `A${r}:P${r}`);
            const rangeList = sheet.getRangeList(ranges);
            const rangeValues = rangeList.getRanges().map(r => r.getValues()[0]);

            for (let i = 0; i < rangeValues.length; i++) {
                const row = rangeValues[i];
                const rowIndex = matchedIndices[i]; // Correct row index

                if (!row[0]) continue; // Skip if no ID (just in case)

                results.push({
                    rowId: rowIndex,
                    id: String(row[0]),
                    createdAt: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : '',
                    occurDate: row[2] ? Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : '',
                    recorder: String(row[3]),
                    user: String(row[4]),
                    userName: String(row[4]),
                    type: String(row[5]),
                    place: String(row[6]),
                    situation: String(row[7]),
                    cause: String(row[8]),
                    response: String(row[9]),
                    prevention: String(row[10]),
                    status: (row[11] || '未承認').toString(),
                    approver: String(row[12]),
                    approvedAt: row[13] ? Utilities.formatDate(new Date(row[13]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : '',
                    returnReason: String(row[14]),
                    returnedAt: row[15] ? Utilities.formatDate(new Date(row[15]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm") : ''
                });
            }
        }

        // Sort by Date (Newest First) - similar to V2 behavior
        results.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

        console.log(`[getApprovalIncidents] Office: ${officeName}, Scanned: ${numRows}, Found: ${results.length}`);
        return { data: results };

    } catch (e) {
        console.error('getApprovalIncidentsByOffice Error:', e);
        return { data: [], debug: { error: e.toString() } };
    }
}


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

