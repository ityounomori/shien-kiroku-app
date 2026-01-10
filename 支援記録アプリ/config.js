// config.gs

// =========================================================================================
// 1. 定数定義
// =========================================================================================

// ■ シート名定義
const SHEET_NAMES = {
    MASTER_USER_LIST: '利用者マスタ',
    MASTER_STAFF_LIST: '職員マスタ',
    MASTER_PHRASES: '定型文マスタ',
    LOG_HISTORY: 'ログ',
    PDF_HISTORY: 'PDF保存履歴',
    RECORD_INPUT: '記録',
    SETTING: '設定値',
    INCIDENT_SHEET: 'incidents',
    INCIDENT_TRASH: 'incident_trash', // 事故報告用ゴミ箱
    TRASH_RECORD: 'ゴミ箱'  // 支援記録用ゴミ箱
};

// ■ 支援記録 列インデックス (変更なし)
const COL_INDEX = {
    TIMESTAMP_AUTO: 1, DATE: 2, USER: 3, RECORDER: 4, ITEM: 5,
    DETAIL_1: 6, DETAIL_2: 7, V_TEMP: 8, V_BP_HIGH: 9, V_BP_LOW: 10,
    V_PULSE: 11, V_SPO2: 12, V_WEIGHT: 13, CONTENT: 14, SEARCH_INDEX: 15
};

// ■ 事故報告 ヘッダー定義 (14列固定) - 追加
const INCIDENT_COLS = [
    'ID (UUID)', '作成日時', '発生日時', '記録者', '対象利用者', '種別',
    '場所', '発生状況', '原因', '対応', '再発防止策',
    'ステータス', '承認者', '承認日時'
];

// =========================================================================================
// 2. ファイルID取得
// =========================================================================================

function getMasterFileId() {
    return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function getLogFileId() {
    return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function getIncidentFileId() {
    // Falls back to master if not found
    const id = PropertiesService.getScriptProperties().getProperty('INCIDENT_FILE_ID');
    return id || getMasterFileId();
}

// getPdfTemplateId は削除 (HTML生成方式に変更のため不要)

// =========================================================================================
// 3. 設定値シート操作
// =========================================================================================

function getSettingValue(key) {
    try {
        const val = PropertiesService.getScriptProperties().getProperty(key);
        return val ? String(val).trim() : '';
    } catch (e) {
        console.error(`[Config] getSettingValue error for ${key}: ${e.message}`);
        return '';
    }
}

function setSettingValue(key, value) {
    try {
        PropertiesService.getScriptProperties().setProperty(key, String(value));
    } catch (e) {
        console.error(`[Config] setSettingValue error for ${key}: ${e.message}`);
        throw new Error(`設定の保存に失敗しました: ${e.message}`);
    }
}

/**
 * 認証モード取得 (soft|strict_all|strict_incident)
 */
function getAuthMode() {
    try {
        const mode = getSettingValue('AUTH_MODE');
        return mode ? String(mode).trim() : 'soft';
    } catch (e) {
        return 'soft';
    }
}

/**
 * セッション有効期限（分）
 */
function getSessionMinutes() {
    try {
        const m = getSettingValue('SESSION_MINUTES');
        return m ? parseInt(m, 10) : 30;
    } catch (e) {
        return 30;
    }
}

/**
 * PINコード桁数
 */
function getPinLength() {
    try {
        // PropertiesService を優先
        const val = PropertiesService.getScriptProperties().getProperty('PIN_LENGTH');
        if (val) return parseInt(val, 10);

        // 取得できない場合は getSettingValue (プロパティ) 経由で取得
        const m = getSettingValue('PIN_LENGTH');
        return m ? parseInt(m, 10) : 4;
    } catch (e) {
        return 4; // v35 Fix: default 4 digits
    }
}

// =========================================================================================
// 4. Dropbox関連
// =========================================================================================

function getDropboxClientId() { return PropertiesService.getScriptProperties().getProperty('DROPBOX_CLIENT_ID'); }
function getDropboxClientSecret() { return PropertiesService.getScriptProperties().getProperty('DROPBOX_CLIENT_SECRET'); }
function getDropboxToken() { return getSettingValue('DROPBOX_ACCESS_TOKEN'); }
function getDropboxRefreshToken() { return getSettingValue('DROPBOX_REFRESH_TOKEN'); }

// =========================================================================================
// 5. ユーティリティ
// =========================================================================================

function getOfficeNameFromFileName() {
    try {
        const fileName = SpreadsheetApp.getActiveSpreadsheet().getName();
        const parts = fileName.split(/[_ ]/);
        if (parts.length >= 2 && parts[1]) {
            return String(parts[1]).trim();
        }
        return '';
    } catch (e) {
        Logger.log(`[Config] NameError: ${e.message}`);
        return '';
    }
}

function getAdminEmail() {
    return PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
}

function getRetentionDays() {
    const days = PropertiesService.getScriptProperties().getProperty('LOG_RETENTION_DAYS');
    return days ? parseInt(days, 10) : 180;
}

// =========================================================================================
// 6. マスタデータ取得
// =========================================================================================

function getUserListByOfficeAuto() {
    const officeName = getOfficeNameFromFileName() || '';
    const sheet = SpreadsheetApp.openById(getMasterFileId()).getSheetByName(SHEET_NAMES.MASTER_USER_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return values
        .filter(row => {
            const uName = String(row[0] || '').trim();
            const uOffice = String(row[1] || '').trim();
            if (uName === '') return false;
            return uOffice === '' || uOffice === officeName;
        })
        .map(row => ({ userName: String(row[0]).trim() }));
}

function getStaffListFromMaster() {
    const officeName = getOfficeNameFromFileName() || '';
    const sheet = SpreadsheetApp.openById(getMasterFileId()).getSheetByName(SHEET_NAMES.MASTER_STAFF_LIST);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const filtered = values.filter(row => {
        const name = String(row[0] || '').trim();
        const office = String(row[1] || '').trim();
        if (name === '') return false;
        return office === '' || office === officeName;
    });
    return filtered.map(row => String(row[0]).trim());
}

function getPhraseMaster() {
    const officeName = getOfficeNameFromFileName() || '';
    const ss = SpreadsheetApp.openById(getMasterFileId());
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_PHRASES);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const filtered = values.filter(row => {
        const content = String(row[2] || '').trim();
        const pOffice = String(row[3] || '').trim();
        if (content === '') return false;
        return pOffice === '' || pOffice === officeName;
    });
    return filtered.map(row => ({
        区分: String(row[1]),
        内容: String(row[2])
    }));
}