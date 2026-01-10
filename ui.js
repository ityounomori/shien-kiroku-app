// ui.gs

/**
 * スプレッドシートのカスタムメニューを作成する。
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();

    // 管理者向けのメニューのみ残す
    ui.createMenu('🔑 管理者メニュー')
        .addItem('Dropbox設定 (要PIN)', 'showAuthDialogWithPinCheck') // auth.js
        .addItem('定型文マスタ初期化 (要PIN)', 'showPhraseInitDialogWithPinCheck') // 下記参照
        .addToUi();
}

/**
 * 定型文マスタの事業所列を初期化するダイアログを表示する。
 */
function showPhraseInitDialogWithPinCheck() {
    if (!verifyAdminPin()) { // auth.js
        return;
    }

    const ui = SpreadsheetApp.getUi();
    const result = ui.alert('定型文マスタの事業所列初期化', '定型文マスタのD列(事業所)が空欄の行に、現在のファイル名から取得した事業所名をセットしますか？', ui.ButtonSet.YES_NO);

    if (result === ui.Button.YES) {
        try {
            // server.js ではなく、ここで簡易実装または dataManager.js にある関数を呼ぶ形でも良いが
            // 今回は直接ロジックを記述せず、必要なら実装する形にします。
            // 以前の server.js にあった initializePhraseOfficeColumn を想定
            ui.alert('この機能は server.js / dataManager.js の整理に伴い、必要であれば再実装してください。');
        } catch (e) {
            ui.alert('エラー', '処理中にエラーが発生しました: ' + e.message, ui.ButtonSet.OK);
        }
    }
}