/**
 * CoreLib: アプリケーションコアライブラリの公開インターフェース
 * 
 * リファクタリングフェーズ2で導入。
 * 将来的に別プロジェクト（AppProject）からライブラリとして読み込まれることを想定。
 */
var CoreLib = {
    /**
     * 認証関連サービス
     */
    Auth: AuthService,

    /**
     * 設定関連ヘルパー
     */
    Config: {
        /**
         * マスタスプレッドシートのIDを設定する (実行時注入)
         * @param {string} id スプレッドシートID
         */
        setMasterFileId: function (id) {
            injectMasterFileId(id);
        },

        /**
         * インシデント記録用スプレッドシートのIDを設定する（実行時注入）
         * @param {string} id スプレッドシートID
         */
        setIncidentFileId: function (id) {
            injectIncidentFileId(id);
        }
    }
};

/**
 * ライブラリの初期化確認用メソッド
 * @returns {Object} バージョン情報とステータス
 */
function getLibraryStatus() {
    return {
        version: '1.0.0',
        status: 'active',
        masterFileIdConfigured: !!PropertiesService.getScriptProperties().getProperty('MASTER_FILE_ID'),
        incidentFileIdConfigured: !!PropertiesService.getScriptProperties().getProperty('INCIDENT_FILE_ID')
    };
}
