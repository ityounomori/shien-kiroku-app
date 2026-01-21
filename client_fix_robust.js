function loadIncidentPendingList(isLoadMore) {
    // Defensive init
    if (typeof APP_STATE === 'undefined') window.APP_STATE = {};
    if (!APP_STATE.pending) APP_STATE.pending = { offset: 0, limit: 50, hasMore: true, loading: false };

    // Default arg handling
    if (typeof isLoadMore === 'undefined') isLoadMore = false;

    if (!isLoadMore) {
        if (APP_STATE.pending.loading) return;
        // [Fix] Use common loader spinner
        const listEl = document.getElementById('incident-pending-list');
        if (listEl) {
            // Check if getLoaderHtml is available, otherwise fallback
            const spinner = (typeof getLoaderHtml === 'function')
                ? '<div class="flex justify-center items-center py-20 w-full">' + getLoaderHtml() + '</div>'
                : '<div class="text-center py-8 text-gray-500"><i class="fas fa-spinner fa-spin mr-2"></i>読み込み中...</div>';
            listEl.innerHTML = spinner;
        }
        document.getElementById('incident-pending-load-more').innerHTML = ''; // Clear button
        APP_STATE.pending = { offset: 0, limit: 50, hasMore: true, loading: false };
    }

    if (APP_STATE.pending.loading) return;
    APP_STATE.pending.loading = true;

    // Show loading indicator in button if "Load More"
    const moreDiv = document.getElementById('incident-pending-load-more');
    if (isLoadMore && moreDiv) {
        moreDiv.innerHTML = '<div class="py-4 text-center"><i class="fas fa-spinner fa-spin text-blue-500"></i></div>';
    }

    const offset = APP_STATE.pending.offset;
    const limit = APP_STATE.pending.limit;

    google.script.run
        .withSuccessHandler(function (response) {
            APP_STATE.pending.loading = false;

            // [Debug] Dump Server Logs
            if (response && response.debugLogs) {
                console.log('[ServerLog] --- Start ---');
                if (Array.isArray(response.debugLogs)) {
                    response.debugLogs.forEach(function (l) { console.log('[ServerLog]', l); });
                }
                console.log('[ServerLog] --- End ---');
            }
            if (response && response.debugRowDump) {
                console.log('[ServerRowDump] First 20 Scanned Rows:');
                if (console.table) console.table(response.debugRowDump);
                else console.log(response.debugRowDump);
            }

            if (!response || !response.data) {
                console.error('Invalid response:', response);
                if (!isLoadMore) document.getElementById('incident-pending-list').innerHTML = '<div class="text-center py-8 text-red-500">エラー: データ取得失敗</div>';
                return;
            }

            const listEl = document.getElementById('incident-pending-list');
            if (!isLoadMore) listEl.innerHTML = '';

            if (response.data.length === 0 && !isLoadMore) {
                listEl.innerHTML = '<div class="text-center py-8 text-gray-400">未承認の報告はありません</div>';
                moreDiv.innerHTML = '';
                return;
            }

            response.data.forEach(function (item) {
                listEl.insertAdjacentHTML('beforeend', createIncidentPendingCard(item));
            });

            APP_STATE.pending.offset += response.data.length;
            APP_STATE.pending.hasMore = response.hasMore;

            // Render "Load More" button if needed
            if (APP_STATE.pending.hasMore) {
                // [Fix] Add relative and z-50 to ensure clickability
                moreDiv.innerHTML = '<div class="flex justify-center py-4 relative z-50">' +
                    '<button id="btn-pending-load-more" onclick="loadIncidentPendingList(true)" class="relative z-50 bg-white border border-gray-300 text-gray-600 font-bold py-2 px-6 rounded-full shadow-sm hover:bg-gray-50 transition flex items-center gap-2">' +
                    '<span>もっと読み込む</span>' +
                    '<i class="fas fa-chevron-down text-sm"></i>' +
                    '</button></div>';

                // Failsafe
                setTimeout(function () {
                    if (APP_STATE.pending.loading) {
                        APP_STATE.pending.loading = false;
                    }
                }, 5000);

            } else {
                moreDiv.innerHTML = '<div class="text-center py-4 text-gray-400 text-sm">すべて表示しました</div>';
            }

        })
        .withFailureHandler(function (error) {
            APP_STATE.pending.loading = false;
            console.error('Pending load error:', error);
            alert('読み込みエラー: ' + error.message);
            if (isLoadMore) {
                document.getElementById('incident-pending-load-more').innerHTML = '<div class="text-center text-red-500">再試行してください</div>';
            }
        })
        .getPendingIncidentsByOffice(config.officeName, limit, offset, whoami);
}
