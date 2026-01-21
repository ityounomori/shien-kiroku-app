
// [Fix] Ensure global APP_STATE.pending is initialized early
// Place this strictly BEFORE document.addEventListener('DOMContentLoaded', ...)

(function () {
    // initialize if missing
    if (typeof APP_STATE === 'undefined') {
        window.APP_STATE = {};
    }
    if (!APP_STATE.pending) {
        APP_STATE.pending = { offset: 0, limit: 50, hasMore: true, loading: false };
    }
})();

// Original Script Logic below
// ...
