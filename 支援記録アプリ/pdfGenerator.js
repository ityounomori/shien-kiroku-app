// pdfGenerator.gs

/**
 * 指定された利用者と期間の記録からHTMLベースでPDFを生成する
 */
function outputPdf(userName, startDateStr, endDateStr, officeName) {
  try {
    // 1. 記録データを取得 (server.js のロジックに合わせる)
    const records = fetchRecordsForPdf(userName, startDateStr, endDateStr, officeName);

    // 2. HTML生成
    const htmlContent = createPdfHtml(userName, startDateStr, endDateStr, records, officeName);
    const blob = Utilities.newBlob(htmlContent, MimeType.HTML);
    const pdfBlob = blob.getAs(MimeType.PDF);

    const fileName = `支援記録_${userName}_${startDateStr.replace(/-/g, '')}-${endDateStr.replace(/-/g, '')}.pdf`;
    pdfBlob.setName(fileName);

    const base64 = Utilities.base64Encode(pdfBlob.getBytes());
    return { base64: base64, fileName: fileName };

  } catch (e) {
    Logger.log(`[outputPdf] Error: ${e.message}`);
    throw e;
  }
}

/**
 * 記録リストと同じロジックでデータを取得する
 */
function fetchRecordsForPdf(userName, startDateStr, endDateStr, officeName) {
  // server.js の getFilesByOffice を使用して正しいファイルIDを取得
  const files = getFilesByOffice(officeName);
  if (!files || !files.recordFileId) throw new Error('記録ファイルIDが見つかりません');

  const ss = SpreadsheetApp.openById(files.recordFileId);
  const sheet = ss.getSheetByName(SHEET_NAMES.RECORD_INPUT);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const start = new Date(startDateStr + 'T00:00:00');
  const end = new Date(endDateStr + 'T23:59:59');

  // データ範囲を一括取得
  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const records = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowUser = String(row[COL_INDEX.USER - 1] || '');

    // ユーザー一致チェック
    if (rowUser !== userName) continue;

    // v35 Fix: 日付パースの堅牢化 (文字列 "yyyy/MM/dd HH:mm" 等に対応)
    const rowDateVal = row[COL_INDEX.DATE - 1];
    let rowDate = rowDateVal;
    if (!(rowDate instanceof Date)) {
      const s = String(rowDateVal).trim();
      // "/Date(...)/" 形式や "yyyy/MM/dd ..." 形式を試行
      const parsed = new Date(s);
      if (!isNaN(parsed.getTime())) {
        rowDate = parsed;
      } else {
        // ハイフンやスラッシュを正規化して再試行
        const cleaned = s.replace(/[\uFF0D]/g, '-').replace(/[\uFF0F]/g, '/').replace(/年/g, '/').replace(/月/g, '/').replace(/日/g, '');
        const p2 = new Date(cleaned);
        if (!isNaN(p2.getTime())) rowDate = p2;
      }
    }

    // それでもDateでない場合、または無効な場合はスキップ
    if (!(rowDate instanceof Date) || isNaN(rowDate.getTime())) continue;

    // 日付範囲チェック
    if (rowDate < start || rowDate > end) continue;

    const item = String(row[COL_INDEX.ITEM - 1] || '');
    const d1 = row[COL_INDEX.DETAIL_1 - 1];
    const d2 = row[COL_INDEX.DETAIL_2 - 1];

    let temp = '', bp = '', pulse_spo2 = '', excretion = '', meal = '';

    // リスト表示用と同じロジックで整形
    if (item === 'バイタル') {
      const t = row[COL_INDEX.V_TEMP - 1];
      const bph = row[COL_INDEX.V_BP_HIGH - 1];
      const bpl = row[COL_INDEX.V_BP_LOW - 1];
      const p = row[COL_INDEX.V_PULSE - 1];
      const s = row[COL_INDEX.V_SPO2 - 1];

      if (t) temp = `${t}℃`;
      if (bph || bpl) bp = `${bph}/${bpl}`;

      const ps = [];
      if (p) ps.push(`P:${p}`);
      if (s) ps.push(`SpO2:${s}%`);
      pulse_spo2 = ps.join(' ');

    } else if (item === '食事') {
      const m = [];
      if (d1) m.push(`摂取:${d1}%`);
      if (d2) m.push(`水:${d2}ml`);
      meal = m.join(' ');

    } else if (item === '排泄') {
      excretion = d1 ? `[${d1}]` : '';
      if (d2) excretion += ` ${d2}`;
    }

    let content = String(row[COL_INDEX.CONTENT - 1] || '');

    // 服薬やその他の詳細を本文に結合
    if (item === '服薬' && d1) {
      content = `【量: ${d1}】\n${content}`;
    } else if (item === 'その他' && d1) {
      content = `【${d1}】\n${content}`;
    }

    records.push({
      date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'MM/dd'),
      time: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'HH:mm'),
      item: item,
      col_temp: temp,
      col_bp: bp,
      col_pulse: pulse_spo2,
      col_excretion: excretion,
      col_meal: meal,
      content: content,
      recorder: String(row[COL_INDEX.RECORDER - 1] || '')
    });
  }

  // 日付順にソート (古い順 = PDFは上から時系列)
  records.sort((a, b) => {
    if (a.date === b.date) return a.time.localeCompare(b.time);
    return a.date.localeCompare(b.date);
  });

  return records;
}

/**
 * 統合型テーブルHTML生成
 */
function createPdfHtml(userName, startDateStr, endDateStr, records, officeName) {
  const style = `
    <style>
      @page { size: A4 landscape; margin: 12mm; }
      body { font-family: sans-serif; font-size: 9pt; color: #333; }
      .header { margin-bottom: 10px; border-bottom: 2px solid #4f46e5; padding-bottom: 5px; }
      h1 { font-size: 16pt; margin: 0; display: inline-block; }
      .meta { float: right; font-size: 10pt; margin-top: 5px; }
      table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      thead { display: table-header-group; }
      tr { page-break-inside: avoid; }
      th, td { border: 0.5pt solid #000; padding: 4px 5px; vertical-align: top; word-wrap: break-word; overflow-wrap: break-word; }
      th { background-color: #f3f4f6; font-weight: bold; text-align: center; height: 25px; }
      .w-datetime { width: 8%; } .w-item { width: 6%; } .w-temp { width: 5%; } .w-bp { width: 7%; }
      .w-pulse { width: 9%; } .w-exc { width: 6%; } .w-meal { width: 10%; } .w-content { width: 40%; } .w-staff { width: 9%; }
      .text-center { text-align: center; } .content-text { white-space: pre-wrap; font-size: 9pt; }
      .item-vital { color: #d32f2f; font-weight: bold; }
    </style>
  `;

  const rowsHtml = records.map(rec => {
    let itemClass = rec.item === 'バイタル' ? 'item-vital' : '';
    return `
      <tr>
        <td class="text-center">${rec.date}<br><span style="font-size:0.85em; color:#555;">${rec.time}</span></td>
        <td class="text-center ${itemClass}">${rec.item}</td>
        <td class="text-center">${rec.col_temp}</td>
        <td class="text-center">${rec.col_bp}</td>
        <td class="text-center">${rec.col_pulse}</td>
        <td class="text-center">${rec.col_excretion}</td>
        <td class="text-center">${rec.col_meal}</td>
        <td><div class="content-text">${rec.content}</div></td>
        <td class="text-center">${rec.recorder}</td>
      </tr>
    `;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8">${style}</head>
    <body>
      <div class="header">
        <h1>支援記録レポート</h1>
        <div class="meta">
          <span><strong>利用者:</strong> ${userName} 様</span> &nbsp;|&nbsp;
          <span><strong>期間:</strong> ${startDateStr} 〜 ${endDateStr}</span> &nbsp;|&nbsp;
          <span><strong>事業所:</strong> ${officeName}</span>
        </div>
      </div>
      <table>
        <thead>
          <tr>
            <th class="w-datetime">日時</th><th class="w-item">区分</th><th class="w-temp">体温</th><th class="w-bp">血圧</th>
            <th class="w-pulse">脈/SpO2</th><th class="w-exc">排泄</th><th class="w-meal">食事/水分</th>
            <th class="w-content">経過記録・特記事項</th><th class="w-staff">記録者</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </body>
    </html>
  `;
}