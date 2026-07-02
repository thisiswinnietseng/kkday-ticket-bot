// =========================================
// KKday 挽單自動化 — Google Sheets 串接
// =========================================
// 使用方式：
// 1. 開一張新的 Google Sheet
// 2. 點上方選單「擴充功能」→「Apps Script」
// 3. 把這段程式碼貼上去（取代原本的內容）
// 4. 點「部署」→「新增部署作業」
//    - 類型選「網頁應用程式」
//    - 執行身分：我（你的 Google 帳號）
//    - 誰可以存取：所有人
// 5. 複製部署後的網址
// 6. 貼到 kkday-ticket-bot/.env 的 SHEETS_WEBHOOK=網址
// =========================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 如果是第一次，自動加標題列
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '時間', '處理人', '訂單編號', '工單 ID',
        '使用日期', '最晚處理', '後拋截止', '功能', '急/一般', '狀態', '錯誤訊息'
      ]);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a9e5c').setFontColor('#ffffff');
    }

    sheet.appendRow([
      data.timestamp || '',
      data.operator || '',
      data.order_id || '',
      data.ticket_id ? '#' + data.ticket_id : '',
      data.use_date || '',
      data.ticket_deadline || '',
      data.follow_deadline || '',
      data.type || '挽單',
      data.is_urgent ? '急單' : '一般單',
      data.success ? '✅ 成功' : '❌ 失敗',
      data.error || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
