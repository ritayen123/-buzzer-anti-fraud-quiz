/**
 * 反詐小測驗 — 題庫 API（Google Apps Script）
 *
 * 部署方式：
 * 1. 在 Google Sheet 中開啟「擴充功能 > Apps Script」
 * 2. 貼上此程式碼
 * 3. 部署 > 新增部署 > 網頁應用程式
 *    - 執行身分：我自己
 *    - 存取權：任何人
 * 4. 複製部署網址，貼到 index.html 的 APPS_SCRIPT_URL
 *
 * API 參數：
 *   ?tag=anti_fraud    篩選標籤（不帶 = 全部）
 *   ?count=3           回傳題數（不帶 = 1）
 *
 * Google Sheet 欄位格式（工作表名稱：題庫）：
 *   A: 題目
 *   B: 選項A
 *   C: 選項B
 *   D: 選項C
 *   E: 選項D
 *   F: 正確答案（數字 0-3，對應選項 A-D）
 *   G: 解說
 *   H: 標籤（例如 anti_fraud）
 *   I: 啟用（TRUE / FALSE）
 */

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('題庫');
  var data = sheet.getDataRange().getValues();
  var tag = (e.parameter.tag || '').trim();
  var count = parseInt(e.parameter.count, 10) || 1;

  var questions = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // 跳過未啟用
    if (row[8] !== true) continue;
    // 標籤篩選
    if (tag && String(row[7]).trim() !== tag) continue;
    // 跳過空題目
    if (!String(row[0]).trim()) continue;

    questions.push({
      question: String(row[0]),
      options: [String(row[1]), String(row[2]), String(row[3]), String(row[4])],
      answer: parseInt(row[5], 10),
      explanation: String(row[6])
    });
  }

  // 隨機洗牌
  for (var j = questions.length - 1; j > 0; j--) {
    var k = Math.floor(Math.random() * (j + 1));
    var temp = questions[j];
    questions[j] = questions[k];
    questions[k] = temp;
  }

  // 取指定題數
  var selected = questions.slice(0, count);

  return ContentService
    .createTextOutput(JSON.stringify({ questions: selected }))
    .setMimeType(ContentService.MimeType.JSON);
}
