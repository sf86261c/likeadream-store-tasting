// 取得試算表與工作表
var SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var SHEET_NAME = "預約資料"; // 請確保與您的工作表名稱一致

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('門市試吃預約系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 取得特定日期的已佔用時段
function getTakenSlots(queryDateStr) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 讀取 A 到 H 欄 (因為日期在 G(索引6)，時段在 H(索引7))
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var taken = [];
  var timeZone = Session.getScriptTimeZone();

  for (var i = 0; i < data.length; i++) {
    var cellDate = data[i][6]; // ✅ 修改為 G 欄：日期
    var cellTime = data[i][7]; // ✅ 修改為 H 欄：時段

    if (!cellDate) continue;

    try {
      var dateStr = "";
      if (cellDate instanceof Date) {
        dateStr = Utilities.formatDate(cellDate, timeZone, "yyyy-MM-dd");
      } else {
        dateStr = Utilities.formatDate(new Date(cellDate), timeZone, "yyyy-MM-dd");
      }

      if (dateStr === queryDateStr) {
        var timeStr = cellTime;
        if (cellTime instanceof Date) {
          timeStr = Utilities.formatDate(cellTime, timeZone, "HH:mm");
        }
        taken.push(String(timeStr));
      }
    } catch (e) {
      Logger.log("Row error: " + e);
      continue;
    }
  }
  return taken;
}

// 處理表單提交 (新增人數、口味、LINE身分、付款方式欄位寫入)
function processForm(formObject) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

  var checkDate = new Date(formObject.date);
  if (checkDate.getDay() === 0) {
      return {status: 'error', message: '預約失敗：週日不開放預約。'};
  }

  var currentTaken = getTakenSlots(formObject.date);
  if (currentTaken.includes(formObject.time)) {
    return {status: 'error', message: '抱歉！該時段剛剛已被搶先預約，請重新選擇。'};
  }

  // 尋找 A 欄真正的最後一列 (防核取方塊干擾)
  var aColumnValues = sheet.getRange("A:A").getValues();
  var trueLastRow = 0;
  for (var i = aColumnValues.length - 1; i >= 0; i--) {
    if (aColumnValues[i][0] !== "") {
      trueLastRow = i + 1;
      break;
    }
  }
  var targetRow = (trueLastRow === 0) ? 2 : trueLastRow + 1;

  // ✅ 擴增寫入資料矩陣 (對應 A 到 N 欄)
  // A:提交時間 B:姓名 C:電話 D:人數 E:口味1 F:口味2 G:日期 H:時段 I:處理狀態
  // J:LINE userId K:LINE顯示名稱 L:付款方式 M:匯款後五碼 N:總金額
  var rowData = [
    new Date(),                     // A欄: 提交時間
    formObject.name,                // B欄: 姓名
    "'" + formObject.phone,         // C欄: 電話
    formObject.peopleCount,         // D欄: 人數
    formObject.flavor1,             // E欄: 口味1
    formObject.flavor2 || "",       // F欄: 口味2 (若只有1人，此值為 undefined，轉換為空字串)
    formObject.date,                // G欄: 預約日期
    formObject.time,                // H欄: 預約時段
    false,                          // I欄: 處理狀態(核取方塊)
    formObject.userId || "",        // J欄: LINE userId
    formObject.lineName || "",      // K欄: LINE顯示名稱
    formObject.paymentMethod || "", // L欄: 付款方式
    formObject.last5 || "",         // M欄: 匯款後五碼
    formObject.totalAmount || ""    // N欄: 總金額
  ];

  // ✅ 寫入範圍改為 1 列 14 欄
  sheet.getRange(targetRow, 1, 1, 14).setValues([rowData]);

  return {status: 'success', message: '預約資料已送出！請等待客服人員聯繫確認才算預約成功。'};
}
