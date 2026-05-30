const SHEET_ID = "1T32661HLRNWSz0vfm5Y3tBYt9XQ58QDLoopYW1WlA5s";
const SHEET_NAME = "客戶回覆資料";

const HEADERS = [
  "填寫時間",
  "客戶稱呼",
  "手機號碼",
  "想找區域",
  "物件類型",
  "屋齡限制",
  "預算",
  "自訂金額",
  "希望房型/幾房",
  "是否需要車位",
  "車位類型",
  "購屋用途",
  "購屋時程",
  "最在意條件",
  "風水/格局忌諱",
  "其他風水說明",
  "其他購屋需求",
  "同意聯繫",
  "買方類型",
  "清楚度分數",
  "最直白建議",
  "來源",
  "後續狀態",
  "承辦備註",
  "下次跟進日期"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse((e.postData && e.postData.contents) || "{}");
    const sheet = getTargetSheet_();
    ensureHeaders_(sheet);

    sheet.appendRow([
      new Date(),
      text_(data.name),
      text_(data.phone),
      text_(data.area),
      list_(data.propertyTypes),
      list_(data.ages),
      text_(data.budget),
      text_(data.customBudget),
      text_(data.rooms),
      text_(data.parking),
      text_(data.parkingType),
      text_(data.purpose),
      text_(data.timeline),
      list_(data.priorities),
      list_(data.fengShui),
      text_(data.fengShuiOther),
      text_(data.notes),
      data.consent ? "是" : "否",
      text_(data.persona),
      text_(data.clarityScore),
      [text_(data.directAdviceTitle), text_(data.directAdviceCopy)].filter(Boolean).join("。"),
      text_(data.source) || "互動網頁",
      "新名單",
      "",
      ""
    ]);

    return json_({ ok: true });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: "buyer lead receiver is running" });
}

function getTargetSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders_(sheet) {
  const existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeader = HEADERS.some((header, index) => existing[index] !== header);
  if (needsHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function list_(value) {
  return Array.isArray(value) ? value.filter(Boolean).join("、") : text_(value);
}

function text_(value) {
  return value === undefined || value === null ? "" : String(value).trim();
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
