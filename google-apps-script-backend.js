const SHEET_ID = "1T32661HLRNWSz0vfm5Y3tBYt9XQ58QDLoopYW1WlA5s";
const SHEET_NAME = "客戶回覆資料";
const NOTIFY_EMAIL = "";

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
  "客戶熱度",
  "熱度判斷",
  "看房節奏",
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
      text_(data.leadTemperature),
      text_(data.leadTemperatureReason),
      text_(data.buyerReadiness),
      [text_(data.directAdviceTitle), text_(data.directAdviceCopy)].filter(Boolean).join("。"),
      text_(data.source) || "互動網頁",
      "新名單",
      "",
      ""
    ]);

    notifyByEmail_(data);

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

function notifyByEmail_(data) {
  try {
    const recipient = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    if (!recipient) return;

    const subject = `新買房測驗名單｜${text_(data.name) || "未填稱呼"}｜${text_(data.leadTemperature) || "未判斷"}`;
    const body = [
      "有新的買房方向測驗完成。",
      "",
      `稱呼：${text_(data.name)}`,
      `手機：${text_(data.phone)}`,
      `區域：${text_(data.area)}`,
      `預算：${text_(data.budget)} ${text_(data.customBudget)}`,
      `物件：${list_(data.propertyTypes)}`,
      `屋齡：${list_(data.ages)}`,
      `時程：${text_(data.timeline)}`,
      `熱度：${text_(data.leadTemperature)}（${text_(data.leadTemperatureReason)}）`,
      `買方類型：${text_(data.persona)}`,
      "",
      "最直白建議：",
      [text_(data.directAdviceTitle), text_(data.directAdviceCopy)].filter(Boolean).join("。"),
      "",
      "請到 Google 試算表查看完整資料。"
    ].join("\n");

    MailApp.sendEmail(recipient, subject, body);
  } catch (error) {
    console.log(`notify failed: ${error}`);
  }
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
