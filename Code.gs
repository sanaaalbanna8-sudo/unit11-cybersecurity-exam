var SHEET_ID = "1kwgeNZT5XDmEkv6jj_eTTEjvBQd9haT95FiZaH7mq7E";
var HEADERS = ["الوقت", "الاسم", "العلامة", "من", "النسبة", "التصنيف", "نتاج أ", "نتاج ب", "نتاج ج", "نتاج د"];

function doPost(e) {
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = {};
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("النتائج");
  if (!sheet) sheet = ss.insertSheet("النتائج");
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== HEADERS[0]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(),
    data.name || "",
    Number(data.correct) || 0,
    Number(data.total) || 0,
    Number(data.percent) || 0,
    data.band || "",
    data.outcomeA || "",
    data.outcomeB || "",
    data.outcomeC || "",
    data.outcomeD || ""
  ]);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MIME_TYPE.JSON);
}

function doGet() {
  return ContentService.createTextOutput("ready");
}
