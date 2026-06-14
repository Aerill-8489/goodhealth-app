function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var custSheet = ss.getSheetByName('顧客') || ss.insertSheet('顧客');
  custSheet.clearContents();
  custSheet.appendRow(['ID','姓名','性別','身高','年齡','分類','電話','健康狀況','過敏/禁忌','搭配產品','購買方式','購買日期','建立日期','最後回訪','目前體重','目標體重','達成率%']);
  (data.customers || []).forEach(function (c) {
    custSheet.appendRow([c.id, c.name, c.gender, c.height, c.age, c.category, c.phone, c.health, c.allergy, c.products, c.lastBuyMethod, c.lastBuyDate, c.createdDate, c.lastVisitDate, c.currentWeight, c.targetWeight, c.progressPct]);
  });

  var repSheet = ss.getSheetByName('回報') || ss.insertSheet('回報');
  repSheet.clearContents();
  repSheet.appendRow(['ID', '顧客', '日期', '備註', '回覆', '保健品', '問答']);
  (data.reports || []).forEach(function (r) {
    repSheet.appendRow([r.id, r.customerName, r.date, r.note, r.reply, r.supps, r.qa]);
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true, customers: (data.customers || []).length, reports: (data.reports || []).length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput('OK - 健康管理同步端點運作中');
}
