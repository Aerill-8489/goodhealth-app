function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  writeSheet(ss, '顧客',
    ['ID','姓名','性別','身高','年齡','分類','電話','健康狀況','過敏/禁忌','搭配產品','購買方式','購買日期','建立日期','最後回訪','目前體重','目標體重','達成率%','計劃名稱','備註'],
    data.customers || [],
    ['id','name','gender','height','age','category','phone','health','allergy','products','lastBuyMethod','lastBuyDate','createdDate','lastVisitDate','currentWeight','targetWeight','progressPct','planName','note']);

  writeSheet(ss, '挑戰期',
    ['顧客ID','顧客姓名','期別','月數','目標減%','起始體重','目標體重','開始日期','驗收日期','搭配保健品'],
    data.periods || [],
    ['customerId','customerName','period','months','targetPct','startWeight','targetWeight','startDate','endDate','supplements']);

  writeSheet(ss, '週記錄',
    ['顧客ID','顧客姓名','期別','週次','日期','體重','體年齡','體脂%','內臟脂肪','皮下脂肪','基礎代謝','BMI','骨骼肌','腰圍','臀圍','大腿圍','手臂圍','備註'],
    data.weeks || [],
    ['customerId','customerName','period','week','date','weight','bodyAge','bodyFat','visceralFat','subFat','bmr','bmi','boneMuscle','waist','hip','thigh','arm','note']);

  writeSheet(ss, '回報',
    ['ID','顧客','日期','備註','回覆','保健品','問答'],
    data.reports || [],
    ['id','customerName','date','note','reply','supps','qa']);

  writeSheet(ss, '使用者',
    ['ID','姓名','角色','狀態','負責分類'],
    data.users || [],
    ['id','name','role','active','dealerCategory']);

  writeSheet(ss, '計劃卡',
    ['ID','名稱','類型','飲水目標','睡眠目標','運動建議','禁忌食物','第一個月注意事項','減重提醒','減脂提醒','建議保健品','備註'],
    data.plans || [],
    ['id','name','type','waterGoal','sleepGoal','exerciseFreq','forbidFoods','monthOneNotes','weightAlert','fatAlert','supplements','notes']);

  writeSheet(ss, '飲食指南',
    ['ID','分類','標題','內容','日期'],
    data.guides || [],
    ['id','category','title','text','date']);

  writeSheet(ss, '回覆模板',
    ['ID','分類','內容'],
    data.templates || [],
    ['id','category','text']);

  writeSheet(ss, '分類',
    ['順序','名稱'],
    data.categories || [],
    ['idx','name']);

  writeSheet(ss, '回訪設定',
    ['項目','數值'],
    data.settings || [],
    ['key','value']);

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    customers: (data.customers || []).length,
    reports: (data.reports || []).length
  })).setMimeType(ContentService.MimeType.JSON);
}

function writeSheet(ss, name, headers, rows, keys) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  var values = [headers];
  rows.forEach(function (row) {
    values.push(keys.map(function (k) {
      var v = row[k];
      return (v === undefined || v === null) ? '' : v;
    }));
  });
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
}

function doGet(e) {
  return ContentService.createTextOutput('OK - 健康管理同步端點運作中');
}
