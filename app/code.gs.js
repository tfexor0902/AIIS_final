function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const { category = "未分類", name = "無名稱", amount = 0, date } = params;

    const spreadsheet = SpreadsheetApp.openById('1BJt6ZqRsIrJfwJlhiwn_6noXDpxVYdVyA1fiBlgR2Pk');
    const sheet = spreadsheet.getSheetByName('記帳明細');

    const columnB = sheet.getRange("B:B").getValues().flat(); 
    let nextRow = columnB.findIndex(value => value === "") + 1; 

    if (nextRow === 0) {
      nextRow = sheet.getLastRow() + 1;
    }

    const rowData = [category, name, amount, date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd")];
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    return ContentService.createTextOutput("已新增").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput(`發生錯誤: ${error.message}`).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * 🌐 網頁部署進入點（已啟用樣板引擎）
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Chart');
  return template.evaluate()
      .setTitle('自動化記帳統計')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 🛠️ 試算表選單彈出視窗（★已修正：使用 evaluate() 允許內嵌 CSS/JS）
 */
function showChartDialog() {
  const template = HtmlService.createTemplateFromFile('Chart');
  const htmlOutput = template.evaluate()
      .setWidth(780)  
      .setHeight(750) 
      .setTitle('我的記帳儀表板');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '我的記帳儀表板');
}

/**
 * 🧠 核心聯件功能：負責把 CSS.html 與 JS.html 的內容即時注入到主網頁中
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 打包儀表板所需的所有圖表數據
function getChartData() {
  const spreadsheet = SpreadsheetApp.openById('1BJt6ZqRsIrJfwJlhiwn_6noXDpxVYdVyA1fiBlgR2Pk');
  const sheet = spreadsheet.getSheetByName('記帳明細');
  
  if (!sheet) {
    return { categoryData: [], dateData: [], weekData: [], rawData: [], budgetInfo: { budget: 0, totalExpense: 0, usePercent: 0 } };
  }
  
  const budget = parseFloat(sheet.getRange("F2").getValue()) || 0;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { categoryData: [], dateData: [], weekData: [], rawData: [], budgetInfo: { budget: budget, totalExpense: 0, usePercent: 0 } }; 
  }

  const categorySummary = {};
  const dateSummary = {};
  let totalExpense = 0; 
  
  const weekNames = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
  const weekSummary = { "星期一": 0, "星期二": 0, "星期三": 0, "星期四": 0, "星期五": 0, "星期六": 0, "星期日": 0 };
  
  const weekDaysMap = { "星期一": {}, "星期二": {}, "星期三": {}, "星期四": {}, "星期五": {}, "星期六": {}, "星期日": {} };
  
  let serializedRawData = [];
  serializedRawData.push([data[0][0], data[0][1], data[0][2], data[0][3]]);

  for (let i = 1; i < data.length; i++) {
    const category = data[i][0];  
    const amountStr = data[i][2]; 
    let dateVal = data[i][3];     
    
    const amount = parseFloat(amountStr) || 0; 
    if (amount <= 0) continue; 
    totalExpense += amount; 

    const categoryName = (category && category.toString().trim()) ? category.toString().trim() : "未分類";
    categorySummary[categoryName] = (categorySummary[categoryName] || 0) + amount;

    let dateStr = "";
    let dayIndex = -1;

    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      dayIndex = dateVal.getDay();
    } else if (dateVal) {
      dateStr = dateVal.toString().split("T")[0].replace(/-/g, '/');
      let parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        dayIndex = parsedDate.getDay();
      }
    }

    if (dateStr && dayIndex !== -1) {
      dateSummary[dateStr] = (dateSummary[dateStr] || 0) + amount;
      const wName = weekNames[dayIndex];
      weekSummary[wName] += amount;
      weekDaysMap[wName][dateStr] = true;
    }
    
    serializedRawData.push([
      categoryName,
      data[i][1] ? data[i][1].toString() : "",
      amount,
      dateStr || "—"
    ]);
  }
  
  const categoryData = [['類別', '總支出']];
  for (let key in categorySummary) { categoryData.push([key, categorySummary[key]]); }
  
  const sortedDates = Object.keys(dateSummary).sort((a, b) => new Date(a) - new Date(b));
  const dateData = [['日期', '每日總支出']];
  sortedDates.forEach(date => { dateData.push([date, dateSummary[date]]); });
  
  const weekOrder = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
  const weekData = [['星期', '平均支出金額']];
  weekOrder.forEach(day => {
    const daysCount = Object.keys(weekDaysMap[day]).length || 1; 
    const averageAmount = Math.round((weekSummary[day] / daysCount) * 100) / 100;
    weekData.push([day, averageAmount]);
  });
  
  const usePercent = budget > 0 ? (totalExpense / budget) * 100 : 0;
  
  return {
    categoryData: categoryData,
    dateData: dateData,
    weekData: weekData, 
    rawData: serializedRawData, 
    budgetInfo: { 
      budget: budget, 
      totalExpense: totalExpense, 
      usePercent: usePercent
    }
  };
}

function getFilteredDateData(targetDateStr) {
  try {
    const spreadsheet = SpreadsheetApp.openById('1BJt6ZqRsIrJfwJlhiwn_6noXDpxVYdVyA1fiBlgR2Pk');
    const sheet = spreadsheet.getSheetByName('記帳明細');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let filteredRows = [];
    const targetFormatted = targetDateStr.trim().replace(/-/g, '/');

    for (let i = 1; i < data.length; i++) {
      let dateVal = data[i][3];
      let rowDateStr = "";
      
      if (dateVal instanceof Date) {
        rowDateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else {
        rowDateStr = dateVal ? dateVal.toString().split("T")[0].replace(/-/g, '/') : "";
      }
      
      if (rowDateStr === targetFormatted) {
        filteredRows.push({
          category: data[i][0] ? data[i][0].toString().trim() : "未分類",
          item: data[i][1] ? data[i][1].toString() : "無備註",
          amount: parseFloat(data[i][2]) || 0
        });
      }
    }
    return filteredRows;
  } catch(e) {
    return [];
  }
}

function getFilteredWeekData(targetWeekName) {
  try {
    const spreadsheet = SpreadsheetApp.openById('1BJt6ZqRsIrJfwJlhiwn_6noXDpxVYdVyA1fiBlgR2Pk');
    const sheet = spreadsheet.getSheetByName('記帳明細');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let filteredRows = [];
    const weekNames = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
    const targetFormatted = targetWeekName.trim();

    for (let i = 1; i < data.length; i++) {
      let dateVal = data[i][3];
      let rowWeekName = "";
      let dateStr = "";
      
      if (dateVal instanceof Date) {
        rowWeekName = weekNames[dateVal.getDay()];
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else if (dateVal) {
        let parsedDate = new Date(dateVal.toString().replace(/-/g, '/'));
        if (!isNaN(parsedDate.getTime())) {
          rowWeekName = weekNames[parsedDate.getDay()];
          dateStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
        }
      }
      
      if (rowWeekName === targetFormatted) {
        filteredRows.push({
          date: dateStr,
          category: data[i][0] ? data[i][0].toString().trim() : "未分類",
          item: data[i][1] ? data[i][1].toString() : "無備註",
          amount: parseFloat(data[i][2]) || 0
        });
      }
    }
    return filteredRows;
  } catch(e) {
    return [];
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('記帳功能')
    .addItem('查看記帳儀表板', 'showChartDialog')
    .addToUi();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('記帳明細');
  if (!sheet) return;

  if (sheet.getRange("F1").getValue() !== "每月預算設定") {
    sheet.getRange("F1").setValue("每月預算設定");
  }

  const currentBudget = sheet.getRange("F2").getValue();
  if (!currentBudget || currentBudget === "") {
    const response = ui.prompt('💰 首次預算設定', '尚未設定本月預算，請輸入金額：', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      const inputText = response.getResponseText().trim();
      const newBudget = parseFloat(inputText);
      if (!isNaN(newBudget) && newBudget > 0) {
        sheet.getRange("F2").setValue(newBudget);
      } else {
        ui.alert('設定失敗', '金額格式不正確，請開啟儀表板再行設定。', ui.ButtonSet.OK);
      }
    }
    return;
  }

  const activeBudget = parseFloat(currentBudget);
  if (activeBudget <= 0) return;

  const data = sheet.getDataRange().getValues();
  let totalExpense = 0;
  for (let i = 1; i < data.length; i++) {
    const amount = parseFloat(data[i][2]) || 0;
    if (amount > 0) totalExpense += amount;
  }

  const usePercent = (totalExpense / activeBudget) * 100;
  const remaining = activeBudget - totalExpense;

  if (usePercent >= 100) {
    ui.alert('🚨 🤯 【破產紅色警報】預算已完全炸裂！', '目前的預算使用率已達到 ' + usePercent.toFixed(1) + '%！\n累積總支出：$' + totalExpense.toLocaleString() + ' 元\n您已經透支了：$' + Math.abs(remaining).toLocaleString() + ' 元！\n\n💡 財務管家：開啟月底吃土防禦壁！', ui.ButtonSet.OK);
  } else if (usePercent >= 90) {
    ui.alert('⚠️ 🔥 【橘色深度預警】錢包正在痛苦哀嚎！', '預算使用率已飆升至 ' + usePercent.toFixed(1) + '%！\n目前已花費：$' + totalExpense.toLocaleString() + ' 元\n僅剩餘額度：$' + remaining.toLocaleString() + ' 元\n\n💡 財務管家：接下來請高舉「非必要不購買」盾牌！', ui.ButtonSet.OK);
  } else if (usePercent >= 80) {
    ui.alert('💡 🚦 【黃色警戒】不知不覺已經花了八成了喔！', '注意！本月預算已消耗了 ' + usePercent.toFixed(1) + '%。\n目前已花費：$' + totalExpense.toLocaleString() + ' 元\n剩餘可用額度：$' + remaining.toLocaleString() + ' 元\n\n💡 財務管家：建議購物慾望先放進購物車冷凍 3 天！', ui.ButtonSet.OK);
  } else {
    ui.alert('✨ 🥦 【進度安全】目前財務狀況非常健康！', '目前的預算使用率為 ' + usePercent.toFixed(1) + '%。\n目前已花費：$' + totalExpense.toLocaleString() + ' 元\n本月還剩下：$' + remaining.toLocaleString() + ' 元可以使用。\n\n💡 財務管家：請保持這個完美的節奏到月底！', ui.ButtonSet.OK);
  }
}

function updateBudget(newBudget) {
  try {
    const amount = parseFloat(newBudget);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: '金額格式不正確' };
    }
    const spreadsheet = SpreadsheetApp.openById('1BJt6ZqRsIrJfwJlhiwn_6noXDpxVYdVyA1fiBlgR2Pk');
    const sheet = spreadsheet.getSheetByName('記帳明細');
    sheet.getRange("F2").setValue(amount);
    return { success: true, budget: amount };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
