// ===================================================================
// 全域常數（集中管理，方便日後修改）
// ===================================================================
const SPREADSHEET_ID = '你的id';
const SHEET_NAME     = '記帳明細';
const BUDGET_CELL    = 'F2';
const Gemini_API     = '你的API_Key'


/**
 * 📥 【功能 1】接收從手機、網頁傳過來的記帳資料 (POST 請求)
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("錯誤: 找不到 postData 內容").setMimeType(ContentService.MimeType.TEXT);
    }

    const params = JSON.parse(e.postData.contents);
    const { category = "未分類", name = "無名稱", amount = 0, date } = params;

    const rawDate      = date ? new Date(date) : new Date();
    const formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet       = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return ContentService.createTextOutput("錯誤:找不到『記帳明細』工作表").setMimeType(ContentService.MimeType.TEXT);
    }

    const rowData = [category, name, Number(amount), formattedDate];
    sheet.appendRow(rowData);

    return ContentService.createTextOutput("已新增").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput("發生錯誤: " + error.message).setMimeType(ContentService.MimeType.TEXT);
  }
}


/**
 * 🌐 【功能 2】網頁部署進入點 (GET 請求)
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Chart');
  return template.evaluate()
      .setTitle('自動化記帳統計')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * 🛠️ 【功能 3】在試算表內部彈出儀表板視窗
 */
function showChartDialog() {
  const template   = HtmlService.createTemplateFromFile('Chart');
  const htmlOutput = template.evaluate()
      .setWidth(780)
      .setHeight(750)
      .setTitle('我的記帳儀表板');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '我的記帳儀表板');
}


/**
 * 🧠 【功能 4】網頁元件拼盤 (組合 CSS 與 JS 檔案)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * 📊 【功能 5】後台大腦：撈出記帳本所有資料，統計成圖表要用的數據
 */
function getChartData() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet       = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { categoryData: [], dateData: [], weekData: [], rawData: [], budgetInfo: { budget: 0, totalExpense: 0, usePercent: 0 }, aiResponse: "" };
  }

  const budget = parseFloat(sheet.getRange(BUDGET_CELL).getValue()) || 0;
  const data   = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return { categoryData: [], dateData: [], weekData: [], rawData: [], budgetInfo: { budget: budget, totalExpense: 0, usePercent: 0 }, aiResponse: "" };
  }

  const categorySummary = {};
  const dateSummary     = {};
  let   totalExpense    = 0;

  const weekNames   = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
  const weekSummary = { "星期一": 0, "星期二": 0, "星期三": 0, "星期四": 0, "星期五": 0, "星期六": 0, "星期日": 0 };
  const weekDaysMap = { "星期一": {}, "星期二": {}, "星期三": {}, "星期四": {}, "星期五": {}, "星期六": {}, "星期日": {} };

  let serializedRawData = [];
  serializedRawData.push([data[0][0], data[0][1], data[0][2], data[0][3]]);

  for (let i = 1; i < data.length; i++) {
    const category = data[i][0];
    const amountStr = data[i][2];
    let   dateVal   = data[i][3];

    const amount = parseFloat(amountStr) || 0;
    if (amount <= 0) continue;

    totalExpense += amount;

    const categoryName = (category && category.toString().trim()) ? category.toString().trim() : "未分類";
    categorySummary[categoryName] = (categorySummary[categoryName] || 0) + amount;

    let dateStr  = "";
    let dayIndex = -1;

    if (dateVal instanceof Date) {
      dateStr  = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
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

  // 整理圖表資料
  const categoryData = [['類別', '總支出']];
  for (let key in categorySummary) { categoryData.push([key, categorySummary[key]]); }

  const sortedDates = Object.keys(dateSummary).sort((a, b) => new Date(a) - new Date(b));
  const dateData    = [['日期', '每日總支出']];
  sortedDates.forEach(date => { dateData.push([date, dateSummary[date]]); });

  const weekOrder = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
  const weekData  = [['星期', '平均支出金額']];
  weekOrder.forEach(day => {
    const daysCount     = Object.keys(weekDaysMap[day]).length || 1;
    const averageAmount = Math.round((weekSummary[day] / daysCount) * 100) / 100;
    weekData.push([day, averageAmount]);
  });

  const usePercent = budget > 0 ? (totalExpense / budget) * 100 : 0;
  const remaining  = budget - totalExpense;

  let categoryTextForWeb = "";
  for (let key in categorySummary) {
    categoryTextForWeb += `${key}: $${categorySummary[key].toLocaleString()} 元; `;
  }

  const webGeminiAdvice = askGeminiAdvisor(totalExpense, remaining, usePercent, categoryTextForWeb);

  return {
    categoryData : categoryData,
    dateData     : dateData,
    weekData     : weekData,
    rawData      : serializedRawData,
    budgetInfo   : { budget: budget, totalExpense: totalExpense, usePercent: usePercent },
    aiResponse   : webGeminiAdvice
  };
}


/**
 * 🔍 【功能 6】互動功能：點選圖表上的「某個日期」，撈出當天的所有消費明細
 */
function getFilteredDateData(targetDateStr) {
  try {
    const spreadsheet   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet         = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    const data            = sheet.getDataRange().getValues();
    let   filteredRows    = [];
    const targetFormatted = targetDateStr.trim().replace(/-/g, '/');

    for (let i = 1; i < data.length; i++) {
      let dateVal    = data[i][3];
      let rowDateStr = "";

      if (dateVal instanceof Date) {
        rowDateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else {
        rowDateStr = dateVal ? dateVal.toString().split("T")[0].replace(/-/g, '/') : "";
      }

      if (rowDateStr === targetFormatted) {
        filteredRows.push({
          category : data[i][0] ? data[i][0].toString().trim() : "未分類",
          item     : data[i][1] ? data[i][1].toString() : "無備註",
          amount   : parseFloat(data[i][2]) || 0
        });
      }
    }
    return filteredRows;
  } catch(e) {
    return [];
  }
}


/**
 * 💾 【功能 7】更新每月預算
 * 修正重點：寫入後強制 flush()，再重新呼叫 getChartData() 確保 AI 讀到最新預算
 */
function updateBudget(newBudgetVal) {
  try {
    const num = parseFloat(newBudgetVal);
    if (isNaN(num) || num <= 0) return { success: false, message: "請輸入大於 0 的有效數字" };

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet       = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: "找不到記帳明細工作表" };

    // 寫入並強制立即刷新，避免 AI 讀到舊值
    sheet.getRange(BUDGET_CELL).setValue(num);
    SpreadsheetApp.flush();

    // 重新取得完整統計（含最新 AI 評語）
    const chartData = getChartData();

    // 防呆：若 Gemini 暫時斷線給預設文字
    const finalAiResponse = (chartData.aiResponse && chartData.aiResponse.trim() !== "")
      ? chartData.aiResponse
      : "🤖 管家剛剛揉了揉眼睛，沒看清帳本，請再試著更新一次預算看看！";

    return {
      success    : true,
      budget     : num,
      usePercent : chartData.budgetInfo.usePercent,
      aiResponse : finalAiResponse
    };
  } catch (e) {
    console.error("更新預算錯誤: " + e.message);
    return { success: false, message: e.message };
  }
}


/**
 * 🔍 【功能 8】互動功能：點選圖表上的「星期幾」，撈出所有屬於該星期幾的消費明細
 */
function getFilteredWeekData(targetWeekName) {
  try {
    const spreadsheet  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet        = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    const data          = sheet.getDataRange().getValues();
    let   filteredRows  = [];
    const weekNames     = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
    const targetFormatted = targetWeekName.trim();

    for (let i = 1; i < data.length; i++) {
      let dateVal    = data[i][3];
      let rowWeekName = "";
      let dateStr     = "";

      if (dateVal instanceof Date) {
        rowWeekName = weekNames[dateVal.getDay()];
        dateStr     = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else if (dateVal) {
        let parsedDate = new Date(dateVal.toString().replace(/-/g, '/'));
        if (!isNaN(parsedDate.getTime())) {
          rowWeekName = weekNames[parsedDate.getDay()];
          dateStr     = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
        }
      }

      if (rowWeekName === targetFormatted) {
        filteredRows.push({
          date     : dateStr,
          category : data[i][0] ? data[i][0].toString().trim() : "未分類",
          item     : data[i][1] ? data[i][1].toString() : "無備註",
          amount   : parseFloat(data[i][2]) || 0
        });
      }
    }
    return filteredRows;
  } catch(e) {
    return [];
  }
}


/**
 * 🔔 【功能 9】打開記帳本時的自動歡迎與理財警告選單
 */
function installableOnOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('記帳功能')
    .addItem('查看記帳儀表板', 'showChartDialog')
    .addToUi();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

  if (sheet.getRange("F1").getValue() !== "每月預算設定") {
    sheet.getRange("F1").setValue("每月預算設定");
  }

  const currentBudget = sheet.getRange(BUDGET_CELL).getValue();
  if (!currentBudget || currentBudget === "") {
    const response = ui.prompt('💰 首次預算設定', '尚未設定本月預算，請輸入金額：', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      const newBudget = parseFloat(response.getResponseText().trim());
      if (!isNaN(newBudget) && newBudget > 0) {
        sheet.getRange(BUDGET_CELL).setValue(newBudget);
      }
    }
    return;
  }

  const activeBudget = parseFloat(currentBudget);
  if (activeBudget <= 0) return;

  const data = sheet.getDataRange().getValues();
  let totalExpense = 0;
  let categoryMap  = {};

  for (let i = 1; i < data.length; i++) {
    const cat    = data[i][0] ? data[i][0].toString().trim() : "未分類";
    const amount = parseFloat(data[i][2]) || 0;
    if (amount > 0) {
      totalExpense += amount;
      categoryMap[cat] = (categoryMap[cat] || 0) + amount;
    }
  }

  const usePercent  = (totalExpense / activeBudget) * 100;
  const remaining   = activeBudget - totalExpense;

  let categoryText = "";
  for (let key in categoryMap) {
    categoryText += `${key}: $${categoryMap[key].toLocaleString()} 元; `;
  }

  sheet.toast("🤖 Gemini AI 管家正在審閱您的帳本...", "系統提示", 3);

  const geminiAdvice = askGeminiAdvisor(totalExpense, remaining, usePercent, categoryText);

  let titlePrefix = "✨ 【安全進度】";
  if (usePercent >= 100) titlePrefix = "🚨 【破產紅色警報】";
  else if (usePercent >= 90) titlePrefix = "⚠️ 【橘色深度預警】";
  else if (usePercent >= 80) titlePrefix = "💡 【黃色警戒】";

  ui.alert(titlePrefix + '智慧理財管家診斷報告', geminiAdvice, ui.ButtonSet.OK);
}


/**
 * 🤖 【功能 10：Gemini AI 智慧理財管家】
 * 徹底修復：精準穿透 Google API 結構，取出正牌的自訂 JSON 字串回傳給前端
 */

function askGeminiAdvisor(totalExpense, remaining, usePercent, categoryText) {
  // === 🔒 核心安全防護線：避免傳入空值 ===
  const safeTotalExpense = (typeof totalExpense === 'number' && !isNaN(totalExpense)) ? totalExpense : 0;
  const safeRemaining    = (typeof remaining === 'number' && !isNaN(remaining)) ? remaining : 0;
  const safeUsePercent   = (typeof usePercent === 'number' && !isNaN(usePercent)) ? usePercent : 0;
  const safeCategoryText = categoryText ? categoryText.toString().trim() : "目前尚無消費分類資料。";

  // === 📅 時間運算中心：計算現在日期與月底倒數 ===
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1; // 月份從 0 開始，故 +1
  const date = now.getDate();
  const todayStr = `${year}/${month}/${date}`;

  // 計算到月底還有幾天
  const lastDayOfMonth = new Date(year, month, 0).getDate(); // 取得當月最後一天是幾號
  const daysLeft = lastDayOfMonth - date;
  let timeLeftStr = "";
  if (daysLeft === 0) {
    timeLeftStr = "今天就是月底最後一天了！";
  } else {
    timeLeftStr = `距離月底還有 ${daysLeft} 天`;
  }

  // ⚠️ 這裡維持使用你原有的金鑰
  const apiKey = Gemini_API; 
  const apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;

  // 建立量身打造的 Prompt 語境（限制繁體中文與回傳格式）
const prompt = `
  你是一位說話直白、幽默帶點毒舌，但內心極度關心使用者的「智慧財務管家」。
  請針對我目前的當月財務狀況進行「一句話的點評」與「一條具體的改善建議」，並在開頭嚴格遵守時間格式輸出。

  【目前財務與時間數據】
  - 今天的日期：${todayStr}
  - 距離月底時間：${timeLeftStr}
  - 目前累積總花費：$${safeTotalExpense.toLocaleString()} 元
  - 剩餘可用額度：$${safeRemaining.toLocaleString()} 元
  - 預算使用率：${safeUsePercent.toFixed(1)}%
  - 帳本分類消費分佈：
  ${safeCategoryText}

  【回傳限制】
  1. 請根據「剩餘可用額度」與「距離月底天數」的比例進行合理評估（例如：若剩下沒幾天但額度還很多，可適度放寬；若剛到月初預算就快爆了，請大力譴責）。
  2. 如果預算使用率 > 90%，請火力全開毒舌，逼我立刻停止消費、月底吃土。
  3. 如果預算在 50% 以下的安全範圍，請給予肯定，但提醒不要得意忘形。
  4. 請直接回傳精簡的繁體中文，**嚴格遵守**以下格式（連標點符號與換行都要一致）：
     📅 管家時間手札：今天日期是 ${todayStr} (${timeLeftStr})
     ⚠️ 管家毒舌點評：(根據數據進行的一句話點評)
     💡 管家強烈建議：(具體的一條行動指南)
     🥶 管家的冷笑話：(依據預算使用率，說一個不超過2行的冷笑話)
  `;

  const payload = {
    "contents": [{ "parts": [{ "text": prompt }] }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true // 允許抓取 API 內部拋出的真實錯誤訊息
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    
    // 穿透 Google 原始大禮包結構，精準抓取文字
    if (json.candidates && json.candidates[0].content.parts[0].text) {
      return json.candidates[0].content.parts[0].text.trim();
    }
    
    // 如果被後端伺服器拒絕，直接白話展示原因
    if (json.error) {
      return `❌ Gemini 拒絕連線，錯誤原因：\n[代碼 ${json.error.code}] ${json.error.message}`;
    }
    
    return "🤖 伺服器回傳結構異常，請稍後再試。";
    
  } catch (e) {
    return "🤖 請求發送失敗: " + e.message;
  }
}
// 主控程式：負責抓取資料並呼叫管家
function runFinancialAdvisor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // ─── 1. 自動抓取試算表數據（請根據您實際的儲存格位置修改） ───
  // 假設：總花費在 B12, 剩餘額度在 B13, 使用率在 B14 (0.85 代表 85%)
  // 分類文字分佈在 D2:E10
  const totalExpense = sheet.getRange("B12").getValue();
  const remaining    = sheet.getRange("B13").getValue();
  const usePercent   = sheet.getRange("B14").getValue() * 100; // 轉為百分比數字
  
  // 抓取分類資料並組合成文字
  const categoryRange = sheet.getRange("D2:E10").getValues();
  let categoryText = "";
  for (let i = 0; i < categoryRange.length; i++) {
    if (categoryRange[i][0] !== "") { // 忽略空行
      categoryText += `${categoryRange[i][0]}: $${categoryRange[i][1].toLocaleString()} 元\n`;
    }
  }

  // ─── 2. 呼叫 Gemini 顧問 ───
  const result = askGeminiAdvisor(totalExpense, remaining, usePercent, categoryText);
  
  // ─── 3. 檢查回傳結果是否失敗 ───
  if (result.includes("❌") || result.includes("🤖") || result.includes("失敗")) {
    
    // 彈出 UI 視窗詢問使用者是否重新連線
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "⚠️ 財務管家罷工中",
      `伺服器回報異常：\n\n${result}\n\n目前伺服器可能較為擁擠，是否要立刻重新嘗試連線？`,
      ui.ButtonSet.YES_NO
    );
    
    if (response == ui.Button.YES) {
      // 【核心功能】使用者點選「是」，立刻重新呼叫本函式（自我重試）
      runFinancialAdvisor(); 
    } else {
      // 使用者點選「否」，將錯誤訊息寫入儲存格（假設寫在 A1）
      sheet.getRange("A1").setValue("診斷失敗，使用者取消重試。");
    }
    
  } else {
    // 診斷成功，將精采的毒舌評論寫入您的目標儲存格（假設寫在 A1）
    sheet.getRange("A1").setValue(result);
  }
}

}
