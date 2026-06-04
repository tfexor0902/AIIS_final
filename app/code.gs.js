/**
 * 📥 【功能 1】接收從手機、網頁 傳過來的記帳資料 (POST 請求)
 * 當有外部系統把資料「推 (POST)」進來時，Google 會自動觸發這個函式。
 */
function doPost(e) {
  try {
    // 1. 安全檢查：確認是否有資料存入，如果沒有則報錯
    if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput("錯誤: 找不到 postData 內容").setMimeType(ContentService.MimeType.TEXT);
      /* return createResponse("錯誤: 找不到 postData 內容"); */
    }

    // 2. 解析資料：把傳過來的 JSON 字串，變成 JavaScript 看得懂的物件
    const params = JSON.parse(e.postData.contents);
    
    // 3. 設定預設值：如果對方沒傳分類就填"未分類"，沒傳金額就當作 0 元
    const { category = "未分類", name = "無名稱", amount = 0, date } = params;

// 4. 處理日期：先建立時間物件
const rawDate = date ? new Date(date) : new Date();

// 將時間物件格式化為「年/月/日」字串 (例如：2026/06/02)
const formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // 5. 連線到 Google 試算表：用網址中間那一串長長的 ID 來鎖定你的記帳本
    const spreadsheet = SpreadsheetApp.openById('你的ID');
    const sheet = spreadsheet.getSheetByName('記帳明細');

    // 防呆機制：萬一你把工作表改名了，系統找不到會主動警告，不會整台當掉
    if (!sheet) {
      return ContentService.createTextOutput("錯誤:找不到『記帳明細』工作表").setMimeType(ContentService.MimeType.TEXT);
  
       /* return createResponse("錯誤: 找不到『記帳明細』工作表"); */
    }

    // 6. 整理排版：把資料排成一陣列 (由左到右分別是：分類、名稱、金額、日期)
    const rowData = [category, name, Number(amount), formattedDate];

    // 7. 寫入記帳本：直接找到工作表最後一列的下一行，把這筆資料整列黏上去！
    sheet.appendRow(rowData);

    // 成功完成，回傳訊息給發送端
          return ContentService.createTextOutput("已新增").setMimeType(ContentService.MimeType.TEXT);

    // return createResponse("已新增");
    
  } catch (error) {
    // 萬一上面任何一個步驟出錯（例如試算表 ID 填錯），就會跑到這裡來捕捉錯誤訊息
          return ContentService.createTextOutput("發生錯誤").setMimeType(ContentService.MimeType.TEXT);

    //return createResponse(`發生錯誤: ${error.message}`);
  }
}


/**
 * 🌐 【功能 2】網頁部署進入點 (GET 請求)
 * 當你在瀏覽器輸入這支程式的網址時，會觸發這個函式，把網頁畫面秀出來。
 */
function doGet() {
  // 讀取名稱為 'Chart.html' 的檔案作為網頁樣板
  const template = HtmlService.createTemplateFromFile('Chart');
  
  // 渲染（執行）樣板並回傳，設定網頁標題，並允許這個網頁可以被內嵌在其他地方
  return template.evaluate()
      .setTitle('自動化記帳統計')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * 🛠️ 【功能 3】在試算表內部彈出儀表板視窗
 */
function showChartDialog() {
  // 一樣先讀取 'Chart.html'
  const template = HtmlService.createTemplateFromFile('Chart');
  
  // 設定彈出視窗的寬高與標題
  const htmlOutput = template.evaluate()
      .setWidth(780)  
      .setHeight(750) 
      .setTitle('我的記帳儀表板');
      
  // 叫 Google 試算表跳出這個前端視窗 (Modal 彈出視窗)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '我的記帳儀表板');
}


/**
 * 🧠 【功能 4】網頁元件拼盤 (組合 CSS 與 JS 檔案)
 * 網頁寫作通常會把 CSS (外觀) 和 JS (動作) 分開寫。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * 📊 【功能 5】後台大腦：撈出記帳本所有資料，統計成圖表要用的數據
 * 這個函式會算好「分類比例」、「每日趨勢」、「星期幾花最多」
 */
function getChartData() {
  const spreadsheet = SpreadsheetApp.openById('你的ID');
  const sheet = spreadsheet.getSheetByName('記帳明細');
  
  // 防呆：如果找不到記帳明細表，就回傳一堆空的預設格式
  if (!sheet) {
    return { categoryData: [], dateData: [], weekData: [], rawData: [], budgetInfo: { budget: 0, totalExpense: 0, usePercent: 0 } };
  }
  
  // 抓取 F2 儲存格裡面的「每月預算金額」
  const budget = parseFloat(sheet.getRange("F2").getValue()) || 0;
  
  // 抓取整個工作表「有寫字的所有範圍」
  const data = sheet.getDataRange().getValues();
  
  // 如果裡面只有第一行（標題列），代表還沒開始記帳，直接回傳基本預算資訊
  if (data.length <= 1) {
    return { categoryData: [], dateData: [], weekData: [], rawData: [], budgetInfo: { budget: budget, totalExpense: 0, usePercent: 0 } }; 
  }

  // 準備用來裝統計結果的空白陣列
  const categorySummary = {}; // 放各分類的總金額，例如：{ "餐飲": 500, "交通": 200 }
  const dateSummary = {};     // 放各日期的總金額，例如：{ "2026/06/01": 700 }
  let totalExpense = 0;       // 累積總花費
  
  // 用來計算星期幾平均花多少的對照表
  const weekNames = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
  const weekSummary = { "星期一": 0, "星期二": 0, "星期三": 0, "星期四": 0, "星期五": 0, "星期六": 0, "星期日": 0 };
  
  // 用來記錄「某個星期幾，總共出現過哪幾天」，好用來算平均值（排除重複日期）
  const weekDaysMap = { "星期一": {}, "星期二": {}, "星期三": {}, "星期四": {}, "星期五": {}, "星期六": {}, "星期日": {} };
  
  // 建立一個乾淨的原始資料陣列，先把第一行標題（分類、名稱、金額、日期）塞進去
  let serializedRawData = [];
  serializedRawData.push([data[0][0], data[0][1], data[0][2], data[0][3]]);

  // ==================== 迴圈開始：逐行掃描記帳本 ====================
  // 陣列從 i = 1 開始（代表第二列），因為第一列是標題
  for (let i = 1; i < data.length; i++) {
    const category = data[i][0];  // A欄：分類
    const amountStr = data[i][2]; // C欄：金額
    let dateVal = data[i][3];     // D欄：日期
    
    // 把金額轉成數字，萬一空格或填錯就當作 0 元
    const amount = parseFloat(amountStr) || 0; 
    if (amount <= 0) continue; // 如果金額小於等於 0 元（收入或打錯），就跳過這一行不統計
    
    totalExpense += amount; // 加上去，累積總支出

    // 處理分類名稱（去掉空格），如果沒填字就叫它 "未分類"
    const categoryName = (category && category.toString().trim()) ? category.toString().trim() : "未分類";
    // 累加該分類的金額
    categorySummary[categoryName] = (categorySummary[categoryName] || 0) + amount;

    let dateStr = "";
    let dayIndex = -1; // -1 代表星期幾還不知道

    // 【日期格式轉換防呆】因為 Google 試算表裡的日期可能是 Date 物件，也可能是純字串
    if (dateVal instanceof Date) {
      // 如果是標準日期物件，轉換成 yyyy/MM/dd 格式的字串
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      dayIndex = dateVal.getDay(); // 抓出這天是星期幾 (0=星期日, 1=星期一...)
    } else if (dateVal) {
      // 如果是字串，先把橫線 - 換成斜線 /，再試著讓程式去辨識它
      dateStr = dateVal.toString().split("T")[0].replace(/-/g, '/');
      let parsedDate = new Date(dateStr);
      if (!isNaN(parsedDate.getTime())) {
        dayIndex = parsedDate.getDay();
      }
    }

    // 如果成功辨識出日期與星期幾，就把它們分類記錄到陣列裡
    if (dateStr && dayIndex !== -1) {
      dateSummary[dateStr] = (dateSummary[dateStr] || 0) + amount; // 累加當天總金額
      
      const wName = weekNames[dayIndex]; // 換算成中文 "星期幾"
      weekSummary[wName] += amount;      // 累加這個星期幾的總金額
      weekDaysMap[wName][dateStr] = true; // 蓋個章，代表這個星期幾包含這一天（去重用）
    }
    
    // 把整理乾淨的資料放進原始資料清單，準備等等傳給網頁的表格看
    serializedRawData.push([
      categoryName,
      data[i][1] ? data[i][1].toString() : "",
      amount,
      dateStr || "—"
    ]);
  }
  // ==================== 迴圈結束 ====================
  
  // 轉換成格式：[['欄位名稱', '欄位名稱'], [資料, 資料]...]
  // 1. 分類圖表資料
  const categoryData = [['類別', '總支出']];
  for (let key in categorySummary) { categoryData.push([key, categorySummary[key]]); }
  
  // 2. 日期圖表資料（依照時間軸從舊到新排序）
  const sortedDates = Object.keys(dateSummary).sort((a, b) => new Date(a) - new Date(b));
  const dateData = [['日期', '每日總支出']];
  sortedDates.forEach(date => { dateData.push([date, dateSummary[date]]); });
  
  // 3. 星期平均圖表資料
  const weekOrder = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
  const weekData = [['星期', '平均支出金額']];
  weekOrder.forEach(day => {
    // 算出這個星期幾在帳本中總共出現過幾天（例如總共有 3 個星期一）
    const daysCount = Object.keys(weekDaysMap[day]).length || 1; 
    // 總金額除以天數 = 平均花費（並四捨五入到小數點後兩位）
    const averageAmount = Math.round((weekSummary[day] / daysCount) * 100) / 100;
    weekData.push([day, averageAmount]);
  });
  
  // 4. 計算預算使用率百分比 (總花費 / 預算 * 100)
  const usePercent = budget > 0 ? (totalExpense / budget) * 100 : 0;
  
  // 最後打包，整包丟回給網頁前端去畫圖
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

function renderDateChart() {
        const startVal = document.getElementById('dateStart') ? document.getElementById('dateStart').value : '';
        const endVal   = document.getElementById('dateEnd')   ? document.getElementById('dateEnd').value   : '';

        let sourceData = chartDataPayload.dateData;
        let filtered = sourceData.slice(1);

        if (startVal) {
          filtered = filtered.filter(r => r[0].replace(/\//g, '-') >= startVal);
        }
        if (endVal) {
          filtered = filtered.filter(r => r[0].replace(/\//g, '-') <= endVal);
        }

        const infoEl = document.getElementById('dateRangeInfo');
        if ((startVal || endVal) && infoEl) {
          const subtotal = filtered.reduce((s, r) => s + r[1], 0);
          infoEl.style.display = 'block';
          infoEl.innerHTML = '已篩選 <strong>' + filtered.length + '</strong> 天，區間花費小計：<strong>$' + subtotal.toLocaleString() + '</strong>';
        } else if (infoEl) {
          infoEl.style.display = 'none';
        }

        const drawData = [sourceData[0], ...filtered];
        if (drawData.length <= 1) {
          document.getElementById('linechart_div').innerHTML = "<div class='no-data'>此區間沒有消費紀錄。</div>";
          return;
        }

        const data = google.visualization.arrayToDataTable(drawData);
        const options = {
          title: '每日花費趨勢 (點擊藍色節點可查看當日消費明細)',
          curveType: 'function',
          pointsVisible: true,
          chartArea: { width: '85%', height: '70%' }
        };
        const chart = new google.visualization.LineChart(document.getElementById('linechart_div'));
        google.visualization.events.addListener(chart, 'select', function() {
          const selectedItem = chart.getSelection()[0];
          if (selectedItem && selectedItem.row !== null && selectedItem.row !== undefined) {
            const dateStr = data.getValue(selectedItem.row, 0);
            showDateDetails(dateStr);
          }
        });
        chart.draw(data, options);
      }


/**
 * 🔍 【功能 6】互動功能：點選圖表上的「某個日期」，撈出當天的所有消費明細
 */
function getFilteredDateData(targetDateStr) {
  try {
    const spreadsheet = SpreadsheetApp.openById('你的ID');
    const sheet = spreadsheet.getSheetByName('記帳明細');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let filteredRows = [];
    const targetFormatted = targetDateStr.trim().replace(/-/g, '/'); // 統一格式

    // 巡邏整張表，只要日期跟被點選的日期一模一樣，就抓出來
    for (let i = 1; i < data.length; i++) {
      let dateVal = data[i][3];
      let rowDateStr = "";
      
      if (dateVal instanceof Date) {
        rowDateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else {
        rowDateStr = dateVal ? dateVal.toString().split("T")[0].replace(/-/g, '/') : "";
      }
      
      // 比對成功，塞進結果清單
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


/**
 * 🔍 【功能 7】互動功能：點選圖表上的「星期幾」，撈出所有屬於該星期幾的消費明細
 */
function getFilteredWeekData(targetWeekName) {
  try {
    const spreadsheet = SpreadsheetApp.openById('你的ID');
    const sheet = spreadsheet.getSheetByName('記帳明細');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let filteredRows = [];
    const weekNames = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
    const targetFormatted = targetWeekName.trim();

    // 巡邏整張表，計算每一行的日期是星期幾，吻合就抓出來
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
      
      // 如果這行資料的星期幾等於點選的星期幾，就打包
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


/**
 * 🔔 【功能 8】打開記帳本時的「自動自動歡迎與理財警告選單」
 * 只要你開瀏覽器打開這個 Google 試算表檔案，這段程式就會自動執行。
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // 在試算表上方工具列，增加一個自訂按鈕選單：『記帳功能』->『查看記帳儀表板』
  ui.createMenu('記帳功能')
    .addItem('查看記帳儀表板', 'showChartDialog')
    .addToUi();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('記帳明細');
  if (!sheet) return;

  // 自動在 F1 儲存格幫你寫上標題，省去手動打字的麻煩
  if (sheet.getRange("F1").getValue() !== "每月預算設定") {
    sheet.getRange("F1").setValue("每月預算設定");
  }

  // 檢查 F2 有沒有填預算，沒填的話就跳出框框逼你輸入
  const currentBudget = sheet.getRange("F2").getValue();
  if (!currentBudget || currentBudget === "") {
    const response = ui.prompt('💰 首次預算設定', '尚未設定本月預算，請輸入金額：', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      const inputText = response.getResponseText().trim();
      const newBudget = parseFloat(inputText);
      if (!isNaN(newBudget) && newBudget > 0) {
        sheet.getRange("F2").setValue(newBudget); // 寫入 F2 預算格
      } else {
        ui.alert('設定失敗', '金額格式不正確，請開啟儀表板再行設定。', ui.ButtonSet.OK);
      }
    }
    return;
  }

  const activeBudget = parseFloat(currentBudget);
  if (activeBudget <= 0) return;

  // 計算目前累積總花費，用來做接下來的彈出警告
  const data = sheet.getDataRange().getValues();
  let totalExpense = 0;
  for (let i = 1; i < data.length; i++) {
    const amount = parseFloat(data[i][2]) || 0;
    if (amount > 0) totalExpense += amount;
  }

  const usePercent = (totalExpense / activeBudget) * 100;
  const remaining = activeBudget - totalExpense;

  // 【貼心理財管家】根據你的花費百分比，彈出不同趣味程度的警告視窗
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


/**
 * ✏️ 【功能 9】修改預算
 * 當你在前端網頁修改預算金額並按下儲存時，網頁會遠端呼叫這個函式，把新數字寫回 F2
 */
function updateBudget(newBudget) {
  try {
    const amount = parseFloat(newBudget);
    // 檢查是不是打錯字或輸入負數
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: '金額格式不正確' };
    }
    const spreadsheet = SpreadsheetApp.openById('你的ID');
    const sheet = spreadsheet.getSheetByName('記帳明細');
    
    // 把新預算直接蓋掉舊的 F2 儲存格
    sheet.getRange("F2").setValue(amount);
    
    return { success: true, budget: amount }; // 回傳成功訊息給網頁
  } catch (e) {
    return { success: false, message: e.message };
  }
}
