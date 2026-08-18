/**
 * 個人視覺化工作儀表板 Web App (v4.7：獨立【⚠️ 逾期任務】頁籤 + 修復週期性任務展延)
 */

const DEFAULT_COLORS = ['#0d6efd', '#198754', '#6f42c1', '#fd7e14', '#0dcaf0', '#d63384'];

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('📱 個人視覺化工作儀表板')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📱 工作儀表板 App')
    .addItem('🚀 1-Click 一鍵最佳化分頁與公式', 'setupDashboardTabs')
    .addItem('🔑 授權並測試「主要 Google 日曆」', 'testCalendarAccess')
    .addItem('🌐 授權並測試「外部 iCal 日曆」', 'testExternalCalendarAccess')
    .addToUi();
}

function forceUrlFetchPermission() {
  const test = UrlFetchApp.fetch("https://www.google.com");
  Logger.log("✅ 外部連線權限授權成功！HTTP 狀態碼：" + test.getResponseCode());
}

function testCalendarAccess() {
  try {
    const cal = CalendarApp.getDefaultCalendar();
    const count = cal.getEventsForDay(new Date()).length;
    SpreadsheetApp.getUi().alert(`✅ 主要日曆連線成功！今日共有 ${count} 筆行程。`);
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ 主要日曆連線提示：' + err.toString());
  }
}

function testExternalCalendarAccess() {
  try {
    const ss = getSpreadsheet();
    const setSheet = ss.getSheetByName('⚙️ 系統設定');
    if (!setSheet || setSheet.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('請先在【⚙️ 系統設定】分頁填入外部日曆 iCal 私人網址！');
      return;
    }

    const icalUrls = setSheet.getRange(2, 4, setSheet.getLastRow() - 1, 2).getDisplayValues();
    let totalExtEvents = 0;
    let urlCount = 0;

    icalUrls.forEach(row => {
      const [urlA, urlB] = row.map(val => String(val || "").trim());
      const url = urlA || urlB || "";
      if (url.startsWith('http')) {
        urlCount++;
        const events = fetchExternalIcsEvents(url, new Date());
        totalExtEvents += events.length;
      }
    });

    if (urlCount === 0) {
      SpreadsheetApp.getUi().alert('未找到有效的 iCal 網址（需為 https:// 開頭）。');
    } else {
      SpreadsheetApp.getUi().alert(`✅ 外部日曆連線測試成功！\n已成功解析 ${urlCount} 個外部日曆，讀取到 ${totalExtEvents} 筆今日行程（含全天、早晚與循環事件）。`);
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ 外部日曆連線失敗：' + err.toString());
  }
}

/**
 * 1-Click 一鍵最佳化試算表分頁與公式
 */
function setupDashboardTabs() {
  const ss = getSpreadsheet();
  const tabNames = ['📊 總覽儀表板', '📋 專案檢核表', '⚙️ 系統設定'];
  
  tabNames.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // 1. ⚙️ 系統設定
  const setSheet = ss.getSheetByName('⚙️ 系統設定');
  const existingIcalUrl = setSheet.getRange('E2').getValue() || setSheet.getRange('D2').getValue() || '';
  
  setSheet.clear();
  setSheet.getRange('A1:F1').setValues([['狀態', '優先級', '來源', '其他日曆 iCal 網址', '專案名稱', '專案代表色']]).setFontWeight('bold').setBackground('#E8EAED');
  setSheet.getRange('A2:A5').setValues([['未開始'], ['進行中'], ['已完成'], ['已延期']]);
  setSheet.getRange('B2:B4').setValues([['高'], ['中'], ['低']]);
  setSheet.getRange('C2:C4').setValues([['Google 日曆'], ['助理排程'], ['手動新增']]);

  if (existingIcalUrl) {
    setSheet.getRange('D2').setValue(existingIcalUrl);
  }

  // 預設範例專案
  setSheet.getRange('E2:F2').setValues([
    ['範例專案A', '#0d6efd']
  ]);

  const statusRule = SpreadsheetApp.newDataValidation().requireValueInRange(setSheet.getRange('A2:A5')).build();
  const priorityRule = SpreadsheetApp.newDataValidation().requireValueInRange(setSheet.getRange('B2:B4')).build();

  // 2. 📋 專案檢核表
  const taskSheet = ss.getSheetByName('📋 專案檢核表');
  if (taskSheet.getLastRow() === 0) {
    taskSheet.getRange('A1:L1').setValues([
      ['任務ID', '專案名稱', '子項目/階段', '任務內容', '負責類別', '優先級', '開始日期', '截止日期', '狀態', '進度/勾選', '備註', '重複週期']
    ]).setFontWeight('bold').setBackground('#E8F0FE');
    taskSheet.getRange('F2:F500').setDataValidation(priorityRule);
    taskSheet.getRange('I2:I500').setDataValidation(statusRule);
  } else {
    if (!taskSheet.getRange('L1').getValue()) {
      taskSheet.getRange('L1').setValue('重複週期').setFontWeight('bold').setBackground('#E8F0FE');
    }
  }

  // 3. 📊 總覽儀表板
  const dashSheet = ss.getSheetByName('📊 總覽儀表板');
  dashSheet.clear();
  dashSheet.getRange('A1:D1').setValues([['總專案數', '未完成任務數', '今日待辦數', '整體完成率']]).setFontWeight('bold').setBackground('#FCE8E6');
  
  dashSheet.getRange('A2').setFormula("=IFERROR(COUNTA(UNIQUE(FILTER('📋 專案檢核表'!B2:B, '📋 專案檢核表'!B2:B<>\"\" ))), 0)");
  dashSheet.getRange('B2').setFormula("=IFERROR(COUNTIFS('📋 專案檢核表'!B2:B, \"<>\"; '📋 專案檢核表'!I2:I, \"<>已完成\"), 0)");
  dashSheet.getRange('C2').setFormula("=IFERROR(COUNTIFS('📋 專案檢核表'!H2:H, TODAY(), '📋 專案檢核表'!I2:I, \"<>已完成\"), 0)");
  dashSheet.getRange('D2').setFormula("=IFERROR(COUNTIF('📋 專案檢核表'!I2:I, \"已完成\") / COUNTIF('📋 專案檢核表'!B2:B, \"<>\"), 0)").setNumberFormat('0.0%');

  dashSheet.getRange('A5:F5').setValues([['專案名稱', '任務總數', '已完成數', '未完成數', '完成率', '狀態']]).setFontWeight('bold').setBackground('#F1F3F4');
  dashSheet.getRange('A6').setFormula("=IFERROR(UNIQUE(FILTER('📋 專案檢核表'!B2:B, '📋 專案檢核表'!B2:B<>\"\" )), \"\")");
  dashSheet.getRange('B6').setFormula("=ARRAYFORMULA(IF(A6:A=\"\", \"\", COUNTIF('📋 專案檢核表'!B:B, A6:A)))");
  dashSheet.getRange('C6').setFormula("=ARRAYFORMULA(IF(A6:A=\"\", \"\", COUNTIFS('📋 專案檢核表'!B:B, A6:A, '📋 專案檢核表'!I:I, \"已完成\")))");
  dashSheet.getRange('D6').setFormula("=ARRAYFORMULA(IF(A6:A=\"\", \"\", B6:B-C6:C))");
  dashSheet.getRange('E6').setFormula("=ARRAYFORMULA(IF(A6:A=\"\", \"\", IFERROR(C6:C/B6:B, 0)))").setNumberFormat('0.0%');
  dashSheet.getRange('F6').setFormula("=ARRAYFORMULA(IF(A6:A=\"\", \"\", IF(E6:E=1, \"已完成\", \"進行中\")))");

  const redundantTabs = ['工作表1', '📅 今日待辦與排程'];
  redundantTabs.forEach(tName => {
    const rSheet = ss.getSheetByName(tName);
    if (rSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(rSheet); } catch(e) {}
    }
  });

  if (SpreadsheetApp.getUi()) {
    SpreadsheetApp.getUi().alert('🎉 試算表結構與公式已初始化完成！');
  }
}

/**
 * 計算有效截止日期字串 (支援 YYYY-MM 月底判定、YYYY 年底判定、YYYY-MM-DD)
 */
function getEffectiveDueDateStr(dueDateStr) {
  if (!dueDateStr || typeof dueDateStr !== 'string') return '';
  const clean = dueDateStr.trim().replace(/\//g, '-').replace(/\./g, '-');
  if (!clean) return '';

  const parts = clean.split('-').map(p => p.trim()).filter(Boolean);
  if (parts.length === 1) {
    const [yStr] = parts;
    if (/^\d{4}$/.test(yStr)) {
      const y = parseInt(yStr, 10);
      return `${y}-12-31`;
    }
  } else if (parts.length === 2) {
    const [yStr, mStr] = parts;
    if (/^\d{4}$/.test(yStr) && /^\d{1,2}$/.test(mStr)) {
      const y = parseInt(yStr, 10);
      const m = parseInt(mStr, 10);
      if (m >= 1 && m <= 12) {
        const lastDay = new Date(y, m, 0).getDate();
        const mm = String(m).padStart(2, '0');
        const dd = String(lastDay).padStart(2, '0');
        return `${y}-${mm}-${dd}`;
      }
    }
  } else if (parts.length === 3) {
    const [yStr, mStr, dStr] = parts;
    if (/^\d{4}$/.test(yStr) && /^\d{1,2}$/.test(mStr) && /^\d{1,2}$/.test(dStr)) {
      const y = parseInt(yStr, 10);
      const m = parseInt(mStr, 10);
      const d = parseInt(dStr, 10);
      const mm = String(m).padStart(2, '0');
      const dd = String(d).padStart(2, '0');
      return `${y}-${mm}-${dd}`;
    }
  }

  const d = new Date(dueDateStr);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return clean;
}

function calculatePriority(dueDateStr) {
  const effDateStr = getEffectiveDueDateStr(dueDateStr);
  if (!effDateStr) return '低';
  
  const parts = effDateStr.split('-');
  if (parts.length !== 3) return '低';
  
  const [yStr, mStr, dStr] = parts;
  const y = parseInt(yStr, 10);
  const m = parseInt(mStr, 10);
  const d = parseInt(dStr, 10);
  
  const due = new Date(y, m - 1, d);
  if (isNaN(due.getTime())) return '低';
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  
  if (diffDays <= 7) return '高';
  return '中';
}

/**
 * 計算下一期重複任務的截止日 (完整修復：防陣列解析錯誤與年份溢出)
 */
function calculateNextDueDate(baseDueDateStr, recurrence) {
  let baseDate = new Date();
  if (baseDueDateStr && baseDueDateStr.trim() !== '') {
    const eff = getEffectiveDueDateStr(baseDueDateStr);
    const parts = eff.split('-');
    if (parts.length === 3) {
      const [yStr, mStr, dStr] = parts;
      const y = parseInt(yStr, 10);
      const m = parseInt(mStr, 10);
      const d = parseInt(dStr, 10);
      baseDate = new Date(y, m - 1, d);
    }
  }

  const nextDate = new Date(baseDate.getTime());

  if (recurrence === '每週' || recurrence.includes('週')) {
    nextDate.setDate(nextDate.getDate() + 7);
  } else if (recurrence === '每月' || recurrence.includes('月')) {
    const origDay = baseDate.getDate();
    nextDate.setMonth(nextDate.getMonth() + 1);
    if (nextDate.getDate() !== origDay) {
      nextDate.setDate(0); // 避免 1/31 變 3/3，自動調整為該月最後一天
    }
  } else if (recurrence === '每年' || recurrence.includes('年')) {
    nextDate.setFullYear(nextDate.getFullYear() + 1);
  } else {
    return '';
  }

  const tz = Session.getScriptTimeZone() || "Asia/Taipei";
  return Utilities.formatDate(nextDate, tz, "yyyy-MM-dd");
}

function parseIcsDateString(dtStr) {
  if (!dtStr) return null;
  const clean = dtStr.replace(/[^0-9TZ]/g, '');
  if (clean.length < 8) return null;

  const y = parseInt(clean.substring(0, 4), 10);
  const m = parseInt(clean.substring(4, 6), 10) - 1;
  const d = parseInt(clean.substring(6, 8), 10);

  if (!clean.includes('T')) {
    return { isAllDay: true, date: new Date(y, m, d) };
  }
  const h = parseInt(clean.substring(9, 11), 10) || 0;
  const min = parseInt(clean.substring(11, 13), 10) || 0;
  const s = parseInt(clean.substring(13, 15), 10) || 0;

  if (clean.endsWith('Z')) {
    const utcMs = Date.UTC(y, m, d, h, min, s);
    return { isAllDay: false, date: new Date(utcMs) };
  } else {
    return { isAllDay: false, date: new Date(y, m, d) };
  }
}

function isIcsEventOccurring(block, targetDate, tz) {
  const targetFormatted = Utilities.formatDate(targetDate, tz, 'yyyyMMdd');
  const targetM = targetDate.getMonth() + 1;
  const targetD = targetDate.getDate();
  const targetDayOfWeek = targetDate.getDay();

  const dtstartExec = /DTSTART(?:;[^:]+)?:([0-9TZ]+)/.exec(block);
  if (!dtstartExec) return false;
  const dtstartRaw = dtstartExec.pop();
  
  const parsedStart = parseIcsDateString(dtstartRaw);
  if (!parsedStart || !parsedStart.date) return false;

  const sDate = parsedStart.date;
  const localStartDateFormatted = Utilities.formatDate(sDate, tz, 'yyyyMMdd');

  if (localStartDateFormatted === targetFormatted) return true;

  const startDateInt = parseInt(localStartDateFormatted, 10);
  const targetDateInt = parseInt(targetFormatted, 10);
  if (startDateInt > targetDateInt) return false;

  const dtendExec = /DTEND(?:;[^:]+)?:([0-9TZ]+)/.exec(block);
  if (dtendExec) {
    const parsedEnd = parseIcsDateString(dtendExec.pop());
    if (parsedEnd && parsedEnd.date) {
      const localEndDateFormatted = Utilities.formatDate(parsedEnd.date, tz, 'yyyyMMdd');
      const endDateInt = parseInt(localEndDateFormatted, 10);
      if (startDateInt <= targetDateInt && targetDateInt < endDateInt) {
        return true;
      }
    }
  }

  const rruleExec = /RRULE:(.*)/.exec(block);
  if (rruleExec) {
    const rrule = rruleExec.pop();

    const untilExec = /UNTIL=([0-9TZ]+)/.exec(rrule);
    if (untilExec) {
      const untilDateInt = parseInt(untilExec.pop().substring(0, 8), 10);
      if (targetDateInt > untilDateInt) return false;
    }

    if (block.includes('EXDATE') && block.includes(targetFormatted)) {
      return false;
    }

    const startLocalD = parseInt(Utilities.formatDate(sDate, tz, 'dd'), 10);
    const startLocalM = parseInt(Utilities.formatDate(sDate, tz, 'MM'), 10);

    if (rrule.includes('FREQ=MONTHLY')) {
      const byMonthDayExec = /BYMONTHDAY=([0-9,]+)/.exec(rrule);
      if (byMonthDayExec) {
        const days = byMonthDayExec.pop().split(',').map(n => parseInt(n, 10));
        if (days.includes(targetD)) return true;
      } else {
        if (startLocalD === targetD) return true;
      }
    }

    if (rrule.includes('FREQ=YEARLY')) {
      if (startLocalM === targetM && startLocalD === targetD) return true;
    }

    if (rrule.includes('FREQ=WEEKLY')) {
      const dayMap = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
      const byDayExec = /BYDAY=([A-Z,]+)/.exec(rrule);
      if (byDayExec) {
        const byDays = byDayExec.pop().split(',');
        const targetDayKey = Object.keys(dayMap).find(k => dayMap[k] === targetDayOfWeek);
        if (byDays.includes(targetDayKey)) return true;
      } else {
        if (sDate.getDay() === targetDayOfWeek) return true;
      }
    }

    if (rrule.includes('FREQ=DAILY')) {
      return true;
    }
  }

  return false;
}

function fetchExternalIcsEvents(icalUrl, targetDate) {
  const events = [];
  if (!icalUrl || !icalUrl.startsWith('http')) return events;
  
  try {
    const res = UrlFetchApp.fetch(icalUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const icsText = res.getContentText();
      const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
      const blocks = icsText.split('BEGIN:VEVENT');
      
      blocks.slice(1).forEach(block => {
        if (isIcsEventOccurring(block, targetDate, tz)) {
          const summaryExec = /SUMMARY:(.*)/.exec(block);
          const summaryRaw = summaryExec ? summaryExec.pop() : '未命名行程';
          const summary = summaryRaw.replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\n/g, ' ').trim();

          const locExec = /LOCATION:(.*)/.exec(block);
          const locRaw = locExec ? locExec.pop() : '';
          const location = locRaw.replace(/\\,/g, ',').replace(/\\;/g, ';').trim();

          const dtstartExec = /DTSTART(?:;[^:]+)?:([0-9TZ]+)/.exec(block);
          const dtstartRaw = dtstartExec ? dtstartExec.pop() : '';
          const dtendExec = /DTEND(?:;[^:]+)?:([0-9TZ]+)/.exec(block);
          const dtendRaw = dtendExec ? dtendExec.pop() : '';

          let timeStr = '全天';
          let startMin = -1;
          let endMin = 1440;
          let isAllDay = true;

          if (dtstartRaw) {
            const parsedStart = parseIcsDateString(dtstartRaw);
            if (parsedStart && !parsedStart.isAllDay && parsedStart.date) {
              isAllDay = false;
              const sDate = parsedStart.date;
              const startFormatted = Utilities.formatDate(sDate, tz, 'HH:mm');
              
              const startH = parseInt(Utilities.formatDate(sDate, tz, 'HH'), 10);
              const startM = parseInt(Utilities.formatDate(sDate, tz, 'mm'), 10);
              startMin = startH * 60 + startM;

              let endFormatted = '';
              const parsedEnd = parseIcsDateString(dtendRaw);
              if (parsedEnd && !parsedEnd.isAllDay && parsedEnd.date) {
                const eDate = parsedEnd.date;
                endFormatted = ' - ' + Utilities.formatDate(eDate, tz, 'HH:mm');
                const endH = parseInt(Utilities.formatDate(eDate, tz, 'HH'), 10);
                const endM = parseInt(Utilities.formatDate(eDate, tz, 'mm'), 10);
                endMin = endH * 60 + endM;
              } else {
                endMin = startMin + 60;
              }
              timeStr = startFormatted + endFormatted;
            }
          }

          events.push({
            time: String(timeStr),
            title: String(summary || '外部日曆行程'),
            location: String(location),
            source: '外部日曆',
            isAllDay: isAllDay,
            startMin: startMin,
            endMin: endMin
          });
        }
      });
    }
  } catch (e) {
    console.error('iCal error:', e);
  }
  return events;
}

function getWebAppDataJson() {
  const result = {
    summary: { totalProjects: 0, pendingTasks: 0, calCount: 0, overallRate: 0 },
    calEvents: [],
    todayTasks: [],
    overdueTasks: [],
    upcomingTasks: [],
    completedTasks: [],
    projects: [],
    existingProjects: [],
    error: ""
  };

  const tz = Session.getScriptTimeZone() || "Asia/Taipei";
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  // 1. Google 主要日曆
  try {
    const cal = CalendarApp.getDefaultCalendar();
    if (cal) {
      const events = cal.getEventsForDay(now);
      result.calEvents = events.map(e => {
        let timeStr = "全天";
        let startMin = -1;
        let endMin = 1440;
        const isAllDay = e.isAllDayEvent();

        if (!isAllDay) {
          const s = e.getStartTime();
          const end = e.getEndTime();
          timeStr = Utilities.formatDate(s, tz, "HH:mm") + " - " + Utilities.formatDate(end, tz, "HH:mm");
          startMin = s.getHours() * 60 + s.getMinutes();
          endMin = end.getHours() * 60 + end.getMinutes();
        }

        return {
          time: String(timeStr),
          title: String(e.getTitle() || "未命名行程"),
          location: String(e.getLocation() || ""),
          source: "主要日曆",
          isAllDay: isAllDay,
          startMin: startMin,
          endMin: endMin
        };
      });
    }
  } catch (err) {
    result.error = "主要日曆存取提示：" + err.toString();
  }

  // 2. 外部日曆 (iCal)
  try {
    const ss = getSpreadsheet();
    const setSheet = ss.getSheetByName("⚙️ 系統設定");
    if (setSheet && setSheet.getLastRow() >= 2) {
      const icalUrls = setSheet.getRange(2, 4, setSheet.getLastRow() - 1, 2).getDisplayValues();
      icalUrls.forEach(urlRow => {
        const [urlA, urlB] = urlRow.map(val => String(val || "").trim());
        const url = urlA || urlB || "";
        if (url && url.startsWith("http")) {
          const extEvents = fetchExternalIcsEvents(url, now);
          result.calEvents = result.calEvents.concat(extEvents);
        }
      });
    }
  } catch (err) {
    console.error('External ical error:', err);
  }

  result.calEvents.sort((a, b) => a.startMin - b.startMin);
  result.summary.calCount = result.calEvents.length;

  // 3. 專案代表色 (使用解構賦值安全讀取)
  const colorMap = {};
  try {
    const ss = getSpreadsheet();
    const setSheet = ss.getSheetByName("⚙️ 系統設定");
    if (setSheet && setSheet.getLastRow() >= 2) {
      const colorRows = setSheet.getRange(2, 5, setSheet.getLastRow() - 1, 2).getDisplayValues();
      colorRows.forEach(cr => {
        const [pName, pColor] = cr.map(val => String(val || "").trim());
        if (pName && pColor) colorMap[pName] = pColor;
      });
    }
  } catch (err) {}

  // 4. 讀取專案檢核表
  try {
    const ss = getSpreadsheet();
    const taskSheet = ss.getSheetByName("📋 專案檢核表");
    
    if (taskSheet && taskSheet.getLastRow() > 1) {
      const maxCols = Math.max(12, taskSheet.getLastColumn());
      const displayData = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, maxCols).getDisplayValues();
      
      let allTasks = [];
      const projectMap = {};
      let autoColorIndex = 0;

      displayData.forEach((rawRow, i) => {
        const cleanedRow = rawRow.map(val => String(val || "").trim());
        const [cTaskId, cProject, cSubItem, cTitle, cCategory, cPriority, cStart, cDueDate, cStatus, cCheck, cRemarks, cRecurrence] = cleanedRow;

        if (cProject && cTitle && cProject !== "FALSE" && cTitle !== "FALSE") {
          const autoPriority = calculatePriority(cDueDate);
          const isCompleted = (cStatus === "已完成");

          if (!colorMap[cProject]) {
            colorMap[cProject] = DEFAULT_COLORS[autoColorIndex % DEFAULT_COLORS.length];
            autoColorIndex++;
          }

          const effDueDate = getEffectiveDueDateStr(cDueDate);
          const taskObj = {
            rowIndex: i + 2,
            id: cTaskId || ('T' + (i + 1)),
            project: cProject,
            projectColor: colorMap[cProject],
            subItem: cSubItem || "一般",
            title: cTitle,
            category: cCategory || "通用",
            priority: autoPriority,
            dueDate: cDueDate || "",
            effectiveDueDate: effDueDate,
            status: cStatus || "未開始",
            isDone: isCompleted,
            recurrence: cRecurrence || "無",
            remarks: cRemarks || ""
          };
          allTasks.push(taskObj);

          if (!projectMap[cProject]) {
            projectMap[cProject] = {
              name: cProject,
              color: colorMap[cProject],
              total: 0,
              done: 0,
              pending: 0
            };
          }
          projectMap[cProject].total += 1;
          if (isCompleted) {
            projectMap[cProject].done += 1;
          } else {
            projectMap[cProject].pending += 1;
          }
        }
      });

      result.existingProjects = Object.keys(projectMap).map(name => ({
        name: name,
        color: projectMap[name].color
      }));

      const totalTaskCount = allTasks.length;
      const completedTaskCount = allTasks.filter(t => t.isDone).length;
      const pendingTaskCount = allTasks.filter(t => !t.isDone).length;
      
      result.summary.totalProjects = result.existingProjects.length;
      result.summary.pendingTasks = pendingTaskCount;
      result.summary.overallRate = totalTaskCount > 0 ? Math.round((completedTaskCount / totalTaskCount) * 100) : 0;

      result.projects = Object.values(projectMap).map(p => {
        const rate = p.total > 0 ? Math.round((p.done / p.total) * 100) : 0;
        return {
          name: p.name,
          color: p.color,
          total: p.total,
          done: p.done,
          pending: p.pending,
          rate: rate,
          status: p.done === p.total && p.total > 0 ? "已完成" : "進行中"
        };
      });

      // 方案 C：分流為 今日待辦 / ⚠️ 逾期任務 / 未來待辦 / 已完成 (含 YYYY-MM 月底防逾期判定)
      result.todayTasks = allTasks.filter(t => !t.isDone && (t.dueDate === todayStr || t.effectiveDueDate === todayStr));
      result.overdueTasks = allTasks.filter(t => !t.isDone && t.effectiveDueDate && t.effectiveDueDate < todayStr).sort((a, b) => {
        return new Date(a.effectiveDueDate) - new Date(b.effectiveDueDate);
      });
      result.upcomingTasks = allTasks.filter(t => !t.isDone && (!t.effectiveDueDate || t.effectiveDueDate > todayStr)).sort((a, b) => {
        if (!a.effectiveDueDate) return 1;
        if (!b.effectiveDueDate) return -1;
        return new Date(a.effectiveDueDate) - new Date(b.effectiveDueDate);
      });
      result.completedTasks = allTasks.filter(t => t.isDone);
    }
  } catch (err) {
    result.error = (result.error ? result.error + " | " : "") + "專案讀取提示：" + err.toString();
  }

  return JSON.stringify(result);
}

function addNewTaskFromWeb(form) {
  try {
    const ss = getSpreadsheet();
    let taskSheet = ss.getSheetByName('📋 專案檢核表');
    if (!taskSheet) {
      setupDashboardTabs();
      taskSheet = ss.getSheetByName('📋 專案檢核表');
    }

    const pName = form.projectName.trim();
    const id = 'T' + new Date().getTime().toString().slice(-6);
    const calculatedPriority = calculatePriority(form.dueDate);

    if (form.projectColor) {
      const setSheet = ss.getSheetByName('⚙️ 系統設定');
      if (setSheet) {
        setSheet.appendRow(['', '', '', '', pName, form.projectColor]);
      }
    }

    taskSheet.appendRow([
      id,
      pName,
      form.subItem || '一般',
      form.taskContent.trim(),
      '通用',
      calculatedPriority,
      new Date(),
      form.dueDate || '',
      '未開始',
      false,
      form.remarks || '',
      form.recurrence || '無'
    ]);

    return '✅ 任務已成功新增至【📋 專案檢核表】！';
  } catch (err) {
    return '❌ 新增失敗：' + err.toString();
  }
}

/**
 * 更新任務狀態並自動展延下一期重複性任務 (使用解構賦值，徹底解決亂碼與年份溢出)
 */
function toggleTaskStatus(rowIndex, newStatus) {
  try {
    const ss = getSpreadsheet();
    const taskSheet = ss.getSheetByName('📋 專案檢核表');
    if (!taskSheet) return '找不到專案檢核表';

    taskSheet.getRange(rowIndex, 9).setValue(newStatus);
    if (newStatus === '已完成') {
      taskSheet.getRange(rowIndex, 10).setValue(true);

      // 檢查是否為重複性任務 (使用陣列解構精確對應欄位)
      const rawRowValues = taskSheet.getRange(rowIndex, 1, 1, 12).getDisplayValues()[0];
      const [
        cTaskId,
        pName,
        subItem,
        taskTitle,
        category,
        cPriority,
        cStart,
        dueDateStr,
        cStatus,
        cCheck,
        remarks,
        rawRecurrence
      ] = rawRowValues.map(val => String(val || "").trim());

      const recurrence = rawRecurrence || '';

      if (recurrence && recurrence !== '無' && recurrence !== '不重複' && recurrence !== '') {
        const nextDueDate = calculateNextDueDate(dueDateStr, recurrence);
        if (nextDueDate) {
          const newId = 'T' + new Date().getTime().toString().slice(-6);
          const newPriority = calculatePriority(nextDueDate);

          taskSheet.appendRow([
            newId,
            pName,
            subItem || '一般',
            taskTitle,
            category || '通用',
            newPriority,
            new Date(),
            nextDueDate,
            '未開始',
            false,
            remarks || '',
            recurrence
          ]);
          return `✅ 任務已完成！並已自動為您排定下一期【${recurrence}】（${nextDueDate}）！`;
        }
      }
    } else {
      taskSheet.getRange(rowIndex, 10).setValue(false);
    }
    return '✅ 狀態已更新為：' + newStatus;
  } catch (err) {
    return '❌ 更新失敗：' + err.toString();
  }
}
