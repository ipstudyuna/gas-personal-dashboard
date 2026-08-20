/**
 * 個人視覺化工作儀表板 Web App
 * v4.7：獨立【⚠️ 逾期任務】頁籤 + 修復週期性任務展延
 * v4.8：內建多角色一鍵週報生成與雙週會議採集
 * v4.9：週報一鍵輸出「下載 Word (.docx)」與「寫入我的雲端週報文件」功能
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

  if (recurrence === '每日' || recurrence.includes('日')) {
    nextDate.setDate(nextDate.getDate() + 1);
  } else if (recurrence === '每週' || recurrence.includes('週')) {
    nextDate.setDate(nextDate.getDate() + 7);
  } else if (recurrence === '每季' || recurrence.includes('季')) {
    const origDay = baseDate.getDate();
    nextDate.setMonth(nextDate.getMonth() + 3);
    if (nextDate.getDate() !== origDay) {
      nextDate.setDate(0); // 避免月底日期溢位（例如 1/31 +3月），自動調整為該月最後一天
    }
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

/**
 * 取得「本週 + 下週」兩週期間的行事曆行程（供週報會議勾選使用）
 * isWork 以簡易關鍵字判斷是否為工作相關行程，預設勾選
 */
function getBiWeeklyCalendarEvents(tz) {
  const events = [];
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const day = d.getDay();
  const diffToMon = (day === 0 ? -6 : 1) - day;

  const thisMon = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diffToMon, 0, 0, 0);
  const thisSun = new Date(thisMon.getFullYear(), thisMon.getMonth(), thisMon.getDate() + 6, 23, 59, 59);
  const nextMon = new Date(thisMon.getFullYear(), thisMon.getMonth(), thisMon.getDate() + 7, 0, 0, 0);
  const nextSun = new Date(thisMon.getFullYear(), thisMon.getMonth(), thisMon.getDate() + 13, 23, 59, 59);

  const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  const workKeywordRegex = /[會研習專案課錄審排協調檢核訪討論工作坊培訓]/;

  // 1. Google 主要日曆
  try {
    const cal = CalendarApp.getDefaultCalendar();
    if (cal) {
      const googleCalEvents = cal.getEvents(thisMon, nextSun);
      googleCalEvents.forEach(e => {
        const s = e.getStartTime();
        const end = e.getEndTime();
        const isAllDay = e.isAllDayEvent();
        const dateStr = `${String(s.getMonth() + 1).padStart(2, '0')}/${String(s.getDate()).padStart(2, '0')} (${weekdays[s.getDay()]})`;
        const timeStr = isAllDay ? '全天' : `${Utilities.formatDate(s, tz, 'HH:mm')} - ${Utilities.formatDate(end, tz, 'HH:mm')}`;
        const title = String(e.getTitle() || '未命名行程');

        const isNextWeek = s >= nextMon;
        const isWork = workKeywordRegex.test(title);

        events.push({
          dateStr: dateStr,
          dateIso: Utilities.formatDate(s, tz, 'yyyy-MM-dd'),
          time: timeStr,
          title: title,
          location: String(e.getLocation() || ''),
          source: '主要日曆',
          isThisWeek: !isNextWeek,
          isNextWeek: isNextWeek,
          isWork: isWork
        });
      });
    }
  } catch (err) {
    console.error('BiWeekly Google Cal error:', err);
  }

  // 2. 外部 iCal 日曆
  try {
    const ss = getSpreadsheet();
    const setSheet = ss.getSheetByName("⚙️ 系統設定");
    if (setSheet && setSheet.getLastRow() >= 2) {
      const icalUrls = setSheet.getRange(2, 4, setSheet.getLastRow() - 1, 2).getDisplayValues();
      icalUrls.forEach(urlRow => {
        const [urlA, urlB] = urlRow.map(val => String(val || "").trim());
        const url = urlA || urlB || "";
        if (url && url.startsWith("http")) {
          for (let i = 0; i < 14; i++) {
            const targetD = new Date(thisMon.getFullYear(), thisMon.getMonth(), thisMon.getDate() + i);
            const extEvents = fetchExternalIcsEvents(url, targetD);
            const isNextWeek = i >= 7;
            const dateStr = `${String(targetD.getMonth() + 1).padStart(2, '0')}/${String(targetD.getDate()).padStart(2, '0')} (${weekdays[targetD.getDay()]})`;

            extEvents.forEach(ee => {
              const isWork = workKeywordRegex.test(ee.title);
              events.push({
                dateStr: dateStr,
                dateIso: Utilities.formatDate(targetD, tz, 'yyyy-MM-dd'),
                time: ee.time,
                title: ee.title,
                location: ee.location,
                source: '外部日曆',
                isThisWeek: !isNextWeek,
                isNextWeek: isNextWeek,
                isWork: isWork
              });
            });
          }
        }
      });
    }
  } catch (err) {
    console.error('BiWeekly iCal error:', err);
  }

  return events.sort((a, b) => a.dateIso.localeCompare(b.dateIso));
}

function getWebAppDataJson() {
  const result = {
    summary: { totalProjects: 0, pendingTasks: 0, calCount: 0, overallRate: 0 },
    calEvents: [],
    biWeeklyCalEvents: [],
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

  // 2b. 雙週（本週+下週）行事曆行程，供週報會議勾選使用
  try {
    result.biWeeklyCalEvents = getBiWeeklyCalendarEvents(tz);
  } catch (err) {
    console.error('biWeeklyCalEvents error:', err);
  }

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

/* ============================================================
 * 週報輸出功能 (v4.9)
 * - generateWeeklyReportDocxBase64：將週報文字轉為 Word (.docx) 檔案並回傳 base64，供前端一鍵下載
 * - appendReportToMyDoc / bindMyReportDoc / getMyReportDocInfo：
 *   讓每位組員綁定「自己專屬」的 Google 文件，之後每次產出週報可一鍵寫入該文件的新頁面（累積存檔）
 *
 * 部署方式（適用「每人各自複製一份範本試算表、各自部署」的散布模式）：
 *   執行身份 (Execute as)：我 (Me)
 *   誰可以存取 (Who has access)：僅限我自己 (Only myself)
 * 因為每個人都是用「自己的帳號」複製範本並部署，此時「以我身份執行」裡的「我」對每個人來說就是他自己，
 * 且存取權限又限制成只有他自己能開啟，「操作者」與「執行身份」剛好是同一人，
 * Session.getActiveUser() 才能正確拿到使用者 email，綁定機制才會生效。
 *
 * 這個功能第一次使用時，必須先手動授權 Google 文件 / 雲端硬碟權限，
 * 否則會出現「你沒有呼叫「DocumentApp.xxx」的權限」錯誤。
 * 解法：在 Apps Script 編輯器頂端的函式下拉選單選擇 authorizeReportFeatures，
 * 按「執行」，跳出授權視窗後選擇帳號 → 進階 → 前往（不安全）→ 允許，即可。
 * ============================================================ */

/**
 * 手動執行一次此函式即可觸發 Google 文件 / 雲端硬碟權限的授權視窗。
 * 授權完成後，週報「下載 Word」與「寫入我的雲端週報文件」功能才能正常運作。
 */
function authorizeReportFeatures() {
  const tempDoc = DocumentApp.create('__週報功能授權測試__' + new Date().getTime());
  const id = tempDoc.getId();
  tempDoc.saveAndClose();
  DriveApp.getFileById(id).setTrashed(true);
  Logger.log('✅ 授權成功！週報「下載 Word」與「寫入我的雲端週報文件」功能現在可以正常使用了。');
}

const REPORT_DOC_MAP_SHEET = '📄 週報文件對照';

/**
 * 從週報 Sheet 表對照表中，依 email 查詢已綁定的 Google 文件 ID
 */
function getUserReportDocId_(email) {
  if (!email) return '';
  const ss = getSpreadsheet();
  const mapSheet = ss.getSheetByName(REPORT_DOC_MAP_SHEET);
  if (!mapSheet) return '';
  const lastRow = mapSheet.getLastRow();
  if (lastRow < 2) return '';
  const data = mapSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.trim().toLowerCase()) {
      return String(data[i][1]).trim();
    }
  }
  return '';
}

/**
 * 從網址或純 ID 字串中解析出 Google 文件 ID
 */
function extractDocId_(input) {
  if (!input) return '';
  const s = String(input).trim();
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  return s;
}

/**
 * 取得目前登入者的週報文件綁定狀態
 */
function getMyReportDocInfo() {
  const email = Session.getActiveUser().getEmail() || '';
  if (!email) {
    return {
      bound: false,
      email: '',
      error: '無法取得您的登入身份。請確認此工具的部署設定為「執行身份：存取網頁應用程式的使用者」，且存取權限已限制為貴組織/網域內的登入使用者。'
    };
  }
  const docId = getUserReportDocId_(email);
  if (!docId) {
    return { bound: false, email: email };
  }
  return {
    bound: true,
    email: email,
    docId: docId,
    url: 'https://docs.google.com/document/d/' + docId + '/edit'
  };
}

/**
 * 綁定（或更新）目前登入者專屬的週報 Google 文件
 */
function bindMyReportDoc(docUrlOrId, displayName) {
  try {
    const email = Session.getActiveUser().getEmail() || '';
    if (!email) {
      return { success: false, message: '無法取得您的登入身份，請確認部署設定（詳見程式碼註解）後再試一次。' };
    }
    const docId = extractDocId_(docUrlOrId);
    if (!docId) {
      return { success: false, message: '請輸入有效的 Google 文件連結。' };
    }

    // 驗證是否能開啟此文件（確認權限）
    try {
      DocumentApp.openById(docId);
    } catch (e) {
      return { success: false, message: '無法開啟此文件，請確認連結正確、且您對該文件有編輯權限：' + e.toString() };
    }

    const ss = getSpreadsheet();
    let mapSheet = ss.getSheetByName(REPORT_DOC_MAP_SHEET);
    if (!mapSheet) {
      mapSheet = ss.insertSheet(REPORT_DOC_MAP_SHEET);
      mapSheet.getRange('A1:D1').setValues([['Email', 'GoogleDocID', '姓名', '最後更新時間']])
        .setFontWeight('bold').setBackground('#E8F0FE');
    }

    const lastRow = mapSheet.getLastRow();
    let foundRow = -1;
    if (lastRow >= 2) {
      const emails = mapSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (String(emails[i][0]).trim().toLowerCase() === email.trim().toLowerCase()) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > 0) {
      mapSheet.getRange(foundRow, 2, 1, 3).setValues([[docId, displayName || '', new Date()]]);
    } else {
      mapSheet.appendRow([email, docId, displayName || '', new Date()]);
    }

    return { success: true, message: '綁定成功！之後點擊「寫入我的週報文件」即會自動將週報加到此文件的新頁面。' };
  } catch (err) {
    return { success: false, message: '綁定失敗：' + err.toString() };
  }
}

/**
 * 週報文件視覺樣式參數（配色參考組員提供的 Word 範本：藍灰色編號、灰色說明文字、淺灰表頭底色）
 */
var REPORT_STYLE_ = {
  BLUE: '#3B6E8F',
  GRAY: '#7B847D',
  DARKGRAY: '#4A4A52',
  HEADER_SHADE: '#F2F5F2',
  BORDER_COLOR: '#D0D5DB'
};

/**
 * 計算「本週一～本週五」的日期範圍標籤，例如 8/17-8/21
 * label：人類可讀（含斜線），safe：檔名安全版本（不含斜線）
 */
function getWeekRangeLabel_(tz) {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const day = d.getDay();
  const diffToMon = (day === 0 ? -6 : 1) - day;
  const mon = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diffToMon);
  const fri = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 4);
  return {
    label: Utilities.formatDate(mon, tz, 'M/d') + '-' + Utilities.formatDate(fri, tz, 'M/d'),
    safe: Utilities.formatDate(mon, tz, 'MMdd') + '-' + Utilities.formatDate(fri, tz, 'MMdd')
  };
}

/**
 * 在文件最上方插入本次週報的日期範圍標題（做為「這是哪一週」的頁面標記）
 */
function appendWeekRangeHeading_(body, weekRangeLabel) {
  const p = body.appendParagraph('📅　' + weekRangeLabel + '　週報');
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.editAsText().setForegroundColor(REPORT_STYLE_.BLUE);
  body.appendHorizontalRule();
}

/**
 * 統一套用表格樣式：細邊框 + 表頭列淺灰底色與粗體（headerRowIndex 傳 -1 表示不需要表頭樣式）
 */
function styleReportTable_(table, headerRowIndex) {
  try {
    table.setBorderWidth(1);
    table.setBorderColor(REPORT_STYLE_.BORDER_COLOR);
  } catch (e) { /* ignore */ }

  if (headerRowIndex !== null && headerRowIndex >= 0 && table.getNumRows() > headerRowIndex) {
    const headerRow = table.getRow(headerRowIndex);
    for (let c = 0; c < headerRow.getNumCells(); c++) {
      const cell = headerRow.getCell(c);
      cell.setBackgroundColor(REPORT_STYLE_.HEADER_SHADE);
      const cellText = cell.editAsText();
      const len = cellText.getText().length;
      if (len > 0) cellText.setBold(0, len - 1, true);
    }
  }
}

/**
 * 將週報文字內容渲染進 Google 文件的 Body（標題／表格／段落），並套上跟組員範本一致的視覺樣式
 * reportText 格式為 updateReportPreview() 產出的純文字週報（含 Markdown 表格語法）
 * 01 章節（上週未完成事項）會渲染成單欄表格，02～05 章節的 Markdown 表格會渲染成有表頭底色的表格
 */
function renderReportIntoBody_(body, reportText) {
  const S = REPORT_STYLE_;
  const lines = String(reportText || '').replace(/\r/g, '').split('\n');
  let i = 0;

  // 1. 開頭標題區（第一個「填表人」或分隔線之前的行）
  const titleLines = [];
  while (i < lines.length && !/^填表人/.test(lines[i].trim()) && !/^=+$/.test(lines[i].trim())) {
    const t = lines[i].trim();
    if (t) titleLines.push(t);
    i++;
  }
  if (titleLines.length > 0) {
    const roleP = body.appendParagraph(titleLines[0]);
    roleP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const roleText = roleP.editAsText();
    roleText.setBold(true);
    roleText.setFontSize(9);
    roleText.setForegroundColor(S.DARKGRAY);
  }
  if (titleLines.length > 1) {
    const titleP = body.appendParagraph(titleLines[1]);
    titleP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const titleText = titleP.editAsText();
    titleText.setBold(true);
    titleText.setFontSize(17);
  }

  // 2. 填表人 meta 行
  if (i < lines.length && /^填表人/.test(lines[i].trim())) {
    const metaP = body.appendParagraph(lines[i].trim());
    metaP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const metaText = metaP.editAsText();
    metaText.setBold(true);
    metaText.setFontSize(10);
    metaText.setForegroundColor(S.GRAY);
    i++;
  }

  body.appendHorizontalRule();

  // 3. 依 "======" 分隔的各章節內容
  let currentTableRows = null;
  let currentSectionNum = '';
  let section01Rows = [];

  const flushTable = function () {
    if (currentTableRows && currentTableRows.length > 0) {
      const table = body.appendTable(currentTableRows);
      styleReportTable_(table, 0);
    }
    currentTableRows = null;
  };

  const flushSection01 = function () {
    if (section01Rows.length > 0) {
      const rows = section01Rows.map(function (r) { return [r]; });
      const table = body.appendTable(rows);
      styleReportTable_(table, -1);
    }
    section01Rows = [];
  };

  for (; i < lines.length; i++) {
    const trimmed = lines[i].trim();

    if (trimmed === '') { continue; }
    if (/^=+$/.test(trimmed)) { flushTable(); flushSection01(); continue; }

    // 章節標題，如「01  上週未完成事項追蹤」：編號用藍灰色小字，標題用黑色粗體大字
    const sectionMatch = trimmed.match(/^(\d{2})(\s+)(\S.*)$/);
    if (sectionMatch) {
      flushTable();
      flushSection01();
      currentSectionNum = sectionMatch[1];

      const h = body.appendParagraph(trimmed);
      h.setSpacingBefore(16);
      h.setSpacingAfter(6);
      const numLen = sectionMatch[1].length + sectionMatch[2].length;
      const text = h.editAsText();
      text.setBold(0, trimmed.length - 1, true);
      text.setFontSize(0, numLen - 1, 10);
      text.setForegroundColor(0, numLen - 1, S.BLUE);
      text.setFontSize(numLen, trimmed.length - 1, 13);
      continue;
    }

    // Markdown 表格列
    if (trimmed.charAt(0) === '|') {
      if (/^\|[\s:\-|]+\|$/.test(trimmed)) { continue; } // 分隔列 | :--- | :--- |
      const cells = trimmed.split('|').slice(1, -1).map(function (c) { return c.trim(); });
      if (!currentTableRows) currentTableRows = [];
      currentTableRows.push(cells);
      continue;
    }

    // 01 章節的自由文字（上週未完成事項清單）渲染成單欄表格，與範本樣式一致
    if (currentSectionNum === '01') {
      section01Rows.push(trimmed);
      continue;
    }

    flushTable();
    body.appendParagraph(trimmed);
  }
  flushTable();
  flushSection01();
}

/**
 * （已停用）原本用 DocumentApp + DriveApp.getAs(MimeType.MICROSOFT_WORD) 產生 .docx 的做法，
 * 會偶發「無法從 application/vnd.google-apps.document 轉換成 ...wordprocessingml.document」錯誤。
 * 目前「下載 Word 週報」改成前端直接組成 MS-Office HTML 存成 .doc（見 Index.html 的
 * downloadWeeklyReportDocx()），不需要伺服器轉檔，也就不會再遇到這個問題。
 * 這個函式保留起來，若之後想改回產生「真正的」.docx 二進位檔，可以再接回來用。
 */
function generateWeeklyReportDocxBase64_DEPRECATED(reportText, meta) {
  meta = meta || {};
  let tempDoc = null;
  try {
    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const weekRange = getWeekRangeLabel_(tz);
    const safeUserName = (meta.userName || '週報').replace(/[\\\/:*?"<>|]/g, '');
    const safeGroupName = (meta.groupName || '').replace(/[\\\/:*?"<>|]/g, '');

    tempDoc = DocumentApp.create('週報暫存_' + safeUserName + '_' + new Date().getTime());
    const body = tempDoc.getBody();
    body.clear();
    appendWeekRangeHeading_(body, weekRange.label);
    renderReportIntoBody_(body, reportText);
    tempDoc.saveAndClose();

    const fileId = tempDoc.getId();
    let docxBlob = null;
    let lastErr = null;
    for (let attempt = 0; attempt < 4 && !docxBlob; attempt++) {
      try {
        Utilities.sleep(attempt === 0 ? 1000 : 1500);
        docxBlob = DriveApp.getFileById(fileId).getAs(MimeType.MICROSOFT_WORD);
      } catch (e) {
        lastErr = e;
      }
    }
    if (!docxBlob) {
      throw lastErr || new Error('轉換為 Word 檔失敗（未知原因）');
    }

    const base64 = Utilities.base64Encode(docxBlob.getBytes());
    const filename = `週報_${safeGroupName}_${safeUserName}_${weekRange.safe}.docx`;
    DriveApp.getFileById(fileId).setTrashed(true);

    return {
      success: true,
      base64: base64,
      filename: filename,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
  } catch (err) {
    if (tempDoc) {
      try { DriveApp.getFileById(tempDoc.getId()).setTrashed(true); } catch (e2) {}
    }
    return { success: false, message: '產生 Word 檔失敗：' + err.toString() };
  }
}

/**
 * 將本次週報，以「新頁面」的方式附加寫入目前登入者已綁定的專屬 Google 文件，
 * 並在該頁最上方自動標註本週日期範圍（例如 8/17-8/21）做為頁面標記
 */
function appendReportToMyDoc(reportText, meta) {
  meta = meta || {};
  try {
    const email = Session.getActiveUser().getEmail() || '';
    if (!email) {
      return { success: false, message: '無法取得您的登入身份，請確認部署設定後再試一次。' };
    }

    const docId = getUserReportDocId_(email);
    if (!docId) {
      return { success: false, needBind: true, message: '尚未綁定您的專屬週報文件，請先在上方貼上您的 Google 文件連結並綁定。' };
    }

    let doc;
    try {
      doc = DocumentApp.openById(docId);
    } catch (e) {
      return { success: false, message: '無法開啟已綁定的文件，請確認您仍有該文件的編輯權限：' + e.toString() };
    }

    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const weekRange = getWeekRangeLabel_(tz);

    const body = doc.getBody();
    if (body.getNumChildren() > 0) {
      body.appendPageBreak();
    }
    appendWeekRangeHeading_(body, weekRange.label);
    renderReportIntoBody_(body, reportText);
    doc.saveAndClose();

    // 同步更新對照表中的姓名，方便日後辨識
    if (meta.userName) {
      bindMyReportDoc(docId, meta.userName);
    }

    return {
      success: true,
      message: `已成功將 ${weekRange.label} 週報寫入您的專屬文件（新頁面）！`,
      url: 'https://docs.google.com/document/d/' + docId + '/edit'
    };
  } catch (err) {
    return { success: false, message: '寫入失敗：' + err.toString() };
  }
}
