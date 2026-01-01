// ================================================================================
// CENTRAL CONFIG CONNECTION
// ================================================================================
const CONFIG_SHEET_ID = '1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk';
const CONFIG_CACHE_DURATION = 300;

// ================================================================================
// ERROR LOGGING TO HUB
// ================================================================================
function logErrorToHub_(errorType, message) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let sheet = ss.getSheetByName('Error Log');
    
    if (!sheet) {
      sheet = ss.insertSheet('Error Log');
      sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'System', 'Error Type', 'Message', 'User']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f4cccc');
    }
    
    let user = 'Unknown';
    try { user = Session.getActiveUser().getEmail(); } catch (_) {}
    
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 5).setValues([[
      new Date().toISOString(),
      'Drive Naming QA',
      errorType,
      String(message).substring(0, 500),
      user
    ]]);
  } catch (e) {
    Logger.log('Could not log error to Hub: ' + e);
  }
}

function getProductsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_products');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  
  const products = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === true) products.push(data[i][0]);
  }
  
  cache.put('dnc_products', JSON.stringify(products), CONFIG_CACHE_DURATION);
  return products;
}

function getPlatformsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_platforms');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf('Platforms');
  
  const platforms = [];
  if (col >= 0) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][col] && data[i][col].toString().trim() !== '') {
        platforms.push(data[i][col].toString().trim());
      }
    }
  }
  
  cache.put('dnc_platforms', JSON.stringify(platforms), CONFIG_CACHE_DURATION);
  return platforms;
}

function getLocalesFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_locales');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf('Locales');
  
  const locales = [];
  if (col >= 0) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][col] && data[i][col].toString().trim() !== '') {
        locales.push(data[i][col].toString().trim());
      }
    }
  }
  
  cache.put('dnc_locales', JSON.stringify(locales), CONFIG_CACHE_DURATION);
  return locales;
}

function getFoldersToMonitorFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_folders');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf('Monitor Folders');
  
  const folders = [];
  if (col >= 0) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][col] && data[i][col].toString().trim() !== '') {
        folders.push(data[i][col].toString().trim());
      }
    }
  }
  
  cache.put('dnc_folders', JSON.stringify(folders), CONFIG_CACHE_DURATION);
  return folders;
}

function getProductAliasesFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_aliases');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const aliasCol = headers.indexOf('Alias');
  const mapsToCol = headers.indexOf('Maps To');
  
  const aliases = {};
  if (aliasCol >= 0 && mapsToCol >= 0) {
    for (let i = 1; i < data.length; i++) {
      const alias = data[i][aliasCol];
      const mapsTo = data[i][mapsToCol];
      if (alias && mapsTo && alias.toString().trim() !== '' && mapsTo.toString().trim() !== '') {
        aliases[alias.toString().trim()] = mapsTo.toString().trim();
      }
    }
  }
  
  cache.put('dnc_aliases', JSON.stringify(aliases), CONFIG_CACHE_DURATION);
  return aliases;
}

function clearDNCConfigCache() {
  CacheService.getScriptCache().removeAll(['dnc_products', 'dnc_platforms', 'dnc_locales', 'dnc_folders', 'dnc_aliases']);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Config cache cleared!', 'Done', 3);
}
// ================================================================================

/*******************************************************
 * DRIVE NAMING CHECKER v3.2 - AUTO-CHECK HISTORY
 * 
 * NEW in v3.2:
 * - ✅ Separate "Auto-Check History" tab for automated triggers
 * - ✅ Week-based grouping with ISO week numbers
 * - ✅ Newest checks at top
 * - ✅ Auto-archive old data (12+ weeks)
 * - ✅ Manual checks still use "Naming Violations" tab
 * - ✅ Slack links point to correct tab
 * 
 * Previous features:
 * - ✅ Real-time progress dialog with live updates
 * - ✅ Async execution (checks one product at a time)
 * - ✅ Manual button tracking in Trigger Status
 * - ✅ Slack integration with breakdown
 *******************************************************/

// ========================================
// 🔧 CONFIGURATION
// ========================================

const DRIVE_CONFIG = {
  get FOLDERS_TO_MONITOR() {
    return getFoldersToMonitorFromConfig_();
  },
  VIOLATIONS_SHEET_NAME: 'Naming Violations',
  AUTO_CHECK_HISTORY_SHEET_NAME: 'Auto-Check History',
  ARCHIVE_SHEET_NAME: 'Archive',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  MAX_DEPTH: 10,
  SKIP_OLD_YEARS: true,
  CURRENT_YEAR: new Date().getFullYear(),
  YEARS_TO_CHECK: 0,
  WEEKS_TO_KEEP: 12
};

const SETTINGS_SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SETTINGS_SHEET_NAME = 'Settings';

const LOG_FOLDER_ID = '12pDwmOkfaIhoosGvZoCtAXkZI99rq-8i';

const VALID_VALUES = {
  get PRODUCTS() {
    return getProductsFromConfig_();
  },
  get PLATFORMS() {
    return getPlatformsFromConfig_();
  },
  get LOCALES() {
    return getLocalesFromConfig_();
  }
};

const COUNTRY_TO_LOCALE = {
  'SPAIN': 'ES', 'GERMANY': 'DE', 'UNITED STATES': 'US', 'UNITED KINGDOM': 'UK',
  'AUSTRALIA': 'AU', 'CANADA': 'CA', 'ITALY': 'IT', 'FRANCE': 'FR',
  'BRAZIL': 'BR', 'MEXICO': 'MX', 'ISRAEL': 'IL'
};

const LAZY_PATTERNS = {
  FORBIDDEN_WORDS: ['temp', 'draft', 'untitled', 'new folder', 'backup', 'old', 'archive'],
  COPY_PATTERNS: [
    /^copy of /i, 
    /^copy \d+ of /i, 
    / - copy$/i,
    / \(\d+\)\.[^.]+$/,
    / \(\d+\)$/
  ],
  TEST_PATTERNS: [/^test$/i, /^test\d+$/i, /^test_/i, /_test$/i, /\(test\)/i, /testtask/i, /testnow/i]
};

// ========================================
// 🌍 GLOBAL STATE
// ========================================

let loggedViolations = new Set();
let debugLog = [];
const STOP_FLAG_PROPERTY = 'DRIVE_CHECKER_STOP';
let startTime = Date.now();

const PROGRESS_KEY = 'PROGRESS_DATA';

// Track current check for auto-history logging
let currentCheckInfo = null;

// ========================================
// 📋 MENU
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 Drive Checker')
    .addItem('▶️ Run Check Now (Today)', 'manualCheckToday')
    .addItem('📅 Check This Week (Mon-Today)', 'manualCheckThisWeek')
    .addItem('📅 Check Last Week (Mon-Sun)', 'checkLastWeek')
    .addItem('🛑 Stop Execution', 'stopExecution')
    .addSeparator()
    .addItem('🗑️ Delete Last Slack Message', 'deleteLastSlackMessage')
    .addItem('🎯 Select & Delete Slack Message', 'selectAndDeleteSlackMessage')
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('✅ Setup All Auto-Checks', 'setupAllTriggers')
      .addItem('❌ Remove All Auto-Checks', 'removeAllTriggers')
      .addItem('📋 View Active Triggers', 'viewActiveTriggers'))
    .addSeparator()
    .addItem('🎨 Setup Professional Sheets', 'setupProfessionalSheets')
    .addItem('🗑️ Clear Manual Violations', 'clearViolationsSheet')
    .addItem('🗄️ Archive Old Checks', 'archiveOldWeeks')
    .addToUi();
}

function stopExecution() {
  PropertiesService.getScriptProperties().setProperty(STOP_FLAG_PROPERTY, 'true');
  SpreadsheetApp.getUi().alert('🛑 Stop signal sent!\n\nThe script will halt at the next folder check.');
  log('🛑 STOP signal set by user');
}

function checkStopFlag() {
  const stopFlag = PropertiesService.getScriptProperties().getProperty(STOP_FLAG_PROPERTY);
  if (stopFlag === 'true') {
    log('🛑 STOP requested by user - halting execution');
    PropertiesService.getScriptProperties().deleteProperty(STOP_FLAG_PROPERTY);
    throw new Error('Execution stopped by user');
  }
}

// ========================================
// 🚀 TRIGGERS
// ========================================

function setupAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  ScriptApp.newTrigger('checkPreviousDay').timeBased().atHour(9).everyDays(1).create();
  
  ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY'].forEach(day => {
    ScriptApp.newTrigger('checkCurrentDay').timeBased().atHour(14).everyWeeks(1).onWeekDay(ScriptApp.WeekDay[day]).create();
    ScriptApp.newTrigger('checkCurrentDay').timeBased().atHour(17).everyWeeks(1).onWeekDay(ScriptApp.WeekDay[day]).create();
  });
  
  ScriptApp.newTrigger('checkThisWeek').timeBased().atHour(13).onWeekDay(ScriptApp.WeekDay.FRIDAY).everyWeeks(1).create();
  
  SpreadsheetApp.getUi().alert(
    '✅ All triggers created!\n\n' +
    '• 9 AM Daily: Previous day violations\n' +
    '• 2 PM Weekdays: Current day violations\n' +
    '• 5 PM Weekdays: Current day violations\n' +
    '• 1 PM Friday: This week violations (Mon-Fri)'
  );
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  SpreadsheetApp.getUi().alert('✅ All automatic triggers removed!');
}

function viewActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert('No active triggers found.\n\nClick "Setup All Auto-Checks" to enable automation.');
    return;
  }
  
  const grouped = {};
  triggers.forEach(trigger => {
    const func = trigger.getHandlerFunction();
    grouped[func] = (grouped[func] || 0) + 1;
  });
  
  let message = `Active Triggers (${triggers.length} total):\n\n`;
  if (grouped['checkPreviousDay']) message += `✅ 9 AM Daily Check (${grouped['checkPreviousDay']})\n   → Previous day violations\n\n`;
  if (grouped['checkCurrentDay']) message += `✅ Weekday Checks (${grouped['checkCurrentDay']})\n   → 2 PM Mon-Fri (5 triggers)\n   → 5 PM Mon-Fri (5 triggers)\n\n`;
  if (grouped['checkThisWeek']) message += `✅ Friday Summary (${grouped['checkThisWeek']})\n   → 1 PM Friday (weekly)\n\n`;
  message += '💡 This is correct - weekday checks need\n   separate triggers for each day.';
  
  SpreadsheetApp.getUi().alert(message);
}

// ========================================
// 📅 CHECK FUNCTIONS
// ========================================

function checkPreviousDay() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);
  const endOfYesterday = new Date(yesterday);
  endOfYesterday.setHours(23, 59, 59, 999);
  runCheckSync(yesterday, endOfYesterday, 'Previous Day');
}

function checkCurrentDay() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const now = new Date();
  const label = now.getHours() >= 16 ? 'Today-Evening' : 'Today';
  runCheckSync(today, now, label);
}

function checkThisWeek() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1));
  monday.setHours(0, 0, 0, 0);
  runCheckSync(monday, new Date(), 'This Week (Mon-Fri)');
}

function manualCheckToday() {
  showProgressDialog();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  startAsyncCheck(today, new Date(), 'Today-Manual');
}

function manualCheckThisWeek() {
  showProgressDialog();
  const today = new Date();
  const dayOfWeek = today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1));
  monday.setHours(0, 0, 0, 0);
  startAsyncCheck(monday, new Date(), 'This Week (Mon-Today)');
}

function checkLastWeek() {
  showProgressDialog();
  const today = new Date();
  const dayOfWeek = today.getDay();
  const lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1) - 7);
  lastMonday.setHours(0, 0, 0, 0);
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);
  lastSunday.setHours(23, 59, 59, 999);
  startAsyncCheck(lastMonday, lastSunday, 'Last Week (Mon-Sun)');
}

// ========================================
// 🎯 ASYNC CHECK SYSTEM
// ========================================

function startAsyncCheck(startDate, endDate, label) {
  // Initialize
  PropertiesService.getScriptProperties().deleteProperty(STOP_FLAG_PROPERTY);
  PropertiesService.getScriptProperties().setProperty('CHECK_START_DATE', startDate.getTime().toString());
  PropertiesService.getScriptProperties().setProperty('CHECK_END_DATE', endDate.getTime().toString());
  PropertiesService.getScriptProperties().setProperty('CHECK_LABEL', label);
  
  clearViolationsSheet();
  
  loggedViolations = new Set();
  debugLog = [];
  startTime = Date.now();
  
  // Save violations to properties (for async access)
  PropertiesService.getScriptProperties().setProperty('LOGGED_VIOLATIONS', JSON.stringify([]));
  
  // Initialize progress
  const progressData = {
    totalFolders: DRIVE_CONFIG.FOLDERS_TO_MONITOR.length,
    foldersChecked: 0,
    filesChecked: 0,
    violations: 0,
    currentFolder: 'Starting...',
    completed: false,
    currentIndex: 0,
    checkLabel: label
  };
  PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, JSON.stringify(progressData));
  
  log('================================================================================');
  log(`🔍 DRIVE NAMING CHECKER - ${label.toUpperCase()}`);
  log(`📅 ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
  log('================================================================================');
  
  // Start checking first folder
  checkNextFolder();
}

function checkNextFolder() {
  try {
    checkStopFlag();
    
    const progressData = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROGRESS_KEY));
    const index = progressData.currentIndex;
    
    if (index >= DRIVE_CONFIG.FOLDERS_TO_MONITOR.length) {
      // All done!
      finishAsyncCheck();
      return;
    }
    
    const folderId = DRIVE_CONFIG.FOLDERS_TO_MONITOR[index];
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    progressData.currentFolder = folderName;
    progressData.foldersChecked = index + 1;
    PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, JSON.stringify(progressData));
    
    log(`📁 Scanning folder: ${folderName}`);
    
    // Load violations from storage
    const storedViolations = JSON.parse(PropertiesService.getScriptProperties().getProperty('LOGGED_VIOLATIONS') || '[]');
    loggedViolations = new Set(storedViolations);
    
    // Check this folder
    checkFolder(folder, 0, folderName);
    
    // Save violations back
    PropertiesService.getScriptProperties().setProperty('LOGGED_VIOLATIONS', JSON.stringify(Array.from(loggedViolations)));
    
    // Update violation count
    progressData.violations = loggedViolations.size;
    progressData.currentIndex = index + 1;
    PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, JSON.stringify(progressData));
    
    // Continue to next folder (with small delay for UI update)
    Utilities.sleep(100);
    checkNextFolder();
    
  } catch (e) {
    if (e.message === 'Execution stopped by user') {
      finishAsyncCheck(true);
    } else {
      logErrorToHub_('Folder Check', e.message || String(e));
      log('❌ ERROR: ' + e.message);
      finishAsyncCheck();
    }
  }
}

function finishAsyncCheck(stopped = false) {
  const progressData = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROGRESS_KEY));
  const label = PropertiesService.getScriptProperties().getProperty('CHECK_LABEL');
  const storedViolations = JSON.parse(PropertiesService.getScriptProperties().getProperty('LOGGED_VIOLATIONS') || '[]');
  
  progressData.completed = true;
  progressData.currentFolder = stopped ? 'Stopped' : 'Complete!';
  PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, JSON.stringify(progressData));
  
  const violationsCount = storedViolations.length;
  
  log('================================================================================');
  log(`✅ Found ${violationsCount} violations for ${label}`);
  log('================================================================================');
  
  const executionTime = (Date.now() - startTime) / 1000;
  log(`⏱️  Total execution time: ${executionTime.toFixed(2)} seconds`);
  
  updateTriggerStatus(label, violationsCount);
  saveDebugLog();
  
  // Cleanup
  PropertiesService.getScriptProperties().deleteProperty('CHECK_START_DATE');
  PropertiesService.getScriptProperties().deleteProperty('CHECK_END_DATE');
  PropertiesService.getScriptProperties().deleteProperty('CHECK_LABEL');
  PropertiesService.getScriptProperties().deleteProperty('LOGGED_VIOLATIONS');
}

// Sync version for automated triggers (no dialog)
function runCheckSync(startDate, endDate, label) {
  try {
    PropertiesService.getScriptProperties().deleteProperty(STOP_FLAG_PROPERTY);
    PropertiesService.getScriptProperties().setProperty('CHECK_START_DATE', startDate.getTime().toString());
    PropertiesService.getScriptProperties().setProperty('CHECK_END_DATE', endDate.getTime().toString());
    PropertiesService.getScriptProperties().setProperty('CHECK_LABEL', label);
    
    // Initialize breakdown tracker
    PropertiesService.getScriptProperties().setProperty('CURRENT_CHECK_BREAKDOWN', '{}');
    
    // Initialize check info for auto-history logging
    currentCheckInfo = {
      label: label,
      timestamp: new Date(),
      isAutoCheck: true
    };
    
    loggedViolations = new Set();
    debugLog = [];
    startTime = Date.now();

    // Initialize file counter
    PropertiesService.getScriptProperties().setProperty('TOTAL_FILES_SCANNED', '0');
    
    log('================================================================================');
    log(`🔍 DRIVE NAMING CHECKER - ${label.toUpperCase()}`);
    log(`📅 ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
    log('================================================================================');
    
    // Setup check header in Auto-Check History
    setupAutoCheckHeader(label, new Date());
    
    DRIVE_CONFIG.FOLDERS_TO_MONITOR.forEach(folderId => {
      checkStopFlag();
      try {
        const folder = DriveApp.getFolderById(folderId);
        log(`📁 Scanning folder: ${folder.getName()}`);
        checkFolder(folder, 0, folder.getName());
      } catch (e) {
        if (e.message === 'Execution stopped by user') throw e;
        log(`❌ ERROR accessing folder ${folderId}: ${e.message}`);
      }
    });
    
    const newViolationsCount = loggedViolations.size;
    
    // Update check header with final count
    updateAutoCheckHeader(newViolationsCount);
    
    log('================================================================================');
    log(`✅ Found ${newViolationsCount} violations for ${label}`);
    log('================================================================================');
    
    updateTriggerStatus(label, newViolationsCount);
    saveDebugLog();

    sendSlackNotification(label, newViolationsCount);
    
    const executionTime = (Date.now() - startTime) / 1000;
    log(`⏱️  Total execution time: ${executionTime.toFixed(2)} seconds`);
    
    // Cleanup
    PropertiesService.getScriptProperties().deleteProperty('CHECK_START_DATE');
    PropertiesService.getScriptProperties().deleteProperty('CHECK_END_DATE');
    PropertiesService.getScriptProperties().deleteProperty('CHECK_LABEL');
    PropertiesService.getScriptProperties().deleteProperty('CURRENT_CHECK_BREAKDOWN');
    
    // Check for fixed violations
    markFixedViolations();
    
    // Reset check info
    currentCheckInfo = null;
    
  } catch (e) {
    if (e.message === 'Execution stopped by user') {
      saveDebugLog();
      SpreadsheetApp.getUi().alert('🛑 Execution Stopped');
    } else {
      log('❌ FATAL ERROR: ' + e.message);
      saveDebugLog();
    }
    currentCheckInfo = null;
  }
}

// ========================================
// 📊 PROGRESS DIALOG
// ========================================

function showProgressDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ProgressDialog')
    .setWidth(450)
    .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Drive Checker Progress');
}

function getProgressUpdate() {
  const progressStr = PropertiesService.getScriptProperties().getProperty(PROGRESS_KEY);
  if (!progressStr) {
    return {
      progress: 0,
      currentFolder: 'Initializing...',
      foldersChecked: 0,
      filesChecked: 0,
      violations: 0,
      completed: false,
      checkLabel: 'Manual Check'
    };
  }
  
  const progressData = JSON.parse(progressStr);
  const progress = (progressData.foldersChecked / progressData.totalFolders) * 100;
  
  return {
    progress: progress,
    currentFolder: progressData.currentFolder,
    foldersChecked: progressData.foldersChecked,
    filesChecked: progressData.filesChecked,
    violations: progressData.violations,
    completed: progressData.completed,
    checkLabel: progressData.checkLabel || 'Manual Check'
  };
}

// ========================================
// 📝 LOGGING
// ========================================

function log(message) {
  const timestamp = new Date().toISOString().split('T')[1].split('.')[0];
  debugLog.push(`[${timestamp}] ${message}`);
  Logger.log(message);
}

function saveDebugLog() {
  try {
    const fileName = `DriveChecker_DebugLog_${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
    const content = debugLog.join('\n');
    try {
      DriveApp.getFolderById(LOG_FOLDER_ID).createFile(fileName, content);
      log(`✅ Debug log saved: ${fileName}`);
    } catch (e) {
      DriveApp.createFile(fileName, content);
      log(`✅ Debug log saved to Drive root: ${fileName}`);
    }
  } catch (e) {
    log(`❌ Failed to save debug log: ${e.message}`);
  }
}

// ========================================
// 📊 SHEET MANAGEMENT
// ========================================

function setupProfessionalSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Manual Violations Sheet
  let violationsSheet = ss.getSheetByName(DRIVE_CONFIG.VIOLATIONS_SHEET_NAME);
  if (!violationsSheet) violationsSheet = ss.insertSheet(DRIVE_CONFIG.VIOLATIONS_SHEET_NAME);
  
  violationsSheet.clear();
  const headers = ['Check Date', 'Timestamp', 'Type', 'Product', 'Name', 'Location', 'Violation', 'Owner', 'Link', 'Status'];
  violationsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = violationsSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  violationsSheet.setFrozenRows(1);
  
  violationsSheet.setColumnWidth(1, 120);
  violationsSheet.setColumnWidth(2, 100);
  violationsSheet.setColumnWidth(3, 80);
  violationsSheet.setColumnWidth(4, 120);
  violationsSheet.setColumnWidth(5, 300);
  violationsSheet.setColumnWidth(6, 250);
  violationsSheet.setColumnWidth(7, 350);
  violationsSheet.setColumnWidth(8, 150);
  violationsSheet.setColumnWidth(9, 200);
  violationsSheet.setColumnWidth(10, 100);
  violationsSheet.setTabColor('#ea4335');
  
  // Setup Auto-Check History Sheet
  setupAutoCheckHistorySheet();
  
  // Setup Archive Sheet
  let archiveSheet = ss.getSheetByName(DRIVE_CONFIG.ARCHIVE_SHEET_NAME);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(DRIVE_CONFIG.ARCHIVE_SHEET_NAME);
    archiveSheet.hideSheet();
    archiveSheet.setTabColor('#9e9e9e');
  }
  
  // Setup Dashboard
  let dashboardSheet = ss.getSheetByName(DRIVE_CONFIG.DASHBOARD_SHEET_NAME);
  if (!dashboardSheet) dashboardSheet = ss.insertSheet(DRIVE_CONFIG.DASHBOARD_SHEET_NAME, 0);
  
  dashboardSheet.clear();
  dashboardSheet.setTabColor('#34a853');
  dashboardSheet.getRange('A1').setValue('📊 Drive Naming Checker Dashboard').setFontSize(16).setFontWeight('bold');
  dashboardSheet.getRange('A3').setValue('Last Check:').setFontWeight('bold');
  dashboardSheet.getRange('B3').setValue(new Date().toLocaleString());
  dashboardSheet.getRange('A4').setValue('Manual Violations:').setFontWeight('bold');
  dashboardSheet.getRange('B4').setFormula(`=COUNTA('${DRIVE_CONFIG.VIOLATIONS_SHEET_NAME}'!A:A)-1`);
  dashboardSheet.getRange('A5').setValue('Auto-Check Violations:').setFontWeight('bold');
  dashboardSheet.getRange('B5').setFormula(`=COUNTIFS('${DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME}'!C:C,"FILE")+COUNTIFS('${DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME}'!C:C,"FOLDER")`);
  dashboardSheet.setColumnWidth(1, 200);
  dashboardSheet.setColumnWidth(2, 200);
  
  setupTriggerStatusSheet();
  
  SpreadsheetApp.getUi().alert('✅ Professional sheets created!');
}

function setupAutoCheckHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  if (!historySheet) historySheet = ss.insertSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  
  historySheet.setTabColor('#673ab7'); // Purple
  historySheet.setFrozenRows(0);
  
  // Set column widths
  historySheet.setColumnWidth(1, 120);
  historySheet.setColumnWidth(2, 100);
  historySheet.setColumnWidth(3, 80);
  historySheet.setColumnWidth(4, 120);
  historySheet.setColumnWidth(5, 300);
  historySheet.setColumnWidth(6, 250);
  historySheet.setColumnWidth(7, 350);
  historySheet.setColumnWidth(8, 150);
  historySheet.setColumnWidth(9, 200);
  historySheet.setColumnWidth(10, 100);
}

function setupTriggerStatusSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let statusSheet = ss.getSheetByName('Trigger Status');
  if (!statusSheet) statusSheet = ss.insertSheet('Trigger Status', 1);
  
  statusSheet.clear();
  statusSheet.setTabColor('#fbbc04');
  
  const headers = ['Check Type', 'Schedule', 'Last Run', 'Status', 'Violations Found'];
  statusSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  statusSheet.getRange(1, 1, 1, headers.length).setBackground('#fbbc04').setFontColor('#000000').setFontWeight('bold').setFontSize(11);
  statusSheet.setFrozenRows(1);
  
  statusSheet.setColumnWidth(1, 250);
  statusSheet.setColumnWidth(2, 200);
  statusSheet.setColumnWidth(3, 180);
  statusSheet.setColumnWidth(4, 100);
  statusSheet.setColumnWidth(5, 150);
  
  const triggers = [
    ['🌅 Previous Day Check', 'Every day at 9 AM', '', '⏳ Pending', ''],
    ['📍 Current Day Check (2 PM)', 'Mon-Fri at 2 PM', '', '⏳ Pending', ''],
    ['📍 Current Day Check (5 PM)', 'Mon-Fri at 5 PM', '', '⏳ Pending', ''],
    ['📊 Weekly Summary', 'Friday at 1 PM', '', '⏳ Pending', ''],
    ['', '', '', '', ''],  // Separator
    ['🖱️ Manual: Check Today', 'On-demand (button)', '', '⏳ Never Run', ''],
    ['🖱️ Manual: Check This Week', 'On-demand (button)', '', '⏳ Never Run', ''],
    ['🖱️ Manual: Check Last Week', 'On-demand (button)', '', '⏳ Never Run', '']
  ];
  
  statusSheet.getRange(2, 1, triggers.length, 5).setValues(triggers);
  for (let i = 2; i <= triggers.length + 1; i++) {
    statusSheet.setRowHeight(i, 30);
    statusSheet.getRange(i, 1, 1, 5).setVerticalAlignment('middle');
  }
  
  statusSheet.getRange(6, 1, 1, 5).setBackground('#f3f3f3');
}

function updateTriggerStatus(triggerName, violationsCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let statusSheet = ss.getSheetByName('Trigger Status');
    if (!statusSheet) {
      setupTriggerStatusSheet();
      statusSheet = ss.getSheetByName('Trigger Status');
    }
    
    const lastRow = statusSheet.getLastRow();
    if (lastRow <= 1) return;
    
    const data = statusSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let rowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      const rowName = data[i][0];
      if ((triggerName === 'Previous Day' && rowName.includes('Previous Day')) ||
          (triggerName === 'Today' && rowName.includes('2 PM')) ||
          (triggerName === 'Today-Evening' && rowName.includes('5 PM')) ||
          (triggerName === 'This Week (Mon-Fri)' && rowName.includes('Weekly Summary')) ||
          (triggerName === 'This Week (Mon-Today)' && rowName.includes('Manual: Check This Week')) ||
          (triggerName === 'Last Week (Mon-Sun)' && rowName.includes('Manual: Check Last Week')) ||
          (triggerName === 'Today-Manual' && rowName.includes('Manual: Check Today'))) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) return;
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
    statusSheet.getRange(rowIndex, 3).setValue(timestamp);
    statusSheet.getRange(rowIndex, 4).setValue('✅ Completed');
    statusSheet.getRange(rowIndex, 5).setValue(violationsCount);
    
    const statusCell = statusSheet.getRange(rowIndex, 4);
    if (violationsCount === 0) statusCell.setBackground('#d4edda');
    else if (violationsCount <= 10) statusCell.setBackground('#fff3cd');
    else statusCell.setBackground('#f8d7da');
    
  } catch (e) {
    log(`⚠️ Could not update trigger status: ${e.message}`);
  }
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    setupProfessionalSheets();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

function clearViolationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DRIVE_CONFIG.VIOLATIONS_SHEET_NAME);
  if (!sheet) return;
  
  // Always remove banding first
  const bandings = sheet.getBandings();
  bandings.forEach(b => b.remove());
  
  const lastRow = sheet.getLastRow();
  const frozenRows = sheet.getFrozenRows() || 0;
  const startRow = frozenRows + 1;
  
  if (lastRow >= startRow) {
    const lastCol = sheet.getLastColumn() || 10;
    const range = sheet.getRange(startRow, 1, lastRow - frozenRows, lastCol);
    range.clear();
    range.setBackground('#ffffff');
  }
}

// ========================================
// 📅 WEEK MANAGEMENT
// ========================================

function getWeekInfo(date) {
  // Get ISO week number
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  
  // Get week start (Monday) and end (Sunday)
  const dayOfWeek = date.getDay();
  const monday = new Date(date);
  monday.setDate(date.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1));
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  const formatDate = (d) => {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${months[d.getMonth()]} ${d.getDate()}`;
  };
  
  return {
    weekNumber: weekNo,
    year: d.getFullYear(),
    monday: monday,
    sunday: sunday,
    label: `📅 WEEK ${weekNo} (${formatDate(monday)}-${formatDate(sunday)}, ${d.getFullYear()})`
  };
}

function setupAutoCheckHeader(checkLabel, timestamp) {
  const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  const weekInfo = getWeekInfo(timestamp);
  
  // Check if week header already exists at top
  const firstCell = sheet.getRange(1, 1).getValue();
  const needsWeekHeader = !firstCell.toString().includes(`WEEK ${weekInfo.weekNumber}`);
  
  let insertRow = 3; // Default: insert after week header (row 1) and separator (row 2)
  
  if (needsWeekHeader) {
    // Insert week header at top
    sheet.insertRowBefore(1);
    sheet.insertRowBefore(1);
    
    const weekHeaderRange = sheet.getRange(1, 1, 1, 10);
    weekHeaderRange.merge();
    weekHeaderRange.setValue(weekInfo.label);
    weekHeaderRange.setBackground('#4285f4');
    weekHeaderRange.setFontColor('#ffffff');
    weekHeaderRange.setFontWeight('bold');
    weekHeaderRange.setFontSize(12);
    weekHeaderRange.setHorizontalAlignment('left');
    weekHeaderRange.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 35);
    
    // Add separator row
    sheet.setRowHeight(2, 10);
    sheet.getRange(2, 1, 1, 10).setBackground('#eeeeee');
  }
  
  // Insert 4 rows for this check block
  sheet.insertRowBefore(insertRow);
  sheet.insertRowBefore(insertRow);
  sheet.insertRowBefore(insertRow);
  sheet.insertRowBefore(insertRow);
  
  // IMPORTANT: Clear inherited formatting from all 4 new rows
  const newRowsRange = sheet.getRange(insertRow, 1, 4, 10);
  newRowsRange.setBackground(null);      // Clear background
  newRowsRange.setFontColor('#000000');  // Black text
  newRowsRange.setFontWeight('normal');  // Normal weight
  newRowsRange.setFontSize(10);          // Default size
  newRowsRange.breakApart();             // Unmerge any merged cells
  
  // Row 1: Check name (will be updated with count later)
  const checkNameRange = sheet.getRange(insertRow, 1, 1, 10);
  checkNameRange.merge();
  checkNameRange.setValue(`▶️ ${getFriendlyCheckName(checkLabel)} - ${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MMM dd, yyyy h:mm a')} (checking...)`);
  checkNameRange.setBackground('#f3f3f3');  // Light gray
  checkNameRange.setFontColor('#000000');   // Black text
  checkNameRange.setFontWeight('bold');
  checkNameRange.setFontSize(11);
  checkNameRange.setHorizontalAlignment('left');
  checkNameRange.setVerticalAlignment('middle');
  sheet.setRowHeight(insertRow, 30);
  
  // Row 2: Column headers
  const headers = ['Check Date', 'Timestamp', 'Type', 'Product', 'Name', 'Location', 'Violation', 'Owner', 'Link', 'Status'];
  const headerRange = sheet.getRange(insertRow + 1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#e8eaf6');   // Light purple
  headerRange.setFontColor('#000000');    // Black text
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  sheet.setRowHeight(insertRow + 1, 25);
  
  // Row 3: Divider (thin row)
  sheet.getRange(insertRow + 2, 1, 1, 10).setBackground('#eeeeee');
  sheet.setRowHeight(insertRow + 2, 5);
  
  // Row 4: Reserved for violations/no violations message (clear it)
  sheet.getRange(insertRow + 3, 1, 1, 10).setBackground(null);
  
  // Store current check row for updates
  PropertiesService.getScriptProperties().setProperty('CURRENT_CHECK_HEADER_ROW', insertRow.toString());
}

function updateAutoCheckHeader(violationsCount) {
  try {
    const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
    const headerRow = parseInt(PropertiesService.getScriptProperties().getProperty('CURRENT_CHECK_HEADER_ROW') || '0');
    
    if (headerRow === 0) return;
    
    const currentText = sheet.getRange(headerRow, 1).getValue().toString();
    const newText = currentText.replace('(checking...)', `(${violationsCount} violation${violationsCount === 1 ? '' : 's'})`);
    sheet.getRange(headerRow, 1).setValue(newText);
    
    // If no violations, add a "clean" row
    if (violationsCount === 0) {
      const insertRow = headerRow + 3; // After headers and divider
      sheet.insertRowAfter(insertRow - 1);
      const cleanMessageRange = sheet.getRange(insertRow, 1, 1, 10);
      cleanMessageRange.merge();
      cleanMessageRange.setValue('✅ No violations found - all files and folders are properly named!');
      cleanMessageRange.setBackground('#d4edda'); // Light green
      cleanMessageRange.setFontColor('#155724'); // Dark green
      cleanMessageRange.setFontWeight('normal');
      cleanMessageRange.setFontStyle('italic');
      cleanMessageRange.setHorizontalAlignment('center');
      cleanMessageRange.setVerticalAlignment('middle');
      sheet.setRowHeight(insertRow, 35);
    }
    
    PropertiesService.getScriptProperties().deleteProperty('CURRENT_CHECK_HEADER_ROW');
  } catch (e) {
    log(`⚠️ Could not update check header: ${e.message}`);
  }
}

function getFriendlyCheckName(label) {
  const mapping = {
    'Previous Day': 'Previous Day Check',
    'Today': 'Today Check (2 PM)',
    'Today-Evening': 'Today Check (5 PM)',
    'This Week (Mon-Fri)': 'Friday Summary'
  };
  return mapping[label] || label;
}

function cleanupWeekStatus() {
  try {
    const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
    const currentWeek = getWeekInfo(new Date());
    
    // Remove 🆕 NEW from violations not in current week
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    
    const statusColumn = sheet.getRange(1, 10, lastRow, 1);
    const values = statusColumn.getValues();
    
    let currentWeekLabel = null;
    for (let i = 0; i < values.length; i++) {
      const cellValue = sheet.getRange(i + 1, 1).getValue().toString();
      
      // Track which week we're in
      if (cellValue.includes('WEEK')) {
        const weekMatch = cellValue.match(/WEEK (\d+)/);
        if (weekMatch) {
          const weekNum = parseInt(weekMatch[1]);
          currentWeekLabel = weekNum === currentWeek.weekNumber ? 'current' : 'old';
        }
      }
      
      // Clear status if not in current week
      if (values[i][0] === '🆕 NEW' && currentWeekLabel === 'old') {
        sheet.getRange(i + 1, 10).setValue('');
      }
    }
  } catch (e) {
    log(`⚠️ Could not cleanup week status: ${e.message}`);
  }
}

function archiveOldWeeks() {
  try {
    const historySheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
    const archiveSheet = getOrCreateSheet(DRIVE_CONFIG.ARCHIVE_SHEET_NAME);
    const currentWeek = getWeekInfo(new Date());
    
    const lastRow = historySheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('ℹ️ No data to archive');
      return;
    }
    
    let archiveStartRow = null;
    let archiveCount = 0;
    
    // Find rows older than 12 weeks
    for (let i = 1; i <= lastRow; i++) {
      const cellValue = historySheet.getRange(i, 1).getValue().toString();
      
      if (cellValue.includes('WEEK')) {
        const weekMatch = cellValue.match(/WEEK (\d+)/);
        if (weekMatch) {
          const weekNum = parseInt(weekMatch[1]);
          const weekAge = currentWeek.weekNumber - weekNum;
          
          if (weekAge > DRIVE_CONFIG.WEEKS_TO_KEEP) {
            archiveStartRow = i;
            break;
          }
        }
      }
    }
    
    if (archiveStartRow) {
      // Copy to archive
      const rowsToArchive = lastRow - archiveStartRow + 1;
      const dataToArchive = historySheet.getRange(archiveStartRow, 1, rowsToArchive, 10).getValues();
      
      if (archiveSheet.getLastRow() === 0) {
        archiveSheet.getRange(1, 1, rowsToArchive, 10).setValues(dataToArchive);
      } else {
        archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive, 10).setValues(dataToArchive);
      }
      
      // Delete from history
      historySheet.deleteRows(archiveStartRow, rowsToArchive);
      archiveCount = rowsToArchive;
      
      SpreadsheetApp.getUi().alert(`✅ Archived ${archiveCount} rows to Archive sheet`);
      log(`✅ Archived ${archiveCount} rows of old data`);
    } else {
      SpreadsheetApp.getUi().alert('ℹ️ No data older than 12 weeks found');
    }
  } catch (e) {
    log(`❌ Archive failed: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ Archive failed: ${e.message}`);
  }
}

// ========================================
// 📝 VIOLATION LOGGING
// ========================================

function logViolation(timestamp, type, name, location, violation, owner, link, folder, skipTimeFilter) {
  // Determine if this is an auto-check or manual check
  if (currentCheckInfo && currentCheckInfo.isAutoCheck) {
    logAutoCheckViolation(timestamp, type, name, location, violation, owner, link, folder, skipTimeFilter);
  } else {
    logManualViolation(timestamp, type, name, location, violation, owner, link, folder, skipTimeFilter);
  }
}

function logManualViolation(timestamp, type, name, location, violation, owner, link, folder, skipTimeFilter) {
  const isDuplicate = violation.includes('Duplicate folder name');
  const key = isDuplicate ? `${location}|${name}|${violation}|${link}` : `${location}|${name}|${violation}`;
  
  if (loggedViolations.has(key)) return;
  
  const startDateStr = PropertiesService.getScriptProperties().getProperty('CHECK_START_DATE');
  const endDateStr = PropertiesService.getScriptProperties().getProperty('CHECK_END_DATE');
  const checkLabel = PropertiesService.getScriptProperties().getProperty('CHECK_LABEL') || 'Manual Check';
  
  if (!isDuplicate && !skipTimeFilter && startDateStr && endDateStr) {
    const startDate = new Date(parseInt(startDateStr));
    const endDate = new Date(parseInt(endDateStr));
    const itemDate = folder ? folder.getLastUpdated() : new Date();
    
    if (itemDate < startDate || itemDate > endDate) {
      log(`  ⏭️  Skipping violation outside date range: ${name}`);
      return;
    }
  }
  
  loggedViolations.add(key);
  
  const product = extractProductName(name, location);
  const sheet = getOrCreateSheet(DRIVE_CONFIG.VIOLATIONS_SHEET_NAME);
  const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
  
  sheet.appendRow([checkLabel, formattedTimestamp, type, product, name, location, violation, owner, link]);
  
  const lastRow = sheet.getLastRow();
  sheet.setRowHeight(lastRow, 21);
  sheet.getRange(lastRow, 1, 1, 10).setWrap(false);
  sheet.getRange(lastRow, 10).setValue('🆕 NEW');
  
  const rowRange = sheet.getRange(lastRow, 1, 1, 9);
  if (type === 'FOLDER') rowRange.setBackground('#fff3cd');
  else if (type === 'FILE') rowRange.setBackground('#d1ecf1');
}

function logAutoCheckViolation(timestamp, type, name, location, violation, owner, link, folder, skipTimeFilter) {
  const isDuplicate = violation.includes('Duplicate folder name');
  const key = isDuplicate ? `${location}|${name}|${violation}|${link}` : `${location}|${name}|${violation}`;
  
  if (loggedViolations.has(key)) return;
  
  const startDateStr = PropertiesService.getScriptProperties().getProperty('CHECK_START_DATE');
  const endDateStr = PropertiesService.getScriptProperties().getProperty('CHECK_END_DATE');
  const checkLabel = PropertiesService.getScriptProperties().getProperty('CHECK_LABEL') || 'Auto Check';
  
  if (!isDuplicate && !skipTimeFilter && startDateStr && endDateStr) {
    const startDate = new Date(parseInt(startDateStr));
    const endDate = new Date(parseInt(endDateStr));
    const itemDate = folder ? folder.getLastUpdated() : new Date();
    
    if (itemDate < startDate || itemDate > endDate) {
      log(`  ⏭️  Skipping violation outside date range: ${name}`);
      return;
    }
  }
  
  loggedViolations.add(key);
  
  const product = extractProductName(name, location);
  
  // Track breakdown in properties as we go
  const props = PropertiesService.getScriptProperties();
  const breakdownJson = props.getProperty('CURRENT_CHECK_BREAKDOWN') || '{}';
  const breakdown = JSON.parse(breakdownJson);
  breakdown[product] = (breakdown[product] || 0) + 1;
  props.setProperty('CURRENT_CHECK_BREAKDOWN', JSON.stringify(breakdown));
  
  const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
  
  // Find the row after the column headers for this check
  const headerRow = parseInt(props.getProperty('CURRENT_CHECK_HEADER_ROW') || '0');
  if (headerRow === 0) {
    // Fallback: append at end
    sheet.appendRow([checkLabel, formattedTimestamp, type, product, name, location, violation, owner, link, '🆕 NEW']);
    const lastRow = sheet.getLastRow();
    sheet.setRowHeight(lastRow, 21);
    const rowRange = sheet.getRange(lastRow, 1, 1, 9);
    if (type === 'FOLDER') rowRange.setBackground('#fff3cd');
    else if (type === 'FILE') rowRange.setBackground('#d1ecf1');
  } else {
    // Insert after header row + 2 (skip column headers and divider)
    const insertRow = headerRow + 3;
    sheet.insertRowAfter(insertRow - 1);
    sheet.getRange(insertRow, 1, 1, 10).setValues([[checkLabel, formattedTimestamp, type, product, name, location, violation, owner, link, '🆕 NEW']]);
    sheet.setRowHeight(insertRow, 21);
    sheet.getRange(insertRow, 1, 1, 10).setWrap(false);
    
    const rowRange = sheet.getRange(insertRow, 1, 1, 9);
    if (type === 'FOLDER') rowRange.setBackground('#fff3cd');
    else if (type === 'FILE') rowRange.setBackground('#d1ecf1');
  }
}

function extractProductName(name, location) {
  const aliases = getProductAliasesFromConfig_();
  
  // First check for exact product match
  for (const product of VALID_VALUES.PRODUCTS) {
    if (name.toUpperCase().startsWith(product.toUpperCase())) return product;
  }
  
  // Then check for aliases
  for (const [alias, canonical] of Object.entries(aliases)) {
    if (name.toUpperCase().startsWith(alias.toUpperCase())) return canonical;
  }
  
  // Check location for product
  for (const product of VALID_VALUES.PRODUCTS) {
    if (location.toUpperCase().includes(product.toUpperCase())) return product;
  }
  
  // Check location for aliases
  for (const [alias, canonical] of Object.entries(aliases)) {
    if (location.toUpperCase().includes(alias.toUpperCase())) return canonical;
  }
  
  return 'Unknown';
}

// [CONTINUED IN NEXT FILE DUE TO LENGTH - THIS INCLUDES ALL FOLDER CHECKING LOGIC]
// The rest of the code (checkFolder, validation functions, Slack integration, etc.) remains exactly the same

// ========================================
// 🔍 FOLDER CHECKING
// ========================================

function checkFolder(folder, depth, parentPath) {
  checkStopFlag();
  
  if (depth > DRIVE_CONFIG.MAX_DEPTH) {
    log(`⚠️  Max depth reached: ${parentPath}`);
    return;
  }
  
  const folderName = folder.getName();
  log(`  ${'  '.repeat(depth)}🔍 [DEPTH=${depth}] "${folderName}"`);
  
  const subfolderDepth = depth + 1;
  const isClairuRoot = folderName.includes('Clairu ADS');
  const isClairuModelFolder = depth === 1 && parentPath.includes('Clairu ADS');
  const isAdsRoot = folderName === '.ADS';
  const insideTellyStickOrFlixy = parentPath === '.ADS > TellyStick' || parentPath === '.ADS > Flixy' ||
                                   parentPath.startsWith('.ADS > TellyStick >') || 
                                   parentPath.startsWith('.ADS > Flixy >');
  
  const subfolders = folder.getFolders();
  const foldersByName = {};
  const allFolders = [];
  
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const subfolderName = subfolder.getName();
    
    if (!foldersByName[subfolderName]) {
      foldersByName[subfolderName] = [];
    }
    foldersByName[subfolderName].push(subfolder);
    allFolders.push(subfolder);
  }
  
  Object.keys(foldersByName).forEach(name => {
    const folders = foldersByName[name];
    if (folders.length > 1) {
      folders.sort((a, b) => a.getLastUpdated() - b.getLastUpdated());
      
      for (let i = 1; i < folders.length; i++) {
        const duplicateFolder = folders[i];
        log(`  ❌ DUPLICATE: "${name}" (created ${duplicateFolder.getLastUpdated().toLocaleDateString()}, original from ${folders[0].getLastUpdated().toLocaleDateString()})`);
        logViolation(new Date(), 'FOLDER', name, parentPath,
          'Duplicate folder name (multiple folders with same name in this location)',
          duplicateFolder.getOwner().getEmail(), duplicateFolder.getUrl(), duplicateFolder, true);
      }
    }
  });
  
  allFolders.forEach(subfolder => {
    const subfolderName = subfolder.getName();
    let isBatchFolder = false;
    
    if (isAdsRoot && subfolderDepth === 1 && (subfolderName === 'TellyStick' || subfolderName === 'Flixy')) {
      isBatchFolder = false;
    } else if (insideTellyStickOrFlixy && subfolderDepth === 2) {
      validateLocaleFolder(subfolder, parentPath);
      isBatchFolder = false;
    } else if (insideTellyStickOrFlixy && subfolderDepth === 3) {
      const result = validateYearOrBatchFolder(subfolder, parentPath);
      isBatchFolder = result.isBatch || result.skipYear;
    } else if (insideTellyStickOrFlixy && subfolderDepth === 4) {
      if (isOrganizationalFolder(subfolderName)) {
        isBatchFolder = false;
      } else {
        validateBatchFolder(subfolder, subfolderDepth, parentPath);
        isBatchFolder = true;
      }
    } else if (insideTellyStickOrFlixy && subfolderDepth >= 5) {
      validateBatchFolder(subfolder, subfolderDepth, parentPath);
      isBatchFolder = true;
    } else if (isClairuRoot && subfolderDepth === 1) {
      isBatchFolder = false;
    } else if (isClairuModelFolder && subfolderDepth === 2) {
      validateLocaleFolder(subfolder, parentPath);
      isBatchFolder = false;
    } else if (isClairuModelFolder && subfolderDepth === 3) {
      const result = validateYearOrBatchFolder(subfolder, parentPath);
      isBatchFolder = result.isBatch || result.skipYear;
    } else if (isClairuModelFolder && subfolderDepth === 4) {
      if (isOrganizationalFolder(subfolderName)) {
        isBatchFolder = false;
      } else {
        validateBatchFolder(subfolder, subfolderDepth, parentPath);
        isBatchFolder = true;
      }
    } else if (isClairuModelFolder && subfolderDepth >= 5) {
      validateBatchFolder(subfolder, subfolderDepth, parentPath);
      isBatchFolder = true;
    } else if (subfolderDepth === 1) {
      validateLocaleFolder(subfolder, parentPath);
    } else if (subfolderDepth === 2) {
      const result = validateYearOrBatchFolder(subfolder, parentPath);
      isBatchFolder = result.isBatch || result.skipYear;
    } else if (subfolderDepth === 3) {
      if (isOrganizationalFolder(subfolderName)) {
        isBatchFolder = false;
      } else {
        validateBatchFolder(subfolder, subfolderDepth, parentPath);
        isBatchFolder = true;
      }
    } else if (subfolderDepth >= 4) {
      validateBatchFolder(subfolder, subfolderDepth, parentPath);
      isBatchFolder = true;
    }
    
    if (!isBatchFolder) {
      checkFolder(subfolder, subfolderDepth, `${parentPath} > ${subfolderName}`);
    }
  });
}

function isOrganizationalFolder(folderName) {
  const name = folderName.trim();
  return /^\d{2}W\d{1,2}(-\d{1,2})?$/i.test(name) ||
         /^\.?\d{2}\s*[A-Za-z]{3}\d{2,4}$/.test(name) ||
         /^\.?\d{2}\s+\d{4}$/.test(name);
}

function looksLikeBatchFolder(name) {
  if (!name.includes('_')) return false;
  const parts = name.split('_');
  if (parts.length < 4) return false;
  return VALID_VALUES.PLATFORMS.includes(parts[1]);
}

function isValidLocaleFolder(folderName) {
  const name = folderName.toUpperCase().trim();
  if (VALID_VALUES.LOCALES.includes(name)) return true;
  if (name.startsWith('.') && VALID_VALUES.LOCALES.includes(name.substring(1))) return true;
  for (const locale of VALID_VALUES.LOCALES) {
    if (name.includes(locale)) return true;
  }
  for (const [country, locale] of Object.entries(COUNTRY_TO_LOCALE)) {
    if (name.includes(country)) return true;
  }
  return false;
}

function validateLocaleFolder(folder, parentPath) {
  const name = folder.getName();
  if (looksLikeBatchFolder(name)) {
    logViolation(new Date(), 'FOLDER', name, parentPath,
      'Batch folder in wrong location (should be inside Locale > Year folders)',
      folder.getOwner().getEmail(), folder.getUrl(), folder);
    return;
  }
  if (!isValidLocaleFolder(name)) {
    logViolation(new Date(), 'FOLDER', name, parentPath,
      `Invalid locale folder - should be one of: ${VALID_VALUES.LOCALES.join(', ')}`,
      folder.getOwner().getEmail(), folder.getUrl(), folder);
  }
}

function validateYearOrBatchFolder(folder, parentPath) {
  const name = folder.getName();
  
  if (/^\d{4}$/.test(name)) {
    const year = parseInt(name);
    if (DRIVE_CONFIG.SKIP_OLD_YEARS) {
      const oldestYearToCheck = DRIVE_CONFIG.CURRENT_YEAR - DRIVE_CONFIG.YEARS_TO_CHECK;
      if (year < oldestYearToCheck) {
        log(`  ⏭️  Skipping old year: ${name}`);
        return { isBatch: false, skipYear: true };
      }
    }
    return { isBatch: false, skipYear: false };
  }
  
  if (name.includes('_')) {
    validateBatchFolder(folder, 2, parentPath);
    return { isBatch: true, skipYear: false };
  }
  
  logViolation(new Date(), 'FOLDER', name, parentPath,
    'Expected year folder (YYYY) or batch folder (Product_Platform_Locale_Batch)',
    folder.getOwner().getEmail(), folder.getUrl(), folder);
  
  return { isBatch: false, skipYear: false };
}

function validateBatchFolder(folder, depth, parentPath) {
  const name = folder.getName();
  
  // Skip year folders - they are not batch folders
  if (/^\d{4}$/.test(name)) {
    log(`  ⏭️  Skipping year folder validation: ${name}`);
    return;
  }
  
  log(`  ✓ Validating batch: ${name}`);
  
  const parts = name.split('_');
  const violations = [];
  
  // Check year placement - extract year from path
  const yearMatch = parentPath.match(/>\s*(\d{4})\s*(?:>|$)/);
  if (yearMatch) {
    const yearViolation = validateYearPlacement(folder, 'Folder', yearMatch[1], parentPath);
    if (yearViolation) {
      violations.push(yearViolation);
    }
  }
  
  if (name !== name.trim()) {
    if (name.startsWith(' ')) {
      violations.push('Folder name has leading space(s) - should be removed');
    }
    if (name.endsWith(' ')) {
      violations.push('Folder name has trailing space(s) - should be removed');
    }
  }
  
  if (parts.length < 4) {
    violations.push('Invalid format - expected: Product_Platform_Locale_Batch');
  } else {
    const sortedProducts = [...VALID_VALUES.PRODUCTS].sort((a, b) => b.length - a.length);
    
    let productName = null;
    let productPartsCount = 0;
    
    for (const product of sortedProducts) {
      if (name.toUpperCase().startsWith(product.toUpperCase() + '_')) {
        productName = product;
        productPartsCount = product.split('_').length;
        break;
      }
    }
    
    if (!productName) {
      const firstPart = parts[0];
      if (firstPart.trim() !== firstPart) {
        if (firstPart.startsWith(' ')) {
          violations.push('Product name has leading space - remove the space before the product name');
        } else if (firstPart.endsWith(' ')) {
          violations.push('Product name has trailing space - remove the space after the product name');
        }
      } else {
        violations.push(`Unknown product: "${parts[0]}"`);
      }
      productPartsCount = 1;
    } else {
      const productPart = parts.slice(0, productPartsCount).join('_');
      if (productPart !== productName) {
        violations.push(`Wrong case: "${productPart}" should be "${productName}"`);
      }
    }
    
    const platformPart = parts[productPartsCount];
    const localePart = parts[productPartsCount + 1];
    const batchPart = parts[productPartsCount + 2];
    
    if (platformPart && !VALID_VALUES.PLATFORMS.includes(platformPart)) {
      violations.push(`Invalid platform: "${platformPart}"`);
    }
    if (localePart && !VALID_VALUES.LOCALES.includes(localePart)) {
      violations.push(`Invalid locale: "${localePart}"`);
    }
    if (batchPart && !/^\d+$/.test(batchPart)) {
      violations.push(`Invalid batch number: "${batchPart}" (should be numeric)`);
    }
  }
  
  violations.push(...checkLazyPatterns(name));
  
  if (violations.length > 0) {
    logViolation(new Date(), 'FOLDER', name, parentPath, violations.join(' | '),
      folder.getOwner().getEmail(), folder.getUrl(), folder);
  }
  
  const batchDate = folder.getLastUpdated();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);
  
  if (batchDate >= cutoffDate) {
    validateBatchFiles(folder, name, parentPath);
  } else {
    log(`    ⏭️  Skipping file validation (batch older than 30 days)`);
  }
}

function checkLazyPatterns(name) {
  const violations = [];
  const beforeParens = name.split('_(')[0].toLowerCase();
  const afterParens = name.includes(')') ? name.split(')').pop().toLowerCase() : '';
  const outsideParens = beforeParens + afterParens;
  
  for (const word of LAZY_PATTERNS.FORBIDDEN_WORDS) {
    if (outsideParens.includes(word)) {
      if (word === 'old' && (outsideParens.includes('mold') || outsideParens.includes('fold') || 
                             outsideParens.includes('bold') || outsideParens.includes('hold'))) {
        continue;
      }
      violations.push(`Contains forbidden word: "${word}"`);
    }
  }
  
  for (const pattern of LAZY_PATTERNS.TEST_PATTERNS) {
    if (pattern.test(name)) {
      violations.push('Contains test-related pattern');
      break;
    }
  }
  
  for (const pattern of LAZY_PATTERNS.COPY_PATTERNS) {
    if (pattern.test(name)) {
      violations.push('Appears to be a copied folder/file');
      break;
    }
  }
  
  return violations;
}

function validateBatchFiles(folder, batchName, batchPath) {
  try {
    const files = folder.getFiles();
    const fileNamesInBatch = {};
    
    const startDateStr = PropertiesService.getScriptProperties().getProperty('CHECK_START_DATE');
    const endDateStr = PropertiesService.getScriptProperties().getProperty('CHECK_END_DATE');
    
    let startDate, endDate;
    if (startDateStr && endDateStr) {
      startDate = new Date(parseInt(startDateStr));
      endDate = new Date(parseInt(endDateStr));
    } else {
      startDate = new Date();
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date();
    }
    
    // Extract year from path for year validation
    const yearMatch = batchPath.match(/>\s*(\d{4})\s*(?:>|$)/);
    const pathYear = yearMatch ? yearMatch[1] : null;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const fileDate = file.getLastUpdated();
      
      const lowerName = fileName.toLowerCase();
      if (lowerName.endsWith('.zip') || lowerName.endsWith('.rar') || lowerName.endsWith('.7z') || 
          lowerName.endsWith('.tar') || lowerName.endsWith('.gz') || lowerName.endsWith('.bz2') ||
          lowerName.endsWith('.tar.gz') || lowerName.endsWith('.tar.bz2')) {
        continue;
      }
      
      // Only process files within the date range
      if (fileDate >= startDate && fileDate <= endDate) {
        // Track total files scanned (only within date range)
        const props = PropertiesService.getScriptProperties();
        const currentCount = parseInt(props.getProperty('TOTAL_FILES_SCANNED') || '0');
        props.setProperty('TOTAL_FILES_SCANNED', (currentCount + 1).toString());
        
        const fileViolations = [];
        
        // Check year placement
        if (pathYear) {
          const yearViolation = validateYearPlacement(file, 'File', pathYear, batchPath);
          if (yearViolation) {
            fileViolations.push(yearViolation);
          }
        }
        
        if (fileNamesInBatch[fileName]) {
          fileViolations.push('Duplicate file name in this batch');
        } else {
          fileNamesInBatch[fileName] = true;
        }
        
        for (const pattern of LAZY_PATTERNS.COPY_PATTERNS) {
          if (pattern.test(fileName)) {
            if (pattern.source.includes('^copy of')) {
              fileViolations.push('File name starts with "Copy of" - should be renamed');
            } else if (pattern.source.includes('\\(\\d+\\)\\.')) {
              const match = fileName.match(/ \((\d+)\)\./);
              if (match) {
                fileViolations.push(`File has copy number "(${match[1]})" before extension - should be removed`);
              } else {
                fileViolations.push('File has copy number pattern - should be removed');
              }
            } else if (pattern.source.includes('\\(\\d+\\)$')) {
              const match = fileName.match(/ \((\d+)\)$/);
              if (match) {
                fileViolations.push(`File ends with "(${match[1]})" - should be removed`);
              } else {
                fileViolations.push('File has copy number pattern - should be removed');
              }
            } else if (pattern.source.includes('- copy$')) {
              fileViolations.push('File name ends with " - copy" - should be renamed');
            }
            break;
          }
        }
        
        const namingViolations = validateFileName(fileName, batchName);
        fileViolations.push(...namingViolations);
        
        if (fileViolations.length > 0) {
          log(`    ❌ FILE: ${fileName}`);
          logViolation(new Date(), 'FILE', fileName, batchPath, fileViolations.join(' | '),
            file.getOwner().getEmail(), file.getUrl(), file);
        }
      } else {
        // Still track duplicate names for files outside date range
        if (!fileNamesInBatch[fileName]) {
          fileNamesInBatch[fileName] = true;
        }
      }
    }
  } catch (e) {
    log(`    ⚠️  Error checking files: ${e.message}`);
  }
}

function validateFileName(fileName, batchName) {
  const violations = [];
  
  if (/^copy of /i.test(fileName)) return violations;
  
  const lowerName = fileName.toLowerCase();
  if (lowerName.endsWith('.zip') || lowerName.endsWith('.rar') || lowerName.endsWith('.7z') || 
      lowerName.endsWith('.tar') || lowerName.endsWith('.gz') || lowerName.endsWith('.bz2') ||
      lowerName.endsWith('.tar.gz') || lowerName.endsWith('.tar.bz2') ||
      lowerName.endsWith('.txt') || lowerName.endsWith('.pdf')) {
    return violations;
  }
  
  const fileNameNoExt = fileName.replace(/\.[^.]+$/, '');
  if (fileNameNoExt !== fileNameNoExt.trim()) {
    if (fileNameNoExt.startsWith(' ')) {
      violations.push('File name has leading space(s) - should be removed');
    }
    if (fileNameNoExt.endsWith(' ')) {
      violations.push('File name has trailing space(s) before extension - should be removed');
    }
    return violations;
  }
  
  const batchParts = batchName.split('_');
  if (batchParts.length < 4) return violations;
  
  const sortedProducts = [...VALID_VALUES.PRODUCTS].sort((a, b) => b.length - a.length);
  let batchProduct = null;
  let productPartsCount = 0;
  
  for (const product of sortedProducts) {
    if (batchName.toUpperCase().startsWith(product.toUpperCase() + '_')) {
      batchProduct = product;
      productPartsCount = product.split('_').length;
      break;
    }
  }
  
  if (!batchProduct) return violations;
  
  const batchPlatform = batchParts[productPartsCount];
  const batchLocale = batchParts[productPartsCount + 1];
  const batchNumber = batchParts[productPartsCount + 2];
  
  const fileParts = fileNameNoExt.split('_');
  
  if (fileParts.length < 4) return violations;
  
  let fileProduct = null;
  let fileProductPartsCount = 0;
  
  for (const product of sortedProducts) {
    const productParts = product.split('_').length;
    const fileProductCandidate = fileParts.slice(0, productParts).join('_');
    
    if (fileProductCandidate.toUpperCase() === product.toUpperCase()) {
      fileProduct = product;
      fileProductPartsCount = productParts;
      
      if (fileProductCandidate !== product) {
        violations.push(`Wrong product case: "${fileProductCandidate}" should be "${product}"`);
      }
      break;
    }
  }
  
  if (!fileProduct) {
    const firstPart = fileParts[0];
    if (firstPart.trim() !== firstPart) {
      if (firstPart.startsWith(' ')) {
        violations.push('Product name has leading space - remove the space before the product name');
      } else if (firstPart.endsWith(' ')) {
        violations.push('Product name has trailing space - remove the space after the product name');
      }
    } else {
      violations.push(`Unknown or misspelled product: "${fileParts[0]}"`);
    }
    return violations;
  }
  
  if (fileParts.length > fileProductPartsCount) {
    const filePlatform = fileParts[fileProductPartsCount];
    if (VALID_VALUES.PLATFORMS.includes(filePlatform) && filePlatform !== batchPlatform) {
      violations.push(`Platform mismatch: file has "${filePlatform}" but batch is "${batchPlatform}"`);
    }
  }
  
  if (fileParts.length > fileProductPartsCount + 1) {
    const fileLocale = fileParts[fileProductPartsCount + 1];
    if (VALID_VALUES.LOCALES.includes(fileLocale) && fileLocale !== batchLocale) {
      violations.push(`Locale mismatch: file has "${fileLocale}" but batch is "${batchLocale}"`);
    }
  }
  
  if (fileParts.length > fileProductPartsCount + 2) {
    const fileBatchNum = fileParts[fileProductPartsCount + 2];
    if (/^\d+$/.test(fileBatchNum) && fileBatchNum !== batchNumber) {
      violations.push(`Batch number mismatch: file has "${fileBatchNum}" but batch is "${batchNumber}"`);
    }
  }
  
  return violations;
}

// ========================================
// 💬 SLACK INTEGRATION
// ========================================

function getSlackSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const data = sheet.getRange('A2:B6').getValues();
    
    const settings = {};
    data.forEach(row => {
      const key = row[0];
      const value = row[1];
      if (key.startsWith('Drive Checker:')) {
        const shortKey = key.replace('Drive Checker: ', '').replace(/ /g, '_').toLowerCase();
        settings[shortKey] = value;
      }
    });
    
    return settings;
  } catch (e) {
    log(`⚠️ Could not load Slack settings: ${e.message}`);
    return null;
  }
}

function sendSlackNotification(checkLabel, violationsCount) {
  const settings = getSlackSettings();
  if (!settings || !settings.slack_bot_token || !settings.channel_id) {
    log('⚠️ Slack settings not configured, skipping notification');
    return;
  }
  
  if (!settings.post_manual_checks && 
      (checkLabel.includes('Manual') || checkLabel === 'Today-Manual')) {
    log('⏭️ Skipping Slack notification for manual check');
    return;
  }
  
  const breakdown = getViolationBreakdown();
  log(`🔍 DEBUG: Breakdown = ${JSON.stringify(breakdown)}`);
  log(`🔍 DEBUG: Violations Count = ${violationsCount}`);
  const isWeeklySummary = checkLabel === 'This Week (Mon-Fri)';
  const message = buildSlackMessage(checkLabel, violationsCount, breakdown, isWeeklySummary);
  const response = postToSlack(settings.slack_bot_token, settings.channel_id, message);
  
  if (response && response.ts) {
    const props = PropertiesService.getScriptProperties();
    const messagesJson = props.getProperty('SLACK_MESSAGE_HISTORY') || '[]';
    const messages = JSON.parse(messagesJson);
    messages.push(response.ts);
    props.setProperty('SLACK_MESSAGE_HISTORY', JSON.stringify(messages));
  }
}

function getViolationBreakdown() {
  // For auto-checks, use the in-memory breakdown tracked during the check
  if (currentCheckInfo && currentCheckInfo.isAutoCheck) {
    const props = PropertiesService.getScriptProperties();
    const breakdownJson = props.getProperty('CURRENT_CHECK_BREAKDOWN') || '{}';
    const breakdown = JSON.parse(breakdownJson);
    
    log(`🔍 getViolationBreakdown: Using in-memory breakdown = ${JSON.stringify(breakdown)}`);
    return breakdown;
  } else {
    // For manual checks, count all violations in the sheet
    const sheet = getOrCreateSheet(DRIVE_CONFIG.VIOLATIONS_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return {};
    
    const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const breakdown = {};
    
    data.forEach(row => {
      const product = row[0] || '';
      if (product && product !== '') {
        breakdown[product] = (breakdown[product] || 0) + 1;
      }
    });
    
    return breakdown;
  }
}

// ========================================
// 📊 ENHANCED ANALYTICS FUNCTIONS
// ========================================

function getWeekStats() {
  const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  const weekInfo = getWeekInfo(new Date());
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return { totalViolations: 0, totalChecks: 0 };
  
  let totalViolations = 0;
  let totalChecks = 0;
  let inCurrentWeek = false;
  
  for (let i = 1; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 1).getValue().toString();
    
    // Check if we're in current week
    if (cellValue.includes(`WEEK ${weekInfo.weekNumber}`)) {
      inCurrentWeek = true;
      continue;
    }
    
    // Stop if we hit next week
    if (inCurrentWeek && cellValue.includes('WEEK') && !cellValue.includes(`WEEK ${weekInfo.weekNumber}`)) {
      break;
    }
    
    // Count checks
    if (inCurrentWeek && cellValue.includes('▶️')) {
      totalChecks++;
    }
    
    // Count violations (rows with actual data)
    if (inCurrentWeek) {
      const type = sheet.getRange(i, 3).getValue().toString();
      if (type === 'FILE' || type === 'FOLDER') {
        totalViolations++;
      }
    }
  }
  
  return { totalViolations, totalChecks };
}

function getFixedCount() {
  const props = PropertiesService.getScriptProperties();
  const lastViolationsJson = props.getProperty('LAST_CHECK_VIOLATIONS') || '[]';
  const lastViolations = new Set(JSON.parse(lastViolationsJson));
  
  // Get current violations
  const currentViolations = new Set(Array.from(loggedViolations));
  
  // Count how many from last check are NOT in current check
  let fixedCount = 0;
  lastViolations.forEach(violation => {
    if (!currentViolations.has(violation)) {
      fixedCount++;
    }
  });
  
  // Store current violations for next check
  props.setProperty('LAST_CHECK_VIOLATIONS', JSON.stringify(Array.from(currentViolations)));
  
  return fixedCount;
}

function getTotalFilesScanned() {
  // This is tracked during the check - retrieve from properties
  const props = PropertiesService.getScriptProperties();
  return parseInt(props.getProperty('TOTAL_FILES_SCANNED') || '0');
}

function getComplianceRate(violationsCount, totalFiles) {
  if (totalFiles === 0) return 100;
  return ((totalFiles - violationsCount) / totalFiles) * 100;
}

function buildComplianceBar(percentage) {
  const filledBlocks = Math.round(percentage / 10);
  const emptyBlocks = 10 - filledBlocks;
  return '▓'.repeat(filledBlocks) + '░'.repeat(emptyBlocks);
}

function getWeeklySummaryData() {
  const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  const weekInfo = getWeekInfo(new Date());
  const lastRow = sheet.getLastRow();
  
  const summary = {
    totalChecks: 0,
    totalViolations: 0,
    totalFilesScanned: 0,
    dailyBreakdown: { Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0 },
    topIssues: {},
    topProducts: {},
    teamPerformance: {}
  };
  
  if (lastRow <= 1) return summary;
  
  let inCurrentWeek = false;
  let currentCheckDate = null;
  
  for (let i = 1; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 1).getValue().toString();
    
    if (cellValue.includes(`WEEK ${weekInfo.weekNumber}`)) {
      inCurrentWeek = true;
      continue;
    }
    
    if (inCurrentWeek && cellValue.includes('WEEK') && !cellValue.includes(`WEEK ${weekInfo.weekNumber}`)) {
      break;
    }
    
    if (inCurrentWeek && cellValue.includes('▶️')) {
      summary.totalChecks++;
      // Extract date from check header
      const dateMatch = cellValue.match(/(\w{3}) (\d{1,2}), (\d{4})/);
      if (dateMatch) {
        const checkDate = new Date(`${dateMatch[1]} ${dateMatch[2]}, ${dateMatch[3]}`);
        const dayName = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][checkDate.getDay()];
        currentCheckDate = dayName;
      }
    }
    
    if (inCurrentWeek) {
      const type = sheet.getRange(i, 3).getValue();
      if (type === 'FILE' || type === 'FOLDER') {
        summary.totalViolations++;
        
        // Track by day
        if (currentCheckDate && summary.dailyBreakdown[currentCheckDate] !== undefined) {
          summary.dailyBreakdown[currentCheckDate]++;
        }
        
        // Track by product
        const product = sheet.getRange(i, 4).getValue() || 'Unknown';
        summary.topProducts[product] = (summary.topProducts[product] || 0) + 1;
        
        // Track by violation type
        const violation = sheet.getRange(i, 7).getValue() || '';
        const issueType = violation.split(':')[0] || violation;
        summary.topIssues[issueType] = (summary.topIssues[issueType] || 0) + 1;
        
        // Track by owner
        const owner = sheet.getRange(i, 8).getValue() || 'Unknown';
        summary.teamPerformance[owner] = (summary.teamPerformance[owner] || 0) + 1;
      }
    }
  }
  
  return summary;
}

function buildSlackMessage(checkLabel, violationsCount, breakdown, isWeeklySummary = false) {
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const isAutoCheck = currentCheckInfo && currentCheckInfo.isAutoCheck;
  const sheetName = isAutoCheck ? DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME : DRIVE_CONFIG.VIOLATIONS_SHEET_NAME;
  const sheet = getOrCreateSheet(sheetName);
  const sheetId = sheet.getSheetId();
  const violationsUrl = `${spreadsheetUrl}#gid=${sheetId}`;
  
  let breakdownText = '';
  if (violationsCount > 0 && Object.keys(breakdown).length > 0) {
    const sorted = Object.entries(breakdown).sort((a, b) => b[1] - a[1]);
    breakdownText = '\n\n*Breakdown by Product:*\n';
    sorted.forEach(([product, count]) => {
      breakdownText += `• ${product}: ${count}\n`;
    });
  }
  
  let statusEmoji = '✅';
  let color = '#10b981';
  if (violationsCount > 10) {
    statusEmoji = '🔴';
    color = '#ef4444';
  } else if (violationsCount > 0) {
    statusEmoji = '⚠️';
    color = '#f59e0b';
  }
  
  const timestamp = new Date().toLocaleString('en-US', { 
  timeZone: 'Europe/Vilnius',  // ← Your timezone
  dateStyle: 'medium',
  timeStyle: 'short'
});
  
  return {
    text: `Drive Checker: ${checkLabel} - ${violationsCount} violations`,
    blocks: isWeeklySummary ? buildWeeklySummaryBlocks(checkLabel, violationsCount) : buildDailyCheckBlocks(checkLabel, violationsCount, breakdown),
    attachments: [
      {
        color: color,
        blocks: []
      }
    ]
  };
}

function buildDailyCheckBlocks(checkLabel, violationsCount, breakdown) {
  log(`🔍 buildDailyCheckBlocks called with: violationsCount=${violationsCount}, breakdown=${JSON.stringify(breakdown)}`);
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const isAutoCheck = currentCheckInfo && currentCheckInfo.isAutoCheck;
  const sheetName = isAutoCheck ? DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME : DRIVE_CONFIG.VIOLATIONS_SHEET_NAME;
  const sheet = getOrCreateSheet(sheetName);
  const sheetId = sheet.getSheetId();
  const violationsUrl = `${spreadsheetUrl}#gid=${sheetId}`;
  
  // Status emoji and color
  let statusEmoji = '✅';
  let color = '#10b981';
  if (violationsCount > 10) {
    statusEmoji = '🔴';
    color = '#ef4444';
  } else if (violationsCount > 0) {
    statusEmoji = '⚠️';
    color = '#f59e0b';
  }
  
  const timestamp = new Date().toLocaleString('en-US', { 
    timeZone: 'Europe/Vilnius',
    dateStyle: 'medium',
    timeStyle: 'short'
  });
  
  // Get analytics
  const weekStats = getWeekStats();
  const fixedCount = getFixedCount();
  const totalFiles = getTotalFilesScanned();
  const complianceRate = getComplianceRate(violationsCount, totalFiles);
  const complianceBar = buildComplianceBar(complianceRate);
  
  // Build breakdown text
  let breakdownText = '';
  if (violationsCount > 0 && Object.keys(breakdown).length > 0) {
    const sorted = Object.entries(breakdown).sort((a, b) => b[1] - a[1]);
    breakdownText = '*Breakdown by Product:*\n';  // ← Removed \n\n
    sorted.forEach(([product, count]) => {
      breakdownText += `• ${product}: ${count}\n`;
    });
  }
  
  // Build week stats text with fixed count
  let weekStatsText = `📅 *This Week:* ${weekStats.totalViolations} total violations across ${weekStats.totalChecks} checks`;
  if (fixedCount > 0) {
    weekStatsText += `\n✅ *${fixedCount} violation${fixedCount === 1 ? '' : 's'} fixed* since last check`;
  }
  
  const blocks = [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: `${statusEmoji} Drive Checker: ${checkLabel}`,
        emoji: true
      }
    },
    {
      type: 'section',
      fields: [
        {
          type: 'mrkdwn',
          text: `*Status:*\n${violationsCount === 0 ? '✅ Clean' : `⚠️ ${violationsCount} violations found`}`
        },
        {
          type: 'mrkdwn',
          text: `*Checked:*\n${timestamp}`
        }
      ]
    }
  ];
  
  // Add compliance bar if we have file count
  if (totalFiles > 0) {
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `📊 *Compliance Rate*\n\`${complianceBar}\` ${complianceRate.toFixed(1)}% (${totalFiles - violationsCount}/${totalFiles} files correct)`
      }
    });
  }
  
  // Add week stats
  blocks.push({
    type: 'context',
    elements: [
      {
        type: 'mrkdwn',
        text: weekStatsText
      }
    ]
  });
  
  // Add breakdown
  if (violationsCount > 0) {
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: breakdownText.length > 0 ? breakdownText : '*Breakdown by Product:*\n_No product data available_'
      }
    });
  } else {
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '_No violations found_'
      }
    });
  }
  
  // Divider and button
  blocks.push({ type: 'divider' });
  blocks.push({
    type: 'actions',
    elements: [
      {
        type: 'button',
        text: {
          type: 'plain_text',
          text: '📊 View Violations Sheet',
          emoji: true
        },
        url: violationsUrl,
        style: violationsCount > 0 ? 'danger' : 'primary'
      }
    ]
  });
  
  return blocks;
}

function buildWeeklySummaryBlocks(checkLabel, violationsCount) {
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
  const sheetId = sheet.getSheetId();
  const violationsUrl = `${spreadsheetUrl}#gid=${sheetId}`;
  const weekInfo = getWeekInfo(new Date());
  const summary = getWeeklySummaryData();
  
  // Format daily breakdown
  const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  let dailyText = '*Daily Breakdown:*\n';
  days.forEach(day => {
    const count = summary.dailyBreakdown[day] || 0;
    const icon = count === 0 ? '✅' : '⚠️';
    dailyText += `${icon} ${day}: ${count} violation${count === 1 ? '' : 's'}\n`;
  });
  
  // Format top issues
  const topIssues = Object.entries(summary.topIssues)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  let issuesText = '*Most Common Problems:*\n';
  topIssues.forEach(([issue, count], index) => {
    issuesText += `${index + 1}. ${issue}: ${count} occurrence${count === 1 ? '' : 's'}\n`;
  });
  
  // Format top products
  const topProducts = Object.entries(summary.topProducts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  let productsText = '*Products Needing Attention:*\n';
  topProducts.forEach(([product, count], index) => {
    productsText += `${index + 1}. ${product}: ${count} violation${count === 1 ? '' : 's'}\n`;
  });
  
  // Format team performance
  const teamPerf = Object.entries(summary.teamPerformance)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  let teamText = '*Contributors:*\n';
  teamPerf.forEach(([owner, count]) => {
    const shortName = owner.split('@')[0];
    teamText += `• ${shortName}: ${count} violation${count === 1 ? '' : 's'}\n`;
  });
  
  const blocks = [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: `📊 ${weekInfo.label}`,
        emoji: true
      }
    },
    {
      type: 'section',
      fields: [
        {
          type: 'mrkdwn',
          text: `*Total Checks:*\n${summary.totalChecks}`
        },
        {
          type: 'mrkdwn',
          text: `*Total Violations:*\n${summary.totalViolations}`
        }
      ]
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: dailyText
      }
    },
    { type: 'divider' },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: issuesText
      }
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: productsText
      }
    },
    { type: 'divider' },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: teamText
      }
    },
    {
      type: 'actions',
      elements: [
        {
          type: 'button',
          text: {
            type: 'plain_text',
            text: '📊 View Full History',
            emoji: true
          },
          url: violationsUrl,
          style: 'primary'
        }
      ]
    }
  ];
  
  return blocks;
}

function postToSlack(token, channel, message) {
  try {
    const url = 'https://slack.com/api/chat.postMessage';
    const payload = {
      channel: channel,
      ...message
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      log(`✅ Slack message posted successfully`);
      return result;
    } else {
      log(`❌ Slack post failed: ${result.error}`);
      return null;
    }
  } catch (e) {
    log(`❌ Slack error: ${e.message}`);
    return null;
  }
}

function deleteSlackMessage(token, channel, messageTs) {
  try {
    const url = 'https://slack.com/api/chat.delete';
    const payload = {
      channel: channel,
      ts: messageTs
    };
    
    UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    log(`🗑️ Deleted old Slack message`);
  } catch (e) {
    log(`⚠️ Could not delete old message: ${e.message}`);
  }
}

function deleteLastSlackMessage() {
  const settings = getSlackSettings();
  if (!settings || !settings.slack_bot_token || !settings.channel_id) {
    SpreadsheetApp.getUi().alert('❌ Slack settings not configured');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Delete Last Message',
    'This will delete the most recent message from the bot in the channel. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    const historyUrl = `https://slack.com/api/conversations.history?channel=${settings.channel_id}&limit=50`;
    const historyResponse = UrlFetchApp.fetch(historyUrl, {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${settings.slack_bot_token}`
      },
      muteHttpExceptions: true
    });
    
    const historyData = JSON.parse(historyResponse.getContentText());
    
    if (!historyData.ok) {
      ui.alert(`❌ Failed to fetch messages: ${historyData.error}`);
      return;
    }
    
    const botUserId = getBotUserId(settings.slack_bot_token);
    const botMessages = historyData.messages.filter(msg => msg.user === botUserId || msg.bot_id);
    
    if (botMessages.length === 0) {
      ui.alert('ℹ️ No bot messages found in the channel');
      return;
    }
    
    const lastMessage = botMessages[0];
    deleteSlackMessage(settings.slack_bot_token, settings.channel_id, lastMessage.ts);
    
    ui.alert('✅ Last bot message deleted!');
    
  } catch (e) {
    ui.alert(`❌ Error: ${e.message}`);
  }
}

function selectAndDeleteSlackMessage() {
  const settings = getSlackSettings();
  if (!settings || !settings.slack_bot_token || !settings.channel_id) {
    SpreadsheetApp.getUi().alert('❌ Slack settings not configured');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    const historyUrl = `https://slack.com/api/conversations.history?channel=${settings.channel_id}&limit=100`;
    const historyResponse = UrlFetchApp.fetch(historyUrl, {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${settings.slack_bot_token}`
      },
      muteHttpExceptions: true
    });
    
    const historyData = JSON.parse(historyResponse.getContentText());
    
    if (!historyData.ok) {
      ui.alert(`❌ Failed to fetch messages: ${historyData.error}`);
      return;
    }
    
    const botUserId = getBotUserId(settings.slack_bot_token);
    const botMessages = historyData.messages.filter(msg => 
      msg.user === botUserId || msg.bot_id || (msg.text && msg.text.includes('Drive Checker'))
    );
    
    if (botMessages.length === 0) {
      ui.alert('ℹ️ No bot messages found in the channel');
      return;
    }
    
    let messageList = 'Select a message to delete:\n\n';
    botMessages.forEach((msg, index) => {
      const date = new Date(parseFloat(msg.ts) * 1000);
      const dateStr = date.toLocaleString('en-US', { 
        month: 'short', 
        day: 'numeric', 
        hour: 'numeric', 
        minute: '2-digit',
        hour12: true 
      });
      
      let preview = 'Unknown';
      if (msg.text) {
        const match = msg.text.match(/Drive Checker: (.+?) -/);
        if (match) {
          preview = match[1];
        }
      } else if (msg.blocks && msg.blocks[0] && msg.blocks[0].text) {
        preview = msg.blocks[0].text.text || 'Message';
      }
      
      messageList += `${index + 1}. ${dateStr} - ${preview}\n`;
    });
    
    messageList += '\nEnter the number of the message to delete (or 0 to cancel):';
    
    const response = ui.prompt('Delete Slack Message', messageList, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const selection = parseInt(response.getResponseText());
    
    if (isNaN(selection) || selection < 1 || selection > botMessages.length) {
      if (selection !== 0) {
        ui.alert('❌ Invalid selection');
      }
      return;
    }
    
    const confirmResponse = ui.alert(
      'Confirm Deletion',
      `Delete message #${selection}?`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    const messageToDelete = botMessages[selection - 1];
    deleteSlackMessage(settings.slack_bot_token, settings.channel_id, messageToDelete.ts);
    
    ui.alert('✅ Message deleted!');
    
  } catch (e) {
    ui.alert(`❌ Error: ${e.message}`);
  }
}

function getBotUserId(token) {
  try {
    const url = 'https://slack.com/api/auth.test';
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    return result.user_id;
  } catch (e) {
    return null;
  }
}

// ========================================
// 🧪 TEST FUNCTIONS
// ========================================

function testSlackIntegration() {
  Logger.log('=== SLACK INTEGRATION TEST ===');
  
  Logger.log('\n1. Checking Settings sheet...');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      Logger.log('❌ FAIL: Settings sheet not found!');
      Logger.log(`   Looking for sheet named: "${SETTINGS_SHEET_NAME}"`);
      Logger.log('   Available sheets: ' + ss.getSheets().map(s => s.getName()).join(', '));
      return;
    }
    Logger.log('✅ PASS: Settings sheet found');
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
    return;
  }
  
  Logger.log('\n2. Loading Slack settings...');
  const settings = getSlackSettings();
  if (!settings) {
    Logger.log('❌ FAIL: getSlackSettings() returned null');
    return;
  }
  Logger.log('✅ PASS: Settings loaded');
  Logger.log('   Settings object: ' + JSON.stringify(settings, null, 2));
  
  Logger.log('\n3. Checking required fields...');
  if (!settings.slack_bot_token) {
    Logger.log('❌ FAIL: slack_bot_token is missing');
    return;
  }
  Logger.log('✅ Token found: ' + settings.slack_bot_token.substring(0, 20) + '...');
  
  if (!settings.channel_id) {
    Logger.log('❌ FAIL: channel_id is missing');
    return;
  }
  Logger.log('✅ Channel found: ' + settings.channel_id);
  Logger.log('✅ Post manual checks: ' + settings.post_manual_checks);
  
  Logger.log('\n4. Building test message...');
  try {
    const breakdown = { 'CamTrix': 5, 'Guardality': 3 };
    const message = buildSlackMessage('TEST CHECK', 8, breakdown);
    Logger.log('✅ PASS: Message built successfully');
    Logger.log('   Message preview: ' + JSON.stringify(message, null, 2));
  } catch (e) {
    Logger.log('❌ FAIL: ' + e.message);
    return;
  }
  
  Logger.log('\n5. Sending test message to Slack...');
  const confirm = SpreadsheetApp.getUi().alert(
    'Send Test Message?',
    'This will send a test message to your Slack channel. Continue?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (confirm !== SpreadsheetApp.getUi().Button.YES) {
    Logger.log('⏭️ Test message skipped by user');
    Logger.log('\n=== TEST COMPLETE (without sending) ===');
    return;
  }
  
  try {
    const breakdown = { 'CamTrix': 5, 'Guardality': 3 };
    const message = buildSlackMessage('TEST CHECK', 8, breakdown);
    const response = postToSlack(settings.slack_bot_token, settings.channel_id, message);
    
    if (response && response.ok) {
      Logger.log('✅ SUCCESS: Test message sent!');
      Logger.log('   Message timestamp: ' + response.ts);
      Logger.log('   Channel: ' + response.channel);
      SpreadsheetApp.getUi().alert('✅ Test message sent successfully! Check your Slack channel.');
    } else {
      Logger.log('❌ FAIL: Slack API returned error');
      Logger.log('   Response: ' + JSON.stringify(response, null, 2));
      SpreadsheetApp.getUi().alert('❌ Failed to send message. Check execution logs for details.');
    }
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
    Logger.log('   Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
  
  Logger.log('\n=== TEST COMPLETE ===');
}

function testBreakdownDebug() {
  // Manually trigger a check
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);
  const endOfYesterday = new Date(yesterday);
  endOfYesterday.setHours(23, 59, 59, 999);
  
  // Run the check
  runCheckSync(yesterday, endOfYesterday, 'Previous Day');
  
  // Check the debug logs
  Logger.log('Check the execution logs for breakdown data');
}

function debugGetViolationBreakdown() {
  // Manually set the check info to simulate an auto-check
  currentCheckInfo = {
    isAutoCheck: true,
    checkLabel: 'Previous Day',
    timestamp: new Date()
  };
  
  // Set the header row to the most recent check
  PropertiesService.getScriptProperties().setProperty('CURRENT_CHECK_HEADER_ROW', '3');
  
  log(`📍 About to get violation breakdown. CURRENT_CHECK_HEADER_ROW = ${PropertiesService.getScriptProperties().getProperty('CURRENT_CHECK_HEADER_ROW')}`);
  const breakdown = getViolationBreakdown();
  log(`📍 Got breakdown: ${JSON.stringify(breakdown)}`);
  
  // Log results
  Logger.log('=== BREAKDOWN RESULTS ===');
  Logger.log(`Total products found: ${Object.keys(breakdown).length}`);
  Logger.log(`Breakdown: ${JSON.stringify(breakdown, null, 2)}`);
  
  // TEST THE SLACK MESSAGE BUILDER
  Logger.log('\n=== TESTING SLACK MESSAGE BUILDER ===');
  const message = buildSlackMessage('Previous Day', 13, breakdown, false);
  Logger.log('Message built successfully!');
  
  // Check if breakdown made it into the message
  const messageText = JSON.stringify(message);
  if (messageText.includes('WiggyDog')) {
    Logger.log('✅ Breakdown IS in the Slack message');
  } else {
    Logger.log('❌ Breakdown NOT in the Slack message');
  }
}

function markFixedViolations() {
  try {
    const sheet = getOrCreateSheet(DRIVE_CONFIG.AUTO_CHECK_HISTORY_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    const data = sheet.getRange(1, 1, lastRow, 10).getValues();
    let fixedCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const type = String(row[2]).trim(); // Column C: Type
      const status = String(row[9]).trim(); // Column J: Status
      
      // Only process violation rows with NEW status
      if ((type !== 'FILE' && type !== 'FOLDER') || status !== '🆕 NEW') continue;
      
      // Convert to string and trim to handle numbers and whitespace
      const name = String(row[4]).trim(); // Column E: Name
      const link = String(row[8]).trim(); // Column I: Link
      
      // Skip year folder violations - they shouldn't have been logged
      if (/^\d{4}$/.test(name)) {
        log(`  ⏭️  Skipping year folder in markFixed: ${name}`);
        continue;
      }
      
      // Extract file/folder ID from link
      let itemId = null;
      if (link.includes('/folders/')) {
        const match = link.match(/\/folders\/([^/?]+)/);
        itemId = match ? match[1] : null;
      } else if (link.includes('/file/d/')) {
        const match = link.match(/\/file\/d\/([^/?]+)/);
        itemId = match ? match[1] : null;
      }
      
      if (!itemId) continue;
      
      try {
        // Try to get the current name of the file/folder
        let currentName;
        if (type === 'FOLDER') {
          const folder = DriveApp.getFolderById(itemId);
          currentName = String(folder.getName()).trim();
        } else {
          const file = DriveApp.getFileById(itemId);
          currentName = String(file.getName()).trim();
        }
        
        // CRITICAL: Compare as trimmed strings to avoid type coercion bugs
        if (currentName !== name) {
          sheet.getRange(i + 1, 10).setValue('✅ FIXED');
          fixedCount++;
          log(`✅ Marked as FIXED: "${name}" → "${currentName}"`);
        }
        // If name is the same, leave as NEW - violation still exists
        
      } catch (e) {
        // File/folder might be deleted or moved
        sheet.getRange(i + 1, 10).setValue('🗑️ DELETED');
        log(`🗑️ Item no longer accessible: ${name}`);
      }
    }
    
    log(`✅ Marked ${fixedCount} violations as FIXED`);
    return fixedCount;
    
  } catch (e) {
    log(`❌ Error marking fixed violations: ${e.message}`);
    return 0;
  }
}

function validateYearPlacement(item, itemType, yearFromPath, parentPath) {
  try {
    const createdDate = item.getDateCreated();
    const createdYear = createdDate.getFullYear();
    const pathYear = parseInt(yearFromPath);
    
    if (isNaN(pathYear)) return null; // Can't validate if year not found in path
    
    if (createdYear !== pathYear) {
      return `Wrong year folder: ${itemType} was created in ${createdYear} but is in "${pathYear}" folder`;
    }
    
    return null; // No violation
  } catch (e) {
    log(`⚠️ Could not validate year placement: ${e.message}`);
    return null;
  }
}

function runDriveCheckerDiagnostics() {
  const startTime = Date.now();
  const results = {
    passed: 0,
    failed: 0,
    warnings: 0,
    tests: []
  };
  
  const log = (category, testName, status, details = '') => {
    const entry = { category, testName, status, details };
    results.tests.push(entry);
    
    if (status === 'PASS') results.passed++;
    else if (status === 'FAIL') results.failed++;
    else if (status === 'WARN') results.warnings++;
    
    const icon = status === 'PASS' ? '✅' : status === 'FAIL' ? '❌' : '⚠️';
    Logger.log(`${icon} [${category}] ${testName}: ${status}${details ? ' - ' + details : ''}`);
  };
  
  Logger.log('');
  Logger.log('╔══════════════════════════════════════════════════════════════════╗');
  Logger.log('║      DRIVE NAMING CHECKER - CONFIG DIAGNOSTICS REPORT           ║');
  Logger.log('║      ' + new Date().toLocaleString() + '                              ║');
  Logger.log('╚══════════════════════════════════════════════════════════════════╝');
  Logger.log('');

  // ============================================================================
  // TEST 1: CONFIG SHEET CONNECTION
  // ============================================================================
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📡 TEST 1: CONFIG SHEET CONNECTION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const configSheet = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    log('Connection', 'Open Config Sheet', 'PASS', 'ID: ' + CONFIG_SHEET_ID.substring(0, 20) + '...');
    
    const requiredTabs = ['Products', 'Valid Values'];
    for (const tabName of requiredTabs) {
      const tab = configSheet.getSheetByName(tabName);
      if (tab) {
        log('Connection', `Tab: ${tabName}`, 'PASS', `${tab.getLastRow()} rows`);
      } else {
        log('Connection', `Tab: ${tabName}`, 'FAIL', 'Tab not found!');
      }
    }
  } catch (e) {
    log('Connection', 'Open Config Sheet', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 2: PRODUCTS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📦 TEST 2: PRODUCTS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const products = VALID_VALUES.PRODUCTS;
    if (Array.isArray(products) && products.length > 0) {
      log('Products', 'Load Products', 'PASS', `${products.length} active product(s)`);
      
      // Check for expected products
      const expectedProducts = ['CamTrix', 'TellyStick', 'Guardality', 'AliveBlue', 'KatuChef'];
      const missing = expectedProducts.filter(p => !products.includes(p));
      
      if (missing.length === 0) {
        log('Products', 'Expected Products Present', 'PASS', 'All key products found');
      } else {
        log('Products', 'Expected Products Present', 'WARN', `Missing: ${missing.join(', ')}`);
      }
      
      // Check for duplicates
      const uniqueProducts = [...new Set(products)];
      if (uniqueProducts.length === products.length) {
        log('Products', 'Duplicate Check', 'PASS', 'No duplicates');
      } else {
        log('Products', 'Duplicate Check', 'WARN', `${products.length - uniqueProducts.length} duplicate(s)`);
      }
      
      Logger.log('   Products: ' + products.slice(0, 8).join(', ') + (products.length > 8 ? '...' : ''));
    } else {
      log('Products', 'Load Products', 'FAIL', 'No products found');
    }
  } catch (e) {
    log('Products', 'Load Products', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 3: PLATFORMS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📱 TEST 3: PLATFORMS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const platforms = VALID_VALUES.PLATFORMS;
    if (Array.isArray(platforms) && platforms.length > 0) {
      log('Platforms', 'Load Platforms', 'PASS', `${platforms.length} platform(s)`);
      Logger.log('   Platforms: ' + platforms.join(', '));
      
      // Check expected platforms
      const expectedPlatforms = ['FB', 'IG', 'TT', 'YT', 'GG'];
      const missing = expectedPlatforms.filter(p => !platforms.includes(p));
      
      if (missing.length === 0) {
        log('Platforms', 'Expected Platforms Present', 'PASS', 'All 5 platforms found');
      } else {
        log('Platforms', 'Expected Platforms Present', 'WARN', `Missing: ${missing.join(', ')}`);
      }
    } else {
      log('Platforms', 'Load Platforms', 'FAIL', 'No platforms found');
    }
  } catch (e) {
    log('Platforms', 'Load Platforms', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 4: LOCALES
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🌍 TEST 4: LOCALES');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const locales = VALID_VALUES.LOCALES;
    if (Array.isArray(locales) && locales.length > 0) {
      log('Locales', 'Load Locales', 'PASS', `${locales.length} locale(s)`);
      Logger.log('   Locales: ' + locales.join(', '));
      
      // Check expected locales
      const expectedLocales = ['US', 'UK', 'DE', 'CA', 'AU', 'LATAM', 'ES'];
      const missing = expectedLocales.filter(l => !locales.includes(l));
      
      if (missing.length === 0) {
        log('Locales', 'Core Locales Present', 'PASS', 'All 7 core locales found');
      } else {
        log('Locales', 'Core Locales Present', 'WARN', `Missing: ${missing.join(', ')}`);
      }
    } else {
      log('Locales', 'Load Locales', 'FAIL', 'No locales found');
    }
  } catch (e) {
    log('Locales', 'Load Locales', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 5: MONITOR FOLDERS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📁 TEST 5: MONITOR FOLDERS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const folders = DRIVE_CONFIG.FOLDERS_TO_MONITOR;
    if (Array.isArray(folders) && folders.length > 0) {
      log('Folders', 'Load Monitor Folders', 'PASS', `${folders.length} folder(s)`);
      
      // Validate folder IDs format (should be ~33 characters)
      const validFormat = folders.filter(f => f && f.length >= 25 && f.length <= 50);
      if (validFormat.length === folders.length) {
        log('Folders', 'Folder ID Format', 'PASS', 'All IDs valid format');
      } else {
        log('Folders', 'Folder ID Format', 'WARN', `${folders.length - validFormat.length} invalid format`);
      }
      
      // Test first folder access
      try {
        const testFolder = DriveApp.getFolderById(folders[0]);
        log('Folders', 'Drive Access Test', 'PASS', `Can access: ${testFolder.getName()}`);
      } catch (e) {
        log('Folders', 'Drive Access Test', 'FAIL', `Cannot access first folder: ${e.message}`);
      }
      
      Logger.log('   Folder IDs: ' + folders.slice(0, 3).map(f => f.substring(0, 15) + '...').join(', ') + (folders.length > 3 ? '...' : ''));
    } else {
      log('Folders', 'Load Monitor Folders', 'FAIL', 'No folders configured');
    }
  } catch (e) {
    log('Folders', 'Load Monitor Folders', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 6: PRODUCT ALIASES (if function exists)
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔄 TEST 6: PRODUCT ALIASES');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    if (typeof getProductAliasesFromConfig_ === 'function') {
      const aliases = getProductAliasesFromConfig_();
      const aliasCount = Object.keys(aliases).length;
      
      if (aliasCount > 0) {
        log('Aliases', 'Load Product Aliases', 'PASS', `${aliasCount} alias(es)`);
        
        for (const [alias, canonical] of Object.entries(aliases)) {
          Logger.log(`   ${alias} → ${canonical}`);
        }
        
        // Test specific expected alias
        if (aliases['ArmouredCard'] === 'Guardality') {
          log('Aliases', 'ArmouredCard→Guardality', 'PASS', 'Alias works correctly');
        } else {
          log('Aliases', 'ArmouredCard→Guardality', 'WARN', 'Alias not found or incorrect');
        }
      } else {
        log('Aliases', 'Load Product Aliases', 'WARN', 'No aliases configured');
      }
    } else {
      log('Aliases', 'Alias Function', 'WARN', 'getProductAliasesFromConfig_ not found - add FIX 1');
    }
  } catch (e) {
    log('Aliases', 'Load Product Aliases', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 7: FUNCTION VALIDATION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔧 TEST 7: FUNCTION VALIDATION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  // Check looksLikeBatchFolder uses VALID_VALUES.PLATFORMS
  try {
    const funcStr = looksLikeBatchFolder.toString();
    if (funcStr.includes('VALID_VALUES.PLATFORMS')) {
      log('Functions', 'looksLikeBatchFolder', 'PASS', 'Uses VALID_VALUES.PLATFORMS');
    } else if (funcStr.includes("'FB'") || funcStr.includes("['FB'")) {
      log('Functions', 'looksLikeBatchFolder', 'FAIL', 'Still has hardcoded platforms - apply FIX 3');
    } else {
      log('Functions', 'looksLikeBatchFolder', 'WARN', 'Could not verify implementation');
    }
  } catch (e) {
    log('Functions', 'looksLikeBatchFolder', 'WARN', 'Could not inspect function');
  }
  
  // Check extractProductName handles aliases
  try {
    const funcStr = extractProductName.toString();
    if (funcStr.includes('getProductAliasesFromConfig_') || funcStr.includes('aliases')) {
      log('Functions', 'extractProductName', 'PASS', 'Handles product aliases');
    } else {
      log('Functions', 'extractProductName', 'WARN', 'Does not handle aliases - apply FIX 4');
    }
  } catch (e) {
    log('Functions', 'extractProductName', 'WARN', 'Could not inspect function');
  }

  // ============================================================================
  // TEST 8: CACHE PERFORMANCE
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('⚡ TEST 8: CACHE PERFORMANCE');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Clear cache
    CacheService.getScriptCache().removeAll(['dnc_products', 'dnc_platforms', 'dnc_locales', 'dnc_folders', 'dnc_aliases']);
    
    // Uncached load
    const uncachedStart = Date.now();
    getProductsFromConfig_();
    getPlatformsFromConfig_();
    getLocalesFromConfig_();
    getFoldersToMonitorFromConfig_();
    const uncachedTime = Date.now() - uncachedStart;
    
    // Cached load
    const cachedStart = Date.now();
    getProductsFromConfig_();
    getPlatformsFromConfig_();
    getLocalesFromConfig_();
    getFoldersToMonitorFromConfig_();
    const cachedTime = Date.now() - cachedStart;
    
    log('Cache', 'Uncached Load', uncachedTime < 5000 ? 'PASS' : 'WARN', `${uncachedTime}ms`);
    log('Cache', 'Cached Load', cachedTime < 100 ? 'PASS' : 'WARN', `${cachedTime}ms`);
    log('Cache', 'Speedup', cachedTime < uncachedTime ? 'PASS' : 'WARN', 
        `${Math.round(uncachedTime / Math.max(cachedTime, 1))}x faster`);
  } catch (e) {
    log('Cache', 'Cache Test', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 9: EDGE CASES
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔬 TEST 9: EDGE CASES');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Test product extraction
    const testName = 'CamTrix_FB_US_123_1_VID.mp4';
    const extracted = extractProductName(testName, '');
    log('EdgeCase', 'Product Extraction', extracted === 'CamTrix' ? 'PASS' : 'FAIL', 
        `"${testName}" → "${extracted}"`);
    
    // Test batch folder detection
    const isBatch = looksLikeBatchFolder('CamTrix_FB_US_123_(Test)');
    log('EdgeCase', 'Batch Folder Detection', isBatch === true ? 'PASS' : 'FAIL',
        `Standard batch folder → ${isBatch}`);
    
    // Test locale validation
    const isValidUS = isValidLocaleFolder('US');
    const isValidLatam = isValidLocaleFolder('.LATAM');
    log('EdgeCase', 'Locale Validation', isValidUS && isValidLatam ? 'PASS' : 'FAIL',
        `US=${isValidUS}, .LATAM=${isValidLatam}`);
        
  } catch (e) {
    log('EdgeCase', 'Edge Case Tests', 'FAIL', e.message);
  }

  // ============================================================================
  // SUMMARY
  // ============================================================================
  const totalTime = Date.now() - startTime;
  
  Logger.log('');
  Logger.log('╔══════════════════════════════════════════════════════════════════╗');
  Logger.log('║                      DIAGNOSTIC SUMMARY                          ║');
  Logger.log('╠══════════════════════════════════════════════════════════════════╣');
  Logger.log(`║  ✅ PASSED:   ${String(results.passed).padEnd(4)} tests                                    ║`);
  Logger.log(`║  ⚠️  WARNINGS: ${String(results.warnings).padEnd(4)} tests                                    ║`);
  Logger.log(`║  ❌ FAILED:   ${String(results.failed).padEnd(4)} tests                                    ║`);
  Logger.log(`║  ⏱️  TIME:     ${String(totalTime + 'ms').padEnd(8)}                                   ║`);
  Logger.log('╠══════════════════════════════════════════════════════════════════╣');
  
  if (results.failed === 0 && results.warnings === 0) {
    Logger.log('║  🎉 ALL TESTS PASSED! Config is working perfectly.              ║');
  } else if (results.failed === 0) {
    Logger.log('║  ✅ Core tests passed. Review warnings above.                   ║');
  } else {
    Logger.log('║  ⚠️  Some tests failed. Review the log above for details.       ║');
  }
  
  Logger.log('╚══════════════════════════════════════════════════════════════════╝');
  
  // Show UI summary
  SpreadsheetApp.getUi().alert(
    'Drive Checker Diagnostics',
    `✅ Passed: ${results.passed}\n⚠️ Warnings: ${results.warnings}\n❌ Failed: ${results.failed}\n\nTime: ${totalTime}ms\n\nSee Execution Log for details.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  return results;
}

/**
 * Quick health check for Drive Naming Checker
 */
function quickDriveCheckerHealthCheck() {
  const checks = [];
  
  try {
    checks.push({ name: 'Config Sheet', ok: !!SpreadsheetApp.openById(CONFIG_SHEET_ID) });
  } catch (e) {
    checks.push({ name: 'Config Sheet', ok: false });
  }
  
  checks.push({ name: 'Products', ok: Array.isArray(VALID_VALUES.PRODUCTS) && VALID_VALUES.PRODUCTS.length > 0 });
  checks.push({ name: 'Platforms', ok: Array.isArray(VALID_VALUES.PLATFORMS) && VALID_VALUES.PLATFORMS.length > 0 });
  checks.push({ name: 'Locales', ok: Array.isArray(VALID_VALUES.LOCALES) && VALID_VALUES.LOCALES.length > 0 });
  checks.push({ name: 'Monitor Folders', ok: Array.isArray(DRIVE_CONFIG.FOLDERS_TO_MONITOR) && DRIVE_CONFIG.FOLDERS_TO_MONITOR.length > 0 });
  
  const passed = checks.filter(c => c.ok).length;
  const total = checks.length;
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${passed}/${total} checks passed`,
    passed === total ? '✅ Config OK' : '⚠️ Issues Found',
    5
  );
  
  return { passed, total, checks };
}

function testErrorLogging() {
  logErrorToHub_('Test', 'This is a test error');
  Logger.log('Done - check Config sheet Error Log tab');
}