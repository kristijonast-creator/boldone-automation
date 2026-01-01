// ================================================================================
// CENTRAL CONFIG CONNECTION
// ================================================================================
const CONFIG_SHEET_ID = '1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk';
const CONFIG_CACHE_DURATION = 300;

function logClickUpTasksCreated(taskCount, productsSummary) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let sheet = ss.getSheetByName('ClickUp Log');
    
    if (!sheet) {
      sheet = ss.insertSheet('ClickUp Log');
      sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Tasks Created', 'Products', 'User', 'Duration']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#d9ead3');
      sheet.setFrozenRows(1);
    }
    
    let user = 'Unknown';
    try { user = Session.getActiveUser().getEmail(); } catch (_) {}
    
    sheet.insertRowAfter(1);
    const newRow = sheet.getRange(2, 1, 1, 5);
    newRow.setValues([[
      new Date().toISOString(),
      taskCount,
      productsSummary || '',
      user,
      ''
    ]]);
    // Clear formatting copied from header
    newRow.setBackground(null).setFontWeight('normal');
  } catch (e) {
    Logger.log('Could not log ClickUp tasks: ' + e);
  }
}

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
    const newRow = sheet.getRange(2, 1, 1, 5);
    newRow.setValues([[
      new Date().toISOString(),
      'Sprint → ClickUp',
      errorType,
      String(message).substring(0, 500),
      user
    ]]);
    // Clear formatting copied from header
    newRow.setBackground(null).setFontWeight('normal');
  } catch (e) {
    Logger.log('Could not log error to Hub: ' + e);
  }
}

function getClickUpApiKeysFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('abc_api_keys');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const keys = {};
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const apiKey = data[i][2];
    const active = data[i][7];
    if (email && apiKey && active === true) {
      keys[email] = apiKey;
    }
  }
  
  cache.put('abc_api_keys', JSON.stringify(keys), CONFIG_CACHE_DURATION);
  return keys;
}

function getAssigneeEmailMapFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('abc_assignees');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const email = data[i][1];
    const active = data[i][7];
    if (name && email && active === true) {
      map[name] = email;
    }
  }
  
  cache.put('abc_assignees', JSON.stringify(map), CONFIG_CACHE_DURATION);
  return map;
}

function getClickUpConfigFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('abc_clickup');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('ClickUp Config');
  const data = sheet.getDataRange().getValues();
  
  const config = {};
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  
  cache.put('abc_clickup', JSON.stringify(config), CONFIG_CACHE_DURATION);
  return config;
}

function clearABCConfigCache() {
  CacheService.getScriptCache().removeAll(['abc_api_keys', 'abc_assignees', 'abc_clickup']);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Config cache cleared!', 'Done', 3);
}
// ================================================================================

/**
 * OAuth Scopes:
 * @scope https://www.googleapis.com/auth/userinfo.email
 * @scope https://www.googleapis.com/auth/spreadsheets
 */

/*******************************************************
 * AUTO-BATCH CREATOR
 * Creates ClickUp tasks from selected spreadsheet rows
 *******************************************************/

// ✅ CONFIGURATION - Lazy loaded to prevent menu from breaking
let _clickupApiKeys = null;
let _clickupConfig = null;
let _assigneeEmailMap = null;

function getClickUpApiKeysLazy_() {
  if (!_clickupApiKeys) {
    _clickupApiKeys = getClickUpApiKeysFromConfig_();
  }
  return _clickupApiKeys;
}

function getClickUpConfigLazy_() {
  if (!_clickupConfig) {
    _clickupConfig = getClickUpConfigFromConfig_();
  }
  return _clickupConfig;
}

function getAssigneeEmailMapLazy_() {
  if (!_assigneeEmailMap) {
    _assigneeEmailMap = getAssigneeEmailMapFromConfig_();
  }
  return _assigneeEmailMap;
}

/**
 * Get the current user's ClickUp API key
 */
function getClickUpApiKey() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Check if email is empty (permission not granted)
    if (!userEmail || userEmail.trim() === '') {
      throw new Error('⚠️ Email permission required.\n\nPlease:\n1. Refresh the page\n2. Grant email access when prompted\n3. Try again');
    }
    
    const apiKeys = getClickUpApiKeysLazy_();
    const apiKey = apiKeys[userEmail];
    
    if (!apiKey || apiKey.includes('PASTE_')) {
      throw new Error('❌ No ClickUp API key configured for ' + userEmail + '.\n\nPlease contact Kristijonas to add your API key to the script.');
    }
    
    Logger.log('✅ Using API key for: ' + userEmail);
    return apiKey;
    
  } catch (e) {
    // If it's our custom error, throw it as-is
    if (e.message.includes('⚠️') || e.message.includes('❌')) {
      throw e;
    }
    
    // Otherwise, it's a permission error
    throw new Error('⚠️ Email permission required.\n\nPlease:\n1. Refresh the page\n2. Grant email access when prompted\n3. Try again\n\nTechnical error: ' + e.message);
  }
}

// Priority mapping
const PRIORITY_MAP = {
  'Urgent': 1,
  'High': 2,
  'Normal': 3,
  'Low': 4
};

/**
 * Add custom menu on open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Auto-Batch')
    .addItem('Create ClickUp Tasks', 'showAutoBatchDialog')
    .addToUi();
}

/**
 * Main entry point - validates selection and shows dialog
 */
function showAutoBatchDialog() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    if (!range) {
      SpreadsheetApp.getUi().alert('❌ Please select rows to create tasks from.');
      return;
    }
    
    // Validate and extract data
    const result = validateAndExtractRows(sheet, range);
    
    if (result.error) {
      SpreadsheetApp.getUi().alert('❌ ' + result.error);
      return;
    }
    
    if (result.totalTasks === 0) {
      SpreadsheetApp.getUi().alert('❌ No valid tasks found in selection.\n\nMake sure column E (Product) is filled.');
      return;
    }
    
    // ✅ NEW: Fetch sprint lists
    try {
      const sprintLists = getSprintLists();
      result.sprintLists = sprintLists;
    } catch (e) {
      SpreadsheetApp.getUi().alert('❌ Error fetching sprint lists: ' + e.message + '\n\nPlease check your Sprint folder setup.');
      return;
    }
    
    // Show dialog
    const html = buildAutoBatchDialog(result);
    SpreadsheetApp.getUi().showModalDialog(html, '🚀 Auto-Batch Creator');
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * Validates selection and extracts task data, grouped by product
 * Handles both contiguous and non-contiguous selections
 * NOW WITH COLUMN VALIDATION!
 */
function validateAndExtractRows(sheet, range) {
  Logger.log('📊 Processing selection...');
  
  const tasksByProduct = {}; // Group tasks by product
  
  // Get all selected ranges (handles non-contiguous selections)
  const ranges = sheet.getActiveRangeList() ? sheet.getActiveRangeList().getRanges() : [range];
  
  Logger.log('📊 Found ' + ranges.length + ' range(s) in selection');
  
  // ✅ NEW: Check which columns are included in the selection
  const selectedColumns = new Set();
  ranges.forEach(function(r) {
    const startCol = r.getColumn();
    const numCols = r.getNumColumns();
    for (let i = 0; i < numCols; i++) {
      selectedColumns.add(startCol + i);
    }
  });
  
  Logger.log('📋 Selected columns: ' + Array.from(selectedColumns).join(', '));
  
  // ✅ NEW: Define required columns (D=4, E=5, F=6, G=7, H=8, I=9, J=10)
  const COLUMNS = {
    PRIORITY: 4,      // D
    PRODUCT: 5,       // E - CRITICAL
    TASK_NAME: 6,     // F - CRITICAL
    QUANTITY: 7,      // G
    DESCRIPTION: 8,   // H
    TIME_ESTIMATE: 9, // I
    ASSIGNEE: 10      // J
  };
  
  // ✅ NEW: Check for missing columns
  const criticalMissing = [];
  const warningMissing = [];
  
  if (!selectedColumns.has(COLUMNS.PRODUCT)) {
    criticalMissing.push('E (Product)');
  }
  if (!selectedColumns.has(COLUMNS.TASK_NAME)) {
    criticalMissing.push('F (Task Name)');
  }
  if (!selectedColumns.has(COLUMNS.PRIORITY)) {
    warningMissing.push('D (Priority) - Tasks will default to "Normal" priority');
  }
  if (!selectedColumns.has(COLUMNS.QUANTITY)) {
    warningMissing.push('G (Quantity) - Pack tasks cannot be created');
  }
  if (!selectedColumns.has(COLUMNS.DESCRIPTION)) {
    warningMissing.push('H (Description) - Tasks will have no description');
  }
  if (!selectedColumns.has(COLUMNS.TIME_ESTIMATE)) {
    warningMissing.push('I (Time Estimate) - Tasks will have no time estimate');
  }
  if (!selectedColumns.has(COLUMNS.ASSIGNEE)) {
    warningMissing.push('J (Assignee) - Tasks will be unassigned');
  }
  
  // ✅ NEW: If critical columns missing, return error immediately
  if (criticalMissing.length > 0) {
    return {
      error: '🔴 CRITICAL: Cannot create tasks without required columns.\n\nMissing columns: ' + criticalMissing.join(', ') + '\n\nPlease select entire rows by clicking row numbers.',
      criticalMissing: criticalMissing,
      warningMissing: [],
      tasksByProduct: {},
      productCount: 0,
      totalTasks: 0
    };
  }
  
  // ✅ NEW: Track processed rows to avoid duplicates
  const processedRows = new Set();
  
  // Process each range
  ranges.forEach(function(r) {
    const startRow = r.getRow();
    const numRows = r.getNumRows();
    
    Logger.log('   Processing rows ' + startRow + ' to ' + (startRow + numRows - 1));
    
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      
      // ✅ NEW: Skip if we've already processed this row
      if (processedRows.has(row)) {
        Logger.log('   ⏭️ Skipping row ' + row + ' (already processed)');
        continue;
      }
      
      // Mark this row as processed
      processedRows.add(row);
      
      // Get values from columns D-J (only if they're selected)
      const priority = selectedColumns.has(COLUMNS.PRIORITY) ? sheet.getRange(row, COLUMNS.PRIORITY).getValue() : '';
      const product = selectedColumns.has(COLUMNS.PRODUCT) ? sheet.getRange(row, COLUMNS.PRODUCT).getValue() : '';
      const taskName = selectedColumns.has(COLUMNS.TASK_NAME) ? sheet.getRange(row, COLUMNS.TASK_NAME).getValue() : '';
      const quantity = selectedColumns.has(COLUMNS.QUANTITY) ? sheet.getRange(row, COLUMNS.QUANTITY).getValue() : '';
      const description = selectedColumns.has(COLUMNS.DESCRIPTION) ? sheet.getRange(row, COLUMNS.DESCRIPTION).getValue() : '';
      const timeEstimate = selectedColumns.has(COLUMNS.TIME_ESTIMATE) ? sheet.getRange(row, COLUMNS.TIME_ESTIMATE).getValue() : '';
      const assignee = selectedColumns.has(COLUMNS.ASSIGNEE) ? sheet.getRange(row, COLUMNS.ASSIGNEE).getValue() : '';
      
      // Skip empty rows (no product)
      if (!product || String(product).trim() === '') {
        Logger.log('   Skipping row ' + row + ' (no product)');
        continue;
      }
      
      const productKey = String(product).trim();
      
      // Initialize product array if it doesn't exist
      if (!tasksByProduct[productKey]) {
        tasksByProduct[productKey] = [];
        Logger.log('   📦 New product found: ' + productKey);
      }
      
      // Add task to this product's array
      tasksByProduct[productKey].push({
        rowNumber: row,
        priority: String(priority || 'Normal').trim(),
        product: productKey,
        taskName: String(taskName || 'Untitled').trim(),
        quantity: quantity && !isNaN(quantity) ? parseInt(quantity) : 0,
        description: String(description || '').trim(),
        timeEstimate: timeEstimate && !isNaN(timeEstimate) ? parseFloat(timeEstimate) : 0,
        assignee: String(assignee || '').trim()
      });
      
      Logger.log('   ✅ Added task: ' + taskName + ' (row ' + row + ')');
    }
  });
  
  // Check if we found any tasks
  const productCount = Object.keys(tasksByProduct).length;
  
  if (productCount === 0) {
    return {
      error: 'No valid rows found. Make sure column E (Product) is filled.',
      criticalMissing: [],
      warningMissing: [],
      tasksByProduct: {},
      productCount: 0,
      totalTasks: 0
    };
  }
  
  // Calculate total tasks
  let totalTasks = 0;
  for (const product in tasksByProduct) {
    totalTasks += tasksByProduct[product].length;
  }
  
  Logger.log('✅ Found ' + totalTasks + ' tasks across ' + productCount + ' product(s)');
  
  // Log summary
  for (const product in tasksByProduct) {
    Logger.log('   📦 ' + product + ': ' + tasksByProduct[product].length + ' task(s)');
  }
  
  return {
    tasksByProduct: tasksByProduct,
    productCount: productCount,
    totalTasks: totalTasks,
    criticalMissing: criticalMissing,
    warningMissing: warningMissing,
    error: null
  };
}

/**
 * Builds the multi-product HTML dialog
 */
function buildAutoBatchDialog(data) {
  const tasksByProduct = data.tasksByProduct;
  const productCount = data.productCount;
  const totalTasks = data.totalTasks;
  const sprintLists = data.sprintLists || []; // ✅ NEW
  
  // Build product sections
  let productSections = '';
  let productIndex = 0;
  
  for (const product in tasksByProduct) {
    const tasks = tasksByProduct[product];
    
    // Build task preview for this product
    let taskPreview = '';
    tasks.forEach(task => {
      const packInfo = task.quantity > 1 ? ` (pack: ${task.quantity} items)` : '';
      taskPreview += `<div style="font-size:11px;padding:2px 0;color:#6b7280">• ${task.taskName}${packInfo}</div>`;
    });
    
    productSections += `
      <div class="product-section">
        <div class="product-header">
          <div class="product-name">📦 ${product}</div>
          <div class="product-count">${tasks.length} task${tasks.length === 1 ? '' : 's'}</div>
        </div>
        <div class="list-select-wrapper">
          <label class="list-label">Select ClickUp List:</label>
          <select class="list-select" id="listSelect${productIndex}" data-product="${product}">
            <option value="">Loading lists...</option>
          </select>
        </div>
        <div class="task-preview">${taskPreview}</div>
      </div>
    `;
    
    productIndex++;
  }
  
  // ✅ NEW: Build sprint dropdown options (NO AUTO-SELECT)
let sprintOptions = '';
sprintLists.forEach(function(sprint) {
  sprintOptions += `<option value="${sprint.id}">${sprint.name}</option>`; // Removed 'selected'
});
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; max-height: 600px; overflow-y: auto; }
    h2 { margin: 0 0 10px; color: #333; font-size: 18px; }
    .info { background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 12px; margin: 10px 0; font-size: 13px; }
    
    /* ✅ UPDATED: Sprint selection box */
.sprint-section {
  background: #f0fdf4;
  border: 2px solid #86efac;
  border-radius: 8px;
  padding: 16px;
  margin: 16px 0;
}
    
.sprint-header {
  font-weight: 700;
  color: #166534;
  font-size: 15px;
  margin-bottom: 10px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.sprint-suggest-btn {
  background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%);
  color: white;
  border: none;
  padding: 8px 16px;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  box-shadow: 0 2px 4px rgba(34, 197, 94, 0.3);
  transition: all 0.2s ease;
  display: flex;
  align-items: center;
  gap: 4px;
}

.sprint-suggest-btn:hover {
  background: linear-gradient(135deg, #16a34a 0%, #15803d 100%);
  box-shadow: 0 4px 8px rgba(34, 197, 94, 0.4);
  transform: translateY(-1px);
}

.sprint-suggest-btn:active {
  transform: translateY(0);
  box-shadow: 0 2px 4px rgba(34, 197, 94, 0.3);
}
    
.sprint-select {
  width: 100%;
  padding: 10px;
  border: 2px solid #86efac;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  background: white;
  color: #166534;
  transition: border-color 0.2s ease;
}

.sprint-select:focus {
  outline: none;
  border-color: #22c55e;
  box-shadow: 0 0 0 3px rgba(34, 197, 94, 0.1);
}

.sprint-select option[value=""] {
  color: #9ca3af;
}

.sprint-helper {
  margin-top: 8px;
  padding: 8px 12px;
  background: #fef3c7;
  border: 1px solid #fbbf24;
  border-radius: 4px;
  font-size: 12px;
  color: #92400e;
  font-weight: 600;
}

/* ✅ NEW: Warning when no sprint selected */
.create-warning {
  background: #fef3c7;
  border: 2px solid #fbbf24;
  border-radius: 6px;
  padding: 12px;
  margin-bottom: 10px;
  font-size: 13px;
  color: #92400e;
  font-weight: 600;
  display: none;
  align-items: center;
  gap: 8px;
}

.create-warning.show {
  display: flex;
}
    
    .warning-box { 
      background: #fef3c7; 
      border: 2px solid #f59e0b; 
      border-radius: 8px; 
      padding: 16px; 
      margin: 10px 0; 
      font-size: 13px; 
    }
    
    .warning-title { 
      font-weight: 700; 
      color: #92400e; 
      font-size: 15px; 
      margin-bottom: 8px; 
      display: flex;
      align-items: center;
      gap: 8px;
    }
    
    .warning-list { 
      margin: 8px 0 8px 20px; 
      color: #78350f; 
    }
    
    .warning-list li { 
      margin: 4px 0; 
    }
    
    .warning-footer {
      margin-top: 12px;
      padding-top: 12px;
      border-top: 1px solid #f59e0b;
      color: #92400e;
      font-weight: 600;
    }
    
    .product-section { 
      background: #f9fafb; 
      border: 1px solid #e5e7eb; 
      border-radius: 8px; 
      padding: 12px; 
      margin: 12px 0; 
    }
    
    .product-header { 
      display: flex; 
      justify-content: space-between; 
      align-items: center; 
      margin-bottom: 10px; 
    }
    
    .product-name { 
      font-weight: 600; 
      font-size: 14px; 
      color: #1f2937; 
    }
    
    .product-count { 
      font-size: 12px; 
      color: #6b7280; 
      background: #e5e7eb; 
      padding: 2px 8px; 
      border-radius: 12px; 
    }
    
    .list-select-wrapper { 
      margin: 10px 0; 
    }
    
    .list-label { 
      display: block; 
      font-size: 12px; 
      font-weight: 600; 
      margin-bottom: 5px; 
      color: #374151; 
    }
    
    .list-select { 
      width: 100%; 
      padding: 8px; 
      border: 1px solid #d1d5db; 
      border-radius: 4px; 
      font-size: 13px; 
      background: white;
    }
    
    .task-preview { 
      margin-top: 10px; 
      padding-top: 10px; 
      border-top: 1px solid #e5e7eb; 
      max-height: 100px; 
      overflow-y: auto; 
    }
    
    .buttons { display: flex; gap: 10px; margin-top: 20px; position: sticky; bottom: 0; background: white; padding: 10px 0; }
    button { padding: 10px 20px; border: none; border-radius: 6px; font-size: 14px; cursor: pointer; font-weight: 600; }
    .btn-primary { background: #7c3aed; color: white; flex: 1; }
    .btn-primary:hover { background: #6d28d9; }
    .btn-secondary { background: #e5e7eb; color: #374151; }
    .btn-secondary:hover { background: #d1d5db; }
    .btn-primary:disabled { background: #d1d5db; cursor: not-allowed; }
    .loading { display: none; text-align: center; margin: 10px 0; color: #7c3aed; font-size: 13px; }

    .progress-container {
      background: #f9fafb;
      border: 1px solid #e5e7eb;
      border-radius: 8px;
      padding: 16px;
      margin: 16px 0;
    }
    
    .progress-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .progress-title {
      font-weight: 600;
      font-size: 14px;
      color: #1f2937;
    }
    
    .progress-stats {
      font-size: 13px;
      color: #6b7280;
      font-weight: 600;
    }
    
    .progress-bar-bg {
      background: #e5e7eb;
      height: 8px;
      border-radius: 4px;
      overflow: hidden;
      margin-bottom: 12px;
    }
    
    .progress-bar {
      background: linear-gradient(90deg, #7c3aed 0%, #a78bfa 100%);
      height: 100%;
      width: 0%;
      transition: width 0.3s ease;
      border-radius: 4px;
    }
    
    .progress-details {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 12px;
      margin-bottom: 8px;
    }
    
    .progress-current {
      color: #4b5563;
      font-weight: 500;
    }
    
    .progress-eta {
      color: #6b7280;
    }
    
    .progress-log {
      max-height: 120px;
      overflow-y: auto;
      background: white;
      border: 1px solid #e5e7eb;
      border-radius: 4px;
      padding: 8px;
      font-size: 11px;
      font-family: monospace;
      color: #374151;
      display: none;
    }
    
    .progress-log.show {
      display: block;
    }
    
    .log-entry {
      padding: 2px 0;
    }
    
    .log-success {
      color: #059669;
    }
    
    .log-error {
      color: #dc2626;
    }
    
    .log-info {
      color: #6b7280;
    }
  </style>
</head>
<body>
  <h2>🚀 Auto-Batch Creator (Multi-Product)</h2>
  
  <div class="info">
    <div><strong>Found:</strong> ${totalTasks} task${totalTasks === 1 ? '' : 's'} across ${productCount} product${productCount === 1 ? '' : 's'}</div>
  </div>
  
  <!-- ✅ NEW: Sprint Selection - Improved UX -->
<div class="sprint-section">
  <div class="sprint-header">
    📅 Select Sprint (Required)
    ${sprintLists.length > 0 ? `<button class="sprint-suggest-btn" onclick="selectSuggestedSprint()" type="button">📌 ${sprintLists[0].name}?</button>` : ''}
  </div>
  <select class="sprint-select" id="sprintSelect">
    <option value="">-- Select a sprint --</option>
    ${sprintOptions}
  </select>
  <div class="sprint-helper" id="sprintHelper" style="display:none;">
    ⚠️ Please select a sprint to continue
  </div>
</div>
  
  ${data.warningMissing && data.warningMissing.length > 0 ? `
    <div class="warning-box">
      <div class="warning-title">
        ⚠️ Incomplete Selection Warning
      </div>
      <div>The following columns are missing from your selection:</div>
      <ul class="warning-list">
        ${data.warningMissing.map(col => '<li>' + col + '</li>').join('')}
      </ul>
      <div class="warning-footer">
        ⚠️ Tasks will be created with incomplete data. You'll need to add these fields manually in ClickUp.
      </div>
    </div>
  ` : ''}
  
  ${productSections}
  
  <div class="progress-container" id="progressContainer" style="display:none;">
    <div class="progress-header">
      <div class="progress-title" id="progressTitle">Creating tasks...</div>
      <div class="progress-stats" id="progressStats">0 / 0</div>
    </div>
    <div class="progress-bar-bg">
      <div class="progress-bar" id="progressBar"></div>
    </div>
    <div class="progress-details" id="progressDetails">
      <div class="progress-current" id="progressCurrent">Preparing...</div>
      <div class="progress-eta" id="progressEta"></div>
    </div>
    <div class="progress-log" id="progressLog"></div>
  </div>
  
  <!-- ✅ NEW: Warning when sprint not selected -->
<div class="create-warning" id="createWarning">
  ⚠️ Please select a sprint before creating tasks
</div>

<div class="buttons">
    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn-primary" id="createBtn" disabled onclick="createAllTasks()">Create All Tasks (${totalTasks})</button>
  </div>
  
  <script>
    const tasksByProduct = ${JSON.stringify(tasksByProduct)};
    const productCount = ${productCount};
    
    // Load ClickUp lists on page load
    window.onload = function() {
      google.script.run
        .withSuccessHandler(populateAllLists)
        .withFailureHandler(showError)
        .getClickUpLists();
    };

    // ✅ NEW: Quick-select suggested sprint
function selectSuggestedSprint() {
  const sprintSelect = document.getElementById('sprintSelect');
  if (sprintSelect.options.length > 1) {
    sprintSelect.selectedIndex = 1; // Select first real option (index 0 is placeholder)
    sprintSelect.dispatchEvent(new Event('change')); // Trigger change event
    
    // Visual feedback
    sprintSelect.style.borderColor = '#22c55e';
    setTimeout(() => {
      sprintSelect.style.borderColor = '#86efac';
    }, 300);
  }
}
    
    function populateAllLists(lists) {
  // Populate all product dropdowns
  for (let i = 0; i < productCount; i++) {
    const select = document.getElementById('listSelect' + i);
    
    select.innerHTML = '<option value="">-- Select a list --</option>';
    
    lists.forEach(list => {
      const option = document.createElement('option');
      option.value = list.id;
      option.textContent = list.spaceName + ' → ' + list.folderName + ' → ' + list.name;
      select.appendChild(option);
    });
    
    select.onchange = checkAllSelected;
  }
  
  // ✅ NEW: Add sprint dropdown listener
  const sprintSelect = document.getElementById('sprintSelect');
  if (sprintSelect) {
    sprintSelect.onchange = checkAllSelected;
  }
}
    
    function checkAllSelected() {
  const createBtn = document.getElementById('createBtn');
  const createWarning = document.getElementById('createWarning');
  const sprintHelper = document.getElementById('sprintHelper');
  const sprintSelect = document.getElementById('sprintSelect');
  
  let allSelected = true;
  let sprintSelected = sprintSelect && sprintSelect.value !== '';
  
  // Check product lists
  for (let i = 0; i < productCount; i++) {
    const select = document.getElementById('listSelect' + i);
    if (!select.value) {
      allSelected = false;
      break;
    }
  }
  
  // Show/hide sprint helper
  if (sprintHelper) {
    sprintHelper.style.display = !sprintSelected && allSelected ? 'block' : 'none';
  }
  
  // Show/hide create warning
  if (createWarning) {
    createWarning.classList.toggle('show', !sprintSelected && allSelected);
  }
  
  // Enable button only if both sprint and all products are selected
  createBtn.disabled = !(allSelected && sprintSelected);
}
    
    function createAllTasks() {
  // ✅ NEW: Get selected sprint
  const sprintSelect = document.getElementById('sprintSelect');
  const sprintListId = sprintSelect.value;
  
  if (!sprintListId) {
    // Highlight the sprint dropdown
    sprintSelect.style.borderColor = '#ef4444';
    sprintSelect.focus();
    
    setTimeout(() => {
      sprintSelect.style.borderColor = '#86efac';
    }, 2000);
    
    alert('❌ Please select a sprint first');
    return;
  }
      
      // Collect list selections for each product
      const productListMap = {};
      
      for (let i = 0; i < productCount; i++) {
        const select = document.getElementById('listSelect' + i);
        const product = select.getAttribute('data-product');
        productListMap[product] = select.value;
      }
      
      const createBtn = document.getElementById('createBtn');
      const progressContainer = document.getElementById('progressContainer');
      
      createBtn.disabled = true;
      progressContainer.style.display = 'block';
      
      // Calculate total tasks
      let totalTasks = 0;
      for (const product in tasksByProduct) {
        totalTasks += tasksByProduct[product].length;
      }
      
      // Initialize progress
      updateProgress(0, totalTasks, 'Starting...', null);
      
      // ✅ NEW: Pass sprintListId to task creation
      createTasksWithProgress(tasksByProduct, productListMap, sprintListId, 0, totalTasks);
    }
    
    function createTasksWithProgress(tasksByProduct, productListMap, sprintListId, currentCount, totalTasks) {
      const products = Object.keys(tasksByProduct);
      let productIndex = 0;
      let taskIndex = 0;
      let successCount = 0;
      let errorCount = 0;
      const startTime = Date.now();
      
      function processNextProduct() {
        if (productIndex >= products.length) {
  // All done!
  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  updateProgress(
    totalTasks, 
    totalTasks, 
    '✅ Complete!', 
    elapsed + 's total',
    true
  );
  
  // ✅ Log to Hub for analytics tracking
  if (successCount > 0) {
    google.script.run.logClickUpTasksCreated(successCount, products.join(', '));
  }

  // ✅ NEW: Collect successful row numbers and color them
  const successfulRows = [];
  let currentIndex = 0;
  
  for (const product in tasksByProduct) {
    const tasks = tasksByProduct[product];
    tasks.forEach(task => {
      if (currentIndex < successCount) {
        successfulRows.push(task.rowNumber);
      }
      currentIndex++;
    });
  }
  
  // Color the rows in the spreadsheet
  if (successfulRows.length > 0) {
    addLog('🎨 Coloring ' + successfulRows.length + ' completed rows...', 'info');
    google.script.run
      .withSuccessHandler(function() {
        addLog('✅ Rows colored successfully', 'success');
        
        setTimeout(function() {
          const message = 'Created ' + successCount + ' of ' + totalTasks + ' tasks successfully' +
            (errorCount > 0 ? '\\n\\n' + errorCount + ' errors occurred' : '') +
            '\\n\\nTime: ' + elapsed + 's' +
            '\\n\\n✅ Rows have been marked as completed';
          alert('✅ ' + message);
          google.script.host.close();
        }, 1000);
      })
      .withFailureHandler(function(error) {
        addLog('⚠️ Could not color rows: ' + error.message, 'error');
        
        setTimeout(function() {
          const message = 'Created ' + successCount + ' of ' + totalTasks + ' tasks successfully' +
            (errorCount > 0 ? '\\n\\n' + errorCount + ' errors occurred' : '') +
            '\\n\\nTime: ' + elapsed + 's';
          alert('✅ ' + message);
          google.script.host.close();
        }, 1000);
      })
      .colorCompletedTaskRows(successfulRows);
  } else {
    setTimeout(function() {
      const message = 'Created ' + successCount + ' of ' + totalTasks + ' tasks successfully' +
        (errorCount > 0 ? '\\n\\n' + errorCount + ' errors occurred' : '') +
        '\\n\\nTime: ' + elapsed + 's';
      alert('✅ ' + message);
      google.script.host.close();
    }, 1000);
  }
  return;
}
        
        const product = products[productIndex];
        const tasks = tasksByProduct[product];
        const listId = productListMap[product];
        
        addLog('📦 Processing ' + product + ' (' + tasks.length + ' tasks)...', 'info');
        
        function processNextTask() {
          if (taskIndex >= tasks.length) {
            // Move to next product
            productIndex++;
            taskIndex = 0;
            processNextProduct();
            return;
          }
          
          const task = tasks[taskIndex];
          const currentTaskNum = currentCount + taskIndex + 1;
          
          // Calculate ETA
          const elapsed = (Date.now() - startTime) / 1000;
          const avgTimePerTask = elapsed / currentTaskNum;
          const remaining = totalTasks - currentTaskNum;
          const eta = (avgTimePerTask * remaining).toFixed(0);
          
          updateProgress(
            currentTaskNum,
            totalTasks,
            '📝 ' + product + ': ' + task.taskName,
            eta > 0 ? 'ETA: ' + eta + 's' : null
          );
          
          // ✅ NEW: Pass sprintListId to createSingleTask
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                successCount++;
                addLog('✅ Created: ' + task.taskName, 'success');
              } else {
                errorCount++;
                addLog('❌ Failed: ' + task.taskName, 'error');
              }
              taskIndex++;
              setTimeout(processNextTask, 100);
            })
            .withFailureHandler(function(error) {
              errorCount++;
              addLog('❌ Error: ' + task.taskName + ' - ' + error.message, 'error');
              taskIndex++;
              setTimeout(processNextTask, 100);
            })
            .createSingleTask(task, listId, sprintListId);
        }
        
        processNextTask();
      }
      
      processNextProduct();
    }
    
    function updateProgress(current, total, message, eta, complete) {
      const percent = (current / total * 100).toFixed(0);
      
      document.getElementById('progressBar').style.width = percent + '%';
      document.getElementById('progressStats').textContent = current + ' / ' + total;
      document.getElementById('progressCurrent').textContent = message;
      
      if (eta) {
        document.getElementById('progressEta').textContent = eta;
      }
      
      if (complete) {
        document.getElementById('progressTitle').textContent = '✅ Complete!';
      }
    }
    
    function addLog(message, type) {
      const log = document.getElementById('progressLog');
      log.classList.add('show');
      
      const entry = document.createElement('div');
      entry.className = 'log-entry log-' + type;
      entry.textContent = message;
      
      log.appendChild(entry);
      log.scrollTop = log.scrollHeight;
    }
    
    function handleSuccess(result) {
      alert('✅ Success!\\n\\n' + result.message);
      google.script.host.close();
    }
    
    function handleError(error) {
      document.getElementById('loading').style.display = 'none';
      document.getElementById('createBtn').disabled = false;
      alert('❌ Error: ' + error.message);
    }
    
    function showError(error) {
      alert('❌ Failed to load ClickUp lists: ' + error.message);
    }
  </script>
</body>
</html>`;
  
  return HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(700); // ✅ Increased height to accommodate sprint dropdown
}

/**
 * Fetches ONLY lists from FB Ads folder
 */
function getClickUpLists() {
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  const config = getClickUpConfigLazy_();
  const spaceId = config.FB_ADS_FOLDER_ID; // Performance Creative space
  
  if (!apiKey || apiKey === 'YOUR_CLICKUP_API_KEY') {
    throw new Error('Please set your ClickUp API key in Config');
  }
  
  try {
    Logger.log('🔍 Fetching FB Ads lists from Performance Creative space');
    
    // Get folders in the Performance Creative space
    const foldersUrl = `https://api.clickup.com/api/v2/space/${spaceId}/folder?archived=false`;
    
    const foldersResponse = UrlFetchApp.fetch(foldersUrl, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (foldersResponse.getResponseCode() !== 200) {
      throw new Error('Cannot fetch folders: ' + foldersResponse.getContentText());
    }
    
    const foldersData = JSON.parse(foldersResponse.getContentText());
    const folders = foldersData.folders || [];
    
    Logger.log('✅ Found ' + folders.length + ' folders in Performance Creative');
    
    // Find the "FB Ads" folder
    let fbAdsFolder = null;
    for (let i = 0; i < folders.length; i++) {
      const folder = folders[i];
      const folderName = folder.name.toLowerCase();
      if (folderName.includes('fb ads') || folderName === 'fb ads') {
        fbAdsFolder = folder;
        Logger.log('✅ Found FB Ads folder: ' + folder.name);
        break;
      }
    }
    
    if (!fbAdsFolder) {
      throw new Error('FB Ads folder not found in Performance Creative space');
    }
    
    // Get all lists from FB Ads folder
    const lists = fbAdsFolder.lists || [];
    const allLists = [];
    
    lists.forEach(function(list) {
      // Only include product lists (A01, A02, etc. or B01, B02, etc.)
      const listName = list.name;
      if (/^[AB]\d{2}_/.test(listName)) {
        allLists.push({
          id: list.id,
          name: list.name,
          folderName: 'FB Ads',
          spaceName: 'Performance Creative'
        });
        Logger.log('   📄 ' + list.name);
      }
    });
    
    Logger.log('✅ Total product lists found: ' + allLists.length);
    
    if (allLists.length === 0) {
      throw new Error('No product lists found in FB Ads folder');
    }
    
    // Sort alphabetically by name
    allLists.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    
    return allLists;
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    throw e;
  }
}

/**
 * Fetches the last 5 Sprint lists from Sprint folder
 */
function getSprintLists() {
  const apiKey = getClickUpApiKey();
  const config = getClickUpConfigLazy_();
  const sprintFolderId = config.SPRINT_FOLDER_ID; // Sprint folder
  
  try {
    Logger.log('🔍 Fetching Sprint lists...');
    
    const url = `https://api.clickup.com/api/v2/folder/${sprintFolderId}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('Cannot fetch sprint folder: ' + response.getContentText());
    }
    
    const folderData = JSON.parse(response.getContentText());
    const lists = folderData.lists || [];
    
    Logger.log('✅ Found ' + lists.length + ' sprint lists');
    
    // Extract sprint numbers and sort by number (descending)
    const sprintLists = lists
      .map(function(list) {
        // Extract number from "Sprint XX" format
        const match = list.name.match(/Sprint\s+(\d+)/i);
        const sprintNumber = match ? parseInt(match[1]) : 0;
        
        return {
          id: list.id,
          name: list.name,
          number: sprintNumber
        };
      })
      .filter(function(sprint) {
        return sprint.number > 0; // Only include valid sprint lists
      })
      .sort(function(a, b) {
        return b.number - a.number; // Sort descending (newest first)
      })
      .slice(0, 5); // Take only the last 5
    
    Logger.log('✅ Returning last 5 sprints:');
    sprintLists.forEach(function(sprint) {
      Logger.log('   📅 ' + sprint.name + ' (ID: ' + sprint.id + ')');
    });
    
    return sprintLists;
    
  } catch (e) {
    Logger.log('❌ Error fetching sprint lists: ' + e);
    throw e;
  }
}

/**
 * Creates ClickUp tasks from the extracted row data
 * Relies on ClickUp automation to create sub-tasks
 */
function createClickUpTasksFromRows(tasks, listId) {
  Logger.log('🚀 Creating ' + tasks.length + ' tasks in list ' + listId);
  
  let successCount = 0;
  let errors = [];
  
  tasks.forEach(function(task, index) {
    try {
      Logger.log('\n📝 Task ' + (index + 1) + '/' + tasks.length + ': ' + task.taskName);
      
      const hasPacks = task.quantity > 1; // ✅ Only treat as pack if quantity is 2 or more
      
      // 🆕 Build task name: Product + Task Name + (Pack Nx)
      let taskName = task.product + ' ' + task.taskName;
      
      if (hasPacks) {
        taskName += ' (Pack ' + task.quantity + 'x)';
      }
      
      Logger.log('📝 Final task name: ' + taskName);
      
      // CREATE MASTER TASK ONLY
      const masterTask = createClickUpTask(listId, {
        name: taskName,
        priority: PRIORITY_MAP[task.priority] || 3,
        description: task.description,
        timeEstimate: task.timeEstimate,
        assignee: task.assignee,
        tags: hasPacks ? ['pack'] : [],
        status: 'TO DO'
      });
      
      if (!masterTask || !masterTask.id) {
        throw new Error('Failed to create master task');
      }
      
      Logger.log('✅ Master task created: ' + masterTask.id);
      
      // ALWAYS wait for automation to create sub-task (both pack and non-pack)
      Logger.log('⏳ Waiting for ClickBot automation to create sub-task...');
      
      let subTask = null;
      let attempts = 0;
      const maxAttempts = 6;
      const waitTime = 500;
      
      while (attempts < maxAttempts && !subTask) {
        attempts++;
        Logger.log('⏳ Attempt ' + attempts + '/' + maxAttempts + ' - waiting ' + (waitTime/1000) + ' second(s)...');
        Utilities.sleep(waitTime);
        
        subTask = getFirstSubTask(masterTask.id, listId);
        
        if (subTask) {
          Logger.log('✅ Found auto-created sub-task after ' + attempts + ' attempt(s): ' + subTask.id);
          
          // Update sub-task name to include "(Pack Nx)" if needed
          if (hasPacks) {
            updateSubTaskName(subTask.id, taskName);
            sleepBetweenCalls();
          }
          
          // ALWAYS set sub-task priority to match master
          updateSubTaskPriority(subTask.id, PRIORITY_MAP[task.priority] || 3);
          sleepBetweenCalls();
          
          // ALWAYS clean the sub-task description
          cleanSubTaskDescription(subTask.id, task.description);
          sleepBetweenCalls();

          // ✅ NEW: Auto-assign Ieva to VID tasks (sub-task only)
          if (task.taskName.toLowerCase().includes('vid')) {
            Logger.log('📹 VID task detected - assigning Ieva to sub-task');
            assignUserToTask(subTask.id, 'ieva.b@commercecore.com');
            sleepBetweenCalls();
          }
          
          break;
        } else {
          Logger.log('⏳ Sub-task not found yet, will retry...');
        }
      }
      
      if (!subTask) {
        const totalWaitTime = (maxAttempts * waitTime) / 1000;
        Logger.log('⚠️ Could not find auto-created sub-task after ' + totalWaitTime + ' seconds');
        errors.push('Row ' + task.rowNumber + ': Sub-task not created by automation within ' + totalWaitTime + 's');
      } else {
        // ✅ NEW: Add Ad Notes to non-pack sub-tasks
        if (!hasPacks) {
          sleepBetweenCalls();
          addAdNotesToSubTask(subTask.id);
        }
        
        // Only create sub-sub-tasks if this is a pack task
        if (hasPacks) {
          Logger.log('📦 Creating ' + task.quantity + ' sub-sub-tasks...');
          
          for (let i = 1; i <= task.quantity; i++) {
            sleepBetweenCalls();
            
            const subSubTask = createClickUpTask(listId, {
              name: 'No' + i,
              priority: null,
              description: '',
              timeEstimate: 0,
              assignee: null,
              tags: [],
              status: 'BACKLOG',
              parent: subTask.id
            });
            
            if (subSubTask && subSubTask.id) {
              Logger.log('✅ Sub-sub-task ' + i + ' created: ' + subSubTask.id);
              
              // ✅ NEW: Add Ad Notes template
              sleepBetweenCalls();
              addAdNotesToSubSubTask(subSubTask.id);
            }
          }
          
          Logger.log('✅ All sub-sub-tasks created successfully');
        }
      }
      
      successCount++;
      
    } catch (e) {
      Logger.log('❌ Error on task ' + (index + 1) + ': ' + e);
      Logger.log('   Task name: ' + task.taskName);
      Logger.log('   Product: ' + task.product);
      Logger.log('   Full error: ' + e.toString());
      Logger.log('   Stack: ' + e.stack);
      errors.push('Row ' + task.rowNumber + ' (' + task.taskName + '): ' + e.message);
    }
  });
  
  const message = 'Created ' + successCount + '/' + tasks.length + ' tasks successfully' + 
    (errors.length > 0 ? '\n\nErrors:\n' + errors.join('\n') : '');
  
  return { success: true, message: message };
}

/**
 * Creates a single ClickUp task with all parameters
 */
function createClickUpTask(listId, params) {
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  
  try {
    // Build task payload
    // ✅ Get user-specific default status
    const userEmail = Session.getActiveUser().getEmail();
    const defaultStatus = (userEmail === 'vismantas.k@commercecore.com') ? 'BACKLOG' : 'TO DO';

    const payload = {
      name: params.name,
      status: params.status || defaultStatus
    };
    
    // Add priority (only if not null)
    if (params.priority !== null && params.priority !== undefined) {
      payload.priority = params.priority;
    }
    
    // Add description
    if (params.description) {
      payload.description = params.description;
    }
    
    // Add time estimate (convert hours to milliseconds)
    if (params.timeEstimate && params.timeEstimate > 0) {
      payload.time_estimate = params.timeEstimate * 3600000;
    }
    
    // Add parent (for sub-tasks)
    if (params.parent) {
      payload.parent = params.parent;
    }
    
    // Add tags
    if (params.tags && params.tags.length > 0) {
      payload.tags = params.tags;
    }
    
    // Add assignee
    if (params.assignee) {
      const userId = getClickUpUserIdByEmail(params.assignee);
      if (userId) {
        payload.assignees = [userId];
      }
    }
    
    Logger.log('📤 Creating task: ' + params.name);
    Logger.log('   Payload: ' + JSON.stringify(payload));
    
    const url = 'https://api.clickup.com/api/v2/list/' + listId + '/task';
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (code !== 200) {
      Logger.log('❌ API Error ' + code + ': ' + responseText);
      throw new Error('ClickUp API error: ' + code);
    }
    
    const task = JSON.parse(responseText);
    Logger.log('✅ Task created with ID: ' + task.id);
    
    return task;
    
  } catch (e) {
    logErrorToHub_('Create Task', e.message || String(e));
    Logger.log('❌ Error creating task: ' + e);
    throw e;
  }
}

/**
 * Gets ClickUp user ID by email address
 * Caches results to avoid repeated API calls
 */
const USER_CACHE = {};

function getClickUpUserIdByEmail(assigneeName) {
  // Check if we have email mapping
  const assigneeMap = getAssigneeEmailMapLazy_();
  const email = assigneeMap[assigneeName];
  if (!email) {
    Logger.log(`⚠️ No email mapping for: ${assigneeName}`);
    return null;
  }
  
  // Check cache
  if (USER_CACHE[email]) {
    Logger.log(`✅ Using cached user ID for ${email}: ${USER_CACHE[email]}`);
    return USER_CACHE[email];
  }
  
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  const config = getClickUpConfigLazy_();
  const teamId = config.TEAM_ID;
  
  try {
    Logger.log(`🔍 Looking up user by email: ${email}`);
    
    const url = `https://api.clickup.com/api/v2/team`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('❌ Failed to fetch team members');
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    const teams = data.teams || [];
    
    // Find user by email in team members
    for (const team of teams) {
      const members = team.members || [];
      for (const member of members) {
        const user = member.user;
        if (user && user.email && user.email.toLowerCase() === email.toLowerCase()) {
          Logger.log(`✅ Found user: ${user.username} (ID: ${user.id})`);
          USER_CACHE[email] = user.id;
          return user.id;
        }
      }
    }
    
    Logger.log(`⚠️ User not found with email: ${email}`);
    return null;
    
  } catch (e) {
    Logger.log('❌ Error looking up user: ' + e);
    return null;
  }
}

/**
 * Helper to add delay between API calls (avoid rate limits)
 */
function sleepBetweenCalls() {
  Utilities.sleep(300);
}

/**
 * Gets the first sub-task of a master task by searching for tasks with parent
 */
function getFirstSubTask(masterTaskId, listId) {
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  
  try {
    Logger.log('🔍 Searching for sub-tasks with parent: ' + masterTaskId);
    
    // Get ALL tasks from the list (including subtasks)
    const url = 'https://api.clickup.com/api/v2/list/' + listId + '/task?subtasks=true&include_closed=false';
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    Logger.log('📡 Response code: ' + code);
    
    if (code !== 200) {
      Logger.log('⚠️ API Error: ' + response.getContentText());
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    const tasks = data.tasks || [];
    
    Logger.log('📋 Found ' + tasks.length + ' total tasks in list');
    
    // Find tasks where parent = masterTaskId
    const subTasks = [];
    for (let i = 0; i < tasks.length; i++) {
      const task = tasks[i];
      if (task.parent && task.parent === masterTaskId) {
        subTasks.push(task);
        Logger.log('   ✅ Found sub-task: ' + task.name + ' (ID: ' + task.id + ')');
      }
    }
    
    Logger.log('📋 Total sub-tasks found: ' + subTasks.length);
    
    if (subTasks.length === 0) {
      Logger.log('⚠️ No sub-tasks found with parent ' + masterTaskId);
      return null;
    }
    
    const firstSubTask = subTasks[0];
    Logger.log('✅ Returning first sub-task: ' + firstSubTask.id);
    
    return firstSubTask;
    
  } catch (e) {
    Logger.log('⚠️ Exception in getFirstSubTask: ' + e);
    return null;
  }
}

/**
 * Removes master task description from sub-task, keeps "Creative: ..." part
 */
function cleanSubTaskDescription(subTaskId, masterDescription) {
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  
  try {
    // First, get the current sub-task description
    Logger.log('🔍 Fetching current sub-task description...');
    
    const getUrl = 'https://api.clickup.com/api/v2/task/' + subTaskId;
    
    const getResponse = UrlFetchApp.fetch(getUrl, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (getResponse.getResponseCode() !== 200) {
      Logger.log('⚠️ Could not fetch sub-task');
      return false;
    }
    
    const taskData = JSON.parse(getResponse.getContentText());
    const currentDescription = taskData.description || '';
    
    Logger.log('📋 Current sub-task description: ' + currentDescription);
    
    // Remove the master description part, keep everything else
    let cleanedDescription = currentDescription;
    
    // Remove master description if it exists in the sub-task description
    if (masterDescription && currentDescription.includes(masterDescription)) {
      cleanedDescription = currentDescription.replace(masterDescription, '').trim();
      Logger.log('🧹 Removed master description');
    }
    
    // Remove any extra newlines
    cleanedDescription = cleanedDescription.replace(/\n\n+/g, '\n').trim();
    
    Logger.log('✨ Cleaned description: ' + cleanedDescription);
    
    // Update the sub-task with cleaned description
    const updateUrl = 'https://api.clickup.com/api/v2/task/' + subTaskId;
    
    const payload = {
      description: cleanedDescription
    };
    
    const updateResponse = UrlFetchApp.fetch(updateUrl, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (updateResponse.getResponseCode() === 200) {
      Logger.log('✅ Sub-task description cleaned successfully');
      return true;
    } else {
      Logger.log('⚠️ Could not update sub-task: ' + updateResponse.getContentText());
      return false;
    }
    
  } catch (e) {
    Logger.log('⚠️ Error cleaning sub-task description: ' + e);
    return false;
  }
}

/**
 * Updates sub-task name to include "(Pack Nx)"
 */
function updateSubTaskName(subTaskId, newName) {
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  
  try {
    Logger.log('📝 Updating sub-task name to: ' + newName);
    
    const url = 'https://api.clickup.com/api/v2/task/' + subTaskId;
    
    const payload = {
      name: newName
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      Logger.log('✅ Sub-task name updated');
      return true;
    } else {
      Logger.log('⚠️ Could not update name: ' + response.getContentText());
      return false;
    }
    
  } catch (e) {
    Logger.log('⚠️ Error updating name: ' + e);
    return false;
  }
}

/**
 * Updates sub-task priority to match master task
 */
function updateSubTaskPriority(subTaskId, priority) {
  const apiKey = getClickUpApiKey(); // ✅ CHANGED
  
  try {
    Logger.log('🎯 Setting sub-task priority to: ' + priority);
    
    const url = 'https://api.clickup.com/api/v2/task/' + subTaskId;
    
    const payload = {
      priority: priority
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      Logger.log('✅ Sub-task priority updated');
      return true;
    } else {
      Logger.log('⚠️ Could not update priority: ' + response.getContentText());
      return false;
    }
    
  } catch (e) {
    Logger.log('⚠️ Error updating priority: ' + e);
    return false;
  }
}

/**
 * Creates tasks for multiple products
 */
function createMultiProductTasks(tasksByProduct, productListMap) {
  Logger.log('🚀 Creating multi-product batch');
  
  let totalSuccess = 0;
  let totalTasks = 0;
  let errors = [];
  
  // Process each product
  for (const product in tasksByProduct) {
    const tasks = tasksByProduct[product];
    const listId = productListMap[product];
    
    if (!listId) {
      errors.push(product + ': No list selected');
      continue;
    }
    
    Logger.log('\n📦 Processing product: ' + product + ' (' + tasks.length + ' tasks)');
    Logger.log('   Target list: ' + listId);
    
    totalTasks += tasks.length;
    
    // Create tasks for this product using existing function
    const result = createClickUpTasksFromRows(tasks, listId);
    
    // Count successes (assuming all succeeded if no error)
    totalSuccess += tasks.length;
  }
  
  const message = 'Created ' + totalSuccess + '/' + totalTasks + ' tasks successfully across ' + 
    Object.keys(tasksByProduct).length + ' product(s)' +
    (errors.length > 0 ? '\n\nErrors:\n' + errors.join('\n') : '');
  
  return { success: true, message: message };
}

/**
 * Creates a single task (for progress tracking)
 * Returns immediately so UI can update
 */
function createSingleTask(task, listId, sprintListId) { // ✅ NEW: Added sprintListId parameter
  try {
    const hasPacks = task.quantity > 1;
    
    // Build task name: Product + Task Name + (Pack Nx)
    let taskName = task.product + ' ' + task.taskName;
    
    if (hasPacks) {
      taskName += ' (Pack ' + task.quantity + 'x)';
    }
    
    // CREATE MASTER TASK
    const userEmail = Session.getActiveUser().getEmail();
    const defaultStatus = (userEmail === 'vismantas.k@commercecore.com') ? 'BACKLOG' : 'TO DO';

    const masterTask = createClickUpTask(listId, {
      name: taskName,
      priority: PRIORITY_MAP[task.priority] || 3,
      description: task.description,
      timeEstimate: task.timeEstimate,
      assignee: task.assignee,
      tags: hasPacks ? ['pack'] : [],
      status: defaultStatus
    });
    
    if (!masterTask || !masterTask.id) {
      return { success: false, error: 'Failed to create master task' };
    }
    
    // ✅ NEW: Add master task to sprint list
    Logger.log('📅 Adding master task to sprint...');
    addTaskToSprintList(masterTask.id, sprintListId);
    sleepBetweenCalls();
    
    // Wait for automation to create sub-task
    let subTask = null;
    let attempts = 0;
    const maxAttempts = 6;
    const waitTime = 500;
    
    while (attempts < maxAttempts && !subTask) {
      attempts++;
      Utilities.sleep(waitTime);
      subTask = getFirstSubTask(masterTask.id, listId);
      if (subTask) break;
    }
    
    if (!subTask) {
      return { success: false, error: 'Sub-task not found' };
    }
    
    // Update sub-task
    if (hasPacks) {
      updateSubTaskName(subTask.id, taskName);
      sleepBetweenCalls();
    }

    updateSubTaskPriority(subTask.id, PRIORITY_MAP[task.priority] || 3);
    sleepBetweenCalls();

    cleanSubTaskDescription(subTask.id, task.description);
    sleepBetweenCalls();

    // Auto-assign Ieva to VID tasks (sub-task only)
    if (task.taskName.toLowerCase().includes('vid')) {
      Logger.log('📹 VID task detected - assigning Ieva to sub-task');
      assignUserToTask(subTask.id, 'ieva.b@commercecore.com');
      sleepBetweenCalls();
    }
    
    // Add Ad Notes for non-pack
    if (!hasPacks) {
      addAdNotesToSubTask(subTask.id);
      sleepBetweenCalls();
    }

    // Create sub-sub-tasks if pack
    if (hasPacks) {
      for (let i = 1; i <= task.quantity; i++) {
        sleepBetweenCalls();
        
        const subSubTask = createClickUpTask(listId, {
          name: 'No' + i,
          priority: null,
          description: '',
          timeEstimate: 0,
          assignee: null,
          tags: [],
          status: 'BACKLOG',
          parent: subTask.id
        });
        
        // Add Ad Notes
        if (subSubTask && subSubTask.id) {
          sleepBetweenCalls();
          addAdNotesToSubSubTask(subSubTask.id);
        }
      }
    }
    
    return { success: true };
    
  } catch (e) {
    Logger.log('❌ Error in createSingleTask: ' + e);
    Logger.log('   Task details: ' + JSON.stringify(task));
    return { 
      success: false, 
      error: 'Task "' + task.taskName + '": ' + e.message 
    };
  }
}

/**
 * Adds Ad Notes template to sub-task description (for non-pack tasks)
 * APPENDS to existing description without removing anything
 */
function addAdNotesToSubTask(subTaskId) {
  const apiKey = getClickUpApiKey();
  
  try {
    Logger.log('📝 Adding Ad Notes to sub-task: ' + subTaskId);
    
    // Get current description
    const getUrl = 'https://api.clickup.com/api/v2/task/' + subTaskId;
    const getResponse = UrlFetchApp.fetch(getUrl, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (getResponse.getResponseCode() !== 200) {
      Logger.log('⚠️ Could not fetch sub-task');
      return false;
    }
    
    const taskData = JSON.parse(getResponse.getContentText());
    const currentDescription = taskData.description || '';
    
    // Append template with bold formatting
    const newDescription = currentDescription + '\n\nAd Notes:\n\nAds Folder:';
    
    // Update
    const updateUrl = 'https://api.clickup.com/api/v2/task/' + subTaskId;
    const updateResponse = UrlFetchApp.fetch(updateUrl, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ description: newDescription }),
      muteHttpExceptions: true
    });
    
    if (updateResponse.getResponseCode() === 200) {
      Logger.log('✅ Ad Notes added to sub-task');
      return true;
    }
    
    return false;
    
  } catch (e) {
    Logger.log('⚠️ Error: ' + e);
    return false;
  }
}

/**
 * Adds Ad Notes template to sub-sub-task description (for pack tasks)
 */
function addAdNotesToSubSubTask(subSubTaskId) {
  const apiKey = getClickUpApiKey();
  
  try {
    Logger.log('📝 Adding Ad Notes to sub-sub-task: ' + subSubTaskId);
    
    const url = 'https://api.clickup.com/api/v2/task/' + subSubTaskId;
    const response = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ description: 'Ad Notes:\n\nAds Folder:' }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      Logger.log('✅ Ad Notes added to sub-sub-task');
      return true;
    }
    
    return false;
    
  } catch (e) {
    Logger.log('⚠️ Error: ' + e);
    return false;
  }
}

/**
 * Adds an existing task to a sprint list (multi-list feature)
 * Uses the task update endpoint to add to multiple lists
 */
function addTaskToSprintList(taskId, sprintListId) {
  const apiKey = getClickUpApiKey();
  
  try {
    Logger.log('📅 Adding task ' + taskId + ' to sprint list ' + sprintListId);
    
    // First, get the task to see its current lists
    const getUrl = 'https://api.clickup.com/api/v2/task/' + taskId;
    
    const getResponse = UrlFetchApp.fetch(getUrl, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (getResponse.getResponseCode() !== 200) {
      Logger.log('⚠️ Could not fetch task: ' + getResponse.getContentText());
      return false;
    }
    
    const taskData = JSON.parse(getResponse.getContentText());
    const currentLists = taskData.list ? [taskData.list.id] : [];
    
    Logger.log('📋 Current lists for task: ' + JSON.stringify(currentLists));
    Logger.log('📋 Adding sprint list: ' + sprintListId);
    
    // Check if already in sprint list
    if (currentLists.includes(sprintListId)) {
      Logger.log('✅ Task already in sprint list');
      return true;
    }
    
    // Add sprint list to the task
    const addUrl = 'https://api.clickup.com/api/v2/list/' + sprintListId + '/task/' + taskId;
    
    const addResponse = UrlFetchApp.fetch(addUrl, {
      method: 'post',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const code = addResponse.getResponseCode();
    const responseText = addResponse.getContentText();
    
    Logger.log('📡 Add to sprint response code: ' + code);
    Logger.log('📡 Add to sprint response body: ' + responseText);
    
    if (code === 200) {
      Logger.log('✅ Task added to sprint list successfully');
      return true;
    } else {
      Logger.log('⚠️ Could not add task to sprint (code ' + code + '): ' + responseText);
      return false;
    }
    
  } catch (e) {
    Logger.log('⚠️ Error adding task to sprint: ' + e);
    Logger.log('⚠️ Stack trace: ' + e.stack);
    return false;
  }
}

/**
 * Colors multiple completed task rows at once (more efficient than one-by-one)
 */
function colorCompletedTaskRows(rowNumbers) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const completedColor = '#fce5cd'; // Light orange (most likely match)
    
    Logger.log('🎨 Coloring ' + rowNumbers.length + ' completed rows...');
    
    // Color each row
    rowNumbers.forEach(function(rowNumber) {
      const range = sheet.getRange(rowNumber, 3, 1, 8); // C to J
      range.setBackground(completedColor);
    });
    
    Logger.log('✅ Successfully colored ' + rowNumbers.length + ' rows');
    return true;
    
  } catch (e) {
    Logger.log('⚠️ Error coloring rows: ' + e);
    return false;
  }
}

/**
 * Assigns a user to a task by email
 */
function assignUserToTask(taskId, userEmail) {
  const apiKey = getClickUpApiKey();
  
  try {
    Logger.log('👤 Assigning ' + userEmail + ' to task ' + taskId);
    
    // Get user ID from email (direct lookup, not through name mapping)
    const userId = getClickUpUserIdByEmailDirect(userEmail);
    
    if (!userId) {
      Logger.log('⚠️ Could not find user ID for ' + userEmail);
      return false;
    }
    
    const url = 'https://api.clickup.com/api/v2/task/' + taskId;
    
    const payload = {
      assignees: {
        add: [userId]
      }
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      Logger.log('✅ User assigned successfully');
      return true;
    } else {
      Logger.log('⚠️ Could not assign user: ' + response.getContentText());
      return false;
    }
    
  } catch (e) {
    Logger.log('⚠️ Error assigning user: ' + e);
    return false;
  }
}

/**
 * Gets ClickUp user ID directly by email (bypasses the name mapping)
 */
function getClickUpUserIdByEmailDirect(email) {
  // Check cache first
  if (USER_CACHE[email]) {
    Logger.log('✅ Using cached user ID for ' + email + ': ' + USER_CACHE[email]);
    return USER_CACHE[email];
  }
  
  const apiKey = getClickUpApiKey();
  
  try {
    Logger.log('🔍 Looking up user by email: ' + email);
    
    const url = 'https://api.clickup.com/api/v2/team';
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('❌ Failed to fetch team members');
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    const teams = data.teams || [];
    
    // Find user by email in team members
    for (const team of teams) {
      const members = team.members || [];
      for (const member of members) {
        const user = member.user;
        if (user && user.email && user.email.toLowerCase() === email.toLowerCase()) {
          Logger.log('✅ Found user: ' + user.username + ' (ID: ' + user.id + ')');
          USER_CACHE[email] = user.id;
          return user.id;
        }
      }
    }
    
    Logger.log('⚠️ User not found with email: ' + email);
    return null;
    
  } catch (e) {
    Logger.log('❌ Error looking up user: ' + e);
    return null;
  }
}

/**
 * Test function to trigger permissions
 */
function testPermissions() {
  try {
    const email = Session.getActiveUser().getEmail();
    Logger.log('✅ Email access granted: ' + email);
    
    const apiKeys = getClickUpApiKeysLazy_();
    if (apiKeys[email]) {
      Logger.log('✅ API key found for ' + email);
    } else {
      Logger.log('⚠️ No API key configured for ' + email);
    }
    
    SpreadsheetApp.getUi().alert('✅ Permissions granted!\n\nYour email: ' + email);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * Test function to check if Ieva can be found in ClickUp
 * This is SAFE to run - it only looks up the user, doesn't assign anything
 */
function testIevaLookup() {
  try {
    Logger.log('🧪 Testing Ieva lookup...');
    
    const userId = getClickUpUserIdByEmailDirect('ieva.b@commercecore.com');
    
    if (userId) {
      SpreadsheetApp.getUi().alert('✅ Success!\n\nIeva found in ClickUp\nUser ID: ' + userId);
    } else {
      SpreadsheetApp.getUi().alert('❌ Failed\n\nIeva not found in ClickUp.\n\nCheck:\n1. Email is correct\n2. Ieva is in your ClickUp workspace');
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

function runAutoBatchDiagnostics() {
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
  Logger.log('║        AUTO-BATCH CREATOR - CONFIG DIAGNOSTICS REPORT           ║');
  Logger.log('║        ' + new Date().toLocaleString() + '                              ║');
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
    
    const requiredTabs = ['Team Members', 'ClickUp Config'];
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
  // TEST 2: CLICKUP API KEYS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔑 TEST 2: CLICKUP API KEYS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const apiKeys = getClickUpApiKeysLazy_();
    if (typeof apiKeys === 'object' && Object.keys(apiKeys).length > 0) {
      const keyCount = Object.keys(apiKeys).length;
      log('APIKeys', 'Load API Keys', 'PASS', `${keyCount} user(s) with keys`);
      
      // Check key format (should start with pk_)
      let validKeys = 0;
      for (const [email, key] of Object.entries(apiKeys)) {
        if (key && key.startsWith('pk_')) {
          validKeys++;
        }
      }
      
      if (validKeys === keyCount) {
        log('APIKeys', 'Key Format Validation', 'PASS', 'All keys start with pk_');
      } else {
        log('APIKeys', 'Key Format Validation', 'WARN', `${keyCount - validKeys} invalid key format(s)`);
      }
      
      // Check current user has key
      try {
        const currentUser = Session.getActiveUser().getEmail();
        if (currentUser && apiKeys[currentUser]) {
          log('APIKeys', 'Current User Key', 'PASS', `Key found for ${currentUser}`);
        } else if (currentUser) {
          log('APIKeys', 'Current User Key', 'WARN', `No key for ${currentUser}`);
        }
      } catch (e) {
        log('APIKeys', 'Current User Key', 'WARN', 'Could not get current user email');
      }
      
      Logger.log('   Users with keys: ' + Object.keys(apiKeys).slice(0, 5).join(', ') + (keyCount > 5 ? '...' : ''));
    } else {
      log('APIKeys', 'Load API Keys', 'FAIL', 'No API keys found');
    }
  } catch (e) {
    log('APIKeys', 'Load API Keys', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 3: ASSIGNEE EMAIL MAP
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('👥 TEST 3: ASSIGNEE EMAIL MAP');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const assignees = getAssigneeEmailMapLazy_();
    if (typeof assignees === 'object' && Object.keys(assignees).length > 0) {
      const count = Object.keys(assignees).length;
      log('Assignees', 'Load Assignee Map', 'PASS', `${count} assignee(s)`);
      
      // Validate email format
      let validEmails = 0;
      for (const [name, email] of Object.entries(assignees)) {
        if (email && email.includes('@')) {
          validEmails++;
        }
      }
      
      if (validEmails === count) {
        log('Assignees', 'Email Format', 'PASS', 'All emails valid');
      } else {
        log('Assignees', 'Email Format', 'WARN', `${count - validEmails} invalid email(s)`);
      }
      
      Logger.log('   Assignees: ' + Object.entries(assignees).slice(0, 5).map(([n, e]) => `${n}→${e.split('@')[0]}`).join(', '));
    } else {
      log('Assignees', 'Load Assignee Map', 'FAIL', 'No assignees found');
    }
  } catch (e) {
    log('Assignees', 'Load Assignee Map', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 4: CLICKUP CONFIG
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('⚙️ TEST 4: CLICKUP CONFIG');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const config = getClickUpConfigLazy_();
    if (typeof config === 'object' && Object.keys(config).length > 0) {
      log('Config', 'Load ClickUp Config', 'PASS', `${Object.keys(config).length} setting(s)`);
      
      // Check required keys
      const requiredKeys = ['TEAM_ID', 'FB_ADS_FOLDER_ID', 'SPRINT_FOLDER_ID'];
      const missingKeys = [];
      
      for (const key of requiredKeys) {
        if (config[key]) {
          Logger.log(`   ${key}: ${config[key]}`);
        } else {
          missingKeys.push(key);
        }
      }
      
      if (missingKeys.length === 0) {
        log('Config', 'Required Keys Present', 'PASS', 'All 4 required keys found');
      } else {
        log('Config', 'Required Keys Present', 'FAIL', `Missing: ${missingKeys.join(', ')}`);
      }
      
      // Check TEAM_ID is being used
      log('Config', 'TEAM_ID', config.TEAM_ID ? 'PASS' : 'FAIL', config.TEAM_ID || 'Not set');
      
    } else {
      log('Config', 'Load ClickUp Config', 'FAIL', 'No config found');
    }
  } catch (e) {
    log('Config', 'Load ClickUp Config', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 5: FUNCTION VALIDATION (Check for hardcoded values)
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔧 TEST 5: FUNCTION VALIDATION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  // Check getClickUpLists uses CLICKUP_CONFIG for spaceId
  try {
    const funcStr = getClickUpLists.toString();
    if (funcStr.includes('CLICKUP_CONFIG.FB_ADS_FOLDER_ID')) {
      log('Functions', 'getClickUpLists - spaceId', 'PASS', 'Uses CLICKUP_CONFIG');
    } else if (funcStr.includes("'90150821271'")) {
      log('Functions', 'getClickUpLists - spaceId', 'FAIL', 'Still hardcoded - apply FIX 1');
    } else {
      log('Functions', 'getClickUpLists - spaceId', 'WARN', 'Could not verify');
    }
  } catch (e) {
    log('Functions', 'getClickUpLists', 'WARN', 'Could not inspect function');
  }
  
  // Check getSprintLists uses CLICKUP_CONFIG for folderId
  try {
    const funcStr = getSprintLists.toString();
    if (funcStr.includes('CLICKUP_CONFIG.SPRINT_FOLDER_ID')) {
      log('Functions', 'getSprintLists - folderId', 'PASS', 'Uses CLICKUP_CONFIG');
    } else if (funcStr.includes("'901512650311'")) {
      log('Functions', 'getSprintLists - folderId', 'FAIL', 'Still hardcoded - apply FIX 2');
    } else {
      log('Functions', 'getSprintLists - folderId', 'WARN', 'Could not verify');
    }
  } catch (e) {
    log('Functions', 'getSprintLists', 'WARN', 'Could not inspect function');
  }

  // ============================================================================
  // TEST 6: CLICKUP API CONNECTION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🌐 TEST 6: CLICKUP API CONNECTION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const apiKey = getClickUpApiKey();
    if (apiKey && apiKey.startsWith('pk_')) {
      log('API', 'Get API Key', 'PASS', 'Key retrieved successfully');
      
      // Test API connection
      try {
        const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2/team', {
          method: 'get',
          headers: {
            'Authorization': apiKey,
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        });
        
        if (response.getResponseCode() === 200) {
          const data = JSON.parse(response.getContentText());
          const teamCount = data.teams ? data.teams.length : 0;
          log('API', 'ClickUp Connection', 'PASS', `Connected, ${teamCount} team(s) accessible`);
        } else {
          log('API', 'ClickUp Connection', 'FAIL', `HTTP ${response.getResponseCode()}`);
        }
      } catch (e) {
        log('API', 'ClickUp Connection', 'FAIL', e.message);
      }
    } else {
      log('API', 'Get API Key', 'FAIL', 'No valid API key');
    }
  } catch (e) {
    log('API', 'Get API Key', 'WARN', e.message);
  }

  // ============================================================================
  // TEST 7: CACHE PERFORMANCE
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('⚡ TEST 7: CACHE PERFORMANCE');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Clear cache
    CacheService.getScriptCache().removeAll(['abc_api_keys', 'abc_assignees', 'abc_clickup']);
    
    // Uncached load
    const uncachedStart = Date.now();
    getClickUpApiKeysFromConfig_();
    getAssigneeEmailMapFromConfig_();
    getClickUpConfigFromConfig_();
    const uncachedTime = Date.now() - uncachedStart;
    
    // Cached load
    const cachedStart = Date.now();
    getClickUpApiKeysFromConfig_();
    getAssigneeEmailMapFromConfig_();
    getClickUpConfigFromConfig_();
    const cachedTime = Date.now() - cachedStart;
    
    log('Cache', 'Uncached Load', uncachedTime < 5000 ? 'PASS' : 'WARN', `${uncachedTime}ms`);
    log('Cache', 'Cached Load', cachedTime < 100 ? 'PASS' : 'WARN', `${cachedTime}ms`);
    log('Cache', 'Speedup', cachedTime < uncachedTime ? 'PASS' : 'WARN', 
        `${Math.round(uncachedTime / Math.max(cachedTime, 1))}x faster`);
  } catch (e) {
    log('Cache', 'Cache Test', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 8: EDGE CASES
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔬 TEST 8: EDGE CASES');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Test priority mapping
    const urgentPriority = PRIORITY_MAP['Urgent'];
    const normalPriority = PRIORITY_MAP['Normal'];
    log('EdgeCase', 'Priority Mapping', urgentPriority === 1 && normalPriority === 3 ? 'PASS' : 'FAIL',
        `Urgent=${urgentPriority}, Normal=${normalPriority}`);
    
    // Test non-existent user API key
    const apiKeys = getClickUpApiKeysLazy_();
    const fakeUserKey = apiKeys['fake.user@example.com'];
    log('EdgeCase', 'Non-existent User Key', fakeUserKey === undefined ? 'PASS' : 'WARN',
        fakeUserKey === undefined ? 'Returns undefined' : 'Unexpected value');
    
    // Test non-existent assignee
    const assignees = getAssigneeEmailMapLazy_();
    const fakeAssignee = assignees['FakeUser'];
    log('EdgeCase', 'Non-existent Assignee', fakeAssignee === undefined ? 'PASS' : 'WARN',
        fakeAssignee === undefined ? 'Returns undefined' : 'Unexpected value');
        
  } catch (e) {
    log('EdgeCase', 'Edge Case Tests', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 9: CROSS-REFERENCE VALIDATION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔀 TEST 9: CROSS-REFERENCE VALIDATION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Check if users with API keys are in assignee map
    const apiKeyEmails = new Set(Object.keys(getClickUpApiKeysLazy_()));
    const assigneeEmails = new Set(Object.values(getAssigneeEmailMapLazy_()));
    
    const keysWithoutAssignee = [...apiKeyEmails].filter(e => !assigneeEmails.has(e));
    
    if (keysWithoutAssignee.length === 0) {
      log('CrossRef', 'API Keys ⊂ Assignees', 'PASS', 'All API key users are assignees');
    } else {
      log('CrossRef', 'API Keys ⊂ Assignees', 'WARN', 
          `${keysWithoutAssignee.length} API key users not in assignee map`);
    }
    
    // Check assignees have API keys
    const assigneesWithoutKeys = [...assigneeEmails].filter(e => !apiKeyEmails.has(e));
    if (assigneesWithoutKeys.length === 0) {
      log('CrossRef', 'Assignees → API Keys', 'PASS', 'All assignees have API keys');
    } else {
      log('CrossRef', 'Assignees → API Keys', 'WARN', 
          `${assigneesWithoutKeys.length} assignee(s) without API keys: ${assigneesWithoutKeys.slice(0,3).join(', ')}`);
    }
  } catch (e) {
    log('CrossRef', 'Cross-Reference', 'FAIL', e.message);
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
    'Auto-Batch Creator Diagnostics',
    `✅ Passed: ${results.passed}\n⚠️ Warnings: ${results.warnings}\n❌ Failed: ${results.failed}\n\nTime: ${totalTime}ms\n\nSee Execution Log for details.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  return results;
}

/**
 * Quick health check for Auto-Batch Creator
 */
function quickAutoBatchHealthCheck() {
  const checks = [];
  
  try {
    checks.push({ name: 'Config Sheet', ok: !!SpreadsheetApp.openById(CONFIG_SHEET_ID) });
  } catch (e) {
    checks.push({ name: 'Config Sheet', ok: false });
  }
  
  try {
    const apiKeys = getClickUpApiKeysLazy_();
    checks.push({ name: 'API Keys', ok: typeof apiKeys === 'object' && Object.keys(apiKeys).length > 0 });
  } catch (e) {
    checks.push({ name: 'API Keys', ok: false });
  }
  
  try {
    const assignees = getAssigneeEmailMapLazy_();
    checks.push({ name: 'Assignees', ok: typeof assignees === 'object' && Object.keys(assignees).length > 0 });
  } catch (e) {
    checks.push({ name: 'Assignees', ok: false });
  }
  
  try {
    const config = getClickUpConfigLazy_();
    checks.push({ name: 'ClickUp Config', ok: typeof config === 'object' && !!config.TEAM_ID });
  } catch (e) {
    checks.push({ name: 'ClickUp Config', ok: false });
  }
  
  try {
    const apiKey = getClickUpApiKey();
    checks.push({ name: 'Current User Key', ok: !!apiKey && apiKey.startsWith('pk_') });
  } catch (e) {
    checks.push({ name: 'Current User Key', ok: false });
  }
  
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