# Google Apps Script Development Skill

## Overview
This skill contains best practices for developing Google Apps Script projects, particularly those involving:
- Web apps with HTML Service
- Google Sheets integration
- External API calls (ClickUp, Slack, etc.)
- Central configuration management

## Architecture Patterns

### Central Configuration
Always use a central config sheet for multi-system setups:

```javascript
const CONFIG_SHEET_ID = 'your-config-sheet-id';
const CONFIG_CACHE_DURATION = 300; // 5 minutes

function getConfigValue_(key) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'config_' + key;
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Config');
  // ... fetch value
  cache.put(cacheKey, value, CONFIG_CACHE_DURATION);
  return value;
}
```

### Caching Strategy
```javascript
// Two-level caching pattern
let _inMemoryCache = null; // Persists within execution

function getData_() {
  // Level 1: In-memory (fastest, within same execution)
  if (_inMemoryCache) return _inMemoryCache;
  
  // Level 2: Script cache (persists across executions)
  const cache = CacheService.getScriptCache();
  const cached = cache.get('data_key');
  if (cached) {
    _inMemoryCache = JSON.parse(cached);
    return _inMemoryCache;
  }
  
  // Level 3: Fetch from source
  const data = fetchFromSheet_();
  cache.put('data_key', JSON.stringify(data), 300);
  _inMemoryCache = data;
  return data;
}

// CRITICAL: Clear both levels when data changes
function clearCache_() {
  CacheService.getScriptCache().remove('data_key');
  _inMemoryCache = null;
}
```

### Function Naming Convention
```javascript
// PUBLIC - Can be called from client-side google.script.run
function saveData(data) { }
function getDashboardData() { }

// PRIVATE - Cannot be called from client-side (underscore suffix)
function fetchFromApi_() { }
function validateData_(data) { }
function getConfigValue_(key) { }
```

⚠️ **CRITICAL BUG PATTERN**: If `google.script.run.someFunction_()` silently fails, the underscore makes it private!

## Web App Patterns

### HTML Service Setup
```javascript
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  
  // Pass server-side data to template
  template.userData = {
    email: Session.getActiveUser().getEmail(),
    name: getUserName_(),
    isAdmin: checkAdmin_()
  };
  
  return template.evaluate()
    .setTitle('App Title')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
```

### Client-Server Communication
```javascript
// CLIENT SIDE (index.html)
google.script.run
  .withSuccessHandler(function(result) {
    if (result.success) {
      // Handle success
    } else {
      showToast('Error: ' + result.error, 'error');
    }
  })
  .withFailureHandler(function(error) {
    showToast('Error: ' + error.message, 'error');
  })
  .serverFunction(data);

// SERVER SIDE (code.gs)
function serverFunction(data) {
  try {
    // Do work
    return { success: true, data: result };
  } catch (e) {
    Logger.log('Error in serverFunction: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
```

### Optimistic UI Pattern
```javascript
// CLIENT: Show immediate feedback, then confirm
function deleteItem(item) {
  // 1. Immediately show "deleting" state
  markAsDeleting(item.id);
  showDeletingToast('Deleting...');
  
  // 2. Call server
  google.script.run
    .withSuccessHandler(function(result) {
      if (result.success) {
        // 3a. Success: Fetch fresh data, remove item
        fetchFreshData();
      } else {
        // 3b. Failed: Restore item, show error
        unmarkAsDeleting(item.id);
        showToast('Failed: ' + result.error, 'error');
      }
    })
    .withFailureHandler(function(error) {
      unmarkAsDeleting(item.id);
      showToast('Error: ' + error.message, 'error');
    })
    .deleteItem(item.id);
}
```

## External API Integration

### ClickUp API Pattern
```javascript
function callClickUpApi_(endpoint, method, payload) {
  const apiKey = getClickUpApiKey_();
  
  const options = {
    method: method || 'GET',
    headers: {
      'Authorization': apiKey,  // NOT "Bearer " prefix!
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  if (payload) {
    options.payload = JSON.stringify(payload);
  }
  
  try {
    const response = UrlFetchApp.fetch(
      'https://api.clickup.com/api/v2/' + endpoint,
      options
    );
    
    const code = response.getResponseCode();
    const body = JSON.parse(response.getContentText());
    
    if (code === 401) {
      // Often NOT an auth issue - check if IDs are correct!
      Logger.log('401 Error - Check Space/Folder IDs');
    }
    
    return { success: code >= 200 && code < 300, data: body, code: code };
  } catch (e) {
    Logger.log('ClickUp API Error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
```

⚠️ **CRITICAL**: ClickUp IDs are strings. Always use `==` not `===` for comparisons.

### Slack Webhook Pattern
```javascript
function sendSlackMessage_(webhookUrl, message, blocks) {
  const payload = {
    text: message
  };
  
  if (blocks) {
    payload.blocks = blocks;
  }
  
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return true;
  } catch (e) {
    Logger.log('Slack error: ' + e.toString());
    return false;
  }
}
```

## Google Sheets Patterns

### Reading Data
```javascript
function getSheetData_(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Convert to array of objects
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}
```

### Writing Data (Avoiding Format Copy Bug)
```javascript
// BAD: insertRowAfter copies formatting from row above
sheet.insertRowAfter(1);
sheet.getRange(2, 1, 1, 5).setValues([data]);  // Gets header formatting!

// GOOD: Clear formatting after insert
sheet.insertRowAfter(1);
const newRow = sheet.getRange(2, 1, 1, 5);
newRow.setValues([data]);
newRow.setBackground(null).setFontWeight('normal');  // Clear copied formatting
```

### Batch Operations
```javascript
// BAD: Multiple API calls
data.forEach(row => {
  sheet.appendRow(row);  // Slow!
});

// GOOD: Single batch operation
sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
```

## Error Handling

### Logging Pattern
```javascript
function riskyOperation_() {
  Logger.log('🚀 Starting riskyOperation_');
  
  try {
    Logger.log('📋 Step 1: Fetching data...');
    const data = fetchData_();
    Logger.log('✅ Fetched ' + data.length + ' items');
    
    Logger.log('📋 Step 2: Processing...');
    const result = process_(data);
    Logger.log('✅ Processed successfully');
    
    return { success: true, data: result };
  } catch (e) {
    Logger.log('❌ Error in riskyOperation_: ' + e.toString());
    Logger.log('📍 Stack: ' + e.stack);
    return { success: false, error: e.toString() };
  }
}
```

### Error Logging to Sheet
```javascript
function logErrorToSheet_(system, errorType, message) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let sheet = ss.getSheetByName('Error Log');
    
    if (!sheet) {
      sheet = ss.insertSheet('Error Log');
      sheet.getRange(1, 1, 1, 5).setValues([
        ['Timestamp', 'System', 'Error Type', 'Message', 'User']
      ]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
    sheet.insertRowAfter(1);
    const newRow = sheet.getRange(2, 1, 1, 5);
    newRow.setValues([[
      new Date().toISOString(),
      system,
      errorType,
      String(message).substring(0, 500),
      Session.getActiveUser().getEmail() || 'Unknown'
    ]]);
    newRow.setBackground(null).setFontWeight('normal');
  } catch (e) {
    Logger.log('Could not log error: ' + e);
  }
}
```

## UI/UX Patterns for Web Apps

### Theme Support
```css
:root {
  --bg: #fdf6f3;
  --card: rgba(255,255,255,0.7);
  --text: #1e293b;
  --accent: #ff6b35;
}

[data-theme="dark"] {
  --bg: #0f172a;
  --card: rgba(30,41,59,0.8);
  --text: #f1f5f9;
  --accent: #ff6b35;
}
```

### Toast Notifications
```javascript
function showToast(message, type = 'info', duration = 3000) {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type} show`;
  toast.textContent = message;
  document.getElementById('toast-container').appendChild(toast);
  
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, duration);
}
```

### Loading States
```javascript
// Show skeleton loaders during data fetch
function renderSkeletonCards(count) {
  return Array(count).fill(0).map(() => `
    <div class="skeleton-card">
      <div class="skeleton-line"></div>
      <div class="skeleton-line short"></div>
    </div>
  `).join('');
}
```

## Common Gotchas

1. **Private functions** (`function_()`) cannot be called from `google.script.run`
2. **insertRowAfter()** copies formatting - always clear it
3. **ClickUp IDs are strings** - use `==` not `===`
4. **Cache has two levels** - clear both in-memory AND script cache
5. **401 errors** often mean wrong IDs, not auth issues
6. **Session.getActiveUser()** may return empty in some contexts
7. **UrlFetchApp timeout** is 60 seconds - plan for it
8. **Spreadsheet operations are slow** - batch them when possible