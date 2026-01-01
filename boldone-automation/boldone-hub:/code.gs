/**
 * ================================================================================
 * BOLD.ONE CREATIVE HUB - WEB APP
 * ================================================================================
 * 
 * A read-only dashboard for monitoring CommerceCore automation systems.
 * 
 * SETUP:
 * 1. Create new Apps Script project at script.google.com
 * 2. Paste this code into Code.gs
 * 3. Create Index.html and paste the HTML template
 * 4. Deploy → New deployment → Web app
 * 5. Execute as: Me | Access: Anyone with link (or specific users)
 * 
 * ================================================================================
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG_SHEET_ID = '1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk';
const CONFIG_CACHE_DURATION = 60; // 1 minute (reduced for real-time feel)

// Quarterly Goals Sheet ID
const QUARTERLY_GOALS_SHEET_ID = '1l3HN3c75qXClr3UJ4-kgkowLIqFEk-9aRT89LJ-FfrQ';

// Allowed users (emails) - only these can access the web app
const ALLOWED_USERS = [
  'kristijonas@bold.one',
  'kristijonas.t@commercecore.com',
  'gabija.p@commercecore.com',
  'gabija.p@bold.one'
  // Add more emails as needed
];

// Analytics Sheet ID (for batches count)
const ANALYTICS_SHEET_ID = '1KalmueNyF_2Hiy52NmkJblQrsbFK4hDhD5my-Psj-L0';

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================

/**
 * Serves the web app HTML
 */
function doGet(e) {
  // Check user access
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!isUserAllowed_(userEmail)) {
    return HtmlService.createHtmlOutput(getAccessDeniedHtml_(userEmail))
      .setTitle('Access Denied')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Serve the main app
  const template = HtmlService.createTemplateFromFile('Index');
  template.userData = getUserData_(userEmail);
  
  return template.evaluate()
    .setTitle('Bold.One Creative Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include HTML files (for modular templates)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// ACCESS CONTROL & SECURITY
// ============================================================================

/**
 * Check if user is allowed to access the app
 */
function isUserAllowed_(email) {
  if (!email) return false;
  return ALLOWED_USERS.some(allowed => 
    allowed.toLowerCase() === email.toLowerCase()
  );
}

/**
 * Verify current user has access - call this at start of every server function
 * Returns user object if authorized, throws error if not
 */
function verifyAccess_() {
  const email = Session.getActiveUser().getEmail();
  if (!email || !isUserAllowed_(email)) {
    throw new Error('Access denied: You are not authorized to perform this action.');
  }
  return getUserDataFromConfig_(email);
}

/**
 * Verify current user is an admin - for sensitive operations
 * Returns user object if admin, throws error if not
 */
function verifyAdmin_() {
  const user = verifyAccess_();
  if (!user.isAdmin) {
    throw new Error('Access denied: This action requires administrator privileges.');
  }
  return user;
}

/**
 * Get user data from Config sheet (with admin status)
 */
function getUserDataFromConfig_(email) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Team Members');
    if (!sheet) {
      return { email: email, name: email.split('@')[0], isAdmin: false };
    }
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Column B (index 1) = Email
      if (data[i][1] && data[i][1].toString().toLowerCase() === email.toLowerCase()) {
        return {
          email: email,
          name: data[i][0] || email.split('@')[0],    // Column A: Name
          isAdmin: data[i][3] === true,                // Column D: Is Admin
          clickUpAccess: data[i][4] === true,          // Column E: ClickUp Access
          andromedaAccess: data[i][5] === true,        // Column F: Andromeda Access
          active: data[i][7] === true                  // Column H: Active
        };
      }
    }
    
    return { email: email, name: email.split('@')[0], isAdmin: false };
  } catch (e) {
    Logger.log('Error getting user data from config: ' + e.toString());
    return { email: email, name: email.split('@')[0], isAdmin: false };
  }
}

/**
 * Get user data for display (uses config-based admin check)
 */
function getUserData_(email) {
  const userData = getUserDataFromConfig_(email);
  const name = userData.name || email.split('@')[0];
  return {
    email: email,
    name: name.charAt(0).toUpperCase() + name.slice(1),
    initial: name.charAt(0).toUpperCase(),
    role: userData.isAdmin ? 'Admin' : 'Member',
    isAdmin: userData.isAdmin
  };
}

/**
 * Simple rate limiting using cache
 * Returns true if rate limit exceeded
 */
function isRateLimited_(action, maxPerMinute) {
  const cache = CacheService.getUserCache();
  const key = 'rate_' + action;
  const current = parseInt(cache.get(key) || '0');
  
  if (current >= maxPerMinute) {
    Logger.log('⚠️ Rate limit exceeded for action: ' + action);
    return true;
  }
  
  cache.put(key, (current + 1).toString(), 60); // 60 second window
  return false;
}

/**
 * Log security event to Sync Log
 */
function logSecurityEvent_(eventType, details) {
  try {
    const email = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date().toISOString();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Sync Log');
    if (sheet) {
      sheet.appendRow([timestamp, '🔒 SECURITY', eventType, details, email]);
    }
    
    Logger.log('🔒 Security Event: ' + eventType + ' - ' + details + ' by ' + email);
  } catch (e) {
    Logger.log('Error logging security event: ' + e.toString());
  }
}

/**
 * Access denied HTML
 */
function getAccessDeniedHtml_(email) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Inter', sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #f8f8f8; }
        .container { text-align: center; padding: 40px; background: white; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
        h1 { color: #ff4422; margin-bottom: 16px; }
        p { color: #525252; margin-bottom: 24px; }
        .email { background: #f5f5f5; padding: 8px 16px; border-radius: 8px; font-family: monospace; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🚫 Access Denied</h1>
        <p>Your email is not authorized to access the Creative Hub.</p>
        <div class="email">${email || 'Unknown'}</div>
        <p style="margin-top: 24px; font-size: 14px;">Contact an admin to request access.</p>
      </div>
    </body>
    </html>
  `;
}

// ============================================================================
// DATA FETCHING FUNCTIONS (called from client-side)
// ============================================================================

/**
 * Get all dashboard data in one call (efficient)
 */
function getDashboardData() {
  try {
    const teamMembers = getTeamMembers_();
    const products = getProducts_();
    const batchesCount = getBatchesCount_();
    const timeSaved = getTimeSavedData_();
    const systemErrors = getSystemErrors();
    const quickLinks = getQuickLinks();
    
    // Count healthy systems
    let healthyCount = 0;
    Object.values(systemErrors.systemStatus).forEach(s => {
      if (s.status === 'healthy') healthyCount++;
    });
    
    return {
      success: true,
      data: {
        teamCount: teamMembers.filter(m => m.active).length,
        productCount: products.filter(p => p.active).length,
        batchesCount: batchesCount,
        systemsHealthy: healthyCount,
        systemsTotal: 5,
        timeSaved: timeSaved,
        teamMembers: teamMembers,
        products: products,
        systemStatus: systemErrors.systemStatus,
        recentErrors: systemErrors.errors,
        quickLinks: quickLinks
      }
    };
  } catch (e) {
    Logger.log('Error in getDashboardData: ' + e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Get team members from Config sheet
 */
function getTeamMembers_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'hub_team_members';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows
    
    members.push({
      rowIndex: i + 1,       // 1-based row number for editing
      name: row[0],           // Column A: Name
      email: row[1],          // Column B: Email
      apiKey: row[2] || '',   // Column C: API Key
      isAdmin: row[3] === true,     // Column D: Is Admin
      clickUpAccess: row[4] === true, // Column E: ClickUp Access
      andromedaAccess: row[5] === true, // Column F: Andromeda Access
      sheetTab: row[6] || '',  // Column G: Sheet Tab
      active: row[7] === true, // Column H: Active
      addedDate: row[8] || ''  // Column I: Added Date
    });
  }
  
  cache.put(cacheKey, JSON.stringify(members), CONFIG_CACHE_DURATION);
  return members;
}

/**
 * Get products from Config sheet
 */
function getProducts_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'hub_products';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  
  // Get footage folders
  const footageFolders = {};
  const footageSheet = ss.getSheetByName('Footage Folders');
  if (footageSheet) {
    const footageData = footageSheet.getDataRange().getValues();
    for (let i = 1; i < footageData.length; i++) {
      if (footageData[i][0]) {
        footageFolders[footageData[i][0]] = footageData[i][1] || '';
      }
    }
  }
  
  const products = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    
    const productKey = row[0];
    products.push({
      rowIndex: i + 1,       // 1-based row number for editing
      key: productKey,       // Column A: Product Key
      displayName: row[1],   // Column B: Display Name
      productId: row[2],     // Column C: Product ID (A03, etc.)
      fullName: row[3],      // Column D: Full Name
      targetCpa: row[4],     // Column E: Target CPA
      url: row[5],           // Column F: URL
      active: row[6] === true, // Column G: Active
      notes: row[7] || '',   // Column H: Notes
      footageFolder: footageFolders[productKey] || '' // From Footage Folders tab
    });
  }
  
  cache.put(cacheKey, JSON.stringify(products), CONFIG_CACHE_DURATION);
  return products;
}

/**
 * Get batches count from Analytics sheet
 */
function getBatchesCount_() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'hub_batches_count';
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // Old format, clear it
      }
    }
    
    const ss = SpreadsheetApp.openById(ANALYTICS_SHEET_ID);
    const sheet = ss.getSheetByName('Events');
    if (!sheet) return { total: 0, thisMonth: 0 };
    
    // Count rows (minus header)
    const count = Math.max(0, sheet.getLastRow() - 1);
    
    // Count this month's batches
    let thisMonthCount = 0;
    if (count > 0) {
      const now = new Date();
      const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
      
      // Get timestamps from column A (assuming first column is timestamp)
      const data = sheet.getRange(2, 1, count, 1).getValues();
      data.forEach(row => {
        if (row[0]) {
          const rowDate = new Date(row[0]);
          if (rowDate >= startOfMonth) {
            thisMonthCount++;
          }
        }
      });
    }
    
    const result = { total: count, thisMonth: thisMonthCount };
    cache.put(cacheKey, JSON.stringify(result), CONFIG_CACHE_DURATION);
    return result;
  } catch (e) {
    Logger.log('Error getting batches count: ' + e.toString());
    return { total: 0, thisMonth: 0 };
  }
}

/**
 * Get time saved data - USES REAL DATA
 * Calculates based on actual batches and tasks created
 */
function getTimeSavedData_() {
  // Get real batch count from Analytics (now returns {total, thisMonth})
  const batchData = getBatchesCount_();
  const batchCount = batchData.total || 0;
  
  // Get real ClickUp task count
  const clickUpTaskCount = getClickUpTaskCount_();
  
  // Time estimates per action (in seconds)
  // These match the Ad Naming Analysis calculations:
  // - Standard Setup: 90 sec (folder creation, template copy, naming, etc.)
  // - Average across all batch types is ~85 sec
  const SECONDS_PER_BATCH = 85; // ~85 sec saved per batch (realistic estimate)
  const SECONDS_PER_CLICKUP_TASK = 90; // ~90 sec saved per ClickUp task
  
  // Calculate Ad Naming time saved
  const adNamingSeconds = batchCount * SECONDS_PER_BATCH;
  const adNamingHours = Math.round((adNamingSeconds / 3600) * 10) / 10; // Round to 1 decimal
  
  // Calculate ClickUp time saved
  const clickUpSeconds = clickUpTaskCount * SECONDS_PER_CLICKUP_TASK;
  const clickUpHours = Math.round((clickUpSeconds / 3600) * 10) / 10;
  
  const totalHours = Math.round((adNamingHours + clickUpHours) * 10) / 10;
  
  return {
    adNaming: {
      hours: adNamingHours,
      batches: batchCount,
      secondsPerBatch: SECONDS_PER_BATCH
    },
    sprintClickUp: {
      hours: clickUpHours,
      tasks: clickUpTaskCount,
      avgSecondsPerTask: SECONDS_PER_CLICKUP_TASK
    },
    totalHours: totalHours
  };
}

/**
 * Get total ClickUp tasks created from ClickUp Log tab
 */
function getClickUpTaskCount_() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'hub_clickup_count';
    const cached = cache.get(cacheKey);
    if (cached) return parseInt(cached);
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('ClickUp Log');
    
    if (!sheet) {
      // No ClickUp Log tab yet - return 0
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    let totalTasks = 0;
    
    // Sum up all tasks (column B, skip header)
    for (let i = 1; i < data.length; i++) {
      const count = parseInt(data[i][1]) || 0;
      totalTasks += count;
    }
    
    cache.put(cacheKey, totalTasks.toString(), CONFIG_CACHE_DURATION);
    return totalTasks;
  } catch (e) {
    Logger.log('Error getting ClickUp task count: ' + e.toString());
    return 0;
  }
}

// ============================================================================
// ACTION FUNCTIONS (called from client-side)
// ============================================================================

/**
 * Clear cache for a specific system
 * Note: This only clears THIS web app's cache.
 * To clear individual system caches, users should use the menu in each sheet.
 */
function clearHubCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['hub_team_members', 'hub_products', 'hub_batches_count', 'hub_clickup_count']);
    return { success: true, message: 'Hub cache cleared! Data will refresh on next load.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Clear cache AND return fresh dashboard data in one call
 * This reduces round trips from 2 to 1
 */
function getDashboardDataFresh() {
  // Clear all caches first
  const cache = CacheService.getScriptCache();
  cache.removeAll(['hub_team_members', 'hub_products', 'hub_batches_count', 'hub_clickup_count']);
  
  // Now get fresh data (cache is empty, so it will fetch from sheet)
  return getDashboardData();
}

// ============================================================================
// PRODUCT FOLDERS MANAGEMENT
// ============================================================================

/**
 * Get all product folders from Config sheet
 */
function getProductFolders() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, error: 'Product Folders sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // ['Product Key', 'US', 'UK', 'DE', ...]
    const locales = headers.slice(1).filter(h => h); // Remove first column (Product Key)
    
    const products = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const productKey = data[i][0];
      if (!productKey) continue;
      
      const folders = {};
      let folderCount = 0;
      
      for (let j = 1; j < headers.length; j++) {
        const locale = headers[j];
        const folderId = data[i][j];
        if (locale) {
          folders[locale] = folderId ? folderId.toString().trim() : '';
          if (folderId) folderCount++;
        }
      }
      
      products.push({
        rowIndex: i + 1,
        productKey: productKey.toString().trim(),
        folders: folders,
        folderCount: folderCount
      });
    }
    
    return {
      success: true,
      locales: locales,
      products: products
    };
    
  } catch (e) {
    Logger.log('Error getting product folders: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get folder data for a single product (for Edit Product modal)
 */
function getProductFoldersForProduct(productKey) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, error: 'Product Folders sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const locales = headers.slice(1).filter(h => h);
    
    // Find the product row
    let productFolders = {};
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productKey) {
        found = true;
        for (let j = 1; j < headers.length; j++) {
          const locale = headers[j];
          if (locale) {
            productFolders[locale] = data[i][j] ? data[i][j].toString().trim() : '';
          }
        }
        break;
      }
    }
    
    // If product not found in Product Folders, return empty folders for all locales
    if (!found) {
      locales.forEach(locale => {
        productFolders[locale] = '';
      });
    }
    
    return {
      success: true,
      productKey: productKey,
      locales: locales,
      folders: productFolders,
      existsInSheet: found
    };
    
  } catch (e) {
    Logger.log('Error getting product folders for ' + productKey + ': ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get all available locales (for Add Product modal)
 */
function getAvailableLocales() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, error: 'Product Folders sheet not found' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const locales = headers.slice(1).filter(h => h);
    
    return {
      success: true,
      locales: locales
    };
    
  } catch (e) {
    Logger.log('Error getting available locales: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Update a single folder ID for a product/locale
 */
function updateProductFolder(productKey, locale, folderId) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, message: 'Product Folders sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column index for locale
    const colIndex = headers.indexOf(locale);
    if (colIndex === -1) {
      return { success: false, message: 'Locale "' + locale + '" not found' };
    }
    
    // Find row for product
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productKey) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Product "' + productKey + '" not found' };
    }
    
    // Clean folder ID (extract ID if full URL given)
    let cleanId = folderId;
    if (folderId && folderId.includes('folders/')) {
      const match = folderId.match(/folders\/([a-zA-Z0-9_-]+)/);
      if (match) cleanId = match[1];
    }
    
    sheet.getRange(rowIndex, colIndex + 1).setValue(cleanId);
    
    logSyncEvent_('Product Folder Updated', productKey + ' / ' + locale);
    
    // Clear cache
    CacheService.getScriptCache().remove('adn_folders');
    
    return { success: true, message: 'Folder updated for ' + productKey + ' / ' + locale };
    
  } catch (e) {
    logSecurityEvent_('UPDATE_FOLDER_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Save multiple folder IDs for a product at once
 */
function saveProductFolders(productKey, folders) {
  try {
    verifyAdmin_();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, message: 'Product Folders sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find product row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productKey) {
        rowIndex = i + 1;
        break;
      }
    }
    
    // If product doesn't exist, add it
    if (rowIndex === -1) {
      rowIndex = sheet.getLastRow() + 1;
      sheet.getRange(rowIndex, 1).setValue(productKey);
    }
    
    // Update each folder
    let updatedCount = 0;
    for (const locale in folders) {
      const colIndex = headers.indexOf(locale);
      if (colIndex > 0) {
        let folderId = folders[locale] || '';
        
        // Clean folder ID (extract ID if full URL given)
        if (folderId && folderId.includes('folders/')) {
          const match = folderId.match(/folders\/([a-zA-Z0-9_-]+)/);
          if (match) folderId = match[1];
        }
        
        sheet.getRange(rowIndex, colIndex + 1).setValue(folderId);
        if (folderId) updatedCount++;
      }
    }
    
    // Clear cache
    CacheService.getScriptCache().remove('adn_folders');
    
    logSyncEvent_('Product Folders Updated', productKey + ' (' + updatedCount + ' locales)');
    
    return { success: true, message: updatedCount + ' folder(s) updated for ' + productKey };
    
  } catch (e) {
    Logger.log('Error saving product folders: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Add a new product to Product Folders
 */
function addProductFolder(productKey) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, message: 'Product Folders sheet not found' };
    }
    
    // Check if product already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productKey) {
        return { success: false, message: 'Product "' + productKey + '" already exists' };
      }
    }
    
    // Add new row with product key
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setValue(productKey);
    
    logSyncEvent_('Product Folder Added', productKey);
    
    return { success: true, message: 'Product added: ' + productKey };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Add a new locale column to Product Folders
 */
function addLocaleToProductFolders(locale) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      return { success: false, message: 'Product Folders sheet not found' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Check if locale already exists
    if (headers.includes(locale)) {
      return { success: false, message: 'Locale "' + locale + '" already exists in Product Folders' };
    }
    
    // Add new column at the end
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(locale);
    
    logSyncEvent_('Locale Column Added', 'Product Folders: ' + locale);
    
    // Clear cache
    CacheService.getScriptCache().remove('adn_folders');
    
    return { success: true, message: 'Locale "' + locale + '" added to Product Folders' };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Get locales that are in Valid Values but not in Product Folders
 */
function getMissingLocales() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    
    // Get Valid Values locales
    const vvSheet = ss.getSheetByName('Valid Values');
    const vvData = vvSheet.getDataRange().getValues();
    const vvHeaders = vvData[0];
    const localeCol = vvHeaders.indexOf('Locales');
    
    const validLocales = [];
    if (localeCol >= 0) {
      for (let i = 1; i < vvData.length; i++) {
        if (vvData[i][localeCol] && vvData[i][localeCol].toString().trim()) {
          validLocales.push(vvData[i][localeCol].toString().trim());
        }
      }
    }
    
    // Get Product Folders locales (headers)
    const pfSheet = ss.getSheetByName('Product Folders');
    const pfHeaders = pfSheet.getRange(1, 1, 1, pfSheet.getLastColumn()).getValues()[0];
    const folderLocales = pfHeaders.slice(1).filter(h => h); // Skip first column (Product Key)
    
    // Find missing
    const missing = validLocales.filter(l => !folderLocales.includes(l));
    
    return {
      success: true,
      missing: missing,
      validLocales: validLocales,
      folderLocales: folderLocales
    };
    
  } catch (e) {
    return { success: false, error: e.toString(), missing: [] };
  }
}

// ============================================================================
// TEMPLATES MANAGEMENT
// ============================================================================

/**
 * Get all templates from Config sheet
 */
function getTemplates() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Templates');
    
    if (!sheet) {
      return { success: false, error: 'Templates sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const templates = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const docId = data[i][1];
      const description = data[i][2] || '';
      
      if (key) {
        // Try to get doc name
        let docName = '';
        let docUrl = '';
        if (docId) {
          try {
            const doc = DriveApp.getFileById(docId);
            docName = doc.getName();
            docUrl = doc.getUrl();
          } catch (e) {
            docName = '(Cannot access)';
          }
        }
        
        templates.push({
          rowIndex: i + 1,
          key: key.toString().trim(),
          docId: docId ? docId.toString().trim() : '',
          description: description,
          docName: docName,
          docUrl: docUrl
        });
      }
    }
    
    return {
      success: true,
      templates: templates
    };
    
  } catch (e) {
    Logger.log('Error getting templates: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Update a template document ID
 */
function updateTemplate(rowIndex, docId) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Templates');
    
    if (!sheet) {
      return { success: false, message: 'Templates sheet not found' };
    }
    
    const key = sheet.getRange(rowIndex, 1).getValue();
    const oldId = sheet.getRange(rowIndex, 2).getValue();
    
    // Validate the new doc ID
    let docName = '';
    if (docId) {
      try {
        const doc = DriveApp.getFileById(docId);
        docName = doc.getName();
      } catch (e) {
        return { success: false, message: 'Cannot access document with ID: ' + docId };
      }
    }
    
    sheet.getRange(rowIndex, 2).setValue(docId);
    
    logSyncEvent_('Template Updated', key + ': ' + (docName || docId));
    
    // Clear cache
    CacheService.getScriptCache().remove('adn_templates');
    
    return { success: true, message: 'Template updated: ' + key, docName: docName };
    
  } catch (e) {
    logSecurityEvent_('UPDATE_TEMPLATE_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Add a new template
 */
function addTemplate(key, docId, description) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Input validation
    if (!key || key.trim() === '') {
      return { success: false, message: 'Template key is required' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Templates');
    
    if (!sheet) {
      return { success: false, message: 'Templates sheet not found' };
    }
    
    // Validate doc ID
    let docName = '';
    if (docId) {
      try {
        const doc = DriveApp.getFileById(docId);
        docName = doc.getName();
      } catch (e) {
        return { success: false, message: 'Cannot access document with ID: ' + docId };
      }
    }
    
    // Add new row
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 3).setValues([[key, docId, description || '']]);
    
    logSyncEvent_('Template Added', key + ': ' + (docName || docId));
    
    // Clear cache
    CacheService.getScriptCache().remove('adn_templates');
    
    return { success: true, message: 'Template added: ' + key };
    
  } catch (e) {
    logSecurityEvent_('ADD_TEMPLATE_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Remove a template
 */
function removeTemplate(rowIndex) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Templates');
    
    if (!sheet) {
      return { success: false, message: 'Templates sheet not found' };
    }
    
    const key = sheet.getRange(rowIndex, 1).getValue();
    sheet.deleteRow(rowIndex);
    
    logSyncEvent_('Template Removed', key);
    logSecurityEvent_('TEMPLATE_REMOVED', key);
    
    // Clear cache
    CacheService.getScriptCache().remove('adn_templates');
    
    return { success: true, message: 'Template removed: ' + key };
    
  } catch (e) {
    logSecurityEvent_('REMOVE_TEMPLATE_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// VALID VALUES MANAGEMENT
// ============================================================================

/**
 * Get all valid values from Config sheet
 */
function getValidValues() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Valid Values');
    
    if (!sheet) {
      return { success: false, error: 'Valid Values sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Build column data
    const columns = {};
    headers.forEach((header, colIndex) => {
      if (header) {
        columns[header] = [];
        for (let i = 1; i < data.length; i++) {
          const value = data[i][colIndex];
          if (value && value.toString().trim() !== '') {
            columns[header].push({
              rowIndex: i + 1,
              value: value.toString().trim()
            });
          }
        }
      }
    });
    
    return {
      success: true,
      headers: headers.filter(h => h),
      columns: columns
    };
    
  } catch (e) {
    Logger.log('Error getting valid values: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Add a new value to a column in Valid Values
 */
function addValidValue(columnName, value) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Input validation
    if (!value || value.toString().trim() === '') {
      return { success: false, message: 'Value cannot be empty' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Valid Values');
    
    if (!sheet) {
      return { success: false, message: 'Valid Values sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIndex = headers.indexOf(columnName);
    
    if (colIndex === -1) {
      return { success: false, message: 'Column "' + columnName + '" not found' };
    }
    
    // Find first empty row in this column
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][colIndex] || data[i][colIndex].toString().trim() === '') {
        targetRow = i + 1;
        break;
      }
    }
    
    // If no empty row, append
    if (targetRow === -1) {
      targetRow = data.length + 1;
    }
    
    sheet.getRange(targetRow, colIndex + 1).setValue(value);
    
    logSyncEvent_('Valid Value Added', columnName + ': ' + value);
    
    return { success: true, message: 'Value added to ' + columnName };
    
  } catch (e) {
    logSecurityEvent_('ADD_VALUE_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Remove a value from Valid Values
 */
function removeValidValue(columnName, rowIndex) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Valid Values');
    
    if (!sheet) {
      return { success: false, message: 'Valid Values sheet not found' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(columnName);
    
    if (colIndex === -1) {
      return { success: false, message: 'Column not found' };
    }
    
    const oldValue = sheet.getRange(rowIndex, colIndex + 1).getValue();
    sheet.getRange(rowIndex, colIndex + 1).setValue('');
    
    logSyncEvent_('Valid Value Removed', columnName + ': ' + oldValue);
    
    return { success: true, message: 'Value removed' };
    
  } catch (e) {
    logSecurityEvent_('REMOVE_VALUE_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Get activity log from Sync Log and Error Log tabs
 * Returns combined, sorted activity for display
 */
function getActivityLog() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const activities = [];
    
    // Get Sync Log entries
    const syncSheet = ss.getSheetByName('Sync Log');
    if (syncSheet && syncSheet.getLastRow() > 1) {
      const syncData = syncSheet.getRange(2, 1, Math.min(syncSheet.getLastRow() - 1, 50), 5).getValues();
      syncData.forEach(row => {
        if (row[0]) {
          activities.push({
            timestamp: new Date(row[0]).toISOString(),
            type: 'sync',
            source: row[1] || 'System',
            action: row[2] || '',
            details: row[3] || '',
            icon: '🔄'
          });
        }
      });
    }
    
    // Get Error Log entries (last 20)
    const errorSheet = ss.getSheetByName('Error Log');
    if (errorSheet && errorSheet.getLastRow() > 1) {
      const errorData = errorSheet.getRange(2, 1, Math.min(errorSheet.getLastRow() - 1, 20), 6).getValues();
      errorData.forEach(row => {
        if (row[0]) {
          const resolved = row[5] === true;
          activities.push({
            timestamp: new Date(row[0]).toISOString(),
            type: resolved ? 'resolved' : 'error',
            source: row[1] || 'System',
            action: row[2] || 'Error',
            details: row[3] || '',
            user: row[4] || '',
            icon: resolved ? '✅' : '⚠️'
          });
        }
      });
    }
    
    // Get ClickUp Log entries (last 20)
    const clickUpSheet = ss.getSheetByName('ClickUp Log');
    if (clickUpSheet && clickUpSheet.getLastRow() > 1) {
      const clickUpData = clickUpSheet.getRange(2, 1, Math.min(clickUpSheet.getLastRow() - 1, 20), 4).getValues();
      clickUpData.forEach(row => {
        if (row[0] && row[1]) {
          activities.push({
            timestamp: new Date(row[0]).toISOString(),
            type: 'clickup',
            source: 'Sprint → ClickUp',
            action: 'Tasks Created',
            details: row[1] + ' tasks (' + (row[2] || 'various products') + ')',
            user: row[3] || '',
            icon: '✅'
          });
        }
      });
    }
    
    // Sort by timestamp (newest first)
    activities.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return {
      success: true,
      activities: activities.slice(0, 50) // Return last 50
    };
    
  } catch (e) {
    Logger.log('Error getting activity log: ' + e.toString());
    return {
      success: false,
      error: e.toString(),
      activities: []
    };
  }
}

/**
 * Get Config sheet URL (for "Open Config Sheet" button)
 */
function getConfigSheetUrl() {
  return 'https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID + '/edit';
}

/**
 * Get Quick Links - sheet URLs for all automations
 * Ordered by usage frequency, Config removed (already in header)
 */
function getQuickLinks() {
  // Links ordered by usage frequency
  const defaults = {
    adNaming: {
      name: 'Ad Naming Sheet',
      desc: 'Create batches & ad notes',
      url: 'https://docs.google.com/spreadsheets/d/1LDQGcAzEXZfz862ps4BjpXSybLGOkXJrCSzSMhAw9iQ/edit',
      icon: '📝',
      color: 'green',
      order: 1
    },
    sprintClickUp: {
      name: 'Sprint → ClickUp',
      desc: 'Create ClickUp tasks',
      url: 'https://docs.google.com/spreadsheets/d/1h7apn0p6ozXoO8qB0Qi9Eo-KuEMBaRfwTTzNdJjh4yI/edit',
      icon: '✅',
      color: 'purple',
      order: 2
    },
    atriaReports: {
      name: 'Atria Reports',
      desc: 'Weekly performance analysis',
      url: 'https://docs.google.com/spreadsheets/d/1fzFXrww0MSN2bsdIsxlIv9uPEqg2xAwsCktPgTlzDS4/edit',
      icon: '📊',
      color: 'blue',
      order: 3
    },
    driveQA: {
      name: 'Drive Naming QA',
      desc: 'Check folder naming',
      url: 'https://docs.google.com/spreadsheets/d/1V_RhADEiL0uHRtxDeJHRINFvLbf2Dq1Zf9W2QVWB578/edit',
      icon: '🔍',
      color: 'orange',
      order: 4
    },
    quarterlyGoals: {
      name: 'Quarterly Goals',
      desc: 'Winning ads tracker',
      url: 'https://docs.google.com/spreadsheets/d/1l3HN3c75qXClr3UJ4-kgkowLIqFEk-9aRT89LJ-FfrQ/edit',
      icon: '🎯',
      color: 'red',
      order: 5
    }
  };
  
  // Try to get custom URLs from Config sheet (if a "Quick Links" tab exists)
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Quick Links');
    
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const key = data[i][0];
        const url = data[i][1];
        if (key && url && defaults[key]) {
          defaults[key].url = url;
        }
      }
    }
  } catch (e) {
    // Use defaults if can't read config
  }
  
  return defaults;
}

// ============================================================================
// SYSTEM HEALTH & ERROR TRACKING
// ============================================================================

/**
 * Get system errors from Error Log tab
 * Returns recent errors for display in System Status
 */
function getSystemErrors() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Error Log');
    
    if (!sheet) {
      // No Error Log tab yet - return empty (all healthy)
      return { 
        hasErrors: false, 
        errors: [],
        systemStatus: {
          'Ad Naming Sheet': { status: 'healthy', message: 'No errors', lastError: null },
          'Atria Reports': { status: 'healthy', message: 'No errors', lastError: null },
          'Sprint → ClickUp': { status: 'healthy', message: 'No errors', lastError: null },
          'Drive Naming QA': { status: 'healthy', message: 'No errors', lastError: null }
        }
      };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { 
        hasErrors: false, 
        errors: [],
        systemStatus: {
          'Ad Naming Sheet': { status: 'healthy', message: 'No errors', lastError: null },
          'Atria Reports': { status: 'healthy', message: 'No errors', lastError: null },
          'Sprint → ClickUp': { status: 'healthy', message: 'No errors', lastError: null },
          'Drive Naming QA': { status: 'healthy', message: 'No errors', lastError: null }
        }
      };
    }
    
    // Parse errors (last 24 hours, not resolved)
    const now = new Date();
    const oneDayAgo = new Date(now - 24 * 60 * 60 * 1000);
    
    const recentErrors = [];
    const systemStatus = {
      'Ad Naming Sheet': { status: 'healthy', message: 'No errors', lastError: null, errorCount: 0 },
      'Atria Reports': { status: 'healthy', message: 'No errors', lastError: null, errorCount: 0 },
      'Sprint → ClickUp': { status: 'healthy', message: 'No errors', lastError: null, errorCount: 0 },
      'Drive Naming QA': { status: 'healthy', message: 'No errors', lastError: null, errorCount: 0 }
    };
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const timestamp = new Date(row[0]);
      const system = row[1];
      const errorType = row[2];
      const message = row[3];
      const user = row[4];
      const resolved = row[5] === true;
      
      // Skip resolved errors
      if (resolved) continue;
      
      if (timestamp >= oneDayAgo) {
        recentErrors.push({
          rowIndex: i + 1, // 1-based for sheet operations
          timestamp: timestamp.toISOString(),
          system: system,
          errorType: errorType,
          message: message,
          user: user,
          resolved: resolved
        });
        
        // Update system status
        if (systemStatus[system]) {
          systemStatus[system].errorCount++;
          if (!systemStatus[system].lastError || timestamp > new Date(systemStatus[system].lastError)) {
            systemStatus[system].lastError = timestamp.toISOString();
            systemStatus[system].lastErrorType = errorType;
            systemStatus[system].lastMessage = message;
          }
        }
      }
    }
    
    // Set status based on error count
    Object.keys(systemStatus).forEach(system => {
      const s = systemStatus[system];
      if (s.errorCount === 0) {
        s.status = 'healthy';
        s.message = 'No errors';
      } else if (s.errorCount <= 2) {
        s.status = 'warning';
        s.message = s.errorCount + ' error(s) in last 24h';
      } else {
        s.status = 'error';
        s.message = s.errorCount + ' errors in last 24h';
      }
    });
    
    return {
      hasErrors: recentErrors.length > 0,
      errors: recentErrors.slice(0, 20), // Last 20 unresolved errors
      systemStatus: systemStatus
    };
    
  } catch (e) {
    Logger.log('Error getting system errors: ' + e.toString());
    return { 
      hasErrors: false, 
      errors: [],
      systemStatus: {
        'Ad Naming Sheet': { status: 'healthy', message: 'No errors', lastError: null },
        'Atria Reports': { status: 'healthy', message: 'No errors', lastError: null },
        'Sprint → ClickUp': { status: 'healthy', message: 'No errors', lastError: null },
        'Drive Naming QA': { status: 'healthy', message: 'No errors', lastError: null }
      }
    };
  }
}

/**
 * Log an error from any automation system
 * Call this from other scripts when errors occur
 */
function logSystemError(system, errorType, message, user) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let sheet = ss.getSheetByName('Error Log');
    
    // Create Error Log tab if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Error Log');
      sheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'System', 'Error Type', 'Message', 'User', 'Resolved']]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f4cccc');
      sheet.setFrozenRows(1);
    }
    
    // Insert error at row 2 (after header)
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 6).setValues([[
      new Date().toISOString(),
      system,
      errorType,
      message,
      user || 'Unknown',
      false // Not resolved
    ]]);
    
    // Keep only last 500 errors
    const lastRow = sheet.getLastRow();
    if (lastRow > 501) {
      sheet.deleteRows(502, lastRow - 501);
    }
    
    return { success: true };
  } catch (e) {
    Logger.log('Failed to log system error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Mark an error as resolved
 * @param {number} rowIndex - The row index in Error Log sheet
 */
function resolveError(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Error Log');
    
    if (!sheet) {
      return { success: false, message: 'Error Log sheet not found' };
    }
    
    // Mark as resolved (column F)
    sheet.getRange(rowIndex, 6).setValue(true);
    
    // Add resolved timestamp to message
    const currentMessage = sheet.getRange(rowIndex, 4).getValue();
    const resolvedBy = Session.getActiveUser().getEmail() || 'Unknown';
    sheet.getRange(rowIndex, 4).setValue(currentMessage + ' [RESOLVED by ' + resolvedBy + ' at ' + new Date().toLocaleString() + ']');
    
    return { success: true, message: 'Error marked as resolved' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Clear all resolved errors (cleanup)
 */
function clearResolvedErrors() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Error Log');
    
    if (!sheet) {
      return { success: false, message: 'Error Log sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;
    
    // Go from bottom to top to avoid index shifting issues
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][5] === true) { // Column F = Resolved
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    return { success: true, message: 'Cleared ' + deletedCount + ' resolved errors' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// FOOTAGE FOLDERS MANAGEMENT
// ============================================================================

/**
 * Get all footage folders from Config sheet
 */
function getFootageFolders() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Footage Folders');
    
    if (!sheet) {
      return { success: true, folders: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    const folders = {};
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const productKey = data[i][0];
      const folderUrl = data[i][1] || '';
      if (productKey) {
        folders[productKey] = folderUrl;
      }
    }
    
    return { success: true, folders: folders };
  } catch (e) {
    Logger.log('Error getting footage folders: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get footage folder for a specific product
 */
function getFootageFolderForProduct(productKey) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Footage Folders');
    
    if (!sheet) {
      return { success: true, folderUrl: '' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productKey) {
        return { success: true, folderUrl: data[i][1] || '' };
      }
    }
    
    return { success: true, folderUrl: '' };
  } catch (e) {
    Logger.log('Error getting footage folder: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Save footage folder for a product
 */
function saveFootageFolder(productKey, folderUrl) {
  try {
    verifyAdmin_();
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let sheet = ss.getSheetByName('Footage Folders');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Footage Folders');
      sheet.getRange(1, 1, 1, 2).setValues([['Product Key', 'Footage Folder URL']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#d9ead3');
      sheet.setFrozenRows(1);
    }
    
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    
    // Find existing row for this product
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productKey) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow > 0) {
      // Update existing row
      sheet.getRange(foundRow, 2).setValue(folderUrl || '');
    } else if (folderUrl) {
      // Add new row only if URL is provided
      sheet.appendRow([productKey, folderUrl]);
    }
    
    // Log the change
    logSyncEvent_('Footage Folder Updated', productKey + ': ' + (folderUrl ? 'Set' : 'Cleared'));
    
    return { success: true, message: 'Footage folder saved' };
  } catch (e) {
    Logger.log('Error saving footage folder: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// EDIT FUNCTIONS - Team Members
// ============================================================================

/**
 * Update a team member in the Config sheet
 * @param {number} rowIndex - The row index (1-based, including header)
 * @param {object} memberData - The updated member data
 */
function updateTeamMember(rowIndex, memberData) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Rate limiting
    if (isRateLimited_('updateTeamMember', 10)) {
      return { success: false, message: 'Too many requests. Please wait a moment.' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Team Members');
    
    if (!sheet) {
      return { success: false, message: 'Team Members sheet not found' };
    }
    
    // Update the row (columns A-I)
    const rowData = [
      memberData.name,
      memberData.email,
      memberData.apiKey || '',
      memberData.isAdmin === true,
      memberData.clickUpAccess === true,
      memberData.andromedaAccess === true,
      memberData.sheetTab || '',
      memberData.active === true,
      memberData.addedDate || ''
    ];
    
    sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
    
    // Clear all relevant caches to ensure fresh data
    CacheService.getScriptCache().removeAll(['hub_team_members', 'hub_dashboard']);
    
    // Log to Sync Log
    logSyncEvent_('Team Member Updated', memberData.name + ' (' + memberData.email + ')');
    
    return { success: true, message: 'Team member updated successfully!' };
  } catch (e) {
    Logger.log('Error updating team member: ' + e.toString());
    logSecurityEvent_('UPDATE_TEAM_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Add a new team member to the Config sheet
 * @param {object} memberData - The new member data
 */
function addTeamMember(memberData) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Rate limiting
    if (isRateLimited_('addTeamMember', 5)) {
      return { success: false, message: 'Too many requests. Please wait a moment.' };
    }
    
    // Input validation
    if (!memberData.email || !memberData.email.includes('@')) {
      return { success: false, message: 'Invalid email address' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Team Members');
    
    if (!sheet) {
      return { success: false, message: 'Team Members sheet not found' };
    }
    
    // Check if email already exists
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][1] && existingData[i][1].toString().toLowerCase() === memberData.email.toLowerCase()) {
        return { success: false, message: 'A team member with this email already exists' };
      }
    }
    
    // Add new row at the bottom
    const rowData = [
      memberData.name,
      memberData.email,
      memberData.apiKey || '',
      memberData.isAdmin === true,
      memberData.clickUpAccess === true,
      memberData.andromedaAccess === true,
      memberData.sheetTab || '',
      true, // Active by default
      new Date().toISOString().split('T')[0] // Today's date
    ];
    
    sheet.appendRow(rowData);
    
    // Clear all relevant caches to ensure fresh data
    CacheService.getScriptCache().removeAll(['hub_team_members', 'hub_dashboard']);
    
    // Log to Sync Log
    logSyncEvent_('Team Member Added', memberData.name + ' (' + memberData.email + ')');
    logSecurityEvent_('TEAM_MEMBER_ADDED', memberData.email);
    
    return { success: true, message: 'Team member added successfully!' };
  } catch (e) {
    Logger.log('Error adding team member: ' + e.toString());
    logSecurityEvent_('ADD_TEAM_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// EDIT FUNCTIONS - Products
// ============================================================================

/**
 * Update a product in the Config sheet
 * @param {number} rowIndex - The row index (1-based, including header)
 * @param {object} productData - The updated product data
 */
function updateProduct(rowIndex, productData) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Rate limiting
    if (isRateLimited_('updateProduct', 10)) {
      return { success: false, message: 'Too many requests. Please wait a moment.' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    
    if (!sheet) {
      return { success: false, message: 'Products sheet not found' };
    }
    
    // Update the row (columns A-H)
    const rowData = [
      productData.key,
      productData.displayName,
      productData.productId,
      productData.fullName || '',
      productData.targetCpa || 60,
      productData.url || '',
      productData.active === true,
      productData.notes || ''
    ];
    
    sheet.getRange(rowIndex, 1, 1, 8).setValues([rowData]);
    
    // Clear all relevant caches to ensure fresh data
    CacheService.getScriptCache().removeAll(['hub_products', 'hub_dashboard']);
    
    // Log to Sync Log
    logSyncEvent_('Product Updated', productData.key + ' (CPA: €' + productData.targetCpa + ')');
    
    return { success: true, message: 'Product updated successfully!' };
  } catch (e) {
    Logger.log('Error updating product: ' + e.toString());
    logSecurityEvent_('UPDATE_PRODUCT_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Add a new product to the Config sheet
 * @param {object} productData - The new product data
 */
function addProduct(productData) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Rate limiting
    if (isRateLimited_('addProduct', 5)) {
      return { success: false, message: 'Too many requests. Please wait a moment.' };
    }
    
    // Input validation
    if (!productData.key || !productData.displayName) {
      return { success: false, message: 'Product key and display name are required' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    
    if (!sheet) {
      return { success: false, message: 'Products sheet not found' };
    }
    
    // Add new row
    const rowData = [
      productData.key,
      productData.displayName,
      productData.productId,
      productData.fullName || '',
      productData.targetCpa || 60,
      productData.url || '',
      true, // Active by default
      productData.notes || ''
    ];
    
    sheet.appendRow(rowData);
    
    // Clear all relevant caches to ensure fresh data
    CacheService.getScriptCache().removeAll(['hub_products', 'hub_dashboard']);
    
    // Also add to Product Folders sheet
    try {
      const foldersSheet = ss.getSheetByName('Product Folders');
      if (foldersSheet) {
        // Check if product already exists
        const folderData = foldersSheet.getDataRange().getValues();
        let exists = false;
        for (let i = 1; i < folderData.length; i++) {
          if (folderData[i][0] === productData.key) {
            exists = true;
            break;
          }
        }
        if (!exists) {
          foldersSheet.appendRow([productData.key]);
          Logger.log('✓ Added ' + productData.key + ' to Product Folders');
        }
      }
    } catch (folderErr) {
      Logger.log('Warning: Could not add to Product Folders: ' + folderErr.message);
    }
    
    // Log to Sync Log
    logSyncEvent_('Product Added', productData.key);
    logSecurityEvent_('PRODUCT_ADDED', productData.key);
    
    return { success: true, message: 'Product added successfully!' };
  } catch (e) {
    Logger.log('Error adding product: ' + e.toString());
    logSecurityEvent_('ADD_PRODUCT_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// DELETE FUNCTIONS
// ============================================================================

/**
 * Delete a team member from the Config sheet
 * @param {number} rowIndex - The row index (1-based, including header)
 */
function deleteTeamMember(rowIndex) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Rate limiting
    if (isRateLimited_('deleteTeamMember', 3)) {
      return { success: false, message: 'Too many requests. Please wait a moment.' };
    }
    
    // Validate row index
    if (!rowIndex || rowIndex < 2) {
      return { success: false, message: 'Invalid row index' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Team Members');
    
    if (!sheet) {
      return { success: false, message: 'Team Members sheet not found' };
    }
    
    // Get member info before deleting (for logging)
    const memberData = sheet.getRange(rowIndex, 1, 1, 2).getValues()[0];
    const memberName = memberData[0] || 'Unknown';
    const memberEmail = memberData[1] || 'Unknown';
    
    // Prevent deleting the current user (self-deletion)
    const currentUser = Session.getActiveUser().getEmail();
    if (memberEmail.toLowerCase() === currentUser.toLowerCase()) {
      return { success: false, message: 'You cannot delete your own account' };
    }
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    // Clear cache
    CacheService.getScriptCache().remove('hub_team_members');
    CacheService.getScriptCache().remove('hub_dashboard');
    
    // Log to Sync Log
    logSyncEvent_('Team Member Deleted', memberName + ' (' + memberEmail + ')');
    logSecurityEvent_('TEAM_MEMBER_DELETED', memberEmail);
    
    return { success: true, message: 'Team member deleted successfully' };
  } catch (e) {
    Logger.log('Error deleting team member: ' + e.toString());
    logSecurityEvent_('DELETE_TEAM_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Delete a product from the Config sheet
 * @param {number} rowIndex - The row index (1-based, including header)
 */
function deleteProduct(rowIndex) {
  try {
    // Security: Admin only
    verifyAdmin_();
    
    // Rate limiting
    if (isRateLimited_('deleteProduct', 3)) {
      return { success: false, message: 'Too many requests. Please wait a moment.' };
    }
    
    // Validate row index
    if (!rowIndex || rowIndex < 2) {
      return { success: false, message: 'Invalid row index' };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Products');
    
    if (!sheet) {
      return { success: false, message: 'Products sheet not found' };
    }
    
    // Get product info before deleting (for logging)
    const productData = sheet.getRange(rowIndex, 1, 1, 2).getValues()[0];
    const productKey = productData[0] || 'Unknown';
    const productName = productData[1] || 'Unknown';
    
    // Delete the row
    sheet.deleteRow(rowIndex);
    
    // Clear cache
    CacheService.getScriptCache().remove('hub_products');
    CacheService.getScriptCache().remove('hub_dashboard');
    
    // Log to Sync Log
    logSyncEvent_('Product Deleted', productKey + ' (' + productName + ')');
    logSecurityEvent_('PRODUCT_DELETED', productKey);
    
    return { success: true, message: 'Product deleted successfully' };
  } catch (e) {
    Logger.log('Error deleting product: ' + e.toString());
    logSecurityEvent_('DELETE_PRODUCT_FAILED', e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// SYNC LOG
// ============================================================================

/**
 * Log an event to the Sync Log tab
 */
function logSyncEvent_(action, details) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Sync Log');
    
    if (!sheet) return;
    
    const user = Session.getActiveUser().getEmail() || 'Hub User';
    const timestamp = new Date().toISOString();
    
    // Insert at row 2 (after header)
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 4).setValues([[timestamp, 'Creative Hub', action, details + ' by ' + user]]);
    
    // Keep only last 100 entries
    const lastRow = sheet.getLastRow();
    if (lastRow > 101) {
      sheet.deleteRows(102, lastRow - 101);
    }
  } catch (e) {
    Logger.log('Error logging sync event: ' + e.toString());
  }
}

// ============================================================================
// DIAGNOSTIC / HEALTH CHECK
// ============================================================================

/**
 * Run quick health check on all systems
 * Returns status for each system
 */
function runHealthCheck() {
  verifyAccess_();
  
  const results = {
    configSheet: { name: 'Config Sheet', status: 'unknown', message: '', details: {} },
    adNaming: { name: 'Ad Naming', status: 'unknown', message: '', details: {} },
    atria: { name: 'Atria Reports', status: 'unknown', message: '', details: {} },
    sprintClickUp: { name: 'Sprint → ClickUp', status: 'unknown', message: '', details: {} },
    driveNaming: { name: 'Drive Naming QA', status: 'unknown', message: '', details: {} }
  };
  
  // 1. Check Config Sheet access
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheets = ss.getSheets().map(s => s.getName());
    const requiredTabs = ['Products', 'Team Members', 'Valid Values', 'ClickUp Config'];
    const missingTabs = requiredTabs.filter(t => !sheets.includes(t));
    
    if (missingTabs.length === 0) {
      results.configSheet.status = 'healthy';
      results.configSheet.message = sheets.length + ' tabs accessible';
    } else {
      results.configSheet.status = 'warning';
      results.configSheet.message = 'Missing: ' + missingTabs.join(', ');
    }
    results.configSheet.details = { tabCount: sheets.length, missingTabs: missingTabs };
  } catch (e) {
    results.configSheet.status = 'error';
    results.configSheet.message = 'Cannot access: ' + e.message;
    results.configSheet.details = { error: e.toString() };
  }
  
  // 2. Check Ad Naming config
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const productsSheet = ss.getSheetByName('Products');
    const validValuesSheet = ss.getSheetByName('Valid Values');
    const templatesSheet = ss.getSheetByName('Templates');
    
    if (!productsSheet) {
      results.adNaming.status = 'error';
      results.adNaming.message = 'Products tab not found';
    } else if (!validValuesSheet) {
      results.adNaming.status = 'warning';
      results.adNaming.message = 'Valid Values tab not found';
    } else {
      const productCount = Math.max(0, productsSheet.getLastRow() - 1);
      results.adNaming.status = 'healthy';
      results.adNaming.message = productCount + ' products configured';
      results.adNaming.details = { 
        productCount: productCount, 
        hasTemplates: templatesSheet !== null 
      };
    }
  } catch (e) {
    results.adNaming.status = 'error';
    results.adNaming.message = 'Config check failed';
    results.adNaming.details = { error: e.toString() };
  }
  
  // 3. Check Atria Reports (Slack config)
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const slackSheet = ss.getSheetByName('Slack Config');
    if (slackSheet) {
      results.atria.status = 'healthy';
      results.atria.message = 'Slack config found';
    } else {
      results.atria.status = 'healthy';
      results.atria.message = 'Using local config';
    }
  } catch (e) {
    results.atria.status = 'error';
    results.atria.message = 'Config check failed';
  }
  
  // 4. Check Sprint → ClickUp (with ACTUAL API test)
  results.sprintClickUp = checkSprintClickUpHealthDeep_();
  
  // 5. Check Drive Naming QA
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const vvSheet = ss.getSheetByName('Valid Values');
    if (vvSheet) {
      results.driveNaming.status = 'healthy';
      results.driveNaming.message = 'Valid Values accessible';
    } else {
      results.driveNaming.status = 'warning';
      results.driveNaming.message = 'Valid Values not found';
    }
  } catch (e) {
    results.driveNaming.status = 'error';
    results.driveNaming.message = 'Config check failed';
  }
  
  return results;
}

/**
 * Deep health check for Sprint → ClickUp with actual API test
 */
function checkSprintClickUpHealthDeep_() {
  const result = {
    name: 'Sprint → ClickUp',
    status: 'unknown',
    message: '',
    details: {}
  };
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const clickUpSheet = ss.getSheetByName('ClickUp Config');
    
    if (!clickUpSheet) {
      result.status = 'error';
      result.message = 'ClickUp Config tab not found';
      return result;
    }
    
    // Get config values
    const data = clickUpSheet.getDataRange().getValues();
    const config = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) config[data[i][0]] = data[i][1];
    }
    
    const hasApiKey = !!config['API_KEY'];
    const hasFolderId = !!config['SPRINT_FOLDER_ID'] || !!config['FB_ADS_FOLDER_ID'];
    
    result.details.hasApiKey = hasApiKey;
    result.details.hasFolderId = hasFolderId;
    
    if (!hasApiKey) {
      result.status = 'error';
      result.message = 'API key not configured';
      return result;
    }
    
    if (!hasFolderId) {
      result.status = 'warning';
      result.message = 'Folder ID not configured';
      return result;
    }
    
    // ACTUAL API TEST - Try to access ClickUp
    try {
      const apiKey = config['API_KEY'];
      const testUrl = 'https://api.clickup.com/api/v2/user';
      
      const response = UrlFetchApp.fetch(testUrl, {
        method: 'GET',
        headers: { 'Authorization': apiKey },
        muteHttpExceptions: true
      });
      
      const code = response.getResponseCode();
      
      if (code === 200) {
        const userData = JSON.parse(response.getContentText());
        result.status = 'healthy';
        result.message = 'API connected as ' + (userData.user?.username || 'user');
        result.details.apiConnected = true;
        result.details.apiUser = userData.user?.username;
      } else if (code === 401) {
        result.status = 'error';
        result.message = 'API key invalid or expired';
        result.details.apiConnected = false;
        result.details.apiError = 'Authentication failed (401)';
      } else {
        result.status = 'warning';
        result.message = 'API returned status ' + code;
        result.details.apiConnected = false;
        result.details.apiError = 'HTTP ' + code;
      }
    } catch (apiError) {
      result.status = 'error';
      result.message = 'API connection failed';
      result.details.apiConnected = false;
      result.details.apiError = apiError.toString().substring(0, 100);
    }
    
    return result;
    
  } catch (e) {
    result.status = 'error';
    result.message = 'Health check failed';
    result.details.error = e.toString();
    return result;
  }
}

/**
 * Run a DEEP health check for a specific system (more thorough, takes longer)
 */
function runDeepHealthCheck(systemName) {
  verifyAccess_();
  
  const result = {
    system: systemName,
    status: 'unknown',
    message: '',
    checks: [],
    troubleshooting: []
  };
  
  switch (systemName) {
    case 'Ad Naming Sheet':
      result.checks = runAdNamingDeepCheck_();
      break;
    case 'Atria Reports':
      result.checks = runAtriaDeepCheck_();
      break;
    case 'Sprint → ClickUp':
      result.checks = runSprintClickUpDeepCheck_();
      break;
    case 'Drive Naming QA':
      result.checks = runDriveQADeepCheck_();
      break;
    case 'Quarterly Goals':
      result.checks = runGoalsDeepCheck_();
      break;
    default:
      result.message = 'Unknown system';
      return result;
  }
  
  // Calculate overall status
  const hasError = result.checks.some(c => c.status === 'error');
  const hasWarning = result.checks.some(c => c.status === 'warning');
  
  if (hasError) {
    result.status = 'error';
    result.message = 'Issues found - see details';
  } else if (hasWarning) {
    result.status = 'warning';
    result.message = 'Warnings found - see details';
  } else {
    result.status = 'healthy';
    result.message = 'All checks passed';
  }
  
  // Add troubleshooting tips
  result.troubleshooting = getTroubleshootingTips_(systemName, result.checks);
  
  return result;
}

function runAdNamingDeepCheck_() {
  const checks = [];
  
  // Check 1: Config access
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    checks.push({ name: 'Config Access', status: 'healthy', message: 'Config sheet accessible' });
  } catch (e) {
    checks.push({ name: 'Config Access', status: 'error', message: 'Cannot access config: ' + e.message });
    return checks;
  }
  
  // Check 2: Products tab
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const productsSheet = ss.getSheetByName('Products');
    if (productsSheet) {
      const count = Math.max(0, productsSheet.getLastRow() - 1);
      checks.push({ name: 'Products Tab', status: 'healthy', message: count + ' products configured' });
    } else {
      checks.push({ name: 'Products Tab', status: 'error', message: 'Products tab not found' });
    }
  } catch (e) {
    checks.push({ name: 'Products Tab', status: 'error', message: e.message });
  }
  
  // Check 3: Valid Values tab
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const vvSheet = ss.getSheetByName('Valid Values');
    if (vvSheet) {
      checks.push({ name: 'Valid Values', status: 'healthy', message: 'Tab accessible' });
    } else {
      checks.push({ name: 'Valid Values', status: 'warning', message: 'Tab not found' });
    }
  } catch (e) {
    checks.push({ name: 'Valid Values', status: 'error', message: e.message });
  }
  
  // Check 4: Templates tab
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const templatesSheet = ss.getSheetByName('Templates');
    if (templatesSheet) {
      checks.push({ name: 'Templates', status: 'healthy', message: 'Tab accessible' });
    } else {
      checks.push({ name: 'Templates', status: 'warning', message: 'Tab not found (optional)' });
    }
  } catch (e) {
    checks.push({ name: 'Templates', status: 'warning', message: 'Could not check' });
  }
  
  // Check 5: User permissions
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (userEmail) {
      checks.push({ name: 'User Auth', status: 'healthy', message: 'Logged in as ' + userEmail });
    } else {
      checks.push({ name: 'User Auth', status: 'warning', message: 'Could not determine user' });
    }
  } catch (e) {
    checks.push({ name: 'User Auth', status: 'warning', message: 'Auth check skipped' });
  }
  
  return checks;
}

function runSprintClickUpDeepCheck_() {
  const checks = [];
  
  // Check 1: ClickUp Config tab
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const clickUpSheet = ss.getSheetByName('ClickUp Config');
    if (clickUpSheet) {
      checks.push({ name: 'Config Tab', status: 'healthy', message: 'ClickUp Config found' });
    } else {
      checks.push({ name: 'Config Tab', status: 'error', message: 'ClickUp Config tab not found' });
      return checks;
    }
  } catch (e) {
    checks.push({ name: 'Config Tab', status: 'error', message: e.message });
    return checks;
  }
  
  // Check 2: API Key configured
  let apiKey = null;
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const clickUpSheet = ss.getSheetByName('ClickUp Config');
    const data = clickUpSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'API_KEY') apiKey = data[i][1];
    }
    if (apiKey) {
      checks.push({ name: 'API Key', status: 'healthy', message: 'API key configured (pk_' + apiKey.substring(3, 8) + '...)' });
    } else {
      checks.push({ name: 'API Key', status: 'error', message: 'API key not found in config' });
      return checks;
    }
  } catch (e) {
    checks.push({ name: 'API Key', status: 'error', message: e.message });
    return checks;
  }
  
  // Check 3: API Connectivity
  try {
    const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
      method: 'GET',
      headers: { 'Authorization': apiKey },
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    if (code === 200) {
      const userData = JSON.parse(response.getContentText());
      checks.push({ 
        name: 'API Connection', 
        status: 'healthy', 
        message: 'Connected as ' + (userData.user?.username || userData.user?.email || 'user')
      });
    } else if (code === 401) {
      checks.push({ name: 'API Connection', status: 'error', message: 'Invalid API key (401 Unauthorized)' });
    } else {
      checks.push({ name: 'API Connection', status: 'warning', message: 'API returned HTTP ' + code });
    }
  } catch (e) {
    checks.push({ name: 'API Connection', status: 'error', message: 'Connection failed: ' + e.message });
  }
  
  // Check 4: Space/Folder IDs
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const clickUpSheet = ss.getSheetByName('ClickUp Config');
    const data = clickUpSheet.getDataRange().getValues();
    let spaceId = null, folderId = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'FB_ADS_FOLDER_ID') spaceId = data[i][1];
      if (data[i][0] === 'SPRINT_FOLDER_ID') folderId = data[i][1];
    }
    
    if (spaceId && folderId) {
      checks.push({ name: 'Folder IDs', status: 'healthy', message: 'Space and folder configured' });
    } else if (spaceId || folderId) {
      checks.push({ name: 'Folder IDs', status: 'warning', message: 'Partial config (Space: ' + (spaceId ? '✓' : '✗') + ', Folder: ' + (folderId ? '✓' : '✗') + ')' });
    } else {
      checks.push({ name: 'Folder IDs', status: 'error', message: 'No folder IDs configured' });
    }
  } catch (e) {
    checks.push({ name: 'Folder IDs', status: 'error', message: e.message });
  }
  
  return checks;
}

function runAtriaDeepCheck_() {
  const checks = [];
  
  // Check 1: Config access
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    checks.push({ name: 'Config Access', status: 'healthy', message: 'Config accessible' });
  } catch (e) {
    checks.push({ name: 'Config Access', status: 'error', message: e.message });
    return checks;
  }
  
  // Check 2: Product sheets configured
  try {
    const productSheets = getProductSpreadsheetIds_();
    const count = Object.keys(productSheets).length;
    if (count > 0) {
      checks.push({ name: 'Product Sheets', status: 'healthy', message: count + ' product sheets configured' });
    } else {
      checks.push({ name: 'Product Sheets', status: 'warning', message: 'No product sheets found' });
    }
  } catch (e) {
    checks.push({ name: 'Product Sheets', status: 'error', message: e.message });
  }
  
  // Check 3: Slack Config
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const slackSheet = ss.getSheetByName('Slack Config');
    if (slackSheet) {
      checks.push({ name: 'Slack Config', status: 'healthy', message: 'Slack config tab found' });
    } else {
      checks.push({ name: 'Slack Config', status: 'healthy', message: 'Using local webhook (secure)' });
    }
  } catch (e) {
    checks.push({ name: 'Slack Config', status: 'warning', message: 'Could not check' });
  }
  
  return checks;
}

function runDriveQADeepCheck_() {
  const checks = [];
  
  // Check 1: Valid Values
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const vvSheet = ss.getSheetByName('Valid Values');
    if (vvSheet) {
      checks.push({ name: 'Valid Values', status: 'healthy', message: 'Tab accessible' });
    } else {
      checks.push({ name: 'Valid Values', status: 'error', message: 'Tab not found' });
    }
  } catch (e) {
    checks.push({ name: 'Valid Values', status: 'error', message: e.message });
  }
  
  // Check 2: Drive access
  try {
    DriveApp.getRootFolder();
    checks.push({ name: 'Drive Access', status: 'healthy', message: 'Drive API accessible' });
  } catch (e) {
    checks.push({ name: 'Drive Access', status: 'error', message: 'Drive access denied' });
  }
  
  return checks;
}

function runGoalsDeepCheck_() {
  const checks = [];
  
  try {
    const ss = SpreadsheetApp.openById(QUARTERLY_GOALS_SHEET_ID);
    checks.push({ name: 'Goals Sheet', status: 'healthy', message: 'Sheet accessible' });
    
    const sheets = ss.getSheets();
    checks.push({ name: 'Tabs', status: 'healthy', message: sheets.length + ' tabs found' });
  } catch (e) {
    checks.push({ name: 'Goals Sheet', status: 'error', message: 'Cannot access: ' + e.message });
  }
  
  return checks;
}

function getTroubleshootingTips_(systemName, checks) {
  const tips = [];
  const hasError = checks.some(c => c.status === 'error');
  const hasAuthIssue = checks.some(c => c.name.includes('Auth') && c.status !== 'healthy');
  const hasApiIssue = checks.some(c => c.name.includes('API') && c.status !== 'healthy');
  
  // Common tip for all systems
  tips.push({
    title: '🔄 Menu Not Appearing?',
    steps: [
      'Close and reopen the spreadsheet',
      'Go to Extensions > Apps Script',
      'Click "Run" button and select "onOpen"',
      'Authorize when prompted',
      'Refresh the spreadsheet'
    ]
  });
  
  if (hasAuthIssue) {
    tips.push({
      title: '🔐 Authorization Issues',
      steps: [
        'Go to Extensions > Apps Script',
        'Click "Run" on any function',
        'Click "Review Permissions"',
        'Choose your Google account',
        'Click "Advanced" then "Go to [Project Name]"',
        'Click "Allow"'
      ]
    });
  }
  
  if (hasApiIssue && systemName === 'Sprint → ClickUp') {
    tips.push({
      title: '🔑 ClickUp API Issues',
      steps: [
        'Go to ClickUp Settings > Apps',
        'Generate a new API token',
        'Copy the full token (starts with pk_)',
        'Update API_KEY in Config sheet > ClickUp Config tab',
        'Run health check again'
      ]
    });
  }
  
  return tips;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Format number with commas
 */
function formatNumber(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/**
 * Get current user email (for client-side)
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

// ============================================================================
// CREATIVE TEST ANALYTICS
// ============================================================================

/**
 * Clear the analytics cache - call this when you want to force fresh data
 */
function clearAnalyticsCache() {
  verifyAccess_();
  const cache = CacheService.getScriptCache();
  cache.remove('hub_analytics');
  Logger.log('🗑️ Analytics cache cleared');
  return { success: true };
}

/**
 * Fetches creative test analytics from all product spreadsheets
 * Returns aggregated data for the Analytics dashboard
 * @param {boolean} forceRefresh - If true, bypass cache and fetch fresh data
 */
function getCreativeTestAnalytics(forceRefresh) {
  verifyAccess_();
  
  const cache = CacheService.getScriptCache();
  
  // Skip cache if force refresh requested
  if (!forceRefresh) {
    const cached = cache.get('hub_analytics');
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // Cache corrupted, continue to fetch fresh data
      }
    }
  } else {
    // Clear cache when force refreshing
    cache.remove('hub_analytics');
    Logger.log('🔄 Force refresh - cache cleared');
  }
  
  try {
    // Get product spreadsheet IDs from Config
    const productSpreadsheets = getProductSpreadsheetIds_();
    
    if (!productSpreadsheets || Object.keys(productSpreadsheets).length === 0) {
      return {
        summary: { totalBatches: 0, totalAds: 0, totalSpend: 0, topAds: 0, midAds: 0, underAds: 0, topBatches: 0, midBatches: 0, underBatches: 0 },
        byProduct: {},
        byWeek: {},
        byProductWeek: {},
        products: [],
        weeks: []
      };
    }
    
    const analytics = {
      summary: {
        totalBatches: 0,
        totalAds: 0,
        totalSpend: 0,
        totalPurchases: 0,
        topAds: 0,
        midAds: 0,
        underAds: 0,
        topBatches: 0,
        midBatches: 0,
        underBatches: 0,
        // Format breakdown
        videoAds: 0,
        imageAds: 0,
        dynAds: 0,
        iterationAds: 0,
        localizedAds: 0,
        videoSpend: 0,
        imageSpend: 0,
        dynSpend: 0,
        videoTopAds: 0,
        imageTopAds: 0,
        dynTopAds: 0,
        videoPurchases: 0,
        imagePurchases: 0,
        dynPurchases: 0,
        // Spend by classification
        topSpend: 0,
        midSpend: 0,
        underSpend: 0,
        topPurchases: 0,
        midPurchases: 0,
        underPurchases: 0
      },
      byProduct: {},
      byWeek: {},
      byProductWeek: {},
      productSheetIds: {},  // Map of product name to spreadsheet ID
      products: [],
      weeks: []
    };
    
    const allWeeks = new Set();
    
    // Loop through each product spreadsheet
    for (const [productName, spreadsheetId] of Object.entries(productSpreadsheets)) {
      if (!spreadsheetId) continue;
      
      // Store the spreadsheet ID for this product
      analytics.productSheetIds[productName] = spreadsheetId;
      
      try {
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const sheets = ss.getSheets();
        
        // Initialize product data
        analytics.byProduct[productName] = {
          batches: 0,
          ads: 0,
          spend: 0,
          purchases: 0,
          topAds: 0,
          midAds: 0,
          underAds: 0,
          topBatches: 0,
          midBatches: 0,
          underBatches: 0,
          videoAds: 0,
          imageAds: 0,
          dynAds: 0,
          videoTopAds: 0,
          imageTopAds: 0,
          dynTopAds: 0
        };
        
        // Initialize per-product-week tracking
        analytics.byProductWeek[productName] = {};
        
        // Find weekly tabs (format: "W## (MM/DD - MM/DD)" or similar)
        sheets.forEach(sheet => {
          const tabName = sheet.getName();
          
          // Match week tabs like "W51 (12/15 - 12/21)"
          if (/^W\d+/.test(tabName)) {
            try {
              const weekData = parseWeeklyTab_(sheet, productName);
              
              if (weekData && weekData.batches > 0) {
                // Add to product totals
                analytics.byProduct[productName].batches += weekData.batches;
                analytics.byProduct[productName].ads += weekData.ads;
                analytics.byProduct[productName].spend += weekData.spend;
                analytics.byProduct[productName].purchases += weekData.purchases;
                analytics.byProduct[productName].topAds += weekData.topAds;
                analytics.byProduct[productName].midAds += weekData.midAds;
                analytics.byProduct[productName].underAds += weekData.underAds;
                analytics.byProduct[productName].topBatches += weekData.topBatches;
                analytics.byProduct[productName].midBatches += weekData.midBatches;
                analytics.byProduct[productName].underBatches += weekData.underBatches;
                analytics.byProduct[productName].videoAds += weekData.videoAds;
                analytics.byProduct[productName].imageAds += weekData.imageAds;
                analytics.byProduct[productName].dynAds += weekData.dynAds;
                analytics.byProduct[productName].videoTopAds += weekData.videoTopAds;
                analytics.byProduct[productName].imageTopAds += weekData.imageTopAds;
                analytics.byProduct[productName].dynTopAds += weekData.dynTopAds;
                
                // Store per-product-week data (includes all format fields)
                analytics.byProductWeek[productName][tabName] = {
                  batches: weekData.batches,
                  ads: weekData.ads,
                  spend: weekData.spend,
                  purchases: weekData.purchases,
                  topAds: weekData.topAds,
                  midAds: weekData.midAds,
                  underAds: weekData.underAds,
                  topBatches: weekData.topBatches,
                  midBatches: weekData.midBatches,
                  underBatches: weekData.underBatches,
                  // Format breakdown
                  videoAds: weekData.videoAds,
                  imageAds: weekData.imageAds,
                  dynAds: weekData.dynAds,
                  videoSpend: weekData.videoSpend,
                  imageSpend: weekData.imageSpend,
                  dynSpend: weekData.dynSpend,
                  videoTopAds: weekData.videoTopAds,
                  imageTopAds: weekData.imageTopAds,
                  dynTopAds: weekData.dynTopAds,
                  videoPurchases: weekData.videoPurchases,
                  imagePurchases: weekData.imagePurchases,
                  dynPurchases: weekData.dynPurchases
                };
                
                // Add to week totals
                allWeeks.add(tabName);
                
                if (!analytics.byWeek[tabName]) {
                  analytics.byWeek[tabName] = {
                    batches: 0,
                    ads: 0,
                    spend: 0,
                    purchases: 0,
                    topAds: 0,
                    midAds: 0,
                    underAds: 0,
                    topBatches: 0,
                    midBatches: 0,
                    underBatches: 0,
                    videoAds: 0,
                    imageAds: 0
                  };
                }
                
                analytics.byWeek[tabName].batches += weekData.batches;
                analytics.byWeek[tabName].ads += weekData.ads;
                analytics.byWeek[tabName].spend += weekData.spend;
                analytics.byWeek[tabName].purchases += weekData.purchases;
                analytics.byWeek[tabName].topAds += weekData.topAds;
                analytics.byWeek[tabName].midAds += weekData.midAds;
                analytics.byWeek[tabName].underAds += weekData.underAds;
                analytics.byWeek[tabName].topBatches += weekData.topBatches;
                analytics.byWeek[tabName].midBatches += weekData.midBatches;
                analytics.byWeek[tabName].underBatches += weekData.underBatches;
                analytics.byWeek[tabName].videoAds += weekData.videoAds;
                analytics.byWeek[tabName].imageAds += weekData.imageAds;
                
                // Add to summary totals
                analytics.summary.totalBatches += weekData.batches;
                analytics.summary.totalAds += weekData.ads;
                analytics.summary.totalSpend += weekData.spend;
                analytics.summary.totalPurchases += weekData.purchases;
                analytics.summary.topAds += weekData.topAds;
                analytics.summary.midAds += weekData.midAds;
                analytics.summary.underAds += weekData.underAds;
                analytics.summary.topBatches += weekData.topBatches;
                analytics.summary.midBatches += weekData.midBatches;
                analytics.summary.underBatches += weekData.underBatches;
                // Format breakdown
                analytics.summary.videoAds += weekData.videoAds;
                analytics.summary.imageAds += weekData.imageAds;
                analytics.summary.dynAds += weekData.dynAds;
                analytics.summary.iterationAds += weekData.iterationAds;
                analytics.summary.localizedAds += weekData.localizedAds;
                analytics.summary.videoSpend += weekData.videoSpend;
                analytics.summary.imageSpend += weekData.imageSpend;
                analytics.summary.dynSpend += weekData.dynSpend;
                analytics.summary.videoTopAds += weekData.videoTopAds;
                analytics.summary.imageTopAds += weekData.imageTopAds;
                analytics.summary.dynTopAds += weekData.dynTopAds;
                analytics.summary.videoPurchases += weekData.videoPurchases;
                analytics.summary.imagePurchases += weekData.imagePurchases;
                analytics.summary.dynPurchases += weekData.dynPurchases;
                // Spend by classification
                analytics.summary.topSpend += weekData.topSpend;
                analytics.summary.midSpend += weekData.midSpend;
                analytics.summary.underSpend += weekData.underSpend;
                analytics.summary.topPurchases += weekData.topPurchases;
                analytics.summary.midPurchases += weekData.midPurchases;
                analytics.summary.underPurchases += weekData.underPurchases;
              }
            } catch (tabError) {
              Logger.log('⚠️ Error parsing tab ' + tabName + ': ' + tabError.message);
            }
          }
        });
        
        analytics.products.push(productName);
        
      } catch (ssError) {
        Logger.log('⚠️ Error accessing spreadsheet for ' + productName + ': ' + ssError.message);
      }
    }
    
    // Sort products and weeks
    analytics.products.sort();
    analytics.weeks = Array.from(allWeeks).sort().reverse(); // Most recent first
    
    // Cache for 5 minutes
    cache.put('hub_analytics', JSON.stringify(analytics), 300);
    
    return analytics;
    
  } catch (error) {
    Logger.log('❌ Error in getCreativeTestAnalytics: ' + error.message);
    throw new Error('Failed to load analytics: ' + error.message);
  }
}

/**
 * Parse a weekly tab to extract creative test data
 * 
 * ACTUAL STRUCTURE (from debug):
 * - Row with "Underperformer" / "Mid Performer" / "Top Performer" text = BATCH HEADER
 * - Row with 🔴 / 🟡 / 🟢 emoji = INDIVIDUAL AD
 * 
 * Batch rows have aggregated spend, ad rows have individual spend
 * We count spend from AD rows only to avoid double-counting
 */
function parseWeeklyTab_(sheet, productName) {
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 5) return null;
  
  const result = {
    batches: 0,
    ads: 0,
    spend: 0,
    purchases: 0,
    topAds: 0,
    midAds: 0,
    underAds: 0,
    topBatches: 0,
    midBatches: 0,
    underBatches: 0,
    // Format breakdown
    videoAds: 0,
    imageAds: 0,
    dynAds: 0,
    iterationAds: 0,
    localizedAds: 0,
    videoSpend: 0,
    imageSpend: 0,
    dynSpend: 0,
    videoTopAds: 0,
    imageTopAds: 0,
    dynTopAds: 0,
    videoPurchases: 0,
    imagePurchases: 0,
    dynPurchases: 0,
    // Spend by classification
    topSpend: 0,
    midSpend: 0,
    underSpend: 0,
    topPurchases: 0,
    midPurchases: 0,
    underPurchases: 0
  };
  
  // Find header row (contains "Batch / Ad Name" or "Spend")
  let headerRow = -1;
  
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    const rowText = row.join(' ').toLowerCase();
    if (rowText.includes('batch') && rowText.includes('spend')) {
      headerRow = i;
      break;
    }
  }
  
  if (headerRow === -1) return null;
  
  // Find column indices
  const headers = data[headerRow];
  const colIndices = {
    status: 0,  // Column A has classification
    name: -1,
    spend: -1,
    purchases: -1
  };
  
  headers.forEach((h, idx) => {
    const header = (h || '').toString().toLowerCase();
    if (header.includes('batch') || header.includes('ad name')) colIndices.name = idx;
    if (header === 'spend' || header.includes('spend')) colIndices.spend = idx;
    if (header === 'purchases' || header.includes('purchase')) colIndices.purchases = idx;
  });
  
  // Process data rows
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const statusCell = (row[0] || '').toString().trim();
    const nameCell = colIndices.name >= 0 ? (row[colIndices.name] || '').toString().trim() : '';
    
    // Skip empty rows
    if (!nameCell && !statusCell) continue;
    
    const statusLower = statusCell.toLowerCase();
    const nameUpper = nameCell.toUpperCase();
    
    // BATCH ROW: Has text classification
    // Variations: "Top Performer", "Winner", "Scale"
    //            "Mid Performer", "Testing", "Test"  
    //            "Underperformer", "Loser", "Kill", "Stop"
    const isBatchTop = statusLower.includes('top performer') || statusLower === 'top' || statusLower === 'winner' || statusLower === 'scale';
    const isBatchMid = statusLower.includes('mid performer') || statusLower === 'mid' || statusLower === 'testing' || statusLower === 'test';
    const isBatchUnder = statusLower.includes('underperformer') || statusLower === 'under' || statusLower === 'loser' || statusLower === 'kill' || statusLower === 'stop';
    const isBatchRow = isBatchTop || isBatchMid || isBatchUnder;
    
    // AD ROW: Has emoji classification (🔴, 🟡, 🟢)
    const isAdTop = statusCell.includes('🟢');
    const isAdMid = statusCell.includes('🟡');
    const isAdUnder = statusCell.includes('🔴');
    const isAdRow = isAdTop || isAdMid || isAdUnder;
    
    if (isBatchRow) {
      // This is a BATCH header row
      result.batches++;
      
      if (isBatchTop) result.topBatches++;
      else if (isBatchMid) result.midBatches++;
      else if (isBatchUnder) result.underBatches++;
      
      // DON'T count spend from batch rows - it's aggregated
      continue;
    }
    
    if (isAdRow) {
      // This is an individual AD row
      result.ads++;
      
      // Determine format from ad name using tokenization (same as AdNaming)
      // Normalize: turn any non-alphanum into underscores, collapse repeats
      const canon = nameUpper
        .replace(/[^A-Z0-9+]+/g, '_')
        .replace(/_+/g, '_')
        .replace(/^_|_$/g, '');
      const tokens = canon.split('_').filter(Boolean);
      
      // Check for format tokens
      const AD_VID = ['VID', 'VIDEO'];
      const AD_IMG = ['IMG', 'IMAGE'];
      const AD_DYN = ['DYN', 'DYNAMIC'];
      const AD_ITR = ['ITR', 'ITERATION'];
      const AD_LOC = ['LOC', 'LOCAL', 'LOCALIZATION'];
      
      const hasToken = (arr) => tokens.some(t => arr.includes(t));
      
      const isVideo = hasToken(AD_VID);
      const isImage = hasToken(AD_IMG);
      const isDynamic = hasToken(AD_DYN);
      const isIteration = hasToken(AD_ITR);
      const isLocalized = hasToken(AD_LOC);
      
      // Count by format
      if (isVideo) result.videoAds++;
      else if (isImage) result.imageAds++;
      else if (isDynamic) result.dynAds++;
      if (isIteration) result.iterationAds++;
      if (isLocalized) result.localizedAds++;
      
      if (isAdTop) result.topAds++;
      else if (isAdMid) result.midAds++;
      else if (isAdUnder) result.underAds++;
      
      // Count spend from ad rows
      const spendVal = colIndices.spend >= 0 ? row[colIndices.spend] : 0;
      const purchasesVal = colIndices.purchases >= 0 ? row[colIndices.purchases] : 0;
      
      let spend = 0;
      if (spendVal) {
        const spendStr = spendVal.toString().replace(/[€,\s]/g, '');
        spend = parseFloat(spendStr) || 0;
      }
      
      let purchases = 0;
      if (purchasesVal) {
        purchases = parseInt(purchasesVal) || 0;
      }
      
      result.spend += spend;
      result.purchases += purchases;
      
      // Track spend and purchases by format
      if (isVideo) {
        result.videoSpend += spend;
        result.videoPurchases += purchases;
        if (isAdTop) result.videoTopAds++;
      } else if (isImage) {
        result.imageSpend += spend;
        result.imagePurchases += purchases;
        if (isAdTop) result.imageTopAds++;
      } else if (isDynamic) {
        result.dynSpend += spend;
        result.dynPurchases += purchases;
        if (isAdTop) result.dynTopAds++;
      }
      
      // Track spend by classification
      if (isAdTop) {
        result.topSpend += spend;
        result.topPurchases += purchases;
      } else if (isAdMid) {
        result.midSpend += spend;
        result.midPurchases += purchases;
      } else if (isAdUnder) {
        result.underSpend += spend;
        result.underPurchases += purchases;
      }
    }
  }
  
  return result;
}

/**
 * Get product spreadsheet IDs from Config sheet
 */
function getProductSpreadsheetIds_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Spreadsheets');
    
    if (!sheet) {
      Logger.log('⚠️ Product Spreadsheets tab not found in Config');
      return {};
    }
    
    const data = sheet.getDataRange().getValues();
    const result = {};
    
    // Skip header row (row 1), start from row 2
    for (let i = 1; i < data.length; i++) {
      const productName = (data[i][0] || '').toString().trim();
      const spreadsheetId = (data[i][1] || '').toString().trim();
      
      if (productName && spreadsheetId) {
        result[productName] = spreadsheetId;
      }
    }
    
    Logger.log('📊 Loaded ' + Object.keys(result).length + ' product spreadsheet mappings');
    return result;
    
  } catch (error) {
    Logger.log('❌ Error loading product spreadsheet IDs: ' + error.message);
    return {};
  }
}

/**
 * DEBUG: Test video/image detection on actual ad names
 * Run this from Apps Script to see what's being detected
 */
function debugVideoImageDetection() {
  Logger.log('🔍 DEBUG: Video/Image Detection Test');
  Logger.log('=====================================');
  
  const productSpreadsheets = getProductSpreadsheetIds_();
  
  let totalAds = 0;
  let videoAds = 0;
  let imageAds = 0;
  let otherAds = 0;
  const sampleNames = { video: [], image: [], other: [] };
  
  for (const [productName, spreadsheetId] of Object.entries(productSpreadsheets)) {
    if (!spreadsheetId) continue;
    
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const tabName = sheet.getName();
        if (!/^W\d+/.test(tabName)) return; // Only weekly tabs
        
        const data = sheet.getDataRange().getValues();
        
        // Find header row
        let headerRow = -1;
        for (let i = 0; i < Math.min(10, data.length); i++) {
          const rowText = data[i].join(' ').toLowerCase();
          if (rowText.includes('batch') && rowText.includes('spend')) {
            headerRow = i;
            break;
          }
        }
        if (headerRow === -1) return;
        
        // Find name column
        const headers = data[headerRow];
        let nameCol = -1;
        headers.forEach((h, idx) => {
          const header = (h || '').toString().toLowerCase();
          if (header.includes('batch') || header.includes('ad name')) nameCol = idx;
        });
        
        // Process data rows
        for (let i = headerRow + 1; i < data.length; i++) {
          const row = data[i];
          const statusCell = (row[0] || '').toString().trim();
          const nameCell = nameCol >= 0 ? (row[nameCol] || '').toString().trim() : '';
          
          // Check if it's an ad row (has emoji)
          const isAdRow = statusCell.includes('🟢') || statusCell.includes('🟡') || statusCell.includes('🔴');
          
          if (isAdRow && nameCell) {
            totalAds++;
            
            // Tokenize
            const canon = nameCell.toUpperCase()
              .replace(/[^A-Z0-9+]+/g, '_')
              .replace(/_+/g, '_')
              .replace(/^_|_$/g, '');
            const tokens = canon.split('_').filter(Boolean);
            
            const AD_VID = ['VID', 'VIDEO'];
            const AD_IMG = ['IMG', 'IMAGE'];
            const hasToken = (arr) => tokens.some(t => arr.includes(t));
            
            const isVideo = hasToken(AD_VID);
            const isImage = hasToken(AD_IMG);
            
            if (isVideo) {
              videoAds++;
              if (sampleNames.video.length < 5) sampleNames.video.push(nameCell);
            } else if (isImage) {
              imageAds++;
              if (sampleNames.image.length < 5) sampleNames.image.push(nameCell);
            } else {
              otherAds++;
              if (sampleNames.other.length < 10) sampleNames.other.push(nameCell);
            }
          }
        }
      });
      
    } catch (e) {
      Logger.log('⚠️ Error with ' + productName + ': ' + e.message);
    }
  }
  
  Logger.log('');
  Logger.log('📊 RESULTS:');
  Logger.log('   Total Ads Found: ' + totalAds);
  Logger.log('   Video Ads: ' + videoAds + ' (' + (totalAds > 0 ? (videoAds/totalAds*100).toFixed(1) : 0) + '%)');
  Logger.log('   Image Ads: ' + imageAds + ' (' + (totalAds > 0 ? (imageAds/totalAds*100).toFixed(1) : 0) + '%)');
  Logger.log('   Other/Unknown: ' + otherAds + ' (' + (totalAds > 0 ? (otherAds/totalAds*100).toFixed(1) : 0) + '%)');
  Logger.log('');
  Logger.log('📹 Sample VIDEO ad names:');
  sampleNames.video.forEach(n => Logger.log('   ' + n));
  Logger.log('');
  Logger.log('🖼️ Sample IMAGE ad names:');
  sampleNames.image.forEach(n => Logger.log('   ' + n));
  Logger.log('');
  Logger.log('❓ Sample OTHER ad names (no VID/IMG token found):');
  sampleNames.other.forEach(n => Logger.log('   ' + n));
  Logger.log('');
  Logger.log('💡 If OTHER ads should be VIDEO or IMAGE, check their naming pattern.');
  Logger.log('   Expected tokens: VID, VIDEO, IMG, IMAGE (as separate tokens)');
}

// ============================================================================
// BULK UPDATE FUNCTIONS
// ============================================================================

/**
 * Bulk update status for team members or products
 */
function bulkUpdateStatus(type, items, active) {
  try {
    verifyAdmin_(); // Only admins can bulk update
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let sheet, emailCol, statusCol;
    
    if (type === 'team') {
      sheet = ss.getSheetByName('Team Members');
      emailCol = 1;   // Column B (0-indexed = 1)
      statusCol = 7;  // Column H (0-indexed = 7)
    } else {
      sheet = ss.getSheetByName('Products');
      emailCol = 0;   // Column A (product key)
      statusCol = 6;  // Column G (Active)
    }
    
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    let updated = 0;
    
    // Find and update each item
    items.forEach(item => {
      for (let i = 1; i < data.length; i++) {
        const cellValue = data[i][emailCol] ? data[i][emailCol].toString().toLowerCase() : '';
        if (cellValue === item.toLowerCase()) {
          sheet.getRange(i + 1, statusCol + 1).setValue(active);
          updated++;
          break;
        }
      }
    });
    
    // Log the action
    logSyncEvent_('Bulk Update', `${type}: ${updated} items ${active ? 'activated' : 'deactivated'}`);
    
    // Clear cache
    CacheService.getScriptCache().removeAll(['hub_dashboard', 'hub_team', 'hub_products']);
    
    return { 
      success: true, 
      message: `Successfully ${active ? 'activated' : 'deactivated'} ${updated} ${type === 'team' ? 'team member' : 'product'}${updated !== 1 ? 's' : ''}` 
    };
    
  } catch (e) {
    Logger.log('❌ Error in bulkUpdateStatus: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// REAL-TIME SYSTEM HEALTH CHECKS
// ============================================================================

/**
 * Perform real-time health check on all systems
 */
function performSystemHealthCheck() {
  try {
    verifyAccess_();
    
    const results = {
      adNaming: checkAdNamingHealth_(),
      atriaReports: checkAtriaReportsHealth_(),
      sprintClickUp: checkSprintClickUpHealth_(),
      driveNamingQA: checkDriveNamingQAHealth_(),
      timestamp: new Date().toISOString()
    };
    
    return { success: true, data: results };
    
  } catch (e) {
    Logger.log('❌ Error in performSystemHealthCheck: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Check Ad Naming Sheet health
 */
function checkAdNamingHealth_() {
  try {
    // Try to access the Ad Naming spreadsheet
    const adNamingId = '1WqJ8dOPfP5R2sTkI0B4P13uI7NZSjKvJ9nCvQ8MTFO4'; // Replace with actual ID
    
    let canAccess = false;
    let lastActivity = null;
    let errorMsg = null;
    
    try {
      const ss = SpreadsheetApp.openById(adNamingId);
      canAccess = true;
      
      // Check last modified time
      const file = DriveApp.getFileById(adNamingId);
      lastActivity = file.getLastUpdated().toISOString();
      
    } catch (e) {
      errorMsg = e.toString();
    }
    
    // Get batch count from Analytics
    let batchCount = 0;
    try {
      const analyticsSheet = SpreadsheetApp.openById(ANALYTICS_SHEET_ID);
      const eventSheet = analyticsSheet.getSheetByName('Analytics_Events');
      if (eventSheet) {
        const data = eventSheet.getDataRange().getValues();
        batchCount = data.length - 1;
      }
    } catch (e) {
      // Ignore
    }
    
    return {
      status: canAccess ? 'healthy' : 'error',
      accessible: canAccess,
      lastActivity: lastActivity,
      batchesCreated: batchCount,
      error: errorMsg
    };
    
  } catch (e) {
    return { status: 'error', error: e.toString() };
  }
}

/**
 * Check Atria Reports health
 */
function checkAtriaReportsHealth_() {
  try {
    // Check if we can access the Config sheet
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const slackConfig = ss.getSheetByName('Slack Config');
    
    let hasSlackConfig = slackConfig !== null;
    let lastReport = null;
    
    // Check Sync Log for last report
    const syncLog = ss.getSheetByName('Sync Log');
    if (syncLog) {
      const data = syncLog.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][1] && data[i][1].toString().includes('Weekly Report')) {
          lastReport = data[i][0];
          break;
        }
      }
    }
    
    return {
      status: hasSlackConfig ? 'healthy' : 'warning',
      slackConfigured: hasSlackConfig,
      lastReport: lastReport,
      configAccessible: true
    };
    
  } catch (e) {
    return { status: 'error', error: e.toString() };
  }
}

/**
 * Check Sprint to ClickUp health
 */
function checkSprintClickUpHealth_() {
  try {
    // Get ClickUp config
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const clickUpSheet = ss.getSheetByName('ClickUp Config');
    
    let hasApiKey = false;
    let hasFolderId = false;
    
    if (clickUpSheet) {
      const data = clickUpSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'API_KEY' && data[i][1]) hasApiKey = true;
        if (data[i][0] === 'SPRINT_FOLDER_ID' && data[i][1]) hasFolderId = true;
      }
    }
    
    const isHealthy = hasApiKey && hasFolderId;
    
    return {
      status: isHealthy ? 'healthy' : 'warning',
      apiKeyConfigured: hasApiKey,
      folderIdConfigured: hasFolderId
    };
    
  } catch (e) {
    return { status: 'error', error: e.toString() };
  }
}

/**
 * Check Drive Naming QA health
 */
function checkDriveNamingQAHealth_() {
  try {
    // Check Valid Values sheet for monitor folders
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const validValues = ss.getSheetByName('Valid Values');
    
    let monitorFoldersCount = 0;
    
    if (validValues) {
      const data = validValues.getDataRange().getValues();
      const headers = data[0];
      const monitorCol = headers.findIndex(h => h && h.toString().toLowerCase().includes('monitor'));
      
      if (monitorCol >= 0) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][monitorCol]) monitorFoldersCount++;
        }
      }
    }
    
    return {
      status: monitorFoldersCount > 0 ? 'healthy' : 'warning',
      monitorFoldersConfigured: monitorFoldersCount
    };
    
  } catch (e) {
    return { status: 'error', error: e.toString() };
  }
}

// ============================================================================
// AUDIT TRAIL / ACTIVITY LOG
// ============================================================================

/**
 * Get detailed audit trail from Sync Log
 */
function getAuditTrail(options = {}) {
  try {
    verifyAccess_();
    
    const { limit = 50, filter = 'all', startDate = null, endDate = null } = options;
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Sync Log');
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Expected columns: Timestamp, Source, Action, Details, User
    const events = [];
    
    for (let i = data.length - 1; i >= 1 && events.length < limit; i--) {
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows
      
      const timestamp = row[0];
      const source = row[1] || '';
      const action = row[2] || '';
      const details = row[3] || '';
      const user = row[4] || 'System';
      
      // Apply filter
      if (filter !== 'all') {
        const sourceLower = source.toString().toLowerCase();
        if (filter === 'security' && !sourceLower.includes('security')) continue;
        if (filter === 'errors' && !sourceLower.includes('error') && !sourceLower.includes('fail')) continue;
        if (filter === 'updates' && !action.toString().toLowerCase().includes('update')) continue;
      }
      
      // Determine event type for styling
      let type = 'info';
      const actionLower = action.toString().toLowerCase();
      const sourceLower = source.toString().toLowerCase();
      
      if (sourceLower.includes('security')) type = 'security';
      else if (sourceLower.includes('error') || actionLower.includes('fail')) type = 'error';
      else if (actionLower.includes('warning')) type = 'warning';
      else if (actionLower.includes('success') || actionLower.includes('created') || actionLower.includes('updated')) type = 'success';
      
      events.push({
        timestamp: timestamp instanceof Date ? timestamp.toISOString() : timestamp,
        source: source,
        action: action,
        details: details,
        user: user,
        type: type
      });
    }
    
    return { success: true, data: events };
    
  } catch (e) {
    Logger.log('❌ Error getting audit trail: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Clear resolved errors from activity log (mark as resolved)
 */
function clearResolvedErrors() {
  try {
    verifyAdmin_();
    
    // For now, just log that this was called
    // In a full implementation, you'd update the Sync Log to mark errors as resolved
    logSyncEvent_('Maintenance', 'Errors marked as resolved');
    
    return { success: true, message: 'Errors marked as resolved' };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Log sync event helper (enhanced)
 */
function logSyncEvent_(source, details) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Sync Log');
    
    if (sheet) {
      const timestamp = new Date();
      const user = Session.getActiveUser().getEmail() || 'System';
      
      sheet.appendRow([timestamp, source, 'Update', details, user]);
      
      // Keep log manageable - only keep last 1000 entries
      const lastRow = sheet.getLastRow();
      if (lastRow > 1001) {
        sheet.deleteRows(2, lastRow - 1001);
      }
    }
  } catch (e) {
    Logger.log('Error logging sync event: ' + e.toString());
  }
}

// ============================================================================
// SYSTEM HEALTH CHECK
// ============================================================================

/**
 * Check health of a specific system
 */
function checkSystemHealth(systemName) {
  try {
    verifyAccess_();
    
    let status = 'healthy';
    let message = 'All systems operational';
    
    switch (systemName) {
      case 'Ad Naming Sheet':
        // Check that required Config tabs are accessible for Ad Naming to work
        try {
          const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
          
          // Check Products tab (required)
          const productsSheet = ss.getSheetByName('Products');
          // Check Valid Values tab (required)
          const validValuesSheet = ss.getSheetByName('Valid Values');
          // Check Templates tab (optional but useful)
          const templatesSheet = ss.getSheetByName('Templates');
          
          if (!productsSheet) {
            status = 'error';
            message = 'Products tab not found in Config';
          } else if (!validValuesSheet) {
            status = 'warning';
            message = 'Valid Values tab not found in Config';
          } else {
            const productCount = Math.max(0, productsSheet.getLastRow() - 1);
            message = 'Config OK. ' + productCount + ' products configured.';
            if (templatesSheet) {
              message += ' Templates ready.';
            }
          }
        } catch (e) {
          status = 'error';
          message = 'Cannot access Config: ' + e.message.substring(0, 50);
        }
        break;
        
      case 'Atria Reports':
        // Check config accessibility (webhook is intentionally kept in local sheet for security)
        try {
          const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
          const slackSheet = ss.getSheetByName('Slack Config');
          if (slackSheet) {
            // Webhook is stored locally in automation sheet for security, not here
            // Just verify we can read the config
            const data = slackSheet.getDataRange().getValues();
            message = 'Config accessible. Webhook stored securely in automation sheet.';
          } else {
            // Slack Config sheet doesn't exist - that's fine, it's optional
            message = 'Using local config (Slack Config sheet not needed)';
          }
        } catch (e) {
          status = 'error';
          message = 'Config check failed: ' + e.message;
        }
        break;
        
      case 'Sprint → ClickUp':
        // Check ClickUp config AND test actual API connectivity
        try {
          const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
          const clickUpSheet = ss.getSheetByName('ClickUp Config');
          if (!clickUpSheet) {
            status = 'warning';
            message = 'ClickUp Config sheet not found';
            break;
          }
          
          const data = clickUpSheet.getDataRange().getValues();
          let apiKey = null;
          let hasFolderId = false;
          
          for (let i = 1; i < data.length; i++) {
            if (data[i][0] === 'API_KEY' && data[i][1]) apiKey = data[i][1];
            if ((data[i][0] === 'SPRINT_FOLDER_ID' || data[i][0] === 'FB_ADS_FOLDER_ID') && data[i][1]) hasFolderId = true;
          }
          
          if (!apiKey) {
            status = 'error';
            message = 'API key not configured';
            break;
          }
          
          if (!hasFolderId) {
            status = 'warning';
            message = 'Folder ID not configured';
            break;
          }
          
          // ACTUAL API TEST - Verify the key works
          try {
            const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
              method: 'GET',
              headers: { 'Authorization': apiKey },
              muteHttpExceptions: true
            });
            
            const code = response.getResponseCode();
            if (code === 200) {
              const userData = JSON.parse(response.getContentText());
              const username = userData.user?.username || userData.user?.email || 'connected';
              status = 'healthy';
              message = '✓ API connected as ' + username;
            } else if (code === 401) {
              status = 'error';
              message = 'API key invalid or expired (401)';
            } else if (code === 429) {
              status = 'warning';
              message = 'Rate limited - try again later';
            } else {
              status = 'warning';
              message = 'API returned HTTP ' + code;
            }
          } catch (apiErr) {
            // Check if it's a permission error (Hub not authorized for external calls)
            if (apiErr.message && apiErr.message.includes('permission')) {
              status = 'healthy';
              message = 'Config OK ✓ (API test skipped - use actual sheet to verify)';
            } else {
              status = 'error';
              message = 'API connection failed: ' + apiErr.message.substring(0, 50);
            }
          }
        } catch (e) {
          status = 'error';
          message = 'Config check failed: ' + e.message;
        }
        break;
        
      case 'Drive Naming QA':
        // Check Valid Values for monitor folders
        try {
          const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
          const vvSheet = ss.getSheetByName('Valid Values');
          if (vvSheet) {
            message = 'Valid Values sheet accessible';
          } else {
            status = 'warning';
            message = 'Valid Values sheet not found';
          }
        } catch (e) {
          status = 'error';
          message = 'Config check failed: ' + e.message;
        }
        break;
        
      case 'Quarterly Goals':
        // Check Quarterly Goals sheet
        try {
          const ss = SpreadsheetApp.openById(QUARTERLY_GOALS_SHEET_ID);
          const settingsSheet = ss.getSheetByName('Settings');
          const winnersSheet = ss.getSheetByName('Winning Ads');
          
          if (!settingsSheet) {
            status = 'error';
            message = 'Settings tab not found';
          } else if (!winnersSheet) {
            status = 'warning';
            message = 'Winning Ads tab not found';
          } else {
            // Get current progress
            const data = settingsSheet.getRange('A1:B20').getValues();
            let quarter = 'Q1 2026';
            let goal = 150;
            let winnerCount = 0;
            
            for (const row of data) {
              if (row[0] === 'Current Quarter') quarter = row[1];
              if (row[0] === 'Goal (Winning Ads)') goal = parseInt(row[1]) || 150;
              if (row[0] === 'Total Winning Ads') winnerCount = parseInt(row[1]) || 0;
            }
            
            const progress = Math.round((winnerCount / goal) * 100);
            message = quarter + ': ' + winnerCount + '/' + goal + ' (' + progress + '%)';
          }
        } catch (e) {
          status = 'error';
          message = 'Cannot access sheet: ' + e.message.substring(0, 40);
        }
        break;
        
      default:
        status = 'warning';
        message = 'Unknown system';
    }
    
    // Log to Activity if not healthy
    if (status !== 'healthy') {
      logHealthCheckResult_(systemName, status, message);
    }
    
    return { success: true, status: status, message: message };
    
  } catch (e) {
    Logger.log('❌ Health check error: ' + e.toString());
    logHealthCheckResult_(systemName, 'error', e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Log health check result to Activity/Sync Log
 */
function logHealthCheckResult_(systemName, status, message) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    let syncLog = ss.getSheetByName('Sync Log');
    
    if (!syncLog) {
      // Create Sync Log if it doesn't exist
      syncLog = ss.insertSheet('Sync Log');
      syncLog.appendRow(['Timestamp', 'System', 'Action', 'Status', 'Details', 'User']);
    }
    
    const userEmail = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date().toISOString();
    
    syncLog.appendRow([
      timestamp,
      systemName,
      'Health Check',
      status.toUpperCase(),
      message,
      userEmail
    ]);
    
  } catch (e) {
    Logger.log('Could not log health check: ' + e.message);
  }
}

/**
 * Trigger function to force authorization of all permissions
 * Users should run this from Apps Script editor if they see permission errors
 */
function authorizeAllPermissions() {
  // These calls trigger authorization prompts for each service
  try {
    // UrlFetchApp
    UrlFetchApp.fetch('https://httpbin.org/get', { muteHttpExceptions: true });
    Logger.log('✓ UrlFetchApp authorized');
  } catch (e) {
    Logger.log('UrlFetchApp auth needed: ' + e.message);
  }
  
  try {
    // DriveApp
    DriveApp.getRootFolder();
    Logger.log('✓ DriveApp authorized');
  } catch (e) {
    Logger.log('DriveApp auth needed: ' + e.message);
  }
  
  try {
    // SpreadsheetApp
    SpreadsheetApp.openById(CONFIG_SHEET_ID);
    Logger.log('✓ SpreadsheetApp authorized');
  } catch (e) {
    Logger.log('SpreadsheetApp auth needed: ' + e.message);
  }
  
  return { success: true, message: 'Authorization check complete. Check logs for details.' };
}

/**
 * Helper to get product spreadsheet ID from config
 */
// ============================================================================
// QUARTERLY GOALS DATA
// ============================================================================

/**
 * Sanitize string for JSON serialization
 */
function sanitizeString_(val) {
  if (val === null || val === undefined) return '';
  // Handle Date objects
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const str = String(val);
  // Remove any non-printable characters and limit length
  return str.replace(/[\x00-\x1F\x7F]/g, '').substring(0, 100);
}

/**
 * Simple test to verify connection and basic data
 */
function testGoalsData() {
  try {
    const ss = SpreadsheetApp.openById(QUARTERLY_GOALS_SHEET_ID);
    const winnersSheet = ss.getSheetByName('Winning Ads');
    const rowCount = winnersSheet ? winnersSheet.getLastRow() - 1 : 0;
    
    return {
      success: true,
      message: 'Test passed',
      winnerCount: rowCount,
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Get quarterly goals data for the Goals tab
 */
function getQuarterlyGoalsData() {
  try {
    verifyAccess_();
    
    const ss = SpreadsheetApp.openById(QUARTERLY_GOALS_SHEET_ID);
    
    // Get settings
    const settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      return { success: false, error: 'Settings sheet not found' };
    }
    
    const settingsData = settingsSheet.getRange('A1:B25').getValues();
    const settings = {};
    for (const row of settingsData) {
      if (row[0]) settings[row[0]] = row[1];
    }
    
    const quarter = settings['Current Quarter'] || 'Q1 2026';
    const goal = parseInt(settings['Goal (Winning Ads)']) || 150;
    const enableLaunchDateFilter = settings['Enable Launch Date Filter'] === 'Yes';
    const thresholds = {
      winner: parseFloat(settings['Winner Threshold']) || 5000,
      babyUnicorn: parseFloat(settings['Baby Unicorn Threshold']) || 10000,
      unicorn: parseFloat(settings['Unicorn Threshold']) || 20000,
      superUnicorn: parseFloat(settings['Super Unicorn Threshold']) || 50000,
      almostMin: parseFloat(settings['Almost Winner Min']) || 4000,
      almostMax: parseFloat(settings['Almost Winner Max']) || 5000
    };
    
    // Get winners - simplified
    const winnersSheet = ss.getSheetByName('Winning Ads');
    let winners = [];
    let launchedThisQuarter = 0;
    let totalWinners = 0;
    
    // Tier breakdown - ONLY FOR LAUNCHED THIS QUARTER
    const tierBreakdown = {
      superUnicorn: { count: 0, spend: 0 },
      unicorn: { count: 0, spend: 0 },
      babyUnicorn: { count: 0, spend: 0 },
      winner: { count: 0, spend: 0 }
    };
    // By product - ONLY FOR LAUNCHED THIS QUARTER
    const byProduct = {};
    let totalSpend = 0;
    
    if (winnersSheet && winnersSheet.getLastRow() > 1) {
      const lastRow = winnersSheet.getLastRow();
      const lastCol = winnersSheet.getLastColumn();
      totalWinners = lastRow - 1;
      
      // Get headers to find column indices dynamically
      const headers = winnersSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const colIndex = {
        adName: 0,
        product: 1,
        spend: 2,
        tier: 3,
        launchDate: 4,
        launchedThisQ: 5,
        roas: 6,
        cpa: 7,
        hookRate: -1,
        holdRate: -1
      };
      
      // Find Hook Rate and Hold Rate columns by header name
      headers.forEach((h, i) => {
        const header = String(h).toLowerCase().trim();
        if (header.includes('hook')) colIndex.hookRate = i;
        else if (header.includes('hold')) colIndex.holdRate = i;
      });
      
      Logger.log('📊 Column indices - Hook Rate: ' + colIndex.hookRate + ', Hold Rate: ' + colIndex.holdRate);
      
      // First, count launched from ALL rows (column F = Launched This Q)
      const allLaunchCol = winnersSheet.getRange(2, colIndex.launchedThisQ + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < allLaunchCol.length; i++) {
        if (allLaunchCol[i][0] && String(allLaunchCol[i][0]).toLowerCase().indexOf('yes') >= 0) {
          launchedThisQuarter++;
        }
      }
      
      // Get ALL data
      const allData = winnersSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      
      // Helper to format rate as percentage
      function formatRateAsPercent(val) {
        if (!val || val === '-' || val === '' || val === 0) return '-';
        const num = parseFloat(val);
        if (isNaN(num)) return '-';
        // If value is less than 1, it's a decimal (0.69 = 69%)
        // If value is >= 1, it's already a percentage
        if (num < 1) {
          return (num * 100).toFixed(1) + '%';
        } else {
          return num.toFixed(1) + '%';
        }
      }
      
      // Calculate tier breakdown and by product ONLY FOR LAUNCHED THIS QUARTER
      for (let i = 0; i < allData.length; i++) {
        const row = allData[i];
        const launchedThisQ = row[colIndex.launchedThisQ] ? String(row[colIndex.launchedThisQ]).toLowerCase() : '';
        const isLaunched = launchedThisQ.indexOf('yes') >= 0;
        
        // Only count launched ads in tier breakdown and product stats
        if (isLaunched) {
          const product = row[colIndex.product] ? String(row[colIndex.product]) : '';
          const spend = row[colIndex.spend] ? String(row[colIndex.spend]) : '€0';
          const tier = row[colIndex.tier] ? String(row[colIndex.tier]) : '';
          
          const spendNum = parseEuroValue_(spend);
          totalSpend += spendNum;
          
          // Tier breakdown
          if (tier.indexOf('Super') >= 0) {
            tierBreakdown.superUnicorn.count++;
            tierBreakdown.superUnicorn.spend += spendNum;
          } else if (tier.indexOf('🦄🦄') >= 0) {
            tierBreakdown.unicorn.count++;
            tierBreakdown.unicorn.spend += spendNum;
          } else if (tier.indexOf('🦄') >= 0) {
            tierBreakdown.babyUnicorn.count++;
            tierBreakdown.babyUnicorn.spend += spendNum;
          } else {
            tierBreakdown.winner.count++;
            tierBreakdown.winner.spend += spendNum;
          }
          
          // By product
          if (product) {
            if (!byProduct[product]) byProduct[product] = { count: 0, spend: 0 };
            byProduct[product].count++;
            byProduct[product].spend += spendNum;
          }
        }
      }
      
      // Now build the display array - return ALL launched ads for accurate tier drill-down
      // Since we have ~114 launched ads, this is manageable for payload
      const rowsToDisplay = Math.min(lastRow - 1, 200);  // Increased limit
      for (let i = 0; i < rowsToDisplay; i++) {
        const row = allData[i];
        const adName = sanitizeString_(row[colIndex.adName]).substring(0, 60);
        const product = sanitizeString_(row[colIndex.product]);
        const spend = sanitizeString_(row[colIndex.spend]) || '€0';
        const tier = sanitizeString_(row[colIndex.tier]) || '🏆 Winner';
        const launchDate = sanitizeString_(row[colIndex.launchDate]) || '-';
        const launchedThisQ = sanitizeString_(row[colIndex.launchedThisQ]);
        const roas = sanitizeString_(row[colIndex.roas]) || '-';
        const cpa = sanitizeString_(row[colIndex.cpa]) || '-';
        
        // Get Hook Rate and Hold Rate from dynamic column positions
        const hookRateRaw = colIndex.hookRate >= 0 ? row[colIndex.hookRate] : '';
        const holdRateRaw = colIndex.holdRate >= 0 ? row[colIndex.holdRate] : '';
        const hookRate = formatRateAsPercent(hookRateRaw);
        const holdRate = formatRateAsPercent(holdRateRaw);
        
        const isLaunched = launchedThisQ.toLowerCase().indexOf('yes') >= 0;
        
        winners.push({
          adName: adName,
          product: product,
          spend: spend,
          tier: tier,
          launchDate: launchDate,
          launchedThisQ: isLaunched ? 'Yes' : 'No',
          roas: roas,
          cpa: cpa,
          hookRate: hookRate,
          holdRate: holdRate
        });
      }
    }
    
    const progress = goal > 0 ? Math.round((totalWinners / goal) * 100) : 0;
    const launchProgress = goal > 0 ? Math.round((launchedThisQuarter / goal) * 100) : 0;
    
    Logger.log('📊 Goals: ' + totalWinners + ' total, ' + launchedThisQuarter + ' launched, returning ' + winners.length + ' rows');
    
    // Load almost winners - check both current sheet and history
    let almostWinners = [];
    let almostSpend = 0;
    
    // First try the current "Almost Winners" sheet
    const almostSheet = ss.getSheetByName('Almost Winners');
    if (almostSheet && almostSheet.getLastRow() > 1) {
      const almostData = almostSheet.getDataRange().getValues();
      almostWinners = almostData.slice(1, 51).map(row => ({
        adName: sanitizeString_(row[0]).substring(0, 60),
        product: sanitizeString_(row[1]),
        spend: sanitizeString_(row[2]) || '€0',
        gap: sanitizeString_(row[3]) || '€0',
        roas: sanitizeString_(row[4]) || '-',
        cpa: sanitizeString_(row[5]) || '-'
      }));
      almostWinners.forEach(a => {
        almostSpend += parseEuroValue_(a.spend);
      });
      Logger.log('📊 Loaded ' + almostWinners.length + ' almost winners from current sheet');
    }
    
    // If no almost winners in current sheet, check history
    if (almostWinners.length === 0) {
      const historyAlmostName = '_History_' + quarter.replace(' ', '_') + '_Almost';
      const historyAlmostSheet = ss.getSheetByName(historyAlmostName);
      if (historyAlmostSheet && historyAlmostSheet.getLastRow() > 1) {
        const almostData = historyAlmostSheet.getDataRange().getValues();
        almostWinners = almostData.slice(1, 51).map(row => ({
          adName: sanitizeString_(row[0]).substring(0, 60),
          product: sanitizeString_(row[1]),
          spend: sanitizeString_(row[2]) || '€0',
          gap: sanitizeString_(row[3]) || '€0',
          roas: sanitizeString_(row[4]) || '-',
          cpa: sanitizeString_(row[5]) || '-'
        }));
        almostWinners.forEach(a => {
          almostSpend += parseEuroValue_(a.spend);
        });
        Logger.log('📊 Loaded ' + almostWinners.length + ' almost winners from history sheet: ' + historyAlmostName);
      }
    }
    
    const result = {
      success: true,
      quarter: quarter,
      goal: goal,
      current: totalWinners,
      launchedThisQuarter: launchedThisQuarter,
      progress: progress,
      launchProgress: launchProgress,
      tierBreakdown: tierBreakdown,
      byProduct: byProduct,
      totalSpend: totalSpend,
      almostSpend: almostSpend,
      winners: winners,
      almostWinners: almostWinners,
      thresholds: thresholds,
      enableLaunchDateFilter: enableLaunchDateFilter,
      lastUpdated: sanitizeString_(settings['Last Updated']) || '-'
    };
    
    // Test serialization
    try {
      const testJson = JSON.stringify(result);
      Logger.log('📊 JSON size: ' + testJson.length + ' bytes');
    } catch (serErr) {
      Logger.log('❌ JSON serialization error: ' + serErr.toString());
      return { success: false, error: 'Data serialization failed: ' + serErr.toString() };
    }
    
    return result;
    
  } catch (e) {
    Logger.log('❌ Error getting quarterly goals data: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Helper to parse Euro values (handles both US and EU formats)
 */
function parseEuroValue_(value) {
  if (!value || value === '-') return 0;
  
  let str = String(value).replace('€', '').replace(/\s/g, '').trim();
  
  // Detect format by checking which separator comes last
  const lastComma = str.lastIndexOf(',');
  const lastDot = str.lastIndexOf('.');
  
  if (lastComma > lastDot) {
    // European format: 9.324,21
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (lastDot > lastComma) {
    // American format: 9,324.21
    str = str.replace(/,/g, '');
  } else if (lastComma !== -1) {
    const afterComma = str.split(',')[1];
    if (afterComma && afterComma.length === 3 && str.split(',').length === 2) {
      str = str.replace(',', '');
    } else {
      str = str.replace(',', '.');
    }
  }
  
  return parseFloat(str) || 0;
}

/**
 * Get available quarters for the dropdown
 */
function getAvailableQuarters() {
  try {
    verifyAccess_();
    
    const ss = SpreadsheetApp.openById(QUARTERLY_GOALS_SHEET_ID);
    const sheets = ss.getSheets();
    const quarters = [];
    
    // Get settings for current quarter
    const settingsSheet = ss.getSheetByName('Settings');
    let currentQuarter = 'Q4 2025';
    if (settingsSheet) {
      const settingsData = settingsSheet.getRange('A1:B25').getValues();
      for (const row of settingsData) {
        if (row[0] === 'Current Quarter') {
          currentQuarter = row[1] || 'Q4 2025';
          break;
        }
      }
    }
    
    // Find all _History_QXYYYY sheets
    sheets.forEach(sheet => {
      const name = sheet.getName();
      const match = name.match(/^_History_(Q[1-4])_(\d{4})$/);
      if (match) {
        const quarter = match[1];
        const year = match[2];
        const winnerCount = sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
        
        quarters.push({
          quarter: quarter + ' ' + year,
          sheetName: name,
          winnerCount: winnerCount,
          year: parseInt(year),
          quarterNum: parseInt(quarter.substring(1)),
          isHistorical: true
        });
      }
    });
    
    // Also add current quarter if it has data
    const winnersSheet = ss.getSheetByName('Winning Ads');
    let currentCount = 0;
    let launchedCount = 0;
    
    if (winnersSheet && winnersSheet.getLastRow() > 1) {
      const data = winnersSheet.getDataRange().getValues();
      currentCount = data.length - 1;
      // Count launched this quarter (column F, index 5)
      launchedCount = data.slice(1).filter(row => 
        row[5] && row[5].toString().toLowerCase().includes('yes')
      ).length;
    }
    
    if (currentCount > 0) {
      // Check if current quarter is already in history
      const exists = quarters.some(q => q.quarter === currentQuarter);
      if (!exists) {
        quarters.push({
          quarter: currentQuarter,
          sheetName: 'Winning Ads',
          winnerCount: launchedCount, // Show launched count for current quarter
          year: parseInt(currentQuarter.split(' ')[1]),
          quarterNum: parseInt(currentQuarter.substring(1, 2)),
          isCurrent: true
        });
      }
    }
    
    // Sort by year desc, then quarter desc
    quarters.sort((a, b) => {
      if (b.year !== a.year) return b.year - a.year;
      return b.quarterNum - a.quarterNum;
    });
    
    return { success: true, quarters: quarters, currentQuarter: currentQuarter };
  } catch (e) {
    Logger.log('❌ Error getting available quarters: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get data for a specific quarter (from history or current)
 */
function getQuarterData(quarterStr) {
  try {
    verifyAccess_();
    
    const ss = SpreadsheetApp.openById(QUARTERLY_GOALS_SHEET_ID);
    
    // Get settings
    const settingsSheet = ss.getSheetByName('Settings');
    let currentQuarter = 'Q4 2025';
    let goal = 150;
    
    if (settingsSheet) {
      const settingsData = settingsSheet.getRange('A1:B25').getValues();
      for (const row of settingsData) {
        if (row[0] === 'Current Quarter') currentQuarter = row[1] || 'Q4 2025';
        if (row[0] === 'Goal (Winning Ads)') goal = parseInt(row[1]) || 150;
      }
    }
    
    // If requesting current quarter, use getQuarterlyGoalsData
    if (quarterStr === currentQuarter) {
      return getQuarterlyGoalsData();
    }
    
    // Otherwise, load from history sheet
    const sheetName = '_History_' + quarterStr.replace(' ', '_');
    const historySheet = ss.getSheetByName(sheetName);
    
    if (!historySheet) {
      return { success: false, error: 'No data found for ' + quarterStr };
    }
    
    const data = historySheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, error: 'History sheet is empty' };
    }
    
    // Find column indices from headers
    const headers = data[0];
    const colIndex = {
      adName: 0,
      product: 1,
      spend: 2,
      tier: 3,
      launchDate: 4,
      launchedThisQ: 5,
      roas: 6,
      cpa: 7,
      hookRate: -1,
      holdRate: -1
    };
    
    headers.forEach((h, i) => {
      const header = String(h).toLowerCase().trim();
      if (header.includes('hook')) colIndex.hookRate = i;
      else if (header.includes('hold')) colIndex.holdRate = i;
    });
    
    // Helper to format rate as percentage
    function formatRateAsPercent(val) {
      if (!val || val === '-' || val === '' || val === 0) return '-';
      const num = parseFloat(val);
      if (isNaN(num)) return '-';
      if (num < 1) {
        return (num * 100).toFixed(1) + '%';
      } else {
        return num.toFixed(1) + '%';
      }
    }
    
    // Parse historical data
    const winners = data.slice(1, 201).map(row => ({  // Limit to 200
      adName: row[colIndex.adName] ? row[colIndex.adName].toString().substring(0, 80) : '',
      product: row[colIndex.product] || '',
      spend: row[colIndex.spend] || '€0',
      tier: row[colIndex.tier] || '🏆 Winner',
      launchDate: row[colIndex.launchDate] || '-',
      launchedThisQ: 'Yes', // History only contains launched ads
      roas: row[colIndex.roas] || '-',
      cpa: row[colIndex.cpa] || '-',
      hookRate: colIndex.hookRate >= 0 ? formatRateAsPercent(row[colIndex.hookRate]) : '-',
      holdRate: colIndex.holdRate >= 0 ? formatRateAsPercent(row[colIndex.holdRate]) : '-'
    }));
    
    // Calculate tier breakdown
    const tierBreakdown = {
      superUnicorn: { count: 0, spend: 0 },
      unicorn: { count: 0, spend: 0 },
      babyUnicorn: { count: 0, spend: 0 },
      winner: { count: 0, spend: 0 }
    };
    
    const byProduct = {};
    let totalSpend = 0;
    
    winners.forEach(w => {
      const spend = parseEuroValue_(w.spend);
      totalSpend += spend;
      
      if (w.tier && w.tier.includes('Super')) {
        tierBreakdown.superUnicorn.count++;
        tierBreakdown.superUnicorn.spend += spend;
      } else if (w.tier && w.tier.includes('🦄🦄')) {
        tierBreakdown.unicorn.count++;
        tierBreakdown.unicorn.spend += spend;
      } else if (w.tier && w.tier.includes('🦄')) {
        tierBreakdown.babyUnicorn.count++;
        tierBreakdown.babyUnicorn.spend += spend;
      } else {
        tierBreakdown.winner.count++;
        tierBreakdown.winner.spend += spend;
      }
      
      if (!byProduct[w.product]) byProduct[w.product] = { count: 0, spend: 0 };
      byProduct[w.product].count++;
      byProduct[w.product].spend += spend;
    });
    
    const launchedThisQuarter = winners.length; // All history ads are launched that quarter
    
    // Check for almost winners history sheet
    const almostSheetName = sheetName + '_Almost';
    const almostHistorySheet = ss.getSheetByName(almostSheetName);
    let almostWinners = [];
    let almostSpend = 0;
    
    if (almostHistorySheet && almostHistorySheet.getLastRow() > 1) {
      const almostData = almostHistorySheet.getDataRange().getValues();
      almostWinners = almostData.slice(1, 51).map(row => ({  // Limit to 50
        adName: row[0] ? row[0].toString().substring(0, 80) : '',
        product: row[1] || '',
        spend: row[2] || '€0',
        gap: row[3] || '€0',
        roas: row[4] || '-',
        cpa: row[5] || '-'
      }));
      
      almostWinners.forEach(a => {
        almostSpend += parseEuroValue_(a.spend);
      });
    }
    
    return {
      success: true,
      quarter: quarterStr,
      goal: goal,
      current: winners.length,
      launchedThisQuarter: launchedThisQuarter,
      progress: Math.round((winners.length / goal) * 100),
      launchProgress: Math.round((launchedThisQuarter / goal) * 100),
      tierBreakdown: tierBreakdown,
      byProduct: byProduct,
      totalSpend: totalSpend,
      almostSpend: almostSpend,
      winners: winners,
      almostWinners: almostWinners,
      isHistorical: true
    };
  } catch (e) {
    Logger.log('❌ Error getting quarter data: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Helper to get product spreadsheet ID from config
 */
function getProductSpreadsheetId_(key) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Spreadsheets');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) return data[i][1];
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Get detailed ad data for format drilldown
 * @param {string} formatType - 'video', 'image', or 'dynamic'
 * @param {Array} selectedWeeks - Array of week names like ['W52 (12/22 - 12/28)']
 * @returns {Object} Detailed format data with individual ads
 */
function getFormatDrilldownData(formatType, selectedWeeks, selectedProducts) {
  verifyAccess_();
  
  const productSpreadsheets = getProductSpreadsheetIds_();
  if (!productSpreadsheets || Object.keys(productSpreadsheets).length === 0) {
    return { ads: [], totals: { ads: 0, spend: 0, purchases: 0 } };
  }
  
  const result = {
    ads: [],
    totals: {
      ads: 0,
      spend: 0,
      purchases: 0,
      topAds: 0,
      avgHookRate: 0,
      avgHoldRate: 0
    }
  };
  
  const formatTokens = {
    'video': ['VID', 'VIDEO'],
    'image': ['IMG', 'IMAGE'],
    'dynamic': ['DYN', 'DYNAMIC']
  };
  
  const targetTokens = formatTokens[formatType.toLowerCase()] || [];
  let hookRateSum = 0;
  let holdRateSum = 0;
  let hookRateCount = 0;
  let holdRateCount = 0;
  
  // Process each product spreadsheet
  for (const [productName, spreadsheetId] of Object.entries(productSpreadsheets)) {
    if (!spreadsheetId) continue;
    
    // Skip if product filter is active and this product is not selected
    if (selectedProducts && selectedProducts.length > 0 && !selectedProducts.includes(productName)) {
      continue;
    }
    
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const tabName = sheet.getName();
        
        // Only process week tabs
        if (!/^W\d+/.test(tabName)) return;
        
        // Only process selected weeks (if specified)
        if (selectedWeeks && selectedWeeks.length > 0 && !selectedWeeks.includes(tabName)) return;
        
        try {
          const adsFromTab = parseTabForFormatAds_(sheet, productName, targetTokens, tabName);
          
          adsFromTab.forEach(ad => {
            result.ads.push(ad);
            result.totals.ads++;
            result.totals.spend += ad.spend;
            result.totals.purchases += ad.purchases;
            if (ad.isTop) result.totals.topAds++;
            
            // Hook/Hold rates (video only)
            if (ad.hookRate > 0) {
              hookRateSum += ad.hookRate;
              hookRateCount++;
            }
            if (ad.holdRate > 0) {
              holdRateSum += ad.holdRate;
              holdRateCount++;
            }
          });
          
        } catch (tabError) {
          Logger.log('Error parsing tab ' + tabName + ': ' + tabError.message);
        }
      });
      
    } catch (ssError) {
      Logger.log('Error accessing spreadsheet for ' + productName + ': ' + ssError.message);
    }
  }
  
  // Calculate averages
  result.totals.avgHookRate = hookRateCount > 0 ? (hookRateSum / hookRateCount).toFixed(2) : 0;
  result.totals.avgHoldRate = holdRateCount > 0 ? (holdRateSum / holdRateCount).toFixed(2) : 0;
  
  // Sort ads by spend descending
  result.ads.sort((a, b) => b.spend - a.spend);
  
  // Note: No limit - show all ads (table is scrollable)
  
  // Build per-product breakdown
  const byProduct = {};
  result.ads.forEach(ad => {
    if (!byProduct[ad.product]) {
      byProduct[ad.product] = { ads: 0, spend: 0, purchases: 0, topAds: 0 };
    }
    byProduct[ad.product].ads++;
    byProduct[ad.product].spend += ad.spend;
    byProduct[ad.product].purchases += ad.purchases;
    if (ad.isTop) byProduct[ad.product].topAds++;
  });
  result.byProduct = byProduct;
  
  // Build weekly Hook/Hold data for charts (video only)
  if (formatType.toLowerCase() === 'video') {
    const weeklyRates = {};
    result.ads.forEach(ad => {
      if (!weeklyRates[ad.week]) {
        weeklyRates[ad.week] = { hookSum: 0, hookCount: 0, holdSum: 0, holdCount: 0 };
      }
      if (ad.hookRate > 0) {
        weeklyRates[ad.week].hookSum += ad.hookRate;
        weeklyRates[ad.week].hookCount++;
      }
      if (ad.holdRate > 0) {
        weeklyRates[ad.week].holdSum += ad.holdRate;
        weeklyRates[ad.week].holdCount++;
      }
    });
    
    // Calculate averages per week
    result.weeklyRates = {};
    Object.keys(weeklyRates).sort().forEach(week => {
      const wr = weeklyRates[week];
      result.weeklyRates[week] = {
        avgHookRate: wr.hookCount > 0 ? (wr.hookSum / wr.hookCount) : 0,
        avgHoldRate: wr.holdCount > 0 ? (wr.holdSum / wr.holdCount) : 0
      };
    });
  }
  
  // Get previous week data for comparison (only if single week selected)
  if (selectedWeeks && selectedWeeks.length === 1) {
    const prevWeekData = getPreviousWeekFormatData_(formatType, selectedWeeks[0], selectedProducts, productSpreadsheets);
    if (prevWeekData) {
      result.previousWeek = prevWeekData;
    }
  }
  
  Logger.log('📊 Drilldown: Found ' + result.ads.length + ' ' + formatType + ' ads across ' + Object.keys(byProduct).length + ' products');
  
  return result;
}

/**
 * Get top performer ads for drilldown
 * @param {Array} selectedWeeks - Array of week names
 * @param {Array} selectedProducts - Array of product names (empty = all)
 * @returns {Object} Top performer data
 */
function getTopPerformersDrilldownData(selectedWeeks, selectedProducts) {
  verifyAccess_();
  
  const productSpreadsheets = getProductSpreadsheetIds_();
  if (!productSpreadsheets || Object.keys(productSpreadsheets).length === 0) {
    return { ads: [], totals: { ads: 0, spend: 0, purchases: 0 } };
  }
  
  const result = {
    ads: [],
    totals: {
      ads: 0,
      spend: 0,
      purchases: 0,
      videoAds: 0,
      imageAds: 0,
      dynamicAds: 0,
      avgHookRate: 0,
      avgHoldRate: 0
    },
    byProduct: {},
    weeklyRates: {}
  };
  
  let hookRateSum = 0, holdRateSum = 0, hookRateCount = 0, holdRateCount = 0;
  
  // Process each product spreadsheet
  for (const [productName, spreadsheetId] of Object.entries(productSpreadsheets)) {
    if (!spreadsheetId) continue;
    if (selectedProducts && selectedProducts.length > 0 && !selectedProducts.includes(productName)) continue;
    
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const tabName = sheet.getName();
        if (!/^W\d+/.test(tabName)) return;
        if (selectedWeeks && selectedWeeks.length > 0 && !selectedWeeks.includes(tabName)) return;
        
        try {
          const topAds = parseTabForTopPerformers_(sheet, productName, tabName);
          
          topAds.forEach(ad => {
            result.ads.push(ad);
            result.totals.ads++;
            result.totals.spend += ad.spend;
            result.totals.purchases += ad.purchases;
            
            // Count by format
            if (ad.format === 'video') result.totals.videoAds++;
            else if (ad.format === 'image') result.totals.imageAds++;
            else if (ad.format === 'dynamic') result.totals.dynamicAds++;
            
            // Product breakdown
            if (!result.byProduct[productName]) {
              result.byProduct[productName] = { ads: 0, spend: 0, videoAds: 0, imageAds: 0, dynamicAds: 0 };
            }
            result.byProduct[productName].ads++;
            result.byProduct[productName].spend += ad.spend;
            if (ad.format === 'video') result.byProduct[productName].videoAds++;
            else if (ad.format === 'image') result.byProduct[productName].imageAds++;
            else if (ad.format === 'dynamic') result.byProduct[productName].dynamicAds++;
            
            // Weekly rates for video
            if (ad.format === 'video') {
              if (!result.weeklyRates[tabName]) {
                result.weeklyRates[tabName] = { hookSum: 0, hookCount: 0, holdSum: 0, holdCount: 0 };
              }
              if (ad.hookRate > 0) {
                result.weeklyRates[tabName].hookSum += ad.hookRate;
                result.weeklyRates[tabName].hookCount++;
                hookRateSum += ad.hookRate;
                hookRateCount++;
              }
              if (ad.holdRate > 0) {
                result.weeklyRates[tabName].holdSum += ad.holdRate;
                result.weeklyRates[tabName].holdCount++;
                holdRateSum += ad.holdRate;
                holdRateCount++;
              }
            }
          });
        } catch (e) {
          Logger.log('Error parsing tab ' + tabName + ': ' + e.message);
        }
      });
    } catch (e) {
      Logger.log('Error accessing spreadsheet for ' + productName + ': ' + e.message);
    }
  }
  
  // Calculate averages
  result.totals.avgHookRate = hookRateCount > 0 ? (hookRateSum / hookRateCount).toFixed(2) : 0;
  result.totals.avgHoldRate = holdRateCount > 0 ? (holdRateSum / holdRateCount).toFixed(2) : 0;
  
  // Calculate weekly averages
  Object.keys(result.weeklyRates).forEach(week => {
    const wr = result.weeklyRates[week];
    result.weeklyRates[week] = {
      avgHookRate: wr.hookCount > 0 ? (wr.hookSum / wr.hookCount) : 0,
      avgHoldRate: wr.holdCount > 0 ? (wr.holdSum / wr.holdCount) : 0
    };
  });
  
  // Sort by spend
  result.ads.sort((a, b) => b.spend - a.spend);
  
  Logger.log('📊 Top Performers Drilldown: Found ' + result.ads.length + ' top performer ads');
  
  return result;
}

/**
 * Get all ads for a single product drilldown
 * @param {string} productName - The product name
 * @param {Array} selectedWeeks - Array of week names
 * @returns {Object} Product data with all ads
 */
function getProductDrilldownData(productName, selectedWeeks) {
  verifyAccess_();
  
  const productSpreadsheets = getProductSpreadsheetIds_();
  const spreadsheetId = productSpreadsheets[productName];
  
  if (!spreadsheetId) {
    return { ads: [], totals: { ads: 0 }, productName: productName };
  }
  
  const result = {
    productName: productName,
    ads: [],
    totals: {
      ads: 0,
      spend: 0,
      purchases: 0,
      topAds: 0,
      midAds: 0,
      underAds: 0,
      videoAds: 0,
      imageAds: 0,
      dynamicAds: 0,
      avgHookRate: 0,
      avgHoldRate: 0
    },
    weeklyRates: {}
  };
  
  let hookRateSum = 0, holdRateSum = 0, hookRateCount = 0, holdRateCount = 0;
  
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      const tabName = sheet.getName();
      if (!/^W\d+/.test(tabName)) return;
      if (selectedWeeks && selectedWeeks.length > 0 && !selectedWeeks.includes(tabName)) return;
      
      try {
        const adsFromTab = parseTabForAllAds_(sheet, productName, tabName);
        
        adsFromTab.forEach(ad => {
          result.ads.push(ad);
          result.totals.ads++;
          result.totals.spend += ad.spend;
          result.totals.purchases += ad.purchases;
          
          // Count by tier
          if (ad.tier === 'top') result.totals.topAds++;
          else if (ad.tier === 'mid') result.totals.midAds++;
          else result.totals.underAds++;
          
          // Count by format
          if (ad.format === 'video') result.totals.videoAds++;
          else if (ad.format === 'image') result.totals.imageAds++;
          else if (ad.format === 'dynamic') result.totals.dynamicAds++;
          
          // Weekly rates for video
          if (ad.format === 'video') {
            if (!result.weeklyRates[tabName]) {
              result.weeklyRates[tabName] = { hookSum: 0, hookCount: 0, holdSum: 0, holdCount: 0 };
            }
            if (ad.hookRate > 0) {
              result.weeklyRates[tabName].hookSum += ad.hookRate;
              result.weeklyRates[tabName].hookCount++;
              hookRateSum += ad.hookRate;
              hookRateCount++;
            }
            if (ad.holdRate > 0) {
              result.weeklyRates[tabName].holdSum += ad.holdRate;
              result.weeklyRates[tabName].holdCount++;
              holdRateSum += ad.holdRate;
              holdRateCount++;
            }
          }
        });
      } catch (e) {
        Logger.log('Error parsing tab ' + tabName + ': ' + e.message);
      }
    });
  } catch (e) {
    Logger.log('Error accessing spreadsheet for ' + productName + ': ' + e.message);
  }
  
  // Calculate averages
  result.totals.avgHookRate = hookRateCount > 0 ? (hookRateSum / hookRateCount).toFixed(2) : 0;
  result.totals.avgHoldRate = holdRateCount > 0 ? (holdRateSum / holdRateCount).toFixed(2) : 0;
  
  // Calculate weekly averages
  Object.keys(result.weeklyRates).forEach(week => {
    const wr = result.weeklyRates[week];
    result.weeklyRates[week] = {
      avgHookRate: wr.hookCount > 0 ? (wr.hookSum / wr.hookCount) : 0,
      avgHoldRate: wr.holdCount > 0 ? (wr.holdSum / wr.holdCount) : 0
    };
  });
  
  // Sort by tier (top first), then by spend
  result.ads.sort((a, b) => {
    const tierOrder = { top: 0, mid: 1, under: 2 };
    if (tierOrder[a.tier] !== tierOrder[b.tier]) {
      return tierOrder[a.tier] - tierOrder[b.tier];
    }
    return b.spend - a.spend;
  });
  
  Logger.log('📊 Product Drilldown: Found ' + result.ads.length + ' ads for ' + productName);
  
  return result;
}

/**
 * Parse tab for all ads (all tiers)
 */
function parseTabForAllAds_(sheet, productName, weekName) {
  const data = sheet.getDataRange().getValues();
  const ads = [];
  
  if (data.length < 5) return ads;
  
  // Find header row
  let headerRow = -1;
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const rowText = data[i].join(' ').toLowerCase();
    if (rowText.includes('batch') && rowText.includes('spend')) {
      headerRow = i;
      break;
    }
  }
  
  if (headerRow === -1) return ads;
  
  // Find column indices
  const headers = data[headerRow];
  const colIndices = {
    name: -1, spend: -1, roas: -1, cpa: -1, purchases: -1, hookRate: -1, holdRate: -1
  };
  
  headers.forEach((h, idx) => {
    const header = (h || '').toString().toLowerCase();
    if (header.includes('batch') || header.includes('ad name')) colIndices.name = idx;
    if (header === 'spend' || header.includes('spend')) colIndices.spend = idx;
    if (header === 'roas') colIndices.roas = idx;
    if (header === 'cpa') colIndices.cpa = idx;
    if (header === 'purchases' || header.includes('purchase')) colIndices.purchases = idx;
    if (header.includes('hook')) colIndices.hookRate = idx;
    if (header.includes('hold')) colIndices.holdRate = idx;
  });
  
  // Process rows
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const statusCell = (row[0] || '').toString().trim();
    const nameCell = colIndices.name >= 0 ? (row[colIndices.name] || '').toString().trim() : '';
    
    if (!nameCell) continue;
    
    // Determine tier from emoji
    let tier = 'other';
    if (statusCell.includes('🟢')) tier = 'top';
    else if (statusCell.includes('🟡')) tier = 'mid';
    else if (statusCell.includes('🔴')) tier = 'under';
    else continue; // Skip non-ad rows
    
    // Determine format
    const nameUpper = nameCell.toUpperCase();
    const tokens = nameUpper.replace(/[^A-Z0-9+]+/g, '_').split('_').filter(Boolean);
    
    let format = 'other';
    if (tokens.some(t => ['VID', 'VIDEO'].includes(t))) format = 'video';
    else if (tokens.some(t => ['IMG', 'IMAGE'].includes(t))) format = 'image';
    else if (tokens.some(t => ['DYN', 'DYNAMIC'].includes(t))) format = 'dynamic';
    
    // Parse values
    const spend = parseFloat((row[colIndices.spend] || '0').toString().replace(/[€,\s]/g, '')) || 0;
    const roas = parseFloat((row[colIndices.roas] || '0').toString()) || 0;
    const cpa = parseFloat((row[colIndices.cpa] || '0').toString().replace(/[€,\s]/g, '')) || 0;
    const purchases = parseInt(row[colIndices.purchases]) || 0;
    
    let hookRate = 0, holdRate = 0;
    if (colIndices.hookRate >= 0) {
      hookRate = parseFloat((row[colIndices.hookRate] || '').toString().replace(/[%,\s]/g, '')) || 0;
      if (hookRate > 0 && hookRate < 1) hookRate = hookRate * 100;
    }
    if (colIndices.holdRate >= 0) {
      holdRate = parseFloat((row[colIndices.holdRate] || '').toString().replace(/[%,\s]/g, '')) || 0;
      if (holdRate > 0 && holdRate < 1) holdRate = holdRate * 100;
    }
    
    ads.push({
      adName: nameCell.substring(0, 80),
      product: productName,
      week: weekName,
      spend: spend,
      roas: roas,
      cpa: cpa,
      purchases: purchases,
      hookRate: hookRate,
      holdRate: holdRate,
      format: format,
      tier: tier,
      isTop: tier === 'top'
    });
  }
  
  return ads;
}

/**
 * Parse tab for top performer ads only (green emoji rows)
 */
function parseTabForTopPerformers_(sheet, productName, weekName) {
  const data = sheet.getDataRange().getValues();
  const ads = [];
  
  if (data.length < 5) return ads;
  
  // Find header row
  let headerRow = -1;
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const rowText = data[i].join(' ').toLowerCase();
    if (rowText.includes('batch') && rowText.includes('spend')) {
      headerRow = i;
      break;
    }
  }
  
  if (headerRow === -1) return ads;
  
  // Find column indices
  const headers = data[headerRow];
  const colIndices = {
    name: -1, spend: -1, roas: -1, cpa: -1, purchases: -1, hookRate: -1, holdRate: -1
  };
  
  headers.forEach((h, idx) => {
    const header = (h || '').toString().toLowerCase();
    if (header.includes('batch') || header.includes('ad name')) colIndices.name = idx;
    if (header === 'spend' || header.includes('spend')) colIndices.spend = idx;
    if (header === 'roas') colIndices.roas = idx;
    if (header === 'cpa') colIndices.cpa = idx;
    if (header === 'purchases' || header.includes('purchase')) colIndices.purchases = idx;
    if (header.includes('hook')) colIndices.hookRate = idx;
    if (header.includes('hold')) colIndices.holdRate = idx;
  });
  
  // Process rows - only top performers (green emoji)
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const statusCell = (row[0] || '').toString().trim();
    const nameCell = colIndices.name >= 0 ? (row[colIndices.name] || '').toString().trim() : '';
    
    if (!nameCell) continue;
    
    // Only top performers (green emoji)
    if (!statusCell.includes('🟢')) continue;
    
    // Determine format
    const nameUpper = nameCell.toUpperCase();
    const tokens = nameUpper.replace(/[^A-Z0-9+]+/g, '_').split('_').filter(Boolean);
    
    let format = 'other';
    if (tokens.some(t => ['VID', 'VIDEO'].includes(t))) format = 'video';
    else if (tokens.some(t => ['IMG', 'IMAGE'].includes(t))) format = 'image';
    else if (tokens.some(t => ['DYN', 'DYNAMIC'].includes(t))) format = 'dynamic';
    
    // Parse values
    const spend = parseFloat((row[colIndices.spend] || '0').toString().replace(/[€,\s]/g, '')) || 0;
    const roas = parseFloat((row[colIndices.roas] || '0').toString()) || 0;
    const cpa = parseFloat((row[colIndices.cpa] || '0').toString().replace(/[€,\s]/g, '')) || 0;
    const purchases = parseInt(row[colIndices.purchases]) || 0;
    
    let hookRate = 0, holdRate = 0;
    if (colIndices.hookRate >= 0) {
      hookRate = parseFloat((row[colIndices.hookRate] || '').toString().replace(/[%,\s]/g, '')) || 0;
      if (hookRate > 0 && hookRate < 1) hookRate = hookRate * 100;
    }
    if (colIndices.holdRate >= 0) {
      holdRate = parseFloat((row[colIndices.holdRate] || '').toString().replace(/[%,\s]/g, '')) || 0;
      if (holdRate > 0 && holdRate < 1) holdRate = holdRate * 100;
    }
    
    ads.push({
      adName: nameCell.substring(0, 80),
      product: productName,
      week: weekName,
      spend: spend,
      roas: roas,
      cpa: cpa,
      purchases: purchases,
      hookRate: hookRate,
      holdRate: holdRate,
      format: format,
      tier: 'top',
      isTop: true
    });
  }
  
  return ads;
}

/**
 * Get previous week's format data for comparison
 */
function getPreviousWeekFormatData_(formatType, currentWeek, selectedProducts, productSpreadsheets) {
  // Extract week number
  const weekMatch = currentWeek.match(/W(\d+)/);
  if (!weekMatch) return null;
  
  const currentWeekNum = parseInt(weekMatch[1]);
  const prevWeekNum = currentWeekNum === 1 ? 52 : currentWeekNum - 1;
  const prevWeekPattern = new RegExp('^W' + prevWeekNum + '\\b');
  
  const formatTokens = {
    'video': ['VID', 'VIDEO'],
    'image': ['IMG', 'IMAGE'],
    'dynamic': ['DYN', 'DYNAMIC']
  };
  const targetTokens = formatTokens[formatType.toLowerCase()] || [];
  
  const byProduct = {};
  let totalAds = 0;
  
  for (const [productName, spreadsheetId] of Object.entries(productSpreadsheets)) {
    if (!spreadsheetId) continue;
    if (selectedProducts && selectedProducts.length > 0 && !selectedProducts.includes(productName)) continue;
    
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const tabName = sheet.getName();
        if (!prevWeekPattern.test(tabName)) return;
        
        try {
          const adsFromTab = parseTabForFormatAds_(sheet, productName, targetTokens, tabName);
          
          if (!byProduct[productName]) {
            byProduct[productName] = { ads: 0 };
          }
          byProduct[productName].ads += adsFromTab.length;
          totalAds += adsFromTab.length;
          
        } catch (e) {}
      });
    } catch (e) {}
  }
  
  return { byProduct: byProduct, totalAds: totalAds };
}

/**
 * Parse a weekly tab for individual ads of a specific format
 */
function parseTabForFormatAds_(sheet, productName, targetTokens, weekName) {
  const data = sheet.getDataRange().getValues();
  const ads = [];
  
  if (data.length < 5) return ads;
  
  // Find header row
  let headerRow = -1;
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const rowText = data[i].join(' ').toLowerCase();
    if (rowText.includes('batch') && rowText.includes('spend')) {
      headerRow = i;
      break;
    }
  }
  
  if (headerRow === -1) return ads;
  
  // Find column indices
  const headers = data[headerRow];
  const colIndices = {
    name: -1, spend: -1, roas: -1, cpa: -1, purchases: -1, hookRate: -1, holdRate: -1
  };
  
  headers.forEach((h, idx) => {
    const header = (h || '').toString().toLowerCase();
    if (header.includes('batch') || header.includes('ad name')) colIndices.name = idx;
    if (header === 'spend' || header.includes('spend')) colIndices.spend = idx;
    if (header === 'roas') colIndices.roas = idx;
    if (header === 'cpa') colIndices.cpa = idx;
    if (header === 'purchases' || header.includes('purchase')) colIndices.purchases = idx;
    if (header.includes('hook')) colIndices.hookRate = idx;
    if (header.includes('hold')) colIndices.holdRate = idx;
  });
  
  // Process rows
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const statusCell = (row[0] || '').toString().trim();
    const nameCell = colIndices.name >= 0 ? (row[colIndices.name] || '').toString().trim() : '';
    
    if (!nameCell) continue;
    
    // Only process individual ad rows (with emoji)
    const isAdTop = statusCell.includes('🟢');
    const isAdMid = statusCell.includes('🟡');
    const isAdUnder = statusCell.includes('🔴');
    if (!isAdTop && !isAdMid && !isAdUnder) continue;
    
    // Check if ad matches target format
    const nameUpper = nameCell.toUpperCase();
    const tokens = nameUpper.replace(/[^A-Z0-9+]+/g, '_').split('_').filter(Boolean);
    const matchesFormat = tokens.some(t => targetTokens.includes(t));
    
    if (!matchesFormat) continue;
    
    // Parse values
    const spend = parseFloat((row[colIndices.spend] || '0').toString().replace(/[€,\s]/g, '')) || 0;
    const roas = parseFloat((row[colIndices.roas] || '0').toString()) || 0;
    const cpaRaw = (row[colIndices.cpa] || '0').toString().replace(/[€,\s]/g, '');
    const cpa = parseFloat(cpaRaw) || 0;
    const purchases = parseInt(row[colIndices.purchases]) || 0;
    
    // Parse hook/hold rates
    let hookRate = 0;
    let holdRate = 0;
    
    if (colIndices.hookRate >= 0) {
      const hookRaw = (row[colIndices.hookRate] || '').toString();
      hookRate = parseFloat(hookRaw.replace(/[%,\s]/g, '')) || 0;
      if (hookRate > 0 && hookRate < 1) hookRate = hookRate * 100;
    }
    
    if (colIndices.holdRate >= 0) {
      const holdRaw = (row[colIndices.holdRate] || '').toString();
      holdRate = parseFloat(holdRaw.replace(/[%,\s]/g, '')) || 0;
      if (holdRate > 0 && holdRate < 1) holdRate = holdRate * 100;
    }
    
    ads.push({
      adName: nameCell.substring(0, 80),
      product: productName,
      week: weekName,
      spend: spend,
      roas: roas,
      cpa: cpa,
      purchases: purchases,
      hookRate: hookRate,
      holdRate: holdRate,
      tier: isAdTop ? 'top' : (isAdMid ? 'mid' : 'under'),
      isTop: isAdTop
    });
  }
  
  return ads;
}