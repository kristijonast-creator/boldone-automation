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
      'Atria Reports',
      errorType,
      String(message).substring(0, 500),
      user
    ]]);
  } catch (e) {
    Logger.log('Could not log error to Hub: ' + e);
  }
}

function getProductLocaleFoldersFromConfig_() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('atria_folders');
    if (cached) return JSON.parse(cached);
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Folders');
    
    if (!sheet) {
      Logger.log('⚠️ Product Folders tab not found in Config sheet');
      return {};
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // ['Product Key', 'US', 'UK', 'DE', ...]
    
    const folders = {};
    for (let i = 1; i < data.length; i++) {
      const productKey = data[i][0];
      if (!productKey) continue;
      
      folders[productKey] = {};
      for (let j = 1; j < headers.length; j++) {
        const locale = headers[j];
        const folderId = data[i][j];
        if (folderId && folderId.toString().trim() !== '') {
          // Convert folder ID to full URL
          folders[productKey][locale] = 'https://drive.google.com/drive/folders/' + folderId.toString().trim();
        } else {
          folders[productKey][locale] = '';
        }
      }
    }
    
    if (Object.keys(folders).length > 0) {
      cache.put('atria_folders', JSON.stringify(folders), CONFIG_CACHE_DURATION);
      Logger.log('✅ Loaded ' + Object.keys(folders).length + ' product folder mappings from Config');
    }
    
    return folders;
  } catch (e) {
    Logger.log('⚠️ Error loading product folders from Config: ' + e.message);
    return {};  // Return empty object so PRODUCT_LOCALE_FOLDERS is defined
  }
}

function getTargetCpaFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('atria_cpa');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  
  const cpa = {};
  for (let i = 1; i < data.length; i++) {
    const productKey = data[i][0];
    const targetCpa = data[i][4]; // Column E = Target CPA
    const active = data[i][6];    // Column G = Active
    
    if (productKey && targetCpa) {
      cpa[productKey] = targetCpa;
    }
  }
  
  cache.put('atria_cpa', JSON.stringify(cpa), CONFIG_CACHE_DURATION);
  return cpa;
}

function getClickUpConfigFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('atria_clickup');
  
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      // Verify the cached data has API_KEY
      if (parsed && parsed.API_KEY) {
        return parsed;
      }
    } catch (e) {
      Logger.log('⚠️ Corrupted ClickUp cache, reloading...');
    }
  }
  
  // Load fresh from sheet
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('ClickUp Config');
  const data = sheet.getDataRange().getValues();
  
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    let value = data[i][1];
    
    // Trim string values
    if (typeof value === 'string') {
      value = value.trim();
    }
    
    if (key) {
      config[key] = value;
    }
  }
  
  // Only cache if we got valid data
  if (config.API_KEY) {
    cache.put('atria_clickup', JSON.stringify(config), CONFIG_CACHE_DURATION);
    Logger.log('✅ ClickUp config cached with ' + Object.keys(config).length + ' keys');
  }
  
  return config;
}

function clearAtriaConfigCache() {
  // Clear script cache
  CacheService.getScriptCache().removeAll([
    'atria_folders', 
    'atria_cpa', 
    'atria_clickup', 
    'atria_slack',
    'atria_product_sheets'
  ]);
  
  // Clear in-memory lazy loaders
  _clickupConfigData = null;
  _productLocaleFolders = null;
  _targetCpa = null;
  
  Logger.log('✅ All config caches cleared');
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Config cache cleared!', 'Done', 3);
  } catch (e) {
    // Ignore if no active spreadsheet
  }
}

/**
 * Get Slack configuration from central Config sheet
 * Returns: { botToken, weeklyChannelId, creativeChannelId, driveNamingChannelId }
 */
function getSlackConfigFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('atria_slack');
  if (cached) return JSON.parse(cached);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Slack Config');
    
    if (!sheet) {
      Logger.log('⚠️ Slack Config tab not found in central config - using local Settings');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    const config = {};
    for (let i = 1; i < data.length; i++) {
      const setting = data[i][0];
      const value = data[i][1];
      if (setting && value) {
        config[setting] = value.toString().trim();
      }
    }
    
    const result = {
      botToken: config['BOT_TOKEN'] || '',
      weeklyChannelId: config['WEEKLY_CHANNEL_ID'] || '',
      creativeChannelId: config['CREATIVE_CHANNEL_ID'] || '',
      driveNamingChannelId: config['DRIVE_NAMING_CHANNEL_ID'] || ''
    };
    
    // Only cache if we have at least the bot token
    if (result.botToken) {
      cache.put('atria_slack', JSON.stringify(result), CONFIG_CACHE_DURATION);
      Logger.log('✅ Loaded Slack config from central Config sheet');
    }
    
    return result;
  } catch (e) {
    Logger.log('⚠️ Error loading Slack config from central Config: ' + e.message);
    return null;
  }
}

/**
 * Get product spreadsheet IDs from central Config sheet
 * Returns: { "ProductKey": "spreadsheet_id", ... }
 */
function getProductSpreadsheetsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('atria_product_sheets');
  if (cached) return JSON.parse(cached);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Product Spreadsheets');
    
    if (!sheet) {
      Logger.log('⚠️ Product Spreadsheets tab not found in central config - using local Settings');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    const productSpreadsheets = {};
    for (let i = 1; i < data.length; i++) {
      const productKey = data[i][0];
      const spreadsheetId = data[i][1];
      
      if (productKey && spreadsheetId && spreadsheetId.toString().trim() !== '') {
        productSpreadsheets[productKey.toString().trim()] = spreadsheetId.toString().trim();
      }
    }
    
    if (Object.keys(productSpreadsheets).length > 0) {
      cache.put('atria_product_sheets', JSON.stringify(productSpreadsheets), CONFIG_CACHE_DURATION);
      Logger.log('✅ Loaded ' + Object.keys(productSpreadsheets).length + ' product spreadsheet mappings from central Config');
    }
    
    return productSpreadsheets;
  } catch (e) {
    Logger.log('⚠️ Error loading product spreadsheets from central Config: ' + e.message);
    return null;
  }
}

// ================================================================================

/**
 * ATRIA TO SLACK - UPDATED VERSION WITH CSV UPLOAD & GOOGLE DRIVE LINKS
 * Handles proper CSV parsing with comma-separated values
 * Adds clickable Google Drive links to ad creatives
 */

// === Product → Locale base folders (LAZY LOADED to prevent menu breaking) ===
let _productLocaleFolders = null;

function getProductLocaleFoldersLazy_() {
  if (!_productLocaleFolders) {
    _productLocaleFolders = getProductLocaleFoldersFromConfig_();
  }
  return _productLocaleFolders;
}

/**
 * Target CPA by product (LAZY LOADED to prevent menu breaking)
 */
let _targetCpa = null;

function getTargetCpaLazy_() {
  if (!_targetCpa) {
    _targetCpa = getTargetCpaFromConfig_();
  }
  return _targetCpa;
}

// === COMPATIBILITY: Define PRODUCT_LOCALE_FOLDERS and TARGET_CPA ===
var PRODUCT_LOCALE_FOLDERS = {};
var TARGET_CPA = {};

function initializeConfigConstants_() {
  if (Object.keys(PRODUCT_LOCALE_FOLDERS).length === 0) {
    PRODUCT_LOCALE_FOLDERS = getProductLocaleFoldersLazy_();
  }
  if (Object.keys(TARGET_CPA).length === 0) {
    TARGET_CPA = getTargetCpaLazy_();
  }
  Logger.log('✅ Config initialized: ' + Object.keys(PRODUCT_LOCALE_FOLDERS).length + ' products, ' + Object.keys(TARGET_CPA).length + ' CPA targets');
}

// ===== MENU =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Atria Reports')
    .addItem('📤 Upload CSV Files', 'showUploadDialog')
    .addItem('📊 Data Info', 'showDataInfo')
    .addSeparator()
    .addItem('📈 Analyze Creative Tests', 'showCreativeTestDialog')
    .addItem('🔄 Check Analysis Progress', 'checkCreativeTestsProgress')        // 🆕 NEW
    .addItem('🗑️ Reset Analysis Progress', 'resetCreativeTestsProgress')       // 🆕 NEW
    .addItem('🧹 Force Cleanup Triggers', 'forceCleanupAllTriggers')           // 🆕 ADD THIS
    .addSeparator()
    .addItem('📊 Generate Weekly Report', 'showUploadDialog')
    .addSeparator()
    .addItem('💬 Send Weekly Report to Slack', 'sendWeeklyReportToSlack')
    .addItem('💬 Send Creative Tests to Slack', 'sendCreativeTestsToSlack')
    .addItem('🗑️ Delete Last Weekly Report', 'deleteLastWeeklyReport')
    .addItem('🗑️ Delete Last Creative Tests Report', 'deleteLastCreativeTestsReport')
    .addSeparator()
    .addItem('📤 Upload Creative Tests to Product Sheets', 'showUploadCreativeTestsDialog')
    .addItem('🔧 Test Product Spreadsheet Connections', 'testProductSpreadsheetConnections')
    .addSeparator()
    .addItem('🔍 Check CSV Data Quality', 'checkCSVDataQuality')
    .addItem('🗑️ Clear All Data', 'clearAllData')
    .addToUi();
}

function showUploadDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
          
          html, body {
            height: 100%;
            overflow: hidden;
          }
          
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: #ffffff;
            display: flex;
            flex-direction: column;
          }
          
          .dialog-header {
            background: #ffffff;
            padding: 24px 28px;
            border-bottom: 1px solid #e8e8e8;
            flex-shrink: 0;
          }
          
          .dialog-header h2 {
            font-size: 20px;
            font-weight: 600;
            margin: 0;
            color: #1f1f1f;
          }
          
          .dialog-body {
            padding: 28px;
            flex: 1;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
          }
          
          .drop-zone {
            border: 2px dashed #d1d1d1;
            border-radius: 8px;
            padding: 48px 32px;
            text-align: center;
            background: #fafafa;
            transition: all 0.2s;
            cursor: pointer;
            flex-shrink: 0;
          }
          
          .drop-zone:hover {
            border-color: #0066cc;
            background: #f5f5f5;
          }
          
          .drop-zone.dragover {
            border-color: #0066cc;
            background: #e8f2ff;
            border-style: solid;
          }
          
          .drop-zone-icon {
            font-size: 48px;
            margin-bottom: 16px;
            opacity: 0.5;
          }
          
          .drop-zone-text {
            font-size: 16px;
            color: #2c2c2c;
            font-weight: 500;
            margin-bottom: 6px;
          }
          
          .drop-zone-subtext {
            font-size: 13px;
            color: #6e6e6e;
            margin-bottom: 20px;
          }
          
          .divider {
            display: flex;
            align-items: center;
            text-align: center;
            color: #999;
            margin: 20px 0;
            font-size: 13px;
          }
          
          .divider::before,
          .divider::after {
            content: '';
            flex: 1;
            border-bottom: 1px solid #e0e0e0;
          }
          
          .divider span {
            padding: 0 14px;
          }
          
          .browse-btn {
            background: #0066cc;
            color: white;
            border: none;
            padding: 12px 28px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }
          
          .browse-btn:hover {
            background: #0052a3;
          }
          
          .browse-btn:active {
            transform: scale(0.98);
          }
          
          .file-input {
            display: none;
          }
          
          .file-list-container {
            margin-top: 20px;
            flex: 1;
            min-height: 0;
            display: flex;
            flex-direction: column;
          }
          
          .file-list {
            max-height: 240px;
            overflow-y: auto;
            overflow-x: hidden;
            padding-right: 4px;
          }
          
          .file-list::-webkit-scrollbar {
            width: 6px;
          }
          
          .file-list::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 3px;
          }
          
          .file-list::-webkit-scrollbar-thumb {
            background: #c1c1c1;
            border-radius: 3px;
          }
          
          .file-list::-webkit-scrollbar-thumb:hover {
            background: #a8a8a8;
          }
          
          .file-item {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 12px;
            background: #f8f8f8;
            border-radius: 6px;
            margin-bottom: 6px;
            transition: all 0.15s;
            border: 1px solid #e8e8e8;
          }
          
          .file-item:hover {
            background: #f0f0f0;
            border-color: #d8d8d8;
          }
          
          .file-icon {
            font-size: 22px;
            flex-shrink: 0;
          }
          
          .file-info {
            flex: 1;
            min-width: 0;
          }
          
          .file-name {
            font-weight: 500;
            color: #2c2c2c;
            font-size: 13px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
          }
          
          .file-size {
            font-size: 12px;
            color: #6e6e6e;
            margin-top: 2px;
          }
          
          .remove-btn {
            background: #ff3b30;
            color: white;
            border: none;
            width: 26px;
            height: 26px;
            border-radius: 50%;
            font-size: 16px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.15s;
            flex-shrink: 0;
            font-weight: 500;
            line-height: 1;
          }
          
          .remove-btn:hover {
            background: #ff453a;
            transform: scale(1.05);
          }
          
          .upload-btn {
            background: #34c759;
            color: white;
            border: none;
            padding: 13px 28px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            width: 100%;
            margin-top: 20px;
            transition: all 0.2s;
            display: none;
            flex-shrink: 0;
          }
          
          .upload-btn.visible {
            display: block;
          }
          
          .upload-btn:hover:not(:disabled) {
            background: #30b350;
          }
          
          .upload-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
          }
          
          .upload-btn:active:not(:disabled) {
            transform: scale(0.98);
          }
          
          #status {
            margin-top: 16px;
            padding: 12px 14px;
            border-radius: 6px;
            display: none;
            font-size: 13px;
            line-height: 1.5;
            flex-shrink: 0;
          }
          
          #status.show {
            display: block;
          }
          
          #status.success {
            background: #d1f4e0;
            border: 1px solid #34c759;
            color: #1e7e34;
          }
          
          #status.error {
            background: #ffe5e5;
            border: 1px solid #ff3b30;
            color: #d32f2f;
          }
          
          .empty-state {
            text-align: center;
            color: #999;
            font-size: 13px;
            padding: 24px 16px;
          }
        </style>
      </head>
      <body>
        <div class="dialog-header">
          <h2>📊 Upload CSV Files</h2>
        </div>
        
        <div class="dialog-body">
          <div class="drop-zone" id="dropZone">
            <div class="drop-zone-icon">📁</div>
            <div class="drop-zone-text">Drag & drop CSV files here</div>
            <div class="drop-zone-subtext">Support for multiple files</div>
            
            <div class="divider">
              <span>or</span>
            </div>
            
            <button type="button" class="browse-btn" id="browseBtn">
              Browse Files
            </button>
          </div>
          
          <input 
            type="file" 
            id="fileInput" 
            class="file-input" 
            accept=".csv" 
            multiple
          >
          
          <div class="file-list-container">
            <div class="file-list" id="fileList">
              <div class="empty-state">No files selected</div>
            </div>
          </div>
          
          <button type="button" class="upload-btn" id="uploadBtn">
            Upload & Generate Report
          </button>
          
          <div id="status"></div>
        </div>
        
        <script>
          (function() {
            var fileInput = document.getElementById('fileInput');
            var uploadBtn = document.getElementById('uploadBtn');
            var browseBtn = document.getElementById('browseBtn');
            var dropZone = document.getElementById('dropZone');
            var fileList = document.getElementById('fileList');
            var status = document.getElementById('status');
            var selectedFiles = [];
            
            // Browse button
            browseBtn.addEventListener('click', function(e) {
              e.preventDefault();
              e.stopPropagation();
              fileInput.click();
            });
            
            // Drop zone click (only if not clicking browse button)
            dropZone.addEventListener('click', function(e) {
              if (e.target !== browseBtn && !browseBtn.contains(e.target)) {
                fileInput.click();
              }
            });
            
            // Drag and drop handlers
            dropZone.addEventListener('dragenter', function(e) {
              e.preventDefault();
              e.stopPropagation();
              dropZone.classList.add('dragover');
            });
            
            dropZone.addEventListener('dragover', function(e) {
              e.preventDefault();
              e.stopPropagation();
            });
            
            dropZone.addEventListener('dragleave', function(e) {
              e.preventDefault();
              e.stopPropagation();
              dropZone.classList.remove('dragover');
            });
            
            dropZone.addEventListener('drop', function(e) {
              e.preventDefault();
              e.stopPropagation();
              dropZone.classList.remove('dragover');
              handleFiles(e.dataTransfer.files);
            });
            
            // File input change
            fileInput.addEventListener('change', function(e) {
              handleFiles(e.target.files);
            });
            
            function handleFiles(files) {
              for (var i = 0; i < files.length; i++) {
                if (files[i].name.toLowerCase().endsWith('.csv')) {
                  selectedFiles.push(files[i]);
                }
              }
              displayFiles();
              if (selectedFiles.length > 0) {
                uploadBtn.classList.add('visible');
              }
            }
            
            function removeFile(index) {
              selectedFiles.splice(index, 1);
              displayFiles();
              if (selectedFiles.length === 0) {
                uploadBtn.classList.remove('visible');
              }
            }
            
            function formatFileSize(bytes) {
              if (bytes === 0) return '0 Bytes';
              var k = 1024;
              var sizes = ['Bytes', 'KB', 'MB', 'GB'];
              var i = Math.floor(Math.log(bytes) / Math.log(k));
              return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
            }
            
            function displayFiles() {
              if (selectedFiles.length === 0) {
                fileList.innerHTML = '<div class="empty-state">No files selected</div>';
                return;
              }
              
              var html = '';
              for (var i = 0; i < selectedFiles.length; i++) {
                html += '<div class="file-item">' +
                  '<div class="file-icon">📄</div>' +
                  '<div class="file-info">' +
                    '<div class="file-name">' + selectedFiles[i].name + '</div>' +
                    '<div class="file-size">' + formatFileSize(selectedFiles[i].size) + '</div>' +
                  '</div>' +
                  '<button class="remove-btn" onclick="removeFile(' + i + ')">×</button>' +
                '</div>';
              }
              fileList.innerHTML = html;
            }
            
            // Make removeFile available globally for onclick
            window.removeFile = removeFile;
            
            // Upload button handler
            uploadBtn.addEventListener('click', handleUpload);
            
            function handleUpload() {
              if (selectedFiles.length === 0) {
                alert('Please select CSV files');
                return;
              }
              
              uploadBtn.disabled = true;
              uploadBtn.textContent = 'Processing...';
              status.className = '';
              
              var readers = [];
              for (var i = 0; i < selectedFiles.length; i++) {
                (function(file, index) {
                  var p = new Promise(function(resolve, reject) {
                    var reader = new FileReader();
                    reader.onload = function(e) { 
                      console.log('File', index, 'loaded:', file.name, 'Length:', e.target.result.length);
                      resolve(e.target.result); 
                    };
                    reader.onerror = function() { 
                      console.error('Failed to read:', file.name);
                      reject(new Error('Failed to read: ' + file.name)); 
                    };
                    reader.readAsText(file);
                  });
                  readers.push(p);
                })(selectedFiles[i], i);
              }
              
              Promise.all(readers)
                .then(function(contents) {
                  console.log('File contents loaded:', contents.length, 'files');
                  
                  // Combine filenames with contents
                  var contentsWithNames = contents.map(function(content, idx) {
                    console.log('File', idx, ':', selectedFiles[idx].name, 'Content length:', content ? content.length : 'NULL');
                    return {
                      name: selectedFiles[idx].name,
                      content: content,
                      csv: content  // Add this as backup
                    };
                  });
                  
                  console.log('Sending to backend:', contentsWithNames.length, 'files');
                  
                  return new Promise(function(resolve, reject) {
                    google.script.run
                      .withSuccessHandler(resolve)
                      .withFailureHandler(reject)
                      .processCSVData(contentsWithNames);
                  });
                })
                .then(function(result) {
                  // Check if CSV Configurator is being shown
                  if (result && result.showingConfigurator) {
                    status.className = 'success show';
                    status.innerHTML = '<strong>⚙️ Configuration Required</strong><br>Please configure the CSV processing in the popup window.';
                    setTimeout(function() { google.script.host.close(); }, 1000);
                    return null; // Don't continue the chain
                  }
                  
                  // Check if we need labeling (duplicates detected) - old flow
                  if (result && result.hasDuplicates) {
                    // Don't call generateReport - modal will handle it
                    status.className = 'success show';
                    status.innerHTML = '<strong>⚠️ Action Required</strong><br>Please provide labels for duplicate products in the popup window.';
                    google.script.host.close();
                    return null; // Don't continue the chain
                  }
                  
                  // No duplicates - continue with report generation
                  return new Promise(function(resolve, reject) {
                    google.script.run
                      .withSuccessHandler(resolve)
                      .withFailureHandler(reject)
                      .generateReport();
                  });
                })
                .then(function(result) {
                  if (result === null) return; // Skip if configurator/modal was shown
                  
                  if (result && result.success) {
                    status.className = 'success show';
                    status.innerHTML = '<strong>✅ Success!</strong><br>' + result.message.replace(/\\n/g, '<br>');
                    setTimeout(function() { google.script.host.close(); }, 2000);
                  } else {
                    throw new Error(result ? result.message : 'Unknown error');
                  }
                })
                .catch(function(error) {
                  uploadBtn.disabled = false;
                  uploadBtn.textContent = 'Upload & Generate Report';
                  status.className = 'error show';
                  status.innerHTML = '<strong>❌ Error</strong><br>' + (error.message || error);
                });
            }
          })();
        </script>
      </body>
    </html>
  `).setWidth(520).setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload CSV Files');
}

function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) {
    throw new Error('Settings sheet not found. Please create a "Settings" sheet.');
  }
  
  // Get local settings (always needed for topN, weekLabel, etc.)
  const localSettings = {
    slackWebhook: sheet.getRange('B2').getValue() || '',
    topN: parseInt(sheet.getRange('B3').getValue()) || 10,
    weekLabel: sheet.getRange('B4').getValue() || '',
    spreadsheetUrl: sheet.getRange('B9').getValue() || ''
  };
  
  // Try to get Slack config from central Config (preferred)
  const centralSlack = getSlackConfigFromConfig_();
  
  if (centralSlack && centralSlack.botToken) {
    // Use central Config for Slack settings
    return {
      ...localSettings,
      slackBotToken: centralSlack.botToken,
      slackChannelIdWeekly: centralSlack.weeklyChannelId,
      slackChannelIdCreative: centralSlack.creativeChannelId
    };
  } else {
    // Fall back to local Settings sheet
    return {
      ...localSettings,
      slackBotToken: sheet.getRange('B6').getValue() || '',
      slackChannelIdWeekly: sheet.getRange('B7').getValue() || '',
      slackChannelIdCreative: sheet.getRange('B8').getValue() || ''
    };
  }
}

/**
 * Get product spreadsheet IDs from Settings sheet
 * Returns object like: { "CamTrix (Main)": "spreadsheet_id", "AliveBlue": "spreadsheet_id" }
 */
function getProductSpreadsheetIds() {
  // Try central Config first
  const centralSpreadsheets = getProductSpreadsheetsFromConfig_();
  
  if (centralSpreadsheets && Object.keys(centralSpreadsheets).length > 0) {
    Logger.log('📋 Using ' + Object.keys(centralSpreadsheets).length + ' product spreadsheet mappings from central Config');
    return centralSpreadsheets;
  }
  
  // Fall back to local Settings sheet
  Logger.log('📋 Falling back to local Settings sheet for product spreadsheet mappings');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) {
    return {};
  }
  
  const productSpreadsheets = {};
  
  // Start from row 11 (after "PRODUCT SPREADSHEET IDs" header in row 10)
  let row = 11;
  while (true) {
    const productName = sheet.getRange(row, 1).getValue();
    const spreadsheetId = sheet.getRange(row, 2).getValue();
    
    // Stop if both cells are empty
    if (!productName && !spreadsheetId) {
      break;
    }
    
    // Only add if both have values
    if (productName && spreadsheetId) {
      productSpreadsheets[productName.toString().trim()] = spreadsheetId.toString().trim();
    }
    
    row++;
    
    // Safety limit
    if (row > 110) {
      break;
    }
  }
  
  Logger.log('Loaded ' + Object.keys(productSpreadsheets).length + ' product spreadsheet mappings from local Settings');
  
  return productSpreadsheets;
}

/**
 * Get top ads list for showing progress
 */
function getTopAdsForProgress() {
  initializeConfigConstants_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet) return [];
  
  const data = dataSheet.getDataRange().getValues();
  const settings = getSettings();
  const report = processAtriaData(data, settings);
  
  return report.topAds.map(ad => {
    let name = ad.adName;
    if (name.includes('||')) {
      name = name.split('||')[0].trim();
    }
    return cleanAdName(name);
  });
}

// ============================================
// PRODUCT LABELING SYSTEM FOR DUPLICATE PRODUCTS
// ============================================

/**
 * ✅ FINAL FIXED VERSION - Meta status prefixes cleaned in correct order
 * Updated processCSVData with duplicate product detection
 * Returns an object with detected duplicates if found
 */
function processCSVData(csvContents, productLabels = null) {
  try {
    // ✅ CRITICAL: Initialize config constants FIRST (needed for product validation)
    initializeConfigConstants_();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dataSheet = ss.getSheetByName('CSV Data');
    
    if (!dataSheet) {
      dataSheet = ss.insertSheet('CSV Data');
    }
    
    // ✅ PRESERVE existing labels from Column I
    const existingLabels = {}; // Map: adName → label
    const existingCsvSources = {}; // Map: adName → csvSource
    
    if (dataSheet.getLastRow() > 0) {
      try {
        const existingData = dataSheet.getDataRange().getValues();
        Logger.log(`📌 Reading ${existingData.length - 1} existing rows to preserve labels...`);
        
        for (let i = 1; i < existingData.length; i++) {
          const adName = String(existingData[i][0] || '').split('||')[0].trim();
          const existingLabel = existingData[i][8] || ''; // Column I (Product Label)
          const existingSource = existingData[i][9] || ''; // Column J (CSV Source)
          
          if (adName && existingLabel) {
            existingLabels[adName] = existingLabel;
            existingCsvSources[adName] = existingSource;
            Logger.log(`  Preserved: ${adName.substring(0, 40)} → "${existingLabel}" (source: ${existingSource})`);
          }
        }
        
        Logger.log(`✅ Preserved ${Object.keys(existingLabels).length} existing labels`);
      } catch (e) {
        Logger.log('⚠️ Could not read existing labels (sheet might be empty): ' + e.toString());
      }
    }
    
    // NOW clear the sheet
    dataSheet.clear();
    
// ============================================
// 🔥 PHASE 1: CHECK IF CSV CONFIGURATOR NEEDED
// Shows unified configurator for mixed products, duplicates, or multiple files
// ============================================
Logger.log('=== PHASE 1: Checking if CSV Configurator needed ===');

// Quick analysis to determine if configurator should be shown
const analysisResult = analyzeCSVFilesForConfigurator(csvContents);
const hasMixed = analysisResult.csvAnalysis.some(csv => csv.isMixed);
const hasDuplicates = analysisResult.csvAnalysis.some(csv => csv.hasDuplicates);
const hasUnknownProducts = analysisResult.unknownProducts && Object.keys(analysisResult.unknownProducts).length > 0;
const multipleComplexFiles = csvContents.length > 2;

if (hasMixed || hasDuplicates || hasUnknownProducts) {
  Logger.log('⚙️ CSV Configurator needed - showing modal...');
  Logger.log('  - Has mixed products: ' + hasMixed);
  Logger.log('  - Has duplicates: ' + hasDuplicates);
  Logger.log('  - Has unknown products: ' + hasUnknownProducts);
  Logger.log('  - Multiple files: ' + multipleComplexFiles);
  
  // Show the new unified CSV Configurator
  showCsvConfiguratorModal(csvContents);
  
  return {
    showingConfigurator: true,
    message: 'Please configure CSV processing in the modal window.'
  };
}

Logger.log('✓ No configurator needed - processing normally');

    // ============================================
    // 🔥 PHASE 2: PROCESS ROWS WITH CORRECT LABELS
    // ============================================
    const allRows = [];
    const targetColumns = ['Ad name', 'Spend', 'ROAS', 'CPA', 'Purchases', 'AOV', 'Hook rate', 'Hold rate', 'Product Label', 'CSV Source'];
    
    // Add header row
    allRows.push(targetColumns);
    
    let totalAdsLoaded = 0;
    
    csvContents.forEach((csvObj, fileIndex) => {
      const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
const csvFileName = typeof csvObj === 'string' ? `File_${fileIndex + 1}` : (csvObj.fileName || csvObj.name);
      
      Logger.log('=== Processing CSV file ' + (fileIndex + 1) + ' ===');
      
      // Parse this CSV - handle both comma and semicolon delimiters
let csvData;
if (csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length) {
  // Semicolon-delimited CSV
  Logger.log('  Detected semicolon delimiter');
  csvData = csv.split('\n').map(line => line.split(';'));
} else {
  // Comma-delimited CSV
  csvData = Utilities.parseCsv(csv);
}
      
      if (csvData.length === 0) return;
      
      const headers = csvData[0];
      Logger.log('Headers: ' + JSON.stringify(headers));
      
      // Find column indices for this specific CSV
      const findCol = (name) => {
        const searchName = name.toLowerCase().trim();
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] && headers[i].toString().toLowerCase().trim() === searchName) {
            return i;
          }
        }
        return -1;
      };
      
      const adNameCol = findCol('ad name');
      const spendCol = findCol('spend');
      const roasCol = findCol('roas');
      const cpaCol = findCol('cpa');
      const purchasesCol = findCol('purchases');
      const aovCol = findCol('aov');
      const hookRateCol = findCol('hook rate');
      const holdRateCol = findCol('hold rate');
      
      Logger.log('Column mapping for file ' + (fileIndex + 1) + ':');
      Logger.log('  Ad name: ' + adNameCol);
      Logger.log('  Spend: ' + spendCol);
      Logger.log('  ROAS: ' + roasCol);
      Logger.log('  CPA: ' + cpaCol);
      Logger.log('  Purchases: ' + purchasesCol);
      Logger.log('  AOV: ' + aovCol);
      Logger.log('  Hook rate: ' + hookRateCol);
      Logger.log('  Hold rate: ' + holdRateCol);
      
      
      // Extract only the columns we need from each row
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        
        if (!row[adNameCol] || row[adNameCol].toString().trim() === '') continue;
        
        // Extract ad name and detect product
        const adName = row[adNameCol] || '';
        let cleanAdName = adName;
        if (cleanAdName.includes('||')) {
          const parts = cleanAdName.split('||');
          for (let j = 0; j < parts.length; j++) {
            if (parts[j].includes('_')) {
              cleanAdName = parts[j].trim();
              break;
            }
          }
        }
        
        // Remove "Test - " prefix if present
        if (cleanAdName.startsWith('Test - ')) {
          cleanAdName = cleanAdName.substring(7);
        }
        
        // ✅ CLEAN the ad name before product detection - CORRECT ORDER
        let adNameForDetection = cleanAdName;

        // ✅ STEP 1: Remove leading brackets ([unititled], [untitled], etc.)
        adNameForDetection = adNameForDetection.replace(/^\[[^\]]*\]\s*/g, '');
        
        // ✅ STEP 2: Remove parentheses prefixes: (blank) rejected:, (anything):
        adNameForDetection = adNameForDetection.replace(/^\([^)]*\)\s*/g, '');
        
        // ✅ STEP 3: Remove status word prefixes: rejected:, not delivering:, etc.
        adNameForDetection = adNameForDetection.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
        
        // ✅ STEP 4: Remove leading brackets AGAIN ([copy:gift], [copy:longlasting], [old copy], etc.)
        adNameForDetection = adNameForDetection.replace(/^\[[^\]]*\]\s*/g, '');

        // ✅ STEP 5: Remove "Post ID || " if present
        if (adNameForDetection.startsWith('Post ID || ')) {
          adNameForDetection = adNameForDetection.substring(11);
        }

        // NOW detect product from the CLEANED name
        const productMatch = adNameForDetection.match(/([A-Za-z0-9]+)_(FB|YT|GDN|TT)_/);
        let product = '';
        if (productMatch) {
          product = normalizeProductName(productMatch[1]);
        }
        
        // Convert hook/hold rates from decimals to percentages
        let hookValue = '';
        let holdValue = '';

        if (hookRateCol !== -1 && row[hookRateCol]) {
          const hookNum = parseFloat(row[hookRateCol]);
          if (!isNaN(hookNum)) {
            hookValue = hookNum < 1 ? (hookNum * 100) : hookNum;
          }
        }

        if (holdRateCol !== -1 && row[holdRateCol]) {
          const holdNum = parseFloat(row[holdRateCol]);
          if (!isNaN(holdNum)) {
            holdValue = holdNum < 1 ? (holdNum * 100) : holdNum;
          }
        }
        
        // ============================================
        // 🎯 LABEL ASSIGNMENT LOGIC
        // ============================================
        const cleanAdNameForLookup = cleanAdName.split('||')[0].trim();
        let productLabel = '';
        let csvSource = fileIndex;
        
        // Priority 1: Restore existing label if this ad had one before
        if (existingLabels[cleanAdNameForLookup]) {
          productLabel = existingLabels[cleanAdNameForLookup];
          csvSource = existingCsvSources[cleanAdNameForLookup] || fileIndex;
          Logger.log(`  🔄 Restored label for ${cleanAdNameForLookup.substring(0, 40)}: "${productLabel}"`);
        } 
        // Priority 2: Apply new labels from modal
        else if (productLabels && product && productLabels[product]) {
          // Get the array of file indices for this product
          const duplicateFileIndexes = detectedProducts[product].map(entry => entry.fileIndex);
          
          // Find THIS file's position in the duplicate array
          const positionInDuplicates = duplicateFileIndexes.indexOf(fileIndex);
          
          Logger.log(`  📍 Product: ${product}, File index: ${fileIndex}, Position in duplicates: ${positionInDuplicates}, Available labels: ${JSON.stringify(productLabels[product])}`);
          
          if (positionInDuplicates !== -1 && productLabels[product][positionInDuplicates]) {
            productLabel = productLabels[product][positionInDuplicates];
            Logger.log(`  🆕 Applied label for ${cleanAdNameForLookup.substring(0, 40)}: "${productLabel}"`);
          }
        }
        
        // ✅ SAFE EXTRACTION: Only extract if column was found (index >= 0)
        // This handles CSVs with different column structures (e.g., some have Status column)
        const safeGet = (colIdx) => colIdx >= 0 ? (row[colIdx] || '') : '';
        
        const extractedRow = [
          safeGet(adNameCol),
          safeGet(spendCol),
          safeGet(roasCol),
          safeGet(cpaCol),
          safeGet(purchasesCol),
          safeGet(aovCol),
          hookValue,
          holdValue,
          productLabel,
          csvSource
        ];
        
        // ✅ DEBUG: Log if spend looks wrong (contains non-numeric characters that aren't currency)
        const spendVal = extractedRow[1];
        if (spendVal && !/^[\d€$,.\s-]+$/.test(spendVal.toString())) {
          Logger.log('⚠️ SUSPICIOUS SPEND VALUE: "' + spendVal + '" for ad: ' + extractedRow[0].substring(0, 50));
          Logger.log('   Headers were: ' + JSON.stringify(headers));
          Logger.log('   Spend col index: ' + spendCol);
        }
        
        allRows.push(extractedRow);
        totalAdsLoaded++;
      }
      
      Logger.log('Extracted ' + (csvData.length - 1) + ' rows from file ' + (fileIndex + 1));
    });
    
    // ============================================
    // WRITE DATA TO SHEET
    // ============================================
    if (allRows.length > 1) {
      dataSheet.getRange(1, 1, allRows.length, targetColumns.length).setValues(allRows);
      Logger.log(`✅ Wrote ${allRows.length - 1} rows to CSV Data sheet`);
      
      // Set proper number formats for all columns
      const dataRange = dataSheet.getRange(2, 1, allRows.length - 1, targetColumns.length);
      
      // Column A: Ad name - plain text
      dataSheet.getRange(2, 1, allRows.length - 1, 1).setNumberFormat('@STRING@');
      
      // Column B: Spend - currency (€)
      dataSheet.getRange(2, 2, allRows.length - 1, 1).setNumberFormat('€#,##0.00');
      
      // Column C: ROAS - plain number with 2 decimals
      dataSheet.getRange(2, 3, allRows.length - 1, 1).setNumberFormat('0.00');
      
      // Column D: CPA - currency (€)
      dataSheet.getRange(2, 4, allRows.length - 1, 1).setNumberFormat('€#,##0.00');
      
      // Column E: Purchases - whole number
      dataSheet.getRange(2, 5, allRows.length - 1, 1).setNumberFormat('0');
      
      // Column F: AOV - currency (€)
      dataSheet.getRange(2, 6, allRows.length - 1, 1).setNumberFormat('€#,##0.00');
      
      // Column G: Hook rate - percentage with 2 decimals
      dataSheet.getRange(2, 7, allRows.length - 1, 1).setNumberFormat('0.00"%"');
      
      // Column H: Hold rate - percentage with 2 decimals
      dataSheet.getRange(2, 8, allRows.length - 1, 1).setNumberFormat('0.00"%"');
      
      // Column I: Product Label - plain text
      dataSheet.getRange(2, 9, allRows.length - 1, 1).setNumberFormat('@STRING@');
      
      // Column J: CSV Source - whole number (no decimals, no %)
      dataSheet.getRange(2, 10, allRows.length - 1, 1).setNumberFormat('0');
      
      Logger.log(`✅ Applied proper number formats to all columns`);
    }

    dataSheet.hideSheet();
    
    // Build product list for Data Info
    const uniqueProducts = new Set();
    Object.keys(detectedProducts).forEach(product => {
      if (productLabels && productLabels[product]) {
        // Add labeled versions
        productLabels[product].forEach((label, idx) => {
          if (label) {
            uniqueProducts.add(`${product} (${label})`);
          } else {
            uniqueProducts.add(product);
          }
        });
      } else {
        uniqueProducts.add(product);
      }
    });
    
    // Update Data Info sheet
    updateDataInfo(totalAdsLoaded, Array.from(uniqueProducts).sort());
    
    return {
      hasDuplicates: false,
      message: `Successfully loaded ${totalAdsLoaded} ads from ${csvContents.length} file(s).`
    };
    
  } catch (error) {
    Logger.log('processCSVData error: ' + error.toString());
    throw new Error('Failed to process CSV: ' + error.toString());
  }
}

/**
 * Store product labels in Script Properties for later use
 */
function storeProductLabels(labels) {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('PRODUCT_LABELS', JSON.stringify(labels));
  Logger.log('✅ Stored product labels: ' + JSON.stringify(labels));
}

/**
 * Retrieve stored product labels
 */
function getStoredProductLabels() {
  const scriptProps = PropertiesService.getScriptProperties();
  const labelsJson = scriptProps.getProperty('PRODUCT_LABELS');
  if (labelsJson) {
    return JSON.parse(labelsJson);
  }
  return null;
}

/**
 * Clear stored product labels
 */
function clearProductLabels() {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.deleteProperty('PRODUCT_LABELS');
  Logger.log('✅ Cleared product labels');
}

/**
 * Show modal for product differentiation
 */
function showProductDifferentiationModal(duplicateInfo) {
  Logger.log('=== SHOWING MODAL ===');
  Logger.log('duplicateInfo: ' + JSON.stringify(duplicateInfo));
  
  const html = HtmlService.createTemplateFromFile('product_differentiation_modal');
  html.duplicateInfo = duplicateInfo;
  
  const htmlOutput = html.evaluate()
    .setWidth(750)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Product Differentiation');
}

/**
 * Show modal for mixed products
 */
function showMixedProductsModal(mixedCsvs) {
  Logger.log('=== SHOWING MIXED PRODUCTS MODAL ===');
  Logger.log('mixedCsvs: ' + JSON.stringify(mixedCsvs));
  
  // ✅ FIX: Properly escape data for embedding in JavaScript
  const mixedCsvsJson = JSON.stringify(mixedCsvs)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r');
  
  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: Arial, sans-serif;
      padding: 20px;
      background: #f5f5f5;
      margin: 0;
    }
    .container {
      max-width: 700px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    h2 {
      color: #e67e22;
      margin-bottom: 10px;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .warning-icon {
      font-size: 28px;
    }
    .description {
      color: #666;
      margin-bottom: 25px;
      font-size: 14px;
      line-height: 1.5;
    }
    .csv-block {
      background: #fff3cd;
      border-left: 4px solid #e67e22;
      padding: 20px;
      margin: 15px 0;
      border-radius: 6px;
    }
    .csv-name {
      font-weight: bold;
      font-size: 15px;
      margin-bottom: 15px;
      color: #333;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .product-line {
      padding: 8px 0;
      font-size: 14px;
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .majority {
      color: #27ae60;
      font-weight: bold;
    }
    .minority {
      color: #e74c3c;
      padding-left: 10px;
    }
    .csv-options {
      margin-top: 15px;
      padding: 15px;
      background: white;
      border-radius: 6px;
      border: 2px solid #dee2e6;
    }
    .csv-options h4 {
      margin: 0 0 12px 0;
      font-size: 14px;
      color: #495057;
    }
    .radio-option {
      display: flex;
      align-items: flex-start;
      padding: 10px;
      margin: 6px 0;
      border-radius: 4px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .radio-option:hover {
      background: rgba(52, 152, 219, 0.1);
    }
    .radio-option input[type="radio"] {
      margin-right: 10px;
      margin-top: 3px;
      cursor: pointer;
      width: 16px;
      height: 16px;
    }
    .radio-option label {
      cursor: pointer;
      flex: 1;
      font-size: 13px;
    }
    .radio-option .option-title {
      font-weight: 600;
      color: #2c3e50;
      margin-bottom: 3px;
    }
    .radio-option .option-desc {
      font-size: 12px;
      color: #6c757d;
      line-height: 1.4;
    }
    .buttons {
      display: flex;
      gap: 10px;
      margin-top: 25px;
    }
    button {
      flex: 1;
      padding: 14px 24px;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 15px;
      font-weight: 600;
      transition: all 0.2s;
    }
    #continueBtn {
      background: #3498db;
      color: white;
    }
    #continueBtn:hover:not(:disabled) {
      background: #2980b9;
    }
    #continueBtn:disabled {
      background: #bdc3c7;
      cursor: not-allowed;
      opacity: 0.6;
    }
    #cancelBtn {
      background: #95a5a6;
      color: white;
    }
    #cancelBtn:hover {
      background: #7f8c8d;
    }
    #status {
      margin-top: 20px;
      padding: 15px;
      border-radius: 6px;
      display: none;
      text-align: center;
    }
    #status.show { display: block; }
    #status.success {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
    }
    #status.error {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>
      <span class="warning-icon">⚠️</span>
      Mixed Products Detected
    </h2>
    
    <div class="description">
      The following CSV file(s) contain multiple products. This usually happens when ads were launched on the wrong account or during testing.
    </div>
    
    <div id="csvList"></div>
    
    <div class="buttons">
      <button id="cancelBtn">Cancel</button>
      <button id="continueBtn" disabled>Continue</button>
    </div>
    
    <div id="status"></div>
  </div>
  
  <script>
    const mixedCsvs = JSON.parse('${mixedCsvsJson}');
    const decisions = {};
    
    console.log('Mixed CSVs:', mixedCsvs);
    
    const csvListDiv = document.getElementById('csvList');
    
    mixedCsvs.forEach((csv, csvIndex) => {
      const block = document.createElement('div');
      block.className = 'csv-block';
      
      let html = '<div class="csv-name">📄 ' + csv.fileName + '</div>';
      html += '<div class="product-line majority">✓ ' + csv.majorityProduct + ': ' + csv.majorityCount + ' ads (' + csv.majorityPercent + '%)</div>';
      
      csv.minorityProducts.forEach(minority => {
        html += '<div class="product-line minority">⚠️ ' + minority.name + ': ' + minority.count + ' ads (' + minority.percent + '%)</div>';
      });
      
      html += '<div class="csv-options">';
      html += '<h4>What to do with this file?</h4>';
      
      const radioName = 'action_' + csv.fileIndex;
      
      html += '<div class="radio-option">';
      html += '<input type="radio" id="' + radioName + '_ignore" name="' + radioName + '" value="ignore" data-file-index="' + csv.fileIndex + '">';
      html += '<label for="' + radioName + '_ignore">';
      html += '<div class="option-title">🗑️ Ignore minority products</div>';
      html += '<div class="option-desc">Exclude ' + csv.minorityProducts.map(m => m.name).join(', ') + ' ads (recommended)</div>';
      html += '</label>';
      html += '</div>';
      
      html += '<div class="radio-option">';
      html += '<input type="radio" id="' + radioName + '_convert" name="' + radioName + '" value="convert" data-file-index="' + csv.fileIndex + '">';
      html += '<label for="' + radioName + '_convert">';
      html += '<div class="option-title">🔄 Treat as ' + csv.majorityProduct + '</div>';
      html += '<div class="option-desc">Rename all ads to ' + csv.majorityProduct + ' (use with caution)</div>';
      html += '</label>';
      html += '</div>';
      
      html += '<div class="radio-option">';
      html += '<input type="radio" id="' + radioName + '_cancel" name="' + radioName + '" value="cancel" data-file-index="' + csv.fileIndex + '">';
      html += '<label for="' + radioName + '_cancel">';
      html += '<div class="option-title">📋 Review this CSV first</div>';
      html += '<div class="option-desc">Skip this file and review manually</div>';
      html += '</label>';
      html += '</div>';
      
      html += '</div>';
      
      block.innerHTML = html;
      csvListDiv.appendChild(block);
    });
    
    const allRadios = document.querySelectorAll('input[type="radio"]');
    const continueBtn = document.getElementById('continueBtn');
    const cancelBtn = document.getElementById('cancelBtn');
    const status = document.getElementById('status');
    
    console.log('Total radio buttons found:', allRadios.length);
    
    allRadios.forEach(radio => {
      radio.addEventListener('change', function() {
        const fileIndex = this.getAttribute('data-file-index');
        const action = this.value;
        
        console.log('File ' + fileIndex + ': selected ' + action);
        
        decisions[fileIndex] = action;
        
        const allDecided = mixedCsvs.every(csv => decisions[csv.fileIndex] !== undefined);
        
        console.log('All decided?', allDecided, 'Decisions:', decisions);
        
        continueBtn.disabled = !allDecided;
      });
    });
    
    cancelBtn.addEventListener('click', () => {
      google.script.host.close();
    });
    
    continueBtn.addEventListener('click', () => {
      console.log('Continue clicked with decisions:', decisions);
      
      const hasCancel = Object.values(decisions).includes('cancel');
      
      if (hasCancel) {
        status.className = 'show';
        status.innerHTML = '⚠️ Upload cancelled. Please review your CSV files.';
        setTimeout(() => {
          google.script.host.close();
        }, 1500);
        return;
      }
      
      continueBtn.disabled = true;
      continueBtn.textContent = 'Processing...';
      status.className = '';
      
      const fullDecisions = {};
      mixedCsvs.forEach(csv => {
        fullDecisions[csv.fileIndex] = {
          action: decisions[csv.fileIndex],
          majorityProduct: csv.majorityProduct,
          minorityProducts: csv.minorityProducts.map(m => m.name)
        };
      });
      
      console.log('Sending to backend:', fullDecisions);
      
      google.script.run
        .withSuccessHandler(result => {
          console.log('Result:', result);
          if (result && result.success) {
            status.className = 'success show';
            status.innerHTML = '✅ ' + result.message;
            setTimeout(() => {
              google.script.host.close();
            }, 1500);
          } else {
            status.className = 'error show';
            status.innerHTML = '❌ ' + (result ? result.message : 'Unknown error');
            continueBtn.disabled = false;
            continueBtn.textContent = 'Continue';
          }
        })
        .withFailureHandler(error => {
          console.error('Error:', error);
          status.className = 'error show';
          status.innerHTML = '❌ Error: ' + error.message;
          continueBtn.disabled = false;
          continueBtn.textContent = 'Continue';
        })
        .applyMixedProductDecisions(fullDecisions);
    });
  </script>
</body>
</html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(750)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Mixed Products Detected');
}

/**
 * Called from the modal when user confirms labels
 */
function applyProductLabels(labels) {
  Logger.log('Received product labels: ' + JSON.stringify(labels));
  
  // Store labels for use during analysis
  storeProductLabels(labels);
  
  // ✅ GET CSV FROM CACHE instead of Script Properties
  const cache = CacheService.getScriptCache();
  const chunksCount = parseInt(cache.get('CSV_CHUNKS_COUNT') || '0');
  
  if (chunksCount === 0) {
    throw new Error('CSV contents not found in cache. Please re-upload files.');
  }
  
  // Reconstruct CSV from chunks
  let csvJson = '';
  for (let i = 0; i < chunksCount; i++) {
    const chunk = cache.get('CSV_CHUNK_' + i);
    if (!chunk) {
      throw new Error('CSV chunk ' + i + ' not found. Please re-upload files.');
    }
    csvJson += chunk;
  }
  
  const csvContents = JSON.parse(csvJson);
  
  // Reprocess CSV with labels
  const result = processCSVData(csvContents, labels);
  
  // Clear cache
  cache.remove('CSV_CHUNKS_COUNT');
  for (let i = 0; i < chunksCount; i++) {
    cache.remove('CSV_CHUNK_' + i);
  }
  
  // Auto-generate report after labeling
  if (result && !result.hasDuplicates) {
    const reportResult = generateReport();
    return reportResult;
  }
  
  return result;
}

/**
 * Apply user's decision about mixed products
 */
function applyMixedProductDecisions(decisions) {
  Logger.log('Received mixed product decisions: ' + JSON.stringify(decisions));
  
  // ✅ GET CSV FROM CACHE instead of Script Properties
  const cache = CacheService.getScriptCache();
  const chunksCount = parseInt(cache.get('CSV_CHUNKS_COUNT') || '0');
  
  if (chunksCount === 0) {
    throw new Error('CSV contents not found in cache. Please re-upload files.');
  }
  
  // Reconstruct CSV from chunks
  let csvJson = '';
  for (let i = 0; i < chunksCount; i++) {
    const chunk = cache.get('CSV_CHUNK_' + i);
    if (!chunk) {
      throw new Error('CSV chunk ' + i + ' not found. Please re-upload files.');
    }
    csvJson += chunk;
  }
  
  let csvContents = JSON.parse(csvJson);
  
  // Process each CSV based on user decision
  Object.keys(decisions).forEach(fileIndexStr => {
    const fileIndex = parseInt(fileIndexStr);
    const decision = decisions[fileIndexStr];
    const action = decision.action;
    
    Logger.log(`Processing file ${fileIndex}: action = ${action}`);
    
    if (action === 'ignore') {
      csvContents[fileIndex] = filterMinorityProducts(csvContents[fileIndex], decision.minorityProducts);
    } else if (action === 'convert') {
      csvContents[fileIndex] = convertMinorityProducts(csvContents[fileIndex], decision.minorityProducts, decision.majorityProduct);
    }
  });
  
  // Now process CSV normally
  const result = processCSVData(csvContents, null);
  
  // Clear cache if successful
  if (result && !result.hasDuplicates && !result.hasMixedProducts) {
    cache.remove('CSV_CHUNKS_COUNT');
    for (let i = 0; i < chunksCount; i++) {
      cache.remove('CSV_CHUNK_' + i);
    }
    
    // Auto-generate report
    const reportResult = generateReport();
    return reportResult;
  }
  
  return result;
}

/**
 * Filter out minority product ads from CSV
 */
function filterMinorityProducts(csvObj, minorityProducts) {
  Logger.log(`\nFiltering minority products from ${csvObj.fileName}`);
  Logger.log(`Minority products to remove: ${minorityProducts.join(', ')}`);
  
  // Create lowercase set for comparison
  const minorityProductsLower = minorityProducts.map(p => p.toLowerCase());
  
  // Parse CSV
  let csvData;
  if (csvObj.csv.includes(';') && csvObj.csv.split('\n')[0].split(';').length > csvObj.csv.split('\n')[0].split(',').length) {
    csvData = csvObj.csv.split('\n').map(line => line.split(';'));
  } else {
    csvData = Utilities.parseCsv(csvObj.csv);
  }
  
  const headers = csvData[0];
  const adNameIndex = headers.indexOf('Ad name');
  
  if (adNameIndex === -1) {
    Logger.log('  No "Ad name" column found');
    return csvObj;
  }
  
  // Filter rows
  const filteredRows = [headers]; // Keep headers
  let removedCount = 0;
  
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row || row.length === 0 || !row[adNameIndex]) {
      filteredRows.push(row);
      continue;
    }
    
    const adName = row[adNameIndex].trim();
    const product = extractProductFromAdName(adName);
    
    // Keep row if it's NOT a minority product (case-insensitive)
    if (!minorityProductsLower.includes(product.toLowerCase())) {
      filteredRows.push(row);
    } else {
      removedCount++;
      Logger.log(`  Removed: ${adName} (${product})`);
    }
  }
  
  Logger.log(`Removed ${removedCount} ads`);
  
  // Convert back to CSV string
  const filteredCsv = filteredRows.map(row => row.join(',')).join('\n');
  
  return {
    fileName: csvObj.fileName,
    csv: filteredCsv
  };
}

/**
 * Convert minority product names to majority product
 */
function convertMinorityProducts(csvObj, minorityProducts, majorityProduct) {
  Logger.log(`\nConverting minority products to ${majorityProduct} in ${csvObj.fileName}`);
  Logger.log(`Minority products to convert: ${minorityProducts.join(', ')}`);
  
  // Create lowercase set for comparison
  const minorityProductsLower = minorityProducts.map(p => p.toLowerCase());
  
  // Parse CSV
  let csvData;
  if (csvObj.csv.includes(';') && csvObj.csv.split('\n')[0].split(';').length > csvObj.csv.split('\n')[0].split(',').length) {
    csvData = csvObj.csv.split('\n').map(line => line.split(';'));
  } else {
    csvData = Utilities.parseCsv(csvObj.csv);
  }
  
  const headers = csvData[0];
  const adNameIndex = headers.indexOf('Ad name');
  
  if (adNameIndex === -1) {
    Logger.log('  No "Ad name" column found');
    return csvObj;
  }
  
  let convertedCount = 0;
  
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row || row.length === 0 || !row[adNameIndex]) continue;
    
    const adName = row[adNameIndex].trim();
    const product = extractProductFromAdName(adName);
    
    // Convert if it's a minority product (case-insensitive)
    if (minorityProductsLower.includes(product.toLowerCase())) {
      // Replace product name in ad name (case-insensitive)
      const newAdName = adName.replace(new RegExp(product, 'gi'), majorityProduct);
      row[adNameIndex] = newAdName;
      convertedCount++;
      Logger.log(`  Converted: ${adName} → ${newAdName}`);
    }
  }
  
  Logger.log(`Converted ${convertedCount} ads`);
  
  // Convert back to CSV string
  const convertedCsv = csvData.map(row => row.join(',')).join('\n');
  
  return {
    fileName: csvObj.fileName,
    csv: convertedCsv
  };
}

/**
 * Updated uploadCSVFiles function to handle duplicates
 */
function uploadCSVFiles() {
  const html = HtmlService.createHtmlOutputFromFile('csv_upload')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload CSV Files');
}

/**
 * Updated handleCSVUpload function
 */
function handleCSVUpload(csvContents) {
  try {
    Logger.log('Starting CSV upload with ' + csvContents.length + ' files...');
    
    // Clear any old product labels
    clearProductLabels();
    
    // Process CSV data - this will show configurator if needed
    const result = processCSVData(csvContents);
    
    // New configurator is being shown - wait for user action
    if (result.showingConfigurator) {
      Logger.log('CSV Configurator shown, waiting for user action...');
      return {
        success: true,
        showingConfigurator: true,
        message: result.message
      };
    }
    
    // Old duplicate handling (kept for backward compatibility)
    if (result.hasDuplicates) {
      Logger.log('Duplicates detected, showing modal...');
      
      // Store CSV contents temporarily for reprocessing
      const scriptProps = PropertiesService.getScriptProperties();
      scriptProps.setProperty('PENDING_CSV_CONTENTS', JSON.stringify(csvContents));
      
      // Show differentiation modal
      showProductDifferentiationModal(result.duplicateInfo);
      
      return {
        success: true,
        needsLabeling: true,
        duplicateInfo: result.duplicateInfo
      };
    }
    
    return {
      success: true,
      needsLabeling: false,
      message: result.message
    };
    
  } catch (error) {
    logErrorToHub_('CSV Upload', error.message || String(error));
    Logger.log('Upload error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Helper to get product with label from ad name
 */
function getProductWithLabel(adName) {
  let cleanAdName = adName;
  if (cleanAdName.includes('||')) {
    const parts = cleanAdName.split('||');
    for (let j = 0; j < parts.length; j++) {
      if (parts[j].includes('_')) {
        cleanAdName = parts[j].trim();
        break;
      }
    }
  }
  
  const productMatch = cleanAdName.match(/([A-Za-z0-9]+)_(FB|YT|GDN|TT)_/);
  if (!productMatch) return null;
  
  let product = normalizeProductName(productMatch[1]);
  
  // Check if we have a label for this product
  const labels = getStoredProductLabels();
  if (!labels || !labels[product]) {
    return product;
  }
  
  // We need to determine which file this ad came from
  // This would be stored in the CSV Data sheet
  // For now, return base product name
  return product;
}

function processAtriaData(data, settings) {
  const headers = data[0];
  
  Logger.log('=== processAtriaData called ===');
  Logger.log('Headers: ' + JSON.stringify(headers));
  
  // Now headers are already normalized: ["Ad name", "Spend", "ROAS", "CPA", "Purchases", "AOV"]
  const adNameCol = 0;
  const spendCol = 1;
  const roasCol = 2;
  const cpaCol = 3;
  const purchasesCol = 4;
  const aovCol = 5;
  
  const rows = data.slice(1).filter(row => {
    if (!row[adNameCol] || row[adNameCol].toString().trim() === '') return false;
    
    const adName = row[adNameCol].toString();
    
    // Skip rows that don't look like ad names
    if (!adName.includes('_')) return false;
    if (adName.toLowerCase().includes('facebook post')) return false;
    if (adName.toLowerCase().includes('post id')) return false;
    
    // Extract actual ad name (remove "Test - BF text ||" prefix)
    let actualAdName = adName;
    if (actualAdName.includes('||')) {
      const parts = actualAdName.split('||');
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].includes('_')) {
          actualAdName = parts[i].trim();
          break;
        }
      }
    }
    
    // ✅ ADD THESE 4 LINES HERE - Remove "Test - " prefix if present
    if (actualAdName.startsWith('Test - ')) {
      actualAdName = actualAdName.substring(7); // Remove "Test - " (7 characters)
    }
    
    // Check for platform identifier (FB, YT, GDN, TT)
    const hasPlatform = /_(FB|YT|GDN|TT)_/.test(actualAdName);
    
    // ✅ FALLBACK: Check if ad starts with a known product name
    const knownProducts = Object.keys(PRODUCT_LOCALE_FOLDERS);
    let startsWithKnownProduct = false;
    for (let i = 0; i < knownProducts.length; i++) {
      const productName = knownProducts[i].toLowerCase();
      const adNameLower = actualAdName.toLowerCase();
      if (adNameLower.startsWith(productName + '_') || adNameLower.startsWith(productName + '-')) {
        startsWithKnownProduct = true;
        break;
      }
    }
    
    // Must have either: platform identifier OR start with known product name
    if (!hasPlatform && !startsWithKnownProduct) return false;
    
    // For platform-based ads: Must have at least 4 underscores
    if (hasPlatform) {
      const underscoreCount = (actualAdName.match(/_/g) || []).length;
      if (underscoreCount < 4) return false;
    }
    
    return true;
  });

  Logger.log('Rows after filtering: ' + rows.length);
  if (rows.length > 0) {
    Logger.log('Sample row:');
    Logger.log('  Ad name: ' + rows[0][adNameCol]);
    Logger.log('  Spend: ' + rows[0][spendCol]);
    Logger.log('  ROAS: ' + rows[0][roasCol]);
    Logger.log('  CPA: ' + rows[0][cpaCol]);
    Logger.log('  Purchases: ' + rows[0][purchasesCol]);
    Logger.log('  AOV: ' + rows[0][aovCol]);
  }
  
  const productLabelCol = 8; // Column I
const csvSourceCol = 9;    // Column J

const processedData = rows.map(row => {
  const rawAdName = row[adNameCol] || '';
  const cleanedAdName = extractActualAdName(rawAdName);
  
  return {
    adName: rawAdName,
    cleanAdName: cleanedAdName,
    spend: parseCurrency(row[spendCol] || '0'),
    spendFormatted: row[spendCol] || '0',
    aov: row[aovCol] || '-',
    roas: row[roasCol] || '-',
    cpm: '-',
    ctr: '-',
    cpcLink: '-',
    holdRate: '-',
    purchases: row[purchasesCol] || '-',
    cpa: row[cpaCol] || '-',
    checkouts: '-',
    hookRate: '-',
    productLabel: row[productLabelCol] || '',
    csvSource: row[csvSourceCol] || ''
  };
});
  
  // Sort by spend and get top ads
  processedData.sort((a, b) => b.spend - a.spend);
  const topAds = processedData; // Don't limit - we'll limit per product later
  
// GROUP ADS BY PRODUCT (including labels)
const unknownProducts = {};  // ✅ Track unknown products for summary
const adsByProduct = {};

topAds.forEach(ad => {
  const parts = ad.cleanAdName.split('_');
  
  // Find platform index
  let platformIndex = -1;
  for (let i = 0; i < parts.length; i++) {
    if (parts[i] === 'FB' || parts[i] === 'YT' || parts[i] === 'GDN' || parts[i] === 'TT') {
      platformIndex = i;
      break;
    }
  }
  
let product = platformIndex > 0 ? parts.slice(0, platformIndex).join('_') : parts[0];
  
// ✅ FALLBACK: If no platform found, check if ad starts with a known product name
if (platformIndex <= 0) {
  const knownProducts = Object.keys(PRODUCT_LOCALE_FOLDERS);
  for (let i = 0; i < knownProducts.length; i++) {
    const productName = knownProducts[i];
    const adNameLower = ad.cleanAdName.toLowerCase();
    if (adNameLower.startsWith(productName.toLowerCase() + '_') || 
        adNameLower.startsWith(productName.toLowerCase() + '-')) {
      product = productName; // Use exact casing from PRODUCT_LOCALE_FOLDERS
      break;
    }
  }
}
  
// ✅ NORMALIZE PRODUCT NAME FIRST (handle rebrands and variants like Clairu → Clairu_M2)
product = normalizeProductName(product);

// ✅ CHECK FOR UNKNOWN PRODUCT REMAPPING (from CSV Configurator decisions)
try {
  const scriptProps = PropertiesService.getScriptProperties();
  const mappingsJson = scriptProps.getProperty('UNKNOWN_PRODUCT_MAPPINGS');
  if (mappingsJson) {
    const unknownMappings = JSON.parse(mappingsJson);
    if (unknownMappings[product] && unknownMappings[product] !== 'ignore') {
      Logger.log(`  🔄 Remapping unknown "${product}" → "${unknownMappings[product]}"`);
      product = unknownMappings[product];
    }
  }
} catch (e) {
  // Ignore mapping errors
}

// ✅ VALIDATE: Only accept products that exist in PRODUCT_LOCALE_FOLDERS (have Drive folders)
const validProducts = Object.keys(PRODUCT_LOCALE_FOLDERS);
let isValidProduct = false;
for (let i = 0; i < validProducts.length; i++) {
  if (validProducts[i].toLowerCase() === product.toLowerCase()) {
    product = validProducts[i]; // Use exact casing
    isValidProduct = true;
    break;
  }
}

// Skip ads with unknown/invalid product names
if (!isValidProduct) {
  // Track unknown products for summary
  if (!unknownProducts[product]) {
    unknownProducts[product] = { count: 0, examples: [] };
  }
  unknownProducts[product].count++;
  if (unknownProducts[product].examples.length < 3) {
    unknownProducts[product].examples.push(ad.cleanAdName.substring(0, 60));
  }
  return; // Skip this ad - forEach continues to next
}
  
// Then normalize against TARGET_CPA keys for exact casing (for CPA lookup)
const productKeys = Object.keys(TARGET_CPA);
for (let i = 0; i < productKeys.length; i++) {
  if (productKeys[i].toLowerCase() === product.toLowerCase()) {
    product = productKeys[i];
    break;
  }
}
  
  // ✅ CRITICAL: Use the productLabel that's ALREADY in the ad object
  // This was set when we created processedData from the CSV rows
  const productLabel = ad.productLabel || '';
  
  // Create final product name WITH label
  let finalProduct = product;
  if (productLabel && productLabel.trim() !== '') {
    finalProduct = `${product} (${productLabel})`;
  }
  
  // Store for later use (and for debugging)
  ad.productWithLabel = finalProduct;
  ad.product = product; // Also store base product name
  
  // Debug log every 20th ad to see what's happening
  const totalAdsProcessed = Object.values(adsByProduct).reduce((sum, arr) => sum + arr.length, 0);
  if (totalAdsProcessed % 20 === 0) {
    Logger.log(`Ad #${totalAdsProcessed}: ${ad.cleanAdName.substring(0, 50)} → Product: "${finalProduct}" (base: "${product}", label: "${productLabel}")`);
  }
  
  if (!adsByProduct[finalProduct]) {
    adsByProduct[finalProduct] = [];
  }
  adsByProduct[finalProduct].push(ad);
});
  
  // GET TOP N ADS PER PRODUCT
  let finalTopAds = [];
  Object.keys(adsByProduct).forEach(product => {
    const productAds = adsByProduct[product].sort((a, b) => b.spend - a.spend);
    const topProductAds = productAds.slice(0, settings.topN);
    finalTopAds = finalTopAds.concat(topProductAds);
  });
  
  finalTopAds.sort((a, b) => b.spend - a.spend);
  
  // Log unknown products summary
  const unknownProductNames = Object.keys(unknownProducts);
  if (unknownProductNames.length > 0) {
    Logger.log('\n⚠️ === UNKNOWN PRODUCTS FOUND (excluded from report) ===');
    unknownProductNames.forEach(p => {
      const info = unknownProducts[p];
      Logger.log(`  ❌ "${p}": ${info.count} ads`);
      info.examples.forEach(ex => Logger.log(`      Example: ${ex}`));
    });
    Logger.log('=== To include these, add them to Product Folders in Config sheet ===\n');
  }
  
  return {
    totalAds: processedData.length,
    topAds: finalTopAds,
    totalSpend: processedData.reduce((sum, ad) => sum + ad.spend, 0),
    linksAdded: 0,
    unknownProducts: unknownProducts
  };
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = i < line.length - 1 ? line[i + 1] : '';
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // Escaped quote inside quotes - add one quote
        current += '"';
        i++; // Skip next quote
      } else {
        // Toggle quote mode
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // Field separator - only when NOT inside quotes
      result.push(current.trim());
      current = '';
    } else {
      // Regular character - add to current field
      current += char;
    }
  }
  
  // Add last field
  result.push(current.trim());
  
  // Remove surrounding quotes from each field
  return result.map(field => {
    // Remove leading/trailing quotes
    return field.replace(/^"(.*)"$/, '$1');
  });
}

/**
 * Generate weekly report with progress tracking
 */
function generateReport() {
  initializeConfigConstants_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet || dataSheet.getLastRow() === 0) {
    SpreadsheetApp.getUi().alert('❌ Error', 'No data found. Please upload CSV files first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return { success: false, message: 'No data found' };  // ← Make sure this returns properly
  }
  
  try {
    const settings = getSettings();
    const data = dataSheet.getDataRange().getValues();
    const report = processAtriaData(data, settings);
    
    // ✅ Make sure report has topAds
    if (!report || !report.topAds || report.topAds.length === 0) {
      return { 
        success: false, 
        message: 'No valid ads found. Make sure ad names follow the format: Product_Platform_Locale_Batch_Variation' 
      };
    }
    
    writeReportToSheet(report, settings);
    
    const message = `✅ Report generated successfully!\n\n` +
                   `📊 Total ads analyzed: ${report.totalAds}\n` +
                   `💰 Total spend: €${report.totalSpend.toFixed(2)}\n` +
                   `🔗 Google Drive links added: ${report.linksAdded || 0}`;
    
    return { success: true, message: message, report: report };
    
  } catch (error) {
    Logger.log('Generate report error: ' + error.toString());
    return { success: false, message: 'Failed to generate report: ' + error.toString() };
  }
}

/**
 * Show progress dialog
 */
function showProgressDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 30px;
            text-align: center;
          }
          h2 {
            color: #4285F4;
            margin-bottom: 30px;
          }
          .progress-container {
            width: 100%;
            background: #f0f0f0;
            border-radius: 10px;
            overflow: hidden;
            margin: 20px 0;
          }
          .progress-bar {
            width: 0%;
            height: 30px;
            background: linear-gradient(90deg, #4285F4, #34A853);
            transition: width 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
          }
          #status {
            font-size: 14px;
            color: #666;
            margin-top: 15px;
            min-height: 20px;
          }
          #eta {
            font-size: 12px;
            color: #999;
            margin-top: 10px;
          }
        </style>
      </head>
      <body>
        <h2>📊 Generating Report</h2>
        <div class="progress-container">
          <div class="progress-bar" id="progressBar">0%</div>
        </div>
        <div id="status">Initializing...</div>
        <div id="eta"></div>
        
        <script>
          let startTime = Date.now();
          
          function updateProgress(percent, message) {
            const bar = document.getElementById('progressBar');
            const status = document.getElementById('status');
            const eta = document.getElementById('eta');
            
            bar.style.width = percent + '%';
            bar.textContent = percent + '%';
            status.textContent = message;
            
            // Calculate ETA
            if (percent > 0 && percent < 100) {
              const elapsed = (Date.now() - startTime) / 1000;
              const estimatedTotal = (elapsed / percent) * 100;
              const remaining = Math.ceil(estimatedTotal - elapsed);
              eta.textContent = 'Estimated time remaining: ' + remaining + 's';
            } else if (percent === 100) {
              eta.textContent = '';
              setTimeout(() => {
                google.script.host.close();
              }, 1000);
            }
          }
          
          // Poll for progress updates
          function checkProgress() {
            google.script.run
              .withSuccessHandler(function(progress) {
                if (progress) {
                  updateProgress(progress.percent, progress.message);
                  if (progress.percent < 100) {
                    setTimeout(checkProgress, 300);
                  }
                }
              })
              .getProgress();
          }
          
          checkProgress();
        </script>
      </body>
    </html>
  `).setWidth(500).setHeight(250);
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'Generating Report');
}

/**
 * Update progress (stores in script properties)
 */
function updateProgress(percent, message) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('PROGRESS', JSON.stringify({ percent: percent, message: message }));
}

/**
 * Get current progress
 */
function getProgress() {
  const props = PropertiesService.getScriptProperties();
  const progress = props.getProperty('PROGRESS');
  return progress ? JSON.parse(progress) : null;
}

function writeReportToSheet(report, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reportSheet = ss.getSheetByName('Weekly Report');
  
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Weekly Report');
  } else {
    reportSheet.clear();
  }
  
  const weekLabel = getWeekLabel(settings.weekLabel);
  let row = 1;
  
  // Title
  reportSheet.getRange(row, 1).setValue('📊 WEEKLY AD PERFORMANCE REPORT')
    .setFontWeight('bold').setFontSize(14);
  row++;
  
  // Week label
  reportSheet.getRange(row, 1).setValue('📅 Week: ' + weekLabel);
  row += 2;
  
  // Product performance header
  reportSheet.getRange(row, 1).setValue('🎯 Product Performance')
    .setFontWeight('bold').setFontSize(12);
  row++;
  
  // Summary stats
  reportSheet.getRange(row, 1).setValue('Total ads analyzed: ' + report.totalAds);
  row++;
  reportSheet.getRange(row, 1).setValue('Total spend: €' + report.totalSpend.toFixed(2));
  row += 2;
  
// GROUP ADS BY PRODUCT (preserve labels from processAtriaData)
const adsByProduct = {};
report.topAds.forEach(ad => {
  // CRITICAL: Always use productWithLabel first (it includes the label)
  let product = ad.productWithLabel || ad.product || 'Unknown';
  
  // Double-check: if productWithLabel is missing but we have label info, reconstruct it
  if (!ad.productWithLabel && ad.productLabel) {
    const baseProduct = ad.product || 'Unknown';
    product = `${baseProduct} (${ad.productLabel})`;
  }
  
  // ✅ NORMALIZE: Handle product name variants (Clairu → Clairu_M2, etc.)
  // Extract base product name (remove label if present)
  let baseProduct = product;
  let label = '';
  const labelMatch = product.match(/^(.+?)\s*\((.+?)\)$/);
  if (labelMatch) {
    baseProduct = labelMatch[1].trim();
    label = labelMatch[2].trim();
  }
  
  // Normalize the base product
  baseProduct = normalizeProductName(baseProduct);
  
  // Reconstruct with label if it existed
  product = label ? `${baseProduct} (${label})` : baseProduct;
  
  if (!adsByProduct[product]) {
    adsByProduct[product] = [];
  }
  adsByProduct[product].push(ad);
});
  
  // Sort products alphabetically
  const productNames = Object.keys(adsByProduct).sort();
  
  let linksAdded = 0;
  let globalIndex = 1;
  
  // CREATE A SEPARATE TABLE FOR EACH PRODUCT
  productNames.forEach(product => {
    const productAds = adsByProduct[product];
    const targetCPA = getTargetCpaLazy_()[product] || 60;
    
    // Product header
    reportSheet.getRange(row, 1).setValue(`📦 ${product.toUpperCase()}`)
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#E8F0FE');
    row++;
    
    reportSheet.getRange(row, 1).setValue(`Target CPA: €${targetCPA} | Total Ads: ${productAds.length} | Total Spend: €${productAds.reduce((sum, ad) => sum + ad.spend, 0).toFixed(2)}`)
      .setFontSize(10)
      .setFontColor('#666666');
    row++;
    
    // Table headers for this product
    const headers = ['Rank', 'Ad Name', 'Spend', 'ROAS', 'CPA', 'Purchases', 'AOV'];
    reportSheet.getRange(row, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF');
    row++;
    
    // Add ads for this product
    const totalAds = productAds.length;
    for (let index = 0; index < productAds.length; index++) {
      const ad = productAds[index];
      
      // Update progress
      const progressPercent = 60 + Math.round((globalIndex / report.topAds.length) * 30);
      updateProgress(progressPercent, `Searching Drive link for ${product} ad ${index + 1}/${totalAds}`);
      globalIndex++;
      
      // Add emoji to clean ad name
      const displayName = addEmojiToAdName(ad.adName, ad.cleanAdName);
      
      // Extract product and locale for Drive search
      const parts = ad.cleanAdName.split('_');
      let platformIndex = -1;
      for (let i = 0; i < parts.length; i++) {
        if (parts[i] === 'FB' || parts[i] === 'YT' || parts[i] === 'GDN' || parts[i] === 'TT') {
          platformIndex = i;
          break;
        }
      }
      
      const productName = platformIndex > 0 ? parts.slice(0, platformIndex).join('_') : parts[0];
      const locale = platformIndex >= 0 && platformIndex + 1 < parts.length ? parts[platformIndex + 1] : '';
      
      // Use fast cached lookup for top 10 per product
      let driveFile = null;
      if (index < 10) {
        driveFile = findAdCreativeFileFast(ad.cleanAdName, productName, locale);
      }
      
      // Store drive URL for Slack
      if (driveFile) {
        ad.driveUrl = driveFile.url;
      }
      
      const rowData = [
        index + 1,
        displayName,
        ad.spendFormatted,
        ad.roas,
        ad.cpa,
        ad.purchases,
        ad.aov
      ];
      
      reportSheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      
      // Add hyperlink if file found
      if (driveFile) {
        const nameCell = reportSheet.getRange(row, 2);
        nameCell.setFormula('=HYPERLINK("' + driveFile.url + '", "' + displayName.replace(/"/g, '""') + '")');
        nameCell.setFontColor('#1a73e8');
        linksAdded++;
      }
      
      // Color code CPA with ±$15 tolerance
const cpaValue = parseFloat(String(ad.cpa).replace(/[^0-9.]/g, ''));
if (!isNaN(cpaValue)) {
  const cpaCell = reportSheet.getRange(row, 5);
  if (cpaValue < targetCPA - 5) {
    cpaCell.setBackground('#D4EDDA').setFontColor('#155724');  // Green
  } else if (cpaValue > targetCPA + 5) {
    cpaCell.setBackground('#F8D7DA').setFontColor('#721C24');  // Red
  } else {
    cpaCell.setBackground('#FFF3CD').setFontColor('#856404');  // Yellow
  }
}
      
      row++;
    }
    
    // Add spacing between products
    row += 2;
  });
  
  // Update report with links added count
  report.linksAdded = linksAdded;
  
  // Auto-resize columns
  for (let i = 1; i <= 7; i++) {
    reportSheet.autoResizeColumn(i);
  }
  
  // Add "Next Steps" section
  reportSheet.getRange(row, 1).setValue('💡 Next Steps:').setFontWeight('bold');
  row++;
  reportSheet.getRange(row, 1).setValue('1. Review the data above');
  row++;
  reportSheet.getRange(row, 1).setValue('2. Add your insights below');
  row++;
  reportSheet.getRange(row, 1).setValue('3. Use "Send to Slack" to post');
  row += 2;
  
  // Add insights section
  reportSheet.getRange(row, 1).setValue('✍️ YOUR INSIGHTS:')
    .setFontWeight('bold').setBackground('#FFF9C4');
  row++;
  reportSheet.getRange(row, 1, 10, 5).setBackground('#FFFEF7');
}

// ===== AD NAME CLEANING (FIXED) =====
function cleanAdName(name) {
  // Find position of _VID, _IMG, or _DYN (case insensitive)
  const vidPos = name.search(/_VID/i);
  const imgPos = name.search(/_IMG/i);
  const dynPos = name.search(/_DYN/i);
  
  // Collect all valid positions and find the earliest
  const positions = [vidPos, imgPos, dynPos].filter(pos => pos !== -1);
  const cutPos = positions.length > 0 ? Math.min(...positions) : -1;
  
  // Trim to that position if found
  const cleanName = cutPos !== -1 ? name.substring(0, cutPos) : name;
  
  // Add emoji based on type (check in priority order)
  if (/DYN/i.test(name)) {
    return '🔄 ' + cleanName; // Dynamic ad
  } else if (/VID/i.test(name)) {
    return '🎥 ' + cleanName; // Video ad
  } else if (/IMG/i.test(name)) {
    return '📷 ' + cleanName; // Image ad
  }
  
  return cleanName;
}

/**
 * Add emoji to a clean ad name based on the file type in the original name
 */
function addEmojiToAdName(originalAdName, cleanAdName) {
  // Determine emoji based on original ad name content
  if (/DYN/i.test(originalAdName)) {
    return '🔄 ' + cleanAdName;
  } else if (/VID/i.test(originalAdName)) {
    return '🎥 ' + cleanAdName;
  } else if (/IMG/i.test(originalAdName)) {
    return '📷 ' + cleanAdName;
  }
  
  // ✅ FALLBACK: Non-standard Heatoor ads (no _FB_) are always videos
  const isNonStandardHeatoor = /^Heatoor[_-]/i.test(cleanAdName) && !/_(FB|YT|GDN|TT)_/.test(cleanAdName);
  if (isNonStandardHeatoor) {
    return '🎥 ' + cleanAdName;
  }
  
  return cleanAdName;
}

// ===== GOOGLE DRIVE INTEGRATION =====

/**
 * Extract folder ID from various Google Drive URL formats
 */
function extractDriveFolderId(url) {
  if (!url) return null;
  
  // Try different URL patterns
  let match = url.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  
  match = url.match(/id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  
  return null;
}

function extractActualAdName(adName) {
  // Remove emoji prefixes
  let cleaned = adName.replace(/^[🎥📷🔄]\s*/, '');
  
  // Remove everything after || and extract the part with underscores
  if (cleaned.includes('||')) {
    const parts = cleaned.split('||');
    // Find the part that looks like a proper ad name (has underscores)
    let foundPart = false;
    for (let i = 0; i < parts.length; i++) {
      if (parts[i].includes('_') && /_(FB|YT|GDN|TT)_/.test(parts[i])) {
        cleaned = parts[i].trim();
        foundPart = true;
        break;
      }
    }
    // ✅ FALLBACK: If no platform-based part found, just take the first part (before ||)
    if (!foundPart) {
      cleaned = parts[0].trim();
    }
  }
  
  // Remove [Copy:...] or [copy:...] or (Test X) prefixes
  let prevCleaned = '';
  while (prevCleaned !== cleaned) {
    prevCleaned = cleaned;
    cleaned = cleaned.replace(/^\[[Cc]opy:[^\]]+\]\s*/, '');
    cleaned = cleaned.replace(/^\([Tt]est\s+[A-Za-z]\)\s*/, '');
  }
  
  // Find the product name
  const products = Object.keys(TARGET_CPA);
  let productFound = null;
  let productIndex = -1;
  
  for (let i = 0; i < products.length; i++) {
    const product = products[i];
    const idx = cleaned.indexOf(product + '_');
    if (idx !== -1) {
      productFound = product;
      productIndex = idx;
      break;
    }
  }
  
  if (productIndex === -1) {
    return cleaned;
  }
  
  // Extract from product onwards
  cleaned = cleaned.substring(productIndex);
  
  // Now extract just Product_Platform_Locale_Batch_Variation
  const parts = cleaned.split('_');
  
  // Find platform
  let platformIdx = -1;
  for (let i = 0; i < parts.length; i++) {
    if (parts[i] === 'FB' || parts[i] === 'YT' || parts[i] === 'GDN' || parts[i] === 'TT') {
      platformIdx = i;
      break;
    }
  }
  
  if (platformIdx === -1) {
    return cleaned;
  }
  
  // Core is: Platform + 3 more parts (Locale, Batch, Variation)
  const stopIdx = platformIdx + 4;
  const coreParts = parts.slice(0, stopIdx);
  
  return coreParts.join('_');
}

// === CACHED FOLDER LOOKUP (handles year subfolders) ===
const _folderCache = {};

function findAdCreativeFileFast(adName, product, locale) {
  // Build cache key
  const cacheKey = `${product}_${locale}`;
  
  // Check if product/locale folder exists
  const folders = PRODUCT_LOCALE_FOLDERS;
  if (!folders[product] || !folders[product][locale]) {
    return null;
  }
  
  const baseFolderUrl = folders[product][locale];
  if (!baseFolderUrl) return null;
  
  const folderId = baseFolderUrl.split('/folders/')[1];
  if (!folderId) return null;
  
  try {
    // Get or build cache for this product/locale
    if (!_folderCache[cacheKey]) {
      _folderCache[cacheKey] = {};
      
      const baseFolder = DriveApp.getFolderById(folderId);
      const topFolders = baseFolder.getFolders();
      
      while (topFolders.hasNext()) {
        const topFolder = topFolders.next();
        const topFolderName = topFolder.getName();
        
        // Check if this is a year folder (2023, 2024, 2025, 2026, etc.)
        if (/^20\d{2}$/.test(topFolderName)) {
          // It's a year folder - look inside for batch folders
          const batchFolders = topFolder.getFolders();
          
          while (batchFolders.hasNext()) {
            const batchFolder = batchFolders.next();
            const batchFolderName = batchFolder.getName();
            
            // Extract batch number from folder name
            const batchMatch = batchFolderName.match(/_(\d+)_?\(/);
            if (batchMatch) {
              const batchNum = batchMatch[1];
              
              // Cache all files in this batch folder
              const files = batchFolder.getFiles();
              const fileList = [];
              while (files.hasNext()) {
                const file = files.next();
                fileList.push({
                  name: file.getName(),
                  id: file.getId(),
                  url: file.getUrl()
                });
              }
              
              _folderCache[cacheKey][batchNum] = fileList;
            }
          }
        } else {
          // Not a year folder - check if it's a batch folder directly
          const batchMatch = topFolderName.match(/_(\d+)_?\(/);
          if (batchMatch) {
            const batchNum = batchMatch[1];
            
            const files = topFolder.getFiles();
            const fileList = [];
            while (files.hasNext()) {
              const file = files.next();
              fileList.push({
                name: file.getName(),
                id: file.getId(),
                url: file.getUrl()
              });
            }
            
            _folderCache[cacheKey][batchNum] = fileList;
          }
        }
      }
      
      Logger.log(`📁 Cached ${Object.keys(_folderCache[cacheKey]).length} batch folders for ${cacheKey}`);
    }
    
    // Parse ad name to get batch number
    const parts = adName.split('_');
    let platformIndex = -1;
    for (let i = 0; i < parts.length; i++) {
      if (['FB', 'YT', 'GDN', 'TT'].includes(parts[i])) {
        platformIndex = i;
        break;
      }
    }
    
    if (platformIndex < 0) return null;
    
    const batch = parts[platformIndex + 2]; // Batch is 2 positions after platform
    if (!batch || !_folderCache[cacheKey][batch]) return null;
    
    // Search for matching file in cached file list
    const searchPattern = adName.toLowerCase();
    const files = _folderCache[cacheKey][batch];
    
    for (const file of files) {
      if (file.name.toLowerCase().startsWith(searchPattern)) {
        return { name: file.name, url: file.url };
      }
    }
    
    return null;
  } catch (e) {
    Logger.log(`⚠️ Fast lookup error for ${adName}: ${e.message}`);
    return null;
  }
}

function clearFolderCache() {
  for (const key in _folderCache) {
    delete _folderCache[key];
  }
  Logger.log('🧹 Folder cache cleared');
}

function findAdCreativeFile(adName, product, locale) {
  Logger.log('=== findAdCreativeFile START ===');
  Logger.log('Input adName: ' + adName);
  Logger.log('Input product: ' + product);
  Logger.log('Input locale: ' + locale);
  
  try {
    // Case-insensitive product lookup
    let actualProduct = null;
    const productLocaleFolders = getProductLocaleFoldersLazy_();
    const productKeys = Object.keys(productLocaleFolders);
    
    for (let i = 0; i < productKeys.length; i++) {
      if (productKeys[i].toLowerCase() === product.toLowerCase()) {
        actualProduct = productKeys[i];
        break;
      }
    }
    
    if (!actualProduct) {
      Logger.log('❌ Product not found in PRODUCT_LOCALE_FOLDERS: ' + product);
      return null;
    }
    
    Logger.log('✓ Product matched: ' + actualProduct);
    
    // Get the locale folder URL for this product
    const productLocales = getProductLocaleFoldersLazy_()[actualProduct];
    
    const folderUrl = productLocales[locale];
    if (!folderUrl) {
      Logger.log('❌ Locale not found for product ' + actualProduct + ': ' + locale);
      return null;
    }
    Logger.log('✓ Folder URL: ' + folderUrl);
    
    // ... rest of the function stays the same
    
    // Extract folder ID and get folder
    const folderId = extractDriveFolderId(folderUrl);
    if (!folderId) {
      Logger.log('❌ Could not extract folder ID from URL');
      return null;
    }
    Logger.log('✓ Folder ID: ' + folderId);
    
    const baseFolder = DriveApp.getFolderById(folderId);
    Logger.log('✓ Base folder accessed: ' + baseFolder.getName());
    
    // Parse ad name: Product_Platform_Locale_Batch_Variation
    const parts = adName.split('_');
    Logger.log('Ad name parts: ' + JSON.stringify(parts));
    
    // For product names with underscores (like Clairu_M2), need to adjust indices
    let platformIndex = -1;
    for (let i = 0; i < parts.length; i++) {
      if (parts[i] === 'FB' || parts[i] === 'YT' || parts[i] === 'GDN' || parts[i] === 'TT') {
        platformIndex = i;
        break;
      }
    }
    
    if (platformIndex === -1) {
      Logger.log('❌ Platform not found in ad name');
      return null;
    }
    Logger.log('✓ Platform index: ' + platformIndex);
    
    const platform = parts[platformIndex]; // "FB"
    const locale2 = parts[platformIndex + 1]; // "US"
    const batch = parts[platformIndex + 2]; // "163"
    const variation = parts[platformIndex + 3]; // "2"
    
    Logger.log('Parsed values:');
    Logger.log('  Platform: ' + platform);
    Logger.log('  Locale: ' + locale2);
    Logger.log('  Batch: ' + batch);
    Logger.log('  Variation: ' + variation);
    
    // Go directly to the batch folder
    const batchFolderName = locale + '_' + batch;
    Logger.log('Looking for batch folder: ' + batchFolderName);
    
    const batchFolder = findBatchFolderDirect(baseFolder, batchFolderName, product);
    
    if (!batchFolder) {
      Logger.log('❌ Batch folder not found: ' + batchFolderName);
      return null;
    }
    Logger.log('✓ Batch folder found: ' + batchFolder.getName());
    
    // Build search pattern
    const searchPattern = product + '_' + platform + '_' + locale + '_' + batch + '_' + variation;
    Logger.log('Search pattern: ' + searchPattern);
    
    // Search only in this batch folder
    const files = batchFolder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      fileCount++;
      
      if (fileName.includes(searchPattern)) {
        Logger.log('✅ MATCH FOUND: ' + fileName);
        return {
          name: fileName,
          url: file.getUrl()
        };
      }
    }
    
    Logger.log('❌ No match found. Searched ' + fileCount + ' files in batch folder');
    return null;
    
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.toString());
    Logger.log('Stack trace: ' + e.stack);
    return null;
  }
}

function findBatchFolderDirect(baseFolder, batchFolderName, product) {
  const parts = batchFolderName.split('_');
  const locale = parts[0];
  const batch = parts[1];
  
  // Check root level first
  let folders = baseFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName();
    
    if (folderName.includes(product)) {
      const pattern = product + '_FB_' + locale + '_' + batch;
      const afterPattern = folderName.indexOf(pattern);
      
      if (afterPattern !== -1) {
        const charAfterPattern = folderName.charAt(afterPattern + pattern.length);
        
        if (charAfterPattern === '_' || charAfterPattern === '(' || charAfterPattern === '') {
          return folder;
        }
      }
    }
  }
  
  // Check inside year folders (2024, 2025, etc.)
  const yearPattern = /^20\d{2}$/;
  folders = baseFolder.getFolders();
  
  while (folders.hasNext()) {
    const yearFolder = folders.next();
    const yearFolderName = yearFolder.getName();
    
    if (yearPattern.test(yearFolderName)) {
      const subfolders = yearFolder.getFolders();
      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        const folderName = folder.getName();
        
        if (folderName.includes(product)) {
          const pattern = product + '_FB_' + locale + '_' + batch;
          const afterPattern = folderName.indexOf(pattern);
          
          if (afterPattern !== -1) {
            const charAfterPattern = folderName.charAt(afterPattern + pattern.length);
            
            if (charAfterPattern === '_' || charAfterPattern === '(' || charAfterPattern === '') {
              return folder;
            }
          }
        }
      }
    }
  }
  
  return null;
}

/**
 * Send report to Slack
 */
function sendWeeklyReportToSlack() {
  initializeConfigConstants_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = ss.getSheetByName('Weekly Report');
  
  if (!reportSheet) {
    SpreadsheetApp.getUi().alert('❌ Error', 'No report found. Please generate a report first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const settings = getSettings();
  
  // ✅ CHANGE THIS LINE
  if (!settings.slackBotToken || !settings.slackChannelIdWeekly) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Slack Bot Token or Weekly Report Channel ID not configured. Please go to Settings first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = reportSheet.getDataRange().getValues();
  
  // Find ALL product tables
  const productTables = [];
  
  for (let i = 0; i < data.length; i++) {
    const cellValue = String(data[i][0]);
    
    if (cellValue.includes('📦')) {
      const productName = cellValue.replace('📦', '').trim();
      
      let headerRow = -1;
      for (let j = i + 1; j < Math.min(i + 5, data.length); j++) {
        if (data[j][0] === 'Rank') {
          headerRow = j;
          break;
        }
      }
      
      if (headerRow === -1) continue;
      
      const productAds = [];
      for (let k = headerRow + 1; k < data.length; k++) {
        const row = data[k];
        
        if (!row[0] || row[0] === '' || String(row[0]).includes('📦')) break;
        if (String(row[1]).includes('Next Steps') || String(row[1]).includes('YOUR INSIGHTS')) break;
        
        // Skip ads with no purchases
        const purchases = row[5];
        const purchaseCount = parseInt(purchases) || 0;
        if (purchaseCount === 0 || purchases === '-' || purchases === '') {
          continue;
        }
        
        const nameCell = reportSheet.getRange(k + 1, 2);
        const formula = nameCell.getFormula();
        let driveUrl = null;
        
        if (formula && formula.includes('HYPERLINK')) {
          const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
          if (urlMatch) {
            driveUrl = urlMatch[1];
          }
        }
        
        let originalAdName = row[1];
originalAdName = originalAdName.replace(/^[\u{1F300}-\u{1F9FF}\u{2600}-\u{26FF}]\s*/u, '');
if (originalAdName.includes('||')) {
  originalAdName = originalAdName.split('||')[0].trim();
}
// Clean the ad name (shorten it and remove extra suffixes)
originalAdName = cleanAdName(originalAdName);
        
        productAds.push({
          rank: row[0],
          adName: row[1],
          originalAdName: originalAdName,
          spend: row[2],
          roas: row[3],
          cpa: row[4],
          purchases: row[5],
          aov: row[6],
          driveUrl: driveUrl
        });
      }
      
      if (productAds.length > 0) {
        productTables.push({
          product: productName,
          ads: productAds
        });
      }
    }
  }
  
  // Get week label
  let weekLabel = 'Unknown';
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).includes('Week:')) {
      weekLabel = String(data[i][0]).replace('📅 Week: ', '');
      break;
    }
  }
  
  const channel = settings.slackChannelIdWeekly;  // ✅ CHANGE THIS LINE
  const sentMessages = [];
  
  // Send header
  let headerText = `*📊 Weekly Ad Performance Report*\n`;
  headerText += `Week of ${weekLabel}\n`;
  headerText += `Total Products: ${productTables.length}\n\n`;
  
  const headerMsg = sendSlackMessageWithBot(headerText, channel);
  if (headerMsg) sentMessages.push(headerMsg);
  Utilities.sleep(1000);
  
// Send each product
let successCount = 0;
productTables.forEach((table, index) => {
  // ✅ FIX: Extract base product name (remove label in parentheses)
  let baseProduct = table.product;
  
  Logger.log('=== PRODUCT: ' + table.product + ' ===');
  Logger.log('Before extraction: "' + baseProduct + '"');
  
  // Remove label like "(Main)" or "(TOP5)" to get base product name
  const labelMatch = baseProduct.match(/^(.+?)\s*\(/);
  if (labelMatch) {
    baseProduct = labelMatch[1].trim();
    Logger.log('After extraction: "' + baseProduct + '"');
  } else {
    Logger.log('No label found - keeping original');
  }
  
  // Get target CPA for this base product (case-insensitive lookup)
  let targetCPA = 60; // default
  for (const key in TARGET_CPA) {
    if (key.toLowerCase() === baseProduct.toLowerCase()) {
      targetCPA = getTargetCpaLazy_()[key];
      Logger.log('✅ MATCHED! key="' + key + '" → targetCPA=' + targetCPA);
      break;
    }
  }
  
  if (targetCPA === 60) {
    Logger.log('❌ NO MATCH - using default 60');
  }
  
  Logger.log('Final targetCPA: ' + targetCPA);
  Logger.log('');
  
  let text = `\`${table.product.toUpperCase()} (TARGET CPA: €${targetCPA})\`\n`;
  text += `Top ${table.ads.length} Highest Spending Ads:\n`;
    
    table.ads.forEach((ad, adIndex) => {
      const cpaValue = parseFloat(String(ad.cpa).replace(/[^0-9.]/g, ''));

let cpaEmoji = '🟢';  // Default to green
if (cpaValue > targetCPA + 5) {
  cpaEmoji = '🔴';  // Red if >€5 over target
} else if (cpaValue >= targetCPA - 5 && cpaValue <= targetCPA + 5) {
  cpaEmoji = '🟡';  // Yellow if within ±€5 of target
}
// Anything below target - €5 stays green (default)
      
      // Use the already-cleaned originalAdName
let adNameDisplay = ad.originalAdName;

if (ad.driveUrl) {
  adNameDisplay = `<${ad.driveUrl}|${adNameDisplay}>`;
}
      
      text += `${adIndex + 1}. ${adNameDisplay}\n`;
      text += `   Spend: $${ad.spend} | ROAS: ${ad.roas} | CPA: ${cpaEmoji} $${ad.cpa}\n`;
    });
    
    text += `\n`;
    
    const productMsg = sendSlackMessageWithBot(text, channel);
    if (productMsg) {
      sentMessages.push(productMsg);
      successCount++;
    }
    
    if (index < productTables.length - 1) {
      Utilities.sleep(1000);
    }
  });
  
  
  // Store message IDs for deletion
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_WEEKLY_REPORT_MESSAGES', JSON.stringify(sentMessages));
  
  SpreadsheetApp.getUi().alert(
    '✅ Success', 
    `Report sent to Slack successfully!\n\nSent ${successCount}/${productTables.length} product reports.`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Save to history
  if (productTables.length > 0 && productTables[0].ads.length > 0) {
    const totalSpend = productTables.reduce((sum, table) => 
      sum + table.ads.reduce((adSum, ad) => 
        adSum + parseFloat(String(ad.spend).replace(/[^0-9.]/g, '')), 0), 0);
    
    const report = {
      topAds: productTables[0].ads,
      weekLabel: weekLabel,
      totalSpend: totalSpend,
      totalAds: productTables.reduce((sum, table) => sum + table.ads.length, 0)
    };
    saveToHistory(report, settings);
  }
}

/**
 * Send message using Slack Bot Token (allows deletion later)
 */
function sendSlackMessageWithBot(text, channel) {
  const settings = getSettings();
  const botToken = settings.slackBotToken;
  
  if (!botToken) {
    Logger.log('No bot token configured');
    return null;
  }
  
  const url = 'https://slack.com/api/chat.postMessage';
  
  const payload = {
    channel: channel,
    text: text,
    mrkdwn: true
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + botToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return {
        success: true,
        ts: data.ts,
        channel: data.channel
      };
    } else {
      Logger.log('Slack API error: ' + data.error);
      return null;
    }
  } catch (e) {
    Logger.log('Error sending message: ' + e.toString());
    return null;
  }
}

/**
 * Delete a Slack message
 */
function deleteSlackMessage(channel, messageTs) {
  const settings = getSettings();
  const botToken = settings.slackBotToken;
  
  if (!botToken) {
    Logger.log('No bot token configured');
    return false;
  }
  
  const url = 'https://slack.com/api/chat.delete';
  
  const payload = {
    channel: channel,
    ts: messageTs
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + botToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      Logger.log('Message deleted successfully');
      return true;
    } else {
      Logger.log('Failed to delete: ' + data.error);
      return false;
    }
  } catch (e) {
    Logger.log('Error deleting message: ' + e.toString());
    return false;
  }
}

/**
 * Delete last report from Slack
 */
function deleteLastWeeklyReport() {
  const props = PropertiesService.getScriptProperties();
  const lastReportData = props.getProperty('LAST_WEEKLY_REPORT_MESSAGES');
  
  if (!lastReportData) {
    SpreadsheetApp.getUi().alert('❌ No Report Found', 'No recent report found to delete.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const messages = JSON.parse(lastReportData);
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Delete Report?', 
    `Delete ${messages.length} messages from Slack?\n\nThis cannot be undone.`, 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  let deletedCount = 0;
  messages.forEach(msg => {
    if (deleteSlackMessage(msg.channel, msg.ts)) {
      deletedCount++;
    }
    Utilities.sleep(500);
  });
  
  props.deleteProperty('LAST_WEEKLY_REPORT_MESSAGES');
  
  ui.alert('✅ Deleted', `Successfully deleted ${deletedCount}/${messages.length} messages.`, ui.ButtonSet.OK);
}

/**
 * Send a message to Slack webhook
 */
function sendSlackMessage(webhookUrl, text) {
  const payload = {
    text: text,
    mrkdwn: true
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    logErrorToHub_('Slack Message', e.message || String(e));
    Logger.log('Slack send error: ' + e.message);
    return false;
  }
}

/**
 * Generate compact report text for Slack with Drive links and CPA performance
 */
function getReportText(report) {
  const settings = getSettings();
  
  let text = `*📊 Weekly Ad Performance Report*\n`;
  text += `Week of ${report.weekLabel}\n\n`;
  
  // Track unknown/invalid products for reporting
  const unknownProducts = {};
  
  // Group by product
  const adsByProduct = {};
  report.topAds.forEach(ad => {
    // Use originalAdName if available, otherwise extract from adName
let nameForGrouping = ad.originalAdName || ad.adName;

// Remove any remaining emojis and extract product
nameForGrouping = nameForGrouping.replace(/^[🎥📷]\s*/, '');
if (nameForGrouping.includes('||')) {
  nameForGrouping = nameForGrouping.split('||')[0].trim();
}

let product = nameForGrouping.split('_')[0];
// ✅ NORMALIZE PRODUCT NAME
product = normalizeProductName(product);
    
    if (!adsByProduct[product]) {
      adsByProduct[product] = [];
    }
    adsByProduct[product].push({ ...ad });
  });
  
  // Display each product's ads
  Object.keys(adsByProduct).forEach(product => {
    const productAds = adsByProduct[product];
    const targetCPA = getTargetCpaLazy_()[product] || 60;
    
    // Product header with code block formatting (entire line in orange/bold)
    text += `\`${product.toUpperCase()} (TARGET CPA: €${targetCPA})\`\n`;
    text += `Top ${productAds.length} Highest Spending Ads:\n`;
    
    productAds.forEach((ad, index) => {
      // Parse CPA value
      const cpaValue = parseFloat(String(ad.cpa).replace(/[^0-9.]/g, ''));
      
      // Determine CPA performance emoji
let cpaEmoji = '🟢';
if (cpaValue > targetCPA + 5) {
  cpaEmoji = '🔴';
} else if (cpaValue >= targetCPA - 5 && cpaValue <= targetCPA + 5) {
  cpaEmoji = '🟡';
}
// If cpaValue < targetCPA - 5, it stays green
      
      // Use the already-cleaned originalAdName
let adNameDisplay = ad.originalAdName;
      
      // Add Drive link if available
      if (ad.driveUrl) {
        adNameDisplay = `<${ad.driveUrl}|${adNameDisplay}>`;
      }
      
      // Compact single-line format
      text += `${index + 1}. ${adNameDisplay}\n`;
      text += `   Spend: $${ad.spend} | ROAS: ${ad.roas} | CPA: ${cpaEmoji} $${ad.cpa}\n`;
    });
    
    text += `\n`;
  });
  
  if (settings.spreadsheetUrl) {
    text += `<${settings.spreadsheetUrl}|📊 View Full Report>\n`;
  }
  
  return text;
}

// ===== UTILITIES =====
function parseCurrency(value) {
  if (!value || value === '-') return 0;
  let str = value.toString().trim();
  
  // Detect European format: has comma but no period, OR ends with € after comma-number
  // Examples: "252,00 €", "1.234,56 €", "252,00"
  const isEuropeanFormat = /\d,\d{2}(\s*€)?$/.test(str) || (/,/.test(str) && !/\./.test(str));
  
  if (isEuropeanFormat) {
    // European: periods are thousands separators, comma is decimal
    str = str.replace(/\./g, '').replace(',', '.');
  }
  
  // Remove currency symbols and any remaining commas (US thousands separator)
  str = str.replace(/[€$,]/g, '').trim();
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function getWeekLabel(customLabel) {
  if (customLabel) return customLabel;
  
  const today = new Date();
  const lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - today.getDay() - 6);
  
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);
  
  const formatDate = (date) => {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getDate();
  };
  
  return formatDate(lastMonday) + ' - ' + formatDate(lastSunday) + ', ' + today.getFullYear();
}

// ===== HISTORY =====
function saveToHistory(report, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName('Historical Data');
  
  if (!historySheet) {
    historySheet = ss.insertSheet('Historical Data');
    const headers = ['Week', 'Date', 'Total Ads', 'Total Spend', 'Top Ad', 'Top Ad Spend', 'Top Ad ROAS'];
    historySheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF');
  }
  
  const topAd = report.topAds[0];
  historySheet.appendRow([
    getWeekLabel(settings.weekLabel),
    new Date(),
    report.totalAds,
    report.totalSpend,
    topAd ? topAd.adName.substring(0, 50) : '',
    topAd ? topAd.spend : 0,
    topAd ? topAd.roas : '-'
  ]);
}

// ===== CLEAR DATA =====
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('⚠️ Clear All Data?', 'This will clear:\n• CSV Data\n• Weekly Report\n• Creative Tests\n• Data Info\n\nAre you sure?', ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const dataSheet = ss.getSheetByName('CSV Data');
    if (dataSheet) dataSheet.clear();
    
    const reportSheet = ss.getSheetByName('Weekly Report');
    if (reportSheet) reportSheet.clear();
    
    const testSheet = ss.getSheetByName('Creative Tests');
    if (testSheet) testSheet.clear();
    
    const infoSheet = ss.getSheetByName('Data Info');
    if (infoSheet) infoSheet.clear();
    
    ui.alert('✅ Cleared!', 'All data has been cleared.', ui.ButtonSet.OK);
  }
}

// ===== SETUP =====
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.getRange('A1:B1').setValues([['Setting', 'Value']])
      .setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    
    settingsSheet.getRange('A2:B4').setValues([
      ['Slack Webhook URL', ''],
      ['Top N Ads', 10],
      ['Week Label (optional)', '']
    ]);
    
    settingsSheet.setColumnWidth(1, 200);
    settingsSheet.setColumnWidth(2, 400);
  }
  
  SpreadsheetApp.getUi().alert('✅ Setup Complete!', 'Use "Upload CSV Files" from the Atria Reports menu.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===== CLICKUP API INTEGRATION (LAZY LOADED to prevent menu breaking) =====
let _clickupConfigData = null;

function getClickUpConfigDataLazy_() {
  if (!_clickupConfigData || !_clickupConfigData.API_KEY) {
    _clickupConfigData = getClickUpConfigFromConfig_();
  }
  return _clickupConfigData;
}

function getFbAdsFolderId_() {
  const config = getClickUpConfigDataLazy_();
  return config.FB_ADS_LISTS_FOLDER_ID || '90151192866';
}

function getClickUpApiKey_() {
  const config = getClickUpConfigDataLazy_();
  return (config.API_KEY || '').trim();
}

function getClickUpTeamId_() {
  const config = getClickUpConfigDataLazy_();
  return config.TEAM_ID || '';
}

function getClickUpSpaceId_() {
  const config = getClickUpConfigDataLazy_();
  // The actual ClickUp Space ID for API calls is FB_ADS_FOLDER_ID (90150821271)
  // SPACE_ID (90151192866) is actually a folder within the space, not the space itself
  return config.FB_ADS_FOLDER_ID || '';
}

function getClickUpFolderId_() {
  const config = getClickUpConfigDataLazy_();
  return config.SPRINT_FOLDER_ID || '';
}

/**
 * Global log buffer for real-time logging
 */
var LOG_BUFFER = [];

function logProgress(message, type) {
  type = type || 'info';
  Logger.log(message);
  
  var log = {
    message: message,
    type: type,
    timestamp: new Date().getTime()
  };
  
  LOG_BUFFER.push(log);
  
  // Also save to cache so other executions can see it
  try {
    var cache = CacheService.getScriptCache();
    var existing = cache.get('LOG_BUFFER');
    var allLogs = existing ? JSON.parse(existing) : [];
    allLogs.push(log);
    
    // Keep only last 100
    if (allLogs.length > 100) {
      allLogs.shift();
    }
    
    cache.put('LOG_BUFFER', JSON.stringify(allLogs), 600); // 10 min expiry
  } catch (e) {
    // Cache errors are non-fatal
    Logger.log('Cache error: ' + e);
  }
}

function getRecentLogs() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('LOG_BUFFER');
    var logs = cached ? JSON.parse(cached) : [];
    Logger.log('getRecentLogs returning ' + logs.length + ' logs');
    return logs;
  } catch (e) {
    Logger.log('Cache read error: ' + e);
    return [];
  }
}

function getAnalysisStatus() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('LOG_BUFFER');
    var logsArray = cached ? JSON.parse(cached) : [];
    
    // Only mark complete if we see the completion message in logs
    var isComplete = logsArray.some(function(log) {
      return log.message && log.message.indexOf('ANALYSIS COMPLETE') > -1;
    });
    
    return {
      logs: logsArray,
      isComplete: isComplete,
      logCount: logsArray.length
    };
  } catch (e) {
    return { logs: [], isComplete: false, logCount: 0, error: e.toString() };
  }
}

function clearLogs() {
  LOG_BUFFER.length = 0;
  try {
    CacheService.getScriptCache().remove('LOG_BUFFER');
  } catch (e) {
    Logger.log('Cache clear error: ' + e);
  }
}

/**
 * Get sprint lists from PC SPRINT folders - but rename them to "Week X"
 * Now includes 2-day weekend buffer (Sat-Sun) before each week
 * Supports multiple sprint folders (current + legacy) for year transitions
 */
function getClickUpSprints() {
  const spaceId = getClickUpSpaceId_();
  const url = `https://api.clickup.com/api/v2/space/${spaceId}/folder?archived=false`;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': getClickUpApiKey_(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      Logger.log('Error: ' + response.getContentText());
      return [];
    }
    
    const data = JSON.parse(response.getContentText());
    
    // Get folder IDs to check (current + legacy if exists)
    const folderIds = [String(getClickUpFolderId_())];
    
    // Check for legacy folder ID in config
    const config = getClickUpConfigDataLazy_();
    const legacyFolderId = config.LEGACY_SPRINT_FOLDER_ID;
const LEGACY_MIN_WEEK = 40; // Only show weeks 40+ from legacy folder
    if (legacyFolderId) {
      folderIds.push(String(legacyFolderId));
      Logger.log('Including legacy sprint folder: ' + legacyFolderId);
    }
    
    // Find all matching sprint folders
    const sprintFolders = data.folders.filter(function(f) { 
      return folderIds.includes(f.id); 
    });
    
    if (sprintFolders.length === 0) {
      Logger.log('No sprint folders found');
      return [];
    }
    
    Logger.log('Found ' + sprintFolders.length + ' sprint folder(s)');
    
    // Combine lists from all sprint folders
    let allLists = [];
    sprintFolders.forEach(function(folder) {
      if (folder.lists) {
        // Add folder name prefix to distinguish years
        folder.lists.forEach(function(list) {
  list._folderName = folder.name; // Track which folder it came from
});

// If legacy folder, only include recent weeks
if (folder.name.includes('2025')) {
  const recentLists = folder.lists.filter(function(list) {
    const weekMatch = list.name.match(/Sprint\s+(\d+)/);
    if (weekMatch) {
      const weekNum = parseInt(weekMatch[1]);
      return weekNum >= LEGACY_MIN_WEEK;
    }
    return false;
  });
  allLists = allLists.concat(recentLists);
} else {
  allLists = allLists.concat(folder.lists);
}
      }
    });
    
    // Sort sprints by number (newest first) and rename to "Week"
    const sprints = allLists.map(function(list) {
      // Replace "Sprint" with "Week" in the display name
      let displayName = list.name.replace(/^Sprint\s+/, 'Week ');
      
      // If from 2025 folder, add year indicator
      if (list._folderName && list._folderName.includes('2025')) {
        displayName = displayName + ' (2025)';
      }
      
      // Extract date range from list name to add weekend buffer
      const dateMatch = list.name.match(/\((\d{1,2}\/\d{1,2})\s*-\s*(\d{1,2}\/\d{1,2})\)/);
      
      let startDate = null;
      let endDate = null;
      
      if (dateMatch) {
        const startStr = dateMatch[1];
        const endStr = dateMatch[2];
        
        // Parse dates - determine year based on folder
        const currentYear = new Date().getFullYear();
        const year = (list._folderName && list._folderName.includes('2025')) ? 2025 : currentYear;
        const startParts = startStr.split('/');
        const endParts = endStr.split('/');
        
        startDate = new Date(year, parseInt(startParts[0]) - 1, parseInt(startParts[1]));
        endDate = new Date(year, parseInt(endParts[0]) - 1, parseInt(endParts[1]));
        
        // Add 2-day weekend buffer to start (go back to Saturday)
        startDate.setDate(startDate.getDate() - 2);
      }
      
      return {
        id: list.id,
        name: displayName,
        originalName: list.name,
        startDate: startDate,
        endDate: endDate,
        folderName: list._folderName
      };
    });
    
    // Sort by start date (newest first), with current year sprints first
    sprints.sort(function(a, b) {
      // Current year sprints come first
      const aIs2025 = a.folderName && a.folderName.includes('2025');
      const bIs2025 = b.folderName && b.folderName.includes('2025');
      if (aIs2025 !== bIs2025) return aIs2025 ? 1 : -1;
      
      // Then sort by date
      if (a.startDate && b.startDate) {
        return b.startDate - a.startDate;
      }
      return 0;
    });
    
    return sprints;
    
  } catch (e) {
    Logger.log('Error fetching sprints: ' + e.message);
    return [];
  }
}

/**
 * Get master task's launch date (simplified - just get the date, no complex logic)
 */
function getMasterLaunchDate(masterTask) {
  // Check master's custom Launch Date field
  if (masterTask.custom_fields) {
    const launchField = masterTask.custom_fields.find(f => 
      f.name && f.name.toLowerCase().includes('launch date')
    );
    if (launchField && launchField.value) {
      const customDate = parseInt(launchField.value);
      if (!isNaN(customDate)) {
        return customDate;
      }
    }
  }
  
  // Fallback to master's date_updated
  if (masterTask.date_updated) {
    return parseInt(masterTask.date_updated);
  }
  
  return null;
}

/**
 * Get launch date from custom field or fallback to date_updated
 * @param {Object} task - ClickUp task object
 * @return {Number|null} - Timestamp in milliseconds
 */
function getTaskLaunchDate(task) {
  // Priority 1: Check for "Launch Date" custom field
  if (task.custom_fields && task.custom_fields.length > 0) {
    const launchField = task.custom_fields.find(field => 
      field.name && field.name.toLowerCase().includes('launch date')
    );
    
    if (launchField && launchField.value) {
      // Custom date fields return timestamp as string, need to parse
      const customDate = parseInt(launchField.value);
      if (!isNaN(customDate)) {
        Logger.log(`  Using custom Launch Date for ${task.name}: ${new Date(customDate).toLocaleDateString()}`);
        return customDate;
      }
    }
  }
  
  // Priority 2: Fallback to date_updated for "launched" status
  const currentStatus = task.status ? task.status.status.toLowerCase() : '';
  
  if (currentStatus === 'launched' && task.date_updated) {
    const dateUpdated = parseInt(task.date_updated);
    Logger.log(`  Fallback to date_updated for ${task.name} (status: launched): ${new Date(dateUpdated).toLocaleDateString()}`);
    return dateUpdated;
  }
  
  // Priority 3: For "closed" status, use date_closed if available
  if (currentStatus === 'closed' && task.date_closed) {
    const dateClosed = parseInt(task.date_closed);
    Logger.log(`  Using date_closed for ${task.name}: ${new Date(dateClosed).toLocaleDateString()}`);
    return dateClosed;
  }
  
  // Last resort: use date_updated anyway
  if (task.date_updated) {
    const dateUpdated = parseInt(task.date_updated);
    Logger.log(`  Last resort - date_updated for ${task.name}: ${new Date(dateUpdated).toLocaleDateString()}`);
    return dateUpdated;
  }
  
  return null;
}

/**
 * Get launch date for a batch with hierarchical fallback
 * Priority: Batch's own Launch Date → Batch's date_updated (if LAUNCHED only) → Parent's Launch Date → Parent's date_updated → Master's Launch Date → Master's date_updated
 */
function getBatchLaunchDate(batch, parentTask, masterTask) {
  // Priority 1: Batch's own Launch Date custom field
  if (batch.custom_fields && batch.custom_fields.length > 0) {
    const launchField = batch.custom_fields.find(f => 
      f.name && f.name.toLowerCase().includes('launch date')
    );
    if (launchField && launchField.value) {
      const customDate = parseInt(launchField.value);
      if (!isNaN(customDate)) {
        Logger.log(`    Using batch's custom Launch Date: ${new Date(customDate).toLocaleDateString()}`);
        return customDate;
      }
    }
  }
  
  // Priority 2: If batch is LAUNCHED (not closed!), use its date_updated
  const batchStatus = batch.status ? batch.status.status.toLowerCase() : '';
  if (batchStatus === 'launched' && batch.date_updated) {
    const dateUpdated = parseInt(batch.date_updated);
    Logger.log(`    Using batch's date_updated (status: ${batchStatus}): ${new Date(dateUpdated).toLocaleDateString()}`);
    return dateUpdated;
  }
  
  // Priority 3: Parent task's Launch Date custom field
  if (parentTask && parentTask.custom_fields) {
    const parentLaunchField = parentTask.custom_fields.find(f => 
      f.name && f.name.toLowerCase().includes('launch date')
    );
    if (parentLaunchField && parentLaunchField.value) {
      const parentDate = parseInt(parentLaunchField.value);
      if (!isNaN(parentDate)) {
        Logger.log(`    Inheriting parent's Launch Date: ${new Date(parentDate).toLocaleDateString()}`);
        return parentDate;
      }
    }
  }
  
  // Priority 4: Parent task's date_updated (if LAUNCHED only)
  if (parentTask && parentTask.status) {
    const parentStatus = parentTask.status.status ? parentTask.status.status.toLowerCase() : '';
    if (parentStatus === 'launched' && parentTask.date_updated) {
      const parentDateUpdated = parseInt(parentTask.date_updated);
      Logger.log(`    Inheriting parent's date_updated (status: ${parentStatus}): ${new Date(parentDateUpdated).toLocaleDateString()}`);
      return parentDateUpdated;
    }
  }
  
  // Priority 5: Master task's Launch Date custom field
  if (masterTask && masterTask.custom_fields) {
    const masterLaunchField = masterTask.custom_fields.find(f => 
      f.name && f.name.toLowerCase().includes('launch date')
    );
    if (masterLaunchField && masterLaunchField.value) {
      const masterDate = parseInt(masterLaunchField.value);
      if (!isNaN(masterDate)) {
        Logger.log(`    Inheriting master's Launch Date: ${new Date(masterDate).toLocaleDateString()}`);
        return masterDate;
      }
    }
  }
  
  // Priority 6: Master task's date_updated (if LAUNCHED only)
  if (masterTask && masterTask.status) {
    const masterStatus = masterTask.status.status ? masterTask.status.status.toLowerCase() : '';
    if (masterStatus === 'launched' && masterTask.date_updated) {
      const masterDateUpdated = parseInt(masterTask.date_updated);
      Logger.log(`    Inheriting master's date_updated (status: ${masterStatus}): ${new Date(masterDateUpdated).toLocaleDateString()}`);
      return masterDateUpdated;
    }
  }
  
  Logger.log(`    ❌ No launch date found at any level`);
  return null;
}

function fetchTasksWithSmartStop(listId, options, sprintStart) {
  const allTasks = [];
  const maxPages = 15;
  let page = 0;
  
  // Calculate date filter: 60 days before sprint start (covers ~8 weeks of batches)
  let dateFilter = '';
  if (sprintStart) {
    const filterDate = sprintStart - (60 * 24 * 60 * 60 * 1000); // 60 days in milliseconds
    dateFilter = '&date_updated_gt=' + filterDate;
    Logger.log('  📅 Filtering tasks updated after: ' + new Date(filterDate).toLocaleDateString());
  }
  
  for (page = 0; page < maxPages; page++) {
    const url = 'https://api.clickup.com/api/v2/list/' + listId + '/task?subtasks=true&include_closed=true&page=' + page + dateFilter;
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());
      
      if (data.tasks && data.tasks.length > 0) {
        allTasks.push.apply(allTasks, data.tasks);
        
        // Stop if we got less than 100 tasks (last page)
        if (data.tasks.length < 100) {
          Logger.log('  ✅ Reached last page at page ' + (page + 1) + ' (' + allTasks.length + ' total tasks)');
          break;
        }
        
        Utilities.sleep(100);
      } else {
        break;
      }
    } catch (e) {
      Logger.log('  ⚠️ Error fetching page ' + page + ': ' + e.toString());
      break;
    }
  }
  
  Logger.log('  📊 Fetched ' + allTasks.length + ' tasks from ' + (page + 1) + ' pages');
  
  return allTasks;
}

function getClickUpSprintTasks(listId, productFilter) {
  initializeConfigConstants_();  // ✅ ADD THIS LINE
  if (productFilter) {
    logProgress('Fetching sprint tasks from ClickUp (filtered to: ' + (Array.isArray(productFilter) ? productFilter.join(', ') : productFilter) + ')...', 'info');
  } else {
    logProgress('Fetching sprint tasks from ClickUp...', 'info');
  }
  
  const sprint = getClickUpSprints().find(function(s) { return s.id === listId; });
  
  if (!sprint) {
    logProgress('Sprint not found for list ID: ' + listId, 'error');
    Logger.log('Sprint not found for list ID: ' + listId);
    return [];
  }
  
  const sprintName = sprint.originalName;
  const displayName = sprint.name;
  
  logProgress('Analyzing: ' + displayName, 'info');
  
  const sprintStart = sprint.startDate;
  const sprintEnd = sprint.endDate;
  
  if (!sprintStart || !sprintEnd) {
    logProgress('Could not get sprint dates from: ' + sprintName, 'error');
    Logger.log('Could not get sprint dates from: ' + sprintName);
    return [];
  }
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': getClickUpApiKey_(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  Logger.log('=== FETCHING BATCHES BY LAUNCH DATE ===');
  Logger.log('Week: ' + displayName);
  Logger.log('Date range: ' + new Date(sprintStart).toLocaleDateString() + ' - ' + new Date(sprintEnd).toLocaleDateString());
  
  const spaceId = getClickUpSpaceId_();
  const folderUrl = 'https://api.clickup.com/api/v2/space/' + spaceId + '/folder?archived=false';
  const folderResponse = UrlFetchApp.fetch(folderUrl, options);
  const folderData = JSON.parse(folderResponse.getContentText());
  const fbAdsFolderId = getFbAdsFolderId_();
  const fbAdsFolder = folderData.folders.find(function(f) { return f.id == fbAdsFolderId || f.id === String(fbAdsFolderId); });
  
  logProgress('Found FB Ads folder with ' + fbAdsFolder.lists.length + ' product lists', 'info');
  
  const batchPattern = /^([A-Za-z_0-9]+)_(FB|YT|GDN|TT)_([A-Z]{2,4})_(\d+)/;
  const seenBatchIds = new Set();
  const allBatchTasks = [];
  const processedTaskIds = new Set();
  
  let totalTasksChecked = 0;
  
  function matchesBatchPattern(taskName) {
    return batchPattern.test(taskName);
  }
  
  function validateTaskStructure(task) {
    const taskNameLower = task.name.toLowerCase();
    if (taskNameLower.startsWith('reviews') || 
        taskNameLower.startsWith('process:') || 
        taskNameLower.startsWith('meets:') || 
        taskNameLower.startsWith('prep:')) {
      return { isValid: true, isAdminTask: true };
    }
    return { isValid: true };
  }

  function fetchTaskWithSubtasks(taskId) {
    const url = 'https://api.clickup.com/api/v2/task/' + taskId + '?include_subtasks=true';
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  }

  const MASTER_TASK_PATTERNS = [
    /Pack\s*\d+/i,
    /Loc\s+(IMG|VID)/i,
    /ITR\s+\d+/i,
    /YT\s+to\s+FB/i,
    /^\d+$/,
    /^[A-Z]{2,}$/
  ];

  function isPotentialMasterTask(taskName) {
    if (matchesBatchPattern(taskName)) return false;
    for (let i = 0; i < MASTER_TASK_PATTERNS.length; i++) {
      if (MASTER_TASK_PATTERNS[i].test(taskName)) return true;
    }
    return true;
  }
  
  fbAdsFolder.lists.forEach(function(productList) {
    try {
      if (productFilter) {
        const normalizedListName = normalizeProductName(productList.name);
        const productsToFetch = Array.isArray(productFilter) ? productFilter : [productFilter];
        const shouldFetch = productsToFetch.some(function(prod) {
          const baseProd = prod.match(/^(.+?)\s*\((.+?)\)$/) ? prod.match(/^(.+?)\s*\((.+?)\)$/)[1] : prod;
          return normalizedListName.toUpperCase() === baseProd.toUpperCase();
        });
        
        if (!shouldFetch) {
          Logger.log('\n⏭️ Skipping ' + productList.name + ' (not in filter)');
          return;
        }
      }
      
      Logger.log('\n=== Checking ' + productList.name + ' ===');
      
      const tasks = fetchTasksWithSmartStop(productList.id, options, sprintStart);
      
      totalTasksChecked += tasks.length;
      
      let foundInList = 0;
      let masterChecked = 0;
      let standaloneChecked = 0;
      
      let apiCallCount = 0;
      
      tasks.forEach(function(task) {
        if (processedTaskIds.has(task.id)) return;
        
        const validation = validateTaskStructure(task);
        if (!validation.isValid || validation.isAdminTask) return;
        
        const matchesBatch = matchesBatchPattern(task.name);

        if (!matchesBatch && isPotentialMasterTask(task.name)) {
          try {
            // ✅ THROTTLE: Small delay every 25 API calls to prevent memory crash
            apiCallCount++;
            if (apiCallCount % 25 === 0) {
              Utilities.sleep(100);
            }
            
            const taskDetails = fetchTaskWithSubtasks(task.id);
            const hasChildren = taskDetails.subtasks && taskDetails.subtasks.length > 0;
            
            if (!hasChildren) return;
            
            masterChecked++;
            
            const child = taskDetails.subtasks[0];
            const childMatchesBatch = matchesBatchPattern(child.name);
            
            if (childMatchesBatch) {
              taskDetails.subtasks.forEach(function(childBatch) {
                if (!matchesBatchPattern(childBatch.name)) return;
                
                const batchLaunchDate = getBatchLaunchDate(childBatch, taskDetails, null);
                if (!batchLaunchDate) return;
                
                const inRange = batchLaunchDate >= sprintStart && batchLaunchDate <= sprintEnd;
                if (!inRange) return;
                
                const match = childBatch.name.match(batchPattern);
                if (match && match[0]) {
                  const batchId = match[0].replace(/[()]/g, '');
                  
                  allBatchTasks.push({
                    name: childBatch.name,
                    batchId: batchId,
                    url: 'https://app.clickup.com/t/' + taskDetails.id,
                    productCode: productList.name,
                    launchDate: batchLaunchDate
                  });
                  
                  processedTaskIds.add(childBatch.id);
                  foundInList++;
                }
              });
            } else {
              try {
                apiCallCount++;
                if (apiCallCount % 25 === 0) {
                  Utilities.sleep(100);
                }
                const childDetails = fetchTaskWithSubtasks(child.id);
                const grandchildren = childDetails.subtasks || [];
                
                grandchildren.forEach(function(grandchild) {
                  if (!matchesBatchPattern(grandchild.name)) return;
                  
                  const batchLaunchDate = getBatchLaunchDate(grandchild, childDetails, taskDetails);
                  if (!batchLaunchDate) return;
                  
                  const inRange = batchLaunchDate >= sprintStart && batchLaunchDate <= sprintEnd;
                  if (!inRange) return;
                  
                  const match = grandchild.name.match(batchPattern);
                  if (match && match[0]) {
                    const batchId = match[0].replace(/[()]/g, '');
                    
                    allBatchTasks.push({
                      name: grandchild.name,
                      batchId: batchId,
                      url: 'https://app.clickup.com/t/' + taskDetails.id,
                      productCode: productList.name,
                      launchDate: batchLaunchDate
                    });
                    
                    processedTaskIds.add(grandchild.id);
                    foundInList++;
                  }
                });
              } catch (e3) {}
            }
          } catch (e2) {}
        }
        else if (matchesBatch) {
          try {
            apiCallCount++;
            if (apiCallCount % 25 === 0) {
              Utilities.sleep(100);
            }
            const taskDetails = fetchTaskWithSubtasks(task.id);
            const hasChildren = taskDetails.subtasks && taskDetails.subtasks.length > 0;
            
            if (hasChildren) {
              masterChecked++;
              
              taskDetails.subtasks.forEach(function(childBatch) {
                if (!matchesBatchPattern(childBatch.name)) return;
                
                const batchLaunchDate = getBatchLaunchDate(childBatch, taskDetails, null);
                if (!batchLaunchDate) return;
                
                const inRange = batchLaunchDate >= sprintStart && batchLaunchDate <= sprintEnd;
                if (!inRange) return;
                
                const match = childBatch.name.match(batchPattern);
                if (match && match[0]) {
                  const batchId = match[0].replace(/[()]/g, '');
                  
                  allBatchTasks.push({
                    name: childBatch.name,
                    batchId: batchId,
                    url: 'https://app.clickup.com/t/' + taskDetails.id,
                    productCode: productList.name,
                    launchDate: batchLaunchDate
                  });
                  
                  processedTaskIds.add(childBatch.id);
                  processedTaskIds.add(task.id);
                  foundInList++;
                }
              });
            } else {
              standaloneChecked++;
              
              const batchLaunchDate = getBatchLaunchDate(taskDetails, null, null);
              const inRange = batchLaunchDate && batchLaunchDate >= sprintStart && batchLaunchDate <= sprintEnd;
              
              if (inRange) {
                const match = task.name.match(batchPattern);
                const batchId = match[1] + '_' + match[2] + '_' + match[3] + '_' + match[4];
                
                if (!seenBatchIds.has(batchId)) {
                  seenBatchIds.add(batchId);
                  
                  allBatchTasks.push({
                    name: task.name,
                    batchId: batchId,
                    url: 'https://app.clickup.com/' + getClickUpTeamId_() + '/v/li/' + task.list.id + '?t=' + task.id,
                    productCode: productList.name
                  });
                  
                  foundInList++;
                }
              }
            }
          } catch (e) {}
        }
      });
      
      if (foundInList > 0) {
        logProgress('  ✓ Found ' + foundInList + ' batches in ' + productList.name, 'success');
      }
      
      Logger.log('  Stats: ' + tasks.length + ' tasks, ' + masterChecked + ' masters, ' + standaloneChecked + ' standalone');
      
      Utilities.sleep(75);
      
    } catch (e) {
      logProgress('  ❌ Error: ' + e.toString(), 'error');
    }
  });
  
  logProgress('Total unique batches found: ' + allBatchTasks.length, 'success');
  
  const uniqueBatchesMap = new Map();
  allBatchTasks.forEach(function(batch) {
    if (!uniqueBatchesMap.has(batch.batchId)) {
      uniqueBatchesMap.set(batch.batchId, batch);
    }
  });
  
  const uniqueBatches = Array.from(uniqueBatchesMap.values());
  
  Logger.log('Batches after deduplication: ' + uniqueBatches.length);
  
  return uniqueBatches;
}

/**
 * Analyze creative test performance
 * NOTE: This fetches ClickUp data each time. For multiple products, use analyzeCreativeTestsWithCache() instead.
 */
function analyzeCreativeTests(sprintListId, selectedProduct) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet || dataSheet.getLastRow() === 0) {
    return { success: false, message: 'No CSV data found. Please upload CSV files first.' };
  }
  
  // Get batch IDs from ClickUp
  const sprintTasks = getClickUpSprintTasks(sprintListId, selectedProduct);
  
  if (sprintTasks.length === 0) {
    return { success: false, message: 'No valid batch tasks found in this sprint.' };
  }
  
  // Filter by product if not "ALL"
  let filteredTasks = sprintTasks;
  let requiredLabel = null;

  if (selectedProduct && selectedProduct !== 'ALL') {
    const labelMatch = selectedProduct.match(/^(.+?)\s*\((.+?)\)$/);
    
    if (labelMatch) {
      const baseProduct = labelMatch[1];
      requiredLabel = labelMatch[2].toUpperCase();
      
      Logger.log('Filtering for base product: ' + baseProduct + ' with label: ' + requiredLabel);
      
      filteredTasks = sprintTasks.filter(function(task) {
        const productMatch = task.batchId.match(/^([A-Za-z_0-9]+)_(?:FB|YT|GDN|TT)/);
        if (!productMatch) return false;
        
        let product = productMatch[1];
        product = normalizeProductName(product);
        
        return product === baseProduct;
      });
    } else {
      filteredTasks = sprintTasks.filter(function(task) {
        const productMatch = task.batchId.match(/^([A-Za-z_0-9]+)_(?:FB|YT|GDN|TT)/);
        if (!productMatch) return false;
        
        let product = productMatch[1];
        product = normalizeProductName(product);
        
        return product === selectedProduct;
      });
    }
  }
  
  Logger.log('=== ANALYZING CREATIVE TESTS ===');
  Logger.log('Selected product: ' + (selectedProduct || 'ALL'));
  Logger.log('Required label: ' + (requiredLabel || 'NONE'));
  Logger.log('Found ' + filteredTasks.length + ' tasks after filtering');
  
  // Get all ad data
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log('CSV Headers: ' + JSON.stringify(headers));
  
  // Column indices
  const adNameCol = 0;
  const spendCol = 1;
  const roasCol = 2;
  const cpaCol = 3;
  const purchasesCol = 4;
  const aovCol = 5;
  const hookRateCol = 6;
  const holdRateCol = 7;
  const productLabelCol = 8;
  const csvSourceCol = 9;
  
  Logger.log('Column mapping:');
  Logger.log('  Ad name: ' + adNameCol);
  Logger.log('  Product Label: ' + productLabelCol);
  Logger.log('  CSV Source: ' + csvSourceCol);
  Logger.log('  Hook rate: ' + hookRateCol);
  Logger.log('  Hold rate: ' + holdRateCol);
  
  // Find ads matching these batch IDs
  const testResults = [];
  
  for (let taskIndex = 0; taskIndex < filteredTasks.length; taskIndex++) {
    const task = filteredTasks[taskIndex];
    const batchId = task.batchId;
    
    Logger.log('\n=== BATCH: ' + batchId + ' ===');
    
    const batchAds = [];
    const seenAds = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let adName = row[adNameCol] || '';
      
      // Extract clean ad name
      let cleanAdName = adName;
      if (cleanAdName.includes('||')) {
        const parts = cleanAdName.split('||');
        for (let j = 0; j < parts.length; j++) {
          if (parts[j].includes('_')) {
            cleanAdName = parts[j].trim();
            break;
          }
        }
      }
      
      if (cleanAdName.startsWith('Test - ')) {
        cleanAdName = cleanAdName.substring(7);
      }

      // ✅ FIX: Remove label prefix like [copy: gift], [Copy:money/MalePOV], etc.
      cleanAdName = cleanAdName.replace(/^\[.*?\]\s*/i, '');
      
      const baseBatchId = batchId.replace(/_\([^)]+\)$/, '');
      const adPattern = new RegExp('^' + baseBatchId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '_\\d+');
        
      if (adPattern.test(cleanAdName)) {
        const spend = parseCurrency(row[spendCol] || '0');
        const roas = parseFloat(row[roasCol]) || 0;
        const cpa = parseCurrency(row[cpaCol] || '0');
        const purchases = parseInt(row[purchasesCol]) || 0;
        
        const productLabel = row[productLabelCol] || '';
        let csvSource = row[csvSourceCol] || '';
        
        if (requiredLabel) {
          const adLabel = productLabel.toUpperCase();
          if (adLabel !== requiredLabel) {
            continue;
          }
        }
        
        csvSource = String(csvSource).replace(/%/g, '').replace(/\.0+$/, '').trim();
        const csvSourceNum = parseFloat(csvSource);
        if (!isNaN(csvSourceNum) && csvSourceNum >= 1) {
          csvSource = '1';
        } else {
          csvSource = '0';
        }
        
        const uniqueKey = cleanAdName + '|' + spend.toFixed(2) + '|' + csvSource;
        
        if (seenAds.has(uniqueKey)) {
          Logger.log('  ⏭️ Skipping duplicate: ' + cleanAdName + ' (spend: €' + spend.toFixed(2) + ', source: ' + csvSource + ')');
          continue;
        }
        
        seenAds.add(uniqueKey);
        
        const isVideo = /VID/i.test(cleanAdName);
        let hookRate = null;
        let holdRate = null;

        if (isVideo) {
          if (row[hookRateCol]) {
            const hookValue = parseFloat(row[hookRateCol]);
            if (!isNaN(hookValue)) {
              hookRate = hookValue;
            }
          }
          
          if (row[holdRateCol]) {
            const holdValue = parseFloat(row[holdRateCol]);
            if (!isNaN(holdValue)) {
              holdRate = holdValue;
            }
          }
        }
        
        Logger.log('  ✓ Added: ' + cleanAdName + ' (spend: €' + spend.toFixed(2) + ', label: ' + productLabel + ', source: ' + csvSource + ')');
        
        // ✅ NO Drive search here - just store the ad name
        batchAds.push({
          adName: cleanAdName,
          spend: spend,
          roas: roas,
          cpa: cpa,
          purchases: purchases,
          aov: parseFloat(row[aovCol]) || 0,
          hookRate: hookRate,
          holdRate: holdRate,
          productLabel: productLabel,
          csvSource: csvSource,
          driveUrl: null  // Will be filled later by writeCreativeTestsToSheet
        });
      }
    }
    
    Logger.log('Total unique ads matched: ' + batchAds.length);
    
    if (batchAds.length > 0) {
      const adsByLabel = {};
      batchAds.forEach(function(ad) {
        const label = ad.productLabel || 'NoLabel';
        if (!adsByLabel[label]) {
          adsByLabel[label] = [];
        }
        adsByLabel[label].push(ad);
      });
      
      Logger.log('Ads grouped by label: ' + JSON.stringify(Object.keys(adsByLabel).map(function(lbl) { return lbl + ':' + adsByLabel[lbl].length; })));
      
      Object.keys(adsByLabel).forEach(function(label) {
        const labelAds = adsByLabel[label];
        
        if (requiredLabel && label.toUpperCase() !== requiredLabel) {
          Logger.log('  ⏭️ Skipping label group: ' + label + ' (not matching ' + requiredLabel + ')');
          return;
        }
        
        Logger.log('\n  Processing label group: ' + label + ' (' + labelAds.length + ' ads)');
        
        const productMatch = batchId.match(/^([A-Za-z_0-9]+)_(FB|YT|GDN|TT)/);
        let product = productMatch ? productMatch[1] : 'Unknown';
        product = normalizeProductName(product);
        
        if (label && label !== 'NoLabel') {
          product = product + ' (' + label + ')';
        }
        
        const targetCPA = getTargetCpaLazy_()[product] || getTargetCpaLazy_()[product.split(' (')[0]] || 60;
        
        const categorizedAds = labelAds.map(function(ad) {
          let adCategory = '🔴 Underperformer';
          let adCategoryColor = '#F8D7DA';
          
          if (ad.spend >= 200 && ad.roas >= 1.1 && ad.cpa <= targetCPA + 15) {
            adCategory = '🟢 Top Performer';
            adCategoryColor = '#D4EDDA';
          }
          else if (ad.roas >= 0.85 && ad.cpa <= targetCPA * 1.5) {
            adCategory = '🟡 Mid Performer';
            adCategoryColor = '#FFF3CD';
          }
          else if (ad.roas >= 1.0 && ad.cpa <= targetCPA * 2.0) {
            adCategory = '🟡 Mid Performer';
            adCategoryColor = '#FFF3CD';
          }
          
          return {
            adName: ad.adName,
            spend: ad.spend,
            roas: ad.roas,
            cpa: ad.cpa,
            purchases: ad.purchases,
            aov: ad.aov,
            hookRate: ad.hookRate,
            holdRate: ad.holdRate,
            productLabel: ad.productLabel,
            csvSource: ad.csvSource,
            driveUrl: ad.driveUrl,
            adCategory: adCategory,
            adCategoryColor: adCategoryColor
          };
        });
        
        const topPerformers = categorizedAds.filter(function(ad) { return ad.adCategory === '🟢 Top Performer'; }).sort(function(a, b) { return b.spend - a.spend; });
        const midPerformers = categorizedAds.filter(function(ad) { return ad.adCategory === '🟡 Mid Performer'; }).sort(function(a, b) { return b.spend - a.spend; });
        const underperformers = categorizedAds.filter(function(ad) { return ad.adCategory === '🔴 Underperformer'; }).sort(function(a, b) { return b.spend - a.spend; });
        
        const sortedAds = topPerformers.concat(midPerformers).concat(underperformers);
        
        let batchCategory = '🔴 Underperformer';
        let batchCategoryColor = '#F8D7DA';
        
        if (topPerformers.length > 0) {
          batchCategory = '🟢 Top Performer';
          batchCategoryColor = '#D4EDDA';
        } else if (midPerformers.length > 0) {
          batchCategory = '🟡 Mid Performer';
          batchCategoryColor = '#FFF3CD';
        }
        
        const totalSpend = labelAds.reduce(function(sum, ad) { return sum + ad.spend; }, 0);
        const totalPurchases = labelAds.reduce(function(sum, ad) { return sum + ad.purchases; }, 0);
        
        testResults.push({
          taskName: task.name,
          taskUrl: task.url,
          batchId: batchId,
          product: product,
          targetCPA: targetCPA,
          adCount: labelAds.length,
          totalSpend: totalSpend,
          totalPurchases: totalPurchases,
          topPerformerCount: topPerformers.length,
          midPerformerCount: midPerformers.length,
          underperformerCount: underperformers.length,
          category: batchCategory,
          categoryColor: batchCategoryColor,
          ads: sortedAds
        });
      });
    }
  }
  
  Logger.log('\n=== FINAL RESULTS ===');
  Logger.log('Total batches analyzed: ' + filteredTasks.length);
  Logger.log('Batches with data: ' + testResults.length);
  
  const sprint = getClickUpSprints().find(function(s) { return s.id === sprintListId; });
  const weekName = sprint ? sprint.name : 'Unknown Week';
  
  if (selectedProduct && selectedProduct !== 'ALL' && testResults.length === 0) {
    Logger.log('No results for ' + selectedProduct + ' - creating empty tab');
    
    let baseProduct = selectedProduct;
    const labelMatch = selectedProduct.match(/^(.+?)\s*\((.+?)\)$/);
    if (labelMatch) {
      baseProduct = labelMatch[1];
    }
    
    const abbreviatedWeekName = getAbbreviatedWeekName(weekName);
    createEmptyCreativeTestTab(baseProduct, abbreviatedWeekName);
    
    return {
      success: true,
      results: [],
      totalBatches: 0,
      analyzedBatches: 0,
      weekName: abbreviatedWeekName
    };
  }
  
  return {
    success: true,
    results: testResults,
    totalBatches: filteredTasks.length,
    analyzedBatches: testResults.length,
    weekName: weekName
  };
}

/**
 * Analyze creative test performance (CACHED VERSION)
 * Uses pre-fetched ClickUp data to avoid redundant API calls when analyzing multiple products
 * 
 * @param {string} sprintListId - The ClickUp sprint list ID
 * @param {string} selectedProduct - The product to analyze (or 'ALL')
 * @param {Array} cachedSprintTasks - Pre-fetched ClickUp batch tasks
 * @returns {Object} Analysis results
 */
function analyzeCreativeTestsWithCache(sprintListId, selectedProduct, cachedSprintTasks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet || dataSheet.getLastRow() === 0) {
    return { success: false, message: 'No CSV data found. Please upload CSV files first.' };
  }
  
  // Use cached tasks instead of fetching
  const sprintTasks = cachedSprintTasks;
  
  if (sprintTasks.length === 0) {
    return { success: false, message: 'No valid batch tasks found in this sprint.' };
  }
  
  // 🔥 NEW DEBUG LOGGING: Show what we're filtering before we filter
  Logger.log('\n=== FILTERING FOR PRODUCT: ' + selectedProduct + ' ===');
  Logger.log('Total batches before filtering: ' + sprintTasks.length);
  
  // Filter by product if not "ALL"
  let filteredTasks = sprintTasks;
  let requiredLabel = null; // Track if we need to filter by label

  if (selectedProduct && selectedProduct !== 'ALL') {
    // 🔥 DEBUG: Show all product names we're checking against
    Logger.log('Checking all batch product names:');
    sprintTasks.forEach(function(t) {
      const productMatch = t.batchId.match(/^([A-Za-z_0-9]+)_(?:FB|YT|GDN|TT)/);
      if (productMatch) {
        const rawProduct = productMatch[1];
        const normalized = normalizeProductName(rawProduct);
        Logger.log('  ' + t.batchId + ' → raw: "' + rawProduct + '" → normalized: "' + normalized + '"');
      }
    });
    
    // Check if selectedProduct has a label suffix like "(TOP5)"
    const labelMatch = selectedProduct.match(/^(.+?)\s*\((.+?)\)$/);
    
    if (labelMatch) {
      const baseProduct = labelMatch[1];
      requiredLabel = labelMatch[2].toUpperCase(); // e.g., "TOP5" - normalize to uppercase
      
      Logger.log('Filtering for base product: ' + baseProduct + ' with label: ' + requiredLabel);
      
      // Filter to base product only
      filteredTasks = sprintTasks.filter(function(task) {
        const productMatch = task.batchId.match(/^([A-Za-z_0-9]+)_(?:FB|YT|GDN|TT)/);
        if (!productMatch) return false;
        
        let product = productMatch[1];
        product = normalizeProductName(product);
        
        return product === baseProduct;
      });
    } else {
      // No label - standard product filtering
      filteredTasks = sprintTasks.filter(function(task) {
        const productMatch = task.batchId.match(/^([A-Za-z_0-9]+)_(?:FB|YT|GDN|TT)/);
        if (!productMatch) return false;
        
        let product = productMatch[1];
        product = normalizeProductName(product);
        
        return product.toUpperCase() === selectedProduct.toUpperCase();  // ✅ CASE-INSENSITIVE!
      });
    }
  }
  
  Logger.log('=== ANALYZING CREATIVE TESTS (CACHED) ===');
  Logger.log('Selected product: ' + (selectedProduct || 'ALL'));
  Logger.log('Required label: ' + (requiredLabel || 'NONE'));
  Logger.log('Found ' + filteredTasks.length + ' tasks after filtering');
  
  // 🔥 DEBUG: Show what batches made it through the filter
  if (filteredTasks.length > 0) {
    Logger.log('Batches that passed filter:');
    filteredTasks.forEach(function(t) {
      Logger.log('  ✓ ' + t.batchId);
    });
  } else {
    Logger.log('❌ NO BATCHES PASSED THE FILTER!');
  }
  
  // Get all ad data
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log('CSV Headers: ' + JSON.stringify(headers));
  
  // Column indices (now normalized from processCSVData)
  const adNameCol = 0;
  const spendCol = 1;
  const roasCol = 2;
  const cpaCol = 3;
  const purchasesCol = 4;
  const aovCol = 5;
  const hookRateCol = 6;
  const holdRateCol = 7;
  const productLabelCol = 8; // Column I - Product Label (Main/TOP5)
  const csvSourceCol = 9;    // Column J - CSV Source (0 or 1)
  
  Logger.log('Column mapping:');
  Logger.log('  Ad name: ' + adNameCol);
  Logger.log('  Product Label: ' + productLabelCol);
  Logger.log('  CSV Source: ' + csvSourceCol);
  Logger.log('  Hook rate: ' + hookRateCol);
  Logger.log('  Hold rate: ' + holdRateCol);
  
  // Find ads matching these batch IDs
  const testResults = [];
  
  for (let taskIndex = 0; taskIndex < filteredTasks.length; taskIndex++) {
    const task = filteredTasks[taskIndex];
    const batchId = task.batchId;
    
    Logger.log('\n=== BATCH: ' + batchId + ' ===');
    
    // ✅ EXTRACT DRIVE URL ONCE PER BATCH (BEFORE THE AD LOOP)
    let batchDriveUrl = null;
    const reportSheet = ss.getSheetByName('Weekly Report');
    
    if (reportSheet) {
      const reportRange = reportSheet.getDataRange();
      const reportData = reportRange.getValues();
      const reportFormulas = reportRange.getFormulas();
      
      // Remove _(description) suffix and extract base batch ID
      const baseBatchId = batchId.replace(/_\([^)]+\)$/, '');
      
      Logger.log('     🔍 Searching Weekly Report for batch: ' + baseBatchId);
      
      for (let r = 1; r < reportData.length; r++) {
        const reportAdName = String(reportData[r][1] || '');
        
        // Remove emoji and whitespace
        const cleanReportName = reportAdName.replace(/[📷🎥📦]/g, '').trim();
        
        if (cleanReportName.startsWith(baseBatchId + '_')) {
          const formula = reportFormulas[r][1];
          if (formula && formula.includes('HYPERLINK')) {
            const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
            if (urlMatch) {
              batchDriveUrl = urlMatch[1];
              Logger.log('     ✅ Found Drive URL: ' + batchDriveUrl);
              break;
            }
          }
        }
      }
      
      if (!batchDriveUrl) {
        Logger.log('     ⚠️ No Drive URL found in Weekly Report');
      }
    }
    
    const batchAds = [];
    
    // Process CSV data rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let adName = row[adNameCol] || '';
      
      // ✅ APPLY FULL META STATUS CLEANING FIRST
      let cleanAdName = adName;
      
      // Remove "(Test X)" prefixes
      cleanAdName = cleanAdName.replace(/^\(Test [A-Z]\)\s*/i, '');
      
      // ✅ STEP 1: Remove leading brackets ([unititled], [untitled], etc.)
      cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
      
      // ✅ STEP 2: Remove parentheses prefixes: (blank) rejected:, (anything):
      cleanAdName = cleanAdName.replace(/^\([^)]*\)\s*/g, '');
      
      // ✅ STEP 3: Remove status word prefixes: rejected:, not delivering:, etc.
      cleanAdName = cleanAdName.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
      
      // ✅ STEP 4: Remove leading brackets AGAIN ([copy:gift], [copy:longlasting], etc.)
      cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
      
      // Remove "Test - " prefix if present
      if (cleanAdName.startsWith('Test - ')) {
        cleanAdName = cleanAdName.substring(7);
      }
      
      // Extract from || notation if present
      if (cleanAdName.includes('||')) {
        const parts = cleanAdName.split('||');
        for (let j = 0; j < parts.length; j++) {
          if (parts[j].includes('_')) {
            cleanAdName = parts[j].trim();
            break;
          }
        }
      }
      
      // ✅ APPLY CLEANING AGAIN AFTER || EXTRACTION
      
      // STEP 1: Remove leading brackets
      cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
      
      // STEP 2: Remove parentheses prefixes
      cleanAdName = cleanAdName.replace(/^\([^)]*\)\s*/g, '');
      
      // STEP 3: Remove status word prefixes
      cleanAdName = cleanAdName.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
      
      // STEP 4: Remove leading brackets AGAIN
      cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
      
      // ✅ FIX: Strip _(description) suffix from batch name for matching
      // Example: "WaveMax_FB_US_313_(New IMG_GiftAngle)" → "WaveMax_FB_US_313"
      const baseBatchId = batchId.replace(/_\([^)]+\)$/, '');
      
      // Match pattern: baseBatchId + _ + variation number (CASE-INSENSITIVE!)
      const adPattern = new RegExp('^' + baseBatchId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '_\\d+', 'i');
      //                                                                                                    ↑
        
      if (adPattern.test(cleanAdName)) {
        const spend = parseCurrency(row[spendCol] || '0');
        const roas = parseFloat(row[roasCol]) || 0;
        const cpa = parseCurrency(row[cpaCol] || '0');
        const purchases = parseInt(row[purchasesCol]) || 0;
        
        // Get product label and CSV source
        const productLabel = row[productLabelCol] || '';
        let csvSource = row[csvSourceCol] || '';
        
        // *** KEY FIX: Filter by required label if specified ***
        if (requiredLabel) {
          const adLabel = productLabel.toUpperCase();
          if (adLabel !== requiredLabel) {
            // Skip this ad - doesn't match required label
            continue;
          }
        }
        
        // Convert CSV source to 0/1 format
        if (csvSource === 'Current') {
          csvSource = '0';
        } else if (csvSource === 'Previous') {
          csvSource = '1';
        }
        
        let hookRate = row[hookRateCol];
        let holdRate = row[holdRateCol];
        
        // Parse percentage strings
        if (typeof hookRate === 'string' && hookRate.includes('%')) {
          hookRate = parseFloat(hookRate.replace('%', '')) / 100;
        } else {
          hookRate = parseFloat(hookRate) || 0;
        }
        
        if (typeof holdRate === 'string' && holdRate.includes('%')) {
          holdRate = parseFloat(holdRate.replace('%', '')) / 100;
        } else {
          holdRate = parseFloat(holdRate) || 0;
        }
        
        batchAds.push({
          adName: cleanAdName,
          spend: spend,
          roas: roas,
          cpa: cpa,
          purchases: purchases,
          aov: parseFloat(row[aovCol]) || 0,
          hookRate: hookRate,
          holdRate: holdRate,
          productLabel: productLabel,
          csvSource: csvSource,
          driveUrl: batchDriveUrl  // ← USE BATCH-LEVEL URL (shared by all ads in this batch)
        });
      }
    }
    
    Logger.log('Total unique ads matched: ' + batchAds.length);
    
    if (batchAds.length > 0) {
      // KEY FIX: GROUP ADS BY THEIR PRODUCT LABEL
      const adsByLabel = {};
      batchAds.forEach(function(ad) {
        const label = ad.productLabel || 'NoLabel';
        if (!adsByLabel[label]) {
          adsByLabel[label] = [];
        }
        adsByLabel[label].push(ad);
      });
      
      Logger.log('Ads grouped by label: ' + JSON.stringify(Object.keys(adsByLabel).map(function(lbl) { return lbl + ':' + adsByLabel[lbl].length; })));
      
      // CREATE A SEPARATE RESULT FOR EACH LABEL
      Object.keys(adsByLabel).forEach(function(label) {
        const labelAds = adsByLabel[label];
        
        // Skip if we're filtering by label and this isn't the right one
        if (requiredLabel && label.toUpperCase() !== requiredLabel) {
          Logger.log('  ⏭️ Skipping label group: ' + label + ' (not matching ' + requiredLabel + ')');
          return;
        }
        
        Logger.log('\n  Processing label group: ' + label + ' (' + labelAds.length + ' ads)');
        
        // Extract product from batch ID
        const productMatch = batchId.match(/^([A-Za-z_0-9]+)_(FB|YT|GDN|TT)/);
        let product = productMatch ? productMatch[1] : 'Unknown';
        product = normalizeProductName(product);
        
        // Apply label to product name
        if (label && label !== 'NoLabel') {
          product = product + ' (' + label + ')';
        }
        
        const targetCPA = getTargetCpaLazy_()[product] || getTargetCpaLazy_()[product.split(' (')[0]] || 60;
        
        // Categorize each ad with UPDATED logic
        const categorizedAds = labelAds.map(function(ad) {
          let adCategory = '🔴 Underperformer';
          let adCategoryColor = '#F8D7DA';
          
          // Top Performer: Spend ≥ $200 AND ROAS ≥ 1.1 AND CPA ≤ targetCPA + $15
          if (ad.spend >= 200 && ad.roas >= 1.1 && ad.cpa <= targetCPA + 15) {
            adCategory = '🟢 Top Performer';
            adCategoryColor = '#D4EDDA';
          }
          // Performer: Two conditions
          else if (ad.roas >= 0.85 && ad.cpa <= targetCPA * 1.5) {
            adCategory = '🟡 Mid Performer';
            adCategoryColor = '#FFF3CD';
          }
          else if (ad.roas >= 1.0 && ad.cpa <= targetCPA * 2.0) {
            adCategory = '🟡 Mid Performer';
            adCategoryColor = '#FFF3CD';
          }
          
          return {
            adName: ad.adName,
            spend: ad.spend,
            roas: ad.roas,
            cpa: ad.cpa,
            purchases: ad.purchases,
            aov: ad.aov,
            hookRate: ad.hookRate,
            holdRate: ad.holdRate,
            productLabel: ad.productLabel,
            csvSource: ad.csvSource,
            driveUrl: ad.driveUrl,  // ← PASS THROUGH DRIVE URL
            adCategory: adCategory,
            adCategoryColor: adCategoryColor
          };
        });
        
        // Sort ads: Winners → Performers → Losers (each sorted by ROAS desc)
        const topPerformers = categorizedAds.filter(function(ad) { return ad.adCategory === '🟢 Top Performer'; }).sort(function(a, b) { return b.spend - a.spend; });
        const midPerformers = categorizedAds.filter(function(ad) { return ad.adCategory === '🟡 Mid Performer'; }).sort(function(a, b) { return b.spend - a.spend; });
        const underperformers = categorizedAds.filter(function(ad) { return ad.adCategory === '🔴 Underperformer'; }).sort(function(a, b) { return b.spend - a.spend; });
        
        const sortedAds = topPerformers.concat(midPerformers).concat(underperformers);
        
        // Determine batch category
        let batchCategory = '🔴 Underperformer';
        let batchCategoryColor = '#F8D7DA';
        
        if (topPerformers.length > 0) {
          batchCategory = '🟢 Top Performer';
          batchCategoryColor = '#D4EDDA';
        } else if (midPerformers.length > 0) {
          batchCategory = '🟡 Mid Performer';
          batchCategoryColor = '#FFF3CD';
        }
        
        const totalSpend = labelAds.reduce(function(sum, ad) { return sum + ad.spend; }, 0);
        const totalPurchases = labelAds.reduce(function(sum, ad) { return sum + ad.purchases; }, 0);
        
        testResults.push({
          taskName: task.name,
          taskUrl: task.url,
          batchId: batchId,
          product: product,
          targetCPA: targetCPA,
          adCount: labelAds.length,
          totalSpend: totalSpend,
          totalPurchases: totalPurchases,
          topPerformerCount: topPerformers.length,
          midPerformerCount: midPerformers.length,
          underperformerCount: underperformers.length,
          category: batchCategory,
          categoryColor: batchCategoryColor,
          ads: sortedAds
        });
      });
    }
  }
  
  Logger.log('\n=== FINAL RESULTS ===');
  Logger.log('Total batches analyzed: ' + filteredTasks.length);
  Logger.log('Batches with data: ' + testResults.length);
  
  // Get week name from sprint info
  const sprint = getClickUpSprints().find(function(s) { return s.id === sprintListId; });
  const weekName = sprint ? sprint.name : 'Unknown Week';
  
  // If analyzing a specific product and NO results found, create empty tab
  if (selectedProduct && selectedProduct !== 'ALL' && testResults.length === 0) {
    Logger.log('No results for ' + selectedProduct + ' - creating empty tab');
    
    // Extract base product name (remove label suffix like "(TOP5)")
    let baseProduct = selectedProduct;
    const labelMatch = selectedProduct.match(/^(.+?)\s*\((.+?)\)$/);
    if (labelMatch) {
      baseProduct = labelMatch[1];
    }
    
    // Convert to abbreviated format
    const abbreviatedWeekName = getAbbreviatedWeekName(weekName);
    
    createEmptyCreativeTestTab(baseProduct, abbreviatedWeekName);
    
    return {
      success: true,
      results: [],
      totalBatches: 0,
      analyzedBatches: 0,
      weekName: abbreviatedWeekName
    };
  }
  
  return {
    success: true,
    results: testResults,
    totalBatches: filteredTasks.length,
    analyzedBatches: testResults.length,
    weekName: weekName
  };
}


/**
 * Create an empty creative test tab with "No creative tests" message
 * Creates tab in the PRODUCT spreadsheet, not ADMIN
 */
function createEmptyCreativeTestTab(productName, weekName) {
  // Get all product spreadsheet IDs from Settings tab
  const productSpreadsheets = getProductSpreadsheetIds();
  
  // Get the product-specific spreadsheet ID
  const productSpreadsheetId = productSpreadsheets[productName];
  
  if (!productSpreadsheetId) {
    Logger.log('❌ No spreadsheet ID found for product: ' + productName);
    Logger.log('Available products: ' + Object.keys(productSpreadsheets).join(', '));
    return;
  }
  
  const productSS = SpreadsheetApp.openById(productSpreadsheetId);
  
  // Convert to abbreviated format
  const abbreviatedWeekName = getAbbreviatedWeekName(weekName);
  
  // Sheet name uses abbreviated format
  const sheetName = abbreviatedWeekName;
  let sheet = productSS.getSheetByName(sheetName);
  
  if (sheet) {
    Logger.log('⚠️ Tab "' + sheetName + '" already exists in ' + productName + ' - will overwrite');
    sheet.clear();
  } else {
    sheet = productSS.insertSheet(sheetName);
    Logger.log('✓ Created new sheet: ' + sheetName + ' in ' + productName + ' spreadsheet');
  }
  
  // Set up header
  sheet.getRange('A1').setValue('📊 CREATIVE TEST ANALYSIS');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  
  sheet.getRange('A2').setValue(abbreviatedWeekName);
  sheet.getRange('A2').setFontSize(11).setFontColor('#666666');
  
  // Main message
  sheet.getRange('A4').setValue('🔍 ' + productName.toUpperCase());
  sheet.getRange('A4').setFontWeight('bold').setFontSize(12).setBackground('#f3f3f3');
  
  sheet.getRange('A6').setValue('No creative tests were conducted this week.');
  sheet.getRange('A6').setFontSize(11).setFontStyle('italic').setFontColor('#999999');
  
  sheet.getRange('A8').setValue('This could mean:');
  sheet.getRange('A8').setFontWeight('bold').setFontSize(10);
  
  sheet.getRange('A9').setValue('• No new batches were launched during ' + abbreviatedWeekName);
  sheet.getRange('A10').setValue('• Batches were launched but had zero purchases');
  sheet.getRange('A11').setValue('• Batches exist in ClickUp but no matching ads in CSV data');
  
  sheet.getRange('A9:A11').setFontSize(10).setFontColor('#666666');
  
  // Format columns
  sheet.setColumnWidth(1, 600);
  sheet.getRange('A1:A11').setVerticalAlignment('top');
  
  // ✅ Reorder week tabs after creating empty tab
  reorderWeekTabs(productSS);
  
  Logger.log('✓ Created/updated empty creative test tab in ' + productName + ' spreadsheet: ' + sheetName);
  
  return { product: productName, created: true };
}

function showCreativeTestDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet || dataSheet.getLastRow() === 0) {
    SpreadsheetApp.getUi().alert('❌ Error', 'No CSV data found. Please upload CSV files first using "Upload CSV & Generate Report".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const sprints = getClickUpSprints();
  
  if (sprints.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Could not fetch sprints from ClickUp. Check your API key.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const products = getProductsFromCSV();
  
  let sprintOptions = '';
  sprints.forEach(sprint => {
    sprintOptions += `<option value="${sprint.id}">${sprint.name}</option>`;
  });
  
  let productCheckboxes = '';
products.forEach(product => {
  productCheckboxes += `
    <div style="padding: 5px 0;">
      <label style="cursor: pointer; display: flex; align-items: center; gap: 8px;">
        <input type="checkbox" class="product-checkbox" value="${product}" style="cursor: pointer;">
        <span>${product}</span>
      </label>
    </div>
  `;
});
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
            margin: 0;
          }
          .container {
            max-width: 700px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          }
          h2 {
            color: #333;
            margin-bottom: 20px;
          }
          label {
            display: block;
            margin-top: 15px;
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
          }
          select {
            width: 100%;
            padding: 10px;
            margin-bottom: 10px;
            border: 2px solid #4285f4;
            border-radius: 4px;
            font-size: 14px;
          }
          select[multiple] {
            height: 150px;
          }
          .select-hint {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
          }
          .legend {
            display: flex;
            gap: 20px;
            font-size: 12px;
            color: #666;
            margin-bottom: 8px;
          }
          .quick-select {
            display: flex;
            gap: 10px;
            margin-top: 10px;
          }
          .quick-btn {
            padding: 6px 12px;
            background: #f0f0f0;
            border: 1px solid #ccc;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
          }
          .quick-btn:hover {
            background: #e0e0e0;
          }
          .checkbox-container {
            margin: 20px 0;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
            border: 2px solid #e0e0e0;
          }
          .checkbox-label {
            display: flex;
            align-items: center;
            gap: 10px;
            cursor: pointer;
            font-weight: normal;
            color: #333;
          }
          input[type="checkbox"] {
            width: 20px;
            height: 20px;
            cursor: pointer;
          }
          .checkbox-help {
            font-size: 12px;
            color: #666;
            margin-top: 8px;
            margin-left: 30px;
          }
          button {
            background: #4285f4;
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            width: 100%;
            margin-top: 20px;
          }
          button:disabled {
            background: #ccc;
            cursor: not-allowed;
          }
          #status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 4px;
            display: none;
          }
          #status.show { display: block; }
          #status.success {
            background: #d4edda;
            color: #155724;
          }
          #status.error {
            background: #f8d7da;
            color: #721c24;
          }
          #logContainer {
            margin-top: 20px;
            padding: 15px;
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            font-family: 'Courier New', monospace;
            font-size: 12px;
            max-height: 300px;
            overflow-y: auto;
            display: none;
          }
          #logContainer.show {
            display: block;
          }
          .log-entry {
            padding: 4px 0;
            border-bottom: 1px solid #e9ecef;
          }
          .log-entry:last-child {
            border-bottom: none;
          }
          .log-entry.info {
            color: #0c5460;
          }
          .log-entry.success {
            color: #155724;
            font-weight: bold;
          }
          .log-entry.warning {
            color: #856404;
          }
          .log-entry.error {
            color: #721c24;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>📈 Analyze Creative Tests</h2>
          
          <label>Select week to analyze:</label>
          <select id="sprintSelect">
            ${sprintOptions}
          </select>
          
          <label>Select products:</label>
<div id="productCheckboxes" style="
  max-height: 200px;
  overflow-y: auto;
  border: 2px solid #4285f4;
  border-radius: 4px;
  padding: 10px;
  background: white;
">
  ${productCheckboxes}
</div>
<div class="select-hint">Click to select/deselect products</div>

<div class="quick-select">
  <button type="button" class="quick-btn" id="selectAllBtn">Select All</button>
  <button type="button" class="quick-btn" id="clearAllBtn">Clear All</button>
</div>
          
          <div class="checkbox-container">
            <label class="checkbox-label">
              <input type="checkbox" id="exportCheckbox">
              <span>📊 Also export to product spreadsheets</span>
            </label>
            <div class="checkbox-help">
              Creates/updates weekly tabs in individual product Google Sheets for easy sharing with team members.
            </div>
          </div>
          
          <button type="button" id="analyzeBtn">Analyze Creative Tests</button>
          
          <div id="logContainer"></div>
          <div id="status"></div>
        </div>
        
        <script>
          const selectAllBtn = document.getElementById('selectAllBtn');
          const clearAllBtn = document.getElementById('clearAllBtn');
          const logContainer = document.getElementById('logContainer');
          
          function addLog(message, type = 'info') {
            const entry = document.createElement('div');
            entry.className = 'log-entry ' + type;
            entry.textContent = new Date().toLocaleTimeString() + ' - ' + message;
            logContainer.appendChild(entry);
            logContainer.scrollTop = logContainer.scrollHeight;
          }
          
          selectAllBtn.addEventListener('click', function() {
  document.querySelectorAll('.product-checkbox').forEach(cb => cb.checked = true);
});

clearAllBtn.addEventListener('click', function() {
  document.querySelectorAll('.product-checkbox').forEach(cb => cb.checked = false);
});
          
          document.getElementById('analyzeBtn').addEventListener('click', function() {
            const sprintId = document.getElementById('sprintSelect').value;
const checkboxes = document.querySelectorAll('.product-checkbox:checked');
const products = Array.from(checkboxes).map(cb => cb.value);
            const shouldExport = document.getElementById('exportCheckbox').checked;  // ✅ FIX: Read checkbox
            const btn = document.getElementById('analyzeBtn');
            const status = document.getElementById('status');
            
            console.log('🔍 DEBUG - shouldExport:', shouldExport);  // Debug
            
            if (products.length === 0) {
              alert('Please select at least one product');
              return;
            }
            
            btn.disabled = true;
            btn.textContent = 'Analyzing...';
            status.className = '';
            status.style.display = 'none';
            
            logContainer.className = 'show';
            logContainer.innerHTML = '';
            
            addLog('Starting creative test analysis...', 'info');
            addLog('Products selected: ' + products.join(', '), 'info');
            addLog('Export to sheets: ' + (shouldExport ? 'YES' : 'NO'), 'info');  // Debug
            
            let shownLogCount = 0; // Track how many logs we've shown
let pollFailures = 0;
const MAX_POLL_FAILURES = 120; // Keep trying for 4 minutes (120 * 2 seconds)

let logPolling = setInterval(function() {
  google.script.run
    .withSuccessHandler(function(logs) {
      pollFailures = 0; // Reset on success
      if (logs && logs.length > shownLogCount) {
        // Only show NEW logs
        for (let i = shownLogCount; i < logs.length; i++) {
          addLog(logs[i].message, logs[i].type);
        }
        shownLogCount = logs.length;
      }
    })
    .withFailureHandler(function(error) {
      pollFailures++;
      console.log('Log poll failed (' + pollFailures + '/' + MAX_POLL_FAILURES + '):', error);
      
      if (pollFailures >= MAX_POLL_FAILURES) {
        addLog('⚠️ Lost connection - check Executions log for status', 'warning');
        clearInterval(logPolling);
      }
      // Otherwise keep trying - the next trigger execution might restore connectivity
    })
    .getRecentLogs();
}, 2000);
            
            google.script.run
              .withSuccessHandler(function(result) {
                clearInterval(logPolling);
                
                // ✅ Wait 2 seconds to let final logs load, then close dialog
                if (result.success) {
                  addLog('✅ Analysis complete! Closing in 3 seconds...', 'success');
                  
                  setTimeout(function() {
                    google.script.host.close(); // ✅ AUTO-CLOSE THE DIALOG
                  }, 3000);
                  
                  // Update UI but don't re-enable button since we're closing
                  btn.textContent = 'Complete ✓';
                  
                  status.className = 'success show';
                  
                  let message = '<strong>✅ Success!</strong><br>' +
                    'Analyzed ' + result.analyzedBatches + ' batches from ' + result.totalBatches + ' tasks.<br><br>';
                  
                  if (shouldExport && result.exportedProducts > 0) {
                    message += '📊 <strong>Product Spreadsheets Updated:</strong><br>';
                    if (result.spreadsheetLinks && result.spreadsheetLinks.length > 0) {
                      result.spreadsheetLinks.forEach(function(link) {
                        message += '• <a href="' + link.url + '" target="_blank">' + link.product + '</a><br>';
                      });
                    } else {
                      message += result.exportedProducts + ' spreadsheet(s) updated<br>';
                    }
                    message += '<br>';
                  }
                  
                  if (result.overwriteWarnings && result.overwriteWarnings.length > 0) {
                    message += '♻️ <strong>Overwrote existing tabs:</strong> ' + result.overwriteWarnings.join(', ') + '<br><br>';
                  }
                  
                  if (shouldExport && result.skippedProducts && result.skippedProducts.length > 0) {
                    message += '⚠️ <strong>Skipped ' + result.skippedProducts.length + ' product(s)</strong> (no spreadsheet ID configured)<br><br>';
                  }
                  
                  if (!shouldExport) {
                    message += '📋 Product spreadsheets were NOT updated (checkbox was unchecked).<br><br>';
                  }
                  
                  message += 'Dialog will close automatically...';
                  
                  status.innerHTML = message;
                } else {
                  // On error, keep dialog open
                  btn.disabled = false;
                  btn.textContent = 'Analyze Creative Tests';
                  addLog('Error: ' + result.message, 'error');
                  status.className = 'error show';
                  status.innerHTML = '<strong>❌ Error:</strong><br>' + result.message;
                }
              })
              .withFailureHandler(function(error) {
                clearInterval(logPolling);
                btn.disabled = false;
                btn.textContent = 'Analyze Creative Tests';
                addLog('Error: ' + error.message, 'error');
                status.className = 'error show';
                status.innerHTML = '<strong>❌ Error:</strong><br>' + error.message;
              })
              .processCreativeTestsParallel(sprintId, products, shouldExport);
          });
        </script>
      </body>
    </html>
  `)
  .setWidth(750)
  .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Analyze Creative Tests');
}

function processCreativeTests(sprintListId, selectedProduct, shouldExport) {
  // ✅ Store the week label for later use
  const sprints = getClickUpSprints();
  const selectedSprint = sprints.find(s => s.id === sprintListId);
  if (selectedSprint) {
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('LAST_ANALYZED_WEEK', selectedSprint.name);
    Logger.log('✅ Stored week label: ' + selectedSprint.name);
  }
  
  const result = analyzeCreativeTests(sprintListId, selectedProduct);
  
  if (!result.success) {
    return result;
  }
  
  // 🆕 ADD THIS: Handle empty tabs for products with no tests
  if (selectedProduct && selectedProduct !== 'ALL' && result.results.length === 0) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const weekLabel = selectedSprint ? selectedSprint.name : 'Unknown Week';
    createEmptyCreativeTestTab(ss, selectedProduct, weekLabel);
    Logger.log('✅ Created empty tab for ' + selectedProduct);
  }
  
  // Write results to main Creative Tests sheet (only if we have results)
  if (result.results.length > 0) {
    writeCreativeTestsToSheet(result.results);
  }
  
  // ✅ CONDITIONAL: Only export if user checked the box
  let exportResult = { productsExported: 0, skippedProducts: [], spreadsheetLinks: [] };
  
  if (shouldExport === true) {
    const weekLabel = selectedSprint ? selectedSprint.name : 'Unknown Week';
    exportResult = exportToProductSpreadsheets(result.results, weekLabel, selectedProduct);
  } else {
    Logger.log('⏭️ Skipping product spreadsheet export (user unchecked option)');
  }
  
  return {
    success: true,
    analyzedBatches: result.analyzedBatches,
    totalBatches: result.totalBatches,
    exportedProducts: exportResult.productsExported,
    skippedProducts: exportResult.skippedProducts,
    spreadsheetLinks: exportResult.spreadsheetLinks,
    message: shouldExport 
      ? `✅ Analyzed ${result.analyzedBatches} batches.\n📊 Created/updated tabs in ${exportResult.productsExported} product spreadsheet(s).`
      : `✅ Analyzed ${result.analyzedBatches} batches.`
  };
}

function writeCreativeTestsToSheet(testResults) {
  // ✅ Initialize config for Drive lookups
  initializeConfigConstants_();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let testSheet = ss.getSheetByName('Creative Tests');
  
  if (!testSheet) {
    testSheet = ss.insertSheet('Creative Tests');
  } else {
    testSheet.clear();
  }
  
  let row = 1;
  
  // Title
  testSheet.getRange(row, 1, 1, 2).merge();
  testSheet.getRange(row, 1).setValue('📈 CREATIVE TEST ANALYSIS')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left');
  row++;

  // Add week label
  const scriptProps = PropertiesService.getScriptProperties();
  const lastAnalyzedWeek = scriptProps.getProperty('LAST_ANALYZED_WEEK');

  if (lastAnalyzedWeek) {
    testSheet.getRange(row, 1, 1, 2).merge();
    testSheet.getRange(row, 1).setValue(lastAnalyzedWeek)
      .setFontSize(11)
      .setHorizontalAlignment('left');
    row++;
  }

  row++; // Blank row before headers
  
  // Headers
  const headers = ['', 'Batch / Ad Name', 'Spend', 'ROAS', 'CPA', 'Purchases', 'Hook %', 'Hold %', 'Target CPA', 'Summary'];
  testSheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  // Set column widths
  testSheet.setColumnWidth(1, 30);   // Narrow emoji column
  testSheet.setColumnWidth(2, 350);  // Ad name column
  testSheet.setColumnWidth(3, 80);   // Spend
  testSheet.setColumnWidth(4, 70);   // ROAS
  testSheet.setColumnWidth(5, 70);   // CPA
  testSheet.setColumnWidth(6, 80);   // Purchases
  testSheet.setColumnWidth(7, 70);   // Hook %
  testSheet.setColumnWidth(8, 70);   // Hold %
  testSheet.setColumnWidth(9, 90);   // Target CPA
  testSheet.setColumnWidth(10, 200); // Summary

  // Set Target CPA column (column 9) to text format to preserve € symbol
  testSheet.getRange(1, 9, testSheet.getMaxRows(), 1).setNumberFormat('@STRING@');
  
  row++;
  
  // Sort results by product first, then by category
  const categoryOrder = {'🟢 Top Performer': 1, '🟡 Mid Performer': 2, '🔴 Underperformer': 3};
  testResults.sort((a, b) => {
    // First sort by product
    if (a.product !== b.product) {
      return a.product.localeCompare(b.product);
    }
    // Then by category
    return categoryOrder[a.category] - categoryOrder[b.category];
  });

  // Group by product
  const resultsByProduct = {};
  testResults.forEach(result => {
    if (!resultsByProduct[result.product]) {
      resultsByProduct[result.product] = [];
    }
    resultsByProduct[result.product].push(result);
  });

  // Write results grouped by product
  Object.keys(resultsByProduct).sort().forEach((productName, productIndex) => {
    const productResults = resultsByProduct[productName];
    
    // Add product header (big, bold, colored)
    testSheet.getRange(row, 1, 1, headers.length).merge();
    testSheet.getRange(row, 1)
      .setValue(`📦 ${productName.toUpperCase()}`)
      .setFontWeight('bold')
      .setFontSize(13)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    row++;
    
    // Add spacing row
    row++;
    
    // Write all batches for this product
    productResults.forEach((result, resultIndex) => {
      // BATCH SUMMARY ROW
      const batchSummary = `${result.adCount} ads: ${result.topPerformerCount}T / ${result.midPerformerCount}M / ${result.underperformerCount}U`;
      
      const batchRowData = [
        [
          result.category,
          result.taskName + ' | ' + result.batchId,
          '€' + result.totalSpend.toFixed(2),
          '',
          '',
          result.totalPurchases,
          '',
          '',
          '€' + result.targetCPA.toFixed(2),
          batchSummary
        ]
      ];
      
      testSheet.getRange(row, 1, 1, headers.length).setValues(batchRowData);
      
      let batchColor = result.categoryColor;
      if (result.category === '🟢 Top Performer') batchColor = '#A8D5A8';
      if (result.category === '🟡 Mid Performer') batchColor = '#FFE082';
      if (result.category === '🔴 Underperformer') batchColor = '#EF9A9A';
      
      testSheet.getRange(row, 1, 1, headers.length)
        .setBackground(batchColor)
        .setFontWeight('bold')
        .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      const taskCell = testSheet.getRange(row, 2);
      const taskText = result.taskName + ' | ' + result.batchId;
      taskCell.setFormula('=HYPERLINK("' + result.taskUrl + '", "' + taskText.replace(/"/g, '""') + '")');
      taskCell.setFontColor('#1a73e8');
      
      row++;
      
      // INDIVIDUAL AD ROWS
      result.ads.forEach((ad, adIndex) => {
        const isVideo = /VID/i.test(ad.adName);
        
        let emoji = '🔴';
        if (ad.adCategory === '🟢 Top Performer') emoji = '🟢';
        if (ad.adCategory === '🟡 Mid Performer') emoji = '🟡';
        
        // ✅ SEARCH GOOGLE DRIVE FOR THIS AD (same logic as Weekly Report)
        const parts = ad.adName.split('_');
        let platformIndex = -1;
        for (let j = 0; j < parts.length; j++) {
          if (parts[j] === 'FB' || parts[j] === 'YT' || parts[j] === 'GDN' || parts[j] === 'TT') {
            platformIndex = j;
            break;
          }
        }
        
        const productName = platformIndex > 0 ? parts.slice(0, platformIndex).join('_') : parts[0];
        const locale = platformIndex >= 0 && platformIndex + 1 < parts.length ? parts[platformIndex + 1] : '';
        
        Logger.log('     🔍 Searching Drive for: ' + ad.adName + ' (product: ' + productName + ', locale: ' + locale + ')');
        
        // Call the same function that Weekly Report uses
        const driveFile = findAdCreativeFile(ad.adName, productName, locale);
        
        if (driveFile) {
          ad.driveUrl = driveFile.url;
          Logger.log('     📎 Found URL in Drive: ' + ad.driveUrl);
        } else {
          Logger.log('     ⚠️ No file found in Drive');
        }
        
        const adRowData = [
          emoji,
          '  ' + ad.adName,  // This will be replaced if we have a driveUrl
          '€' + ad.spend.toFixed(2),
          ad.roas.toFixed(2),
          '€' + ad.cpa.toFixed(2),
          ad.purchases,
          isVideo && ad.hookRate ? (ad.hookRate * 100).toFixed(2) + '%' : '',
          isVideo && ad.holdRate ? (ad.holdRate * 100).toFixed(2) + '%' : '',
          '',
          ''
        ];
        
        testSheet.getRange(row, 1, 1, headers.length).setValues([adRowData]);
        
        // Add hyperlink to ad name if driveUrl exists
        if (ad.driveUrl) {
          const adNameCell = testSheet.getRange(row, 2); // Column B
          const displayName = '  ' + ad.adName;
          adNameCell.setFormula('=HYPERLINK("' + ad.driveUrl + '", "' + displayName.replace(/"/g, '""') + '")');
          adNameCell.setFontColor('#1a73e8'); // Blue hyperlink color
        }
        testSheet.getRange(row, 1, 1, headers.length)
          .setBackground(ad.adCategoryColor)
          .setFontWeight('normal')
          .setVerticalAlignment('middle');
        
        testSheet.getRange(row, 3, 1, 6).setHorizontalAlignment('center');
        
        row++;
      });
      
      // Add spacing between batches
      row++;
    });
    
    // Add thick separator between products
    if (productIndex < Object.keys(resultsByProduct).length - 1) {
      testSheet.getRange(row, 1, 1, headers.length)
        .setBackground('#000000')
        .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      row++;
      row++; // Extra spacing
    }
  });
  
  // Freeze header row
  testSheet.setFrozenRows(4);
}


/**
 * Normalize product names to handle rebrands, variants, and case variations
 */
function normalizeProductName(productName) {
  if (!productName) return productName;
  
  // ✅ STEP 1: Strip the product code prefix first (A03_, B02_, etc.)
  const cleanName = productName.replace(/^[A-Z]\d+_/, '');
  
  // ✅ STEP 2: Case-insensitive canonical mappings (check this FIRST)
  const canonicalNames = {
    'wavemax': 'WaveMax',
    'auranaturals': 'AuraNaturals',
    'wellheater': 'WellHeater',
    'camtrix': 'CamTrix',
    'aliveblue': 'AliveBlue',
    'katuchef': 'KatuChef',
    'oribreeze': 'OriBreeze',
    'epicooler': 'EpiCooler',
    'wiggydog': 'WiggyDog',
    'guardality': 'Guardality',
    'tellystick': 'TellyStick',
    'heatoor': 'Heatoor',
    'cleanlixmold': 'CleanlixMold',
    'cleanlixpowder': 'CleanlixPowder',
    'cleanlix': 'Cleanlix',
    'therawolfneuro': 'TheraWolfNeuro',
    'therawolfrelief': 'TheraWolfRelief',
    'therawolf': 'TheraWolf',
    'clairu_m2': 'Clairu_M2',
    'clairu': 'Clairu_M2',
    'ozoori': 'Ozoori',
    'lumenlight': 'LumenLight',
    'armouredcard': 'Guardality'
  };
  
  // Check canonical name (case-insensitive)
  const lowerCleanName = cleanName.toLowerCase();
  if (canonicalNames[lowerCleanName]) {
    return canonicalNames[lowerCleanName];
  }
  
  // ✅ STEP 3: ClickUp list name mappings (exact matches)
  const listNameMappings = {
    'Mini Wifi Camera': 'CamTrix',
    'TV Antena': 'WaveMax',
    'Heater': 'WellHeater',
    'AC Heater': 'Heatoor',
    'Hydrogen Bottle': 'AliveBlue',
    'Card Protector': 'Guardality',
    'Cutting board': 'KatuChef',
    'Cutting Board': 'KatuChef',
    'TV Stick': 'TellyStick',
    'Mini Portable Cooler': 'OriBreeze',
    'AC Cooler': 'EpiCooler',
    'Mold Remover Gel': 'CleanlixMold',
    'Robot Puppy': 'WiggyDog',
    'Fridge Deodorizer': 'Ozoori',
    'Fridge deodoriser': 'Ozoori',
    'Head FlashLight': 'LumenLight',
    'NeuroBalm': 'TheraWolfNeuro',
    'ReliefBalm': 'TheraWolfRelief',
    'Ionic Air Purifier': 'Clairu_M2'
  };
  
  // Check exact list name matches
  if (listNameMappings[cleanName]) {
    return listNameMappings[cleanName];
  }
  
  // ✅ STEP 4: Handle special compound names (Clairu, TheraWolf, Cleanlix)
  if (lowerCleanName.includes('clairu') || lowerCleanName.includes('ionic')) {
    return 'Clairu_M2';
  }
  
  if (lowerCleanName.includes('therawolf')) {
    if (lowerCleanName.includes('neuro')) {
      return 'TheraWolfNeuro';
    }
    if (lowerCleanName.includes('relief')) {
      return 'TheraWolfRelief';
    }
    return 'TheraWolf';
  }
  
  if (lowerCleanName.includes('cleanlix')) {
    if (lowerCleanName.includes('mold')) {
      return 'CleanlixMold';
    }
    if (lowerCleanName.includes('powder')) {
      return 'CleanlixPowder';
    }
    return 'Cleanlix';
  }
  
  // ✅ STEP 5: Final fallback - check if it's a known product with wrong case
  // This catches things like "Wavemax" that might slip through
  for (const key in canonicalNames) {
    if (lowerCleanName === key || lowerCleanName.replace(/[_-]/g, '') === key.replace(/[_-]/g, '')) {
      return canonicalNames[key];
    }
  }
  
  // ✅ STEP 6: Default - return cleaned name as-is
  return cleanName;
}

/**
 * Extract product name from ad name
 * Returns the product name or 'Unknown' if not found
 * ✅ FIXED: Now recognizes compound product names like Clairu_M2, TheraWolf_Neuro
 * ✅ FIXED: Strips _orig suffixes and ad ID patterns (e.g., _FB_US_266_3_orig)
 */
function extractProductFromAdName(adName) {
  if (!adName) return 'Unknown';
  
  let cleanAdName = adName.toString().trim();
  
  // Remove "(Test X)" prefixes at the very beginning
  cleanAdName = cleanAdName.replace(/^\(Test [A-Z]\)\s*/i, '');
  
  // ✅ STEP 1: Remove leading brackets ([unititled], [untitled], etc.)
  cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
  
  // ✅ STEP 2: Remove parentheses prefixes: (blank) rejected:, (anything):
  cleanAdName = cleanAdName.replace(/^\([^)]*\)\s*/g, '');
  
  // ✅ STEP 3: Remove status word prefixes: rejected:, not delivering:, etc.
  cleanAdName = cleanAdName.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
  
  // ✅ STEP 4: Remove leading brackets AGAIN ([copy:gift], [copy:longlasting], [old copy], etc.)
  cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
  
  // Remove "Test - " prefix if present
  if (cleanAdName.startsWith('Test - ')) {
    cleanAdName = cleanAdName.substring(7);
  }
  
  // Remove "Post ID || " prefix if present
  if (cleanAdName.startsWith('Post ID || ')) {
    cleanAdName = cleanAdName.substring(11);
  }
  
  // Extract from || notation if present
  if (cleanAdName.includes('||')) {
    const parts = cleanAdName.split('||');
    for (let i = 0; i < parts.length; i++) {
      if (parts[i].includes('_')) {
        cleanAdName = parts[i].trim();
        break;
      }
    }
  }
  
  // ✅ Apply cleaning AGAIN after || extraction
  // STEP 1: Remove leading brackets
  cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
  
  // STEP 2: Remove parentheses prefixes
  cleanAdName = cleanAdName.replace(/^\([^)]*\)\s*/g, '');
  
  // STEP 3: Remove status word prefixes
  cleanAdName = cleanAdName.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
  
  // STEP 4: Remove leading brackets AGAIN
  cleanAdName = cleanAdName.replace(/^\[[^\]]*\]\s*/g, '');
  
  // Remove "Post ID || " if present after || extraction
  if (cleanAdName.startsWith('Post ID || ')) {
    cleanAdName = cleanAdName.substring(11);
  }
  
  // ✅ NEW LOGIC: Check for compound product names FIRST
  // List of known compound products (underscore is part of the product name)
  const compoundProducts = [
    'Clairu_M2',
    'TheraWolf_Neuro', 
    'TheraWolfNeuro',
    'TheraWolf_Relief',
    'TheraWolfRelief',
    'Cleanlix_Mold',
    'CleanlixMold',
    'Cleanlix_Powder',
    'CleanlixPowder',
    'DermaSonic+'
  ];
  
  // Check if ad name starts with any compound product
  for (let i = 0; i < compoundProducts.length; i++) {
    const compound = compoundProducts[i];
    // Check with underscore separator (Clairu_M2_FB_...)
    if (cleanAdName.startsWith(compound + '_')) {
      return normalizeProductName(compound);
    }
    // Check without underscore for legacy formats
    const compoundNoUnderscore = compound.replace(/_/g, '');
    if (cleanAdName.startsWith(compoundNoUnderscore + '_')) {
      return normalizeProductName(compound);
    }
  }
  
  // ✅ IMPROVED LOGIC: Extract just the product name (first segment before platform)
  // Match: ProductName_FB_ or ProductName_YT_ etc.
  const productMatch = cleanAdName.match(/^([A-Za-z0-9]+)_(?:FB|YT|GDN|TT|SNAP|IG)_/);
  if (productMatch) {
    let product = productMatch[1];
    product = normalizeProductName(product);
    return product;
  }
  
  return 'Unknown';
}


/**
 * ✅ FIXED VERSION of getProductsFromCSV
 * 
 * This function reads the "CSV Data" sheet and extracts unique products
 * for the "Analyze Creative Tests" popup.
 * 
 * CRITICAL: Must apply the SAME cleaning logic as extractProductFromAdName!
 */
function getProductsFromCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet || dataSheet.getLastRow() === 0) {
    return [];
  }
  
  const data = dataSheet.getDataRange().getValues();
  
  // Column indices
  const adNameCol = 0;
  const productLabelCol = 8; // Column I
  
  const products = new Set();
  
  // Process all rows
  for (let i = 1; i < data.length; i++) {
    let adName = data[i][adNameCol] || '';
    const label = data[i][productLabelCol] || '';
    
    if (!adName) continue;
    
    // ✅ APPLY FULL CLEANING SEQUENCE (same as extractProductFromAdName)
    
    // Remove "(Test X)" prefixes
    adName = adName.replace(/^\(Test [A-Z]\)\s*/i, '');
    
    // ✅ STEP 1: Remove leading brackets ([unititled], [untitled], etc.)
    adName = adName.replace(/^\[[^\]]*\]\s*/g, '');
    
    // ✅ STEP 2: Remove parentheses prefixes: (blank) rejected:, (anything):
    adName = adName.replace(/^\([^)]*\)\s*/g, '');
    
    // ✅ STEP 3: Remove status word prefixes: rejected:, not delivering:, etc.
    adName = adName.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
    
    // ✅ STEP 4: Remove leading brackets AGAIN ([copy:gift], [copy:longlasting], etc.)
    adName = adName.replace(/^\[[^\]]*\]\s*/g, '');
    
    // Remove "Test - " prefix
    if (adName.startsWith('Test - ')) {
      adName = adName.substring(7);
    }
    
    // Extract clean ad name from || notation
    if (adName.includes('||')) {
      const parts = adName.split('||');
      for (let j = 0; j < parts.length; j++) {
        if (parts[j].includes('_')) {
          adName = parts[j].trim();
          break;
        }
      }
    }
    
    // ✅ APPLY CLEANING AGAIN AFTER || EXTRACTION
    
    // STEP 1: Remove leading brackets
    adName = adName.replace(/^\[[^\]]*\]\s*/g, '');
    
    // STEP 2: Remove parentheses prefixes
    adName = adName.replace(/^\([^)]*\)\s*/g, '');
    
    // STEP 3: Remove status word prefixes
    adName = adName.replace(/^(rejected|not delivering|delivering|active|paused|pending|approved|disapproved):\s*/gi, '');
    
    // STEP 4: Remove leading brackets AGAIN
    adName = adName.replace(/^\[[^\]]*\]\s*/g, '');
    
    // Split by underscore and find platform
    const parts = adName.split('_');
    let platformIndex = -1;
    for (let j = 0; j < parts.length; j++) {
      if (parts[j] === 'FB' || parts[j] === 'YT' || parts[j] === 'GDN' || parts[j] === 'TT') {
        platformIndex = j;
        break;
      }
    }
    
    if (platformIndex === -1) continue;
    
    // Product is everything BEFORE the platform
    let product = platformIndex > 0 ? parts.slice(0, platformIndex).join('_') : parts[0];
    
    // Normalize product name (handle rebrands)
    product = normalizeProductName(product);
    
    // ✅ CRITICAL FIX: Normalize case for WellHeater
    if (product.toLowerCase() === 'wellheater') {
      product = 'WellHeater';
    }
    
    // Add label to product name if it exists
    if (label) {
      product = `${product} (${label})`;
    }
    
    products.add(product);
  }
  
  // Convert to array and sort alphabetically
  return Array.from(products).sort();
}


function updateDataInfo(adsLoaded, products) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let infoSheet = ss.getSheetByName('Data Info');
  
  if (!infoSheet) {
    infoSheet = ss.insertSheet('Data Info');
  } else {
    infoSheet.clear();
  }
  
  const now = new Date();
  
  infoSheet.getRange(1, 1).setValue('📊 DATA UPLOAD INFO')
    .setFontWeight('bold')
    .setFontSize(14);
  
  // ✅ Clean up product list - remove duplicates and sort
  const uniqueProducts = [...new Set(products)].sort();
  const productListText = uniqueProducts.join(', ');
  
  infoSheet.getRange(3, 1, 5, 2).setValues([
    ['Last Upload:', now.toLocaleString()],
    ['Total Ads Loaded:', adsLoaded],
    ['Products Detected:', uniqueProducts.length],
    ['Product List:', productListText],
    ['Status:', '✅ Ready for analysis']
  ]);
  
  infoSheet.getRange(3, 1, 5, 1).setFontWeight('bold');
  infoSheet.setColumnWidth(1, 150);
  infoSheet.setColumnWidth(2, 600); // Wider column for product list
  
  // Enable text wrapping for product list cell (row 6, column B)
  infoSheet.getRange(6, 2).setWrap(true);
  infoSheet.getRange(6, 2).setVerticalAlignment('top');
}

function showDataInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let infoSheet = ss.getSheetByName('Data Info');
  
  if (!infoSheet) {
    SpreadsheetApp.getUi().alert('No Data', 'No CSV data has been uploaded yet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  infoSheet.activate();
}

function sendCreativeTestsToSlack() {
  // ✅ Always get fresh Slack config when opening this dialog
  CacheService.getScriptCache().remove('atria_slack');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  
  if (!settings.slackBotToken || !settings.slackChannelIdCreative) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Slack Bot Token or Creative Tests Channel ID not configured. Please go to Settings first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Show source selection dialog
  const html = buildCreativeTestsSourceDialog();
  SpreadsheetApp.getUi().showModalDialog(html, '📈 Send Creative Tests to Slack');
}

function buildCreativeTestsSourceDialog() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; }
        body { 
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
          padding: 16px; 
          margin: 0;
          color: #333;
        }
        .source-section {
          margin-bottom: 16px;
          padding-bottom: 16px;
          border-bottom: 1px solid #e0e0e0;
        }
        .source-section h3 { margin: 0 0 12px 0; font-size: 14px; }
        .source-option {
          padding: 12px;
          border: 2px solid #e0e0e0;
          border-radius: 8px;
          margin-bottom: 8px;
          cursor: pointer;
          transition: all 0.2s;
        }
        .source-option:hover { border-color: #4285f4; background: #f8f9fa; }
        .source-option.selected { border-color: #4285f4; background: #e8f0fe; }
        .source-option label { display: flex; align-items: center; gap: 10px; cursor: pointer; }
        .source-option input[type="radio"] { width: 18px; height: 18px; }
        .source-title { font-weight: 500; }
        .source-desc { font-size: 12px; color: #666; margin-top: 4px; }
        .week-select { 
          margin-top: 12px; 
          padding: 12px;
          background: #f8f9fa;
          border-radius: 4px;
          display: none;
        }
        .week-select.visible { display: block; }
        .week-select label { font-size: 13px; font-weight: 500; }
        .week-select select { 
          width: 100%; 
          padding: 8px; 
          margin-top: 6px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 14px;
        }
        .button-row { 
          display: flex; 
          justify-content: flex-end; 
          gap: 8px;
          margin-top: 16px;
        }
        .btn {
          padding: 10px 20px;
          border-radius: 4px;
          font-size: 14px;
          cursor: pointer;
          border: none;
        }
        .btn-cancel { background: #f5f5f5; border: 1px solid #ddd; }
        .btn-cancel:hover { background: #e8e8e8; }
        .btn-next { background: #4285f4; color: white; }
        .btn-next:hover { background: #3367d6; }
        .btn-next:disabled { background: #ccc; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <div class="source-section">
        <h3>📂 Select data source:</h3>
        
        <div class="source-option selected" onclick="selectSource('tab')">
          <label>
            <input type="radio" name="source" value="tab" checked>
            <div>
              <div class="source-title">Creative Tests tab (this sheet)</div>
              <div class="source-desc">Use data from the consolidated Creative Tests tab</div>
            </div>
          </label>
        </div>
        
        <div class="source-option" onclick="selectSource('spreadsheets')">
          <label>
            <input type="radio" name="source" value="spreadsheets">
            <div>
              <div class="source-title">Product spreadsheets</div>
              <div class="source-desc">Pull latest data from individual product spreadsheets</div>
            </div>
          </label>
        </div>
        
        <div class="source-option" onclick="selectSource('summary')">
          <label>
            <input type="radio" name="source" value="summary">
            <div>
              <div class="source-title">📊 Weekly Summary Only</div>
              <div class="source-desc">Send summary footer with no top performers, no tests, and all product links</div>
            </div>
          </label>
        </div>
        
        <div class="week-select" id="weekSelect">
          <label>Select week:</label>
          <select id="weekDropdown">
            <option value="latest">Latest week (auto-detect)</option>
          </select>
        </div>
      </div>
      
      <div class="button-row">
        <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button class="btn btn-next" id="nextBtn" onclick="loadProducts()">Next →</button>
      </div>
      
      <script>
        function selectSource(source) {
          document.querySelectorAll('.source-option').forEach(el => el.classList.remove('selected'));
          event.currentTarget.classList.add('selected');
          document.querySelector('input[value="' + source + '"]').checked = true;
          document.getElementById('weekSelect').classList.toggle('visible', source === 'spreadsheets');
        }
        
        function loadProducts() {
          const source = document.querySelector('input[name="source"]:checked').value;
          const week = document.getElementById('weekDropdown').value;
          
          document.getElementById('nextBtn').disabled = true;
          document.getElementById('nextBtn').textContent = 'Loading...';
          
          if (source === 'summary') {
            // Go directly to summary sending
            google.script.run
              .withSuccessHandler(function() { google.script.host.close(); })
              .withFailureHandler(function(error) {
                alert('Error: ' + error);
                document.getElementById('nextBtn').disabled = false;
                document.getElementById('nextBtn').textContent = 'Next →';
              })
              .sendCreativeTestsSummaryFooter();
          } else {
            google.script.run
              .withSuccessHandler(function() {})
              .withFailureHandler(function(error) {
                alert('Error: ' + error);
                document.getElementById('nextBtn').disabled = false;
                document.getElementById('nextBtn').textContent = 'Next →';
              })
              .showCreativeTestsProductSelector(source, week);
          }
        }
      </script>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(320);
}

function showCreativeTestsProductSelector(source, week) {
  let productList;
  let weekLabel;
  
  if (source === 'tab') {
    const result = getProductsFromCreativeTestsTab();
    productList = result.products;
    weekLabel = result.weekLabel;
  } else {
    const result = getProductsFromSpreadsheets(week);
    productList = result.products;
    weekLabel = result.weekLabel;
  }
  
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('CT_SLACK_SOURCE', source);
  scriptProps.setProperty('CT_SLACK_WEEK', week);
  scriptProps.setProperty('CT_SLACK_WEEK_LABEL', weekLabel || '');
  
  const html = buildCreativeTestsProductDialog(productList, source, weekLabel);
  SpreadsheetApp.getUi().showModalDialog(html, '📈 Send Creative Tests to Slack');
}

function getProductsFromCreativeTestsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  if (!testSheet) {
    return { products: [], weekLabel: null };
  }
  
  const data = testSheet.getDataRange().getValues();
  const results = parseCreativeTestsSheet(data, testSheet);
  
  const scriptProps = PropertiesService.getScriptProperties();
  const weekLabel = scriptProps.getProperty('LAST_ANALYZED_WEEK');
  
  const filteredResults = results.filter(batch => 
    batch.category.includes('Top Performer') || batch.category.includes('Mid Performer')
  );
  
  const productStats = {};
  filteredResults.forEach(batch => {
    if (!productStats[batch.product]) {
      productStats[batch.product] = { batches: 0, topPerformers: 0, hasData: true };
    }
    productStats[batch.product].batches++;
    productStats[batch.product].topPerformers += batch.topPerformerAds.length;
  });
  
  const productList = Object.keys(productStats).sort().map(product => ({
    name: product,
    batches: productStats[product].batches,
    topPerformers: productStats[product].topPerformers,
    hasData: true
  }));
  
  return { products: productList, weekLabel: weekLabel };
}

function getProductsFromSpreadsheets(week) {
  const productSpreadsheetIds = getProductSpreadsheetIds();
  const productList = [];
  let detectedWeekLabel = null;
  
  Object.keys(productSpreadsheetIds).sort().forEach(productName => {
    const spreadsheetId = productSpreadsheetIds[productName];
    
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheets = spreadsheet.getSheets();
      
      const weekTabs = sheets.filter(s => /^W\d+\s*\(/.test(s.getName()))
        .sort((a, b) => {
          const weekA = parseInt(a.getName().match(/W(\d+)/)[1]);
          const weekB = parseInt(b.getName().match(/W(\d+)/)[1]);
          return weekB - weekA;
        });
      
      if (weekTabs.length === 0) {
        productList.push({
          name: productName,
          batches: 0,
          topPerformers: 0,
          hasData: false,
          noWeekTabs: true
        });
        return;
      }
      
      let targetTab;
      if (week === 'latest') {
        targetTab = weekTabs[0];
      } else {
        targetTab = weekTabs.find(s => s.getName().includes(week)) || weekTabs[0];
      }
      
      if (!detectedWeekLabel) {
        detectedWeekLabel = targetTab.getName();
      }
      
      const data = targetTab.getDataRange().getValues();
      const stats = parseProductSpreadsheetTab(data);
      
      productList.push({
        name: productName,
        batches: stats.batches,
        topPerformers: stats.topPerformers,
        hasData: stats.hasData,
        noCreativeTests: stats.noCreativeTests
      });
      
    } catch (e) {
      Logger.log('Error accessing ' + productName + ': ' + e.toString());
      productList.push({
        name: productName,
        batches: 0,
        topPerformers: 0,
        hasData: false,
        error: true
      });
    }
  });
  
  return { products: productList, weekLabel: detectedWeekLabel };
}

function parseProductSpreadsheetTab(data) {
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const cellValue = String(data[i][0] || '').toLowerCase();
    if (cellValue.includes('no creative tests were conducted')) {
      return { batches: 0, topPerformers: 0, hasData: false, noCreativeTests: true };
    }
  }
  
  let batches = 0;
  let topPerformers = 0;
  
  for (let i = 0; i < data.length; i++) {
    const col0 = String(data[i][0] || '');
    
    if (col0.includes('Top Performer') || col0.includes('Mid Performer') || col0.includes('Underperformer')) {
      batches++;
    }
    
    if (col0.includes('🟢')) {
      topPerformers++;
    }
  }
  
  return { 
    batches: batches, 
    topPerformers: topPerformers, 
    hasData: batches > 0 || topPerformers > 0,
    noCreativeTests: false
  };
}

function buildCreativeTestsProductDialog(productList, source, weekLabel) {
  const sourceLabel = source === 'tab' ? 'Creative Tests tab' : 'Product spreadsheets';
  
  const productRows = productList.map((p, idx) => {
    const canSelect = p.hasData && p.topPerformers > 0;
    let statusIcon = '';
    let statusText = '';
    
    if (p.error) {
      statusIcon = '❌';
      statusText = 'Error accessing spreadsheet';
    } else if (p.noWeekTabs) {
      statusIcon = '⚠️';
      statusText = 'No week tabs found';
    } else if (p.noCreativeTests) {
      statusIcon = '📭';
      statusText = 'No creative tests this week';
    } else if (p.topPerformers === 0) {
      statusIcon = '⚠️';
      statusText = '0 top performers';
    } else {
      statusText = p.batches + ' batch' + (p.batches !== 1 ? 'es' : '') + ', ' + p.topPerformers + ' top performer' + (p.topPerformers !== 1 ? 's' : '');
    }
    
    const disabledAttr = !canSelect ? 'disabled' : '';
    const checkedAttr = canSelect ? 'checked' : '';
    const rowClass = !canSelect ? 'no-performers' : '';
    
    return '<div class="product-row ' + rowClass + '">' +
      '<label>' +
      '<input type="checkbox" name="product" value="' + p.name + '" ' + checkedAttr + ' ' + disabledAttr + '>' +
      '<span class="product-name">' + p.name + ' ' + statusIcon + '</span>' +
      '<span class="product-stats">' + statusText + '</span>' +
      '</label></div>';
  }).join('');
  
  const htmlContent = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    '* { box-sizing: border-box; }' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; margin: 0; color: #333; }' +
    '.header { margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid #e0e0e0; }' +
    '.header h3 { margin: 0 0 4px 0; font-size: 14px; }' +
    '.header .source-info { font-size: 12px; color: #666; margin-bottom: 12px; }' +
    '.select-buttons { margin-bottom: 12px; }' +
    '.select-buttons button { padding: 6px 12px; margin-right: 8px; border: 1px solid #ddd; border-radius: 4px; background: #f5f5f5; cursor: pointer; font-size: 12px; }' +
    '.select-buttons button:hover { background: #e8e8e8; }' +
    '.product-list { max-height: 280px; overflow-y: auto; border: 1px solid #e0e0e0; border-radius: 4px; padding: 8px; margin-bottom: 16px; }' +
    '.product-row { padding: 8px; border-bottom: 1px solid #f0f0f0; }' +
    '.product-row:last-child { border-bottom: none; }' +
    '.product-row label { display: flex; align-items: center; cursor: pointer; gap: 8px; }' +
    '.product-row.no-performers { opacity: 0.5; }' +
    '.product-row.no-performers label { cursor: not-allowed; }' +
    '.product-name { font-weight: 500; flex: 1; }' +
    '.product-stats { font-size: 11px; color: #666; }' +
    'input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; }' +
    '.button-row { display: flex; justify-content: space-between; align-items: center; padding-top: 12px; border-top: 1px solid #e0e0e0; }' +
    '.btn { padding: 10px 20px; border-radius: 4px; font-size: 14px; cursor: pointer; border: none; }' +
    '.btn-back { background: #f5f5f5; border: 1px solid #ddd; }' +
    '.btn-back:hover { background: #e8e8e8; }' +
    '.btn-send { background: #4285f4; color: white; }' +
    '.btn-send:hover { background: #3367d6; }' +
    '.btn-send:disabled { background: #ccc; cursor: not-allowed; }' +
    '.selected-count { font-size: 12px; color: #666; margin-bottom: 8px; }' +
    '.right-buttons { display: flex; gap: 8px; }' +
    '</style></head><body>' +
    '<div class="header">' +
    '<h3>Select products to include:</h3>' +
    '<div class="source-info">Source: ' + sourceLabel + (weekLabel ? ' • ' + weekLabel : '') + '</div>' +
    '<div class="select-buttons">' +
    '<button onclick="selectAll()">✅ Select All</button>' +
    '<button onclick="deselectAll()">☐ Deselect All</button>' +
    '</div>' +
    '<div class="selected-count" id="selectedCount">0 products selected</div>' +
    '</div>' +
    '<div class="product-list">' + productRows + '</div>' +
    '<div class="button-row">' +
    '<button class="btn btn-back" onclick="goBack()">← Back</button>' +
    '<div class="right-buttons">' +
    '<button class="btn btn-back" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-send" id="sendBtn" onclick="sendSelected()">Send to Slack</button>' +
    '</div></div>' +
    '<script>' +
    'var source = "' + source + '";' +
    'function updateCount() {' +
    '  var checked = document.querySelectorAll("input[name=product]:checked:not(:disabled)").length;' +
    '  document.getElementById("selectedCount").textContent = checked + " product" + (checked !== 1 ? "s" : "") + " selected";' +
    '  document.getElementById("sendBtn").disabled = checked === 0;' +
    '}' +
    'function selectAll() { document.querySelectorAll("input[name=product]:not(:disabled)").forEach(function(cb) { cb.checked = true; }); updateCount(); }' +
    'function deselectAll() { document.querySelectorAll("input[name=product]").forEach(function(cb) { cb.checked = false; }); updateCount(); }' +
    'function goBack() { google.script.run.sendCreativeTestsToSlack(); }' +
    'function sendSelected() {' +
    '  var selected = Array.from(document.querySelectorAll("input[name=product]:checked")).map(function(cb) { return cb.value; });' +
    '  if (selected.length === 0) { alert("Please select at least one product"); return; }' +
    '  document.getElementById("sendBtn").disabled = true;' +
    '  document.getElementById("sendBtn").textContent = "Sending...";' +
    '  google.script.run.withSuccessHandler(function() { google.script.host.close(); }).withFailureHandler(function(error) { alert("Error: " + error); document.getElementById("sendBtn").disabled = false; document.getElementById("sendBtn").textContent = "Send to Slack"; }).sendCreativeTestsToSlackFromSource(selected, source);' +
    '}' +
    'document.querySelectorAll("input[name=product]").forEach(function(cb) { cb.addEventListener("change", updateCount); });' +
    'updateCount();' +
    '</script></body></html>';
  
  return HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(500);
}

function sendCreativeTestsToSlackFromSource(selectedProducts, source) {
  const scriptProps = PropertiesService.getScriptProperties();
  const weekLabel = scriptProps.getProperty('CT_SLACK_WEEK_LABEL');
  const settings = getSettings();
  const channel = settings.slackChannelIdCreative;
  
  let allResults;
  
  if (source === 'tab') {
    allResults = getCreativeTestDataFromTab(selectedProducts);
  } else {
    allResults = getCreativeTestDataFromSpreadsheets(selectedProducts);
  }
  
  if (!allResults || Object.keys(allResults).length === 0) {
    SpreadsheetApp.getUi().alert('❌ Error', 'No data found for selected products.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const sentMessages = [];
  
  const headerBlocks = [
    { "type": "header", "text": { "type": "plain_text", "text": "📈 CREATIVE TEST RESULTS", "emoji": true } }
  ];
  
  if (weekLabel) {
    headerBlocks.push({ "type": "section", "text": { "type": "mrkdwn", "text": "*" + weekLabel + "*" } });
  }
  
  headerBlocks.push({ "type": "divider" });
  
  const headerMsg = sendSlackBlockMessage(headerBlocks, channel);
  if (headerMsg) sentMessages.push({ ts: headerMsg.ts, channel: headerMsg.channel });
  Utilities.sleep(1000);
  
  const products = Object.keys(allResults).sort();
  products.forEach((product, productIndex) => {
    const productData = allResults[product];
    
    if (productData.topPerformers.length === 0) return;
    
    const productBlocks = [
      { "type": "header", "text": { "type": "plain_text", "text": product.toUpperCase(), "emoji": true } },
      { "type": "section", "text": { "type": "mrkdwn", "text": "`TOP PERFORMERS 🟢`" } }
    ];
    
    // Show only top 3 in main message
    const top3 = productData.topPerformers.slice(0, 3);
    const remaining = productData.topPerformers.slice(3);
    
    top3.forEach(ad => {
      const hookHold = ad.isVideo && ad.hookRate && ad.holdRate 
        ? " | Hook: " + parseFloat(ad.hookRate).toFixed(2) + "% | Hold: " + parseFloat(ad.holdRate).toFixed(2) + "%" : '';
      
      const adNameDisplay = ad.driveUrl ? "<" + ad.driveUrl + "|" + ad.adName + ">" : ad.adName;
      
      productBlocks.push({
        "type": "section",
        "text": { "type": "mrkdwn", "text": "*" + adNameDisplay + "*\n€" + ad.spend + " | ROAS: " + ad.roas + " | CPA: €" + ad.cpa + " | " + ad.purchases + " purchases" + hookHold }
      });
    });
    
    // Add note about more in thread if there are remaining top performers
    if (remaining.length > 0) {
      productBlocks.push({ 
        "type": "context", 
        "elements": [{ "type": "mrkdwn", "text": "_" + remaining.length + " more top performer" + (remaining.length !== 1 ? "s" : "") + " in thread ↓_" }] 
      });
    }
    
    if (productData.midPerformers.length > 0) {
      productBlocks.push({ "type": "context", "elements": [{ "type": "mrkdwn", "text": "🟡 _Mid Performing ads: " + productData.midPerformers.length + "_" }] });
    }
    
    if (productIndex < products.length - 1) {
      productBlocks.push({ "type": "divider" });
    }
    
    const productMsg = sendSlackBlockMessage(productBlocks, channel);
    if (productMsg) sentMessages.push({ ts: productMsg.ts, channel: productMsg.channel });
    Utilities.sleep(500);
    
    // Send remaining top performers to thread
    if (remaining.length > 0 && productMsg) {
      const threadBlocks = [
        { "type": "section", "text": { "type": "mrkdwn", "text": "`MORE TOP PERFORMERS 🟢`" } }
      ];
      
      remaining.forEach(ad => {
        const hookHold = ad.isVideo && ad.hookRate && ad.holdRate 
          ? " | Hook: " + parseFloat(ad.hookRate).toFixed(2) + "% | Hold: " + parseFloat(ad.holdRate).toFixed(2) + "%" : '';
        
        const adNameDisplay = ad.driveUrl ? "<" + ad.driveUrl + "|" + ad.adName + ">" : ad.adName;
        
        threadBlocks.push({
          "type": "section",
          "text": { "type": "mrkdwn", "text": "*" + adNameDisplay + "*\n€" + ad.spend + " | ROAS: " + ad.roas + " | CPA: €" + ad.cpa + " | " + ad.purchases + " purchases" + hookHold }
        });
      });
      
      const threadMsg = sendSlackThreadReply(threadBlocks, channel, productMsg.ts);
      if (threadMsg) sentMessages.push({ ts: threadMsg, channel: channel });
      Utilities.sleep(500);
    }
    
    Utilities.sleep(500);
  });
  
  // ✅ Add weekly summary footer at the end (2 messages)
  Utilities.sleep(1000);
  const { blocks1, blocks2 } = buildWeeklySummaryBlocks();
  
  const summaryMsg1 = sendSlackBlockMessage(blocks1, channel);
  if (summaryMsg1) sentMessages.push({ ts: summaryMsg1.ts, channel: summaryMsg1.channel });
  
  Utilities.sleep(500);
  
  const summaryMsg2 = sendSlackBlockMessage(blocks2, channel);
  if (summaryMsg2) sentMessages.push({ ts: summaryMsg2.ts, channel: summaryMsg2.channel });
  
  scriptProps.setProperty('LAST_CREATIVE_TEST_MESSAGES', JSON.stringify(sentMessages));
  
  SpreadsheetApp.getUi().alert('✅ Success', 'Creative Tests sent to Slack!\n\nSent ' + products.length + ' product reports + weekly summary.', SpreadsheetApp.getUi().ButtonSet.OK);
  
  return { success: true, count: products.length };
}

/**
 * Build the weekly summary blocks for Slack
 * Returns array of Slack blocks
 */
function buildWeeklySummaryBlocks() {
  const productSpreadsheetIds = getProductSpreadsheetIds();
  const allProducts = Object.keys(productSpreadsheetIds).sort();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  const noTopPerformers = [];
  const noTests = [];
  const productsWithTopPerformers = new Set();
  const productBatchCounts = {};
  
  if (testSheet) {
    const data = testSheet.getDataRange().getValues();
    const results = parseCreativeTestsSheet(data, testSheet);
    
    results.forEach(batch => {
      const product = batch.product;
      if (!productBatchCounts[product]) {
        productBatchCounts[product] = 0;
      }
      productBatchCounts[product]++;
      
      if (batch.topPerformerAds && batch.topPerformerAds.length > 0) {
        productsWithTopPerformers.add(product);
      }
    });
  }
  
  allProducts.forEach(product => {
    const productUpper = product.toUpperCase();
    const batchCount = productBatchCounts[product] || productBatchCounts[productUpper] || 0;
    
    if (batchCount > 0) {
      if (!productsWithTopPerformers.has(product) && !productsWithTopPerformers.has(productUpper)) {
        noTopPerformers.push({ name: product, batches: batchCount });
      }
    } else {
      noTests.push(product);
    }
  });
  
  const blocks = [
    { "type": "header", "text": { "type": "plain_text", "text": "📊 WEEKLY SUMMARY", "emoji": true } },
    { "type": "divider" }
  ];
  
  if (noTopPerformers.length > 0) {
    let noTopText = "*⚠️ No Top Performers This Week:*\n";
    noTopPerformers.forEach(p => {
      const spreadsheetId = productSpreadsheetIds[p.name];
      const link = spreadsheetId ? "https://docs.google.com/spreadsheets/d/" + spreadsheetId : null;
      const productDisplay = link ? "<" + link + "|" + p.name + ">" : p.name;
      noTopText += "• " + productDisplay + " (" + p.batches + " batch" + (p.batches !== 1 ? "es" : "") + " tested)\n";
    });
    blocks.push({ "type": "section", "text": { "type": "mrkdwn", "text": noTopText } });
  }
  
  if (noTests.length > 0) {
    let noTestsText = "*📭 No Tests This Week:*\n";
    noTests.forEach(product => {
      const spreadsheetId = productSpreadsheetIds[product];
      const link = spreadsheetId ? "https://docs.google.com/spreadsheets/d/" + spreadsheetId : null;
      const productDisplay = link ? "<" + link + "|" + product + ">" : product;
      noTestsText += "• " + productDisplay + "\n";
    });
    blocks.push({ "type": "section", "text": { "type": "mrkdwn", "text": noTestsText } });
  }
  
  blocks.push({ "type": "divider" });
  
  let allSheetsText = "*📁 All Product Sheets:*\n";
  const productLinks = allProducts.map(product => {
    const spreadsheetId = productSpreadsheetIds[product];
    const link = "https://docs.google.com/spreadsheets/d/" + spreadsheetId;
    return "<" + link + "|" + product + ">";
  });
  allSheetsText += productLinks.join(" | ");
  
  blocks.push({ "type": "section", "text": { "type": "mrkdwn", "text": allSheetsText } });
  
  return blocks;
}

/**
 * Get top performer count from a product's previous week tab
 */
function getPreviousWeekTopPerformers(spreadsheetId, currentWeekAbbrev) {
  try {
    // Parse current week number (e.g., "W52 (12/22 - 12/28)" -> 52)
    const weekMatch = currentWeekAbbrev.match(/W(\d+)/);
    if (!weekMatch) return { count: 0, exists: false };
    
    const currentWeekNum = parseInt(weekMatch[1]);
    const prevWeekNum = currentWeekNum === 1 ? 52 : currentWeekNum - 1;
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    // Find previous week tab (could be "W51 (12/15 - 12/21)" or similar)
    let prevSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i].getName();
      if (sheetName.match(new RegExp('^W' + prevWeekNum + '\\b'))) {
        prevSheet = sheets[i];
        break;
      }
    }
    
    if (!prevSheet) {
      return { count: 0, exists: false };
    }
    
    // Count top performers (rows with 🟢 in column A)
    const data = prevSheet.getDataRange().getValues();
    let topCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const cellA = String(data[i][0]);
      const cellB = String(data[i][1]);
      
      // Count individual ad rows that are top performers (🟢 emoji, not batch headers)
      if (cellA.includes('🟢') && !cellB.includes(' | ')) {
        topCount++;
      }
    }
    
    return { count: topCount, exists: true };
  } catch (e) {
    Logger.log('Error getting previous week data for ' + spreadsheetId + ': ' + e.message);
    return { count: 0, exists: false };
  }
}

/**
 * Build the weekly summary blocks for Slack (returns array of 2 block sets for 2 messages)
 */
function buildWeeklySummaryBlocks() {
  const productSpreadsheetIds = getProductSpreadsheetIds();
  const allProducts = Object.keys(productSpreadsheetIds).sort();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  const scriptProps = PropertiesService.getScriptProperties();
  const weekLabel = scriptProps.getProperty('LAST_ANALYZED_WEEK') || 'This Week';
  const weekAbbrev = getAbbreviatedWeekName(weekLabel);
  
  // Track products by category
  const withTopPerformers = [];  // Products with top performers
  const noTopPerformers = [];    // Tested but no top performers
  const noTests = [];            // Not tested at all
  
  const productBatchCounts = {};
  const productTopCounts = {};
  
  if (testSheet) {
    const data = testSheet.getDataRange().getValues();
    let currentProduct = null;
    
    for (let i = 0; i < data.length; i++) {
      const cellA = String(data[i][0]);
      const cellB = String(data[i][1]);
      
      // Product header
      if (cellA.includes('📦')) {
        currentProduct = cellA.replace('📦', '').trim();
        if (!productBatchCounts[currentProduct]) {
          productBatchCounts[currentProduct] = 0;
          productTopCounts[currentProduct] = 0;
        }
        continue;
      }
      
      // Batch row (has " | " in name)
      if (currentProduct && cellB.includes(' | ')) {
        productBatchCounts[currentProduct]++;
        continue;
      }
      
      // Individual ad row - count top performers
      if (currentProduct && cellA.includes('🟢') && !cellB.includes(' | ')) {
        productTopCounts[currentProduct]++;
      }
    }
  }
  
  // Categorize products and get previous week comparison
  allProducts.forEach(product => {
    const productUpper = product.toUpperCase();
    const batchCount = productBatchCounts[product] || productBatchCounts[productUpper] || 0;
    const topCount = productTopCounts[product] || productTopCounts[productUpper] || 0;
    
    if (batchCount > 0) {
      if (topCount > 0) {
        // Get previous week comparison
        const spreadsheetId = productSpreadsheetIds[product];
        const prevWeek = getPreviousWeekTopPerformers(spreadsheetId, weekAbbrev);
        
        let arrow = '➡️';
        let diff = 0;
        let isNew = false;
        
        if (!prevWeek.exists) {
          arrow = '🔺';
          diff = topCount;
          isNew = true;
        } else if (topCount > prevWeek.count) {
          arrow = '🔺';
          diff = topCount - prevWeek.count;
        } else if (topCount < prevWeek.count) {
          arrow = '🔻';
          diff = prevWeek.count - topCount;
        }
        
        withTopPerformers.push({
          name: product,
          topCount: topCount,
          arrow: arrow,
          diff: diff,
          isNew: isNew,
          prevWeekNum: weekAbbrev.match(/W(\d+)/)?.[1] === '1' ? 52 : parseInt(weekAbbrev.match(/W(\d+)/)?.[1] || 1) - 1
        });
      } else {
        noTopPerformers.push({ name: product, batches: batchCount });
      }
    } else {
      noTests.push(product);
    }
  });
  
  // Sort by top count descending
  withTopPerformers.sort((a, b) => b.topCount - a.topCount);
  
  // === BUILD MESSAGE 1: Performance Overview ===
  const blocks1 = [];
  
  // Header with week label
  blocks1.push({ 
    "type": "header", 
    "text": { "type": "plain_text", "text": "📊 WEEKLY SUMMARY - " + weekAbbrev, "emoji": true } 
  });
  
  // Stats header
  const statsText = "✨ " + withTopPerformers.length + " products with top performers | ⚠️ " + noTopPerformers.length + " tested without top performers | 📭 " + noTests.length + " not tested";
  blocks1.push({ "type": "section", "text": { "type": "mrkdwn", "text": statsText } });
  
  blocks1.push({ "type": "divider" });
  
  // Top Performers section
  if (withTopPerformers.length > 0) {
    let topText = "*🏆 Top Performers This Week:*\n";
    withTopPerformers.forEach(p => {
      const spreadsheetId = productSpreadsheetIds[p.name];
      const link = spreadsheetId ? "https://docs.google.com/spreadsheets/d/" + spreadsheetId : null;
      const productDisplay = link ? "<" + link + "|" + p.name + ">" : p.name;
      
      let comparison = '';
      if (p.arrow === '➡️') {
        comparison = " ➡️ (same as W" + p.prevWeekNum + ")";
      } else if (p.isNew) {
        comparison = " 🔺 (+" + p.diff + " vs W" + p.prevWeekNum + ") 🆕";
      } else if (p.arrow === '🔺') {
        comparison = " 🔺 (+" + p.diff + " vs W" + p.prevWeekNum + ")";
      } else {
        comparison = " 🔻 (-" + p.diff + " vs W" + p.prevWeekNum + ")";
      }
      
      const performerWord = p.topCount === 1 ? "top performer" : "top performers";
      topText += "• " + productDisplay + " - " + p.topCount + " " + performerWord + comparison + "\n";
    });
    blocks1.push({ "type": "section", "text": { "type": "mrkdwn", "text": topText } });
  }
  
  // No Top Performers section
  if (noTopPerformers.length > 0) {
    let noTopText = "*⚠️ No Top Performers This Week:*\n";
    noTopPerformers.forEach(p => {
      const spreadsheetId = productSpreadsheetIds[p.name];
      const link = spreadsheetId ? "https://docs.google.com/spreadsheets/d/" + spreadsheetId : null;
      const productDisplay = link ? "<" + link + "|" + p.name + ">" : p.name;
      noTopText += "• " + productDisplay + " (" + p.batches + " batch" + (p.batches !== 1 ? "es" : "") + " tested)\n";
    });
    blocks1.push({ "type": "section", "text": { "type": "mrkdwn", "text": noTopText } });
  }
  
  // === BUILD MESSAGE 2: Not Tested + Links ===
  const blocks2 = [];
  
  // No Tests section
  if (noTests.length > 0) {
    let noTestsText = "*📭 No Tests This Week:*\n";
    noTests.forEach(product => {
      const spreadsheetId = productSpreadsheetIds[product];
      const link = spreadsheetId ? "https://docs.google.com/spreadsheets/d/" + spreadsheetId : null;
      const productDisplay = link ? "<" + link + "|" + product + ">" : product;
      noTestsText += "• " + productDisplay + "\n";
    });
    blocks2.push({ "type": "section", "text": { "type": "mrkdwn", "text": noTestsText } });
  }
  
  blocks2.push({ "type": "divider" });
  
  // All Product Sheets section
  let allSheetsText = "*📁 All Product Sheets:*\n";
  const productLinks = allProducts.map(product => {
    const spreadsheetId = productSpreadsheetIds[product];
    const link = "https://docs.google.com/spreadsheets/d/" + spreadsheetId;
    return "<" + link + "|" + product + ">";
  });
  allSheetsText += productLinks.join(" | ");
  
  blocks2.push({ "type": "section", "text": { "type": "mrkdwn", "text": allSheetsText } });
  
  return { blocks1: blocks1, blocks2: blocks2 };
}

/**
 * Send only the weekly summary footer to Slack
 * Shows: top performers with comparison, no top performers, no tests, and all product links
 */
function sendCreativeTestsSummaryFooter() {
  const settings = getSettings();
  const channel = settings.slackChannelIdCreative;
  
  if (!settings.slackBotToken || !channel) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Slack Bot Token or Creative Tests Channel ID not configured.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const { blocks1, blocks2 } = buildWeeklySummaryBlocks();
  const sentMessages = [];
  
  // Send message 1 (Performance Overview)
  const result1 = sendSlackBlockMessage(blocks1, channel);
  if (result1) {
    sentMessages.push({ ts: result1.ts, channel: result1.channel });
  }
  
  Utilities.sleep(500);
  
  // Send message 2 (Not Tested + Links)
  const result2 = sendSlackBlockMessage(blocks2, channel);
  if (result2) {
    sentMessages.push({ ts: result2.ts, channel: result2.channel });
  }
  
  if (sentMessages.length > 0) {
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('LAST_CREATIVE_TEST_MESSAGES', JSON.stringify(sentMessages));
    SpreadsheetApp.getUi().alert('✅ Success', 'Weekly Summary sent to Slack!', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to send messages to Slack.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getCreativeTestDataFromTab(selectedProducts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  if (!testSheet) return {};
  
  const data = testSheet.getDataRange().getValues();
  const results = parseCreativeTestsSheet(data, testSheet);
  
  const filteredResults = results.filter(batch => 
    selectedProducts.includes(batch.product) &&
    (batch.category.includes('Top Performer') || batch.category.includes('Mid Performer'))
  );
  
  const byProduct = {};
  filteredResults.forEach(batch => {
    if (!byProduct[batch.product]) {
      byProduct[batch.product] = { topPerformers: [], midPerformers: [] };
    }
    byProduct[batch.product].topPerformers.push(...batch.topPerformerAds);
    byProduct[batch.product].midPerformers.push(...batch.midPerformerAds);
  });
  
  Object.keys(byProduct).forEach(product => {
    byProduct[product].topPerformers.sort((a, b) => parseFloat(b.spend) - parseFloat(a.spend));
    byProduct[product].midPerformers.sort((a, b) => parseFloat(b.spend) - parseFloat(a.spend));
  });
  
  return byProduct;
}

function getCreativeTestDataFromSpreadsheets(selectedProducts) {
  const productSpreadsheetIds = getProductSpreadsheetIds();
  const scriptProps = PropertiesService.getScriptProperties();
  const week = scriptProps.getProperty('CT_SLACK_WEEK') || 'latest';
  
  const byProduct = {};
  
  selectedProducts.forEach(productName => {
    const spreadsheetId = productSpreadsheetIds[productName];
    if (!spreadsheetId) return;
    
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheets = spreadsheet.getSheets();
      
      const weekTabs = sheets.filter(s => /^W\d+\s*\(/.test(s.getName()))
        .sort((a, b) => {
          const weekA = parseInt(a.getName().match(/W(\d+)/)[1]);
          const weekB = parseInt(b.getName().match(/W(\d+)/)[1]);
          return weekB - weekA;
        });
      
      if (weekTabs.length === 0) return;
      
      let targetTab;
      if (week === 'latest') {
        targetTab = weekTabs[0];
      } else {
        targetTab = weekTabs.find(s => s.getName().includes(week)) || weekTabs[0];
      }
      
      const data = targetTab.getDataRange().getValues();
      const productData = parseProductSpreadsheetForSlack(data, targetTab);
      
      if (productData.topPerformers.length > 0 || productData.midPerformers.length > 0) {
        byProduct[productName] = productData;
      }
      
    } catch (e) {
      Logger.log('Error reading ' + productName + ': ' + e.toString());
    }
  });
  
  return byProduct;
}

function parseProductSpreadsheetForSlack(data, sheet) {
  const topPerformers = [];
  const midPerformers = [];
  
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const cellValue = String(data[i][0] || '').toLowerCase();
    if (cellValue.includes('no creative tests were conducted')) {
      return { topPerformers: [], midPerformers: [] };
    }
  }
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const col0 = String(row[0] || '');
    const adName = String(row[1] || '').trim();
    
    if (adName.includes(' | ')) continue;
    
    const isTopPerformer = col0.includes('🟢');
    const isMidPerformer = col0.includes('🟡');
    
    if (!isTopPerformer && !isMidPerformer) continue;
    if (!adName || adName === '') continue;
    
    const spend = parseCurrency(row[2] || '0');
    const roas = parseFloat(row[3]) || 0;
    const cpa = parseCurrency(row[4] || '0');
    const purchases = parseInt(row[5]) || 0;
    
    const isVideo = /VID/i.test(adName);
    let hookRate = null;
    let holdRate = null;
    
    if (isVideo && row[6]) {
      const hookValue = row[6];
      if (typeof hookValue === 'string' && hookValue.includes('%')) {
        hookRate = parseFloat(hookValue.replace('%', '')).toFixed(2);
      } else {
        const hookNum = parseFloat(hookValue);
        if (!isNaN(hookNum)) {
          hookRate = hookNum < 1 ? (hookNum * 100).toFixed(2) : hookNum.toFixed(2);
        }
      }
    }
    
    if (isVideo && row[7]) {
      const holdValue = row[7];
      if (typeof holdValue === 'string' && holdValue.includes('%')) {
        holdRate = parseFloat(holdValue.replace('%', '')).toFixed(2);
      } else {
        const holdNum = parseFloat(holdValue);
        if (!isNaN(holdNum)) {
          holdRate = holdNum < 1 ? (holdNum * 100).toFixed(2) : holdNum.toFixed(2);
        }
      }
    }
    
    let driveUrl = null;
    if (sheet) {
      try {
        const nameCell = sheet.getRange(i + 1, 2);
        const formula = nameCell.getFormula();
        if (formula && formula.includes('HYPERLINK')) {
          const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
          if (urlMatch) driveUrl = urlMatch[1];
        }
      } catch (e) {}
    }
    
    const ad = {
      adName: adName.replace(/^[\s]+/, '').trim(),
      spend: spend.toFixed(2),
      roas: roas.toFixed(2),
      cpa: cpa.toFixed(2),
      purchases: purchases,
      hookRate: hookRate,
      holdRate: holdRate,
      isVideo: isVideo,
      driveUrl: driveUrl
    };
    
    if (isTopPerformer) {
      topPerformers.push(ad);
    } else {
      midPerformers.push(ad);
    }
  }
  
  topPerformers.sort((a, b) => parseFloat(b.spend) - parseFloat(a.spend));
  midPerformers.sort((a, b) => parseFloat(b.spend) - parseFloat(a.spend));
  
  return { topPerformers, midPerformers };
}

function parseCreativeTestsSheet(data, testSheet) {
  const results = [];
  let currentBatch = null;
  let currentProduct = null;  // Track current product header
  
  for (let i = 3; i < data.length; i++) {
    const row = data[i];
    const category = String(row[0]);
    const nameCell = String(row[1]);
    
    // ✅ CORRECT PRODUCT HEADER DETECTION
    // Product headers are in Column A with 📦 emoji and Column B is empty
    if (category.includes('📦') && nameCell === '') {
      currentProduct = category.replace('📦', '').trim();
      Logger.log('✅ Found product header: ' + currentProduct);
      continue;
    }
    
    // Check if this is a batch row - has " | " separator in name cell
    const isBatchRow = nameCell.includes(' | ') && 
                       (category.includes('Top Performer') || category.includes('Mid Performer') || category.includes('Underperformer'));
    
    if (isBatchRow) {
      // Parse: "Task Name | BatchID"
      const parts = nameCell.split(' | ');
      const taskName = parts[0];
      const batchId = parts[1];
      
      const spend = parseCurrency(row[2] || '0');
      const purchases = parseInt(row[5]) || 0;
      const targetCPA = parseFloat(String(row[8]).replace(/[^0-9.]/g, '')) || 60;
      const summary = String(row[9]); // "8 ads: 3T / 0M / 5U"
      
      const summaryMatch = summary.match(/(\d+)T\s*\/\s*(\d+)M\s*\/\s*(\d+)U/);
      const topPerformerCount = summaryMatch ? parseInt(summaryMatch[1]) : 0;
      const midPerformerCount = summaryMatch ? parseInt(summaryMatch[2]) : 0;
      const underperformerCount = summaryMatch ? parseInt(summaryMatch[3]) : 0;
      
      // Use the product header we found
      let product = currentProduct || 'Unknown';
      
      if (!currentProduct) {
        Logger.log('⚠️ No product header found for batch: ' + batchId);
      }
      
      currentBatch = {
        category: category.replace(/[🟢🟡🔴]/g, '').trim(),
        batchId: batchId,
        taskName: taskName,
        product: product,
        totalSpend: spend,
        totalPurchases: purchases,
        targetCPA: targetCPA,
        adCount: topPerformerCount + midPerformerCount + underperformerCount,
        topPerformerCount: topPerformerCount,
        midPerformerCount: midPerformerCount,
        underperformerCount: underperformerCount,
        topPerformerAds: [],
        midPerformerAds: [],
        ads: []
      };
      
      results.push(currentBatch);
    }
    // Check if this is an ad row (has emoji in column A, name in column B)
    else if (currentBatch && category.match(/[🟢🟡🔴]/)) {
      const adName = String(row[1]).replace(/^[\s]+/, '').trim();
      
      if (adName.includes(' | ')) continue;
      
      const spend = parseCurrency(row[2] || '0');
      const roas = parseFloat(row[3]) || 0;
      const cpa = parseCurrency(row[4] || '0');
      const purchases = parseInt(row[5]) || 0;
      
      // Parse hook and hold rates
      let hookRate = null;
      let holdRate = null;

      const isVideo = /VID/i.test(adName);

      if (isVideo && row[6]) {
        const hookValue = row[6];
        if (typeof hookValue === 'string' && hookValue.includes('%')) {
          const hookNum = parseFloat(hookValue.replace('%', ''));
          if (!isNaN(hookNum)) {
            hookRate = hookNum.toFixed(1);
          }
        } else {
          const hookNum = parseFloat(hookValue);
          if (!isNaN(hookNum)) {
            hookRate = hookNum < 1 ? (hookNum * 100).toFixed(1) : hookNum.toFixed(1);
          }
        }
      }

      if (isVideo && row[7]) {
        const holdValue = row[7];
        if (typeof holdValue === 'string' && holdValue.includes('%')) {
          const holdNum = parseFloat(holdValue.replace('%', ''));
          if (!isNaN(holdNum)) {
            holdRate = holdNum.toFixed(1);
          }
        } else {
          const holdNum = parseFloat(holdValue);
          if (!isNaN(holdNum)) {
            holdRate = holdNum < 1 ? (holdNum * 100).toFixed(1) : holdNum.toFixed(1);
          }
        }
      }
      
      // Extract driveUrl from hyperlink formula
      let driveUrl = null;
      if (testSheet) {
        try {
          const nameCell = testSheet.getRange(i + 1, 2); // Column B (Ad Name)
          const formula = nameCell.getFormula();
          
          if (formula && formula.includes('HYPERLINK')) {
            const urlMatch = formula.match(/HYPERLINK\("([^"]+)"/);
            if (urlMatch) {
              driveUrl = urlMatch[1];
            }
          }
        } catch (e) {
          // Ignore errors if sheet access fails
          Logger.log('Could not extract driveUrl: ' + e);
        }
      }
      
      const ad = {
        adName: adName,
        spend: spend.toFixed(2),
        roas: roas.toFixed(2),
        cpa: cpa.toFixed(2),
        purchases: purchases,
        hookRate: hookRate,
        holdRate: holdRate,
        isVideo: isVideo,
        driveUrl: driveUrl  // ← ADD THIS
      };
      
      if (category.includes('🟢')) {
        currentBatch.topPerformerAds.push(ad);
      } else if (category.includes('🟡')) {
        currentBatch.midPerformerAds.push(ad);
      }
    }
  }
  
  return results;
}

function deleteLastCreativeTestsReport() {
  const settings = getSettings();
  
  if (!settings.slackBotToken || !settings.slackChannelIdCreative) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Slack Bot Token or Creative Tests Channel ID not configured.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Show loading dialog first
  const loadingHtml = HtmlService.createHtmlOutput('<div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; padding: 40px; text-align: center;"><p>Loading messages from Slack...</p></div>')
    .setWidth(400)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(loadingHtml, '🗑️ Delete Creative Tests Messages');
  
  // Fetch messages and show selector
  showDeleteMessagesDialog();
}

function showDeleteMessagesDialog() {
  const settings = getSettings();
  const channel = settings.slackChannelIdCreative;
  
  // Fetch bot messages from Slack
  const messages = fetchBotMessagesFromSlack(channel, 50);
  
  if (!messages || messages.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ No Messages', 'No bot messages found in the Creative Tests channel.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const html = buildDeleteMessagesDialog(messages, channel);
  SpreadsheetApp.getUi().showModalDialog(html, '🗑️ Delete Creative Tests Messages');
}

function fetchBotMessagesFromSlack(channel, limit) {
  const settings = getSettings();
  
  const url = 'https://slack.com/api/conversations.history?channel=' + channel + '&limit=' + limit;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + settings.slackBotToken
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      Logger.log('Slack API error: ' + data.error);
      return [];
    }
    
    // Get all messages (including deleted parent messages that might have thread replies)
    const allMessages = data.messages || [];
    
    // Filter to bot messages OR deleted messages with replies
    const relevantMessages = allMessages.filter(msg => {
      const isBot = msg.bot_id || msg.subtype === 'bot_message';
      const isDeletedWithReplies = msg.subtype === 'tombstone' && msg.reply_count > 0;
      const hasBlocks = msg.blocks && msg.blocks.length > 0;
      
      return (isBot && hasBlocks) || isDeletedWithReplies;
    });
    
    // For each message with replies, fetch the thread replies
    relevantMessages.forEach(msg => {
      msg.threadReplies = [];
      
      // Check if message has replies (reply_count > 0)
      if (msg.reply_count && msg.reply_count > 0) {
        const threadReplies = fetchThreadReplies(channel, msg.ts);
        msg.threadReplies = threadReplies;
      }
    });
    
    // Also check for orphaned thread replies (where parent was deleted)
    // These show up as messages with thread_ts different from ts
    const orphanedThreadParents = findOrphanedThreadParents(channel, allMessages);
    
    // Add orphaned thread parents to our list
    orphanedThreadParents.forEach(orphan => {
      // Check if we already have this parent
      const exists = relevantMessages.find(m => m.ts === orphan.ts);
      if (!exists) {
        relevantMessages.push(orphan);
      }
    });
    
    return relevantMessages;
    
  } catch (e) {
    Logger.log('Error fetching Slack messages: ' + e.toString());
    return [];
  }
}

function fetchThreadReplies(channel, threadTs) {
  const settings = getSettings();
  
  const url = 'https://slack.com/api/conversations.replies?channel=' + channel + '&ts=' + threadTs;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + settings.slackBotToken
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      Logger.log('Slack API error fetching replies: ' + data.error);
      return [];
    }
    
    // Filter out the parent message (first message is the parent)
    // and return only bot replies
    const replies = (data.messages || []).filter(msg => {
      const isReply = msg.ts !== threadTs;
      const isBot = msg.bot_id || msg.subtype === 'bot_message';
      return isReply && isBot;
    });
    
    return replies;
    
  } catch (e) {
    Logger.log('Error fetching thread replies: ' + e.toString());
    return [];
  }
}

function findOrphanedThreadParents(channel, allMessages) {
  const settings = getSettings();
  const orphanedParents = [];
  
  // Look for "tombstone" messages (deleted messages) that have reply_count > 0
  allMessages.forEach(msg => {
    if (msg.subtype === 'tombstone' && msg.reply_count && msg.reply_count > 0) {
      // Fetch the thread replies for this deleted parent
      const threadReplies = fetchThreadReplies(channel, msg.ts);
      
      if (threadReplies.length > 0) {
        orphanedParents.push({
          ts: msg.ts,
          isDeleted: true,
          blocks: [{ type: 'section', text: { text: '(Deleted message with thread replies)' } }],
          threadReplies: threadReplies,
          reply_count: threadReplies.length
        });
      }
    }
  });
  
  return orphanedParents;
}

function buildDeleteMessagesDialog(messages, channel) {
  const messageRows = messages.map((msg, idx) => {
    // Extract preview text from blocks
    let preview = '';
    let isHeader = false;
    
    if (msg.blocks) {
      for (let i = 0; i < msg.blocks.length; i++) {
        const block = msg.blocks[i];
        if (block.type === 'header' && block.text) {
          preview = block.text.text || '';
          isHeader = true;
          break;
        } else if (block.type === 'section' && block.text) {
          preview = block.text.text || '';
          break;
        }
      }
    }
    
    // Clean up preview - remove Slack formatting and escape HTML
    preview = preview
      .replace(/<[^|>]+\|([^>]+)>/g, '$1')  // Convert <url|text> to just text
      .replace(/<[^>]+>/g, '')              // Remove any remaining <...> tags
      .replace(/[*`_~]/g, '')               // Remove markdown formatting
      .replace(/&/g, '&amp;')               // Escape HTML special chars
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .substring(0, 50);
    if (preview.length === 50) preview += '...';
    
    // Mark deleted parent messages
    if (msg.isDeleted) {
      preview = '🗑️ ' + preview;
    }
    
    // Format timestamp
    const date = new Date(parseFloat(msg.ts) * 1000);
    const timeStr = date.toLocaleDateString() + ' ' + date.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    
    // Thread indicator
    const threadCount = msg.threadReplies ? msg.threadReplies.length : 0;
    const threadIndicator = threadCount > 0 ? ' <span class="thread-badge">+' + threadCount + ' in thread</span>' : '';
    
    // Store thread reply ts values as data attribute
    const threadTs = msg.threadReplies ? msg.threadReplies.map(r => r.ts).join(',') : '';
    
    return '<div class="message-row">' +
      '<label>' +
      '<input type="checkbox" name="message" value="' + msg.ts + '" data-thread="' + threadTs + '">' +
      '<div class="message-content">' +
      '<span class="message-preview">' + preview + threadIndicator + '</span>' +
      '<span class="message-time">' + timeStr + '</span>' +
      '</div>' +
      '</label></div>';
  }).join('');
  
  const htmlContent = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    '* { box-sizing: border-box; }' +
    'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; margin: 0; color: #333; }' +
    '.header { margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid #e0e0e0; }' +
    '.header h3 { margin: 0 0 8px 0; font-size: 14px; }' +
    '.select-buttons { margin-bottom: 12px; }' +
    '.select-buttons button { padding: 6px 12px; margin-right: 8px; border: 1px solid #ddd; border-radius: 4px; background: #f5f5f5; cursor: pointer; font-size: 12px; }' +
    '.select-buttons button:hover { background: #e8e8e8; }' +
    '.message-list { max-height: 300px; overflow-y: auto; border: 1px solid #e0e0e0; border-radius: 4px; padding: 8px; margin-bottom: 16px; }' +
    '.message-row { padding: 10px 8px; border-bottom: 1px solid #f0f0f0; }' +
    '.message-row:last-child { border-bottom: none; }' +
    '.message-row label { display: flex; align-items: flex-start; cursor: pointer; gap: 10px; }' +
    '.message-content { flex: 1; }' +
    '.message-preview { display: block; font-weight: 500; margin-bottom: 4px; }' +
    '.message-time { font-size: 11px; color: #666; }' +
    '.thread-badge { font-size: 10px; background: #e8f0fe; color: #4285f4; padding: 2px 6px; border-radius: 10px; margin-left: 6px; font-weight: normal; }' +
    'input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; margin-top: 2px; }' +
    '.button-row { display: flex; justify-content: space-between; align-items: center; padding-top: 12px; border-top: 1px solid #e0e0e0; }' +
    '.btn { padding: 10px 20px; border-radius: 4px; font-size: 14px; cursor: pointer; border: none; }' +
    '.btn-cancel { background: #f5f5f5; border: 1px solid #ddd; }' +
    '.btn-cancel:hover { background: #e8e8e8; }' +
    '.btn-delete { background: #d93025; color: white; }' +
    '.btn-delete:hover { background: #b5221b; }' +
    '.btn-delete:disabled { background: #ccc; cursor: not-allowed; }' +
    '.selected-count { font-size: 12px; color: #666; }' +
    '.warning { font-size: 12px; color: #d93025; margin-top: 8px; }' +
    '</style></head><body>' +
    '<div class="header">' +
    '<h3>Select messages to delete:</h3>' +
    '<div class="select-buttons">' +
    '<button onclick="selectAll()">✅ Select All</button>' +
    '<button onclick="deselectAll()">☐ Deselect All</button>' +
    '</div>' +
    '<div class="selected-count" id="selectedCount">0 messages selected</div>' +
    '</div>' +
    '<div class="message-list">' + messageRows + '</div>' +
    '<div class="warning">⚠️ This action cannot be undone</div>' +
    '<div class="button-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-delete" id="deleteBtn" onclick="deleteSelected()" disabled>Delete Selected</button>' +
    '</div>' +
    '<script>' +
    'var channel = "' + channel + '";' +
    'function updateCount() {' +
    '  var checked = document.querySelectorAll("input[name=message]:checked").length;' +
    '  var totalWithThreads = 0;' +
    '  document.querySelectorAll("input[name=message]:checked").forEach(function(cb) {' +
    '    totalWithThreads++;' +
    '    var threadTs = cb.getAttribute("data-thread");' +
    '    if (threadTs) totalWithThreads += threadTs.split(",").filter(function(t) { return t; }).length;' +
    '  });' +
    '  document.getElementById("selectedCount").textContent = checked + " message" + (checked !== 1 ? "s" : "") + " selected" + (totalWithThreads > checked ? " (" + totalWithThreads + " including threads)" : "");' +
    '  document.getElementById("deleteBtn").disabled = checked === 0;' +
    '}' +
    'function selectAll() { document.querySelectorAll("input[name=message]").forEach(function(cb) { cb.checked = true; }); updateCount(); }' +
    'function deselectAll() { document.querySelectorAll("input[name=message]").forEach(function(cb) { cb.checked = false; }); updateCount(); }' +
    'function deleteSelected() {' +
    '  var selected = [];' +
    '  document.querySelectorAll("input[name=message]:checked").forEach(function(cb) {' +
    '    selected.push(cb.value);' +
    '    var threadTs = cb.getAttribute("data-thread");' +
    '    if (threadTs) {' +
    '      threadTs.split(",").filter(function(t) { return t; }).forEach(function(ts) { selected.push(ts); });' +
    '    }' +
    '  });' +
    '  if (selected.length === 0) { alert("Please select at least one message"); return; }' +
    '  document.getElementById("deleteBtn").disabled = true;' +
    '  document.getElementById("deleteBtn").textContent = "Deleting...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) { google.script.host.close(); })' +
    '    .withFailureHandler(function(error) { alert("Error: " + error); document.getElementById("deleteBtn").disabled = false; document.getElementById("deleteBtn").textContent = "Delete Selected"; })' +
    '    .deleteSelectedSlackMessages(selected, channel);' +
    '}' +
    'document.querySelectorAll("input[name=message]").forEach(function(cb) { cb.addEventListener("change", updateCount); });' +
    'updateCount();' +
    '</script></body></html>';
  
  return HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(480);
}

function deleteSelectedSlackMessages(messageTimestamps, channel) {
  const settings = getSettings();
  let deletedCount = 0;
  
  messageTimestamps.forEach(ts => {
    const url = 'https://slack.com/api/chat.delete';
    
    const payload = {
      channel: channel,
      ts: ts
    };
    
    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + settings.slackBotToken,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());
      
      if (data.ok) {
        deletedCount++;
      } else {
        Logger.log('Failed to delete message ' + ts + ': ' + data.error);
      }
    } catch (e) {
      Logger.log('Error deleting message: ' + e.toString());
    }
    
    Utilities.sleep(300);
  });
  
  SpreadsheetApp.getUi().alert('✅ Deleted', 'Successfully deleted ' + deletedCount + '/' + messageTimestamps.length + ' messages.', SpreadsheetApp.getUi().ButtonSet.OK);
  
  return { success: true, deleted: deletedCount };
}

/**
 * Send message using Slack Block Kit
 */
function sendSlackBlockMessage(blocks, channel) {
  const settings = getSettings();
  const botToken = settings.slackBotToken;
  
  if (!botToken) {
    Logger.log('No bot token configured');
    return null;
  }
  
  const url = 'https://slack.com/api/chat.postMessage';
  
  const payload = {
    channel: channel,
    blocks: blocks
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + botToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return {
        success: true,
        ts: data.ts,
        channel: data.channel
      };
    } else {
      Logger.log('Slack API error: ' + data.error);
      return null;
    }
  } catch (e) {
    Logger.log('Error sending message: ' + e.toString());
    return null;
  }
}

function sendSlackThreadReply(blocks, channel, parentTs) {
  const settings = getSettings();
  
  const payload = {
    channel: channel,
    thread_ts: parentTs,
    blocks: blocks
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + settings.slackBotToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.ok) {
      Logger.log('Slack thread reply error: ' + result.error);
      return null;
    }
    
    return result.ts;
  } catch (e) {
    Logger.log('Error sending Slack thread reply: ' + e.toString());
    return null;
  }
}


/**
 * Export creative test results to product-specific spreadsheets
 */
function exportToProductSpreadsheets(testResults, weekLabel, selectedProduct) {
  // Get product spreadsheet mappings from Settings
  const productSpreadsheetIds = getProductSpreadsheetIds();
  
  if (Object.keys(productSpreadsheetIds).length === 0) {
    Logger.log('⚠️ No product spreadsheet mappings found in Settings');
    return { 
      productsExported: 0, 
      skippedProducts: ['No spreadsheet mappings configured'], 
      spreadsheetLinks: [] 
    };
  }
  
  // Convert to abbreviated format
  const abbreviatedWeekLabel = getAbbreviatedWeekName(weekLabel);
  
  Logger.log('=== EXPORTING TO PRODUCT SPREADSHEETS ===');
  Logger.log('Week: ' + abbreviatedWeekLabel);
  Logger.log('Selected product filter: ' + (selectedProduct || 'ALL'));
  Logger.log('Available mappings: ' + Object.keys(productSpreadsheetIds).length);
  
  // Create case-insensitive lookup map
  const productSpreadsheetIdsLower = {};
  Object.keys(productSpreadsheetIds).forEach(key => {
    productSpreadsheetIdsLower[key.toLowerCase()] = {
      originalKey: key,
      spreadsheetId: productSpreadsheetIds[key]
    };
  });
  
  // Group results by product
  const byProduct = {};
  testResults.forEach(batch => {
    if (!byProduct[batch.product]) {
      byProduct[batch.product] = [];
    }
    byProduct[batch.product].push(batch);
  });
  
  Logger.log('Products with data: ' + Object.keys(byProduct).length);
  
  let productsExported = 0;
  const skippedProducts = [];
  const spreadsheetLinks = [];
  const overwriteWarnings = [];
  const emptyTabs = [];
  
  // For each product, update its spreadsheet
  Object.keys(byProduct).sort().forEach(productName => {
    const productBatches = byProduct[productName];
    
    Logger.log('\n📦 Processing: ' + productName);
    
    // Case-insensitive lookup
    const productNameLower = productName.toLowerCase();
    const mapping = productSpreadsheetIdsLower[productNameLower];
    
    if (!mapping) {
      Logger.log('  ⚠️ No spreadsheet ID configured for: ' + productName);
      skippedProducts.push(productName + ' (not in Config)');
      return;
    }
    
    const spreadsheetId = mapping.spreadsheetId;
    Logger.log('  ✓ Found mapping: ' + mapping.originalKey + ' → ' + spreadsheetId);
    
    // Open the spreadsheet
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      Logger.log('  ✓ Opened spreadsheet: ' + spreadsheet.getName());
    } catch (e) {
      Logger.log('  ❌ Error opening spreadsheet: ' + e.toString());
      skippedProducts.push(productName + ' (Invalid ID)');
      return;
    }
    
    // Use abbreviated format for tab name
    const tabName = abbreviatedWeekLabel;
    Logger.log('  Tab name: ' + tabName);
    
    // Check if tab exists
    let sheet = spreadsheet.getSheetByName(tabName);
    
    if (sheet) {
      Logger.log('  ⚠️ Tab already exists - will overwrite');
      overwriteWarnings.push(productName);
      sheet.clear();
    } else {
      Logger.log('  ✨ Creating new tab');
      sheet = spreadsheet.insertSheet(tabName);
    }
    
    // Write data to this product's tab (pass abbreviated week label)
    writeProductDataToSheet(sheet, productBatches, abbreviatedWeekLabel, productName);
    
    // Reorder week tabs from newest to oldest
    reorderWeekTabs(spreadsheet);
    
    productsExported++;
    spreadsheetLinks.push({
      product: productName,
      url: spreadsheet.getUrl() + '#gid=' + sheet.getSheetId()
    });
    
    Logger.log('  ✅ Success! URL: ' + spreadsheet.getUrl());
  });
  
  Logger.log('\n=== EXPORT COMPLETE ===');
  Logger.log('Products exported: ' + productsExported);
  Logger.log('Products skipped: ' + skippedProducts.length);
  Logger.log('Tabs overwritten: ' + overwriteWarnings.length);
  
  return { 
    productsExported: productsExported,
    skippedProducts: skippedProducts,
    spreadsheetLinks: spreadsheetLinks,
    overwriteWarnings: overwriteWarnings,
    emptyTabs: emptyTabs
  };
}

/**
 * Write product data to a sheet (same format as Creative Tests tab)
 */
function writeProductDataToSheet(sheet, batches, weekLabel, productName) {
  // ✅ FIX: Clear sheet completely first (it was already created/cleared in exportToProductSpreadsheets)
  sheet.clear();
  
  let row = 1;
  
  // Title
  sheet.getRange(row, 1, 1, 2).merge();
  sheet.getRange(row, 1).setValue('📈 CREATIVE TEST ANALYSIS')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left');
  row++;
  
  // Week label
  sheet.getRange(row, 1, 1, 2).merge();
  sheet.getRange(row, 1).setValue(weekLabel)
    .setFontSize(11)
    .setHorizontalAlignment('left');
  row++;
  
  // Product header
  sheet.getRange(row, 1, 1, 10).merge();
  sheet.getRange(row, 1).setValue(`📦 ${productName.toUpperCase()}`)
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  row++;
  
  row++; // Blank row
  
  // Headers (same as main Creative Tests sheet)
  const headers = ['', 'Batch / Ad Name', 'Spend', 'ROAS', 'CPA', 'Purchases', 'Hook %', 'Hold %', 'Target CPA', 'Summary'];
  sheet.getRange(row, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  // Set column widths
  sheet.setColumnWidth(1, 30);   // Emoji
  sheet.setColumnWidth(2, 350);  // Ad name
  sheet.setColumnWidth(3, 80);   // Spend
  sheet.setColumnWidth(4, 70);   // ROAS
  sheet.setColumnWidth(5, 70);   // CPA
  sheet.setColumnWidth(6, 80);   // Purchases
  sheet.setColumnWidth(7, 70);   // Hook %
  sheet.setColumnWidth(8, 70);   // Hold %
  sheet.setColumnWidth(9, 90);   // Target CPA

  // Set Target CPA column to text format
  sheet.getRange(1, 9, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');

  sheet.setColumnWidth(10, 200); // Summary
  
  row++;
  
  // Sort batches by category (Winners → Performers → Poor)
  const categoryOrder = {'🟢 Top Performer': 1, '🟡 Mid Performer': 2, '🔴 Underperformer': 3};
  batches.sort((a, b) => {
    const aOrder = categoryOrder[a.category] || 999;
    const bOrder = categoryOrder[b.category] || 999;
    return aOrder - bOrder;
  });
  
  // DEBUG: Check batch properties
  if (batches.length > 0) {
    Logger.log('DEBUG BATCH PROPERTIES: ' + JSON.stringify({
      adCount: batches[0].adCount,
      topPerformerCount: batches[0].topPerformerCount,
      midPerformerCount: batches[0].midPerformerCount,
      underperformerCount: batches[0].underperformerCount,
      adsLength: batches[0].ads ? batches[0].ads.length : 'NO ADS ARRAY'
    }));
  }
  
  // Write batches
  batches.forEach((batch, batchIndex) => {
    // Batch summary row
    const batchSummary = `${batch.adCount} ads: ${batch.topPerformerCount}T / ${batch.midPerformerCount}M / ${batch.underperformerCount}U`;
    
    const batchRowData = [
      batch.category,
      batch.taskName + ' | ' + batch.batchId,
      '€' + batch.totalSpend.toFixed(2),
      '',
      '',
      batch.totalPurchases,
      '',
      '',
      '€' + batch.targetCPA.toFixed(2),
      batchSummary
    ];
    
    sheet.getRange(row, 1, 1, headers.length).setValues([batchRowData]);
    
    // Apply batch formatting
    let batchColor = batch.categoryColor;
    if (batch.category === '🟢 Top Performer') batchColor = '#A8D5A8';
    if (batch.category === '🟡 Mid Performer') batchColor = '#FFE082';
    if (batch.category === '🔴 Underperformer') batchColor = '#EF9A9A';
    
    sheet.getRange(row, 1, 1, headers.length)
      .setBackground(batchColor)
      .setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Add ClickUp link
    const taskCell = sheet.getRange(row, 2);
    const taskText = batch.taskName + ' | ' + batch.batchId;
    taskCell.setFormula('=HYPERLINK("' + batch.taskUrl + '", "' + taskText.replace(/"/g, '""') + '")');
    taskCell.setFontColor('#1a73e8');
    
    row++;
    
    // Individual ads
    batch.ads.forEach((ad, adIndex) => {
      const isVideo = /VID/i.test(ad.adName);
      let emoji = '🔴';
      if (ad.adCategory === '🟢 Top Performer') emoji = '🟢';
      if (ad.adCategory === '🟡 Mid Performer') emoji = '🟡';
      
      // ✅ SEARCH GOOGLE DRIVE FOR THIS AD (same logic as writeCreativeTestsToSheet)
      const parts = ad.adName.split('_');
      let platformIndex = -1;
      for (let j = 0; j < parts.length; j++) {
        if (parts[j] === 'FB' || parts[j] === 'YT' || parts[j] === 'GDN' || parts[j] === 'TT') {
          platformIndex = j;
          break;
        }
      }
      
      const adProductName = platformIndex > 0 ? parts.slice(0, platformIndex).join('_') : parts[0];
      const locale = platformIndex >= 0 && platformIndex + 1 < parts.length ? parts[platformIndex + 1] : '';
      
      Logger.log('     🔍 Searching Drive for: ' + ad.adName + ' (product: ' + adProductName + ', locale: ' + locale + ')');
      
      // Call the same function that Weekly Report uses
      const driveFile = findAdCreativeFile(ad.adName, adProductName, locale);
      
      let driveUrl = null;
      if (driveFile) {
        driveUrl = driveFile.url;
        Logger.log('     📎 Found URL in Drive: ' + driveUrl);
      } else {
        Logger.log('     ⚠️ No file found in Drive');
      }
      
      const adRowData = [
        emoji,
        '  ' + ad.adName,
        '€' + ad.spend.toFixed(2),
        ad.roas.toFixed(2),
        '€' + ad.cpa.toFixed(2),
        ad.purchases,
        isVideo && ad.hookRate ? (ad.hookRate < 1 ? (ad.hookRate * 100).toFixed(2) : parseFloat(ad.hookRate).toFixed(2)) + '%' : '',
        isVideo && ad.holdRate ? (ad.holdRate < 1 ? (ad.holdRate * 100).toFixed(2) : parseFloat(ad.holdRate).toFixed(2)) + '%' : '',
        '',
        ''
      ];
      
      sheet.getRange(row, 1, 1, headers.length).setValues([adRowData]);
      
      // Add hyperlink to ad name if driveUrl exists
      if (driveUrl) {
        const adNameCell = sheet.getRange(row, 2); // Column B
        const displayName = '  ' + ad.adName;
        adNameCell.setFormula('=HYPERLINK("' + driveUrl + '", "' + displayName.replace(/"/g, '""') + '")');
        adNameCell.setFontColor('#1a73e8'); // Blue hyperlink color
      }
      
      sheet.getRange(row, 1, 1, headers.length)
        .setBackground(ad.adCategoryColor)
        .setFontWeight('normal')
        .setVerticalAlignment('middle');
      
      // Center align numerical columns
      sheet.getRange(row, 3, 1, 6).setHorizontalAlignment('center');
      
      row++;
    });
    
    // Add spacing between batches
    row++;
  });
  
  // Freeze header rows (title, week, product header, column headers)
  sheet.setFrozenRows(5);
}


/**
 * Split TheraWolf products and remove minority products
 * This runs BEFORE mixed products detection so TheraWolf never triggers the popup
 * Automatically removes minority products (< 20%) to keep only the majority product
 */
function splitTheraWolfProducts(csvContents) {
  Logger.log('=== SPLITTING THERAWOLF PRODUCTS ===');
  
  const newCsvContents = [];
  
  csvContents.forEach((csvObj, fileIndex) => {
    const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
    const csvFileName = typeof csvObj === 'string' ? `File_${fileIndex + 1}` : (csvObj.fileName || csvObj.name);
    
    // Parse CSV
    let csvData;
    if (csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length) {
      csvData = csv.split('\n').map(line => line.split(';'));
    } else {
      csvData = Utilities.parseCsv(csv);
    }
    
    const headers = csvData[0];
    const adNameIndex = headers.indexOf('Ad name');
    
    if (adNameIndex === -1) {
      Logger.log(`  File ${fileIndex}: No ad name column - keeping as-is`);
      newCsvContents.push(csvObj);
      return;
    }
    
    // ✅ KEY FIX: Use extractProductFromAdName to properly identify products
    const productCounts = {};
    
    for (let i = 1; i < csvData.length; i++) {
      const adName = csvData[i][adNameIndex];
      if (!adName) continue;
      
      // Use the SAME logic as detectMixedProductsInCSVs
      const product = extractProductFromAdName(adName);

if (product && product !== 'Unknown') {
  // ✅ FIX: Normalize compound products before counting
  let normalizedProduct = normalizeProductName(product);
  const productLower = normalizedProduct.toLowerCase();
  productCounts[productLower] = (productCounts[productLower] || 0) + 1;
      }
    }
    
    Logger.log(`  File ${fileIndex} (${csvFileName}):`);
    Logger.log(`    Products detected: ${JSON.stringify(productCounts)}`);
    
    // Check if this CSV has BOTH TheraWolfNeuro AND TheraWolfRelief
    const hasNeuro = Object.keys(productCounts).some(p => p.includes('therawolf') && p.includes('neuro'));
    const hasRelief = Object.keys(productCounts).some(p => p.includes('therawolf') && p.includes('relief'));
    
    // If no TheraWolf products or only one type, keep as-is
    if (!hasNeuro && !hasRelief) {
      Logger.log(`    → No TheraWolf products, keeping as-is`);
      newCsvContents.push(csvObj);
      return;
    }
    
    if (!hasNeuro || !hasRelief) {
      Logger.log(`    → Only one TheraWolf type, keeping as-is`);
      newCsvContents.push(csvObj);
      return;
    }
    
    // Both types present - calculate percentages
    const neuroKey = Object.keys(productCounts).find(p => p.includes('therawolf') && p.includes('neuro'));
    const reliefKey = Object.keys(productCounts).find(p => p.includes('therawolf') && p.includes('relief'));
    
    const neuroCount = productCounts[neuroKey] || 0;
    const reliefCount = productCounts[reliefKey] || 0;
    const totalTheraWolf = neuroCount + reliefCount;
    
    const neuroPercent = totalTheraWolf > 0 ? (neuroCount / totalTheraWolf) : 0;
    const reliefPercent = totalTheraWolf > 0 ? (reliefCount / totalTheraWolf) : 0;
    
    Logger.log(`    TheraWolfNeuro: ${neuroCount} (${(neuroPercent * 100).toFixed(1)}%)`);
    Logger.log(`    TheraWolfRelief: ${reliefCount} (${(reliefPercent * 100).toFixed(1)}%)`);
    
    // CASE 1: One is minority (< 20%) - remove it
    if (neuroPercent < 0.20 || reliefPercent < 0.20) {
      const keepNeuro = neuroPercent >= 0.20;
      
      Logger.log(`    → Removing minority, keeping only ${keepNeuro ? 'Neuro' : 'Relief'}`);
      
      const filteredRows = [headers];
      
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        const adName = row[adNameIndex];
        
        if (!adName) continue;
        
        const product = extractProductFromAdName(adName);
        const productLower = product ? product.toLowerCase() : '';
        
        // Keep if it's the majority type OR if it's not TheraWolf at all
        const isNeuro = productLower.includes('therawolf') && productLower.includes('neuro');
        const isRelief = productLower.includes('therawolf') && productLower.includes('relief');
        const isTheraWolf = isNeuro || isRelief;
        
        if (!isTheraWolf) {
          // Keep non-TheraWolf products
          filteredRows.push(row);
        } else if ((keepNeuro && isNeuro) || (!keepNeuro && isRelief)) {
          // Keep majority TheraWolf type
          filteredRows.push(row);
        }
        // Otherwise skip (minority)
      }
      
      const filteredContent = filteredRows.map(row => row.join(',')).join('\n');
      newCsvContents.push({
        fileName: csvFileName,
        csv: filteredContent,
        content: filteredContent
      });
      
      Logger.log(`    → Filtered CSV: ${filteredRows.length - 1} ads`);
      return;
    }
    
    // CASE 2: Both significant (>= 20%) - split into 2 CSVs
    Logger.log(`    → Both types significant, splitting into 2 CSVs`);
    
    const neuroRows = [headers];
    const reliefRows = [headers];
    
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      const adName = row[adNameIndex];
      
      if (!adName) continue;
      
      const product = extractProductFromAdName(adName);
      const productLower = product ? product.toLowerCase() : '';
      
      if (productLower.includes('therawolf') && productLower.includes('neuro')) {
        neuroRows.push(row);
      } else if (productLower.includes('therawolf') && productLower.includes('relief')) {
        reliefRows.push(row);
      }
      // Drop non-TheraWolf products when splitting
    }
    
    // Create Neuro CSV
    if (neuroRows.length > 1) {
      const neuroContent = neuroRows.map(row => row.join(',')).join('\n');
      newCsvContents.push({
        fileName: csvFileName.replace(/\.(csv|CSV)$/, '') + ' - TheraWolfNeuro.csv',
        csv: neuroContent,
        content: neuroContent
      });
      Logger.log(`    → Created TheraWolfNeuro CSV: ${neuroRows.length - 1} ads`);
    }
    
    // Create Relief CSV
    if (reliefRows.length > 1) {
      const reliefContent = reliefRows.map(row => row.join(',')).join('\n');
      newCsvContents.push({
        fileName: csvFileName.replace(/\.(csv|CSV)$/, '') + ' - TheraWolfRelief.csv',
        csv: reliefContent,
        content: reliefContent
      });
      Logger.log(`    → Created TheraWolfRelief CSV: ${reliefRows.length - 1} ads`);
    }
  });
  
  Logger.log(`✓ TheraWolf splitting complete: ${csvContents.length} input → ${newCsvContents.length} output CSVs`);
  
  return newCsvContents;
}

/**
 * Detect if a single CSV has multiple products (likely a mistake)
 * Returns { hasMixedProducts: boolean, mixedCsvs: [] }
 */
/**
 * Detect if individual CSVs have multiple products (likely a mistake)
 * Now works on RAW CSV contents before upload
 * Returns { hasMixedProducts: boolean, mixedCsvs: [...] }
 */
function detectMixedProductsInCSVs(csvContents) {
  Logger.log('═══════════════════════════════════════════════════════════');
  Logger.log('DEBUG: detectMixedProductsInCSVs() START');
  Logger.log('═══════════════════════════════════════════════════════════');
  
  const mixedCsvs = [];
  
  // Compound products to skip entirely
  const compoundProductsToSkip = ['Clairu_M2', 'CleanlixMold', 'CleanlixPowder', 'TheraWolfNeuro', 'TheraWolfRelief'];
  
  csvContents.forEach((csvObj, fileIndex) => {
    const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
    const csvFileName = typeof csvObj === 'string' ? `File_${fileIndex + 1}` : (csvObj.fileName || csvObj.name);
    
    Logger.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Analyzing CSV: ' + csvFileName);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    
    // Parse CSV
    let csvData;
    if (csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length) {
      csvData = csv.split('\n').map(line => line.split(';'));
    } else {
      csvData = Utilities.parseCsv(csv);
    }
    
    if (csvData.length === 0) return;
    
    const headers = csvData[0];
    const adNameIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
    
    if (adNameIndex === -1) {
      Logger.log('  No "Ad name" column found');
      return;
    }
    
    Logger.log('Total ads: ' + (csvData.length - 1));
    
    // Count products
    const productCounts = {};
    const productOriginalCase = {};
    
    Logger.log('\nExtracting products from ads:');
    for (let i = 1; i < csvData.length; i++) {
      const adName = csvData[i][adNameIndex];
      if (!adName) continue;
      
      const extractedProduct = extractProductFromAdName(adName.toString());
      
      if (i <= 5) {  // Show first 5 ads as examples
        Logger.log('  Ad: ' + adName);
        Logger.log('    → Extracted: "' + extractedProduct + '"');
      }
      
      if (extractedProduct && extractedProduct !== 'Unknown') {
        const normalizedProduct = normalizeProductName(extractedProduct);
        const productLower = normalizedProduct.toLowerCase();
        
        if (i <= 5) {
          Logger.log('    → Normalized: "' + normalizedProduct + '"');
        }
        
        productCounts[productLower] = (productCounts[productLower] || 0) + 1;
        
        if (!productOriginalCase[productLower]) {
          productOriginalCase[productLower] = normalizedProduct;
        }
      }
    }
    
    if (csvData.length > 6) {
      Logger.log('  ... and ' + (csvData.length - 6) + ' more ads');
    }
    
    const detectedProducts = Object.keys(productCounts);
    
    Logger.log('\n📊 Product count summary:');
    detectedProducts.forEach(productLower => {
      const productName = productOriginalCase[productLower];
      const count = productCounts[productLower];
      const percent = Math.round((count / (csvData.length - 1)) * 100);
      Logger.log('  • ' + productName + ': ' + count + ' ads (' + percent + '%)');
    });
    
    // Check if contains compound products (skip these CSVs)
    let hasCompoundProduct = false;
    for (let i = 0; i < detectedProducts.length; i++) {
      const productName = productOriginalCase[detectedProducts[i]];
      if (compoundProductsToSkip.includes(productName)) {
        hasCompoundProduct = true;
        Logger.log('\n⏭️  SKIPPING: CSV contains compound product "' + productName + '"');
        break;
      }
    }
    
    if (hasCompoundProduct) {
      Logger.log('✓ Skipped - compound product');
      return;
    }
    
    // Check for TheraWolf special case
    const hasTheraWolf = detectedProducts.some(p => 
      productOriginalCase[p].toLowerCase().includes('therawolf')
    );
    
    if (hasTheraWolf) {
      Logger.log('\n⏭️  SKIPPING: CSV contains TheraWolf (special multi-product handling)');
      Logger.log('✓ Skipped - TheraWolf');
      return;
    }
    
    // If multiple products detected
    if (detectedProducts.length > 1) {
      Logger.log('\n⚠️  MIXED PRODUCTS DETECTED! (' + detectedProducts.length + ' different products)');
      
      const sortedProducts = detectedProducts.sort((a, b) => 
        productCounts[b] - productCounts[a]
      );
      
      const majorityProductLower = sortedProducts[0];
      const majorityProduct = productOriginalCase[majorityProductLower];
      const majorityCount = productCounts[majorityProductLower];
      const majorityPercent = Math.round((majorityCount / (csvData.length - 1)) * 100);
      
      const minorityProducts = [];
      
      for (let i = 1; i < sortedProducts.length; i++) {
        const minorityProductLower = sortedProducts[i];
        const minorityProduct = productOriginalCase[minorityProductLower];
        const minorityCount = productCounts[minorityProductLower];
        const minorityPercent = Math.round((minorityCount / (csvData.length - 1)) * 100);
        
        minorityProducts.push({
          name: minorityProduct,
          count: minorityCount,
          percent: minorityPercent
        });
      }
      
      Logger.log('  Majority: ' + majorityProduct + ' (' + majorityPercent + '%)');
      minorityProducts.forEach(mp => {
        Logger.log('  Minority: ' + mp.name + ' (' + mp.percent + '%)');
      });
      
      mixedCsvs.push({
        fileIndex: fileIndex,
        fileName: csvFileName,
        totalAds: csvData.length - 1,
        majorityProduct: majorityProduct,
        majorityCount: majorityCount,
        majorityPercent: majorityPercent,
        minorityProducts: minorityProducts
      });
      
      Logger.log('❌ Added to mixed CSVs list');
    } else {
      Logger.log('\n✅ SINGLE PRODUCT - OK');
    }
  });
  
  Logger.log('\n═══════════════════════════════════════════════════════════');
  Logger.log('DEBUG: detectMixedProductsInCSVs() END');
  Logger.log('Total mixed CSVs found: ' + mixedCsvs.length);
  Logger.log('═══════════════════════════════════════════════════════════\n');
  
  return {
    hasMixedProducts: mixedCsvs.length > 0,
    mixedCsvs: mixedCsvs
  };
}

/**
 * TEST FUNCTION 1: Check what ClickUp returns for task comments/history
 * Replace 'YOUR_TASK_ID' with an actual task ID from ClickUp
 */

/**
 * TEST FUNCTION 2: Search for status changes in comments
 */

/**
 * TEST FUNCTION 3: Get task details including custom fields
 */

/**
 * TEST FUNCTION 4: Compare multiple tasks
 */

/**
 * HELPER: Try to find launch date in comments
 */


function processCreativeTestsMultiple(sprintListId, selectedProducts, shouldExport) {
  // Check if resuming from checkpoint
  const checkpoint = getCheckpoint();
  
  if (checkpoint.isActive) {
    // This is a resume - call the resume function instead
    resumeCreativeTests();
    return;
  }
  
  // New run - save checkpoint and start
  clearLogs();
  
  const sprints = getClickUpSprints();
  const selectedSprint = sprints.find(s => s.id === sprintListId);
  if (selectedSprint) {
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('LAST_ANALYZED_WEEK', selectedSprint.name);
  }
  
  const weekLabel = selectedSprint ? selectedSprint.name : 'Unknown Week';
  
  // Save checkpoint
  saveCheckpoint(selectedProducts, 0, sprintListId, shouldExport, weekLabel);
  
  logProgress('🚀 Starting analysis for ' + selectedProducts.length + ' products', 'info');
  logProgress('Products: ' + selectedProducts.join(', '), 'info');
  
  // Process first product
  resumeCreativeTests();
}

function getCheckpoint() {
  const props = PropertiesService.getScriptProperties();
  
  return {
    isActive: props.getProperty('ct_checkpoint_active') === 'true',
    products: JSON.parse(props.getProperty('ct_checkpoint_products') || '[]'),
    currentIndex: parseInt(props.getProperty('ct_checkpoint_index') || '0'),
    weekId: props.getProperty('ct_checkpoint_week'),
    shouldExport: props.getProperty('ct_checkpoint_export') === 'true',
    weekLabel: props.getProperty('ct_checkpoint_label'),
    startTime: props.getProperty('ct_checkpoint_start')
  };
}

function saveCheckpoint(products, currentIndex, weekId, shouldExport, weekLabel) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ct_checkpoint_active', 'true');
  props.setProperty('ct_checkpoint_products', JSON.stringify(products));
  props.setProperty('ct_checkpoint_index', currentIndex.toString());
  props.setProperty('ct_checkpoint_week', weekId);
  props.setProperty('ct_checkpoint_export', shouldExport ? 'true' : 'false');
  props.setProperty('ct_checkpoint_label', weekLabel);
  
  if (!props.getProperty('ct_checkpoint_start')) {
    props.setProperty('ct_checkpoint_start', new Date().toISOString());
  }
}

function clearCheckpoint() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('ct_checkpoint_active');
  props.deleteProperty('ct_checkpoint_products');
  props.deleteProperty('ct_checkpoint_index');
  props.deleteProperty('ct_checkpoint_week');
  props.deleteProperty('ct_checkpoint_export');
  props.deleteProperty('ct_checkpoint_label');
  props.deleteProperty('ct_checkpoint_start');
  
  cleanupTriggers();
}

/**
 * Clean up ALL automation triggers (comprehensive cleanup)
 */
function cleanupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let cleaned = 0;
  
  // List of ALL trigger functions used by this script
  const triggerFunctions = [
    'resumeCreativeTests',
    'runCreativeTestsWithAutoRestart',
    'dummyFunction'
  ];
  
  triggers.forEach(function(trigger) {
    const handlerName = trigger.getHandlerFunction();
    if (triggerFunctions.indexOf(handlerName) > -1) {
      try {
        ScriptApp.deleteTrigger(trigger);
        cleaned++;
      } catch (e) {
        Logger.log('Could not delete trigger: ' + e.toString());
      }
    }
  });
  
  if (cleaned > 0) {
    Logger.log('✅ Cleaned up ' + cleaned + ' triggers');
  }
  
  return cleaned;
}

/**
 * Safe trigger creation with automatic cleanup and retry
 */
function createTriggerSafely(functionName, delayMs) {
  const MAX_RETRIES = 3;
  
  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      // Always clean up old triggers BEFORE creating new one
      cleanupTriggers();
      Utilities.sleep(500);
      
      // Check current trigger count
      const currentTriggers = ScriptApp.getProjectTriggers();
      if (currentTriggers.length >= 18) {
        Logger.log('⚠️ Too many triggers (' + currentTriggers.length + '), force cleaning...');
        forceCleanupAllTriggers();
        Utilities.sleep(1000);
      }
      
      // Now create the new trigger
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .after(delayMs)
        .create();
      
      Logger.log('✅ Trigger created for ' + functionName);
      return true;
      
    } catch (e) {
      Logger.log('⚠️ Trigger creation failed (attempt ' + attempt + '): ' + e.toString());
      
      if (e.message && e.message.indexOf('too many triggers') > -1) {
        Logger.log('🧹 Force cleaning all triggers...');
        forceCleanupAllTriggers();
        Utilities.sleep(2000);
      }
      
      if (attempt === MAX_RETRIES) {
        Logger.log('❌ Failed to create trigger after ' + MAX_RETRIES + ' attempts');
        return false;
      }
    }
  }
  return false;
}

/**
 * Force cleanup ALL project triggers
 */
function forceCleanupAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let cleaned = 0;
  
  triggers.forEach(function(trigger) {
    try {
      ScriptApp.deleteTrigger(trigger);
      cleaned++;
    } catch (e) {
      Logger.log('Could not delete trigger: ' + e.toString());
    }
  });
  
  Logger.log('✅ Force deleted ' + cleaned + ' triggers');
  return cleaned;
}

/**
 * Resume creative tests processing with checkpoint support
 */
function resumeCreativeTests() {
  const checkpoint = getCheckpoint();
  
  // Check if we're done
  if (!checkpoint.isActive || checkpoint.currentIndex >= checkpoint.products.length) {
    const duration = checkpoint.startTime ? 
      Math.round((new Date() - new Date(checkpoint.startTime)) / 1000 / 60) + ' minutes' : 
      'unknown';
    
    logProgress('🎉 ALL PRODUCTS COMPLETED! (' + duration + ')', 'success');
    
    // Write all accumulated results at the end
    const allResults = getAccumulatedResults();
    if (allResults && allResults.length > 0) {
      logProgress('Writing all results to sheet...', 'info');
      writeCreativeTestsToSheet(allResults);
      logProgress('✅ Results written!', 'success');
    }
    
    clearAccumulatedResults();
    clearCheckpoint();
    return;
  }
  
  const currentProduct = checkpoint.products[checkpoint.currentIndex];
  const progress = (checkpoint.currentIndex + 1) + '/' + checkpoint.products.length;
  
  logProgress('📊 Processing: ' + currentProduct + ' (' + progress + ')', 'info');
  Logger.log('='.repeat(60));
  Logger.log('PRODUCT: ' + currentProduct + ' (' + progress + ')');
  Logger.log('='.repeat(60));
  
  // Process current product
  try {
    logProgress('  Analyzing ' + currentProduct + '...', 'info');
    
    const result = analyzeCreativeTests(checkpoint.weekId, currentProduct);
    
    if (result.success) {
      if (result.results && result.results.length > 0) {
        logProgress('  ✓ Analyzed ' + result.results.length + ' batches', 'success');
        
        // Accumulate results instead of writing immediately
        addToAccumulatedResults(result.results);
        
        // Export if needed
        if (checkpoint.shouldExport) {
          const exportResult = exportToProductSpreadsheets(result.results, checkpoint.weekLabel, currentProduct);
          if (exportResult.productsExported > 0) {
            logProgress('  ✓ Exported to product spreadsheet', 'success');
          }
        }
      } else {
        logProgress('  ⚠️ No batches found for ' + currentProduct, 'warning');
        
        if (checkpoint.shouldExport) {
          let baseProduct = currentProduct;
          const labelMatch = currentProduct.match(/^(.+?)\s*\((.+?)\)$/);
          if (labelMatch) {
            baseProduct = labelMatch[1];
          }
          createEmptyCreativeTestTab(baseProduct, checkpoint.weekLabel);
          logProgress('  → Created empty tab', 'info');
        }
      }
    } else {
      logProgress('  ❌ Error: ' + (result.message || 'Unknown error'), 'error');
    }
    
    logProgress('✅ Completed: ' + currentProduct, 'success');
    
  } catch (e) {
    logProgress('❌ Error in ' + currentProduct + ': ' + e.toString(), 'error');
    Logger.log('ERROR: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
  }
  
  // Move to next product
  const nextIndex = checkpoint.currentIndex + 1;
  
  if (nextIndex < checkpoint.products.length) {
    // Save checkpoint for next product
    saveCheckpoint(
      checkpoint.products, 
      nextIndex, 
      checkpoint.weekId, 
      checkpoint.shouldExport, 
      checkpoint.weekLabel
    );
    
    Logger.log('⏩ Scheduling next product...');
    logProgress('⏭️ Next: ' + checkpoint.products[nextIndex], 'info');
    
    // Use safe trigger creation with auto-retry
    const triggerCreated = createTriggerSafely('resumeCreativeTests', 3000);
    
    if (!triggerCreated) {
      // If trigger creation totally failed, try direct recursion as fallback
      Logger.log('⚠️ Trigger creation failed, trying direct call...');
      Utilities.sleep(5000);
      resumeCreativeTests();
    }
    
  } else {
    // All done - write accumulated results
    const duration = checkpoint.startTime ? 
      Math.round((new Date() - new Date(checkpoint.startTime)) / 1000 / 60) + ' minutes' : 
      'unknown';
    
    logProgress('🎉 ALL PRODUCTS COMPLETED! (' + duration + ')', 'success');
    
    const allResults = getAccumulatedResults();
    if (allResults && allResults.length > 0) {
      logProgress('Writing all results to sheet...', 'info');
      writeCreativeTestsToSheet(allResults);
      logProgress('✅ Results written!', 'success');
    }
    
    clearAccumulatedResults();
    clearCheckpoint();
  }
}

// ========================================
// NEW HELPER FUNCTIONS - ADD THESE TOO
// ========================================

function addToAccumulatedResults(results) {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty('accumulated_results');
  
  let allResults = [];
  if (existing) {
    try {
      allResults = JSON.parse(existing);
    } catch (e) {
      Logger.log('Error parsing accumulated results: ' + e.toString());
    }
  }
  
  // Add new results
  allResults = allResults.concat(results);
  
  // Save back
  props.setProperty('accumulated_results', JSON.stringify(allResults));
  Logger.log('✓ Accumulated ' + allResults.length + ' total results');
}

function getAccumulatedResults() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty('accumulated_results');
  
  if (!existing) return [];
  
  try {
    return JSON.parse(existing);
  } catch (e) {
    Logger.log('Error parsing accumulated results: ' + e.toString());
    return [];
  }
}

function clearAccumulatedResults() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('accumulated_results');
  Logger.log('✓ Cleared accumulated results');
}

function authorizeTriggers() {
  // This will force the authorization prompt
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Current triggers: ' + triggers.length);
  
  // Try to create a trigger - this MUST show auth prompt
  ScriptApp.newTrigger('dummyFunction')
    .timeBased()
    .after(1000)
    .create();
  
  Logger.log('✅ Trigger created! Authorization successful!');
  
  // Clean up
  var newTriggers = ScriptApp.getProjectTriggers();
  newTriggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'dummyFunction') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  SpreadsheetApp.getUi().alert(
    '✅ Success!',
    'Authorization complete! You can now use auto-restart.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function dummyFunction() {
  Logger.log('Test function');
}

function resetCreativeTestsProgress() {
  clearCheckpoint();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ Progress Reset', 
    'Checkpoint and triggers have been cleared.\n\n' +
    'You can now start a fresh analysis.',
    ui.ButtonSet.OK);
}

function checkCreativeTestsProgress() {
  const checkpoint = getCheckpoint();
  
  if (!checkpoint.isActive) {
    Logger.log('ℹ️ No analysis in progress');
    SpreadsheetApp.getUi().alert('No analysis currently running');
    return;
  }
  
  const completed = checkpoint.currentIndex;
  const total = checkpoint.products.length;
  const current = checkpoint.products[checkpoint.currentIndex];
  const elapsed = checkpoint.startTime ? 
    Math.round((new Date() - new Date(checkpoint.startTime)) / 1000 / 60) + ' minutes' : 
    'unknown';
  
  const message = 
    'Analysis in Progress\n\n' +
    'Products: ' + checkpoint.products.join(', ') + '\n\n' +
    'Progress: ' + completed + '/' + total + '\n' +
    'Current: ' + current + '\n' +
    'Elapsed: ' + elapsed;
  
  Logger.log(message);
  SpreadsheetApp.getUi().alert('Progress', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * ONE-TIME AUTHORIZATION FOR TRIGGERS
 * Run this once manually to grant permissions
 */

/**
 * Dummy function for authorization test
 */

function testProductSpreadsheetConnections() {
  Logger.log('=== TESTING PRODUCT SPREADSHEET CONNECTIONS ===\n');
  
  const productSpreadsheetIds = getProductSpreadsheetIds();
  
  if (Object.keys(productSpreadsheetIds).length === 0) {
    Logger.log('❌ No product spreadsheet mappings found in Settings');
    SpreadsheetApp.getUi().alert(
      '❌ No Mappings Found',
      'No product spreadsheet IDs found in Settings sheet.\n\n' +
      'Please add them starting at row 11:\n' +
      'Column A: Product Name\n' +
      'Column B: Spreadsheet ID',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log('Found ' + Object.keys(productSpreadsheetIds).length + ' product mappings:\n');
  
  const results = [];
  
  Object.keys(productSpreadsheetIds).forEach(productName => {
    const spreadsheetId = productSpreadsheetIds[productName];
    
    Logger.log('Testing: ' + productName);
    Logger.log('  ID: ' + spreadsheetId);
    
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const name = spreadsheet.getName();
      const url = spreadsheet.getUrl();
      
      Logger.log('  ✅ SUCCESS');
      Logger.log('  Name: ' + name);
      Logger.log('  URL: ' + url);
      Logger.log('');
      
      results.push({
        product: productName,
        status: '✅ Connected',
        name: name,
        url: url
      });
      
    } catch (e) {
      Logger.log('  ❌ FAILED: ' + e.toString());
      Logger.log('');
      
      results.push({
        product: productName,
        status: '❌ Failed',
        error: e.toString()
      });
    }
  });
  
  // Show results in UI
  let message = 'PRODUCT SPREADSHEET CONNECTION TEST\n\n';
  
  const successful = results.filter(r => r.status.includes('✅'));
  const failed = results.filter(r => r.status.includes('❌'));
  
  if (successful.length > 0) {
    message += '✅ SUCCESSFUL (' + successful.length + '):\n';
    successful.forEach(r => {
      message += '  • ' + r.product + '\n';
      message += '    ' + r.name + '\n\n';
    });
  }
  
  if (failed.length > 0) {
    message += '\n❌ FAILED (' + failed.length + '):\n';
    failed.forEach(r => {
      message += '  • ' + r.product + '\n';
      message += '    Error: ' + r.error.substring(0, 80) + '\n\n';
    });
  }
  
  message += '\nCheck Logs (Ctrl+Enter or Cmd+Enter) for full details.';
  
  SpreadsheetApp.getUi().alert('Connection Test Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showUploadCreativeTestsDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  // Get week label
  const scriptProps = PropertiesService.getScriptProperties();
  const weekLabel = scriptProps.getProperty('LAST_ANALYZED_WEEK') || 'Unknown Week';
  
  // Get ALL configured products from Config
  const productSpreadsheetIds = getProductSpreadsheetIds();
  const allConfiguredProducts = Object.keys(productSpreadsheetIds).sort();
  
  if (allConfiguredProducts.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Error', 'No product spreadsheets configured. Please add them to the Config sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get products that have data in Creative Tests sheet
  const productsWithData = new Set();
  
  if (testSheet) {
    const data = testSheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      const cellValue = String(data[i][0]);
      if (cellValue.includes('📦')) {
        const productName = cellValue.replace('📦', '').trim().toUpperCase();
        productsWithData.add(productName);
      }
    }
  }
  
  // Build product options - show ALL configured products, mark which have data
  let productOptions = '';
  allConfiguredProducts.forEach(product => {
    const hasData = productsWithData.has(product.toUpperCase());
    const indicator = hasData ? '✅' : '📭';
    const label = hasData ? product : product + ' (no tests)';
    productOptions += `<option value="${product}" data-hasdata="${hasData}">${indicator} ${label}</option>`;
  });
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
            margin: 0;
          }
          .container {
            max-width: 600px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          }
          h2 {
            color: #333;
            margin-bottom: 20px;
          }
          .info-box {
            background: #e8f4fd;
            border-left: 4px solid #2196F3;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
          }
          label {
            display: block;
            margin-top: 15px;
            margin-bottom: 5px;
            font-weight: bold;
            color: #555;
          }
          select {
            width: 100%;
            padding: 10px;
            margin-bottom: 10px;
            border: 2px solid #4285f4;
            border-radius: 4px;
            font-size: 14px;
          }
          select[multiple] {
            height: 300px;
          }
          .select-hint {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
          }
          .quick-select {
            display: flex;
            gap: 10px;
            margin-top: 10px;
          }
          .quick-btn {
            padding: 6px 12px;
            background: #f0f0f0;
            border: 1px solid #ccc;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
          }
          .quick-btn:hover {
            background: #e0e0e0;
          }
          button {
            background: #4285f4;
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            width: 100%;
            margin-top: 20px;
          }
          button:disabled {
            background: #ccc;
            cursor: not-allowed;
          }
          #status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 4px;
            display: none;
            max-height: 300px;
            overflow-y: auto;
          }
          #status.show { display: block; }
          #status.success {
            background: #d4edda;
            color: #155724;
          }
          #status.error {
            background: #f8d7da;
            color: #721c24;
          }
          #status.info {
            background: #d1ecf1;
            color: #0c5460;
          }
          .modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0,0,0,0.5);
            z-index: 1000;
            align-items: center;
            justify-content: center;
          }
          .modal-overlay.show {
            display: flex;
          }
          .modal-content {
            background: white;
            padding: 30px;
            border-radius: 10px;
            max-width: 500px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
          }
          .modal-content h3 {
            margin-top: 0;
            color: #e67e22;
          }
          .product-list {
            background: #fff3cd;
            padding: 15px;
            border-radius: 4px;
            margin: 15px 0;
            max-height: 200px;
            overflow-y: auto;
          }
          .modal-buttons {
            display: flex;
            gap: 10px;
            margin-top: 20px;
          }
          .modal-buttons button {
            flex: 1;
            margin-top: 0;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>📤 Upload Creative Tests to Product Sheets</h2>
          
          <div class="info-box">
            <strong>Week:</strong> ${weekLabel}<br>
            <strong>Action:</strong> Create/update tabs in product spreadsheets
          </div>
          
          <label>Select products to upload:</label>
          <div class="legend">
            <span>✅ Has creative test data</span>
            <span>📭 No tests (will create empty tab)</span>
          </div>
          <select id="productSelect" multiple>
            ${productOptions}
          </select>
          <div class="select-hint">Hold Ctrl (Windows) or Cmd (Mac) to select multiple</div>
          
          <div class="quick-select">
            <button type="button" class="quick-btn" id="selectAllBtn">Select All</button>
            <button type="button" class="quick-btn" id="selectDataBtn">Select With Data Only</button>
            <button type="button" class="quick-btn" id="clearAllBtn">Clear All</button>
          </div>
          
          <button type="button" id="uploadBtn">Upload to Product Sheets</button>
          <div id="status"></div>
        </div>
        
        <!-- Confirmation Modal -->
        <div id="confirmModal" class="modal-overlay">
          <div class="modal-content">
            <h3>⚠️ Existing Tabs Detected</h3>
            <p>The following products already have tabs for <strong>${weekLabel}</strong>:</p>
            <div class="product-list" id="existingProductsList"></div>
            <p>What would you like to do?</p>
            <div class="modal-buttons">
              <button type="button" id="btnCancel" style="background: #95a5a6;">Cancel</button>
              <button type="button" id="btnSelective" style="background: #3498db;">Choose Which</button>
              <button type="button" id="btnOverrideAll" style="background: #e67e22;">Override All</button>
            </div>
          </div>
        </div>
        
        <script>
          const weekLabel = '${weekLabel}';
          const productSelect = document.getElementById('productSelect');
          const selectAllBtn = document.getElementById('selectAllBtn');
          const selectDataBtn = document.getElementById('selectDataBtn');
          const clearAllBtn = document.getElementById('clearAllBtn');
          const confirmModal = document.getElementById('confirmModal');
          
          // Quick select buttons
          selectAllBtn.addEventListener('click', function() {
            for (let i = 0; i < productSelect.options.length; i++) {
              productSelect.options[i].selected = true;
            }
          });
          
          selectDataBtn.addEventListener('click', function() {
            for (let i = 0; i < productSelect.options.length; i++) {
              // Only select products that have data
              productSelect.options[i].selected = productSelect.options[i].dataset.hasdata === 'true';
            }
          });
          
          clearAllBtn.addEventListener('click', function() {
            for (let i = 0; i < productSelect.options.length; i++) {
              productSelect.options[i].selected = false;
            }
          });
          
          // Main upload button
          document.getElementById('uploadBtn').addEventListener('click', function() {
            const selectedOptions = Array.from(productSelect.selectedOptions);
            const products = selectedOptions.map(opt => opt.value);
            
            if (products.length === 0) {
              alert('Please select at least one product');
              return;
            }
            
            showStatus('Checking for existing tabs...', 'info');
            
            // First check which tabs already exist
            google.script.run
              .withSuccessHandler(function(existingTabs) {
                if (existingTabs && existingTabs.length > 0) {
                  // Show confirmation modal
                  showConfirmationModal(existingTabs, products);
                } else {
                  // No existing tabs, proceed directly
                  performUpload(products);
                }
              })
              .withFailureHandler(function(error) {
                showStatus('Error checking existing tabs: ' + error.message, 'error');
              })
              .checkExistingTabsForProducts(products, weekLabel);
          });
          
          function showConfirmationModal(existingTabs, allProducts) {
            const existingList = document.getElementById('existingProductsList');
            existingList.innerHTML = existingTabs.map(t => '• ' + t).join('<br>');
            confirmModal.classList.add('show');
            
            document.getElementById('btnCancel').onclick = function() {
              confirmModal.classList.remove('show');
              showStatus('Upload cancelled', 'info');
            };
            
            document.getElementById('btnOverrideAll').onclick = function() {
              confirmModal.classList.remove('show');
              performUpload(allProducts);
            };
            
            document.getElementById('btnSelective').onclick = function() {
              confirmModal.classList.remove('show');
              showSelectiveDialog(existingTabs, allProducts);
            };
          }
          
          function showSelectiveDialog(existingProductNames, allProducts) {
            const options = allProducts.map((p, i) => {
              const status = existingProductNames.includes(p) ? ' (⚠️ will overwrite)' : ' (✓ new tab)';
              return (i + 1) + '. ' + p + status;
            }).join('\\n');
            
            const choice = prompt(
              'Select which products to process:\\n\\n' + options + 
              '\\n\\nEnter product numbers separated by commas\\n(or type "all" to process all, "new" for only new tabs):'
            );
            
            if (!choice) {
              showStatus('Upload cancelled', 'info');
              return;
            }
            
            const choiceLower = choice.trim().toLowerCase();
            let finalProducts = [];
            
            if (choiceLower === 'all') {
              finalProducts = allProducts;
            } else if (choiceLower === 'new') {
              finalProducts = allProducts.filter(p => !existingProductNames.includes(p));
            } else {
              // Parse numbers like "1,2,3"
              const indices = choice.split(',').map(s => parseInt(s.trim()) - 1);
              finalProducts = indices
                .filter(i => i >= 0 && i < allProducts.length)
                .map(i => allProducts[i]);
            }
            
            if (finalProducts.length === 0) {
              showStatus('No valid products selected', 'error');
              return;
            }
            
            performUpload(finalProducts);
          }
          
          function performUpload(products) {
            const btn = document.getElementById('uploadBtn');
            
            btn.disabled = true;
            btn.textContent = 'Uploading...';
            showStatus('Uploading to ' + products.length + ' product spreadsheet(s)...', 'info');
            
            google.script.run
              .withSuccessHandler(function(result) {
                btn.disabled = false;
                btn.textContent = 'Upload to Product Sheets';
                
                if (result.success) {
                  let message = '<strong>✅ Success!</strong><br><br>';
                  
                  if (result.spreadsheetLinks && result.spreadsheetLinks.length > 0) {
                    message += '📊 <strong>Updated Spreadsheets:</strong><br>';
                    result.spreadsheetLinks.forEach(function(link) {
                      message += '• <a href="' + link.url + '" target="_blank">' + link.product + '</a><br>';
                    });
                  }
                  
                  if (result.emptyTabs && result.emptyTabs.length > 0) {
                    message += '<br>📭 <strong>Empty tabs created (no tests):</strong> ' + result.emptyTabs.join(', ');
                  }
                  
                  if (result.skippedProducts && result.skippedProducts.length > 0) {
                    message += '<br>⚠️ <strong>Skipped:</strong> ' + result.skippedProducts.join(', ');
                  }
                  
                  showStatus(message, 'success');
                } else {
                  showStatus('<strong>❌ Error:</strong><br>' + result.message, 'error');
                }
              })
              .withFailureHandler(function(error) {
                btn.disabled = false;
                btn.textContent = 'Upload to Product Sheets';
                showStatus('<strong>❌ Error:</strong><br>' + error.message, 'error');
              })
              .uploadExistingCreativeTestsDirect(products);
          }
          
          function showStatus(message, type) {
            const status = document.getElementById('status');
            status.className = 'show ' + type;
            status.innerHTML = message;
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(700)
  .setHeight(750);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Creative Tests to Product Sheets');
}

function uploadExistingCreativeTests(selectedProducts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  if (!testSheet) {
    return { success: false, message: 'Creative Tests sheet not found' };
  }
  
  // Get week label
  const scriptProps = PropertiesService.getScriptProperties();
  const weekLabel = scriptProps.getProperty('LAST_ANALYZED_WEEK') || 'Unknown Week';
  
  // Parse the Creative Tests sheet
  const data = testSheet.getDataRange().getValues();
  const testResults = parseCreativeTestsSheetForUpload(data, selectedProducts);
  
  if (testResults.length === 0) {
    return { success: false, message: 'No batches found for selected products' };
  }
  
  // Upload to product spreadsheets
  const exportResult = exportToProductSpreadsheets(testResults, weekLabel, null);
  
  return {
    success: true,
    productsExported: exportResult.productsExported,
    skippedProducts: exportResult.skippedProducts,
    spreadsheetLinks: exportResult.spreadsheetLinks
  };
}

function parseCreativeTestsSheetForUpload(data, selectedProducts) {
  const results = [];
  let currentBatch = null;
  let currentProduct = null;
  
  for (let i = 3; i < data.length; i++) {
    const row = data[i];
    const category = String(row[0]);
    const nameCell = String(row[1]);
    
    // Product header detection
    if (category.includes('📦') && nameCell === '') {
      currentProduct = category.replace('📦', '').trim();
      Logger.log('Found product header: ' + currentProduct);
      
      // Skip if not in selected products (case-insensitive)
      const selectedUpper = selectedProducts.map(p => p.toUpperCase());
      if (!selectedUpper.includes(currentProduct.toUpperCase())) {
        currentProduct = null;
      }
      continue;
    }
    
    // Skip if we're not in a selected product
    if (!currentProduct) continue;
    
    // Batch row detection
    const isBatchRow = nameCell.includes(' | ') && 
                       (category.includes('Top Performer') || category.includes('Mid Performer') || category.includes('Underperformer'));
    
    if (isBatchRow) {
      const parts = nameCell.split(' | ');
      const taskName = parts[0];
      const batchId = parts[1];
      
      const spend = parseCurrency(row[2] || '0');
      const purchases = parseInt(row[5]) || 0;
      const targetCPA = parseFloat(String(row[8]).replace(/[^0-9.]/g, '')) || 60;
      const summary = String(row[9]);
      
      const summaryMatch = summary.match(/(\d+)T\s*\/\s*(\d+)M\s*\/\s*(\d+)U/);
      const topPerformerCount = summaryMatch ? parseInt(summaryMatch[1]) : 0;
      const midPerformerCount = summaryMatch ? parseInt(summaryMatch[2]) : 0;
      const underperformerCount = summaryMatch ? parseInt(summaryMatch[3]) : 0;
      
      currentBatch = {
        category: category.replace(/[🟢🟡🔴]/g, '').trim(),
        categoryColor: category.includes('🟢') ? '#D4EDDA' : (category.includes('🟡') ? '#FFF3CD' : '#F8D7DA'),
        batchId: batchId,
        taskName: taskName,
        taskUrl: '', // We don't have this stored, but it's okay
        product: currentProduct,
        totalSpend: spend,
        targetCPA: targetCPA,
        totalPurchases: purchases,
        adCount: topPerformerCount + midPerformerCount + underperformerCount,
        topPerformerCount: topPerformerCount,
        midPerformerCount: midPerformerCount,
        underperformerCount: underperformerCount,
        ads: []
      };
      
      results.push(currentBatch);
    }
    // Ad row detection
    else if (currentBatch && category.match(/[🟢🟡🔴]/)) {
      const adName = String(row[1]).replace(/^[\s]+/, '').trim();
      
      if (adName.includes(' | ')) continue;
      
      const spend = parseCurrency(row[2] || '0');
      const roas = parseFloat(row[3]) || 0;
      const cpa = parseCurrency(row[4] || '0');
      const purchases = parseInt(row[5]) || 0;
      
      const isVideo = /VID/i.test(adName);
      let hookRate = null;
      let holdRate = null;
      
      if (isVideo && row[6]) {
        const hookValue = row[6];
        if (typeof hookValue === 'string' && hookValue.includes('%')) {
          hookRate = parseFloat(hookValue.replace('%', ''));
        } else {
          const hookNum = parseFloat(hookValue);
          if (!isNaN(hookNum)) {
            hookRate = hookNum < 1 ? hookNum * 100 : hookNum;
          }
        }
      }
      
      if (isVideo && row[7]) {
        const holdValue = row[7];
        if (typeof holdValue === 'string' && holdValue.includes('%')) {
          holdRate = parseFloat(holdValue.replace('%', ''));
        } else {
          const holdNum = parseFloat(holdValue);
          if (!isNaN(holdNum)) {
            holdRate = holdNum < 1 ? holdNum * 100 : holdNum;
          }
        }
      }
      
      const adCategory = category.includes('🟢') ? '🟢 Top Performer' : (category.includes('🟡') ? '🟡 Mid Performer' : '🔴 Underperformer');
      const adCategoryColor = category.includes('🟢') ? '#D4EDDA' : (category.includes('🟡') ? '#FFF3CD' : '#F8D7DA');
      
      currentBatch.ads.push({
        adName: adName,
        spend: spend,
        roas: roas,
        cpa: cpa,
        purchases: purchases,
        hookRate: hookRate,
        holdRate: holdRate,
        adCategory: adCategory,
        adCategoryColor: adCategoryColor
      });
    }
  }
  
  return results;
}

/**
 * DEBUG FUNCTION - Trace the entire creative tests flow
 */


/**
 * Debug function to test cached creative tests analysis
 * Tests multiple products to verify caching performance
 */

/**
 * Quick performance test - compares cached vs uncached for 2-3 products
 * Completes in ~10-15 minutes instead of 30+
 */

/**
 * Simple test - just run the NEW cached approach
 * No comparison, just pure performance test
 */

/**
 * Check which tabs already exist for selected products
 */
function checkExistingTabs(products, weekName) {
  Logger.log('=== checkExistingTabs START ===');
  
  // ✅ Use central Config via getProductSpreadsheetIds() instead of local Settings
  const productSpreadsheetIds = getProductSpreadsheetIds();
  
  if (Object.keys(productSpreadsheetIds).length === 0) {
    Logger.log('ERROR: No product spreadsheet mappings found');
    return [];
  }
  
  // Create case-insensitive lookup map
  const productSpreadsheetIdsLower = {};
  Object.keys(productSpreadsheetIds).forEach(key => {
    productSpreadsheetIdsLower[key.toLowerCase()] = {
      originalKey: key,
      spreadsheetId: productSpreadsheetIds[key]
    };
  });
  
  const existingTabs = [];
  
  // Convert to abbreviated format
  const weekNameAbbreviated = getAbbreviatedWeekName(weekName);
  
  Logger.log('Looking for tab name: "' + weekNameAbbreviated + '"');
  Logger.log('Available mappings: ' + Object.keys(productSpreadsheetIds).join(', '));
  
  for (let i = 0; i < products.length; i++) {
    const product = products[i];
    const productLower = product.toLowerCase();
    
    Logger.log('---');
    Logger.log('Checking product #' + (i+1) + ': ' + product);
    
    // Case-insensitive lookup
    const mapping = productSpreadsheetIdsLower[productLower];
    
    if (!mapping) {
      Logger.log('NOT FOUND in Config: ' + product);
      continue;
    }
    
    const spreadsheetId = mapping.spreadsheetId;
    Logger.log('Found mapping: ' + mapping.originalKey + ' → ' + spreadsheetId.substring(0, 20) + '...');
    
    try {
      const targetSs = SpreadsheetApp.openById(spreadsheetId);
      const existingSheet = targetSs.getSheetByName(weekNameAbbreviated);
      
      if (existingSheet) {
        Logger.log('✓ TAB EXISTS for ' + product);
        existingTabs.push(product);
      } else {
        Logger.log('✗ Tab does not exist for ' + product);
      }
    } catch (error) {
      Logger.log('ERROR accessing spreadsheet: ' + error.message);
    }
  }
  
  Logger.log('=== checkExistingTabs END ===');
  Logger.log('Total existing tabs found: ' + existingTabs.length);
  Logger.log('Products with existing tabs: ' + JSON.stringify(existingTabs));
  
  return existingTabs;
}

// NEW HELPER FUNCTION
function normalizeProductNameForLookup(product) {
  // Convert product name to match Settings sheet format
  
  // First, handle special cases with labels in parentheses
  if (product.indexOf('(MAIN)') !== -1) {
    // "CAMTRIX (MAIN)" -> "CamTrix (Main)"
    var baseName = product.replace('(MAIN)', '').trim();
    return toProperCase(baseName) + ' (Main)';
  }
  
  if (product.indexOf('(TOP5)') !== -1) {
    // "CAMTRIX (TOP5)" -> "CamTrix (TOP5)"
    var baseName = product.replace('(TOP5)', '').trim();
    return toProperCase(baseName) + ' (TOP5)';
  }
  
  // Handle compound names without spaces
  var normalized = product.replace(/\s+/g, ''); // Remove spaces first
  
  // Special case mappings - EXACT matches from Settings sheet
  var mappings = {
    'ALIVEBLUE': 'AliveBlue',
    'CAMTRIX': 'CamTrix',
    'THERAWOLFNEURO': 'TheraWolfNeuro',
    'THERAWOLFRELIEF': 'TheraWolfRelief',
    'CLEANLIXMOLD': 'CleanlixMold',
    'GUARDALITY': 'Guardality',
    'WELLHEATER': 'WellHeater',
    'KATUCHEF': 'KatuChef',
    'CLAIRU_M2': 'Clairu_M2',
    'ORIBREEZE': 'OriBreeze',
    'WAVEMAX': 'Wavemax',
    'ROBOTPUPPY': 'RobotPuppy',
    'AURANATURALS': 'AuraNaturals'
  };
  
  return mappings[normalized] || normalized;
}

// Helper function for proper case conversion
function toProperCase(str) {
  // For known products, use exact mapping
  var upperStr = str.toUpperCase();
  var mappings = {
    'CAMTRIX': 'CamTrix',
    'ALIVEBLUE': 'AliveBlue',
    'THERAWOLF': 'TheraWolf',
    'CLEANLIX': 'Cleanlix',
    'GUARDALITY': 'Guardality',
    'WELLHEATER': 'WellHeater',
    'KATUCHEF': 'KatuChef',
    'CLAIRU': 'Clairu',
    'ORIBREEZE': 'OriBreeze',
    'WAVEMAX': 'Wavemax',
    'ROBOTPUPPY': 'RobotPuppy',
    'AURANATURALS': 'AuraNaturals'
  };
  
  return mappings[upperStr] || (str.charAt(0).toUpperCase() + str.slice(1).toLowerCase());
}

/**
 * Show confirmation dialog for existing tabs
 */
function showOverrideConfirmation(existingTabs, weekName) {
  const ui = SpreadsheetApp.getUi();
  
  const productList = existingTabs
    .map(function(tab) { return '• ' + tab.product.toUpperCase(); })
    .join('\n');
  
  const message = 
    'The following product sheets already have a tab named "' + weekName + '":\n\n' +
    productList + '\n\n' +
    'What would you like to do?';
  
  const response = ui.alert(
    'Existing Tabs Detected',
    message,
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    return 'OVERRIDE_ALL';
  } else if (response === ui.Button.NO) {
    return 'SELECTIVE';
  } else {
    return 'CANCEL_ALL';
  }
}

/**
 * Let user select which products to override
 */
function selectiveOverride(existingTabs, allSelectedProducts, weekName) {
  const ui = SpreadsheetApp.getUi();
  
  // Create a list of products with existing tabs
  const existingProductNames = existingTabs.map(function(tab) { 
    return tab.product; 
  });
  
  // Build checkbox list
  const productOptions = allSelectedProducts.map(function(product) {
    const hasExisting = existingProductNames.indexOf(product) !== -1;
    const status = hasExisting ? ' (⚠️ will overwrite)' : ' (✓ new tab)';
    return product + status;
  }).join('\n');
  
  const message = 
    'Select which products to process:\n\n' +
    productOptions + '\n\n' +
    'Enter product names separated by commas\n' +
    '(or type "all" to process all, "new" for only new tabs):';
  
  const response = ui.prompt(
    'Select Products',
    message,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return [];
  }
  
  const userInput = response.getResponseText().trim().toLowerCase();
  
  if (userInput === 'all') {
    return allSelectedProducts;
  }
  
  if (userInput === 'new') {
    // Only process products without existing tabs
    return allSelectedProducts.filter(function(product) {
      return existingProductNames.indexOf(product) === -1;
    });
  }
  
  // Parse comma-separated list
  const selectedNames = userInput
    .split(',')
    .map(function(name) { return name.trim(); });
  
  // Match against available products (case-insensitive)
  return allSelectedProducts.filter(function(product) {
    return selectedNames.some(function(name) {
      return product.toLowerCase() === name;
    });
  });
}

/**
 * Check existing tabs and return list (called from HTML dialog)
 */
function checkExistingTabsForProducts(products, weekName) {
  Logger.log('=== checkExistingTabsForProducts DEBUG ===');
  Logger.log('Input products: ' + JSON.stringify(products));
  Logger.log('Week name: ' + weekName);
  
  const result = checkExistingTabs(products, weekName);
  
  Logger.log('Result from checkExistingTabs: ' + JSON.stringify(result));
  Logger.log('Number of existing tabs found: ' + result.length);
  
  return result;
}

/**
 * Upload directly without checking (confirmation already done in dialog)
 */
function uploadExistingCreativeTestsDirect(selectedProducts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  // Get week label
  const scriptProps = PropertiesService.getScriptProperties();
  const weekLabel = scriptProps.getProperty('LAST_ANALYZED_WEEK') || 'Unknown Week';
  
  // Convert to abbreviated format
  const abbreviatedWeekLabel = getAbbreviatedWeekName(weekLabel);
  
  // Get products with data from Creative Tests sheet
  const productsWithData = new Set();
  let testResults = [];
  
  if (testSheet) {
    const data = testSheet.getDataRange().getValues();
    
    // Find products that have data
    for (let i = 0; i < data.length; i++) {
      const cellValue = String(data[i][0]);
      if (cellValue.includes('📦')) {
        const productName = cellValue.replace('📦', '').trim().toUpperCase();
        productsWithData.add(productName);
      }
    }
    
    // Parse test results for products WITH data
    const productsToExport = selectedProducts.filter(p => productsWithData.has(p.toUpperCase()));
    if (productsToExport.length > 0) {
      testResults = parseCreativeTestsSheetForUpload(data, productsToExport);
    }
  }
  
  // Separate selected products into those with data and those without
  const productsForDataTabs = selectedProducts.filter(p => productsWithData.has(p.toUpperCase()));
  const productsForEmptyTabs = selectedProducts.filter(p => !productsWithData.has(p.toUpperCase()));
  
  Logger.log('=== UPLOAD CREATIVE TESTS ===');
  Logger.log('Products with data: ' + productsForDataTabs.join(', '));
  Logger.log('Products for empty tabs: ' + productsForEmptyTabs.join(', '));
  
  let exportResult = { productsExported: 0, skippedProducts: [], spreadsheetLinks: [], emptyTabs: [] };
  
  // Export products with data
  if (testResults.length > 0) {
    exportResult = exportToProductSpreadsheets(testResults, abbreviatedWeekLabel, null);
  }
  
  // Create empty tabs for products without data
  const emptyTabsCreated = [];
  productsForEmptyTabs.forEach(product => {
    try {
      createEmptyCreativeTestTab(product, weekLabel);
      emptyTabsCreated.push(product);
      Logger.log('✅ Created empty tab for: ' + product);
    } catch (e) {
      Logger.log('❌ Failed to create empty tab for ' + product + ': ' + e.message);
      exportResult.skippedProducts.push(product + ' (error creating empty tab)');
    }
  });
  
  return {
    success: true,
    message: 'Upload complete',
    productsExported: exportResult.productsExported,
    skippedProducts: exportResult.skippedProducts,
    spreadsheetLinks: exportResult.spreadsheetLinks,
    emptyTabs: emptyTabsCreated
  };
}


/**
 * Converts week name to abbreviated format
 * "Week 48 (11/24 - 11/30)" → "W48 (11/24 - 11/30)"
 * "W48 (11/24 - 11/30)" → "W48 (11/24 - 11/30)" (no change)
 */
function getAbbreviatedWeekName(weekName) {
  // Convert "Week 51 (12/15 - 12/21) (2025)" to "W51 (12/15 - 12/21)"
  let abbreviated = weekName.replace(/^Week (\d+)/, 'W$1');
  // Remove year suffix like " (2025)" or " (2026)"
  abbreviated = abbreviated.replace(/\s*\(\d{4}\)\s*$/, '');
  return abbreviated;
}

/**
 * Reorder week tabs so newest weeks come first (W52, W51, W50...)
 * Only affects tabs that match the W## pattern
 */
function reorderWeekTabs(spreadsheet) {
  try {
    const sheets = spreadsheet.getSheets();
    
    // Find all week tabs (W1, W2, ... W52 format)
    const weekSheets = [];
    const otherSheets = [];
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      const weekMatch = name.match(/^W(\d+)/);
      if (weekMatch) {
        weekSheets.push({
          sheet: sheet,
          weekNum: parseInt(weekMatch[1]),
          name: name
        });
      } else {
        otherSheets.push(sheet);
      }
    });
    
    // Sort week tabs by week number descending (newest first)
    weekSheets.sort((a, b) => b.weekNum - a.weekNum);
    
    // Move week tabs to the beginning, in order
    weekSheets.forEach((item, index) => {
      item.sheet.activate();
      spreadsheet.moveActiveSheet(index + 1);
    });
    
    Logger.log('✅ Reordered ' + weekSheets.length + ' week tabs (newest first)');
  } catch (e) {
    Logger.log('⚠️ Could not reorder tabs: ' + e.message);
  }
}

/**
 * DEBUG FUNCTION 1: Test product extraction on sample ad names
 */

/**
 * DEBUG FUNCTION 2: Simulate mixed products detection on real CSV data
 */

/**
 * DEBUG FUNCTION 3: Test specific ad name to see full extraction flow
 */

/**
 * DEBUG FUNCTION 4: Show what normalizeProductName does
 */

/**
 * DEBUG: Test the WiggyDog YT to FB 2 task specifically
 */

/**
 * DEBUG: Find WiggyDog task ID
 */

/**
 * DEBUG: Test getMasterLaunchDate function
 */

/**
 * DEBUG: Simulate the two-pass logic
 */

/**
 * DEBUG: Test WiggyDog batch detection with full two-pass logic
 */

/**
 * DEBUG: Test if WiggyDog master task passes the optimization filters
 */

/**
 * DEBUG: Find ALL WiggyDog tasks and their batches
 */

/**
 * DEBUG: Full simulation of WiggyDog YT to FB 2 processing with updated pattern
 */

/**
 * DEBUG: Check if WiggyDog batches have CSV data
 */

/**
 * DEBUG: Check what getClickUpSprintTasks returns for Week 49
 */

/**
 * DEBUG: Find WaveMax tasks and show structure
 */

/**
 * DEBUG: Simpler approach - just check what getClickUpSprintTasks returns
 */


// Run the test
// debugRefactoredLogic();


// Run comprehensive test
// debugComprehensiveTest();

/**
 * COMPREHENSIVE PRODUCT DEBUGGING
 * This script will:
 * 1. Show all product lists in FB Ads folder
 * 2. Test refactored logic against each product
 * 3. Verify productCode is added to all batches
 */

// ===== COPY YOUR ACTUAL getClickUpSprintTasks FUNCTION HERE =====
// (The refactored version from getClickUpSprintTasks_REFACTORED.js)


// Helper function to match your CSV Data products
function identifyCSVSource(productName, batchName) {
  /**
   * Based on your CSV Data, the "CSV Source" column has these mappings:
   * 0 = AuraNaturals
   * 1 = CamTrix (Main)
   * 2 = CleanlixMold
   * 3 = WiggyDog
   * 4 = Guardality
   * 5 = OriBreeze
   * 6 = Clairu_M2
   * 7 = AliveBlue
   * 8 = KatuChef
   * 9 = CamTrix (TOP5)
   * 10 = WaveMax
   * 11 = WellHeater
   * 12 = TheraWolfRelief (DE)
   * 13 = TheraWolfRelief (US/ES)
   */
  
  const csvSourceMap = {
    'AuraNaturals': 0,
    'CamTrix': [1, 9], // Main and TOP5
    'CleanlixMold': 2,
    'WiggyDog': 3,
    'Guardality': 4,
    'OriBreeze': 5,
    'Clairu': 6,
    'AliveBlue': 7,
    'KatuChef': 8,
    'WaveMax': 10,
    'WellHeater': 11,
    'TheraWolf': [12, 13] // DE and US/ES
  };
  
  // Check if batch name contains product identifier
  for (const [product, source] of Object.entries(csvSourceMap)) {
    if (batchName.includes(product)) {
      return { product, csvSource: source };
    }
  }
  
  return { product: productName, csvSource: null };
}

// Run the test
// testAllProducts();

/**
 * COMPREHENSIVE PRODUCT TEST - CORRECT CLICKUP NAMES
 * Tests ALL products with their exact ClickUp list names
 */


/**
 * DEBUG CamTrix TOP5 batches
 */

/**
 * DEBUG AliveBlue batch 92
 */

/**
 * DEBUG ALL AliveBlue batches in Week 49
 */

/**
 * DEBUG AliveBlue master tasks
 */

// ============================================================================
// DEBUGGING FUNCTION: Check API response for specific products
// This mimics your actual workflow
// ============================================================================


// ============================================================================
// DEBUGGING FUNCTION: Test the new code logic on WiggyDog
// ============================================================================


/**
 * Export execution logs to Google Drive
 * Call this at the END of processCreativeTestsMultiple to save logs
 */

/**
 * 🔍 DEBUG FUNCTION: Investigate Why Specific Batches Are Missing
 * 
 * FIXED VERSION - Uses the same CLICKUP_API_KEY constant from your script
 */

// ===== DEBUG FUNCTIONS FOR WAVEMAX 315/317 =====

/**
 * DEBUG 1: Fetch A10_TV Antena list and show ALL tasks with 315 or 317
 */

/**
 * DEBUG 2: Fetch specific task by ID
 */

/**
 * DEBUG 3: Compare list query vs individual fetch
 */

/**
 * Fetch ALL tasks from a list (handles pagination)
 */


// 🚀 HYBRID SOLUTION: Fetch only what you need, not everything

/**
 * Smart fetch strategy:
 * 1. Use ClickUp's search/filter API to get tasks with Launch Date in range
 * 2. Fetch recently updated batch-named tasks
 * 3. Merge results
 * 
 * Result: ~50-100 tasks instead of 1,468!
 */


// 🎯 SIMPLE PAGINATION: Always fetch 10 pages, no early stopping


// 🚀 FASTEST SOLUTION: Direct ClickUp REST API call

/**
 * Get tasks with Launch Date in range - NO pagination needed!
 * @param {string} listId - ClickUp list ID
 * @param {number} startDate - Week start (timestamp)
 * @param {number} endDate - Week end (timestamp)
 * @returns {Array} - Filtered tasks
 */

/**
 * Process multiple products - PARALLEL execution
 * @param {Array} productList - List of ClickUp products
 * @param {number} startDate - Week start
 * @param {number} endDate - Week end
 * @returns {Array} - All tasks from all products
 */

/**
 * Test the new approach
 */

/**
 * TEST FUNCTION - Run this in Apps Script!
 */

/**
 * FULL TEST - Process ALL products
 */

/**
 * DEBUG: Check what dates are actually in the tasks
 */

/**
 * AUTO-RESTART VERSION
 * Processes one product at a time, saves progress, then triggers itself again
 * This avoids timeout by breaking work into small chunks
 */

function runCreativeTestsWithAutoRestart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = ss.getSheetByName('_PROGRESS') || ss.insertSheet('_PROGRESS');
  
  // Get or initialize progress
  let lastProcessedIndex = progressSheet.getRange('A1').getValue() || 0;
  
  // Get all products to process
  const allProducts = ['WaveMax', 'CamTrix', 'AliveBlue']; // Add your products here
  
  if (lastProcessedIndex >= allProducts.length) {
    // We're done!
    Logger.log('✅ ALL PRODUCTS PROCESSED!');
    progressSheet.getRange('A1').setValue(0); // Reset for next run
    return;
  }
  
  // Process next product
  const currentProduct = allProducts[lastProcessedIndex];
  Logger.log('🔄 Processing product ' + (lastProcessedIndex + 1) + '/' + allProducts.length + ': ' + currentProduct);
  
  try {
    // Run Creative Tests for this ONE product
    runCreativeTestsAnalysis(currentProduct);
    
    // Mark as processed
    lastProcessedIndex++;
    progressSheet.getRange('A1').setValue(lastProcessedIndex);
    
    // If more products remain, trigger next run
    if (lastProcessedIndex < allProducts.length) {
      Logger.log('⏩ Triggering next product in 5 seconds...');
      Utilities.sleep(5000);
      
      // Trigger itself again!
      ScriptApp.newTrigger('runCreativeTestsWithAutoRestart')
        .timeBased()
        .after(1000) // 1 second from now
        .create();
      
      Logger.log('✅ Trigger created for next product');
    } else {
      Logger.log('🎉 ALL DONE!');
    }
    
  } catch (e) {
    Logger.log('❌ Error processing ' + currentProduct + ': ' + e.toString());
    // Still continue to next product
    lastProcessedIndex++;
    progressSheet.getRange('A1').setValue(lastProcessedIndex);
  }
}

/**
 * Run Creative Tests for a single product
 */
function runCreativeTestsAnalysis(productFilter) {
  // This is your existing Creative Tests function
  // Just modify it to accept a single product parameter
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  
  // Set the product filter temporarily
  const filterCell = settingsSheet.getRange('B4'); // Adjust to your actual cell
  const originalFilter = filterCell.getValue();
  
  filterCell.setValue(productFilter);
  
  try {
    // Call your main Creative Tests function
    analyzeCreativeTests(); // Or whatever your main function is called
  } finally {
    // Restore original filter
    filterCell.setValue(originalFilter);
  }
}

/**
 * Reset progress (run this to start fresh)
 */
function resetProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = ss.getSheetByName('_PROGRESS');
  if (progressSheet) {
    progressSheet.getRange('A1').setValue(0);
  }
  Logger.log('✅ Progress reset');
}

/**
 * This is called when user clicks "Analyze Creative Tests" button
 * OR when auto-resuming from checkpoint
 */
function startCreativeTestsWithCheckpoint() {
  const checkpoint = getCheckpoint();
  
  if (checkpoint.isActive) {
    // Resume existing run
    resumeCreativeTests();
  } else {
    // New run - show dialog
    showProductSelectionDialog();
  }
}

function clearEverything() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('accumulated_results');
  props.deleteProperty('ct_checkpoint_active');
  props.deleteProperty('ct_checkpoint_products');
  props.deleteProperty('ct_checkpoint_index');
  props.deleteProperty('ct_checkpoint_week');
  props.deleteProperty('ct_checkpoint_export');
  props.deleteProperty('ct_checkpoint_label');
  props.deleteProperty('ct_checkpoint_start');
  Logger.log('✅ Everything cleared!');
}


// Helper function - you may already have this


function forceWriteAccumulatedResults() {
  const allResults = getAccumulatedResults();
  Logger.log('Found ' + allResults.length + ' accumulated results');
  
  if (allResults && allResults.length > 0) {
    writeCreativeTestsToSheet(allResults);
    Logger.log('✅ Results written to sheet!');
  } else {
    Logger.log('❌ No accumulated results found');
  }
}

function skipToNextProduct() {
  const checkpoint = getCheckpoint();
  Logger.log('Current index: ' + checkpoint.currentIndex);
  Logger.log('Current product: ' + checkpoint.products[checkpoint.currentIndex]);
  Logger.log('Next product: ' + checkpoint.products[checkpoint.currentIndex + 1]);
  
  // Skip to next (WellHeater)
  saveCheckpoint(
    checkpoint.products,
    checkpoint.currentIndex + 1,
    checkpoint.weekId,
    checkpoint.shouldExport,
    checkpoint.weekLabel
  );
  
  Logger.log('✅ Skipped! Now run resumeCreativeTests() to continue');
}

function finishRemainingProducts() {
  const weekId = PropertiesService.getScriptProperties().getProperty('ct_checkpoint_week');
  const weekLabel = PropertiesService.getScriptProperties().getProperty('ct_checkpoint_label');
  
  Logger.log('Week ID: ' + weekId);
  Logger.log('Week Label: ' + weekLabel);
  
  // Process WellHeater
  Logger.log('Processing WellHeater...');
  const result1 = analyzeCreativeTests(weekId, 'WellHeater');
  if (result1.success && result1.results.length > 0) {
    appendCreativeTestsToSheet(result1.results);
    Logger.log('✅ WellHeater done: ' + result1.results.length + ' batches');
  }
  
  // Process WiggyDog
  Logger.log('Processing WiggyDog...');
  const result2 = analyzeCreativeTests(weekId, 'WiggyDog');
  if (result2.success && result2.results.length > 0) {
    appendCreativeTestsToSheet(result2.results);
    Logger.log('✅ WiggyDog done: ' + result2.results.length + ' batches');
  }
  
  // Clean up checkpoint
  clearCheckpoint();
  clearAccumulatedResults();
  Logger.log('🎉 All done!');
}

function appendCreativeTestsToSheet(testResults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  if (!testSheet) {
    Logger.log('❌ Creative Tests sheet not found!');
    return;
  }
  
  // Find last row with data
  const lastRow = testSheet.getLastRow();
  let row = lastRow + 2; // Leave a blank row
  
  // Group by product
  const resultsByProduct = {};
  testResults.forEach(result => {
    const product = result.product || 'Unknown';
    if (!resultsByProduct[product]) {
      resultsByProduct[product] = [];
    }
    resultsByProduct[product].push(result);
  });
  
  // Write each product
  Object.keys(resultsByProduct).sort().forEach(productName => {
    const productResults = resultsByProduct[productName];
    
    // Product header
    testSheet.getRange(row, 1, 1, 10).merge();
    testSheet.getRange(row, 1).setValue('📦 ' + productName)
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#E8F0FE');
    row++;
    
    // Write batches
    productResults.forEach(result => {
      const rowData = [
        result.isVideo ? '🎥' : '📷',
        result.batchName || result.adName,
        result.spend || '-',
        result.roas || '-',
        result.cpa || '-',
        result.purchases || '-',
        result.hookRate || '-',
        result.holdRate || '-',
        result.targetCpa || '-',
        result.summary || ''
      ];
      testSheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
    
    row++; // Blank row between products
  });
  
  Logger.log('✅ Appended ' + testResults.length + ' results starting at row ' + (lastRow + 2));
}

// ========================================
// PARALLEL CREATIVE TESTS PROCESSING
// ========================================

/**
 * Start parallel processing of all products
 */
function processCreativeTestsParallel(sprintListId, selectedProducts, shouldExport) {
  const props = PropertiesService.getScriptProperties();
  
  // Clear any existing parallel state
  clearParallelState();
  
  // Get week label
  const sprints = getClickUpSprints();
  const selectedSprint = sprints.find(s => s.id === sprintListId);
  const weekLabel = selectedSprint ? selectedSprint.name : 'Unknown Week';
  
  // Store parallel config
  props.setProperty('CT_PARALLEL_ACTIVE', 'true');
  props.setProperty('CT_PARALLEL_WEEK_ID', sprintListId);
  props.setProperty('CT_PARALLEL_WEEK_LABEL', weekLabel);
  props.setProperty('CT_PARALLEL_EXPORT', shouldExport ? 'true' : 'false');
  props.setProperty('CT_PARALLEL_START_TIME', new Date().toISOString());
  
  // Initialize queue - each product as {name, status}
  const queue = selectedProducts.map(function(product) {
    return { name: product, status: 'pending' };
  });
  props.setProperty('CT_PARALLEL_QUEUE', JSON.stringify(queue));
  
  Logger.log('=== STARTING PARALLEL PROCESSING ===');
  Logger.log('Products: ' + selectedProducts.length);
  Logger.log('Week: ' + weekLabel);
  
  // Create parallel workers (6 simultaneous workers is a good balance)
  const NUM_WORKERS = 6;
  for (let i = 0; i < NUM_WORKERS; i++) {
    ScriptApp.newTrigger('parallelWorker')
      .timeBased()
      .after((i * 1500) + 1000) // Stagger by 1.5 seconds
      .create();
  }
  
  Logger.log('Created ' + NUM_WORKERS + ' parallel workers');
  
  // Show progress sidebar
  showParallelProgressSidebar();
}

/**
 * Worker function - picks up next pending product and processes it
 */
function parallelWorker() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  
  // Check if parallel processing is still active
  if (props.getProperty('CT_PARALLEL_ACTIVE') !== 'true') {
    return;
  }
  
  // Try to get a product to process (with lock to prevent race conditions)
  let productToProcess = null;
  
  try {
    // Wait up to 10 seconds for lock
    if (!lock.tryLock(10000)) {
      Logger.log('Could not acquire lock, will retry');
      scheduleWorkerRetry();
      return;
    }
    
    // Get queue and find next pending product
    const queue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
    
    for (let i = 0; i < queue.length; i++) {
      if (queue[i].status === 'pending') {
        productToProcess = queue[i].name;
        queue[i].status = 'running';
        queue[i].startTime = new Date().toISOString();
        props.setProperty('CT_PARALLEL_QUEUE', JSON.stringify(queue));
        break;
      }
    }
    
    lock.releaseLock();
  } catch (e) {
    Logger.log('Lock error: ' + e.toString());
    try { lock.releaseLock(); } catch (e2) {}
    scheduleWorkerRetry();
    return;
  }
  
  // If no product to process, check if we're done
  if (!productToProcess) {
    checkIfAllDone();
    return;
  }
  
  // Process the product
  Logger.log('=== Worker processing: ' + productToProcess + ' ===');
  const startTime = new Date();
  
  const weekId = props.getProperty('CT_PARALLEL_WEEK_ID');
  const weekLabel = props.getProperty('CT_PARALLEL_WEEK_LABEL');
  const shouldExport = props.getProperty('CT_PARALLEL_EXPORT') === 'true';
  
  let resultCount = 0;
  let error = null;
  
  try {
    const result = analyzeCreativeTests(weekId, productToProcess);
    
    if (result.success && result.results && result.results.length > 0) {
      resultCount = result.results.length;
      
      // Store results
      const safeKey = productToProcess.replace(/[^a-zA-Z0-9]/g, '_');
      props.setProperty('CT_PARALLEL_RESULT_' + safeKey, JSON.stringify(result.results));
      
      // Export if needed
      if (shouldExport) {
        try {
          exportToProductSpreadsheets(result.results, weekLabel, productToProcess);
        } catch (exportError) {
          Logger.log('Export error: ' + exportError.toString());
        }
      }
      
      Logger.log('✅ ' + productToProcess + ': ' + resultCount + ' batches');
    } else {
      // No results - create empty tab if exporting
      if (shouldExport) {
        try {
          let baseProduct = productToProcess;
          const labelMatch = productToProcess.match(/^(.+?)\s*\((.+?)\)$/);
          if (labelMatch) baseProduct = labelMatch[1];
          createEmptyCreativeTestTab(baseProduct, weekLabel);
          
          // ✅ Mark this as empty tab created
          try {
            lock.tryLock(5000);
            const currentQueue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
            for (let qi = 0; qi < currentQueue.length; qi++) {
              if (currentQueue[qi].name === productToProcess) {
                currentQueue[qi].emptyTab = true;
                break;
              }
            }
            props.setProperty('CT_PARALLEL_QUEUE', JSON.stringify(currentQueue));
            lock.releaseLock();
          } catch (lockErr) {
            try { lock.releaseLock(); } catch (e2) {}
          }
        } catch (e) {
          Logger.log('⚠️ Could not create empty tab: ' + e.toString());
        }
      }
      Logger.log('📭 ' + productToProcess + ': No batches');
    }
  } catch (e) {
    error = e.toString();
    Logger.log('❌ ' + productToProcess + ': ' + error);
  }
  
  // Update queue with result
  try {
    lock.tryLock(10000);
    const queue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
    
    for (let i = 0; i < queue.length; i++) {
      if (queue[i].name === productToProcess) {
        queue[i].status = error ? 'error' : 'done';
        queue[i].count = resultCount;
        queue[i].error = error;
        queue[i].duration = Math.round((new Date() - startTime) / 1000);
        break;
      }
    }
    
    props.setProperty('CT_PARALLEL_QUEUE', JSON.stringify(queue));
    lock.releaseLock();
  } catch (e) {
    try { lock.releaseLock(); } catch (e2) {}
  }
  
  // Schedule this worker to pick up another product
  scheduleWorkerRetry();
}

/**
 * Schedule worker to try again
 */
function scheduleWorkerRetry() {
  const props = PropertiesService.getScriptProperties();
  
  // Check if still active
  if (props.getProperty('CT_PARALLEL_ACTIVE') !== 'true') {
    return;
  }
  
  // Check if any products still pending
  const queue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
  const hasPending = queue.some(function(p) { return p.status === 'pending'; });
  
  if (hasPending) {
    ScriptApp.newTrigger('parallelWorker')
      .timeBased()
      .after(2000)
      .create();
  } else {
    // No more pending, check if all done
    checkIfAllDone();
  }
}

/**
 * Check if all products are done and trigger finalization
 */
function checkIfAllDone() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  
  if (props.getProperty('CT_PARALLEL_ACTIVE') !== 'true') {
    return;
  }
  
  try {
    if (!lock.tryLock(5000)) return;
    
    const queue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
    const allDone = queue.every(function(p) { 
      return p.status === 'done' || p.status === 'error'; 
    });
    
    if (allDone) {
      // Check if already scheduled finalization
      if (props.getProperty('CT_PARALLEL_FINALIZING') === 'true') {
        lock.releaseLock();
        return;
      }
      
      // Mark as finalizing
      props.setProperty('CT_PARALLEL_FINALIZING', 'true');
      lock.releaseLock();
      
      // Schedule write as separate execution
      Logger.log('All products done, scheduling write to sheet...');
      ScriptApp.newTrigger('combineParallelResults')
        .timeBased()
        .after(2000)
        .create();
    } else {
      lock.releaseLock();
    }
  } catch (e) {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * Combine all parallel results and write to sheet
 */
function combineParallelResults() {
  const props = PropertiesService.getScriptProperties();
  
  // Mark writing phase
  props.setProperty('CT_PARALLEL_WRITING', 'true');
  
  const queue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
  const weekLabel = props.getProperty('CT_PARALLEL_WEEK_LABEL');
  const startTime = props.getProperty('CT_PARALLEL_START_TIME');
  
  Logger.log('=== COMBINING RESULTS ===');
  
  const allResults = [];
  let totalBatches = 0;
  let successCount = 0;
  let errorCount = 0;
  
  queue.forEach(function(item) {
    if (item.status === 'done') {
      successCount++;
      const safeKey = item.name.replace(/[^a-zA-Z0-9]/g, '_');
      const resultJson = props.getProperty('CT_PARALLEL_RESULT_' + safeKey);
      
      if (resultJson) {
        try {
          const results = JSON.parse(resultJson);
          allResults.push.apply(allResults, results);
          totalBatches += results.length;
        } catch (e) {}
      }
    } else if (item.status === 'error') {
      errorCount++;
    }
  });
  
  Logger.log('Total batches: ' + totalBatches);
  Logger.log('Success: ' + successCount + ', Errors: ' + errorCount);
  
  // Write to sheet
  if (allResults.length > 0) {
    // ✅ FIX: Update LAST_ANALYZED_WEEK before writing so the sheet shows correct week
    props.setProperty('LAST_ANALYZED_WEEK', weekLabel);
    
    writeCreativeTestsToSheet(allResults);
    Logger.log('✅ Written to Creative Tests sheet');
  }
  
  // Mark as complete
  props.setProperty('CT_PARALLEL_COMPLETE', 'true');
  props.setProperty('CT_PARALLEL_WRITING', 'false');
  
  // Calculate duration
  const duration = startTime ? 
    Math.round((new Date() - new Date(startTime)) / 1000 / 60) : 0;
  
  // Store summary for sidebar
  props.setProperty('CT_PARALLEL_SUMMARY', JSON.stringify({
    total: queue.length,
    batches: totalBatches,
    success: successCount,
    errors: errorCount,
    duration: duration
  }));
  
  // Clean up triggers (but keep state for sidebar to show completion)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    const name = trigger.getHandlerFunction();
    if (name === 'parallelWorker' || name === 'checkParallelProgress' || name === 'combineParallelResults') {
      try { ScriptApp.deleteTrigger(trigger); } catch (e) {}
    }
  });
  
  Logger.log('✅ Parallel processing complete! Duration: ' + duration + ' minutes');
}

/**
 * Clear all parallel state
 */
function clearParallelState() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  
  Object.keys(allProps).forEach(function(key) {
    if (key.startsWith('CT_PARALLEL_')) {
      props.deleteProperty(key);
    }
  });
  
  // Clean up triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    const name = trigger.getHandlerFunction();
    if (name === 'parallelWorker' || name === 'checkParallelProgress' || name === 'combineParallelResults') {
      try { ScriptApp.deleteTrigger(trigger); } catch (e) {}
    }
  });
  
  Logger.log('Cleared parallel state');
}

/**
 * Show progress sidebar
 */
function showParallelProgressSidebar() {
  const html = HtmlService.createHtmlOutput(buildParallelProgressHtml())
    .setTitle('Creative Tests Progress')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Build progress HTML
 */
function buildParallelProgressHtml() {
  return '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body { font-family: -apple-system, sans-serif; padding: 16px; }' +
    'h3 { margin: 0 0 16px 0; }' +
    '.product { padding: 8px; margin: 4px 0; border-radius: 4px; display: flex; justify-content: space-between; align-items: center; }' +
    '.product.done { background: #e6f4ea; }' +
    '.product.running { background: #fff3e0; }' +
    '.product.error { background: #fce8e6; }' +
    '.product.pending { background: #f5f5f5; }' +
    '.status { font-size: 12px; }' +
    '.progress-bar { height: 8px; background: #e0e0e0; border-radius: 4px; margin: 16px 0; }' +
    '.progress-fill { height: 100%; background: #4285f4; border-radius: 4px; transition: width 0.5s; }' +
    '.stats { font-size: 13px; color: #666; margin-bottom: 16px; }' +
    '.phase-indicator { padding: 12px; border-radius: 6px; margin-bottom: 16px; text-align: center; }' +
    '.phase-indicator.writing { background: #fff3e0; color: #e65100; }' +
    '.phase-indicator.complete { background: #e6f4ea; color: #1e7e34; }' +
    '.summary { font-size: 12px; color: #666; margin-top: 16px; padding-top: 16px; border-top: 1px solid #e0e0e0; }' +
    '.summary div { margin: 4px 0; }' +
    '</style></head><body>' +
    '<h3 id="title">⏳ Processing...</h3>' +
    '<div class="progress-bar"><div class="progress-fill" id="progressFill" style="width: 0%"></div></div>' +
    '<div class="stats" id="stats">Starting...</div>' +
    '<div id="phaseIndicator"></div>' +
    '<div id="productList"></div>' +
    '<div id="summary" class="summary" style="display:none;"></div>' +
    '<script>' +
    'function refresh() {' +
    '  google.script.run.withSuccessHandler(function(data) {' +
    '    if (!data) { setTimeout(refresh, 2000); return; }' +
    '    var pct = data.total > 0 ? Math.round((data.completed / data.total) * 100) : 0;' +
    '    document.getElementById("progressFill").style.width = pct + "%";' +
    '    document.getElementById("stats").innerHTML = data.completed + "/" + data.total + " products (" + pct + "%)";' +
    '    ' +
    '    var phaseEl = document.getElementById("phaseIndicator");' +
    '    if (data.phase === "writing") {' +
    '      phaseEl.className = "phase-indicator writing";' +
    '      phaseEl.innerHTML = "📝 Writing to sheet...";' +
    '      document.getElementById("title").innerHTML = "📝 Saving...";' +
    '    } else if (data.phase === "finalizing") {' +
    '      phaseEl.className = "phase-indicator writing";' +
    '      phaseEl.innerHTML = "⏳ Preparing results...";' +
    '    } else if (data.phase === "complete") {' +
    '      phaseEl.className = "phase-indicator complete";' +
    '      phaseEl.innerHTML = "✅ All done!";' +
    '      document.getElementById("title").innerHTML = "✅ Complete!";' +
    '      if (data.summary) {' +
    '        var sumEl = document.getElementById("summary");' +
    '        sumEl.style.display = "block";' +
    '        sumEl.innerHTML = "<div><strong>Summary</strong></div>" +' +
    '          "<div>Total batches: " + data.summary.batches + "</div>" +' +
    '          "<div>Duration: " + data.summary.duration + " min</div>" +' +
    '          "<div>Errors: " + data.summary.errors + "</div>";' +
    '      }' +
    '    } else {' +
    '      phaseEl.innerHTML = "";' +
    '    }' +
    '    ' +
    '    var html = "";' +
    '    data.products.forEach(function(p) {' +
    '      html += "<div class=\\"product " + p.status + "\\"><span>" + p.name + "</span><span class=\\"status\\">" + p.icon + "</span></div>";' +
    '    });' +
    '    document.getElementById("productList").innerHTML = html;' +
    '    ' +
    '    if (data.phase !== "complete") { setTimeout(refresh, 2000); }' +
    '  }).getParallelProgressData();' +
    '}' +
    'refresh();' +
    '</script></body></html>';
}

/**
 * Get progress data for sidebar
 */
function getParallelProgressData() {
  const props = PropertiesService.getScriptProperties();
  
  if (props.getProperty('CT_PARALLEL_ACTIVE') !== 'true') {
    return null;
  }
  
  let queue = JSON.parse(props.getProperty('CT_PARALLEL_QUEUE') || '[]');
  const isWriting = props.getProperty('CT_PARALLEL_WRITING') === 'true';
  const isComplete = props.getProperty('CT_PARALLEL_COMPLETE') === 'true';
  const summaryJson = props.getProperty('CT_PARALLEL_SUMMARY');
  
  // No artificial timeout - let workers run as long as needed
  
  let completed = 0;
  
  const products = queue.map(function(item) {
    if (item.status === 'done' || item.status === 'error') completed++;
    
    var icon = item.status === 'pending' ? '⏳' : 
               item.status === 'running' ? '🔄' : 
               item.status === 'done' ? (item.emptyTab ? '📭 0' : '✅ ' + (item.count || 0)) : 
               '❌';
    
    return { name: item.name, status: item.status, icon: icon };
  });
  
  // Determine phase
  let phase = 'processing';
  if (completed >= queue.length) {
    if (isComplete) {
      phase = 'complete';
    } else if (isWriting) {
      phase = 'writing';
    } else {
      phase = 'finalizing';
    }
  }
  
  // Parse summary if complete
  let summary = null;
  if (summaryJson) {
    try { summary = JSON.parse(summaryJson); } catch (e) {}
  }
  
  return { 
    total: queue.length, 
    completed: completed, 
    products: products,
    phase: phase,
    summary: summary
  };
}

function testAtriaConfig() {
  initializeConfigConstants_();
  const folders = PRODUCT_LOCALE_FOLDERS;
  const cpa = TARGET_CPA;
  
  Logger.log('Products with folders: ' + Object.keys(folders).length);
  Logger.log('Sample folder (CamTrix US): ' + (folders.CamTrix ? folders.CamTrix.US : 'NOT FOUND'));
  
  Logger.log('Products with CPA: ' + Object.keys(cpa).length);
  Logger.log('Sample CPA (CamTrix): ' + cpa.CamTrix);
  Logger.log('Sample CPA (AliveBlue): ' + cpa.AliveBlue);
}

function runAtriaReportsDiagnostics() {
  initializeConfigConstants_();
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
  Logger.log('║          ATRIA REPORTS - CONFIG DIAGNOSTICS REPORT              ║');
  Logger.log('║          ' + new Date().toLocaleString() + '                            ║');
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
    
    const requiredTabs = ['Products', 'Product Folders', 'ClickUp Config'];
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
  // TEST 2: PRODUCT LOCALE FOLDERS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📁 TEST 2: PRODUCT LOCALE FOLDERS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const folders = PRODUCT_LOCALE_FOLDERS;
    if (typeof folders === 'object' && Object.keys(folders).length > 0) {
      const productCount = Object.keys(folders).length;
      log('Folders', 'Load Product Folders', 'PASS', `${productCount} product(s)`);
      
      // Count total folder links
      let totalLinks = 0;
      let emptyLinks = 0;
      
      for (const prodKey of Object.keys(folders)) {
        const prodFolders = folders[prodKey];
        for (const locale of Object.keys(prodFolders)) {
          const link = prodFolders[locale];
          if (link && link.includes('drive.google.com')) {
            totalLinks++;
          } else {
            emptyLinks++;
          }
        }
      }
      
      log('Folders', 'Folder Links', 'PASS', `${totalLinks} links, ${emptyLinks} empty`);
      
      // Test specific folder lookup
      const testCases = [
        { product: 'CamTrix', locale: 'US' },
        { product: 'AliveBlue', locale: 'DE' }
      ];
      
      for (const tc of testCases) {
        const folder = folders[tc.product] && folders[tc.product][tc.locale];
        if (folder && folder.includes('drive.google.com')) {
          Logger.log(`   ✅ ${tc.product}/${tc.locale}: Found`);
        } else {
          Logger.log(`   ⚠️ ${tc.product}/${tc.locale}: Empty or missing`);
        }
      }
    } else {
      log('Folders', 'Load Product Folders', 'FAIL', 'No folders found');
    }
  } catch (e) {
    log('Folders', 'Load Product Folders', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 3: TARGET CPA
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('💰 TEST 3: TARGET CPA');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const cpa = TARGET_CPA;
    if (typeof cpa === 'object' && Object.keys(cpa).length > 0) {
      const productCount = Object.keys(cpa).length;
      log('CPA', 'Load Target CPA', 'PASS', `${productCount} product(s) with CPA`);
      
      // Show sample values
      const sampleProducts = ['CamTrix', 'AliveBlue', 'TellyStick', 'Guardality'];
      for (const prod of sampleProducts) {
        if (cpa[prod]) {
          Logger.log(`   ${prod}: $${cpa[prod]}`);
        }
      }
      
      // Check for reasonable CPA values (should be 20-200 typically)
      const values = Object.values(cpa);
      const avgCPA = values.reduce((a, b) => a + Number(b), 0) / values.length;
      
      if (avgCPA > 0 && avgCPA < 500) {
        log('CPA', 'CPA Values Sanity', 'PASS', `Average CPA: $${avgCPA.toFixed(2)}`);
      } else {
        log('CPA', 'CPA Values Sanity', 'WARN', `Average CPA: $${avgCPA.toFixed(2)} (unusual)`);
      }
    } else {
      log('CPA', 'Load Target CPA', 'FAIL', 'No CPA data found');
    }
  } catch (e) {
    log('CPA', 'Load Target CPA', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 4: CLICKUP CONFIG
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔗 TEST 4: CLICKUP CONFIG');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  // Check if using centralized config or hardcoded
  try {
    // Check if getClickUpConfigFromConfig_ exists
    if (typeof getClickUpConfigFromConfig_ === 'function') {
      const config = getClickUpConfigFromConfig_();
      log('ClickUp', 'Config Function Exists', 'PASS', 'Using centralized config');
      
      // Check required keys
      const requiredKeys = ['TEAM_ID', 'FB_ADS_FOLDER_ID', 'SPRINT_FOLDER_ID'];
      const missingKeys = requiredKeys.filter(k => !config[k]);
      
      if (missingKeys.length === 0) {
        log('ClickUp', 'Required Keys', 'PASS', 'All keys present');
        Logger.log(`   TEAM_ID: ${config.TEAM_ID}`);
        Logger.log(`   FB_ADS_FOLDER_ID: ${config.FB_ADS_FOLDER_ID}`);
        Logger.log(`   SPRINT_FOLDER_ID: ${config.SPRINT_FOLDER_ID}`);
      } else {
        log('ClickUp', 'Required Keys', 'WARN', `Missing: ${missingKeys.join(', ')}`);
      }
      
      // Check if API_KEY is in config (for centralized approach)
      if (config.API_KEY && config.API_KEY.startsWith('pk_')) {
        log('ClickUp', 'API Key in Config', 'PASS', 'Centralized API key found');
      } else {
        log('ClickUp', 'API Key in Config', 'WARN', 'API key not in config (may be hardcoded)');
      }
    } else {
      log('ClickUp', 'Config Function', 'WARN', 'getClickUpConfigFromConfig_ not found - add FIX 1');
    }
    
        // Check if using centralized config (via lazy loader)
    try {
      const configData = getClickUpConfigDataLazy_();
      if (configData && Object.keys(configData).length > 0) {
        log('ClickUp', 'ClickUp Config Data', 'PASS', 'Using centralized ClickUp config');
      }
    } catch (e) {
      log('ClickUp', 'ClickUp Config Data', 'WARN', 'Could not load config data');
    }
    
  } catch (e) {
    log('ClickUp', 'ClickUp Config', 'FAIL', e.message);
  }

// ============================================================================
  // TEST 5: SLACK CONFIG (CENTRAL)
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('💬 TEST 5: SLACK CONFIG (CENTRAL)');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    if (typeof getSlackConfigFromConfig_ === 'function') {
      const slackConfig = getSlackConfigFromConfig_();
      if (slackConfig && slackConfig.botToken) {
        log('Slack', 'Load Slack Config', 'PASS', 'From central Config');
        Logger.log('   Bot Token: ****' + slackConfig.botToken.slice(-4));
        Logger.log('   Weekly Channel: ' + (slackConfig.weeklyChannelId || 'NOT SET'));
        Logger.log('   Creative Channel: ' + (slackConfig.creativeChannelId || 'NOT SET'));
      } else {
        log('Slack', 'Load Slack Config', 'WARN', 'Not in central Config - using local Settings');
      }
    } else {
      log('Slack', 'Slack Config Function', 'WARN', 'getSlackConfigFromConfig_ not found - add new function');
    }
  } catch (e) {
    log('Slack', 'Slack Config', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 6: PRODUCT SPREADSHEETS (CENTRAL)
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📊 TEST 6: PRODUCT SPREADSHEETS (CENTRAL)');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    if (typeof getProductSpreadsheetsFromConfig_ === 'function') {
      const productSheets = getProductSpreadsheetsFromConfig_();
      if (productSheets && Object.keys(productSheets).length > 0) {
        log('Spreadsheets', 'Load Product Spreadsheets', 'PASS', 
            Object.keys(productSheets).length + ' products from central Config');
        Logger.log('   Sample: ' + Object.keys(productSheets).slice(0, 3).join(', '));
      } else {
        log('Spreadsheets', 'Load Product Spreadsheets', 'WARN', 'Not in central Config - using local Settings');
      }
    } else {
      log('Spreadsheets', 'Product Spreadsheets Function', 'WARN', 'getProductSpreadsheetsFromConfig_ not found - add new function');
    }
  } catch (e) {
    log('Spreadsheets', 'Product Spreadsheets', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 7: CLICKUP API CONNECTION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🌐 TEST 5: CLICKUP API CONNECTION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Use whatever API key is available
    const apiKey = (typeof CLICKUP_API_KEY !== 'undefined') ? CLICKUP_API_KEY : null;
    
    if (apiKey && apiKey.startsWith('pk_')) {
      log('API', 'API Key Available', 'PASS', 'Key found');
      
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
          log('API', 'ClickUp Connection', 'PASS', 'Connected successfully');
        } else {
          log('API', 'ClickUp Connection', 'FAIL', `HTTP ${response.getResponseCode()}`);
        }
      } catch (e) {
        log('API', 'ClickUp Connection', 'FAIL', e.message);
      }
    } else {
      log('API', 'API Key', 'FAIL', 'No valid API key found');
    }
  } catch (e) {
    log('API', 'API Test', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 8: CACHE PERFORMANCE
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('⚡ TEST 6: CACHE PERFORMANCE');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Clear cache
    CacheService.getScriptCache().removeAll(['atria_folders', 'atria_cpa', 'atria_clickup']);
    
    // Uncached load
    const uncachedStart = Date.now();
    getProductLocaleFoldersFromConfig_();
    getTargetCpaFromConfig_();
    const uncachedTime = Date.now() - uncachedStart;
    
    // Cached load
    const cachedStart = Date.now();
    getProductLocaleFoldersFromConfig_();
    getTargetCpaFromConfig_();
    const cachedTime = Date.now() - cachedStart;
    
    log('Cache', 'Uncached Load', uncachedTime < 5000 ? 'PASS' : 'WARN', `${uncachedTime}ms`);
    log('Cache', 'Cached Load', cachedTime < 100 ? 'PASS' : 'WARN', `${cachedTime}ms`);
    log('Cache', 'Speedup', cachedTime < uncachedTime ? 'PASS' : 'WARN', 
        `${Math.round(uncachedTime / Math.max(cachedTime, 1))}x faster`);
  } catch (e) {
    log('Cache', 'Cache Test', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 9: EDGE CASES & DATA VALIDATION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔬 TEST 7: EDGE CASES & DATA VALIDATION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Test non-existent product CPA
    const fakeCPA = getTargetCpaLazy_()['NonExistentProduct123'];
    log('EdgeCase', 'Non-existent Product CPA', fakeCPA === undefined ? 'PASS' : 'WARN',
        fakeCPA === undefined ? 'Returns undefined' : 'Unexpected value');
    
    // Test non-existent folder
    const fakeFolder = getProductLocaleFoldersLazy_()['NonExistentProduct123'];
    log('EdgeCase', 'Non-existent Product Folder', fakeFolder === undefined ? 'PASS' : 'WARN',
        fakeFolder === undefined ? 'Returns undefined' : 'Unexpected value');
    
    // Test LATAM/ES folder aliasing
    const folders = getProductLocaleFoldersLazy_();
    if (folders.CamTrix) {
      const latam = folders.CamTrix.LATAM;
      const es = folders.CamTrix.ES;
      if (latam && es && latam === es) {
        log('EdgeCase', 'LATAM/ES Folder Alias', 'PASS', 'Same folder for both');
      } else if (latam || es) {
        log('EdgeCase', 'LATAM/ES Folder Alias', 'WARN', 'Different or partial folders');
      }
    }
    
    // Test normalizeProductName function if it exists
    if (typeof normalizeProductName === 'function') {
      const normalized = normalizeProductName('camtrix');
      log('EdgeCase', 'Product Name Normalization', normalized === 'CamTrix' ? 'PASS' : 'WARN',
          `"camtrix" → "${normalized}"`);
    }
  } catch (e) {
    log('EdgeCase', 'Edge Case Tests', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 10: CROSS-REFERENCE VALIDATION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔀 TEST 8: CROSS-REFERENCE VALIDATION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const cpaProducts = new Set(Object.keys(TARGET_CPA));
    const folderProducts = new Set(Object.keys(PRODUCT_LOCALE_FOLDERS));
    
    // Check if CPA products have folders
    const cpaWithoutFolders = [...cpaProducts].filter(p => !folderProducts.has(p));
    if (cpaWithoutFolders.length === 0) {
      log('CrossRef', 'CPA Products → Folders', 'PASS', 'All CPA products have folders');
    } else {
      log('CrossRef', 'CPA Products → Folders', 'WARN', 
          `${cpaWithoutFolders.length} CPA products without folders: ${cpaWithoutFolders.slice(0,5).join(', ')}`);
    }
    
    // Check if folder products have CPA
    const foldersWithoutCPA = [...folderProducts].filter(p => !cpaProducts.has(p));
    if (foldersWithoutCPA.length === 0) {
      log('CrossRef', 'Folder Products → CPA', 'PASS', 'All folder products have CPA');
    } else {
      log('CrossRef', 'Folder Products → CPA', 'WARN', 
          `${foldersWithoutCPA.length} folder products without CPA: ${foldersWithoutCPA.join(', ')}`);
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
    'Atria Reports Diagnostics',
    `✅ Passed: ${results.passed}\n⚠️ Warnings: ${results.warnings}\n❌ Failed: ${results.failed}\n\nTime: ${totalTime}ms\n\nSee Execution Log for details.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  return results;
}

/**
 * Quick health check for Atria Reports
 */
function quickAtriaHealthCheck() {
  const checks = [];
  
  try {
    checks.push({ name: 'Config Sheet', ok: !!SpreadsheetApp.openById(CONFIG_SHEET_ID) });
  } catch (e) {
    checks.push({ name: 'Config Sheet', ok: false });
  }
  
  try {
    const folders = getProductLocaleFoldersLazy_();
    checks.push({ name: 'Product Folders', ok: typeof folders === 'object' && Object.keys(folders).length > 0 });
  } catch (e) {
    checks.push({ name: 'Product Folders', ok: false });
  }
  
  try {
    const cpa = getTargetCpaLazy_();
    checks.push({ name: 'Target CPA', ok: typeof cpa === 'object' && Object.keys(cpa).length > 0 });
  } catch (e) {
    checks.push({ name: 'Target CPA', ok: false });
  }
  
  try {
    const apiKey = getClickUpApiKey_();
    checks.push({ name: 'ClickUp API Key', ok: apiKey && apiKey.startsWith('pk_') });
  } catch (e) {
    checks.push({ name: 'ClickUp API Key', ok: false });
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
}// ================================================================================
// ENHANCED CSV CONFIGURATOR - Backend Functions
// Add these functions to AtriaSlackAutomation.gs
// ================================================================================



/**
 * Analyzes CSV files and returns detailed information for the configurator
 * @param {Array} csvContents - Array of CSV objects with fileName and content
 * @returns {Object} - Analysis results
 */
function analyzeCSVFilesForConfigurator(csvContents) {
  Logger.log('═══════════════════════════════════════════════════════════');
  Logger.log('CSV CONFIGURATOR: Analyzing ' + csvContents.length + ' files');
  Logger.log('═══════════════════════════════════════════════════════════');
  
  const analysis = [];
  const allProductsSet = new Set();
  const productToFiles = {}; // Track which files have which products
  const unknownProducts = {}; // Track products not in PRODUCT_LOCALE_FOLDERS
  
  // Get list of valid products (those with Drive folders configured)
  const validProductsList = Object.keys(PRODUCT_LOCALE_FOLDERS).map(p => p.toLowerCase());
  
  // Phase 1: Analyze each CSV
  csvContents.forEach((csvObj, fileIndex) => {
    const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
    const csvFileName = typeof csvObj === 'string' ? `File_${fileIndex + 1}` : (csvObj.fileName || csvObj.name);
    
    Logger.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Analyzing: ' + csvFileName);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    
    // Parse CSV
    let csvData;
    if (csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length) {
      csvData = csv.split('\n').map(line => line.split(';'));
    } else {
      csvData = Utilities.parseCsv(csv);
    }
    
    if (csvData.length === 0) {
      Logger.log('  ⚠️ Empty CSV');
      analysis.push({
        fileIndex: fileIndex,
        fileName: csvFileName,
        totalAds: 0,
        products: {},
        isMixed: false,
        hasDuplicates: false,
        isEmpty: true
      });
      return;
    }
    
    const headers = csvData[0];
    const adNameIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
    const spendIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'spend');
    
    if (adNameIndex === -1) {
      Logger.log('  ⚠️ No "Ad name" column found');
      analysis.push({
        fileIndex: fileIndex,
        fileName: csvFileName,
        totalAds: csvData.length - 1,
        products: {},
        isMixed: false,
        hasDuplicates: false,
        noAdNameColumn: true
      });
      return;
    }
    
    // Count products and collect stats
    const products = {};
    let totalSpend = 0;
    const totalAds = csvData.length - 1;
    
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      const adName = row[adNameIndex];
      if (!adName) continue;
      
      // Extract product
      const product = extractProductFromAdName(adName.toString());
      
      if (product && product !== 'Unknown') {
        const normalizedProduct = normalizeProductName(product);
        
        // Check if this is a known/valid product
        const isValidProduct = validProductsList.includes(normalizedProduct.toLowerCase());
        
        if (!isValidProduct) {
          // Track unknown product
          if (!unknownProducts[normalizedProduct]) {
            unknownProducts[normalizedProduct] = {
              count: 0,
              files: [],
              sampleAds: []
            };
          }
          unknownProducts[normalizedProduct].count++;
          if (!unknownProducts[normalizedProduct].files.includes(csvFileName)) {
            unknownProducts[normalizedProduct].files.push(csvFileName);
          }
          if (unknownProducts[normalizedProduct].sampleAds.length < 2) {
            unknownProducts[normalizedProduct].sampleAds.push(adName.toString().substring(0, 60));
          }
          continue; // Skip unknown products - don't add to analysis
        }
        
        if (!products[normalizedProduct]) {
          products[normalizedProduct] = {
            count: 0,
            spend: 0,
            percent: 0,
            sampleAds: []
          };
        }
        
        products[normalizedProduct].count++;
        
        // Track spend if available
        if (spendIndex !== -1 && row[spendIndex]) {
          const spendStr = row[spendIndex].toString().replace(/[€$,]/g, '').replace(',', '.');
          const spend = parseFloat(spendStr) || 0;
          products[normalizedProduct].spend += spend;
          totalSpend += spend;
        }
        
        // Store sample ads (first 3)
        if (products[normalizedProduct].sampleAds.length < 3) {
          products[normalizedProduct].sampleAds.push(adName.toString().substring(0, 60));
        }
        
        allProductsSet.add(normalizedProduct);
        
        // Track product to file mapping
        if (!productToFiles[normalizedProduct]) {
          productToFiles[normalizedProduct] = [];
        }
        if (!productToFiles[normalizedProduct].some(f => f.fileIndex === fileIndex)) {
          productToFiles[normalizedProduct].push({
            fileIndex: fileIndex,
            fileName: csvFileName
          });
        }
      }
    }
    
    // Calculate percentages
    Object.keys(products).forEach(p => {
      products[p].percent = totalAds > 0 ? Math.round((products[p].count / totalAds) * 100) : 0;
    });
    
    // Check if mixed
    const productKeys = Object.keys(products);
    const isMixed = productKeys.length > 1;
    
    // Check for compound products (TheraWolf, etc)
    const compoundProducts = ['TheraWolfRelief', 'TheraWolfNeuro', 'CleanlixMold', 'CleanlixPowder', 'Clairu_M2'];
    const hasCompound = productKeys.some(p => compoundProducts.includes(p));
    
    // Determine majority/minority
    let majorityProduct = null;
    let minorityProducts = [];
    
    if (isMixed) {
      const sorted = productKeys.sort((a, b) => products[b].count - products[a].count);
      majorityProduct = sorted[0];
      minorityProducts = sorted.slice(1);
    }
    
    Logger.log('  Products: ' + productKeys.join(', '));
    Logger.log('  Total ads: ' + totalAds);
    Logger.log('  Is mixed: ' + isMixed);
    Logger.log('  Has compound: ' + hasCompound);
    
    analysis.push({
      fileIndex: fileIndex,
      fileName: csvFileName,
      totalAds: totalAds,
      totalSpend: totalSpend,
      products: products,
      isMixed: isMixed && !hasCompound, // Don't flag compound as mixed
      hasCompound: hasCompound,
      majorityProduct: majorityProduct,
      minorityProducts: minorityProducts,
      hasDuplicates: false, // Will be set in Phase 2
      duplicateWith: []
    });
  });
  
  // Phase 2: Check for duplicates across files
  Logger.log('\n═══════════════════════════════════════════════════════════');
  Logger.log('Checking for duplicate products across files...');
  Logger.log('═══════════════════════════════════════════════════════════');
  
  Object.keys(productToFiles).forEach(product => {
    const files = productToFiles[product];
    
    // Filter to only files where this is a majority product (>20%)
    const majorityFiles = files.filter(f => {
      const csvAnalysis = analysis[f.fileIndex];
      const productData = csvAnalysis.products[product];
      return productData && productData.percent > 20;
    });
    
    if (majorityFiles.length > 1) {
      Logger.log('  Duplicate: ' + product + ' in ' + majorityFiles.length + ' files');
      
      // Mark each file as having duplicates
      majorityFiles.forEach(f => {
        analysis[f.fileIndex].hasDuplicates = true;
        analysis[f.fileIndex].duplicateWith = majorityFiles.filter(of => of.fileIndex !== f.fileIndex);
      });
    }
  });
  
  Logger.log('\n═══════════════════════════════════════════════════════════');
  Logger.log('Analysis complete');
  Logger.log('═══════════════════════════════════════════════════════════');
  
  return {
    csvAnalysis: analysis,
    allProducts: Array.from(allProductsSet).sort(),
    productToFiles: productToFiles,
    unknownProducts: unknownProducts
  };
}

/**
 * Check CSV data quality - identify unknown products before generating report
 * Can be run anytime after CSV upload to see what products will be excluded
 */
function checkCSVDataQuality() {
  initializeConfigConstants_();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet || dataSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('❌ No CSV data found.\n\nPlease upload CSV files first using "Upload CSV Files" menu.');
    return;
  }
  
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const adNameCol = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
  
  if (adNameCol === -1) {
    SpreadsheetApp.getUi().alert('❌ No "Ad name" column found in CSV Data sheet.');
    return;
  }
  
  const validProducts = Object.keys(PRODUCT_LOCALE_FOLDERS);
  const validProductsLower = validProducts.map(p => p.toLowerCase());
  const unknownProducts = {};
  const validCounts = {};
  
  for (let i = 1; i < data.length; i++) {
    const adName = String(data[i][adNameCol] || '');
    if (!adName) continue;
    
    // Extract product name (same logic as processAtriaData)
    let cleanAdName = adName;
    if (cleanAdName.includes('||')) {
      cleanAdName = cleanAdName.split('||')[0].trim();
    }
    if (cleanAdName.includes('[')) {
      cleanAdName = cleanAdName.replace(/\[.*?\]/g, '').trim();
    }
    
    const parts = cleanAdName.split('_');
    let platformIndex = -1;
    for (let j = 0; j < parts.length; j++) {
      if (['FB', 'YT', 'GDN', 'TT'].includes(parts[j])) {
        platformIndex = j;
        break;
      }
    }
    
    let product = platformIndex > 0 ? parts.slice(0, platformIndex).join('_') : parts[0];
    product = normalizeProductName(product);
    
    // Check if valid
    const validIndex = validProductsLower.indexOf(product.toLowerCase());
    if (validIndex >= 0) {
      const exactProduct = validProducts[validIndex];
      validCounts[exactProduct] = (validCounts[exactProduct] || 0) + 1;
    } else {
      if (!unknownProducts[product]) {
        unknownProducts[product] = { count: 0, examples: [] };
      }
      unknownProducts[product].count++;
      if (unknownProducts[product].examples.length < 2) {
        unknownProducts[product].examples.push(cleanAdName.substring(0, 55));
      }
    }
  }
  
  // Build summary message
  let message = '📊 CSV DATA QUALITY CHECK\\n';
  message += '═══════════════════════════════════════\\n\\n';
  
  // Valid products
  const validProductNames = Object.keys(validCounts).sort();
  message += '✅ VALID PRODUCTS (' + validProductNames.length + '):\\n';
  validProductNames.forEach(p => {
    message += '   • ' + p + ': ' + validCounts[p] + ' ads\\n';
  });
  
  // Unknown products
  const unknownList = Object.keys(unknownProducts).sort();
  if (unknownList.length > 0) {
    message += '\\n⚠️ UNKNOWN PRODUCTS (' + unknownList.length + '):\\n';
    message += '   (will be EXCLUDED from report)\\n\\n';
    unknownList.forEach(p => {
      const info = unknownProducts[p];
      message += '   ❌ "' + p + '": ' + info.count + ' ads\\n';
      info.examples.forEach(ex => {
        message += '      └─ ' + ex + '...\\n';
      });
    });
    message += '\\n💡 To include these products:\\n';
    message += '   Add them to "Product Folders" tab\\n';
    message += '   in BoldOne Creative Config sheet.';
  } else {
    message += '\\n✅ All products are recognized!\\n';
    message += '   Ready to generate report.';
  }
  
  // Show in dialog
  const htmlOutput = HtmlService.createHtmlOutput(
    '<pre style="font-family: Consolas, Monaco, monospace; font-size: 12px; white-space: pre-wrap; line-height: 1.5; padding: 10px;">' + 
    message.replace(/\\n/g, '\n') + 
    '</pre>'
  ).setWidth(480).setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 CSV Data Quality Check');
}

/**
 * Shows the enhanced CSV Configurator modal
 * @param {Array} csvContents - Array of CSV objects
 */
function showCsvConfiguratorModal(csvContents) {
  try {
    // Store CSV contents in cache for later processing
    const cache = CacheService.getScriptCache();
    const csvJson = JSON.stringify(csvContents);
    const chunkSize = 90000;
    const chunks = [];
    
    for (let i = 0; i < csvJson.length; i += chunkSize) {
      chunks.push(csvJson.substring(i, i + chunkSize));
    }
    
    cache.put('CSV_CONFIGURATOR_CHUNKS', chunks.length.toString(), 600);
    chunks.forEach((chunk, idx) => {
      cache.put('CSV_CONFIGURATOR_CHUNK_' + idx, chunk, 600);
    });
    
    Logger.log('✓ Stored CSV data in ' + chunks.length + ' cache chunks');
    
    // Analyze CSVs
    const analysisResult = analyzeCSVFilesForConfigurator(csvContents);
    
    Logger.log('Analysis result: ' + analysisResult.csvAnalysis.length + ' files analyzed');
    Logger.log('Products found: ' + analysisResult.allProducts.join(', '));
    
    // Store analysis result in cache for the modal to fetch
    const analysisJson = JSON.stringify({
      csvAnalysis: analysisResult.csvAnalysis,
      allProducts: analysisResult.allProducts,
      unknownProducts: analysisResult.unknownProducts || {}
    });
    cache.put('CSV_CONFIGURATOR_ANALYSIS', analysisJson, 600);
    
    // Create modal HTML - simpler approach without template variables
    const htmlContent = HtmlService.createHtmlOutputFromFile('csv_configurator_modal')
      .setWidth(950)
      .setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(htmlContent, 'CSV Configurator');
    
  } catch (e) {
    Logger.log('Error showing CSV Configurator: ' + e.toString());
    throw new Error('Failed to show CSV Configurator: ' + e.toString());
  }
}

/**
 * Get CSV Configurator analysis data (called from modal)
 */
function getCsvConfiguratorData() {
  const cache = CacheService.getScriptCache();
  const analysisJson = cache.get('CSV_CONFIGURATOR_ANALYSIS');
  
  if (!analysisJson) {
    return { error: 'No data found. Please upload CSV files first.' };
  }
  
  try {
    return JSON.parse(analysisJson);
  } catch (e) {
    return { error: 'Failed to parse analysis data: ' + e.toString() };
  }
}

/**
 * Apply user decisions from CSV Configurator and process files
 * @param {Object} decisions - User decisions for each CSV
 * @returns {Object} - Processing result
 */
function applyCsvConfiguratorDecisions(decisions) {
  try {
    Logger.log('═══════════════════════════════════════════════════════════');
    Logger.log('Applying CSV Configurator decisions (v2)');
    Logger.log('Decisions: ' + JSON.stringify(decisions));
    Logger.log('═══════════════════════════════════════════════════════════');
    
    // Extract decision components
    const globalOptions = decisions.globalOptions || { combineSameProducts: true };
    const mixedFileDecisions = decisions.mixedFileDecisions || {};
    const productDecisions = decisions.productDecisions || {};
    const unknownProductDecisions = decisions.unknownProductDecisions || {};
    
    Logger.log('Global options: ' + JSON.stringify(globalOptions));
    Logger.log('Mixed file decisions: ' + JSON.stringify(mixedFileDecisions));
    Logger.log('Unknown product decisions: ' + JSON.stringify(unknownProductDecisions));
    
    // Store unknown product mappings for use during processing
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('UNKNOWN_PRODUCT_MAPPINGS', JSON.stringify(unknownProductDecisions));
    
    // Retrieve CSV contents from cache
    const cache = CacheService.getScriptCache();
    const chunksCount = parseInt(cache.get('CSV_CONFIGURATOR_CHUNKS') || '0');
    
    if (chunksCount === 0) {
      throw new Error('CSV contents not found in cache. Please re-upload files.');
    }
    
    let csvJson = '';
    for (let i = 0; i < chunksCount; i++) {
      const chunk = cache.get('CSV_CONFIGURATOR_CHUNK_' + i);
      if (!chunk) {
        throw new Error('CSV chunk ' + i + ' not found. Please re-upload files.');
      }
      csvJson += chunk;
    }
    
    const csvContents = JSON.parse(csvJson);
    Logger.log('Loaded ' + csvContents.length + ' CSV files from cache');
    
    // STEP 1: Process each CSV based on mixed file decisions
    const processedCsvs = [];
    
    csvContents.forEach((csvObj, idx) => {
      const mixedDecision = mixedFileDecisions[idx];
      const fileName = typeof csvObj === 'string' ? `File_${idx + 1}` : (csvObj.fileName || csvObj.name);
      
      // Store original index for later label lookup
      if (typeof csvObj !== 'string') {
        csvObj.originalIndex = idx;
      }
      
      Logger.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('Processing CSV ' + idx + ': ' + fileName);
      Logger.log('Mixed decision: ' + (mixedDecision || 'none'));
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      
      if (mixedDecision === 'skip') {
        Logger.log('  ⏭️ Skipped');
        return; // Skip this file
      }
      
      if (mixedDecision === 'split') {
        // Split into separate CSVs by product
        Logger.log('  ✂️ Splitting by product...');
        const splitCsvs = splitCsvByProduct(csvObj);
        Logger.log('    Split into ' + splitCsvs.length + ' files');
        splitCsvs.forEach(sc => {
          // Preserve original index for split CSVs
          sc.originalIndex = idx;
          Logger.log('      - ' + sc.splitProduct + ': ' + sc.fileName);
          processedCsvs.push(sc);
        });
        return;
      }
      
      if (mixedDecision === 'keep_majority') {
        // Filter out minority products
        Logger.log('  🗑️ Keeping only majority product...');
        const filteredCsv = filterMinorityProductsFromCsv(csvObj);
        filteredCsv.originalIndex = idx;
        processedCsvs.push(filteredCsv);
        return;
      }
      
      // No special decision - keep as-is
      Logger.log('  ✅ Keeping as-is');
      processedCsvs.push(csvObj);
    });
    
    Logger.log('\n═══════════════════════════════════════════════════════════');
    Logger.log('After processing decisions: ' + processedCsvs.length + ' CSV segments');
    Logger.log('═══════════════════════════════════════════════════════════');
    
    // STEP 2: If combineSameProducts is ON, we process normally
    // The main processCSVData will handle the combination via CSV Source column
    // But we need to ensure same products get the same CSV Source
    
    if (globalOptions.combineSameProducts) {
      Logger.log('🔗 Combine same products is ON - grouping by product...');
      
      // Group CSVs by their detected product
      const productGroups = {};
      const csvToOriginalIdx = new Map(); // Track original file index for labels
      
      processedCsvs.forEach((csvObj, idx) => {
        const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
        const fileName = typeof csvObj === 'string' ? `File_${idx + 1}` : (csvObj.fileName || csvObj.name);
        const originalIdx = csvObj.originalIndex !== undefined ? csvObj.originalIndex : idx;
        
        // Detect the main product in this CSV
        const product = detectMainProductInCsv(csv);
        
        if (product) {
          if (!productGroups[product]) {
            productGroups[product] = [];
          }
          // Store with original index for label lookup
          csvObj.originalIndex = originalIdx;
          productGroups[product].push(csvObj);
          Logger.log('  ' + fileName + ' → ' + product + ' (original idx: ' + originalIdx + ')');
        } else {
          // Unknown product - keep separate
          if (!productGroups['_unknown_']) {
            productGroups['_unknown_'] = [];
          }
          productGroups['_unknown_'].push(csvObj);
          Logger.log('  ' + fileName + ' → Unknown product');
        }
      });
      
      // Process each product group based on per-product decisions
      const finalCsvs = [];
      
      Object.keys(productGroups).forEach(product => {
        const group = productGroups[product];
        const productDecision = productDecisions[product];
        
        Logger.log('\n  Processing product: ' + product + ' (' + group.length + ' files)');
        Logger.log('  Decision: ' + JSON.stringify(productDecision));
        
        // Check if this product should be kept separate with labels
        if (productDecision && productDecision.action === 'separate') {
          Logger.log('  📑 Keeping separate with labels...');
          
          // Apply labels to each file
          group.forEach(csvObj => {
            const originalIdx = csvObj.originalIndex;
            const label = productDecision.labels?.[originalIdx] || '';
            
            if (label) {
              Logger.log('    File ' + originalIdx + ' → Label: "' + label + '"');
              csvObj.productLabel = label;
            } else {
              // No label provided - use filename as fallback
              const fileName = typeof csvObj === 'string' ? `File_${originalIdx + 1}` : (csvObj.fileName || csvObj.name);
              Logger.log('    File ' + originalIdx + ' → No label (using filename)');
            }
            
            finalCsvs.push(csvObj);
          });
        } else if (group.length === 1) {
          // Only one file for this product - keep as-is
          Logger.log('  Single file - keeping as-is');
          finalCsvs.push(group[0]);
        } else {
          // Multiple files and should combine
          Logger.log('  🔗 Combining ' + group.length + ' files...');
          let combined = group[0];
          for (let i = 1; i < group.length; i++) {
            combined = combineCsvFiles(combined, group[i]);
          }
          // Mark as combined
          combined.combinedProduct = product;
          combined.combinedFromCount = group.length;
          finalCsvs.push(combined);
        }
      });
      
      Logger.log('\nAfter processing: ' + finalCsvs.length + ' final CSV segments');
      
      // Clear cache
      clearConfiguratorCache(cache, chunksCount);
      
      // Process the final CSVs
      return processAndReturn(finalCsvs);
    }
    
    // STEP 3: If combineSameProducts is OFF, check for per-product separate decisions with labels
    Logger.log('🔗 Combine same products is OFF - checking for labels...');
    
    // Apply any labels from productDecisions
    processedCsvs.forEach((csvObj, idx) => {
      const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
      const product = detectMainProductInCsv(csv);
      
      if (product && productDecisions[product] && productDecisions[product].action === 'separate') {
        const originalIdx = csvObj.originalIndex !== undefined ? csvObj.originalIndex : idx;
        const label = productDecisions[product].labels?.[originalIdx] || '';
        
        if (label) {
          Logger.log('  Applying label "' + label + '" to file ' + idx);
          csvObj.productLabel = label;
        }
      }
    });
    
    // Clear cache
    clearConfiguratorCache(cache, chunksCount);
    
    // Process as separate files
    return processAndReturn(processedCsvs);
    
  } catch (e) {
    Logger.log('Error applying decisions: ' + e.toString());
    return {
      success: false,
      message: 'Error: ' + e.toString()
    };
  }
}

/**
 * Helper: Clear configurator cache
 */
function clearConfiguratorCache(cache, chunksCount) {
  cache.remove('CSV_CONFIGURATOR_CHUNKS');
  cache.remove('CSV_CONFIGURATOR_ANALYSIS');
  for (let i = 0; i < chunksCount; i++) {
    cache.remove('CSV_CONFIGURATOR_CHUNK_' + i);
  }
}

/**
 * Helper: Process CSVs and return result
 */
function processAndReturn(csvs) {
  if (csvs.length === 0) {
    return {
      success: false,
      message: 'No CSV files to process after applying decisions.'
    };
  }
  
  // Call the main processCSVData function with a flag to skip configurator
  const result = processCSVDataDirect(csvs);
  
  if (result && result.success === false) {
    return result;
  }
  
  return {
    success: true,
    message: 'Successfully processed ' + csvs.length + ' CSV segment(s)!'
  };
}

/**
 * Detect the main (majority) product in a CSV
 */
function detectMainProductInCsv(csv) {
  try {
    let csvData;
    if (csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length) {
      csvData = csv.split('\n').map(line => line.split(';'));
    } else {
      csvData = Utilities.parseCsv(csv);
    }
    
    if (csvData.length < 2) return null;
    
    const headers = csvData[0];
    const adNameIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
    
    if (adNameIndex === -1) return null;
    
    // Count products
    const productCounts = {};
    
    for (let i = 1; i < csvData.length; i++) {
      const adName = csvData[i][adNameIndex];
      if (!adName) continue;
      
      const product = extractProductFromAdName(adName.toString());
      if (product && product !== 'Unknown') {
        const normalized = normalizeProductName(product);
        productCounts[normalized] = (productCounts[normalized] || 0) + 1;
      }
    }
    
    // Find majority product
    let maxCount = 0;
    let mainProduct = null;
    
    Object.keys(productCounts).forEach(p => {
      if (productCounts[p] > maxCount) {
        maxCount = productCounts[p];
        mainProduct = p;
      }
    });
    
    return mainProduct;
  } catch (e) {
    Logger.log('Error detecting main product: ' + e.toString());
    return null;
  }
}

/**
 * Process CSV data directly (bypasses configurator check)
 */
function processCSVDataDirect(csvContents) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dataSheet = ss.getSheetByName('CSV Data');
    
    if (!dataSheet) {
      dataSheet = ss.insertSheet('CSV Data');
    }
    
    // Preserve existing labels
    const existingLabels = {};
    if (dataSheet.getLastRow() > 0) {
      try {
        const existingData = dataSheet.getDataRange().getValues();
        for (let i = 1; i < existingData.length; i++) {
          const adName = String(existingData[i][0] || '').split('||')[0].trim();
          const existingLabel = existingData[i][8] || '';
          if (adName && existingLabel) {
            existingLabels[adName] = existingLabel;
          }
        }
        Logger.log('Preserved ' + Object.keys(existingLabels).length + ' existing labels');
      } catch (e) {
        Logger.log('Could not read existing labels: ' + e.toString());
      }
    }
    
    // Clear sheet
    dataSheet.clear();
    
    // Process rows
    const allRows = [];
    const targetColumns = ['Ad name', 'Spend', 'ROAS', 'CPA', 'Purchases', 'AOV', 'Hook rate', 'Hold rate', 'Product Label', 'CSV Source'];
    allRows.push(targetColumns);
    
    let totalAdsLoaded = 0;
    
    csvContents.forEach((csvObj, fileIndex) => {
      const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
      const csvFileName = typeof csvObj === 'string' ? `File_${fileIndex + 1}` : (csvObj.fileName || csvObj.name);
      const productLabel = csvObj.productLabel || '';
      
      Logger.log('Processing: ' + csvFileName);
      
      let csvData;
      if (csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length) {
        csvData = csv.split('\n').map(line => line.split(';'));
      } else {
        csvData = Utilities.parseCsv(csv);
      }
      
      if (csvData.length === 0) return;
      
      const headers = csvData[0];
      
      const findCol = (name) => {
        const searchName = name.toLowerCase().trim();
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] && headers[i].toString().toLowerCase().trim() === searchName) {
            return i;
          }
        }
        return -1;
      };
      
      const adNameCol = findCol('ad name');
      const spendCol = findCol('spend');
      const roasCol = findCol('roas');
      const cpaCol = findCol('cpa');
      const purchasesCol = findCol('purchases');
      const aovCol = findCol('aov');
      const hookRateCol = findCol('hook rate');
      const holdRateCol = findCol('hold rate');
      
      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        if (!row[adNameCol] || row[adNameCol].toString().trim() === '') continue;
        
        const adName = row[adNameCol] || '';
        const cleanAdName = adName.toString().split('||')[0].trim();
        
        // Check for existing label
        let label = productLabel || existingLabels[cleanAdName] || '';
        
        // Build output row
        const outputRow = [
          adName,
          row[spendCol] || '',
          row[roasCol] || '',
          row[cpaCol] || '',
          row[purchasesCol] || '',
          row[aovCol] || '',
          row[hookRateCol] || '',
          row[holdRateCol] || '',
          label,
          fileIndex + 1  // CSV Source
        ];
        
        allRows.push(outputRow);
        totalAdsLoaded++;
      }
    });
    
    // Write to sheet
    if (allRows.length > 1) {
      dataSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
      Logger.log('Wrote ' + (allRows.length - 1) + ' rows to CSV Data sheet');
      
      // Format header
      dataSheet.getRange(1, 1, 1, allRows[0].length).setFontWeight('bold').setBackground('#f3f4f6');
      dataSheet.setFrozenRows(1);
    }
    
    // ✅ GENERATE THE WEEKLY REPORT
    Logger.log('📊 Generating Weekly Report...');
    const reportResult = generateReport();
    Logger.log('✅ Weekly Report generated: ' + JSON.stringify(reportResult));
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Loaded ' + totalAdsLoaded + ' ads from ' + csvContents.length + ' file(s). Weekly Report generated!',
      '✅ CSV Upload Complete',
      5
    );
    
    return { success: true, totalAds: totalAdsLoaded, reportGenerated: true };
    
  } catch (e) {
    Logger.log('Error in processCSVDataDirect: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Filter out minority products from a CSV
 */
function filterMinorityProductsFromCsv(csvObj) {
  const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
  const fileName = typeof csvObj === 'string' ? 'Filtered' : (csvObj.fileName || csvObj.name);
  
  let csvData;
  const delimiter = csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length ? ';' : ',';
  
  if (delimiter === ';') {
    csvData = csv.split('\n').map(line => line.split(';'));
  } else {
    csvData = Utilities.parseCsv(csv);
  }
  
  if (csvData.length === 0) return csvObj;
  
  const headers = csvData[0];
  const adNameIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
  
  if (adNameIndex === -1) return csvObj;
  
  // Find majority product
  const productCounts = {};
  for (let i = 1; i < csvData.length; i++) {
    const adName = csvData[i][adNameIndex];
    if (!adName) continue;
    const product = extractProductFromAdName(adName.toString());
    if (product && product !== 'Unknown') {
      const normalized = normalizeProductName(product);
      productCounts[normalized] = (productCounts[normalized] || 0) + 1;
    }
  }
  
  const majorityProduct = Object.keys(productCounts).reduce((a, b) => 
    productCounts[a] > productCounts[b] ? a : b
  );
  
  // Filter rows
  const filteredRows = [headers];
  for (let i = 1; i < csvData.length; i++) {
    const adName = csvData[i][adNameIndex];
    if (!adName) continue;
    const product = extractProductFromAdName(adName.toString());
    const normalized = product ? normalizeProductName(product) : null;
    
    if (normalized === majorityProduct || !normalized) {
      filteredRows.push(csvData[i]);
    }
  }
  
  // Convert back to CSV string
  // ✅ PROPERLY ESCAPE CSV VALUES
  const escapeCsvValue = (val) => {
    if (val === null || val === undefined) return '';
    const str = val.toString();
    if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };
  
  const newCsv = filteredRows.map(row => row.map(escapeCsvValue).join(delimiter)).join('\n');
  
  Logger.log('  Filtered from ' + csvData.length + ' to ' + filteredRows.length + ' rows (keeping ' + majorityProduct + ')');
  
  return {
    fileName: fileName,
    csv: newCsv,
    content: newCsv
  };
}

/**
 * Split a CSV into separate CSVs by product
 */
function splitCsvByProduct(csvObj) {
  const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
  const fileName = typeof csvObj === 'string' ? 'Split' : (csvObj.fileName || csvObj.name);
  
  let csvData;
  const delimiter = csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length ? ';' : ',';
  
  if (delimiter === ';') {
    csvData = csv.split('\n').map(line => line.split(';'));
  } else {
    csvData = Utilities.parseCsv(csv);
  }
  
  if (csvData.length === 0) return [csvObj];
  
  const headers = csvData[0];
  const adNameIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
  
  if (adNameIndex === -1) return [csvObj];
  
  // Group rows by product
  const productRows = {};
  
  for (let i = 1; i < csvData.length; i++) {
    const adName = csvData[i][adNameIndex];
    if (!adName) continue;
    const product = extractProductFromAdName(adName.toString());
    const normalized = product ? normalizeProductName(product) : 'Unknown';
    
    if (!productRows[normalized]) {
      productRows[normalized] = [headers];
    }
    productRows[normalized].push(csvData[i]);
  }
  
  // Create separate CSVs
  const splitCsvs = [];
  Object.keys(productRows).forEach(product => {
    const rows = productRows[product];
    // ✅ PROPERLY ESCAPE CSV VALUES (handle commas, quotes, newlines in values)
    const escapeCsvValue = (val) => {
      if (val === null || val === undefined) return '';
      const str = val.toString();
      // If value contains delimiter, quote, or newline, wrap in quotes and escape internal quotes
      if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    };
    
    const newCsv = rows.map(row => row.map(escapeCsvValue).join(delimiter)).join('\n');
    const newFileName = fileName.replace('.csv', '') + '_' + product + '.csv';
    
    Logger.log('  Split: ' + product + ' → ' + (rows.length - 1) + ' rows');
    
    splitCsvs.push({
      fileName: newFileName,
      csv: newCsv,
      content: newCsv,
      splitFrom: fileName,
      splitProduct: product
    });
  });
  
  return splitCsvs;
}

/**
 * Apply per-product decisions to a CSV
 */
function applyPerProductDecisions(csvObj, perProductDecisions) {
  if (!perProductDecisions || Object.keys(perProductDecisions).length === 0) {
    return csvObj;
  }
  
  const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
  const fileName = typeof csvObj === 'string' ? 'Modified' : (csvObj.fileName || csvObj.name);
  
  let csvData;
  const delimiter = csv.includes(';') && csv.split('\n')[0].split(';').length > csv.split('\n')[0].split(',').length ? ';' : ',';
  
  if (delimiter === ';') {
    csvData = csv.split('\n').map(line => line.split(';'));
  } else {
    csvData = Utilities.parseCsv(csv);
  }
  
  if (csvData.length === 0) return csvObj;
  
  const headers = csvData[0];
  const adNameIndex = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'ad name');
  
  if (adNameIndex === -1) return csvObj;
  
  // Process each row based on decisions
  const processedRows = [headers];
  
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    const adName = row[adNameIndex];
    if (!adName) continue;
    
    const product = extractProductFromAdName(adName.toString());
    const normalized = product ? normalizeProductName(product) : null;
    
    if (!normalized) {
      processedRows.push(row);
      continue;
    }
    
    const decision = perProductDecisions[normalized];
    
    if (!decision || decision === 'keep') {
      processedRows.push(row);
    } else if (decision === 'ignore') {
      // Skip this row
      Logger.log('    Ignored: ' + adName.substring(0, 40));
    } else if (decision.startsWith('convert_')) {
      // Convert to another product by modifying ad name
      const targetProduct = decision.replace('convert_', '');
      const newAdName = adName.toString().replace(
        new RegExp(normalized + '_(FB|YT|GDN|TT)_', 'gi'),
        targetProduct + '_$1_'
      );
      row[adNameIndex] = newAdName;
      processedRows.push(row);
      Logger.log('    Converted: ' + normalized + ' → ' + targetProduct);
    }
  }
  
  // Convert back to CSV string
  const newCsv = processedRows.map(row => row.join(delimiter)).join('\n');
  
  Logger.log('  Per-product: ' + csvData.length + ' → ' + processedRows.length + ' rows');
  
  return {
    fileName: fileName,
    csv: newCsv,
    content: newCsv
  };
}

/**
 * Add a label to a CSV object
 */
function addLabelToCsv(csvObj, label) {
  const csv = typeof csvObj === 'string' ? csvObj : (csvObj.csv || csvObj.content);
  const fileName = typeof csvObj === 'string' ? 'Labeled' : (csvObj.fileName || csvObj.name);
  
  return {
    fileName: fileName,
    csv: csv,
    content: csv,
    productLabel: label
  };
}

/**
 * Combine two CSV files
 */
function combineCsvFiles(csvObj1, csvObj2) {
  const csv1 = typeof csvObj1 === 'string' ? csvObj1 : (csvObj1.csv || csvObj1.content);
  const csv2 = typeof csvObj2 === 'string' ? csvObj2 : (csvObj2.csv || csvObj2.content);
  const fileName1 = typeof csvObj1 === 'string' ? 'File1' : (csvObj1.fileName || csvObj1.name);
  const fileName2 = typeof csvObj2 === 'string' ? 'File2' : (csvObj2.fileName || csvObj2.name);
  
  // Parse both
  const delimiter = csv1.includes(';') && csv1.split('\n')[0].split(';').length > csv1.split('\n')[0].split(',').length ? ';' : ',';
  
  let csvData1, csvData2;
  if (delimiter === ';') {
    csvData1 = csv1.split('\n').map(line => line.split(';'));
    csvData2 = csv2.split('\n').map(line => line.split(';'));
  } else {
    csvData1 = Utilities.parseCsv(csv1);
    csvData2 = Utilities.parseCsv(csv2);
  }
  
  // Get headers from both
  const headers1 = csvData1[0];
  const headers2 = csvData2[0];
  
  // ✅ CRITICAL: Normalize headers to standard columns
  // Standard columns we care about: Ad name, Spend, ROAS, CPA, Purchases, AOV, Hook rate, Hold rate
  const standardCols = ['ad name', 'spend', 'roas', 'cpa', 'purchases', 'aov', 'hook rate', 'hold rate'];
  
  // Build standard headers (use original casing from first CSV that has it)
  const standardHeaders = standardCols.map(col => {
    for (let h of headers1) {
      if (h && h.toString().toLowerCase().trim() === col) return h;
    }
    for (let h of headers2) {
      if (h && h.toString().toLowerCase().trim() === col) return h;
    }
    return col.charAt(0).toUpperCase() + col.slice(1); // Fallback capitalized
  });
  
  // Helper to find column index in a headers array
  const findColIn = (headers, colName) => {
    const search = colName.toLowerCase().trim();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase().trim() === search) {
        return i;
      }
    }
    return -1;
  };
  
  // Build column mappings for each CSV
  const mapping1 = standardCols.map(col => findColIn(headers1, col));
  const mapping2 = standardCols.map(col => findColIn(headers2, col));
  
  Logger.log('  Headers1: ' + JSON.stringify(headers1));
  Logger.log('  Headers2: ' + JSON.stringify(headers2));
  Logger.log('  Mapping1: ' + JSON.stringify(mapping1));
  Logger.log('  Mapping2: ' + JSON.stringify(mapping2));
  
  // Extract and normalize data from CSV1
  const combinedRows = [standardHeaders];
  
  for (let i = 1; i < csvData1.length; i++) {
    const row = csvData1[i];
    const normalizedRow = mapping1.map(colIdx => colIdx >= 0 ? (row[colIdx] || '') : '');
    combinedRows.push(normalizedRow);
  }
  
  // Extract and normalize data from CSV2
  for (let i = 1; i < csvData2.length; i++) {
    const row = csvData2[i];
    const normalizedRow = mapping2.map(colIdx => colIdx >= 0 ? (row[colIdx] || '') : '');
    combinedRows.push(normalizedRow);
  }
  
  // Convert back to CSV string
  // ✅ PROPERLY ESCAPE CSV VALUES
  const escapeCsvValue = (val) => {
    if (val === null || val === undefined) return '';
    const str = val.toString();
    if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };
  
  const newCsv = combinedRows.map(row => row.map(escapeCsvValue).join(delimiter)).join('\n');

  const newFileName = fileName1.replace('.csv', '') + '_combined_' + fileName2.replace('.csv', '') + '.csv';
  
  Logger.log('  Combined: ' + (csvData1.length - 1) + ' + ' + (csvData2.length - 1) + ' = ' + (combinedRows.length - 1) + ' rows');
  Logger.log('  Standard headers: ' + standardHeaders.join(', '));
  
  return {
    fileName: newFileName,
    csv: newCsv,
    content: newCsv,
    combinedFrom: [fileName1, fileName2]
  };
}

/**
 * Updated processCSVData to use CSV Configurator for complex cases
 * Replace the existing mixed products detection section with this
 */
function shouldShowCsvConfigurator(csvContents) {
  const analysis = analyzeCSVFilesForConfigurator(csvContents);
  
  // Show configurator if:
  // 1. Any CSV has mixed products
  // 2. Any product appears in multiple CSVs (duplicates)
  // 3. More than 2 CSV files uploaded
  
  const hasMixed = analysis.csvAnalysis.some(csv => csv.isMixed);
  const hasDuplicates = analysis.csvAnalysis.some(csv => csv.hasDuplicates);
  const multipleFiles = csvContents.length > 2;
  
  return hasMixed || hasDuplicates || multipleFiles;
}

/**
 * 🔧 DIAGNOSTIC: Test if generateReport works
 * Run this from Apps Script Editor directly
 */
function testGenerateReport() {
  initializeConfigConstants_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  Logger.log('=== DIAGNOSTIC: testGenerateReport ===');
  
  if (!dataSheet) {
    Logger.log('❌ CSV Data sheet not found!');
    SpreadsheetApp.getUi().alert('CSV Data sheet not found!');
    return;
  }
  
  const lastRow = dataSheet.getLastRow();
  Logger.log('CSV Data sheet has ' + lastRow + ' rows');
  
  if (lastRow === 0) {
    Logger.log('❌ CSV Data sheet is empty!');
    SpreadsheetApp.getUi().alert('CSV Data sheet is empty! No data to generate report from.');
    return;
  }
  
  // Show first few rows
  const preview = dataSheet.getRange(1, 1, Math.min(5, lastRow), 10).getValues();
  Logger.log('Preview of CSV Data:');
  preview.forEach((row, i) => Logger.log('  Row ' + i + ': ' + row.slice(0, 3).join(' | ')));
  
  // Try to generate report
  Logger.log('Calling generateReport()...');
  try {
    const result = generateReport();
    Logger.log('generateReport() result: ' + JSON.stringify(result));
    
    if (result.success) {
      SpreadsheetApp.getUi().alert('✅ Success!\n\n' + result.message);
    } else {
      SpreadsheetApp.getUi().alert('❌ Failed:\n\n' + result.message);
    }
  } catch (e) {
    Logger.log('❌ Error: ' + e.toString());
    SpreadsheetApp.getUi().alert('❌ Error:\n\n' + e.toString());
  }
}

function diagnoseConfigAccess() {
  const CONFIG_SHEET_ID = '1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk';
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    Logger.log('✅ Can access Config sheet: ' + ss.getName());
    
    const sheets = ss.getSheets().map(s => s.getName());
    Logger.log('📋 Available tabs: ' + sheets.join(', '));
    
    const productFolders = ss.getSheetByName('Product Folders');
    if (productFolders) {
      Logger.log('✅ Product Folders tab exists with ' + productFolders.getLastRow() + ' rows');
    } else {
      Logger.log('❌ Product Folders tab NOT found!');
    }
    
  } catch (e) {
    Logger.log('❌ Cannot access Config sheet: ' + e.message);
  }
}

function testConfigConstants() {
  initializeConfigConstants_();
  Logger.log('PRODUCT_LOCALE_FOLDERS keys: ' + Object.keys(PRODUCT_LOCALE_FOLDERS));
  Logger.log('TARGET_CPA keys: ' + Object.keys(TARGET_CPA));
}

/**
 * Test Drive link lookup - run this to see what's happening
 */
function testDriveLinkLookup() {
  initializeConfigConstants_();
  
  Logger.log('=== DRIVE LINK DIAGNOSTIC ===');
  Logger.log('');
  
  // Test with a known product
  const testProducts = ['Guardality', 'CamTrix', 'EpiCooler'];
  
  testProducts.forEach(product => {
    Logger.log(`\n📦 Testing: ${product}`);
    
    // Check if product exists in PRODUCT_LOCALE_FOLDERS
    if (!PRODUCT_LOCALE_FOLDERS[product]) {
      Logger.log(`  ❌ Not found in PRODUCT_LOCALE_FOLDERS`);
      return;
    }
    
    const locales = Object.keys(PRODUCT_LOCALE_FOLDERS[product]);
    Logger.log(`  ✅ Found locales: ${locales.join(', ')}`);
    
    // Test first locale
    const testLocale = locales[0];
    const folderUrl = PRODUCT_LOCALE_FOLDERS[product][testLocale];
    Logger.log(`  📁 Testing locale ${testLocale}: ${folderUrl}`);
    
    const folderId = folderUrl.split('/folders/')[1];
    if (!folderId) {
      Logger.log(`  ❌ Could not extract folder ID from URL`);
      return;
    }
    
    try {
      const baseFolder = DriveApp.getFolderById(folderId);
      Logger.log(`  ✅ Folder access OK: ${baseFolder.getName()}`);
      
      // List subfolders
      const subfolders = baseFolder.getFolders();
      let subCount = 0;
      while (subfolders.hasNext() && subCount < 5) {
        const sub = subfolders.next();
        const subName = sub.getName();
        Logger.log(`    📂 Subfolder: "${subName}"`);
        
        // Test batch extraction
        const batchMatch1 = subName.match(/_(\d+)_?\(/);  // Original regex
        const batchMatch2 = subName.match(/_(\d+)/);       // More flexible
        const batchMatch3 = subName.match(/(\d+)/);        // Most flexible
        
        Logger.log(`       Original regex /_(\d+)_?\(/: ${batchMatch1 ? batchMatch1[1] : 'NO MATCH'}`);
        Logger.log(`       Flexible regex /_(\d+)/: ${batchMatch2 ? batchMatch2[1] : 'NO MATCH'}`);
        Logger.log(`       Any number /(\d+)/: ${batchMatch3 ? batchMatch3[1] : 'NO MATCH'}`);
        
        subCount++;
      }
      
      Logger.log(`  Total subfolders checked: ${subCount}`);
      
    } catch (e) {
      Logger.log(`  ❌ Error: ${e.message}`);
    }
  });
  
  Logger.log('\n=== DIAGNOSTIC COMPLETE ===');
  Logger.log('Check the logs above to see which regex works for your folder naming.');
}

/**
 * Run this function to diagnose ClickUp connection issues
 * Check View > Logs after running
 */
function diagnoseClickUpConnection() {
  Logger.log('========== CLICKUP CONNECTION DIAGNOSTIC ==========');
  
  // Step 1: Check if config loads
  Logger.log('\n--- STEP 1: Loading Config ---');
  let config;
  try {
    config = getClickUpConfigDataLazy_();
    Logger.log('✅ Config loaded successfully');
    Logger.log('Config keys: ' + Object.keys(config).join(', '));
  } catch (e) {
    Logger.log('❌ Failed to load config: ' + e.message);
    return;
  }
  
  // Step 2: Check each config value
  Logger.log('\n--- STEP 2: Config Values ---');
  Logger.log('SPACE_ID: ' + (config.SPACE_ID || '❌ MISSING'));
  Logger.log('FB_ADS_FOLDER_ID: ' + (config.FB_ADS_FOLDER_ID || '❌ MISSING'));
  Logger.log('SPRINT_FOLDER_ID: ' + (config.SPRINT_FOLDER_ID || '❌ MISSING'));
  Logger.log('LEGACY_SPRINT_FOLDER_ID: ' + (config.LEGACY_SPRINT_FOLDER_ID || '(not set)'));
  Logger.log('API_KEY: ' + (config.API_KEY ? config.API_KEY.substring(0, 10) + '...' : '❌ MISSING'));
  Logger.log('TEAM_ID: ' + (config.TEAM_ID || '❌ MISSING'));
  
  // Step 3: Check getter functions
  Logger.log('\n--- STEP 3: Getter Function Results ---');
  try {
    Logger.log('getClickUpSpaceId_(): ' + getClickUpSpaceId_());
  } catch (e) {
    Logger.log('❌ getClickUpSpaceId_() failed: ' + e.message);
  }
  
  try {
    Logger.log('getClickUpFolderId_(): ' + getClickUpFolderId_());
  } catch (e) {
    Logger.log('❌ getClickUpFolderId_() failed: ' + e.message);
  }
  
  try {
    const apiKey = getClickUpApiKey_();
    Logger.log('getClickUpApiKey_(): ' + (apiKey ? apiKey.substring(0, 10) + '...' : '❌ EMPTY'));
  } catch (e) {
    Logger.log('❌ getClickUpApiKey_() failed: ' + e.message);
  }
  
  // Step 4: Test API connection
  Logger.log('\n--- STEP 4: Testing API Connection ---');
  const spaceId = getClickUpSpaceId_();
  const apiKey = getClickUpApiKey_();
  
  if (!spaceId) {
    Logger.log('❌ Cannot test API - SPACE_ID is empty');
    return;
  }
  
  if (!apiKey) {
    Logger.log('❌ Cannot test API - API_KEY is empty');
    return;
  }
  
  const url = 'https://api.clickup.com/api/v2/space/' + spaceId + '/folder?archived=false';
  Logger.log('API URL: ' + url);
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    Logger.log('Response Code: ' + responseCode);
    
    if (responseCode === 200) {
      Logger.log('✅ API connection successful!');
      const data = JSON.parse(response.getContentText());
      Logger.log('Number of folders found: ' + data.folders.length);
      
      // List all folders
      Logger.log('\n--- Available Folders ---');
      data.folders.forEach(function(folder) {
        Logger.log('  ID: ' + folder.id + ' | Name: ' + folder.name);
      });
      
      // Check if our target folder exists
      const targetFolderId = getClickUpFolderId_();
      Logger.log('\n--- Checking Target Folder ---');
      Logger.log('Looking for folder ID: ' + targetFolderId);
      
      const targetFolder = data.folders.find(function(f) {
        return f.id == targetFolderId;
      });
      
      if (targetFolder) {
        Logger.log('✅ Target folder found: ' + targetFolder.name);
        Logger.log('Lists in folder: ' + (targetFolder.lists ? targetFolder.lists.length : 0));
      } else {
        Logger.log('❌ Target folder NOT FOUND in this space!');
        Logger.log('Available folder IDs: ' + data.folders.map(function(f) { return f.id; }).join(', '));
      }
      
    } else if (responseCode === 401) {
      Logger.log('❌ Authentication failed - API key is invalid');
      Logger.log('Response: ' + response.getContentText());
    } else if (responseCode === 404) {
      Logger.log('❌ Space not found - SPACE_ID may be wrong');
      Logger.log('Response: ' + response.getContentText());
    } else {
      Logger.log('❌ Unexpected error');
      Logger.log('Response: ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ API call failed: ' + e.message);
  }
  
  Logger.log('\n========== END DIAGNOSTIC ==========');
}

function testApiKeyDirect() {
  // Paste your API key here directly to test
  const apiKey = 'pk_100652070_MF1MRV4VX1ETPN5B192QG8BJS3RETZH6';
  
  const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': apiKey },
    muteHttpExceptions: true
  });
  
  Logger.log('Response: ' + response.getResponseCode());
  Logger.log('Body: ' + response.getContentText());
}

function debugGetClickUpSprints() {
  Logger.log('========== DEBUG getClickUpSprints ==========');
  
  // Step 1: Get config values
  const spaceId = getClickUpSpaceId_();
  const folderId = getClickUpFolderId_();
  const apiKey = getClickUpApiKey_();
  
  Logger.log('Space ID: ' + spaceId);
  Logger.log('Folder ID: ' + folderId);
  Logger.log('API Key (trimmed): ' + apiKey.substring(0, 15) + '...');
  
  // Step 2: Make API call
  const url = 'https://api.clickup.com/api/v2/space/' + spaceId + '/folder?archived=false';
  Logger.log('URL: ' + url);
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('Response code: ' + response.getResponseCode());
  
  if (response.getResponseCode() !== 200) {
    Logger.log('Response body: ' + response.getContentText());
    return;
  }
  
  const data = JSON.parse(response.getContentText());
  Logger.log('Total folders in space: ' + data.folders.length);
  
  // Step 3: List all folders
  Logger.log('\n--- All Folders in Space ---');
  data.folders.forEach(function(f) {
    Logger.log('  ' + f.id + ' = ' + f.name + ' (lists: ' + (f.lists ? f.lists.length : 0) + ')');
  });
  
  // Step 4: Check if our folder ID matches
  Logger.log('\n--- Looking for folder ID: ' + folderId + ' ---');
  const matchingFolder = data.folders.find(function(f) {
    return f.id == folderId || String(f.id) === String(folderId);
  });
  
  if (matchingFolder) {
    Logger.log('✅ FOUND: ' + matchingFolder.name);
  } else {
    Logger.log('❌ NOT FOUND - folder ID does not exist in this space!');
  }
  
  // Step 5: Now test actual getClickUpSprints
  Logger.log('\n--- Testing getClickUpSprints() ---');
  try {
    const sprints = getClickUpSprints();
    Logger.log('Sprints returned: ' + sprints.length);
    if (sprints.length > 0) {
      Logger.log('First 3 sprints:');
      sprints.slice(0, 3).forEach(function(s) {
        Logger.log('  ' + s.id + ' = ' + s.name);
      });
    }
  } catch (e) {
    Logger.log('❌ getClickUpSprints() threw error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
  
  Logger.log('========== END DEBUG ==========');
}

function debugApiKeyBytes() {
  // First, clear the cache
  CacheService.getScriptCache().remove('atria_clickup');
  Logger.log('Cache cleared');
  
  // Force fresh load
  _clickupConfigData = null;  // Reset lazy loader
  
  const config = getClickUpConfigFromConfig_();  // Direct load, no cache
  const apiKey = config.API_KEY;
  
  Logger.log('Raw API key: [' + apiKey + ']');
  Logger.log('Length: ' + apiKey.length);
  Logger.log('Expected: 44');
  
  // Show character codes
  let codes = [];
  for (let i = 0; i < apiKey.length; i++) {
    codes.push(apiKey.charCodeAt(i));
  }
  Logger.log('Char codes: ' + codes.join(','));
  
  // Test with the key we got
  const response1 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': apiKey },
    muteHttpExceptions: true
  });
  Logger.log('Config key response: ' + response1.getResponseCode());
  
  // Test with hardcoded
  const hardcoded = 'pk_100652070_MF1MRV4VX1ETPN5B192QG8BJS3RETZH6';
  const response2 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': hardcoded },
    muteHttpExceptions: true
  });
  Logger.log('Hardcoded key response: ' + response2.getResponseCode());
  
  // Are they equal?
  Logger.log('Keys match: ' + (apiKey === hardcoded));
  Logger.log('Trimmed match: ' + (apiKey.trim() === hardcoded));
}

function debugGetClickUpSprints() {
  Logger.log('========== DEBUG getClickUpSprints ==========');
  
  // Clear cache first!
  CacheService.getScriptCache().remove('atria_clickup');
  _clickupConfigData = null;
  
  const spaceId = getClickUpSpaceId_();
  const folderId = getClickUpFolderId_();
  const apiKey = getClickUpApiKey_();
  
  Logger.log('Space ID: ' + spaceId);
  Logger.log('Folder ID: ' + folderId);
  Logger.log('API Key: ' + apiKey.substring(0, 15) + '...');
  
  const url = 'https://api.clickup.com/api/v2/space/' + spaceId + '/folder?archived=false';
  
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  Logger.log('Response code: ' + response.getResponseCode());
  
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    Logger.log('✅ Found ' + data.folders.length + ' folders');
    
    // Now test actual function
    const sprints = getClickUpSprints();
    Logger.log('getClickUpSprints() returned: ' + sprints.length + ' sprints');
  } else {
    Logger.log('Response: ' + response.getContentText());
  }
}

function debugApiKeyDirect() {
  // Clear everything
  CacheService.getScriptCache().remove('atria_clickup');
  
  // Read directly from sheet - NO cache, NO lazy loader
  const ss = SpreadsheetApp.openById('1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk');  // Fixed typo
  const sheet = ss.getSheetByName('ClickUp Config');
  const data = sheet.getDataRange().getValues();
  
  let apiKeyFromSheet = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'API_KEY') {
      apiKeyFromSheet = data[i][1];
      break;
    }
  }
  
  Logger.log('API key from sheet: [' + apiKeyFromSheet + ']');
  Logger.log('Length: ' + apiKeyFromSheet.length);
  Logger.log('Typeof: ' + typeof apiKeyFromSheet);
  
  // Test this key
  const response1 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': apiKeyFromSheet },
    muteHttpExceptions: true
  });
  Logger.log('Direct from sheet: ' + response1.getResponseCode());
  
  // Now test what getClickUpApiKey_() returns
  const apiKeyFromGetter = getClickUpApiKey_();
  Logger.log('From getter: [' + apiKeyFromGetter + ']');
  Logger.log('Getter length: ' + apiKeyFromGetter.length);
  
  const response2 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': apiKeyFromGetter },
    muteHttpExceptions: true
  });
  Logger.log('From getter: ' + response2.getResponseCode());
  
  // Are they the same?
  Logger.log('Same value: ' + (apiKeyFromSheet === apiKeyFromGetter));
}

function debugGetClickUpSprintsStepByStep() {
  Logger.log('========== STEP BY STEP DEBUG ==========');
  
  // Clear cache
  CacheService.getScriptCache().remove('atria_clickup');
  _clickupConfigData = null;
  
  // Step 1: Get values
  const spaceId = getClickUpSpaceId_();
  const apiKey = getClickUpApiKey_();
  
  Logger.log('Step 1 - spaceId: ' + spaceId);
  Logger.log('Step 1 - apiKey: ' + apiKey.substring(0,15) + '...');
  
  // Step 2: Build URL exactly as getClickUpSprints does
  const url = `https://api.clickup.com/api/v2/space/${spaceId}/folder?archived=false`;
  Logger.log('Step 2 - URL: ' + url);
  
  // Step 3: Build options exactly as getClickUpSprints does
  const options = {
    method: 'get',
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  Logger.log('Step 3 - Auth header: ' + options.headers.Authorization.substring(0,15) + '...');
  
  // Step 4: Make the call
  const response = UrlFetchApp.fetch(url, options);
  Logger.log('Step 4 - Response code: ' + response.getResponseCode());
  
  if (response.getResponseCode() !== 200) {
    Logger.log('Step 4 - Error body: ' + response.getContentText());
  } else {
    Logger.log('Step 4 - SUCCESS!');
    const data = JSON.parse(response.getContentText());
    Logger.log('Folders found: ' + data.folders.length);
  }
  
  // Step 5: Now call actual getClickUpSprints and see what happens
  Logger.log('\n--- Now calling actual getClickUpSprints() ---');
  try {
    const result = getClickUpSprints();
    Logger.log('Result length: ' + result.length);
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}

function testClickUpAfterFix() {
  // Force complete reset
  CacheService.getScriptCache().remove('atria_clickup');
  _clickupConfigData = null;
  
  Logger.log('=== Testing ClickUp Connection ===');
  
  // Now get fresh config
  const config = getClickUpConfigFromConfig_();
  Logger.log('Config keys: ' + Object.keys(config).join(', '));
  Logger.log('API_KEY length: ' + (config.API_KEY || '').length);
  Logger.log('API_KEY: ' + (config.API_KEY || '').substring(0, 15) + '...');
  
  // Test API
  const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': config.API_KEY },
    muteHttpExceptions: true
  });
  
  Logger.log('API Response: ' + response.getResponseCode());
  
  if (response.getResponseCode() === 200) {
    Logger.log('✅ API KEY WORKS!');
    
    // Now test getClickUpSprints
    const sprints = getClickUpSprints();
    Logger.log('Sprints found: ' + sprints.length);
  } else {
    Logger.log('❌ API still failing: ' + response.getContentText());
  }
}

function testClickUpFinal() {
  // Complete reset
  clearAtriaConfigCache();
  
  Logger.log('=== FINAL CLICKUP TEST ===');
  
  // Test 1: Direct config load
  const config = getClickUpConfigFromConfig_();
  Logger.log('1. Config loaded: ' + Object.keys(config).length + ' keys');
  Logger.log('   API_KEY: ' + (config.API_KEY ? config.API_KEY.substring(0,15) + '...' : 'MISSING'));
  
  // Test 2: Lazy loader
  const lazyConfig = getClickUpConfigDataLazy_();
  Logger.log('2. Lazy config: ' + Object.keys(lazyConfig).length + ' keys');
  Logger.log('   API_KEY: ' + (lazyConfig.API_KEY ? lazyConfig.API_KEY.substring(0,15) + '...' : 'MISSING'));
  
  // Test 3: API key getter
  const apiKey = getClickUpApiKey_();
  Logger.log('3. getClickUpApiKey_(): ' + (apiKey ? apiKey.substring(0,15) + '...' : 'EMPTY'));
  
  // Test 4: Direct API call
  const response = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': apiKey },
    muteHttpExceptions: true
  });
  Logger.log('4. API test: ' + response.getResponseCode());
  
  // Test 5: getClickUpSprints
  Logger.log('5. Calling getClickUpSprints()...');
  const sprints = getClickUpSprints();
  Logger.log('   Result: ' + sprints.length + ' sprints');
  
  if (sprints.length > 0) {
    Logger.log('   First 3: ' + sprints.slice(0,3).map(s => s.name).join(', '));
  }
  
  Logger.log('=== TEST COMPLETE ===');
}

function testCorrectSpaceId() {
  clearAtriaConfigCache();
  
  const config = getClickUpConfigFromConfig_();
  
  // Test with the CORRECT space ID (FB_ADS_FOLDER_ID)
  const correctSpaceId = config.FB_ADS_FOLDER_ID;  // Should be 90150821271
  const wrongSpaceId = config.SPACE_ID;  // This is 90151192866
  
  Logger.log('Testing with FB_ADS_FOLDER_ID: ' + correctSpaceId);
  
  const response = UrlFetchApp.fetch(
    'https://api.clickup.com/api/v2/space/' + correctSpaceId + '/folder?archived=false',
    {
      headers: { 'Authorization': config.API_KEY },
      muteHttpExceptions: true
    }
  );
  
  Logger.log('Response: ' + response.getResponseCode());
  
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    Logger.log('✅ SUCCESS! Found ' + data.folders.length + ' folders');
    data.folders.forEach(f => Logger.log('  - ' + f.id + ': ' + f.name));
  } else {
    Logger.log('❌ Failed: ' + response.getContentText());
  }
}

function testSprintsFinal() {
  clearAtriaConfigCache();
  
  Logger.log('=== TESTING SPRINTS ===');
  
  const sprints = getClickUpSprints();
  
  Logger.log('Found ' + sprints.length + ' sprints');
  
  if (sprints.length > 0) {
    Logger.log('First 5 sprints:');
    sprints.slice(0, 5).forEach(s => {
      Logger.log('  - ' + s.name + ' (ID: ' + s.id + ')');
    });
  }
  
  Logger.log('=== TEST COMPLETE ===');
}

function debugClickUpStepByStep() {
  Logger.log('========== COMPREHENSIVE CLICKUP DEBUG ==========');
  
  // Step 0: Show current state of lazy loader
  Logger.log('Step 0: _clickupConfigData = ' + (_clickupConfigData ? 'SET' : 'NULL'));
  
  // Step 1: Clear everything
  Logger.log('\nStep 1: Clearing all caches...');
  CacheService.getScriptCache().remove('atria_clickup');
  _clickupConfigData = null;
  Logger.log('  Cache cleared, lazy loader reset');
  
  // Step 2: Load config directly (bypassing lazy loader)
  Logger.log('\nStep 2: Loading config directly from sheet...');
  const ss = SpreadsheetApp.openById('1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk');
  const sheet = ss.getSheetByName('ClickUp Config');
  const data = sheet.getDataRange().getValues();
  
  const directConfig = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    let value = data[i][1];
    if (typeof value === 'string') value = value.trim();
    if (key) directConfig[key] = value;
  }
  
  Logger.log('  Direct config keys: ' + Object.keys(directConfig).join(', '));
  Logger.log('  API_KEY from sheet: [' + (directConfig.API_KEY || 'MISSING') + ']');
  Logger.log('  API_KEY length: ' + (directConfig.API_KEY || '').length);
  
  // Step 3: Test API with direct config
  Logger.log('\nStep 3: Testing API with direct config...');
  const testResponse1 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': directConfig.API_KEY },
    muteHttpExceptions: true
  });
  Logger.log('  Direct API test: ' + testResponse1.getResponseCode());
  
  // Step 4: Now call getClickUpConfigFromConfig_()
  Logger.log('\nStep 4: Calling getClickUpConfigFromConfig_()...');
  const configFromFunction = getClickUpConfigFromConfig_();
  Logger.log('  Function returned keys: ' + Object.keys(configFromFunction).join(', '));
  Logger.log('  API_KEY from function: [' + (configFromFunction.API_KEY || 'MISSING') + ']');
  
  // Step 5: Test API with function config
  Logger.log('\nStep 5: Testing API with function config...');
  const testResponse2 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': configFromFunction.API_KEY },
    muteHttpExceptions: true
  });
  Logger.log('  Function config API test: ' + testResponse2.getResponseCode());
  
  // Step 6: Check lazy loader state
  Logger.log('\nStep 6: Checking _clickupConfigData...');
  Logger.log('  _clickupConfigData = ' + (_clickupConfigData ? 'SET' : 'NULL'));
  if (_clickupConfigData) {
    Logger.log('  _clickupConfigData.API_KEY: [' + (_clickupConfigData.API_KEY || 'MISSING') + ']');
  }
  
  // Step 7: Call getClickUpApiKey_()
  Logger.log('\nStep 7: Calling getClickUpApiKey_()...');
  const apiKeyFromGetter = getClickUpApiKey_();
  Logger.log('  getClickUpApiKey_() returned: [' + apiKeyFromGetter + ']');
  Logger.log('  Length: ' + apiKeyFromGetter.length);
  
  // Step 8: Test API with getter
  Logger.log('\nStep 8: Testing API with getClickUpApiKey_()...');
  const testResponse3 = UrlFetchApp.fetch('https://api.clickup.com/api/v2/user', {
    headers: { 'Authorization': apiKeyFromGetter },
    muteHttpExceptions: true
  });
  Logger.log('  Getter API test: ' + testResponse3.getResponseCode());
  
  // Step 9: Now test the actual space call
  Logger.log('\nStep 9: Testing space API call...');
  const spaceId = getClickUpSpaceId_();
  Logger.log('  Space ID: ' + spaceId);
  
  const spaceUrl = 'https://api.clickup.com/api/v2/space/' + spaceId + '/folder?archived=false';
  Logger.log('  URL: ' + spaceUrl);
  
  const spaceResponse = UrlFetchApp.fetch(spaceUrl, {
    headers: { 
      'Authorization': apiKeyFromGetter,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  Logger.log('  Space API response: ' + spaceResponse.getResponseCode());
  
  if (spaceResponse.getResponseCode() !== 200) {
    Logger.log('  Error: ' + spaceResponse.getContentText());
  } else {
    const spaceData = JSON.parse(spaceResponse.getContentText());
    Logger.log('  SUCCESS! Found ' + spaceData.folders.length + ' folders');
  }
  
  // Step 10: Now call getClickUpSprints()
  Logger.log('\nStep 10: Calling getClickUpSprints()...');
  const sprints = getClickUpSprints();
  Logger.log('  Result: ' + sprints.length + ' sprints');
  
  Logger.log('\n========== DEBUG COMPLETE ==========');
}

function quickSpaceIdTest() {
  _clickupConfigData = null;
  CacheService.getScriptCache().remove('atria_clickup');
  
  const spaceId = getClickUpSpaceId_();
  Logger.log('Space ID: ' + spaceId);
  Logger.log('Expected: 90150821271');
  Logger.log('Match: ' + (spaceId === '90150821271'));
  
  if (spaceId === '90150821271') {
    const sprints = getClickUpSprints();
    Logger.log('Sprints: ' + sprints.length);
  }
}

function quickSpaceIdTest2() {
  _clickupConfigData = null;
  CacheService.getScriptCache().remove('atria_clickup');
  
  const spaceId = getClickUpSpaceId_();
  Logger.log('Space ID: ' + spaceId);
  Logger.log('Type: ' + typeof spaceId);
  Logger.log('Expected: 90150821271');
  Logger.log('Match: ' + (String(spaceId) === '90150821271'));
  
  // Test sprints
  Logger.log('\nTesting getClickUpSprints()...');
  const sprints = getClickUpSprints();
  Logger.log('Sprints found: ' + sprints.length);
  
  if (sprints.length > 0) {
    Logger.log('First 5:');
    sprints.slice(0, 5).forEach(s => Logger.log('  - ' + s.name));
  }
}

function debugCreativeTestsProgress() {
  clearAtriaConfigCache();
  
  Logger.log('=== DEBUG CREATIVE TESTS PROGRESS ===');
  
  // Check config values
  const config = getClickUpConfigFromConfig_();
  Logger.log('\n1. Config Values:');
  Logger.log('   FB_ADS_FOLDER_ID: ' + config.FB_ADS_FOLDER_ID + ' (used as Space ID)');
  Logger.log('   FB_ADS_LISTS_FOLDER_ID: ' + config.FB_ADS_LISTS_FOLDER_ID + ' (used for FB Ads folder)');
  Logger.log('   SPACE_ID: ' + config.SPACE_ID);
  
  // What getFbAdsFolderId_ returns
  Logger.log('\n2. getFbAdsFolderId_() returns: ' + getFbAdsFolderId_());
  Logger.log('   Expected (from working code): 90151192866');
  
  // Test the API call
  const spaceId = getClickUpSpaceId_();
  Logger.log('\n3. getClickUpSpaceId_() returns: ' + spaceId);
  
  const apiKey = getClickUpApiKey_();
  const folderUrl = 'https://api.clickup.com/api/v2/space/' + spaceId + '/folder?archived=false';
  
  Logger.log('\n4. Fetching folders from: ' + folderUrl);
  
  const response = UrlFetchApp.fetch(folderUrl, {
    headers: { 'Authorization': apiKey },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log('   ERROR: ' + response.getContentText());
    return;
  }
  
  const data = JSON.parse(response.getContentText());
  Logger.log('   Found ' + data.folders.length + ' folders');
  
  // List all folders
  Logger.log('\n5. All folders in space:');
  data.folders.forEach(f => {
    Logger.log('   - ' + f.id + ': ' + f.name + ' (' + (f.lists ? f.lists.length : 0) + ' lists)');
  });
  
  // Find FB Ads folder
  const fbAdsFolderId = getFbAdsFolderId_();
  Logger.log('\n6. Looking for FB Ads folder with ID: ' + fbAdsFolderId);
  
  const fbAdsFolder = data.folders.find(f => f.id == fbAdsFolderId || String(f.id) === String(fbAdsFolderId));
  
  if (fbAdsFolder) {
    Logger.log('   ✅ FOUND: ' + fbAdsFolder.name + ' with ' + fbAdsFolder.lists.length + ' product lists');
    Logger.log('\n7. Product lists in FB Ads folder:');
    fbAdsFolder.lists.slice(0, 10).forEach(l => {
      Logger.log('   - ' + l.name);
    });
  } else {
    Logger.log('   ❌ NOT FOUND!');
    Logger.log('   Available folder IDs: ' + data.folders.map(f => f.id).join(', '));
  }
}

function testNormalizeProductName() {
  const testNames = [
    'A35_Hydrogen Bottle',
    'A03_Mini Wifi Camera', 
    'A10_TV Antena',
    'AliveBlue',
    'CamTrix',
    'WaveMax'
  ];
  
  testNames.forEach(name => {
    const normalized = normalizeProductName(name);
    Logger.log(name + ' → ' + normalized);
  });
}

function debugAliveBlueCreativeTests() {
  clearAtriaConfigCache();
  
  Logger.log('=== DEBUG ALIVEBLUE CREATIVE TESTS ===');
  
  // Get sprints
  const sprints = getClickUpSprints();
  Logger.log('Found ' + sprints.length + ' sprints');
  
  // Find Week 52 (2025)
  const week52 = sprints.find(s => s.name.includes('Week 52') || s.name.includes('52 (2025)'));
  
  if (!week52) {
    Logger.log('ERROR: Could not find Week 52 (2025)');
    Logger.log('Available sprints:');
    sprints.forEach(s => Logger.log('  - ' + s.name));
    return;
  }
  
  Logger.log('Using sprint: ' + week52.name + ' (ID: ' + week52.id + ')');
  
  // Step 1: Check what batches exist for AliveBlue in this sprint
  Logger.log('\n--- Step 1: Get ClickUp tasks for AliveBlue ---');
  const sprintTasks = getClickUpSprintTasks(week52.id, 'AliveBlue');
  Logger.log('AliveBlue tasks found: ' + sprintTasks.length);
  
  if (sprintTasks.length > 0) {
    Logger.log('AliveBlue batch IDs:');
    sprintTasks.forEach(t => Logger.log('  - ' + t.batchId));
  }
  
  // Step 2: Check CSV Data for AliveBlue ads
  Logger.log('\n--- Step 2: Check CSV Data for AliveBlue ads ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('CSV Data');
  
  if (!dataSheet) {
    Logger.log('ERROR: No CSV Data sheet found!');
    return;
  }
  
  const data = dataSheet.getDataRange().getValues();
  const aliveBlueAds = [];
  
  for (let i = 1; i < data.length; i++) {
    const adName = String(data[i][0] || '');
    if (adName.toLowerCase().includes('aliveblue')) {
      aliveBlueAds.push(adName.split('||')[0].trim());
    }
  }
  
  Logger.log('AliveBlue ads in CSV: ' + aliveBlueAds.length);
  if (aliveBlueAds.length > 0) {
    Logger.log('First 10 AliveBlue ads:');
    aliveBlueAds.slice(0, 10).forEach(a => Logger.log('  - ' + a.substring(0, 70)));
  }
  
  // Step 3: Run analyzeCreativeTests
  Logger.log('\n--- Step 3: Run analyzeCreativeTests ---');
  const result = analyzeCreativeTests(week52.id, 'AliveBlue');
  
  Logger.log('Result success: ' + result.success);
  Logger.log('Result message: ' + (result.message || 'none'));
  Logger.log('Results length: ' + (result.results ? result.results.length : 0));
  
  if (result.results && result.results.length > 0) {
    Logger.log('Batches found:');
    result.results.forEach(r => Logger.log('  - ' + r.batchId + ' (' + (r.ads ? r.ads.length : 0) + ' ads)'));
  }
  
  Logger.log('\n=== DEBUG COMPLETE ===');
}

function debugProductSpreadsheetMappings() {
  clearAtriaConfigCache();
  
  Logger.log('=== DEBUG PRODUCT SPREADSHEET MAPPINGS ===');
  
  const mappings = getProductSpreadsheetIds();
  
  Logger.log('\nAll mappings (' + Object.keys(mappings).length + ' total):');
  Object.keys(mappings).sort().forEach(key => {
    Logger.log('  "' + key + '" → ' + mappings[key].substring(0, 20) + '...');
  });
  
  // Check for AuraNaturals and KatuChef specifically
  Logger.log('\n--- Checking AuraNaturals ---');
  Logger.log('  "AuraNaturals" exists: ' + ('AuraNaturals' in mappings));
  Logger.log('  "auranaturals" exists: ' + ('auranaturals' in mappings));
  Logger.log('  "AURANATURALS" exists: ' + ('AURANATURALS' in mappings));
  
  Logger.log('\n--- Checking KatuChef ---');
  Logger.log('  "KatuChef" exists: ' + ('KatuChef' in mappings));
  Logger.log('  "katuchef" exists: ' + ('katuchef' in mappings));
  Logger.log('  "KATUCHEF" exists: ' + ('KATUCHEF' in mappings));
  
  // Now check what product names come from the Creative Tests data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName('Creative Tests');
  
  if (testSheet) {
    const data = testSheet.getDataRange().getValues();
    const foundProducts = new Set();
    
    for (let i = 0; i < data.length; i++) {
      const cellValue = String(data[i][1] || ''); // Column B
      
      // Check if it's a product header (starts with emoji flag)
      if (cellValue.startsWith('🟡') || cellValue.startsWith('🔴') || cellValue.startsWith('🟢')) {
        // Skip - these are batch headers
      } else if (cellValue.match(/^[A-Z]+$/i) || cellValue.includes('NATURALS') || cellValue.includes('CHEF')) {
        // This might be a product section header
        if (data[i][0] && String(data[i][0]).includes('📦')) {
          foundProducts.add(cellValue);
        }
      }
    }
    
    // Also scan for product headers in column A
    for (let i = 0; i < data.length; i++) {
      const cellA = String(data[i][0] || '');
      if (cellA.startsWith('📦')) {
        const productName = cellA.replace('📦', '').trim();
        foundProducts.add(productName);
      }
    }
    
    Logger.log('\n--- Products found in Creative Tests sheet ---');
    foundProducts.forEach(p => {
      const inMappings = p in mappings || p.toLowerCase() in Object.keys(mappings).reduce((acc, k) => { acc[k.toLowerCase()] = true; return acc; }, {});
      Logger.log('  "' + p + '" - in mappings: ' + inMappings);
    });
  }
  
  Logger.log('\n=== DEBUG COMPLETE ===');
}

function testHookHoldFormatting() {
  // Test the decimal to percentage conversion
  const testValues = [0.30, 0.5025, 0.0074, 1.0];
  
  Logger.log('=== Hook/Hold Formatting Test ===');
  testValues.forEach(val => {
    const formatted = (val * 100).toFixed(2) + '%';
    Logger.log(val + ' → ' + formatted);
  });
  
  // Expected output:
  // 0.30 → 30.00%
  // 0.5025 → 50.25%
  // 0.0074 → 0.74%
  // 1.0 → 100.00%
}