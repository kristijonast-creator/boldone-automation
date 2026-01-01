// ================================================================================
// CENTRAL CONFIG CONNECTION
// ================================================================================
const CONFIG_SHEET_ID = '1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk';
const CONFIG_CACHE_DURATION = 300;

function getAdminEmailsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_admins');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const admins = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const isAdmin = data[i][3];
    const active = data[i][7];
    if (email && isAdmin === true && active === true) {
      admins.push(email);
    }
  }
  
  cache.put('adn_admins', JSON.stringify(admins), CONFIG_CACHE_DURATION);
  return admins;
}

function getClickUpAllowedUsersFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_clickup_users');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const clickupAccess = data[i][4];
    const active = data[i][7];
    if (email && clickupAccess === true && active === true) {
      users.push(email);
    }
  }
  
  cache.put('adn_clickup_users', JSON.stringify(users), CONFIG_CACHE_DURATION);
  return users;
}

function getAndromedaUsersFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_andromeda_users');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const andromedaAccess = data[i][5];
    const active = data[i][7];
    if (email && andromedaAccess === true && active === true) {
      users.push(email);
    }
  }
  
  cache.put('adn_andromeda_users', JSON.stringify(users), CONFIG_CACHE_DURATION);
  return users;
}

function getLocalesFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_locales');
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
  
  cache.put('adn_locales', JSON.stringify(locales), CONFIG_CACHE_DURATION);
  return locales;
}

function getProductIndexFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_product_index');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  
  const index = {};
  for (let i = 1; i < data.length; i++) {
    const productKey = data[i][0];
    const productId = data[i][2];
    const productFullName = data[i][3];
    const productUrl = data[i][5];
    
    if (productKey) {
      index[productKey] = {
        id: productId || '',
        name: productFullName || '',
        url: productUrl || ''
      };
    }
  }
  
  cache.put('adn_product_index', JSON.stringify(index), CONFIG_CACHE_DURATION);
  return index;
}

function getProductLocaleFoldersFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_folders');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Product Folders');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const folders = {};
  for (let i = 1; i < data.length; i++) {
    const productKey = data[i][0];
    if (!productKey) continue;
    
    folders[productKey] = {};
    for (let j = 1; j < headers.length; j++) {
      const locale = headers[j];
      const folderId = data[i][j];
      if (folderId && folderId.toString().trim() !== '') {
        folders[productKey][locale] = 'https://drive.google.com/drive/folders/' + folderId.toString().trim();
      } else {
        folders[productKey][locale] = '';
      }
    }
  }
  
  cache.put('adn_folders', JSON.stringify(folders), CONFIG_CACHE_DURATION);
  return folders;
}

/**
 * Get template document IDs from central Config sheet
 * Returns: { IMG: 'docId', VID: 'docId', ITR: 'docId', LOC: 'docId', YTTOFB: 'docId', ANDROMEDA_VID: 'docId' }
 */
function getTemplatesFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_templates');
  if (cached) return JSON.parse(cached);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Templates');
    
    if (!sheet) {
      Logger.log('⚠️ Templates tab not found in central config - using hardcoded values');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    const templates = {};
    for (let i = 1; i < data.length; i++) {
      const templateKey = data[i][0];
      const documentId = data[i][1];
      
      if (templateKey && documentId && documentId.toString().trim() !== '') {
        templates[templateKey.toString().trim()] = documentId.toString().trim();
      }
    }
    
    if (Object.keys(templates).length > 0) {
      cache.put('adn_templates', JSON.stringify(templates), CONFIG_CACHE_DURATION);
      Logger.log('✅ Loaded ' + Object.keys(templates).length + ' templates from central Config');
    }
    
    return templates;
  } catch (e) {
    Logger.log('⚠️ Error loading templates from central Config: ' + e.message);
    return null;
  }
}

/**
 * Get footage folder URLs from central Config sheet
 * Returns: { "ProductKey": "https://drive.google.com/...", ... }
 */
function getFootageFoldersFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_footage');
  if (cached) return JSON.parse(cached);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Footage Folders');
    
    if (!sheet) {
      Logger.log('⚠️ Footage Folders tab not found in central config - using hardcoded values');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    const folders = {};
    for (let i = 1; i < data.length; i++) {
      const productKey = data[i][0];
      const folderUrl = data[i][1];
      
      if (productKey) {
        // Store even empty values so we know the product exists
        folders[productKey.toString().trim()] = (folderUrl || '').toString().trim();
      }
    }
    
    if (Object.keys(folders).length > 0) {
      cache.put('adn_footage', JSON.stringify(folders), CONFIG_CACHE_DURATION);
      Logger.log('✅ Loaded ' + Object.keys(folders).length + ' footage folder mappings from central Config');
    }
    
    return folders;
  } catch (e) {
    Logger.log('⚠️ Error loading footage folders from central Config: ' + e.message);
    return null;
  }
}

function clearAdNamingConfigCache() {
  CacheService.getScriptCache().removeAll([
    'adn_admins', 
    'adn_clickup_users', 
    'adn_andromeda_users', 
    'adn_locales', 
    'adn_product_index', 
    'adn_folders', 
    'adn_team_members', 
    'adn_allowed_sheets',
    'adn_templates',    // NEW
    'adn_footage'       // NEW
  ]);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Config cache cleared!', 'Done', 3);
}

function getTeamMembersFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_team_members');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][1];
    const active = data[i][7];
    if (email && active === true) {
      members.push(email);
    }
  }
  
  cache.put('adn_team_members', JSON.stringify(members), CONFIG_CACHE_DURATION);
  return members;
}

function getClickUpApiKeyFromConfig_() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return null;
    
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName('Team Members');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userEmail && data[i][7] === true) {
        return data[i][2] || null; // Column C = ClickUp API Key
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function getAllowedSheetsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('adn_allowed_sheets');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Team Members');
  const data = sheet.getDataRange().getValues();
  
  const sheets = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];  // Column A = Name
    const active = data[i][7]; // Column H = Active
    if (name && active === true) {
      sheets.push(name);
    }
  }
  
  cache.put('adn_allowed_sheets', JSON.stringify(sheets), CONFIG_CACHE_DURATION);
  return sheets;
}
// ================================================================================

// ===== ANALYTICS CONFIGURATION =====
const ANALYTICS_ENABLED = true;  // Set to false to disable all tracking
const ANALYTICS_SHEET_ID = '1KalmueNyF_2Hiy52NmkJblQrsbFK4hDhD5my-Psj-L0';  // ✅ Paste the ID from Step 1

const ADMIN_EMAILS = getAdminEmailsFromConfig_();
const TRACK_ADMINS = true;

const CLICKUP_ALLOWED_USERS = getClickUpAllowedUsersFromConfig_();
const CLICKUP_AUTOFILL_USERS = getClickUpAllowedUsersFromConfig_(); // Same as ClickUp access
const ANDROMEDA_USERS = getAndromedaUsersFromConfig_();
// ===================================

function onOpen() {
  try {
    normalizeProductLocaleFolders();
    
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Ad Notes')
      .addItem('Create Ad Notes + Open', 'createAdNotesAndOpen')
      .addItem('Create Ad Notes + Choose Folder', 'createAdNotesWithPicker')
      .addItem('Create Ad Notes + Paste Link', 'createAdNotesChooseFolderPrompt')
      .addItem('Create Ad Notes + Auto Route', 'createAdNotesAutoRoute')
      .addSeparator()
      .addItem('🔄 Refresh all Batch IDs', 'refreshAllBatchIds');
    
    // ✅ Safe email access for admin check
  try {
    const currentUser = Session.getActiveUser().getEmail();
    if (ADMIN_EMAILS.includes(currentUser)) {
      menu.addSeparator()
          .addItem('📊 View Analytics Dashboard', 'showAnalyticsDashboard')
          .addItem('📥 Export Analytics Data', 'exportAnalyticsData')
          .addItem('🔄 Recalculate Stats from Sheet', 'recalculateAnalyticsStats')
          .addItem('👥 Check Team Analytics Status', 'generateTeamAnalyticsReport')
          .addItem('📋 Complete Team Report', 'generateCompleteTeamReport');
    }
  } catch (emailError) {
    // User hasn't granted email permission - skip admin menu
    Logger.log('⚠️ Could not check admin status: ' + emailError);
  }
    
    menu.addToUi();
  } catch (e) {
    Logger.log('❌ onOpen failed: ' + e);
  }
}

// Add this constant at the top of your script (after ADMIN_EMAILS)
const TEAM_MEMBERS = getTeamMembersFromConfig_();

// ✅ ClickUp API Configuration
const CLICKUP_API_KEY = getClickUpApiKeyFromConfig_();  // Get from: https://app.clickup.com/settings/apps
const CLICKUP_ENABLED = true;  // Set to true after adding your API key

// Template file IDs - check central Config first, fall back to hardcoded
const TEMPLATES_CONFIG = getTemplatesFromConfig_() || {};
const IMG_TEMPLATE_ID = TEMPLATES_CONFIG['IMG'] || '1zmyCJLrNLxOgsRperFRH319Ctaj8LPne3xwuYgPoI1w';
const VID_TEMPLATE_ID = TEMPLATES_CONFIG['VID'] || '1RKXPnmirIelXitg8HVRuD8r3afDloLRqqjoD6y6AoXc';
const ITR_TEMPLATE_ID = TEMPLATES_CONFIG['ITR'] || '1LKlyX1j0sNVJYYH3WdGTip4ZdyTJ92Ob5tnwzfZlth8';
const LOC_TEMPLATE_ID = TEMPLATES_CONFIG['LOC'] || '15-YSsBz_8gDY9s2rTp02ry23ZzraYdJcGdMsy7oZY-4';
const YTTOFB_TEMPLATE_ID = TEMPLATES_CONFIG['YTTOFB'] || '1nA8iyXEjcG38O46EF87zSW-fyZWrQNtw48ec-yK3VMQ';
const ANDROMEDA_VID_TEMPLATE_ID = TEMPLATES_CONFIG['ANDROMEDA_VID'] || '1LPY644EsF-bLjUwrt4oy_pZXbR6P-04CAy-yh-bAWnk';

// Optional default parent folder
const PARENT_FOLDER_ID = '';
const CREATE_SUBFOLDER = false;

// Fallback per-locale (optional)
const LOCALE_BASE_FOLDERS = {
  US: '', UK: '', DE: '', CA: '', AU: '', LATAM: '', ES: '',
  FR: '', IT: '', IL: '', EU: ''
};

// Org sharing
const ORG_DOMAIN = 'commercecore.com';

// All supported locales (order doesn’t matter)
const LOCALES = getLocalesFromConfig_();

// Picker API key, same Cloud project as this Apps Script
const PICKER_DEVELOPER_KEY = 'AIzaSyARSwrKz-9ipuZMJyPhFs9iFACyMcFAbX8';

// PRODUCT INDEX
const PRODUCT_INDEX = getProductIndexFromConfig_();

// === Product → Locale base folders ===
// Key = product domain (from start of C31). Added [Ax] tags in comments.
const PRODUCT_LOCALE_FOLDERS = getProductLocaleFoldersFromConfig_();


// FOOTAGE LINKS - check central Config first, fall back to hardcoded
const FOOTAGE_CONFIG = getFootageFoldersFromConfig_();
const FOLDER_LINKS = FOOTAGE_CONFIG || {
  // Fallback hardcoded values (only used if Config tab doesn't exist)
  EpiCooler: 'https://drive.google.com/drive/folders/15qj9M0m11ShfnvgKyaoUtC4upC6Vtimd',
  Ozoori: 'https://drive.google.com/drive/folders/1pxF-P2YMeJsqOoJCe0UhexqhgCpq6TDF',
  Heatoor: 'https://drive.google.com/drive/folders/12RpUYCeRXGN3DKAqfMw-KyItm5U2BEBq',
  WiggyDog: 'https://drive.google.com/drive/folders/1AET_w4i4OPxpsHjGFn7qDqA0oQuVXmdd',
  PrimaFocus: '',
  OmniHear: 'https://drive.google.com/drive/folders/1mT7jpjXaAhjCtKOE0wc0h4RziILuM5sX',
  Clairu_M2: 'https://drive.google.com/drive/u/0/folders/1glDIEoPCGEF13UR1tjBj4hPE7ZzFzhnk',
  CamTrix: 'https://drive.google.com/drive/folders/1jCrAaVyvV1zURcFExthu8jccoDiUs0Mk',
  HammerDex: '',
  WellHeater: 'https://drive.google.com/drive/folders/1jn6GeQgHNqOG3Un32wYPhqFDfnlsR8h6',
  WaveMax: 'https://drive.google.com/drive/folders/1wW14AT70gm2-EVK7XH0153KMxSDFlFeK',
  LumenLight: '',
  SleepZee: '',
  Repellio: '',
  Clairu: 'https://drive.google.com/drive/folders/1glDIEoPCGEF13UR1tjBj4hPE7ZzFzhnk',
  Blitron: '',
  Phoxfer: '',
  LoweSkin: '',
  OriBreeze: 'https://drive.google.com/drive/folders/1CYp1kWaHpCMVG19xXiiSFwgpdQC_RBRG',
  Zapfie: '',
  NuraFix: '',
  Cleanlix: '',
  LoveAndShine: '',
  JayShoes: '',
  NuraDerma: '',
  Epibella: '',
  CleanlixPowder: 'https://drive.google.com/drive/folders/1g6mE8O_3SEQIqqM64bc74a2FFryPzT2_',
  CleanlixMold: 'https://drive.google.com/drive/folders/1B5rsMz-wdmtOFoIQ3JwsQs93mKOviBPS',
  VacuumSealer: '',
  DermaSonic: '',
  Guardality: 'https://drive.google.com/drive/folders/1eyEqM2iSOJMnNygTaqjIIt9U8OYpTOzv',
  TVStick: '',
  Flixy: '',
  TellyStick: 'https://drive.google.com/drive/folders/1U0YKbwrwdZeTCNmfrWooKPQcomfSkOZV',
  TheraWolfRelief: 'https://drive.google.com/drive/folders/1EK3o966nv4WTsH8gQahuzxCZkgWotuHw',
  TheraWolfNeuro: 'https://drive.google.com/drive/folders/1kSYm3aIi_4K55JAVGKCNCoBjddmHvKYq',
  BioVitals: '',
  AuraNaturals: 'https://drive.google.com/drive/folders/1B8YmG7BWb7oUPVAasD4mr63nCetbLdlt',
  PlanetWash: '',
  EmaWell: '',
  BRYT: '',
  KatuChef: 'https://drive.google.com/drive/folders/1gTcSZlOAUC84OEpitVyOIE49bCu7D-RV',
  AliveBlue: 'https://drive.google.com/drive/folders/11VZ3TIR9u_R3wCcJXaYLTYq0MAv08zH6'
};

function debugDiagnostics() {
  try {
    const res = getCreateDialogDiagnosticsPlus();
    Logger.log(JSON.stringify(res, null, 2));
  } catch (e) {
    Logger.log("💥 " + e + "\n" + (e.stack || ""));
    throw e;
  }
}

function sanityCheck() {
  Logger.log("✅ Apps Script runtime working");
  return "ok";
}

/* ========= Buttons ========= */

function createAdNotesWithPicker() {
  const html = HtmlService.createHtmlOutput(buildPickerHtml_Debug())
    .setWidth(640)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select target folder');
}

function detectProductStage() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = sheet.getRange('C31').getValue();
  const product = name.split('_')[0];
  const id = getProductId(product);
  return { id, name: product };
}

function detectLocaleStage(productName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const val = sheet.getRange('C31').getValue();
  const locale = extractLocale(val);
  return locale || '—';
}

function detectPlatformStage(productName, locale) {
  const val = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('C31').getValue();
  const platform = extractPlatform(val);
  return platform || '—';
}

function createAdNotesChooseFolderPrompt() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = ui.prompt('Target folder', 'Paste a Drive FOLDER link or ID. You can also paste a FILE link and I will use its parent folder.', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const raw = (res.getResponseText() || '').trim();
    if (!raw) { ui.alert('Empty input. Paste a Drive folder link or ID.'); return; }

    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);

    const url = createAdNotesCore(raw);

    lock.releaseLock();

    const html = '<script>window.open("' + url + '","_blank");google.script.host.close();</script>';
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(10).setHeight(10), 'Opening…');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Failed: ' + (e && e.message ? e.message : e));
  }
}

function createAdNotesAndOpen() {
  const url = createAdNotesCore(null);
  const html = '<script>window.open("' + url + '","_blank");google.script.host.close();</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(10).setHeight(10), 'Opening…');
}

function openCreateDialog() {
  // 1) build the base HTML (your existing popup)
  var htmlText = buildCreateDialogHtml('{}');

  // 2) inject the ClickUp copy script just before </body>
  htmlText = injectRenderDoneBoxScript(htmlText);
  htmlText = injectReopenLastCreatedScript(htmlText);

  // 3) serve it
  var html = HtmlService.createHtmlOutput(htmlText)
    .setWidth(680)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Create');
}


/** 
 * Customizes the Done box depending on which action created content.
 * Called from inside the HTML dialog via google.script.run or inline script.
 */
function injectRenderDoneBoxScript(html) {
  var script =
    '<script>' +
    'function renderDoneBoxByAction(type, urls){' +
      'urls = urls || {};' +
      
      // ✅ DEBUG: Log what we received
      'console.log("🔍 renderDoneBoxByAction called");' +
      'console.log("   type:", type);' +
      'console.log("   urls:", JSON.stringify(urls));' +
      'console.log("   urls.deleted:", urls.deleted);' +
      'console.log("   urls.deleted === true:", urls.deleted === true);' +
      
      // ✅ NEW: Check if this record was deleted (only if explicitly marked)
      'if(urls.deleted === true){' +
        'console.log("⚠️ This record was deleted");' +
        'var doneModal=document.getElementById("doneModal");' +
        'var isModal=doneModal && doneModal.style.display==="flex";' +
        'var suffix=isModal?"_modal":"";' +
        'var doneBox=document.getElementById("doneBox"+suffix);' +
        'if(doneBox){' +
          'doneBox.innerHTML="";' +
          'var title=document.createElement("div");' +
          'title.className="done-title";' +
          'title.textContent="🗑️ This batch was deleted";' +
          'doneBox.appendChild(title);' +
          'var msg=document.createElement("p");' +
          'msg.style.cssText="text-align:center;color:#6b7280;font-size:14px;margin:20px;line-height:1.6;";' +
          'msg.textContent="The batch you created was deleted by an admin on "+new Date(urls.deletedAt||Date.now()).toLocaleString()+". The folders and documents are in Google Drive Trash and can be restored if needed.";' +
          'doneBox.appendChild(msg);' +
          'doneBox.classList.add("show");' +
        '}' +
        'return;' +
      '}' +

      // ✅ NEW: Reset admin delete button state before rendering new content
      'var adminSection=document.getElementById("adminDeleteSection");' +
      'var adminDeleteBtn=document.getElementById("adminDeleteBtn");' +
      'if(adminSection) adminSection.style.display="none";' +
      'if(adminDeleteBtn){' +
        'adminDeleteBtn.disabled=false;' +
        'adminDeleteBtn.textContent="🗑️ Delete Test Batch (Admin Only)";' +
        'adminDeleteBtn.style.background="";' +
        'adminDeleteBtn.style.borderColor="";' +
        'adminDeleteBtn.style.color="";' +
      '}' +
      
      // ✅ NEW: Remove any "deleted successfully" messages from previous runs
      'var doneBox=document.getElementById("doneBox");' +
      'if(doneBox){' +
        'var oldMsgs=doneBox.querySelectorAll("p");' +
        'for(var i=0;i<oldMsgs.length;i++){' +
          'if(oldMsgs[i].textContent.includes("You can create")){' +
            'oldMsgs[i].remove();' +
          '}' +
        '}' +
      '}' +
      
      // Detect if we're in modal or inline mode
      'var doneModal=document.getElementById("doneModal");' +
      'var isModal = doneModal && doneModal.style.display === "flex";' +
      
      // Use modal suffix if in modal, otherwise no suffix
      'var suffix = isModal ? "_modal" : "";' +
      
      'var doneBox = document.getElementById("doneBox");' +
      'var title = doneBox ? doneBox.querySelector(".done-title") : null;' +
      'var docRow = document.getElementById("docRow" + suffix);' +
      'var adsRow = document.getElementById("adsRow" + suffix);' +
      'var openDocBtn = document.getElementById("openDocBtn" + suffix);' +
      'var openAdsBtn = document.getElementById("openAdsBtn" + suffix);' +
      'var copyDocBtn = document.getElementById("copyDocBtn" + suffix);' +
      'var copyAdsBtn = document.getElementById("copyAdsBtn" + suffix);' +
      'var copyClickupBtn = document.getElementById("copyClickupBtn" + suffix);' +
      'var copyStatus = document.getElementById("copyStatus" + suffix);' +

      // Reset visibility
      'if(docRow) docRow.style.display="none";' +
      'if(adsRow) adsRow.style.display="none";' +
      'if(copyStatus) copyStatus.style.display="none";' +

      // Resolve URLs - normalize the structure
      'var docUrl = urls.docUrl || "";' +
      'var adsUrl = urls.adsUrl || "";' +
      'var adNotesUrl = urls.adNotesUrl || "";' +
      
      'console.log("🎨 Rendering done box:", {type: type, docUrl: docUrl, adsUrl: adsUrl, isModal: isModal});' +

      // Handle different creation types
'if(type==="notes"){' +
  'if(title) title.textContent="✅ Ad Notes document created!";' +
  'if(docRow && docUrl) {' +
    'docRow.style.display="flex";' +
    'if(openDocBtn) openDocBtn.href = docUrl;' +
    'if(copyDocBtn) {' +
      'copyDocBtn.onclick = function(){' +
        'navigator.clipboard.writeText(docUrl);' +
        'copyDocBtn.classList.add("copied");' +
        'setTimeout(function(){ copyDocBtn.classList.remove("copied"); }, 1000);' +
      '};' +
    '}' +
  '}' +
     '}else if(type==="folder"){' +
  'if(title) title.textContent="✅ Folder created successfully!";' +
  'if(adsRow && adsUrl) {' +
    'adsRow.style.display="flex";' +
    'if(openAdsBtn) openAdsBtn.href = adsUrl;' +
    'if(copyAdsBtn) {' +
      'copyAdsBtn.onclick = function(){' +
        'navigator.clipboard.writeText(adsUrl);' +
        'copyAdsBtn.classList.add("copied");' +
        'setTimeout(function(){ copyAdsBtn.classList.remove("copied"); }, 1000);' +
      '};' +
    '}' +
  '}' +
'}else{' +
        // standard or custom
        'if(title) title.textContent="✅ Done! Your Ad Notes were created.";' +
  'if(docRow && docUrl) {' +
    'docRow.style.display="flex";' +
    'if(openDocBtn) openDocBtn.href = docUrl;' +
    'if(copyDocBtn) {' +
      'copyDocBtn.onclick = function(){' +
        'navigator.clipboard.writeText(docUrl);' +
        'copyDocBtn.classList.add("copied");' +
        'setTimeout(function(){ copyDocBtn.classList.remove("copied"); }, 1000);' +
      '};' +
    '}' +
  '}' +
  'if(adsRow && adsUrl) {' +
    'adsRow.style.display="flex";' +
    'if(openAdsBtn) openAdsBtn.href = adsUrl;' +
    'if(copyAdsBtn) {' +
      'copyAdsBtn.onclick = function(){' +
        'navigator.clipboard.writeText(adsUrl);' +
        'copyAdsBtn.classList.add("copied");' +
        'setTimeout(function(){ copyAdsBtn.classList.remove("copied"); }, 1000);' +
      '};' +
    '}' +
  '}' +
'}' +

// ✅ NEW: Setup admin delete button
      'console.log("🔍 Checking if user is admin...");' +
'google.script.run' +
  '.withSuccessHandler(function(isAdmin){' +
    'if(isAdmin && (urls.adsUrl || urls.docUrl)){' +
      'var adminSection=document.getElementById("adminDeleteSection");' +
      'var adminDeleteBtn=document.getElementById("adminDeleteBtn");' +
      'if(adminSection) adminSection.style.display="block";' +
      'if(adminDeleteBtn){' +
        
        // ✅ NEW: Handle both folder deletion (Standard/Custom) and doc deletion (Notes Only)
        'var targetType = urls.adsUrl ? "folder" : "document";' +
        'var targetUrl = urls.adsUrl || urls.docUrl;' +
        'var match = targetUrl.match(/\\/(folders|d)\\/([a-zA-Z0-9_-]+)/);' +
        'var targetId = match ? match[2] : null;' +
        
        'if(targetId){' +
          'adminDeleteBtn.onclick=function(){' +
            'var confirmMsg = targetType === "folder" ?' +
              '"⚠️ DELETE TEST BATCH?\\n\\nThis will move the entire batch folder to trash.\\n\\nYou can restore it from Google Drive trash if needed.\\n\\nContinue?" :' +
              '"⚠️ DELETE TEST DOCUMENT?\\n\\nThis will move the Ad Notes document to trash.\\n\\nYou can restore it from Google Drive trash if needed.\\n\\nContinue?";' +
            'if(!confirm(confirmMsg)) return;' +
            'adminDeleteBtn.disabled=true;' +
            'adminDeleteBtn.textContent="⏳ Deleting...";' +
            
            // ✅ Call appropriate delete function based on type
            'var deleteFunc = targetType === "folder" ? "deleteBatchFolderForDialog" : "deleteDocumentForDialog";' +
            
            'google.script.run' +
              '.withSuccessHandler(function(result){' +
                'if(result.status==="deleted"){' +
                  'var docRow=document.getElementById("docRow");' +
                  'var adsRow=document.getElementById("adsRow");' +
                  'if(docRow) docRow.style.display="none";' +
                  'if(adsRow) adsRow.style.display="none";' +
                  'var clickupSection=document.querySelector(".clickup-section");' +
                  'if(clickupSection) clickupSection.style.display="none";' +
                  'var clickupRenameBox=document.getElementById("clickupRenameBox");' +
                  'if(clickupRenameBox) {clickupRenameBox.classList.remove("show");clickupRenameBox.style.display="none";}' +
                  'var clickupAutoFillBox=document.getElementById("clickupAutoFillBox");' +
                  'if(clickupAutoFillBox) clickupAutoFillBox.style.display="none";' +
                  'var doneTitle=document.querySelector(".done-title");' +
                  'if(doneTitle) doneTitle.textContent= targetType === "folder" ? "🗑️ Test batch deleted successfully!" : "🗑️ Test document deleted successfully!";' +
                  'var adminSection=document.getElementById("adminDeleteSection");' +
                  'if(adminSection) adminSection.style.display="none";' +
                  'var doneBox=document.getElementById("doneBox");' +
                  'if(doneBox){' +
                    'var msg=document.createElement("p");' +
                    'msg.style.cssText="text-align:center;color:#6b7280;font-size:13px;margin:16px 0 0 0;";' +
                    'msg.textContent="You can create a new " + (targetType === "folder" ? "batch" : "document") + " now using the buttons above.";' +
                    'doneBox.appendChild(msg);' +
                  '}' +
                  'google.script.run.markLastCreatedAsDeleted();' +
                '}else{' +
                  'alert("Error: "+result.message);' +
                  'adminDeleteBtn.disabled=false;' +
                  'adminDeleteBtn.textContent="🗑️ Delete Test Batch";' +
                '}' +
              '})' +
              '.withFailureHandler(function(err){' +
                'alert("Error deleting: "+err.message);' +
                'adminDeleteBtn.disabled=false;' +
                'adminDeleteBtn.textContent="🗑️ Delete Test Batch";' +
              '})' +
              '[deleteFunc](targetId);' +
          '};' +
        '}' +
      '}' +
    '}' +
  '})' +
        '.withFailureHandler(function(err){' +
          'console.log("⚠️ Could not check admin status:",err);' +
        '})' +
        '.checkIfCurrentUserIsAdmin();' +
      
      // Setup ClickUp copy button
      'console.log("🎨 Building ClickUp copy text...");' +

      // Setup ClickUp copy button
      'console.log("🎨 Building ClickUp copy text...");' +
      'console.log("   urls object:", JSON.stringify(urls));' +
      
            // ✅ IMPROVED: Build as plain text list (no auto-linking)
      'var lines = [];' +
      'if(docUrl) {' +
        'lines.push("Ad Notes:");' +
        'lines.push(docUrl.trim());' +
        'lines.push("");' + // Empty line for spacing
      '}' +
      'if(adsUrl) {' +
        'lines.push("Ads folder:");' +
        'lines.push(adsUrl.trim());' +
        'lines.push("");' +
      '}' +
      'if(urls.originalAdUrl && urls.originalAdName) {' +
        'lines.push("Batch / Ad:");' +
        'lines.push(urls.originalAdUrl.trim());' +
      '}' +
      
      // ✅ Join with line breaks
      'var copyText = lines.join("\\n");' +
      
      'console.log("   Final copyText:", copyText);' +
      
      'if(copyClickupBtn && copyText){' +
        'copyClickupBtn.onclick = function(){' +
          'navigator.clipboard.writeText(copyText).then(function(){' +
            'if(copyStatus){ copyStatus.style.display="block"; setTimeout(function(){ copyStatus.style.display="none"; },1500); }' +
          '});' +
        '};' +
      '}' +
    '}' +
    '</script>';
  
  return html + script;
}


function injectReopenLastCreatedScript(html) {
  var script =
    '<script>' +
    '(function(){' +
      'document.addEventListener("click", function(e){' +
        'var btn = e.target.closest("#reopenLastBtn");' +
        'if(!btn) return;' +
        'btn.disabled = true;' +
        'btn.textContent = "⟳ Loading...";' +
        'google.script.run' +
          '.withSuccessHandler(function(record){' +
            'console.log("🧠 getLastCreatedRecord returned:", record);' +
            'btn.disabled = false;' +
            'btn.textContent = "⟳ Reopen Last Created";' +
            
            'if(!record || !record.urls){ ' +
              'alert("No recent creation found."); ' +
              'return; ' +
            '}' +
            
            // ✅ Debug: Log the urls object
            'console.log("📦 record.urls:", JSON.stringify(record.urls));' +
            'console.log("   - docUrl:", record.urls.docUrl);' +
            'console.log("   - adsUrl:", record.urls.adsUrl);' +
            'console.log("   - originalAdUrl:", record.urls.originalAdUrl);' +
            'console.log("   - originalAdName:", record.urls.originalAdName);' +
            
            // Show the modal
            'var modal = document.getElementById("doneModal");' +
            'if(!modal){ alert("Modal not found."); return; }' +
            'modal.style.display = "flex";' +
            
            // ✅ NEW: Normalize the urls object before passing
            'var normalizedUrls = {' +
              'docUrl: record.urls.docUrl || "",'+
              'adsUrl: record.urls.adsUrl || "",'+
              'adNotesUrl: record.urls.adNotesUrl || "",'+
              'originalAdUrl: record.urls.originalAdUrl || "",'+  // ✅ ADD THIS
              'originalAdName: record.urls.originalAdName || ""'+ // ✅ ADD THIS
            '};' +
            
            'console.log("📦 normalizedUrls passed to render:", JSON.stringify(normalizedUrls));' +
            
            // Render the content with normalized urls
            'renderDoneBoxByAction(record.type, normalizedUrls);' +
          '})' +
          '.withFailureHandler(function(err){' +
            'console.error("❌ Failed to get last record:", err);' +
            'btn.disabled = false;' +
            'btn.textContent = "⟳ Reopen Last Created";' +
            'alert("Failed to retrieve last created record: " + err.message);' +
          '})' +
          '.getLastCreatedRecord();' +
      '});' +
    '})();' +
    '</script>';

  var pos = html.lastIndexOf('</body>');
  if (pos === -1) return html + script;
  return html.slice(0, pos) + script + html.slice(pos);
}

/**
 * Check if current user can use ClickUp Auto-Fill
 * @returns {boolean}
 */
function checkClickUpAutoFillPermission() {
  try {
    const currentUser = Session.getActiveUser().getEmail();
    return CLICKUP_AUTOFILL_USERS.includes(currentUser);
  } catch (e) {
    Logger.log('⚠️ Could not check Auto-Fill permission: ' + e);
    return false;
  }
}

/**
 * Check if Andromeda template selection should be shown
 * @returns {Object} {hasPermission: boolean, isNewVid: boolean}
 */
function checkAndromedaFeature() {
  try {
    const hasPermission = checkAndromedaPermission();
    
    const sh = SpreadsheetApp.getActiveSheet();
    const c31 = (sh.getRange('C31').getDisplayValue() || '').trim().toUpperCase();
    
    // Check if it's a New VID ad
    const isNewVid = c31.includes('NEW') && (c31.includes('VID') || c31.includes('VIDEO'));
    
    // ✅ ADD LOGGING
    Logger.log('🔍 Andromeda Check:');
    Logger.log('   C31 value: ' + c31);
    Logger.log('   Has permission: ' + hasPermission);
    Logger.log('   Contains NEW: ' + c31.includes('NEW'));
    Logger.log('   Contains VID: ' + c31.includes('VID'));
    Logger.log('   Contains VIDEO: ' + c31.includes('VIDEO'));
    Logger.log('   Final isNewVid: ' + isNewVid);
    
    return {
      hasPermission: hasPermission,
      isNewVid: isNewVid
    };
  } catch (e) {
    Logger.log('⚠️ checkAndromedaFeature error: ' + e);
    return { hasPermission: false, isNewVid: false };
  }
}

/**
 * Check if current user can access Andromeda template
 * @returns {boolean}
 */
function checkAndromedaPermission() {
  try {
    const currentUser = Session.getActiveUser().getEmail();
    return ANDROMEDA_USERS.includes(currentUser);
  } catch (e) {
    Logger.log('⚠️ Could not check Andromeda permission: ' + e);
    return false;
  }
}

function createAdNotesChooseFolderPromptForDialog(populateAdName /*, forceCreate unused */) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    const res = ui.prompt(
      'Target folder',
      'Paste a Drive FOLDER link/ID (or a FILE link/ID — I will use its parent).',
      ui.ButtonSet.OK_CANCEL
    );
    if (res.getSelectedButton() !== ui.Button.OK) return { status: 'cancelled' };

    const raw = (res.getResponseText() || '').trim();
    if (!raw) return { status: 'cancelled' };

    const targetFolder = resolveFolderStrict(raw);
    const url = createAdNotesCoreOpt(targetFolder.getId(), populateAdName === true);

    // Also scaffold product→locale→year/<C31>/Ad Notes (report both URLs)
    let adNotesUrl = '';
    let adsUrl = '';
    try {
      const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
      const fullName = sanitizeBatchName(fullNameRaw).clean;
      const locale = parseLocaleFromC31(fullName);
      const domain = parseDomainFromC31(fullName);
      const year = detectTargetYear(fullName);

      const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
      if (baseFolder) {
        const yearFolder = getOrCreateSubfolder(baseFolder, year);
        const batchFolder = getOrCreateSubfolder(yearFolder, fullName);
        const adNotes = getOrCreateSubfolder(batchFolder, 'Ad Notes');
        adNotesUrl = toDriveUrl_(adNotes);
        adsUrl = toDriveUrl_(batchFolder);
      }
    } catch (_) { }

    return { status: 'created', url, adNotesUrl, adsUrl };
  } catch (e) {
    logErrorToHub_('Folder Creation', e.message || String(e));  // ← ADD THIS
    const msg = String(e && e.message ? e.message : e);
    const quota = /Service invoked too many times|Rate Limit|User rate limit/i.test(msg);
    const perm = /insufficient permissions|file not found|no access/i.test(msg);
    if (quota) return { status: 'error', message: 'Drive is throttling requests (rate limit). Please try again shortly.' };
    if (perm) return { status: 'error', message: 'You might not have access to that folder or the template. Please request access and retry.' };
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

/** Return a short “A33 CleanlixMold” label if known; else just the domain. */
function productShortLabel_(domain) {
  const meta = PRODUCT_INDEX[domain];
  if (meta && meta.id) return meta.id + ' ' + (domain || '');
  return domain || '';
}

/** Return subfolder if exists (no creation). */
function getSubfolderIfExists(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function createAdNotesUsingExistingBatchForDialog(batchFolderId, populateAdName, origAdMode) {
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(20 * 1000); } catch (_) { }

  try {
    const batchFolder = DriveApp.getFolderById(batchFolderId);
    const adNotesFolder = getOrCreateSubfolder(batchFolder, 'Ad Notes');
    const url = createAdNotesCoreOpt(adNotesFolder.getId(), populateAdName === true, { origAdMode: origAdMode || 'auto' });
    return {
      status: 'created',
      url,
      adNotesUrl: toDriveUrl_(adNotesFolder),
      adsUrl: toDriveUrl_(batchFolder)
    };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    
    // ✅ NEW: Log the error
    try {
      const sh = SpreadsheetApp.getActiveSheet();
      const c31 = sh.getRange('C31').getDisplayValue().trim();
      const domain = parseDomainFromC31(c31);
      
      logAnalyticsEventSafe_({
        action: 'Use Existing Batch',
        product: domain || '',
        success: false,
        error: msg.substring(0, 200)
      });
    } catch (_) {
      // Silent fail
    }
    
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

function attemptStandardSetupWithBatchCheck(populateAdName, origAdMode, videoTemplate) {
  const sh = SpreadsheetApp.getActiveSheet();
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
    if (!fullNameRaw) return { status: 'error', message: 'C31 is empty.' };
    const fullName = sanitizeBatchName(fullNameRaw).clean;

    const parts = parseNameParts(fullName);
    const domain = parseDomainFromC31(fullName);
    const locale = parts.locale;
    const batchId = parts.batch;
    const year = detectTargetYear(fullName);

    // 👇 Add this line for debugging
    Logger.log('Resolved product=' + domain + ', locale=' + locale + ', batchId=' + batchId);

    if (!domain || !locale || !batchId) {
      return { status: 'error', message: 'Could not fully parse product / locale / batch ID from C31. Make sure naming is correct.' };
    }

    const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
    if (!baseFolder) {
      return {
        status: 'error',
        message: '🟡 Missing base folder for ' + domain + ' / ' + locale +
          '. Add it in PRODUCT_LOCALE_FOLDERS or create that locale folder under the product\'s Ads folder and click Refresh.'
      };
    }

    // ✅ smart container (year or flat)
    const containerForCreate = getOrCreateYearContainerFlexible_(baseFolder, year);
    const containerForScan = pickScanContainer_(baseFolder, year);

    // 🔎 duplicate by Batch ID (scan in the correct container)
    const dupByBatch = getDuplicateBatchByIdPayload_Flexible_(baseFolder, year, fullName);
    if (dupByBatch && dupByBatch.status === 'exists') {
      return {
        status: 'batchDup',
        batchId: dupByBatch.batchId,
        existingFolderId: dupByBatch.existing.folderId || '',
        existingFolderUrl: dupByBatch.existing.folderUrl || '',
        yearFolderUrl: toDriveUrl_(containerForScan) || '',
        yearFolderId: containerForScan.getId() || '',
        meta: { domain, locale, year, c31Name: fullName }
      };
    }

    // No conflict → do normal create flow inside containerForCreate
    const dupByName = checkDuplicateInContainer_(containerForCreate, fullName);
    if (dupByName.status !== 'none') {
      dupByName.status = 'exists';
      return dupByName;
    }

    const batchFolder = getOrCreateSubfolder(containerForCreate, fullName);
    const adNotesFolder = getOrCreateSubfolder(batchFolder, 'Ad Notes');
    getOrCreateSubfolder(batchFolder, 'Work Files'); // ✅ NEW: Create Work Files folder
    
    // ✅ Pass opts object to capture original ad info
    const opts = { 
      origAdMode: origAdMode || 'auto',
      videoTemplate: videoTemplate || 'standard' // ✅ ADD THIS LINE
    };
    let url;
    try {
      Logger.log('🔧 About to create Ad Notes doc...');
      Logger.log('   Template ID will be determined by C31/C36');
      Logger.log('   Target folder: ' + adNotesFolder.getId());
      
      url = createAdNotesCoreOpt(adNotesFolder.getId(), populateAdName === true, opts);
      
      Logger.log('✅ Ad Notes doc created successfully: ' + url);
    } catch (docError) {
      logErrorToHub_('Ad Notes Creation', docError.message || String(docError));  // ← ADD THIS
      Logger.log('❌ Failed to create Ad Notes doc: ' + docError);
      Logger.log('   Stack: ' + (docError.stack || 'no stack'));
      
      // Return a more detailed error to the user
      return {
        status: 'error',
        message: '❌ Failed to create Ad Notes document.\n\n' +
                 'Error: ' + (docError.message || String(docError)) + '\n\n' +
                 '🔍 This usually means:\n' +
                 '1. Template file permissions issue\n' +
                 '2. Template file was deleted or moved\n' +
                 '3. Network/Drive API issue\n\n' +
                 'Please contact Kristijonas with this error message.'
      };
    }
    
    const result = { 
      status: 'created', 
      url, 
      adNotesUrl: toDriveUrl_(adNotesFolder), 
      adsUrl: toDriveUrl_(batchFolder) 
    };
    
    Logger.log('🔍 Checking opts for original ad info...');
    Logger.log('   opts.originalAdUrl: ' + (opts.originalAdUrl || 'MISSING'));
    Logger.log('   opts.originalAdName: ' + (opts.originalAdName || 'MISSING'));
    
    // ✅ ADD TO RESULT OBJECT (so the dialog can use it immediately)
    if (opts.originalAdUrl) {
      result.originalAdUrl = opts.originalAdUrl;
      result.originalAdName = opts.originalAdName;
      Logger.log('✅ Added original ad to result: ' + opts.originalAdName);
    } else {
      Logger.log('⚠️ No original ad info found in opts');
    }
    
    // ✅ ALSO save to record (for "Reopen Last Created")
    const recordData = {
      docUrl: result.url,
      adsUrl: result.adsUrl,
      adNotesUrl: result.adNotesUrl
    };
    
    if (opts.originalAdUrl) {
      recordData.originalAdUrl = opts.originalAdUrl;
      recordData.originalAdName = opts.originalAdName;
    }
    
    saveLastCreatedRecord_('standard', recordData);

    /**
 * Mark the last created record as deleted (so Reopen button shows friendly message)
 */
function markLastCreatedAsDeleted() {
  try {
    const props = PropertiesService.getUserProperties();
    const raw = props.getProperty('lastCreatedRecord');
    
    if (!raw) return;
    
    const record = JSON.parse(raw);
    record.deleted = true;
    record.deletedAt = new Date().toISOString();
    
    props.setProperty('lastCreatedRecord', JSON.stringify(record));
    Logger.log('✅ Marked last created record as deleted');
  } catch (e) {
    Logger.log('⚠️ Failed to mark record as deleted: ' + e);
  }
}
    
    // ✅ FIXED BLOCK:
    const c31Upper = fullNameRaw.toUpperCase();
    const typeRaw = sh.getRange('C36').getDisplayValue().trim();
    const typeUpper = (typeRaw || '').toUpperCase();
    const isItr = typeUpper.includes('ITR') || c31Upper.includes('ITR');
    
    let templateType = 'IMG';
    if (isItr) templateType = 'ITR';
    else if (typeUpper.includes('VID')) templateType = 'VID';
    
    logAnalyticsEventSafe_({
      action: 'Standard Setup',
      product: domain,
      locale: locale,
      batchId: batchId,
      template: templateType,
      success: true,
      additionalData: { year: year, populateAdName: populateAdName }
    });
    
    return result;


  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    const quota = /Service invoked too many times|Rate Limit|User rate limit/i.test(msg);
    const perm = /insufficient permissions|file not found|no access/i.test(msg);
    
    // ✅ NEW: Log the error
    const sh = SpreadsheetApp.getActiveSheet();
    const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
    const fullName = sanitizeBatchName(fullNameRaw).clean;
    const parts = parseNameParts(fullName);
    const domain = parseDomainFromC31(fullName);
    
    logAnalyticsEventSafe_({
      action: 'Standard Setup',
      product: domain || '',
      locale: parts.locale || '',
      batchId: parts.batch || '',
      success: false,
      error: msg.substring(0, 200)  // First 200 chars of error
    });
    
    if (quota) return { status: 'error', message: 'Drive is throttling requests (rate limit). Please try again shortly.' };
    if (perm) return { status: 'error', message: 'You might not have access to the target folder or the template. Please request access and retry.' };
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

function attemptCustomSetupWithBatchCheck(rawFolderInput, populateAdName, origAdMode, videoTemplate) {
  const sh = SpreadsheetApp.getActiveSheet();
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
    if (!fullNameRaw) return { status: 'error', message: 'C31 is empty.' };
    const fullName = sanitizeBatchName(fullNameRaw).clean;

    const parts = parseNameParts(fullName);
    const domain = parseDomainFromC31(fullName);
    const locale = parts.locale;
    const batchId = parts.batch;
    const year = detectTargetYear(fullName);

    if (!domain || !locale || !batchId) {
      return { status: 'error', message: 'Could not fully parse product / locale / batch ID from C31. Make sure naming is correct.' };
    }

    const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
    if (!baseFolder) {
      return {
        status: 'error',
        message: '🟡 Missing base folder for ' + domain + ' / ' + locale +
          '. Add it in PRODUCT_LOCALE_FOLDERS or create that locale folder under the product\'s Ads folder and click Refresh.'
      };
    }

    const containerForScan = pickScanContainer_(baseFolder, year);

    // same batch ID check (scan in target container)
    const dupByBatch = getDuplicateBatchByIdPayload_Flexible_(baseFolder, year, fullName);
    if (dupByBatch && dupByBatch.status === 'exists') {
      return {
        status: 'batchDup',
        batchId: dupByBatch.batchId,
        existingFolderId: dupByBatch.existing.folderId || '',
        existingFolderUrl: dupByBatch.existing.folderUrl || '',
        yearFolderUrl: toDriveUrl_(containerForScan) || '',
        yearFolderId: containerForScan.getId() || '',
        meta: { domain, locale, year, c31Name: fullName }
      };
    }

        // No conflict -> proceed with your current custom flow
    const result = createAdNotesFromPastedFolderForDialog(rawFolderInput, populateAdName, origAdMode, videoTemplate);
    
    // ✅ NEW: Try to create Work Files in the batch folder if we can find it
    if (result && result.status === 'created' && result.adsUrl) {
      try {
        const match = result.adsUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
        if (match) {
          const batchFolder = DriveApp.getFolderById(match[1]);
          getOrCreateSubfolder(batchFolder, 'Work Files');
        }
      } catch (e) {
        Logger.log('⚠️ Could not create Work Files in custom setup: ' + e);
      }
    }

if (result && result.status === 'created') {
  // ✅ normalize and save
  saveLastCreatedRecord_('custom', {
      docUrl: result.url || '',
      adsUrl: result.adsUrl || '',
      adNotesUrl: result.adNotesUrl || ''
    });
    
    // ✅ ADD THIS:
    logAnalyticsEventSafe_({
      action: 'Custom Setup',
      product: domain,
      locale: locale,
      batchId: batchId,
      success: true
    });
  }

  return result;


  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    const quota = /Service invoked too many times|Rate Limit|User rate limit/i.test(msg);
    const perm = /insufficient permissions|file not found|no access/i.test(msg);
    
    // ✅ NEW: Log the error
    const sh = SpreadsheetApp.getActiveSheet();
    const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
    const fullName = sanitizeBatchName(fullNameRaw).clean;
    const parts = parseNameParts(fullName);
    const domain = parseDomainFromC31(fullName);
    
    logAnalyticsEventSafe_({
      action: 'Custom Setup',
      product: domain || '',
      locale: parts.locale || '',
      batchId: parts.batch || '',
      success: false,
      error: msg.substring(0, 200)
    });
    
    if (quota) return { status: 'error', message: 'Drive is throttling requests (rate limit). Please try again shortly.' };
    if (perm) return { status: 'error', message: 'You might not have access to that folder or the template. Please request access and retry.' };
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}


function nukeExistingBatchAndRecreate_(batchFolderId, populateAdName) {
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    if (batchFolderId) {
      try {
        const f = DriveApp.getFolderById(batchFolderId);
        f.setTrashed(true); // goes to Drive Trash
      } catch (e) {
        Logger.log('⚠️ Could not trash batch folder: ' + e);
      }
    }

    // After nuking we force-create fresh with current C31
    return createAdNotesAutoRouteForDialog(populateAdName, /*forceCreate=*/ true);

  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    
    // ✅ NEW: Log the error
    try {
      const sh = SpreadsheetApp.getActiveSheet();
      const c31 = sh.getRange('C31').getDisplayValue().trim();
      const domain = parseDomainFromC31(c31);
      
      logAnalyticsEventSafe_({
        action: 'Replace Batch',
        product: domain || '',
        success: false,
        error: msg.substring(0, 200)
      });
    } catch (_) {
      // Silent fail
    }
    
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

/** Check for duplicates without creating anything (strict). */
function checkDuplicateInYear_(yearFolder, batchName) {
  const batchFolder = findExistingBatchFolder(yearFolder, batchName); // no creation
  if (!batchFolder) return { status: 'none' };

  const adNotes = getSubfolderIfExists(batchFolder, 'Ad Notes'); // no creation
  let doc = null;
  if (adNotes) {
    const it = adNotes.getFilesByName(batchName);
    if (it.hasNext()) doc = it.next();
  }

  const payload = {
    status: doc ? 'doc' : (adNotes ? 'adnotes' : 'folder'),
    existing: {
      folderId: batchFolder.getId(),
      folderUrl: toDriveUrl_(batchFolder),
      adNotesId: adNotes ? adNotes.getId() : '',
      adNotesUrl: adNotes ? adNotes.getUrl() : '',
      docUrl: doc ? doc.getUrl() : ''
    }
  };
  return payload;
}

/**
 * Scan the target Year folder for any batch folder that matches the same Batch ID as C31.
 * Returns {status:'none'} if not found or if we can't resolve base/locale.
 * If found, returns {status:'exists', existing:{folderId,folderUrl,adNotesId,adNotesUrl,docUrl}}
 */
function getDuplicateBatchByIdPayload_() {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31Raw = sh.getRange('C31').getDisplayValue().trim();
  if (!c31Raw) return { status: 'none' };

  const fullName = sanitizeBatchName(c31Raw).clean;
  const parts = parseNameParts(fullName);
  const domain = parseDomainFromC31(fullName);
  const locale = parts.locale;
  const batchId = parts.batch;               // number token, e.g. 113
  if (!domain || !locale || !batchId) return { status: 'none' };

  const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
  if (!baseFolder) return { status: 'none' };

  const year = detectTargetYear(fullName);
  const yearFolder = getOrCreateSubfolder(baseFolder, year);

  // Find any batch folder that has the same batch token
  const folderIter = yearFolder.getFolders();
  let matched = null;
  while (folderIter.hasNext()) {
    const f = folderIter.next();
    const p = parseNameParts(f.getName());
    if (
  p.batch === batchId &&
  p.product &&
  p.product.toUpperCase() === domain.toUpperCase() &&
  p.locale &&
  p.locale.toUpperCase() === locale.toUpperCase()
) {
  matched = f;
  break;
}

  }
  if (!matched) return { status: 'none' };

  // Gather Ad Notes + Doc (doc must match current <C31> name to treat as "the existing doc")
  let adNotes = getSubfolderIfExists(matched, 'Ad Notes');
  let doc = null;
  if (adNotes) {
    const it = adNotes.getFilesByName(fullName);
    if (it.hasNext()) doc = it.next();
  }

  return {
    status: 'exists',
    existing: {
      folderId: matched.getId(),
      folderUrl: toDriveUrl_(matched),
      adNotesId: adNotes ? adNotes.getId() : '',
      adNotesUrl: adNotes ? adNotes.getUrl() : '',
      docUrl: doc ? doc.getUrl() : ''
    },
    batchId: batchId
  };
}

function getDuplicateBatchByIdPayload_Flexible_(baseFolder, year, fullName) {
  try {
    const parts = parseNameParts(fullName);
    const domain = parseDomainFromC31(fullName);
    const locale = parts.locale;
    const batchId = parts.batch;
    if (!domain || !locale || !batchId) return { status: 'none' };

    // choose best container for scanning, but do not create any folder
    const container = pickScanContainer_(baseFolder, year);

    const it = container.getFolders();
    let matched = null;
    while (it.hasNext()) {
      const f = it.next();
      const p = parseNameParts(f.getName());
      if (
  p.batch === batchId &&
  p.product &&
  p.product.toUpperCase() === domain.toUpperCase() &&
  p.locale &&
  p.locale.toUpperCase() === locale.toUpperCase()
) {
  matched = f;
  break;
}

    }
    if (!matched) return { status: 'none' };

    // check inside matched for Ad Notes + doc with exact C31 name
    let adNotes = getSubfolderIfExists(matched, 'Ad Notes');
    let doc = null;
    if (adNotes) {
      const fit = adNotes.getFilesByName(fullName);
      if (fit.hasNext()) doc = fit.next();
    }

    return {
      status: 'exists',
      existing: {
        folderId: matched.getId(),
        folderUrl: toDriveUrl_(matched),
        adNotesId: adNotes ? adNotes.getId() : '',
        adNotesUrl: adNotes ? adNotes.getUrl() : '',
        docUrl: doc ? doc.getUrl() : ''
      },
      batchId: batchId
    };
  } catch (_) {
    return { status: 'none' };
  }
}

/**
 * Full diagnostics + duplicate-by-batch-ID check (used by Refresh and initial load).
 * Keeps the original structure you expect on the client, plus .dup.
 */
/**
 * FAST staged detection for popup
 *  - Instant: product / locale / platform from C31
 *  - Async parallel: Drive folder and template access
 *  - Returns in ~0.4s, feels instant
 */
function getCreateDialogDiagnosticsPlus() {
  try {
    Logger.log("🧠 Starting diagnostics…");

    const sh = SpreadsheetApp.getActiveSheet();
    const c31 = (sh.getRange('C31').getDisplayValue() || '').trim();
    const c3 = (sh.getRange('C3').getDisplayValue() || '').trim();
    const c16 = (sh.getRange('C16').getDisplayValue() || '').trim();

    if (typeof parseNameParts !== "function") throw new Error("Missing: parseNameParts()");
    if (typeof parseDomainFromC31 !== "function") throw new Error("Missing: parseDomainFromC31()");
    if (typeof PRODUCT_INDEX === "undefined") throw new Error("Missing: PRODUCT_INDEX");
    if (typeof PRODUCT_LOCALE_FOLDERS === "undefined") throw new Error("Missing: PRODUCT_LOCALE_FOLDERS");

    const p31 = parseNameParts(c31);
    const domain = parseDomainFromC31(c31);
    const locale = p31.locale;
    const prodMeta = PRODUCT_INDEX[domain] || {};

    // quick mismatches (optional)
    const pC3 = parseNameParts(c3);
    const pC16 = parseNameParts(c16);
    // Decide which field to compare against based on C31 pattern
    let sourceField = 'C3';
    const upper = c31.toUpperCase();

    if (upper.includes('ITR') || upper.includes('LOC') || upper.includes('LOCAL') || upper.includes('YTTOFB')) {
      sourceField = 'C16';
    } else if (upper.includes('NEW')) {
      sourceField = 'C3';
    }

    const srcParts = sourceField === 'C3' ? pC3 : pC16;
    const mismatches = [];
    const cmp = (a, b) => String(a || '').toUpperCase() === String(b || '').toUpperCase();
    if (p31.product && srcParts.product && !cmp(p31.product, srcParts.product)) mismatches.push({ field: 'product', c31: p31.product, src: srcParts.product });
    if (p31.platform && srcParts.platform && !cmp(p31.platform, srcParts.platform)) mismatches.push({ field: 'platform', c31: p31.platform, src: srcParts.platform });
    if (p31.locale && srcParts.locale && !cmp(p31.locale, srcParts.locale)) mismatches.push({ field: 'locale', c31: p31.locale, src: srcParts.locale });
    if (p31.batch && srcParts.batch && !cmp(p31.batch, srcParts.batch)) mismatches.push({ field: 'batch', c31: p31.batch, src: srcParts.batch });
    if (p31.adType && srcParts.adType && !cmp(p31.adType, srcParts.adType)) mismatches.push({ field: 'adType', c31: p31.adType, src: srcParts.adType });

    const diag = {
      isFb: p31.platform === 'FB',
      platformToken: p31.platform || '',
      domain,
      locale,
      productId: prodMeta.id || '',
      productLabel: (prodMeta.id ? prodMeta.id + ' ' : '') + (domain || ''),
      hasBase: !!(PRODUCT_LOCALE_FOLDERS[domain] && PRODUCT_LOCALE_FOLDERS[domain][locale]),
      missingLocale: false,
      consistency: mismatches.length ? { mismatches, sourceField } : null,
      templates: verifyTemplatesAccessible(),
      andromeda: {
        hasPermission: checkAndromedaPermission(),
        isNewVid: c31.toUpperCase().includes('NEW') && (c31.toUpperCase().includes('VID') || c31.toUpperCase().includes('VIDEO'))
      }
    };

    // duplicate-by-batch-id payload
    const dup = getDuplicateBatchByIdPayload_();

    Logger.log("✅ Finishing diagnostics");
    return { diag, dup };                     // <-- this is what the HTML needs
  } catch (err) {
    Logger.log("🚨 getCreateDialogDiagnosticsPlus error: " + err + "\n" + (err.stack || ""));
    // Still return a shape the UI can handle so it stops spinning
    return { diag: { error: String(err), isFb: false, platformToken: '', domain: '', locale: '', productId: '', productLabel: '', hasBase: false, missingLocale: true, consistency: null, templates: { IMG: false, VID: false, ITR: false, errors: [String(err)] } }, dup: { status: 'none' } };
  }
}

/**
 * Simple helper to get C31 value for the dialog
 */
function getC31Value() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    return sh.getRange('C31').getDisplayValue() || '';
  } catch (e) {
    return '';
  }
}

/**
 * Pre-fetch the original ad search result for iteration ads
 * Called immediately when dialog opens
 * Returns: { cached: true, result: {...} } or { cached: false }
 */
function prefetchOriginalAdIfIteration() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    const c31 = (sh.getRange('C31').getDisplayValue() || '').toUpperCase();
    const c16 = sh.getRange('C16').getDisplayValue().trim();
    
    // Only pre-fetch if this is an iteration ad
    if (!c31.includes('ITR') && !c31.includes('ITERATION')) {
      return { cached: false, reason: 'Not an iteration ad' };
    }
    
    if (!c16) {
      return { cached: false, reason: 'C16 is empty' };
    }
    
    Logger.log('🔄 Pre-fetching original ad for iteration...');
    
    // Parse the original ad info
    const parsed = parseOriginalAdId(c16);
    if (!parsed.hasOrig) {
      return { cached: false, reason: 'No original ad ID found' };
    }
    
    // Extract locale from origId
    const origLocale = parsed.origId.split('_')[0];
    
    // Try to find the product/locale folder
    const baseFolder = resolveBaseFolderForProductLocale(parsed.product, origLocale);
    if (!baseFolder) {
      return { cached: false, reason: 'Base folder not found' };
    }
    
    // Do the search
    const searchResult = findOriginalAdFile(parsed.origId, parsed.product, parsed.adType, baseFolder);
    
    Logger.log('✅ Pre-fetch complete: ' + searchResult.status);
    
    // ✅ Serialize the result for caching (can't store DriveApp objects)
    const cacheableResult = {
      status: searchResult.status,
      searchUrl: searchResult.searchUrl || ''
    };
    
    if (searchResult.status === 'found' && searchResult.file) {
      cacheableResult.file = {
        name: searchResult.file.getName(),
        url: searchResult.file.getUrl()
      };
    } else if (searchResult.status === 'multiple' && searchResult.files) {
      cacheableResult.files = searchResult.files.map(f => ({
        name: f.getName(),
        url: f.getUrl()
      }));
    } else if (searchResult.message) {
      cacheableResult.message = searchResult.message;
    }
    
    // Store in cache
    const cache = {
      timestamp: new Date().getTime(),
      c16: c16,
      result: cacheableResult
    };
    
    PropertiesService.getUserProperties().setProperty('origAdCache', JSON.stringify(cache));
    
    return { cached: true, result: searchResult };
    
  } catch (e) {
    Logger.log('⚠️ Pre-fetch failed: ' + e);
    return { cached: false, reason: String(e) };
  }
}

/**
 * Get cached original ad search result if still valid
 * Returns the cached result or null if expired/invalid
 */
function getCachedOriginalAdResult() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    const c16 = sh.getRange('C16').getDisplayValue().trim();
    
    const raw = PropertiesService.getUserProperties().getProperty('origAdCache');
    if (!raw) return null;
    
    const cache = JSON.parse(raw);
    
    // Check if cache is still valid (same C16 value, less than 5 minutes old)
    const age = new Date().getTime() - cache.timestamp;
    const maxAge = 5 * 60 * 1000; // 5 minutes
    
    if (cache.c16 !== c16) {
      Logger.log('⚠️ Cache invalid: C16 changed');
      return null;
    }
    
    if (age > maxAge) {
      Logger.log('⚠️ Cache expired (age: ' + Math.round(age/1000) + 's)');
      return null;
    }
    
    Logger.log('✅ Using cached result (age: ' + Math.round(age/1000) + 's)');
    return cache.result;
    
  } catch (e) {
    return null;
  }
}

/**
 * Background async verifier (runs silently ~2s later)
 */
function backgroundVerifyTemplatesAndCache() {
  try {
    verifyTemplatesAccessible(); // warms up DriveApp cache
  } catch (_) { }
}

function overrideAdNotesDocForDialog(batchFolderId, populateAdName) {
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(20 * 1000); } catch (_) { }

  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getActiveSheet();
    const c31Name = sh.getRange('C31').getDisplayValue().trim();
    if (!c31Name) return { status: 'error', message: 'C31 is empty.' };

    const batchFolder = DriveApp.getFolderById(batchFolderId);
    const adNotes = getOrCreateSubfolder(batchFolder, 'Ad Notes');

    const it = adNotes.getFilesByName(c31Name);
    if (it.hasNext()) {
      const oldDoc = it.next();
      try { oldDoc.setTrashed(true); }
      catch (e) { try { oldDoc.setName(c31Name + ' (OLD ' + new Date().toISOString().slice(0, 10) + ')'); } catch (_) { } }
    }

    const url = createAdNotesCoreOpt(adNotes.getId(), populateAdName === true);
    return { status: 'created', url, adNotesUrl: adNotes.getUrl(), batchUrl: batchFolder.getUrl() };
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    
    // ✅ NEW: Log the error
    try {
      const sh = SpreadsheetApp.getActiveSheet();
      const c31 = sh.getRange('C31').getDisplayValue().trim();
      const domain = parseDomainFromC31(c31);
      
      logAnalyticsEventSafe_({
        action: 'Override Doc',
        product: domain || '',
        success: false,
        error: msg.substring(0, 200)
      });
    } catch (_) {
      // Silent fail
    }
    
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

function detectTargetYear(fullName) {
  const m = String(fullName || '').match(/_(20\d{2})_/); // e.g., _2025_
  const currentYear = new Date().getFullYear();
  let year = m ? m[1] : String(currentYear);
  
  // Allow current year and one year before/after for flexibility
  const yearNum = parseInt(year);
  if (yearNum < 2024 || yearNum > currentYear + 1) {
    year = String(currentYear);
  }
  
  return year;
}

function normalizeProductLocaleFolders() {
  Object.keys(PRODUCT_LOCALE_FOLDERS).forEach(domain => {
    const map = PRODUCT_LOCALE_FOLDERS[domain] || {};
    LOCALES.forEach(loc => {
      if (!(loc in map)) map[loc] = ''; // add missing locale with empty string
    });
    PRODUCT_LOCALE_FOLDERS[domain] = map;
  });
}

/* ========= Core ========= */
function createAdNotesCore(targetFolderId) {
  return createAdNotesCoreOpt(targetFolderId, true); // default behavior unchanged
}

function createAdNotesAutoRouteForDialog(populateAdName, forceCreate) {
  const sh = SpreadsheetApp.getActiveSheet();
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
    if (!fullNameRaw) throw new Error('C31 is empty');

    const { clean: fullName } = sanitizeBatchName(fullNameRaw);
    const locale = parseLocaleFromC31(fullName);
    if (!locale) throw new Error('Could not detect a supported locale in C31.');

    const domain = parseDomainFromC31(fullName);
    if (!domain) throw new Error('Could not parse product domain from C31: ' + fullName);

    const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
    if (!baseFolder) {
      return {
        status: 'error',
        message: '🟡 Missing base folder for ' + domain + ' / ' + locale +
          '. Add it in PRODUCT_LOCALE_FOLDERS or create that locale folder under the product’s Ads folder and click Refresh.'
      };
    }

    const year = detectTargetYear(fullName);

    // 🟡 NEW: choose a smart container (year folder or flat root)
    const container = getOrCreateYearContainerFlexible_(baseFolder, year);

    // Strict duplicate check inside the chosen container
    const dup = checkDuplicateInContainer_(container, fullName);
    if (dup.status !== 'none' && !forceCreate) {
      dup.status = 'exists';
      return dup;
    }

    // Create path if needed inside the chosen container
    const batchFolder = dup.status === 'none'
      ? getOrCreateSubfolder(container, fullName)
      : findExistingBatchFolder(container, fullName);

    const adNotesFolder = getOrCreateSubfolder(batchFolder, 'Ad Notes');

    const url = createAdNotesCoreOpt(adNotesFolder.getId(), populateAdName === true);

    const result = {
      status: 'created',
      url,
      adNotesUrl: toDriveUrl_(adNotesFolder),
      adsUrl: toDriveUrl_(batchFolder)
    };

    // ✅ normalize and save
    saveLastCreatedRecord_('standard', {
      docUrl: url,  // ✅ Direct value
      adsUrl: toDriveUrl_(batchFolder),
      adNotesUrl: toDriveUrl_(adNotesFolder)
    });

    return result;

  } catch (e) {
    logErrorToHub_('Folder Creation', e.message || String(e));
    const msg = String(e && e.message ? e.message : e);
    const quota = /Service invoked too many times|Rate Limit|User rate limit/i.test(msg);
    const perm = /insufficient permissions|file not found|no access/i.test(msg);
    
    // ✅ NEW: Log the error
    const sh = SpreadsheetApp.getActiveSheet();
    const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
    const { clean: fullName } = sanitizeBatchName(fullNameRaw);
    const locale = parseLocaleFromC31(fullName);
    const domain = parseDomainFromC31(fullName);
    
    logAnalyticsEventSafe_({
      action: 'Auto Route',
      product: domain || '',
      locale: locale || '',
      success: false,
      error: msg.substring(0, 200)
    });
    
    if (quota) return { status: 'error', message: 'Drive is throttling requests (rate limit). Please try again in ~30–60s.' };
    if (perm) return { status: 'error', message: 'You might not have access to the target folder or template. Please request access and retry.' };
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

/** ---------- Discovered locale cache (persists across sessions) ---------- */
function _cacheKey(domain, locale) { return 'LOC_' + domain + '__' + locale; }

function cacheSetLocaleFolder(domain, locale, folderId) {
  try {
    PropertiesService.getScriptProperties()
      .setProperty(_cacheKey(domain, locale), folderId);
  } catch (e) { }
  // also reflect in current session map so it’s instant now
  if (PRODUCT_LOCALE_FOLDERS[domain]) {
    PRODUCT_LOCALE_FOLDERS[domain][locale] = folderId;
  }
}

function cacheGetLocaleFolder(domain, locale) {
  // 1) fast path: in-memory map
  const v = PRODUCT_LOCALE_FOLDERS[domain] && PRODUCT_LOCALE_FOLDERS[domain][locale];
  if (v) return v;

  // 2) persisted property
  try {
    return PropertiesService.getScriptProperties().getProperty(_cacheKey(domain, locale)) || '';
  } catch (e) {
    return '';
  }
}

/** Optional: clear all cached locale links (add a menu item to call it) */
function clearLocaleCache() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(k => { if (k.indexOf('LOC_') === 0) props.deleteProperty(k); });
  SpreadsheetApp.getUi().alert('Locale cache cleared.');
}


function createAdNotesChooseFolderPromptForDialog(populateAdName, /* forceCreate unused here */) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    const res = ui.prompt('Target folder', 'Paste a Drive FOLDER link/ID (or a FILE link/ID, parent will be used).', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return { status: 'cancelled' };

    const raw = (res.getResponseText() || '').trim();
    if (!raw) return { status: 'cancelled' };

    const targetFolder = resolveFolderStrict(raw);
    const url = createAdNotesCoreOpt(targetFolder.getId(), populateAdName === true);

    // Soft scaffold (no moves)
    let scaffoldUrl = '';
    try {
      const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
      const fullName = sanitizeBatchName(fullNameRaw).clean;
      const locale = parseLocaleFromC31(fullName);
      const domain = parseDomainFromC31(fullName);
      const year = detectTargetYear(fullName);

      const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
      if (baseFolder) {
        const yearFolder = getOrCreateSubfolder(baseFolder, year);
        const batchFolder = getOrCreateSubfolder(yearFolder, fullName);
        const adNotes = getOrCreateSubfolder(batchFolder, 'Ad Notes');
        scaffoldUrl = toDriveUrl_(adNotes);
      }
    } catch (_) { }

    return { status: 'created', url, folderUrl: scaffoldUrl };
  } catch (e) {
    logErrorToHub_('Custom Setup', e.message || String(e));
    const msg = String(e && e.message ? e.message : e);
    const quota = /Service invoked too many times|Rate Limit|User rate limit/i.test(msg);
    const perm = /insufficient permissions|file not found|no access/i.test(msg);
    if (quota) return { status: 'error', message: 'Drive is throttling requests (rate limit). Please try again shortly.' };
    if (perm) return { status: 'error', message: 'You might not have access to that folder or the template. Please request access and retry.' };
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

// --- Add this helper (anywhere with the other helpers) ---
function getOrCreateYearFolder(localeFolder, fullName) {
  const year = detectTargetYear(fullName);           // '2024' | '2025' | fallback
  return getOrCreateSubfolder(localeFolder, year);   // reuses your existing helper
}

function findFolderByName(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

/** Return a child folder by exact name if it exists. */
function getSubfolderIfExists(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

/** Find <year> subfolder if it exists (no creation). */
function findYearFolderIfExists_(baseFolder, year) {
  try {
    const it = baseFolder.getFoldersByName(String(year));
    return it.hasNext() ? it.next() : null;
  } catch (_) { return null; }
}

/** Does the locale folder already organize by year? (has any 20xx child) */
function hasAnyYearChild_(baseFolder) {
  const re = /^20\d{2}$/;
  const it = baseFolder.getFolders();
  while (it.hasNext()) {
    const name = it.next().getName();
    if (re.test(name)) return true;
  }
  return false;
}

/** Does the locale folder look "flat" (batches directly under it)? */
function hasAnyBatchChild_(baseFolder) {
  const it = baseFolder.getFolders();
  while (it.hasNext()) {
    const name = it.next().getName();
    const p = parseNameParts(name);
    if (p && p.batch) return true; // has a batch-like token
  }
  return false;
}

/**
 * Pick or create the container where the batch should live.
 * - If <year> exists: return it.
 * - Else if there are other year folders: create <year> and return it.
 * - Else if there are batch-like folders directly under locale: use locale root (no year).
 * - Else (empty/unclear): create <year>.
 */
function getOrCreateYearContainerFlexible_(baseFolder, year) {
  const yearName = String(year);
  const yearExisting = findYearFolderIfExists_(baseFolder, yearName);
  if (yearExisting) return yearExisting; // already exists

  const hasYear = hasAnyYearChild_(baseFolder);
  const hasFlat = hasAnyBatchChild_(baseFolder);

  if (hasYear) {
    // keep year-based structure
    return getOrCreateSubfolder(baseFolder, yearName);
  }
  if (hasFlat) {
    // respect flat legacy structure
    return baseFolder;
  }
  // empty/unclear → start year structure
  return getOrCreateSubfolder(baseFolder, yearName);
}

/**
 * Pick the best container for SCANNING (no creation).
 * - Prefer existing <year>.
 * - If not found, scan the locale root (flat).
 */
function pickScanContainer_(baseFolder, year) {
  const y = findYearFolderIfExists_(baseFolder, String(year));
  return y || baseFolder;
}

/** Generic duplicate check inside any container (no creation). */
function checkDuplicateInContainer_(containerFolder, batchName) {
  const batchFolder = findExistingBatchFolder(containerFolder, batchName);
  if (!batchFolder) return { status: 'none' };

  const adNotes = getSubfolderIfExists(batchFolder, 'Ad Notes');
  let doc = null;
  if (adNotes) {
    const it = adNotes.getFilesByName(batchName);
    if (it.hasNext()) doc = it.next();
  }
  return {
    status: doc ? 'doc' : (adNotes ? 'adnotes' : 'folder'),
    existing: {
      folderId: batchFolder.getId(),
      folderUrl: toDriveUrl_(batchFolder),
      adNotesId: adNotes ? adNotes.getId() : '',
      adNotesUrl: adNotes ? adNotes.getUrl() : '',
      docUrl: doc ? doc.getUrl() : ''
    }
  };
}

function createAdNotesNotesOnlyForDialog(populateAdName, origAdMode, videoTemplate) {
  try {
    const folderId = PARENT_FOLDER_ID ? PARENT_FOLDER_ID : null;
    
    // ✅ Pass opts to capture original ad info
    const opts = { 
      origAdMode: origAdMode || 'auto',
      videoTemplate: videoTemplate || 'standard' // ✅ ADD THIS LINE
    };
    const url = createAdNotesCoreOpt(folderId, populateAdName === true, opts);
    
    const result = { status: 'created', url, adNotesUrl: '', adsUrl: '' };

    // ✅ ADD to result object (so dialog can use it)
    if (opts.originalAdUrl) {
      result.originalAdUrl = opts.originalAdUrl;
      result.originalAdName = opts.originalAdName;
    }

    // ✅ ALSO save to record (for "Reopen Last Created")
    const recordData = {
      docUrl: url,
      adsUrl: '',
      adNotesUrl: ''
    };
    
    if (opts.originalAdUrl) {
      recordData.originalAdUrl = opts.originalAdUrl;
      recordData.originalAdName = opts.originalAdName;
    }

    saveLastCreatedRecord_('notes', recordData);

    // ✅ ADD THIS:
    logAnalyticsEventSafe_({
      action: 'Notes Only',
      product: parseDomainFromC31(SpreadsheetApp.getActiveSheet().getRange('C31').getValue()),
      success: true
    });

    return result;

  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    
    // ✅ NEW: Log the error
    try {
      const sh = SpreadsheetApp.getActiveSheet();
      const c31 = sh.getRange('C31').getDisplayValue().trim();
      const domain = parseDomainFromC31(c31);
      
      logAnalyticsEventSafe_({
        action: 'Notes Only',
        product: domain || '',
        success: false,
        error: msg.substring(0, 200)
      });
    } catch (_) {
      // Silent fail if we can't log
    }
    
    return { status: 'error', message: msg };
  }
}


/* ========= Helpers ========= */

/**
 * Resolve the base folder for a given product domain + locale.
 * Priority:
 *  1) If a direct mapping exists and is valid (we can open it) → use it.
 *  2) ✅ NEW: Handle LATAM/ES aliases - if user asks for one but folder is named the other, use it anyway
 *  3) Otherwise, try any known locale for that product, go to its parent (Ads),
 *     and look for a sibling folder named exactly <locale>. If found → use it
 *     and update the mapping to the FOLDER ID (not URL).
 * Options:
 *  - opts.noCache === true → ignore the stored mapping and always re-check Drive.
 */
function resolveBaseFolderForProductLocale(domain, locale, opts) {
  const noCache = !!(opts && opts.noCache);
  const productMap = PRODUCT_LOCALE_FOLDERS[domain];
  if (!productMap) return null;

  // ✅ NEW: Handle LATAM/ES aliases
  // If user asks for LATAM or ES, check if the configured folder exists and accept either name
  if (locale === 'LATAM' || locale === 'ES') {
    const directId = extractAnyDriveId(productMap[locale] || '');
    if (directId) {
      try {
        const testFolder = DriveApp.getFolderById(directId);
        const folderName = testFolder.getName().toUpperCase();
        
        Logger.log(`🔍 Checking ${locale} folder: found folder named "${testFolder.getName()}"`);
        
        // If the folder is named LATAM, ES, or contains either word, accept it for both locales
        const isLatamOrEs = folderName.includes('LATAM') || 
                           folderName.includes('ES') || 
                           folderName === 'LATAM' || 
                           folderName === 'ES';
        
        if (isLatamOrEs) {
          Logger.log(`✅ Using "${testFolder.getName()}" folder for ${locale} locale`);
          testFolder.getName(); // Touch to verify access
          return testFolder;
        }
      } catch (e) {
        Logger.log(`⚠️ Could not access ${locale} folder ID: ${directId}`);
        // Fall through to discovery logic
      }
    }
  }

  // 1) Direct mapping (only if not bypassing cache)
  if (!noCache) {
    const direct = productMap[locale];
    const directId = extractAnyDriveId(direct || '');
    if (directId) {
      try {
        const f = DriveApp.getFolderById(directId);
        // touch to ensure it exists / accessible
        f.getName();
        return f;
      } catch (_) {
        // fall through to discovery
      }
    }
  }

  // 2) Discovery via any known locale → parent Ads → sibling named <locale>
  const locales = Object.keys(productMap);
  for (const loc of locales) {
    const id = extractAnyDriveId(productMap[loc] || '');
    if (!id) continue;
    try {
      const knownLocaleFolder = DriveApp.getFolderById(id);
      knownLocaleFolder.getName(); // access check
      const parents = knownLocaleFolder.getParents();
      if (!parents.hasNext()) continue;

      const adsFolder = parents.next(); // Product → Channel → Ads
      
      // ✅ NEW: For LATAM/ES, search for either folder name
      if (locale === 'LATAM' || locale === 'ES') {
        // Try to find any folder named LATAM or ES
        const siblings = adsFolder.getFolders();
        while (siblings.hasNext()) {
          const sibling = siblings.next();
          const siblingName = sibling.getName().toUpperCase();
          
          const isLatamOrEs = siblingName.includes('LATAM') || 
                             siblingName.includes('ES') || 
                             siblingName === 'LATAM' || 
                             siblingName === 'ES';
          
          if (isLatamOrEs) {
            Logger.log(`✅ Discovered ${siblingName} folder for ${locale} via sibling search`);
            // Cache with canonical folder ID to keep it clean
            PRODUCT_LOCALE_FOLDERS[domain][locale] = sibling.getId();
            return sibling;
          }
        }
      } else {
        // Original logic for other locales
        const siblings = adsFolder.getFoldersByName(locale);
        if (siblings.hasNext()) {
          const found = siblings.next();
          // Cache with canonical folder ID to keep it clean
          PRODUCT_LOCALE_FOLDERS[domain][locale] = found.getId();
          return found;
        }
      }
    } catch (_) {
      // ignore and try next mapped locale
    }
  }

  // Nothing found
  return null;
}

/** Optional: clear a cached mapping if you ever need to force-refresh manually. */
function invalidateCachedLocaleFolder(domain, locale) {
  if (PRODUCT_LOCALE_FOLDERS[domain]) {
    PRODUCT_LOCALE_FOLDERS[domain][locale] = '';
  }
}

function parseDomainFromC31(s) {
  const m = String(s).trim().match(/^([A-Za-z0-9+]+)_/);
  if (!m) return '';
  return m[1].replace(/^"+|"+$/g, '');
}

function insertFootageLinkIntoFolderCell(body, domainKey) {
  const link = FOLDER_LINKS[domainKey] || '';
  const found = body.findText('{{FOLDER}}');
  if (!found) return;
  const el = found.getElement();
  const para = el.getParent().asParagraph();
  const cell = para.getParent();
  clearElementChildren(cell);
  const p = cell.appendParagraph('📂 FOOTAGE');
  const t = p.editAsText();
  if (link) t.setLinkUrl(0, t.getText().length - 1, link);
  t.setBold(true);
  t.setFontSize(9);
  try { t.setFontFamily('Lexend'); } catch (e) { }
}

function setOrgWideAccess(file, domain) {
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT); }
  catch (err) { Logger.log('Error setting domain access: ' + err); }
}

function copyTemplateToFolder(templateId, newName, targetFolder) {
  try {
    // Try to access the template
    const templateFile = DriveApp.getFileById(templateId);
    
    // Verify we can read it
    try {
      templateFile.getName();
    } catch (accessError) {
      throw new Error(
        '❌ Cannot access template file.\n\n' +
        'Template ID: ' + templateId + '\n\n' +
        'This usually means:\n' +
        '1. Template is not shared with you\n' +
        '2. Template was deleted or moved\n' +
        '3. You need to request access\n\n' +
        'Contact Kristijonas to resolve.\n\n' +
        'Technical: ' + accessError.message
      );
    }
    
    // ✅ FIXED: Always use root Drive if no target folder specified
    if (targetFolder) {
      return templateFile.makeCopy(newName, targetFolder);
    } else {
      // Create in root Drive, not in template's parent folder
      return templateFile.makeCopy(newName, DriveApp.getRootFolder());
    }
    
  } catch (e) {
    // Log for debugging
    Logger.log('❌ copyTemplateToFolder failed:');
    Logger.log('   Template ID: ' + templateId);
    Logger.log('   New name: ' + newName);
    
    let debugUser = 'unknown';
    try {
      debugUser = Session.getActiveUser().getEmail();
    } catch (e) {
      debugUser = 'anonymous (no email permission)';
    }
    Logger.log('   User: ' + debugUser);
    
    throw e; // Re-throw to surface to user
  }
}

function getOrCreateSubfolder(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

function replaceAll(body, map) {
  Object.keys(map).forEach(k => body.replaceText(escapeForReplace(k), map[k] ?? ''));
}

function escapeForReplace(s) {
  return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

function clearElementChildren(container) {
  while (container.getNumChildren && container.getNumChildren() > 0) {
    container.removeChild(container.getChild(0));
  }
}

function getCreateDialogDiagnostics() {
  // Just wraps the fast version for compatibility
  return getCreateDialogDiagnosticsPlus();
}


function verifyTemplatesAccessible() {
  const out = { IMG: false, VID: false, ITR: false, ANDROMEDA: false, errors: [] }; // ✅ ADDED ANDROMEDA
  const map = [
    { key: 'IMG', id: IMG_TEMPLATE_ID },
    { key: 'VID', id: VID_TEMPLATE_ID },
    { key: 'ITR', id: ITR_TEMPLATE_ID },
    { key: 'ANDROMEDA', id: ANDROMEDA_VID_TEMPLATE_ID } // ✅ ADDED THIS LINE
  ];
  map.forEach(({ key, id }) => {
    try {
      if (!id) throw new Error(key + ' template ID is empty');
      const f = DriveApp.getFileById(id); // throws if no access
      f.getName(); // touch
      out[key] = true;
    } catch (e) {
      out[key] = false;
      out.errors.push(key + ': ' + (e && e.message ? e.message : String(e)));
    }
  });
  return out;
}

function createAdNotesCoreOpt(targetFolderId, populateAdName, opts) {
  opts = opts || {};
  const origAdMode = opts.origAdMode || 'auto'; // 'auto' or 'blank'
  const videoTemplate = opts.videoTemplate || 'standard'; // ✅ NEW: 'standard' or 'andromeda'
  
  const sh = SpreadsheetApp.getActiveSheet(); // ✅ ADD NEW LINE HERE

  const fullName = sh.getRange('C31').getDisplayValue().trim();
  if (!fullName) throw new Error('C31 is empty');

  const domain = parseDomainFromC31(fullName);
  if (!domain) throw new Error('Could not parse meta domain from C31: ' + fullName);

  const prodMeta = PRODUCT_INDEX[domain];
  if (!prodMeta) throw new Error('Unknown meta domain: ' + domain);

  const docName = fullName;
  const productDisplay = prodMeta.id + ' ' + prodMeta.name + ' (' + domain + ')';
  const creative = sh.getName();

// ===== Template selection logic =====
const c31Raw = (sh.getRange('C31').getDisplayValue() || '').trim();
const c31Upper = c31Raw.toUpperCase();
const typeRaw = sh.getRange('C36').getDisplayValue().trim();
const typeUpper = (typeRaw || '').toUpperCase();

let templateId = '';
let isItrTemplate = false;
let isLocTemplate = false;

// ✅ FIXED: Extract text inside parentheses only
const parenMatch = c31Raw.match(/\(([^)]+)\)/);
const insideParens = parenMatch ? parenMatch[1].toUpperCase() : '';

Logger.log('🔍 Inside parentheses: "' + insideParens + '"');

// 1) Check for LOC/LOCALIZATION immediately after (
if (/^(LOC|LOCAL)/i.test(insideParens) && LOC_TEMPLATE_ID) {
  templateId = LOC_TEMPLATE_ID;
  isLocTemplate = true;
  Logger.log('✅ LOC template selected (found "' + insideParens + '" after "(")');

// 2) Check for ITR/ITERATION immediately after (
} else if (/^(ITR|ITER)/i.test(insideParens)) {
  templateId = ITR_TEMPLATE_ID;
  isItrTemplate = true;
  Logger.log('✅ ITR template selected (found "' + insideParens + '" after "(")');

// 3) "YTtoFB" template
} else if (insideParens.includes('YTTOFB') && YTTOFB_TEMPLATE_ID) {
  templateId = YTTOFB_TEMPLATE_ID;
  Logger.log('✅ YTtoFB template selected');

// 4) Explicitly handle "New DYN" → always use IMG template
} else if (c31Upper.includes('NEW DYN') || c31Upper.includes('NEWDYN')) {
  templateId = IMG_TEMPLATE_ID;
  Logger.log('✅ IMG template selected (New DYN)');

// ✅ NEW: 5) Check for New VID with Andromeda option
} else if (c31Upper.includes('NEW') && (c31Upper.includes('VID') || c31Upper.includes('VIDEO'))) {
  if (videoTemplate === 'andromeda' && ANDROMEDA_VID_TEMPLATE_ID) {
    templateId = ANDROMEDA_VID_TEMPLATE_ID;
    Logger.log('✨ Andromeda VID template selected');
  } else {
    templateId = VID_TEMPLATE_ID;
    Logger.log('✅ Standard VID template selected');
  }

// 6) Existing fallback logic from C36
} else if (typeUpper.includes('ITR')) {
  templateId = ITR_TEMPLATE_ID;
  isItrTemplate = true;
  Logger.log('✅ ITR template selected (from C36)');
} else if (typeUpper.includes('VID')) {
  templateId = VID_TEMPLATE_ID;
  Logger.log('✅ VID template selected (from C36)');
} else {
  templateId = IMG_TEMPLATE_ID;
  Logger.log('✅ IMG template selected (default)');
}

  if (!templateId) {
    throw new Error('No template ID resolved. Check LOC_TEMPLATE_ID / YTTOFB_TEMPLATE_ID / existing template IDs.');
  }

  // ✅ NEW: Verify template access BEFORE trying to copy
Logger.log('🔍 Checking template access...');
  Logger.log('   Template ID: ' + templateId);
  
  let debugUser = 'unknown';
  try {
    debugUser = Session.getActiveUser().getEmail();
  } catch (e) {
    debugUser = 'anonymous (no email permission)';
  }
  Logger.log('   User: ' + debugUser);
  
  try {
    const testFile = DriveApp.getFileById(templateId);
    const testName = testFile.getName();
    Logger.log('✅ Template accessible: "' + testName + '"');
    
    // Try to check if we can actually copy it
    const owner = testFile.getOwner();
    Logger.log('   Template owner: ' + (owner ? owner.getEmail() : 'Unknown'));
    
  } catch (accessError) {
    Logger.log('❌ Cannot access template: ' + accessError);
    
    const templateType = isLocTemplate ? 'LOC' : 
                        (c31Upper.includes('YTTOFB') ? 'YTTOFB' :
                        (c31Upper.includes('ITR') ? 'ITR' :
                        (typeUpper.includes('VID') ? 'VID' : 'IMG')));
    
    throw new Error(
      '❌ Cannot access the ' + templateType + ' template.\n\n' +
      'Template ID: ' + templateId + '\n\n' +
      '🔧 Fix:\n' +
      '1. Contact Kristijonas\n' +
      '2. Ask them to share the template with you as Editor\n' +
      '3. Or check if the template still exists\n\n' +
      'Technical error: ' + accessError.message
    );
  }

  // ===== Figure out the locale from the ad name =====
  // We already have parseNameParts() which gives us .locale
  const parts = parseNameParts(fullName);
  const localeCode = parts.locale || '';
  const localeHuman = mapLocaleCodeToLanguage(localeCode); // <-- new helper

  // ===== Pick/create the target folder =====
  let targetFolder = null;
  if (targetFolderId) {
    // Use the specified folder (from Standard/Custom Setup)
    targetFolder = DriveApp.getFolderById(targetFolderId);
  } else if (PARENT_FOLDER_ID && PARENT_FOLDER_ID.trim() !== '') {
    // Use default parent folder if defined
    const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
    targetFolder = parent;
  }
  // ✅ If targetFolder is still null, file will be created in root Drive

  // Copy template -> new file
  const newFile = copyTemplateToFolder(templateId, docName, targetFolder);

  // Open doc body for replacement
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  // ===== Build replacement map =====
  const replacements = {
    '{{PRODUCT}}': productDisplay,
    '{{CREATIVE}}': creative,
    '{{DOC_NAME}}': docName,
    '{{DOMAIN}}': domain
  };

  // Optional {{ADNAME}} from C3/C16 logic
  if (populateAdName === true) {
    const adName = getAdNameFromSheet();
    if (adName) {
      replacements['{{ADNAME}}'] = adName;
    }
  }

  // NEW: If it's the LOC template AND we have a human-readable language,
  // inject {{LOCALE}} = that language.
  //
  // Example:
  //   C31: "OriBreeze_FB_ES_35_(Loc VID_Senior_Invention)"
  //   parts.locale -> "ES"
  //   localeHuman  -> "Spanish"
  //
  //   We will replace {{LOCALE}} with "Spanish".
  if (isLocTemplate && localeHuman) {
    replacements['{{LOCALE}}'] = localeHuman;
  }

// ✅ NEW: {{OG_AD}} for iteration/localization template
if (isItrTemplate || isLocTemplate) {
  const templateType = isItrTemplate ? 'iteration' : 'localization';
  Logger.log('🔧 Processing ' + templateType + ' template, origAdMode: ' + origAdMode);
  
  if (origAdMode === 'blank') {
    // ✅ User chose to fill manually (cross-product inspiration)
    replacements['{{OG_AD}}'] = '[Fill in manually - taking inspiration from another product]';
    Logger.log('ℹ️ Original ad left blank per user choice (cross-product iteration)');
    
  } else {
      // origAdMode === 'auto' (default behavior)
      const adName = getAdNameFromSheet(); // Gets C16 for iterations
      const parsed = parseOriginalAdId(adName);
      
      if (!parsed.hasOrig) {
        // No original ID found
        replacements['{{OG_AD}}'] = '⚠️ No original ad reference found in ad name';
      } else {
        // ✅ NEW: Extract ad type from ORIGINAL ad ID, not from iteration name
        // Example: If origId is "US_226_3", look in that batch's files to determine type
        let originalAdType = parsed.adType; // Fallback to parsed type
        
        // ✅ BETTER: Try BOTH image and video, prioritize finding ANY match
        Logger.log('🔍 Searching for original ad (will try both IMG and VID if needed)');
        
        let searchResult = getCachedOriginalAdResult();
        
        if (searchResult) {
          Logger.log('⚡ Using pre-fetched result (instant!)');
          Logger.log('   Cache status: ' + searchResult.status);
          Logger.log('   Cache has file: ' + !!(searchResult.file));
        } else {
          // Not cached, do fresh search
          const origLocale = parsed.origId.split('_')[0];
          let targetSearchFolder = null;
          
          try {
            const baseFolder = resolveBaseFolderForProductLocale(parsed.product, origLocale);
            if (baseFolder) {
              targetSearchFolder = baseFolder;
              Logger.log('✅ Found product/locale folder for faster search: ' + baseFolder.getName());
            } else {
              Logger.log('⚠️ Could not find product/locale folder, will search master folder');
            }
          } catch (e) {
            Logger.log('⚠️ Error resolving target folder: ' + e);
          }
          
          // ✅ NEW: Try the parsed ad type first
          Logger.log('🔎 First attempt: Looking for ' + originalAdType + ' files');
          searchResult = findOriginalAdFile(parsed.origId, parsed.product, originalAdType, targetSearchFolder);
          
          // ✅ NEW: If not found, try the OPPOSITE type
          if (searchResult.status === 'notfound') {
            const oppositeType = (originalAdType === 'IMG') ? 'VID' : 'IMG';
            Logger.log('⚠️ Not found as ' + originalAdType + ', trying ' + oppositeType + ' instead');
            searchResult = findOriginalAdFile(parsed.origId, parsed.product, oppositeType, targetSearchFolder);
            
            if (searchResult.status === 'found') {
              Logger.log('✅ Found as ' + oppositeType + ' file!');
            }
          }
        }
        
        if (searchResult.status === 'found') {
          // Success! Insert hyperlinked filename
          Logger.log('✅ Found result processing...');
          
          // Handle both fresh results (DriveApp.File) and cached results (plain object)
          let fileName, fileUrl;
          
          if (typeof searchResult.file.getName === 'function') {
            // Fresh result from Drive
            Logger.log('   Type: Fresh Drive result');
            fileName = searchResult.file.getName();
            fileUrl = searchResult.file.getUrl();
          } else {
            // Cached result (plain object)
            Logger.log('   Type: Cached result');
            Logger.log('   searchResult.file: ' + JSON.stringify(searchResult.file));
            fileName = searchResult.file.name || searchResult.file.fileName || 'Original Ad';
            fileUrl = searchResult.file.url || searchResult.file.fileUrl || '';
          }
          
          Logger.log('   Final fileName: ' + fileName);
          Logger.log('   Final fileUrl: ' + fileUrl);
          
          replacements['{{OG_AD}}'] = fileName;
          
          // Store for hyperlinking after text replacement
          replacements['__OG_AD_URL__'] = fileUrl;
          
          // ✅ Store in opts so we can save it to the record later
          if (opts) {
            opts.originalAdUrl = fileUrl;
            opts.originalAdName = fileName;
          }
          
        } else if (searchResult.status === 'multiple') {
          // Multiple files found - error
          const count = searchResult.files.length;
          replacements['{{OG_AD}}'] = '⚠️ ERROR: Found ' + count + ' files with ID ' + parsed.origId + '. Check search: ' + searchResult.searchUrl;
          replacements['__OG_AD_IS_ERROR__'] = true;
          
        } else if (searchResult.status === 'notfound') {
          // Not found
          replacements['{{OG_AD}}'] = '⚠️ Original ad not found: ' + parsed.origId + ' | Search here: ' + searchResult.searchUrl;
          replacements['__OG_AD_IS_ERROR__'] = true;
          
        } else {
          // Error occurred
          replacements['{{OG_AD}}'] = '⚠️ Error searching for original: ' + (searchResult.message || 'Unknown error');
          replacements['__OG_AD_IS_ERROR__'] = true;
        }
      }
    }
}

// Do the replacements
replaceAll(body, replacements);

  // 🆕 Handle {{PRODUCT}} and {{PRODUCT_URL}} with hyperlink support
  const meta = PRODUCT_INDEX[domain] || {};
  const productName = meta.name || domain;
  const productUrl = meta.url || '';

  body.replaceText('{{PRODUCT}}', productName);

  if (productUrl) {
    const range = body.findText('{{PRODUCT_URL}}');
    if (range) {
      const el = range.getElement();
      el.setText(productUrl);
      el.setLinkUrl(productUrl);
    }
  } else {
    body.replaceText('{{PRODUCT_URL}}', '');
  }

  // Insert FOOTAGE link if we have one
  insertFootageLinkIntoFolderCell(body, domain);

// ✅ NEW: Hyperlink and style {{OG_AD}} if successful (works for both ITR and LOC)
if ((isItrTemplate || isLocTemplate) && replacements['__OG_AD_URL__']) {
  const ogRange = body.findText(escapeForReplace(replacements['{{OG_AD}}']));
  if (ogRange) {
    const el = ogRange.getElement().asText();
    const start = ogRange.getStartOffset();
    const end = ogRange.getEndOffsetInclusive();
    
    el.setLinkUrl(start, end, replacements['__OG_AD_URL__']);
    el.setForegroundColor(start, end, '#1a73e8'); // Blue
    el.setUnderline(start, end, true);
  }
}

// ✅ NEW: Style error messages in red + bold (works for both ITR and LOC)
if ((isItrTemplate || isLocTemplate) && replacements['__OG_AD_IS_ERROR__']) {
  const errorText = replacements['{{OG_AD}}'];
  const ogRange = body.findText(escapeForReplace(errorText));
  if (ogRange) {
    const el = ogRange.getElement().asText();
    const start = ogRange.getStartOffset();
    const end = ogRange.getEndOffsetInclusive();
    
    el.setForegroundColor(start, end, '#dc2626'); // Red
    el.setBold(start, end, true);
    
    // If there's a URL in the error message, hyperlink it
    const urlMatch = errorText.match(/(https:\/\/[^\s]+)/);
    if (urlMatch) {
      const url = urlMatch[1];
      const urlStart = errorText.indexOf(url);
      if (urlStart !== -1) {
        el.setLinkUrl(start + urlStart, start + urlStart + url.length - 1, url);
      }
    }
  }
}

doc.saveAndClose();

  // ✅ Get the document URL (needed for return and popup)
  const url = newFile.getUrl();

  // Org sharing stays the same
  try {
    newFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (e) {
    // not fatal
  }

  SpreadsheetApp.getActive().toast('Ad Notes created: ' + docName);
  return url;
}

/** Save last created info per user so they can reopen it later */
function saveLastCreatedRecord_(type, urls) {
  try {
    Logger.log('💾 Saving record - Input type: ' + type);
    Logger.log('💾 Saving record - Input urls: ' + JSON.stringify(urls));
    
    // Normalize the urls object to always have the same shape
    const normalizedUrls = {
      docUrl: urls.docUrl || urls.url || '', // ✅ ADD urls.url as fallback
      adsUrl: urls.adsUrl || '',
      adNotesUrl: urls.adNotesUrl || '',
      originalAdUrl: urls.originalAdUrl || '',      // ✅ ADD THIS
      originalAdName: urls.originalAdName || ''     // ✅ ADD THIS
    };
    
    Logger.log('💾 Normalized URLs: ' + JSON.stringify(normalizedUrls));
    
    const payload = {
      type: type || 'standard',
      urls: normalizedUrls,
      timestamp: new Date().toISOString()
    };
    
    PropertiesService.getUserProperties().setProperty('lastCreatedRecord', JSON.stringify(payload));
    Logger.log('✅ Saved last created record: ' + JSON.stringify(payload));
  } catch (e) {
    Logger.log('⚠️ Failed to save last created record: ' + e);
  }
}

/** Retrieve last created info for this user */
function getLastCreatedRecord_() {
  try {
    const raw = PropertiesService.getUserProperties().getProperty('lastCreatedRecord');
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

/** Public wrapper so HTML can call this safely */
function getLastCreatedRecord() {
  const record = getLastCreatedRecord_();
  Logger.log('🧠 getLastCreatedRecord returned: ' + JSON.stringify(record));
  return record;
}


function getAdNameFromSheet() {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31 = (sh.getRange('C31').getDisplayValue() || '').toUpperCase();

  // ✅ FIXED: Extract text inside parentheses to determine source
  const parenMatch = c31.match(/\(([^)]+)\)/);
  const insideParens = parenMatch ? parenMatch[1].toUpperCase() : '';
  
  // Check what's immediately after the opening parenthesis
  if (/^(ITR|ITER|LOC|LOCAL|YTTOFB)/i.test(insideParens)) {
    return sh.getRange('C16').getDisplayValue().trim();
  }
  
  // Detect "NEW" even if surrounded by parentheses or mixed case
  if (/\bNEW\b/i.test(c31) || /\(NEW/i.test(c31)) {
    return sh.getRange('C3').getDisplayValue().trim();
  }

  // Default fallback to C3
  return sh.getRange('C3').getDisplayValue().trim();
}

function resolveFolderStrict(input) {
  const id = extractAnyDriveId(input);
  if (!id) throw new Error('Could not read a Drive ID from your input');
  try {
    const f = DriveApp.getFolderById(id);
    f.getName();
    return f;
  } catch (e1) {
    try {
      const file = DriveApp.getFileById(id);
      const parents = file.getParents();
      if (!parents.hasNext()) throw new Error('Selected file has no parent folder');
      const parent = parents.next();
      parent.getName();
      return parent;
    } catch (e2) {
      throw new Error('Cannot access the target. Paste a folder link or a file link you can access. Details: ' + e2);
    }
  }
}

// --- Drop-in replacement: more robust ad-type & token parsing
function parseNameParts(str) {
  const raw = String(str || '').trim();
  if (!raw) return { product: '', platform: '', locale: '', batch: '', adType: '' };

  // Normalize: turn any non-alphanum into underscores, collapse repeats
  const canon = raw
    .toUpperCase()
    .replace(/[^A-Z0-9+]+/g, '_')  // keep A–Z, 0–9, and plus (for domains like DermaSonic+)
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');

  const tokens = canon.split('_').filter(Boolean);

  // Helper to read first matching token
  const first = (arr, set) => {
    for (const t of arr) if (set.has(t)) return t;
    return '';
  };

  // Product = first token from ORIGINAL (preserve case & plus), but sanitize parens/quotes.
  const cleanFirst = String((raw.split('_')[0] || '').trim())
    .replace(/^"+|"+$/g, '')
    .replace(/^\(+|\)+$/g, '')
    .trim();

  const product = cleanFirst;

  // Platform map (FB/META/IG→FB, TT/TIKTOK→TT, YT/YOUTUBE→YT, GD/GGL/GOOGLE→GD)
  const platformMap = { FB: 'FB', META: 'FB', IG: 'FB', TT: 'TT', TIKTOK: 'TT', YT: 'YT', YOUTUBE: 'YT', GD: 'GD', GGL: 'GD', GOOGLE: 'GD' };
  const platformRaw = tokens[1] || ''; // typical position
  const platform = platformMap[platformRaw] || platformMap[canon.match(/\b(FB|META|IG|TT|TIKTOK|YT|YOUTUBE|GD|GGL|GOOGLE)\b/)?.[1]] || (platformRaw || '');

  // Locale detection (3rd token preferred; else scan)
  const LOCALES = new Set(['US', 'UK', 'CA', 'AU', 'ES', 'DE', 'FR', 'IT', 'IL', 'EU', 'LATAM']);
  let locale = tokens[2] || '';
  if (!LOCALES.has(locale)) {
    locale = first(tokens, LOCALES);
  }

  // Batch ID: first 1–6 digit token anywhere
  let batch = '';
  for (const t of tokens) { if (/^\d{1,6}$/.test(t)) { batch = t; break; } }

  // Ad Type: check tokens anywhere (VID / IMG / DYN / CAR). Also accept VIDEO/IMAGE/DYNAMIC/CAROUSEL.
  // Priority if multiple appear by accident: VID > IMG > DYN > CAR
  const AD_VID = new Set(['VID', 'VIDEO']);
  const AD_IMG = new Set(['IMG', 'IMAGE']);
  const AD_DYN = new Set(['DYN', 'DYNAMIC']);
  const AD_CAR = new Set(['CAR', 'CAROUSEL']);
  const AD_LOC = new Set(['LOC', 'LOCAL', 'LOCALIZATION']);

  let adType = '';
  const has = (set) => tokens.some(t => set.has(t));
  if (has(AD_VID)) adType = 'VID';
  else if (has(AD_IMG)) adType = 'IMG';
  else if (has(AD_DYN)) adType = 'DYN';
  else if (has(AD_CAR)) adType = 'CAR';
  else if (has(AD_LOC)) adType = 'LOC';

  return { product, platform, locale, batch, adType };
}

// Decide which field to compare: C3 for NEW, C16 for Itr
function pickAdSourceFieldByC31() {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31 = (sh.getRange('C31').getDisplayValue() || '').toUpperCase();
  if (c31.includes('ITR')) return { field: 'C16', value: sh.getRange('C16').getDisplayValue().trim() };
  if (c31.includes('NEW')) return { field: 'C3', value: sh.getRange('C3').getDisplayValue().trim() };
  return { field: '', value: '' };
}

function createAdNotesFromPastedFolderForDialog(raw, populateAdName, origAdMode, videoTemplate) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const lock = LockService.getDocumentLock();
  try { lock.waitLock(30 * 1000); } catch (_) { }

  try {
    raw = (raw || '').trim();
    if (!raw) return { status: 'error', message: 'No input provided.' };

    const targetFolder = resolveFolderStrict(raw);
    const opts = {
      origAdMode: origAdMode || 'auto',
      videoTemplate: videoTemplate || 'standard'
    };
    const url = createAdNotesCoreOpt(targetFolder.getId(), populateAdName === true, opts);

    // Try to scaffold the standard route and return both URLs
    let adNotesUrl = '';
    let adsUrl = '';
    try {
      const fullNameRaw = sh.getRange('C31').getDisplayValue().trim();
      const fullName = sanitizeBatchName(fullNameRaw).clean;
      const locale = parseLocaleFromC31(fullName);
      const domain = parseDomainFromC31(fullName);
      const year = detectTargetYear(fullName);

      const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
      if (baseFolder) {
        const yearFolder = getOrCreateSubfolder(baseFolder, year);
        const batchFolder = getOrCreateSubfolder(yearFolder, fullName);
        const adNotes = getOrCreateSubfolder(batchFolder, 'Ad Notes');
        adNotesUrl = toDriveUrl_(adNotes);
        adsUrl = toDriveUrl_(batchFolder);
      }
    } catch (_) { }

    return { status: 'created', url, adNotesUrl, adsUrl };
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    const quota = /Service invoked too many times|Rate Limit|User rate limit/i.test(msg);
    const perm = /insufficient permissions|file not found|no access/i.test(msg);
    if (quota) return { status: 'error', message: 'Drive is throttling requests (rate limit). Please try again shortly.' };
    if (perm) return { status: 'error', message: 'You might not have access to that folder or the template. Please request access and retry.' };
    return { status: 'error', message: msg };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

function getNameConsistencyDiagnostics() {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31Raw = (sh.getRange('C31').getDisplayValue() || '').trim();
  const c3Raw = (sh.getRange('C3').getDisplayValue() || '').trim();
  const c16Raw = (sh.getRange('C16').getDisplayValue() || '').trim();

  if (!c31Raw) {
    return { sourceField: '', mismatches: [], partsC31: {}, partsSrc: {} };
  }

  // ✅ FIXED: Extract text inside parentheses to determine source field
  const parenMatch = c31Raw.match(/\(([^)]+)\)/);
  const insideParens = parenMatch ? parenMatch[1].toUpperCase() : '';
  
  let sourceField = 'C3';
  let sourceRaw = c3Raw;
  
  // Check what's immediately after the opening parenthesis
  if (/^(ITR|ITER|LOC|LOCAL|YTTOFB)/i.test(insideParens)) {
    sourceField = 'C16';
    sourceRaw = c16Raw;
  } else if (/^NEW/i.test(insideParens) || insideParens === '' || c31Raw.toUpperCase().includes('NEW')) {
    sourceField = 'C3';
    sourceRaw = c3Raw;
  }

  const p31 = parseNameParts(c31Raw);
  const pSrc = parseNameParts(sourceRaw);

  const mismatches = [];
  const cmp = (a, b) => String(a || '').trim().toUpperCase() === String(b || '').trim().toUpperCase();

  if (p31.product && pSrc.product && !cmp(p31.product, pSrc.product)) mismatches.push({ field: 'product', c31: p31.product, src: pSrc.product });
  if (p31.platform && pSrc.platform && !cmp(p31.platform, pSrc.platform)) mismatches.push({ field: 'platform', c31: p31.platform, src: pSrc.platform });
  if (p31.locale && pSrc.locale && !cmp(p31.locale, pSrc.locale)) mismatches.push({ field: 'locale', c31: p31.locale, src: pSrc.locale });
  if (p31.batch && pSrc.batch && !cmp(p31.batch, pSrc.batch)) mismatches.push({ field: 'batch', c31: p31.batch, src: pSrc.batch });
  if (p31.adType && pSrc.adType && !cmp(p31.adType, pSrc.adType)) mismatches.push({ field: 'adType', c31: p31.adType, src: pSrc.adType });

  // ✅ NEW: Check for TVStick/Flixy → TellyStick rename
  let renameWarning = null;
  const c31Product = parseDomainFromC31(c31Raw).toUpperCase();
  const srcProduct = parseDomainFromC31(sourceRaw).toUpperCase();
  
  const oldNames = ['TVSTICK', 'FLIXY'];
  const c31NeedsRename = oldNames.includes(c31Product);
  const srcNeedsRename = oldNames.includes(srcProduct);
  
  if (c31NeedsRename || srcNeedsRename) {
    let message = '⚠️ <strong>Product renamed:</strong> ';
    const oldName = c31NeedsRename ? c31Product : srcProduct;
    
    if (oldName === 'FLIXY') {
      message += '<strong>Flixy</strong> is now called <strong>TellyStick</strong>.';
    } else {
      message += '<strong>TVStick</strong> is now called <strong>TellyStick</strong>.';
    }
    
    message += '<br><br><strong>What to do:</strong><ul style="margin:8px 0 0 20px;padding:0;">';
    
    if (c31NeedsRename) {
      message += '<li>Update <strong>C31</strong>: Change "' + oldName.charAt(0) + oldName.slice(1).toLowerCase() + '" to "TellyStick"</li>';
    }
    
    if (srcNeedsRename) {
      const cellToChange = sourceField === 'C3' ? 'C4' : 'C17';
      message += '<li>Update <strong>' + cellToChange + '</strong>: Change the product dropdown to "TellyStick"</li>';
    }
    
    message += '<li>Click <strong>Refresh consistency</strong> after updating</li></ul>';
    message += '<br><em style="font-size:12px;color:#64748b;">Need help? Contact Patricija for furhter questions.</em>';
    
    renameWarning = {
      message: message,
      affectedFields: {
        c31: c31NeedsRename,
        source: srcNeedsRename,
        sourceCell: sourceField === 'C3' ? 'C4' : 'C17'
      }
    };
  }

  // ✅ NEW: Check Andromeda feature
  const hasAndromedaPermission = checkAndromedaPermission();
  const c31Upper = c31Raw.toUpperCase();
  const isNewVid = c31Upper.includes('NEW') && (c31Upper.includes('VID') || c31Upper.includes('VIDEO'));

  return {
    sourceField,
    mismatches,
    partsC31: p31,
    partsSrc: pSrc,
    c31Value: c31Raw,
    renameWarning: renameWarning,
    andromeda: {
      hasPermission: hasAndromedaPermission,
      isNewVid: isNewVid
    }
  };
}

function findExistingBatchFolder(parentYearFolder, batchName) {
  const it = parentYearFolder.getFoldersByName(batchName);
  return it.hasNext() ? it.next() : null;
}

function findExistingDocInAdNotes(adNotesFolder, docName) {
  const it = adNotesFolder.getFilesByName(docName);
  return it.hasNext() ? it.next() : null;
}

function toDriveUrl_(entry) {
  try {
    return entry && entry.getUrl ? entry.getUrl() : '';
  } catch (_) { return ''; }
}

function sanitizeBatchName(raw) {
  // Keep original exactly as-is unless it starts with whitespace.
  const s0 = String(raw == null ? '' : raw);
  if (!s0) return { clean: '', changed: false };

  // Only remove leading whitespace (spacebar gap at start).
  if (/^\s+/.test(s0)) {
    const clean = s0.replace(/^\s+/, '');
    return { clean, changed: true };
  }
  return { clean: s0, changed: false };
}

function extractAnyDriveId(input) {
  if (!input) return '';
  const m1 = input.match(/\/folders\/([a-zA-Z0-9_-]{10,})/);
  if (m1) return m1[1];
  const m2 = input.match(/\/file\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m2) return m2[1];
  const m3 = input.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m3) return m3[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(input)) return input;
  return '';
}

/**
 * Delete a batch folder (admin only, for testing)
 * Moves folder to trash, doesn't permanently delete
 */
function deleteBatchFolderForDialog(folderId) {
  try {
    // ✅ Admin check
    let currentUser = '';
    try {
      currentUser = Session.getActiveUser().getEmail();
    } catch (e) {
      return { status: 'error', message: 'Could not verify user permissions.' };
    }
    
    if (!ADMIN_EMAILS.includes(currentUser)) {
      return { status: 'error', message: 'This function is only available to admins.' };
    }
    
    if (!folderId) {
      return { status: 'error', message: 'No folder ID provided.' };
    }
    
    // Get the folder
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    // Move to trash (not permanent delete)
    folder.setTrashed(true);
    
    Logger.log('🗑️ Admin deleted test batch: ' + folderName);
    
    return { 
      status: 'deleted', 
      message: 'Batch folder "' + folderName + '" moved to trash successfully.' 
    };
    
  } catch (e) {
    Logger.log('❌ Delete batch folder failed: ' + e);
    return { 
      status: 'error', 
      message: 'Failed to delete folder: ' + (e.message || String(e)) 
    };
  }
}

/* ========= Picker HTML with debug ========= */

function buildPickerHtml_Debug() {
  const devKey = PICKER_DEVELOPER_KEY;
  return `
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Drive folder picker</title>
<style>
  @keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}
  body{font-family:system-ui,Arial,sans-serif;margin:0}
  #log{padding:12px;font-size:13px;white-space:pre-wrap}
  #err{padding:12px;background:#f7f7f7;border-top:1px solid #ddd;color:#900;display:none;white-space:pre-wrap}
</style>
<script>

function showLoading(on){
  const el = document.getElementById('loadingAnim');
  if (!el) return;
  el.style.display = on ? 'block' : 'none';
}


let oauthToken=null, pickerApiLoaded=false;
function log(s){const el=document.getElementById('log'); el.textContent+=s+"\\n";}
function showErr(e){const el=document.getElementById('err'); el.style.display='block'; el.textContent=String(e);}
function loadPicker(){
  log('Loading Google APIs...');
  // If api.js fails to load within 8s, show hint
  setTimeout(()=>{ if(!window.gapi){ showErr('google APIs failed to load. Allow https://apis.google.com and third-party cookies.'); } },8000);
  gapi.load('picker', {'callback':()=>{pickerApiLoaded=true;log('Picker API loaded'); maybeCreatePicker();}});
  gapi.load('auth',   {'callback':()=>{log('Requesting OAuth token'); google.script.run.withSuccessHandler(t=>{oauthToken=t;log('Token received, length '+String(t||'').length); maybeCreatePicker();}).withFailureHandler(e=>showErr('getOAuthToken failed: '+(e&&e.message?e.message:e))).getOAuthToken();}});
}
function maybeCreatePicker(){
  if(!pickerApiLoaded || !oauthToken) return;
  if(!'${devKey}' || '${devKey}'==='YOUR_API_KEY'){ showErr('Developer key missing in script'); return; }
  try{
    const view=new google.picker.DocsView(google.picker.ViewId.FOLDERS)
      .setSelectFolderEnabled(true)
      .setMimeTypes('application/vnd.google-apps.folder')
      .setMode(google.picker.DocsViewMode.LIST);
    const picker=new google.picker.PickerBuilder()
      .setOAuthToken(oauthToken)
      .setDeveloperKey('${devKey}')
      .addView(view)
      .setCallback(data=>{
        if(data.action===google.picker.Action.PICKED && data.docs && data.docs.length){
          const id=data.docs[0].id;
          google.script.run.withSuccessHandler(url=>{window.open(url,'_blank');google.script.host.close();})
            .withFailureHandler(e=>showErr(e&&e.message?e.message:e))
            .createAdNotesInFolder(id);
        }else if(data.action===google.picker.Action.CANCEL){google.script.host.close();}
      })
      .build();
    picker.setVisible(true);
    log('Picker shown. If still blank, API key project mismatch or cookies blocked.');
  }catch(e){ showErr(e); }
}
</script>
<script src="https://apis.google.com/js/api.js?onload=loadPicker"></script>
</head>
<body>
  <div id="log">Starting…</div>
  <div id="err"></div>
</body>
</html>`;
}

// Expose token to Picker
function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}
function parseLocaleFromC31(s) {
  // Find the first underscore-separated token that matches one of the supported LOCALES
  const parts = String(s || '').split('_').map(x => x.trim()).filter(Boolean);
  for (let i = 0; i < parts.length; i++) {
    const t = parts[i].toUpperCase();
    if (Array.isArray(LOCALES) && LOCALES.indexOf(t) !== -1) return t;
  }
  return '';
}

function mapLocaleCodeToLanguage(localeCode) {
  // normalize
  const code = String(localeCode || '').toUpperCase().trim();

  // you can extend this map any time
  const MAP = {
    US: 'English (US)',
    UK: 'English (UK)',
    CA: 'English (Canada)',
    AU: 'English (Australia)',
    ES: 'Spanish',
    LATAM: 'Spanish (LATAM)',
    DE: 'German',
    FR: 'French',
    IT: 'Italian',
    IL: 'Hebrew',
    EU: 'English (EU)', // you tell me what you want here
    IL: 'Israel',
  };

  // fallback: just return the code itself if we don't know
  return MAP[code] || code || '';
}

/**
 * Extract the original ad ID from an iteration ad name.
 * Example: "WaveMax_FB_US_306_1_orig_US_64_2_VID_..." → "US_64_2"
 */
function parseOriginalAdId(adName) {
  const tokens = String(adName || '').split('_').filter(Boolean);
  
  // 1. Product is always first token
  const product = tokens[0] || '';
  
  // 2. Find ad type (VID or IMG)
  let adType = null;
  for (const t of tokens) {
    const upper = t.toUpperCase();
    if (upper === 'VID' || upper === 'VIDEO') {
      adType = 'VID';
      break;
    }
    if (upper === 'IMG' || upper === 'IMAGE') {
      adType = 'IMG';
      break;
    }
  }
  
  // 3. Find "orig" marker (case-insensitive)
  let origIndex = -1;
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i].toLowerCase() === 'orig') {
      origIndex = i;
      break;
    }
  }
  
  // 4. Extract original ID
  let origId = '';
  
  if (origIndex !== -1) {
    // Pattern: ..._orig_LOCALE_BATCH_VAR_...
    if (origIndex + 3 < tokens.length) {
      origId = tokens[origIndex + 1] + '_' + tokens[origIndex + 2] + '_' + tokens[origIndex + 3];
    }
  } else {
    // No "orig" marker - look for second LOCALE_BATCH_VAR pattern
    const LOCALES = new Set(['US', 'UK', 'CA', 'AU', 'ES', 'DE', 'FR', 'IT', 'IL', 'EU', 'LATAM']);
    let localeMatches = [];
    
    for (let i = 0; i < tokens.length - 2; i++) {
      const upper = tokens[i].toUpperCase();
      if (LOCALES.has(upper)) {
        // Check if next two tokens are numbers (batch and variation)
        if (/^\d+$/.test(tokens[i + 1]) && /^\d+$/.test(tokens[i + 2])) {
          localeMatches.push({
            index: i,
            id: tokens[i] + '_' + tokens[i + 1] + '_' + tokens[i + 2]
          });
        }
      }
    }
    
    // If we found 2+ locale sequences, the second one is the original
    if (localeMatches.length >= 2) {
      origId = localeMatches[1].id;
    }
  }
  
  return {
    origId: origId,
    product: product,
    adType: adType,
    hasOrig: !!origId
  };
}

/**
 * Get files from folder with smart depth control
 * If folder looks like a batch folder, only scan 1 level
 * Otherwise scan 2 levels (root + year folders + their subfolders)
 */
function getAllFilesRecursive_(folder) {
  const files = [];
  const folderName = folder.getName();
  const yearPattern = /^20\d{2}$/;
  
  // Check if this looks like a batch folder (contains batch pattern like "FB_US_252_")
  const isBatchFolder = /FB_[A-Z]{2}_\d+_\d+/i.test(folderName);
  
  Logger.log('🔍 Starting search in: ' + folderName + (isBatchFolder ? ' (batch folder)' : ''));
  
  if (isBatchFolder) {
    // ✅ FAST PATH: We're already in a batch folder, just scan files here
    Logger.log('   ⚡ Fast scan: Only checking files in this folder');
    const batchFiles = folder.getFiles();
    let count = 0;
    while (batchFiles.hasNext()) {
      files.push(batchFiles.next());
      count++;
    }
    Logger.log('   ✓ Found ' + count + ' files');
    return files;
  }
  
  // SLOW PATH: Need to search through folder structure
  Logger.log('📁 Scanning root level...');
  const rootFiles = folder.getFiles();
  let rootCount = 0;
  while (rootFiles.hasNext()) {
    files.push(rootFiles.next());
    rootCount++;
  }
  Logger.log('   ✓ Found ' + rootCount + ' files in root');
  
  Logger.log('📁 Checking subfolders...');
  const subfolders = folder.getFolders();
  
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const name = subfolder.getName();
    
    if (yearPattern.test(name)) {
      // Year folder - scan it and its immediate subfolders
      Logger.log('   📅 Scanning year folder: ' + name);
      
      const yearFiles = subfolder.getFiles();
      let yearCount = 0;
      while (yearFiles.hasNext()) {
        files.push(yearFiles.next());
        yearCount++;
      }
      Logger.log('      ✓ Found ' + yearCount + ' files in ' + name);
      
      // Check one level of subfolders
      const yearSubfolders = subfolder.getFolders();
      let subCount = 0;
      while (yearSubfolders.hasNext()) {
        const yearSub = yearSubfolders.next();
        const subName = yearSub.getName();
        
        const subFiles = yearSub.getFiles();
        let subFileCount = 0;
        while (subFiles.hasNext()) {
          files.push(subFiles.next());
          subFileCount++;
        }
        
        if (subFileCount > 0) {
          Logger.log('      📂 ' + name + '/' + subName + ': ' + subFileCount + ' files');
        }
        subCount++;
      }
      
    } else {
      Logger.log('   ⏭️  Skipping: ' + name);
    }
  }
  
  Logger.log('✅ Search complete: ' + files.length + ' total files');
  return files;
}

/**
 * Helper: Search for files matching the pattern
 * FIXED: More flexible matching - checks if origId appears anywhere in filename
 */
function searchForOriginalAd_(folder, product, origId, adType, separator) {
  const matched = [];
  
  // Get all files recursively
  const allFiles = getAllFilesRecursive_(folder);
  
  Logger.log('🔎 Searching for: product=' + product + ', origId=' + origId + ', adType=' + adType + ', separator=' + separator);
  
  const productLower = product.toLowerCase();
  const origIdPattern = origId.replace(/_/g, separator).toUpperCase();
  
  Logger.log('   Pattern to match: starts with "' + productLower + '" AND contains "' + origIdPattern + '"');
  
  let checkedCount = 0;
  for (const file of allFiles) {
    const name = file.getName();
    const nameLower = name.toLowerCase();
    const nameUpper = name.toUpperCase();
    checkedCount++;
    
    // 1. Must start with product name (case-insensitive)
    if (!nameLower.startsWith(productLower)) continue;
    
    // 2. Check file type matches ad type
    if (adType === 'VID') {
      const mimeType = file.getMimeType();
      if (!mimeType.includes('video')) {
        continue;
      }
    } else if (adType === 'IMG') {
      const mimeType = file.getMimeType();
      if (!mimeType.includes('image')) {
        continue;
      }
    }
    
    // 3. ✅ SIMPLIFIED: Just check if the pattern exists anywhere in the filename
    const tokens = nameUpper.split(/[_-]/).filter(Boolean);
    
    // Build all possible 3-token sequences and check if any match
    for (let i = 0; i < tokens.length - 2; i++) {
      const sequence = tokens[i] + separator + tokens[i + 1] + separator + tokens[i + 2];
      
      if (sequence === origIdPattern) {
        Logger.log('   ✅ MATCH FOUND: ' + name);
        Logger.log('      Matched sequence: ' + sequence);
        matched.push(file);
        break; // Stop checking this file
      }
    }
  }
  
  Logger.log('   Checked ' + checkedCount + ' files, found ' + matched.length + ' matches');
  return matched;
}

/**
 * Try to find the specific batch folder for the original ad
 * Returns the folder if found, null otherwise
 * Example: origId "US_252_3" → finds folder "CamTrix_FB_US_252_1" or "CamTrix_FB_US_252_2"
 */
function findBatchFolder_(baseFolder, product, origId) {
  try {
    // Extract just the batch number (e.g., "US_252_3" → "US_252")
    const tokens = origId.split('_');
    if (tokens.length < 2) return null;
    
    const locale = tokens[0]; // "US"
    const batchNum = tokens[1]; // "252"
    const batchPrefix = product + '_FB_' + locale + '_' + batchNum; // "CamTrix_FB_US_252"
    
    Logger.log('🎯 Smart search: Looking for batch folder matching "' + batchPrefix + '_*"');
    
    const yearPattern = /^20\d{2}$/;
    
    // Helper to search folders for batch match
    function searchFoldersForBatch(container, containerName) {
      const folders = container.getFolders();
      while (folders.hasNext()) {
        const folder = folders.next();
        const name = folder.getName();
        
        // Check if folder name starts with our batch prefix
        // E.g., "CamTrix_FB_US_252_1" or "CamTrix_FB_US_252_2"
        if (name.toUpperCase().startsWith(batchPrefix.toUpperCase() + '_')) {
          Logger.log('   ✅ Found batch folder in ' + containerName + ': ' + name);
          return folder;
        }
      }
      return null;
    }
    
    // 1. Try root level first
    Logger.log('   Searching root level...');
    let batchFolder = searchFoldersForBatch(baseFolder, 'root');
    if (batchFolder) return batchFolder;
    
    // 2. Try year folders
    Logger.log('   Searching year folders...');
    const yearFolders = baseFolder.getFolders();
    while (yearFolders.hasNext()) {
      const yearFolder = yearFolders.next();
      const yearName = yearFolder.getName();
      
      if (yearPattern.test(yearName)) {
        Logger.log('   Checking year: ' + yearName);
        batchFolder = searchFoldersForBatch(yearFolder, yearName);
        if (batchFolder) return batchFolder;
      }
    }
    
    Logger.log('   ⚠️ No batch folder found, will scan all files');
    return null;
    
  } catch (e) {
    Logger.log('   ⚠️ Error finding batch folder: ' + e);
    return null;
  }
}

/**
/**
 * Search for the original ad file in Drive.
 * First tries to find the specific batch folder, then searches only there.
 * Falls back to scanning all files if batch folder not found.
 */
function findOriginalAdFile(origId, product, adType, targetFolder) {
  const MASTER_ADS_FOLDER_ID = '151s5U5ZemSfaZ4JKXRWm_MhRiJzw0hHq';
  
  try {
    // Use provided folder, or fall back to master
    let searchFolder;
    if (targetFolder) {
      searchFolder = targetFolder;
      Logger.log('🔍 Searching in specific folder: ' + targetFolder.getName());
    } else {
      searchFolder = DriveApp.getFolderById(MASTER_ADS_FOLDER_ID);
      Logger.log('🔍 Searching in master folder (fallback)');
    }
    
    // ✅ NEW: Try to find the specific batch folder first (FAST)
    const batchFolder = findBatchFolder_(searchFolder, product, origId);
    if (batchFolder) {
      Logger.log('🚀 Smart search: Scanning only batch folder (fast path)');
      searchFolder = batchFolder; // Search only inside this folder
    } else {
      Logger.log('📂 Standard search: Scanning all files (slow path)');
    }
    
    // ✅ FIXED: Try both underscore and dash separators
    Logger.log('🔎 Searching with underscores: ' + origId);
    let matchedFiles = searchForOriginalAd_(searchFolder, product, origId, adType, '_');
    
    // If not found with underscores, try dashes
    if (matchedFiles.length === 0) {
      Logger.log('🔎 Not found with underscores, trying dashes: ' + origId.replace(/_/g, '-'));
      const origIdDash = origId.replace(/_/g, '-');
      matchedFiles = searchForOriginalAd_(searchFolder, product, origIdDash, adType, '-');
    }
    
    // Build search URL for user to check manually
    const searchQuery = 'title:' + origId + ' title:' + product;
    const searchUrl = 'https://drive.google.com/drive/u/0/search?q=' + encodeURIComponent(searchQuery);
    
    Logger.log('📊 Search complete: found ' + matchedFiles.length + ' files');
    
    if (matchedFiles.length === 0) {
      return { status: 'notfound', searchUrl: searchUrl };
    }
    
    if (matchedFiles.length === 1) {
      Logger.log('✅ Found exactly 1 file: ' + matchedFiles[0].getName());
      return { status: 'found', file: matchedFiles[0], searchUrl: searchUrl };
    }
    
    // Multiple matches - return the first one (could be improved to pick best match)
    Logger.log('⚠️ Found multiple files (' + matchedFiles.length + '), returning first match');
    return { status: 'found', file: matchedFiles[0], searchUrl: searchUrl };
    
  } catch (e) {
    Logger.log('❌ Error finding original ad: ' + e);
    return { status: 'error', message: String(e) };
  }
}

/**
 * Save ClickUp link and update task name + hierarchy in one call
 */
function saveClickupLinkAndUpdateTask(clickupUrl) {
  try {
    Logger.log('🔧 Starting ClickUp integration...');
    Logger.log('   URL received: ' + clickupUrl);
    
    // Get the batch name from last created record
    const props = PropertiesService.getUserProperties();
    const raw = props.getProperty('lastCreatedRecord');
    
    Logger.log('   Last created record exists: ' + (!!raw));
    
    if (!raw) {
      Logger.log('❌ No last created record found');
      return { 
        status: 'error', 
        message: 'No batch found. Create folders first, then paste ClickUp URL.' 
      };
    }
    
    const record = JSON.parse(raw);
    Logger.log('   Record parsed: ' + JSON.stringify(record));
    
    // Extract batch name from adsUrl
    let batchName = '';
    if (record.urls && record.urls.adsUrl) {
      Logger.log('   adsUrl found: ' + record.urls.adsUrl);
      
      try {
        const match = record.urls.adsUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
        if (match) {
          Logger.log('   Folder ID extracted: ' + match[1]);
          const folder = DriveApp.getFolderById(match[1]);
          batchName = folder.getName();
          Logger.log('   ✅ Batch name: ' + batchName);
        } else {
          Logger.log('⚠️ No folder ID match in URL');
        }
      } catch (e) {
        Logger.log('⚠️ Could not get folder name: ' + e);
      }
    } else {
      Logger.log('⚠️ No adsUrl in record');
    }
    
    if (!batchName) {
      Logger.log('❌ Could not determine batch name');
      return { 
        status: 'error', 
        message: 'Could not determine batch folder name. Make sure you created the folders first.' 
      };
    }
    
    Logger.log('📝 Calling updateClickUpTaskHierarchy...');
    
    // ✅ Call the new hierarchy function instead of the old one
    const updateResult = updateClickUpTaskHierarchy(clickupUrl, batchName);
    
    Logger.log('📡 Update result: ' + JSON.stringify(updateResult));
    
    // Save the link regardless of update success
    record.urls = record.urls || {};
    record.urls.clickupUrl = clickupUrl;
    props.setProperty('lastCreatedRecord', JSON.stringify(record));
    
    Logger.log('💾 Saved ClickUp URL to record');
    
    // Return the update result
    if (updateResult.status === 'success') {
      // Add helpful details
      updateResult.message = batchName;  // Keep backward compatibility
      updateResult.displayMessage = updateResult.details || (batchName + ' - Main task + ' + (updateResult.subtasksRenamed || 0) + ' sub-tasks renamed');
    }
    
    return updateResult;
    
  } catch (e) {
    Logger.log('❌ Error in saveClickupLinkAndUpdateTask: ' + e);
    Logger.log('   Stack: ' + (e.stack || 'no stack'));
    return { 
      status: 'error', 
      message: String(e.message || e) 
    };
  }
}

/**
 * Get ClickUp task link from last created record
 */
function getClickupLinkFromRecord() {
  try {
    const props = PropertiesService.getUserProperties();
    const raw = props.getProperty('lastCreatedRecord');
    
    if (!raw) return '';
    
    const record = JSON.parse(raw);
    return (record.urls && record.urls.clickupUrl) || '';
  } catch (e) {
    return '';
  }
}

/**
 * Update ClickUp task name via API
 * @param {string} taskUrl - Full ClickUp task URL (supports multiple formats)
 * @param {string} newName - New task name (batch folder name)
 */
function updateClickUpTaskName(taskUrl, newName) {
  Logger.log('🚀 updateClickUpTaskName called');
  Logger.log('   URL: ' + taskUrl);
  Logger.log('   New name: ' + newName);
  
  if (!CLICKUP_ENABLED) {
    Logger.log('⚠️ ClickUp is disabled');
    return { 
      status: 'disabled', 
      message: 'ClickUp integration is disabled. Enable it by setting CLICKUP_ENABLED = true and adding your API key.' 
    };
  }
  
  if (!CLICKUP_API_KEY || CLICKUP_API_KEY === 'YOUR_CLICKUP_API_KEY_HERE') {
    Logger.log('⚠️ API key not configured');
    return { 
      status: 'error', 
      message: 'ClickUp API key not configured. Add your API key to CLICKUP_API_KEY constant.' 
    };
  }
  
  try {
    // Clean the URL
    taskUrl = taskUrl.trim();
    
    Logger.log('🔍 Parsing URL: ' + taskUrl);
    
    // Extract task ID - try multiple patterns
    let taskId = null;
    
    // Pattern 1: Full URL with workspace
    // https://app.clickup.com/t/WORKSPACE/TASK_ID
    let match = taskUrl.match(/\/t\/[^\/]+\/([a-zA-Z0-9]+)/);
    if (match) {
      taskId = match[1];
      Logger.log('✅ Pattern 1 matched (full URL): ' + taskId);
    }
    
    // Pattern 2: Short URL
    // https://app.clickup.com/t/TASK_ID
    if (!taskId) {
      match = taskUrl.match(/\/t\/([a-zA-Z0-9]+)(?:\/|$)/);
      if (match) {
        taskId = match[1];
        Logger.log('✅ Pattern 2 matched (short URL): ' + taskId);
      }
    }
    
    // Pattern 3: Most permissive - just grab anything after /t/
    if (!taskId) {
      match = taskUrl.match(/\/t\/([a-zA-Z0-9]+)/);
      if (match) {
        taskId = match[1];
        Logger.log('✅ Pattern 3 matched (permissive): ' + taskId);
      }
    }
    
    // Pattern 4: Just the task ID by itself
    if (!taskId && /^[a-zA-Z0-9]+$/.test(taskUrl)) {
      taskId = taskUrl;
      Logger.log('✅ Pattern 4 matched (raw ID): ' + taskId);
    }
    
    if (!taskId) {
      Logger.log('❌ Could not extract task ID from: ' + taskUrl);
      return { 
        status: 'error', 
        message: 'Could not extract task ID from URL.\n\nYour URL: ' + taskUrl + '\n\nTry copying the full URL from your browser address bar.'
      };
    }
    
    Logger.log('📝 Task ID: ' + taskId);
    Logger.log('📝 Updating to name: ' + newName);
    
    // Call ClickUp API
    const apiUrl = 'https://api.clickup.com/api/v2/task/' + taskId;
    Logger.log('🌐 API URL: ' + apiUrl);
    
    const options = {
      method: 'put',
      headers: {
        'Authorization': CLICKUP_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        name: newName
      }),
      muteHttpExceptions: true
    };
    
    Logger.log('📡 Sending request to ClickUp API...');
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('📡 Response Code: ' + responseCode);
    Logger.log('📡 Response Body: ' + responseText);
    
    if (responseCode === 200) {
      Logger.log('✅ SUCCESS! Task renamed to: ' + newName);
      return { 
        status: 'success', 
        message: newName,
        taskId: taskId
      };
    } else {
      Logger.log('❌ API Error: ' + responseCode);
      
      let errorMsg = 'Failed to update task';
      try {
        const errorData = JSON.parse(responseText);
        errorMsg = errorData.err || errorData.error || errorData.message || errorMsg;
      } catch (e) {
        errorMsg = responseText || errorMsg;
      }
      
      return { 
        status: 'error', 
        message: 'ClickUp API error (' + responseCode + '):\n' + errorMsg + '\n\nTask ID: ' + taskId
      };
    }
    
  } catch (e) {
    Logger.log('❌ Exception: ' + e);
    Logger.log('   Stack: ' + (e.stack || 'no stack'));
    return { 
      status: 'error', 
      message: 'Error: ' + (e.message || String(e)) 
    };
  }
}

/**
 * Update ClickUp task name + all sub-tasks (FIXED VERSION - gets sub-tasks correctly)
 */
function updateClickUpTaskHierarchy(taskUrl, newName) {
  Logger.log('🚀 updateClickUpTaskHierarchy called');
  Logger.log('   URL: ' + taskUrl);
  Logger.log('   New name: ' + newName);
  
  if (!CLICKUP_ENABLED) {
    return { 
      status: 'disabled', 
      message: 'ClickUp integration is disabled.' 
    };
  }
  
  const apiKey = CLICKUP_API_KEY;
  if (!apiKey || apiKey.includes('YOUR_CLICKUP_API_KEY')) {
    return { 
      status: 'error', 
      message: 'ClickUp API key not configured.' 
    };
  }
  
  try {
    taskUrl = taskUrl.trim();
    
    // Extract task ID
    let taskId = null;
    let match = taskUrl.match(/\/t\/[^\/]+\/([a-zA-Z0-9]+)/);
    if (match) taskId = match[1];
    if (!taskId) {
      match = taskUrl.match(/\/t\/([a-zA-Z0-9]+)(?:\/|$)/);
      if (match) taskId = match[1];
    }
    if (!taskId && /^[a-zA-Z0-9]+$/.test(taskUrl)) {
      taskId = taskUrl;
    }
    
    if (!taskId) {
      return { 
        status: 'error', 
        message: 'Could not extract task ID from URL.' 
      };
    }
    
    Logger.log('📝 Task ID extracted: ' + taskId);
    
    // ✅ FIX: Add query parameter to include subtasks
    const mainTaskUrl = 'https://api.clickup.com/api/v2/task/' + taskId + '?include_subtasks=true';
    Logger.log('🌐 Fetching main task from: ' + mainTaskUrl);
    
    const mainTaskResponse = UrlFetchApp.fetch(mainTaskUrl, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const mainTaskCode = mainTaskResponse.getResponseCode();
    Logger.log('📡 Response code: ' + mainTaskCode);
    
    if (mainTaskCode !== 200) {
      return { 
        status: 'error', 
        message: 'Cannot access task. API returned: ' + mainTaskCode 
      };
    }
    
    const mainTask = JSON.parse(mainTaskResponse.getContentText());
    Logger.log('✅ Main task found: ' + mainTask.name);
    
    // ✅ Check if subtasks are in the response now
    const subtasksList = mainTask.subtasks || [];
    Logger.log('📊 Main task has ' + subtasksList.length + ' sub-tasks');
    
    if (subtasksList.length > 0) {
      Logger.log('🔍 Sub-tasks details:');
      for (let i = 0; i < subtasksList.length; i++) {
        Logger.log('   Sub-task ' + (i+1) + ': ID=' + subtasksList[i].id + ', Name=' + subtasksList[i].name);
      }
    }
    
    // Step 2: Check for Packs (sub-sub-tasks)
    if (subtasksList.length > 0) {
      Logger.log('🔍 Checking for Packs (sub-sub-tasks)...');
      
      const firstSubtask = subtasksList[0];
      const subtaskCheckUrl = 'https://api.clickup.com/api/v2/task/' + firstSubtask.id + '?include_subtasks=true';
      const subtaskCheckResponse = UrlFetchApp.fetch(subtaskCheckUrl, {
        method: 'get',
        headers: {
          'Authorization': apiKey,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      });
      
      if (subtaskCheckResponse.getResponseCode() === 200) {
        const subtaskData = JSON.parse(subtaskCheckResponse.getContentText());
        
        if (subtaskData.subtasks && subtaskData.subtasks.length > 0) {
          Logger.log('❌ Pack detected (has sub-sub-tasks)');
          return {
            status: 'error',
            message: '🚫 This is a Pack (has sub-sub-tasks)\n\n' +
                     'This tool does not support Packs yet.\n\n' +
                     'Packs have: Main Task → Sub-tasks → Sub-sub-tasks\n' +
                     'This tool only works with: Main Task → Sub-tasks'
          };
        }
      }
    }
    
    // Step 3: Rename main task
    Logger.log('📝 Renaming main task to: ' + newName);
    const updateMainUrl = 'https://api.clickup.com/api/v2/task/' + taskId;
    const updateMainResponse = UrlFetchApp.fetch(updateMainUrl, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ name: newName }),
      muteHttpExceptions: true
    });
    
    const updateMainCode = updateMainResponse.getResponseCode();
    Logger.log('📡 Main task update response: ' + updateMainCode);
    
    if (updateMainCode !== 200) {
      Logger.log('❌ Main task update failed');
      return { 
        status: 'error', 
        message: 'Failed to update main task. API error: ' + updateMainCode 
      };
    }
    
    Logger.log('✅ Main task renamed successfully');
    
    // Step 4: Rename all sub-tasks
    let subtasksRenamed = 0;
    let subtaskErrors = [];
    
    if (subtasksList.length > 0) {
      Logger.log('📝 Starting to rename ' + subtasksList.length + ' sub-tasks...');
      
      for (let i = 0; i < subtasksList.length; i++) {
        const subtask = subtasksList[i];
        const subtaskUrl = 'https://api.clickup.com/api/v2/task/' + subtask.id;
        
        Logger.log('   🔄 Sub-task ' + (i+1) + '/' + subtasksList.length + ': ' + subtask.name);
        
        try {
          const updateSubResponse = UrlFetchApp.fetch(subtaskUrl, {
            method: 'put',
            headers: {
              'Authorization': apiKey,
              'Content-Type': 'application/json'
            },
            payload: JSON.stringify({ name: newName }),
            muteHttpExceptions: true
          });
          
          const subCode = updateSubResponse.getResponseCode();
          
          if (subCode === 200) {
            subtasksRenamed++;
            Logger.log('      ✅ Successfully renamed');
          } else {
            const errorMsg = 'Sub-task ' + (i+1) + ' failed with code ' + subCode;
            subtaskErrors.push(errorMsg);
            Logger.log('      ❌ ' + errorMsg);
          }
        } catch (e) {
          const errorMsg = 'Sub-task ' + (i+1) + ' exception: ' + e;
          subtaskErrors.push(errorMsg);
          Logger.log('      ❌ Exception: ' + e);
        }
      }
    } else {
      Logger.log('⚠️ No sub-tasks to rename');
    }
    
    Logger.log('📊 Final count: ' + subtasksRenamed + ' sub-tasks renamed successfully');
    
    // Step 5: Return success with details
    let message = newName;
    let details = '✅ Main task renamed';
    
    if (subtasksRenamed > 0) {
      details += '\n✅ ' + subtasksRenamed + ' sub-task(s) renamed';
    }
    
    if (subtaskErrors.length > 0) {
      details += '\n⚠️ ' + subtaskErrors.length + ' sub-task(s) failed';
    }
    
    return { 
      status: 'success', 
      message: message,
      taskId: taskId,
      subtasksRenamed: subtasksRenamed,
      details: details,
      totalSubtasks: subtasksList.length,
      errors: subtaskErrors.length > 0 ? subtaskErrors : null
    };
    
  } catch (e) {
    Logger.log('❌ Fatal exception: ' + e);
    return { 
      status: 'error', 
      message: 'Error: ' + (e.message || String(e)) 
    };
  }
}

/**
 * Admin-only: Delete a single document (for Notes Only flow)
 * @param {string} fileId - The document ID to delete
 * @returns {Object} Status object
 */
function deleteDocumentForDialog(fileId) {
  try {
    // Check admin permission
    if (!checkIfCurrentUserIsAdmin()) {
      return { status: 'error', message: 'Admin permission required' };
    }
    
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    
    Logger.log('🗑️ Admin deleting document: ' + fileName);
    
    // Move to trash
    file.setTrashed(true);
    
    Logger.log('✅ Document moved to trash');
    
    return { 
      status: 'deleted',
      message: 'Document "' + fileName + '" moved to trash'
    };
    
  } catch (e) {
    Logger.log('❌ Error deleting document: ' + e);
    return { 
      status: 'error', 
      message: e.message || String(e)
    };
  }
}

/**
 * ClickUp Auto-Fill: Rename tasks + update sub-task descriptions
 * @param {string} taskUrl - Main task URL
 * @param {string} newName - New task name
 * @param {Object} links - {docUrl, adsUrl, originalAdUrl, originalAdName}
 * @returns {Object} Result with status and details
 */
function clickUpAutoFill(taskUrl, newName, links) {
  Logger.log('🚀 ClickUp Auto-Fill started');
  Logger.log('   URL: ' + taskUrl);
  Logger.log('   New name: ' + newName);
  Logger.log('   Links: ' + JSON.stringify(links));
  
  if (!CLICKUP_ENABLED) {
    return { 
      status: 'disabled', 
      message: 'ClickUp integration is disabled.' 
    };
  }
  
  const apiKey = CLICKUP_API_KEY;
  if (!apiKey || apiKey.includes('YOUR_CLICKUP_API_KEY')) {
    return { 
      status: 'error', 
      message: 'ClickUp API key not configured.' 
    };
  }
  
  try {
    taskUrl = taskUrl.trim();
    
    // Extract task ID
    let taskId = null;
    let match = taskUrl.match(/\/t\/[^\/]+\/([a-zA-Z0-9]+)/);
    if (match) taskId = match[1];
    if (!taskId) {
      match = taskUrl.match(/\/t\/([a-zA-Z0-9]+)(?:\/|$)/);
      if (match) taskId = match[1];
    }
    if (!taskId && /^[a-zA-Z0-9]+$/.test(taskUrl)) {
      taskId = taskUrl;
    }
    
    if (!taskId) {
      return { 
        status: 'error', 
        message: 'Could not extract task ID from URL.' 
      };
    }
    
    Logger.log('📝 Task ID: ' + taskId);
    
    // Step 1: Get main task + subtasks
    const mainTaskUrl = 'https://api.clickup.com/api/v2/task/' + taskId + '?include_subtasks=true';
    const mainTaskResponse = UrlFetchApp.fetch(mainTaskUrl, {
      method: 'get',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (mainTaskResponse.getResponseCode() !== 200) {
      return { 
        status: 'error', 
        message: 'Cannot access task. Check URL and permissions.' 
      };
    }
    
    const mainTask = JSON.parse(mainTaskResponse.getContentText());
    const subtasksList = mainTask.subtasks || [];
    
    Logger.log('✅ Main task: ' + mainTask.name);
    Logger.log('📊 Found ' + subtasksList.length + ' sub-tasks');
    
    // Step 2: Check for Packs
    if (subtasksList.length > 0) {
      const firstSubtask = subtasksList[0];
      const subtaskCheckUrl = 'https://api.clickup.com/api/v2/task/' + firstSubtask.id + '?include_subtasks=true';
      const subtaskCheckResponse = UrlFetchApp.fetch(subtaskCheckUrl, {
        method: 'get',
        headers: {
          'Authorization': apiKey,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      });
      
      if (subtaskCheckResponse.getResponseCode() === 200) {
        const subtaskData = JSON.parse(subtaskCheckResponse.getContentText());
        if (subtaskData.subtasks && subtaskData.subtasks.length > 0) {
          Logger.log('❌ Pack detected');
          return {
            status: 'error',
            message: 'This is a Pack (has sub-sub-tasks). This tool does not support Packs yet.'
          };
        }
      }
    }
    
    // Step 3: Rename main task
    Logger.log('📝 Renaming main task...');
    const updateMainUrl = 'https://api.clickup.com/api/v2/task/' + taskId;
    const updateMainResponse = UrlFetchApp.fetch(updateMainUrl, {
      method: 'put',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ name: newName }),
      muteHttpExceptions: true
    });
    
    if (updateMainResponse.getResponseCode() !== 200) {
      return { 
        status: 'error', 
        message: 'Failed to rename main task.' 
      };
    }
    
    Logger.log('✅ Main task renamed');
    
    // Step 4: Build the description content (match "Copy for ClickUp" format exactly)
const descriptionLines = [];

// Match the exact format: "Label:\nURL"
if (links.docUrl) {
  descriptionLines.push('Ad Notes:');
  descriptionLines.push(links.docUrl);
}

if (links.adsUrl) {
  descriptionLines.push('Ads folder:');
  descriptionLines.push(links.adsUrl);
}

if (links.originalAdUrl) {
  descriptionLines.push('Batch / Ad:');
  descriptionLines.push(links.originalAdUrl);
}

const newDescriptionContent = descriptionLines.join('\n');
    Logger.log('📝 New description content prepared');
    
    // Step 5: Process each sub-task
    let subtasksRenamed = 0;
    let descriptionsUpdated = 0;
    let errors = [];
    
    if (subtasksList.length > 0) {
      Logger.log('📝 Processing ' + subtasksList.length + ' sub-tasks...');
      
      for (let i = 0; i < subtasksList.length; i++) {
        const subtask = subtasksList[i];
        const subtaskUrl = 'https://api.clickup.com/api/v2/task/' + subtask.id;
        
        Logger.log('   🔄 Sub-task ' + (i+1) + ': ' + subtask.name);
        
        try {
          // Get full subtask details
          const subtaskDetailResponse = UrlFetchApp.fetch(subtaskUrl, {
            method: 'get',
            headers: {
              'Authorization': apiKey,
              'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
          });
          
          if (subtaskDetailResponse.getResponseCode() !== 200) {
            errors.push('Sub-task ' + (i+1) + ': Could not fetch details');
            continue;
          }
          
          const subtaskDetail = JSON.parse(subtaskDetailResponse.getContentText());
          const currentDescription = subtaskDetail.description || subtaskDetail.text_content || '';
          
          Logger.log('      Current description length: ' + currentDescription.length);
          
          // Parse existing description to find "Creative: Name" line
          let newDescription = '';
          const lines = currentDescription.split('\n');
          let foundCreativeLine = false;
          let creativeLine = '';
          
          for (let j = 0; j < lines.length; j++) {
            const line = lines[j].trim();
            if (line.toLowerCase().startsWith('creative:')) {
              foundCreativeLine = true;
              creativeLine = lines[j];
              break;
            }
          }
          
          // Build new description
          if (foundCreativeLine) {
            newDescription = creativeLine + '\n\n' + newDescriptionContent;
            Logger.log('      ✅ Found Creative line, preserving it');
          } else {
            newDescription = newDescriptionContent;
            Logger.log('      ℹ️ No Creative line found, adding links only');
          }
          
          // Update subtask: rename + update description
          const updatePayload = {
            name: newName,
            description: newDescription
          };
          
          const updateSubResponse = UrlFetchApp.fetch(subtaskUrl, {
            method: 'put',
            headers: {
              'Authorization': apiKey,
              'Content-Type': 'application/json'
            },
            payload: JSON.stringify(updatePayload),
            muteHttpExceptions: true
          });
          
          const subCode = updateSubResponse.getResponseCode();
          
          if (subCode === 200) {
            subtasksRenamed++;
            descriptionsUpdated++;
            Logger.log('      ✅ Renamed + description updated');
          } else {
            errors.push('Sub-task ' + (i+1) + ': Update failed (code ' + subCode + ')');
            Logger.log('      ❌ Failed with code ' + subCode);
          }
          
        } catch (e) {
          errors.push('Sub-task ' + (i+1) + ': ' + e.message);
          Logger.log('      ❌ Exception: ' + e);
        }
      }
    }
    
    Logger.log('📊 Final results:');
    Logger.log('   Main task: renamed ✅');
    Logger.log('   Sub-tasks renamed: ' + subtasksRenamed);
    Logger.log('   Descriptions updated: ' + descriptionsUpdated);
    Logger.log('   Errors: ' + errors.length);
    
    // Build result message
    let details = 'Main task renamed';
    
    if (subtasksRenamed > 0) {
      details += '\n' + subtasksRenamed + ' sub-task(s) renamed';
    }
    
    if (descriptionsUpdated > 0) {
      details += '\n' + descriptionsUpdated + ' description(s) updated';
    }
    
    if (errors.length > 0) {
      details += '\n' + errors.length + ' error(s) occurred';
    }
    
    return { 
      status: 'success', 
      message: newName,
      taskId: taskId,
      subtasksRenamed: subtasksRenamed,
      descriptionsUpdated: descriptionsUpdated,
      details: details,
      errors: errors.length > 0 ? errors : null
    };
    
  } catch (e) {
    Logger.log('❌ Fatal exception: ' + e);
    return { 
      status: 'error', 
      message: 'Error: ' + (e.message || String(e)) 
    };
  }
}

/**
 * Save ClickUp link and do full auto-fill
 */
function saveClickupLinkAndAutoFill(clickupUrl) {
  try {
    Logger.log('🔧 Starting ClickUp Auto-Fill...');
    
    const props = PropertiesService.getUserProperties();
    const raw = props.getProperty('lastCreatedRecord');
    
    if (!raw) {
      return { 
        status: 'error', 
        message: 'No batch found. Create folders first, then use ClickUp Auto-Fill.' 
      };
    }
    
    const record = JSON.parse(raw);
    
    // Get batch name
    let batchName = '';
    if (record.urls && record.urls.adsUrl) {
      try {
        const match = record.urls.adsUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
        if (match) {
          const folder = DriveApp.getFolderById(match[1]);
          batchName = folder.getName();
        }
      } catch (e) {
        Logger.log('⚠️ Could not get folder name: ' + e);
      }
    }
    
    if (!batchName) {
      return { 
        status: 'error', 
        message: 'Could not determine batch folder name.' 
      };
    }
    
    // Prepare links object
    const links = {
      docUrl: record.urls.docUrl || '',
      adsUrl: record.urls.adsUrl || '',
      originalAdUrl: record.urls.originalAdUrl || '',
      originalAdName: record.urls.originalAdName || ''
    };
    
    Logger.log('📝 Calling ClickUp Auto-Fill...');
    
    // Call the auto-fill function
    const result = clickUpAutoFill(clickupUrl, batchName, links);
    
    // Save the link
    if (result.status === 'success') {
      record.urls = record.urls || {};
      record.urls.clickupUrl = clickupUrl;
      props.setProperty('lastCreatedRecord', JSON.stringify(record));
      Logger.log('💾 Saved ClickUp URL to record');
    }
    
    return result;
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    return { 
      status: 'error', 
      message: String(e.message || e) 
    };
  }
}

/**
 * Test ClickUp URL parsing (temporary debug function)
 */
function testClickUpUrl() {
  const testUrl = 'https://app.clickup.com/t/86c6h3903';
  
  Logger.log('Testing URL: ' + testUrl);
  
  // Test pattern 1
  const pattern1 = testUrl.match(/\/t\/[^\/]+\/([a-zA-Z0-9]+)/);
  Logger.log('Pattern 1 (with workspace): ' + (pattern1 ? pattern1[1] : 'NO MATCH'));
  
  // Test pattern 2
  const pattern2 = testUrl.match(/\/t\/([a-zA-Z0-9]+)/);
  Logger.log('Pattern 2 (short URL): ' + (pattern2 ? pattern2[1] : 'NO MATCH'));
  
  // Test pattern 3 (everything after /t/)
  const pattern3 = testUrl.match(/\/t\/(.+?)(?:\/|$)/);
  Logger.log('Pattern 3 (flexible): ' + (pattern3 ? pattern3[1] : 'NO MATCH'));
  
  // Show what we got
  if (pattern2) {
    Logger.log('✅ Extracted task ID: ' + pattern2[1]);
  }
}

function getCreateDialogDiagnosticsFast() {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31 = sh.getRange('C31').getDisplayValue().trim();
  const c3 = sh.getRange('C3').getDisplayValue().trim();
  const c16 = sh.getRange('C16').getDisplayValue().trim();

  const p31 = parseNameParts(c31);
  const pC3 = parseNameParts(c3);
  const pC16 = parseNameParts(c16);

  const domain = parseDomainFromC31(c31);
  const isFb = (p31.platform === 'FB');
  const locale = p31.locale;

  const productMap = PRODUCT_LOCALE_FOLDERS[domain] || {};
  const hasBaseDirect = !!extractAnyDriveId(productMap[locale] || '');

  const mismatches = [];
  const pushIf = (field, a, b) => {
    if (!a && !b) return;
    if (String(a || '').toUpperCase() !== String(b || '').toUpperCase()) {
      mismatches.push({ field, c31: a || '—', src: b || '—' });
    }
  };
  pushIf('platform', p31.platform, srcParts.platform);
  pushIf('locale', p31.locale, srcParts.locale);
  pushIf('batch', p31.batch, srcParts.batch);
  pushIf('adType', p31.adType, srcParts.adType);
  pushIf('product', p31.product, srcParts.product);

  return {
    isFb,
    platformToken: p31.platform || '',
    domain: domain || '',
    locale: locale || '',
    productLabel: productShortLabel_(domain), // << "A33 CleanlixMold"
    hasBase: hasBaseDirect,
    missingLocale: (!hasBaseDirect && !!domain && !!locale),
    consistency: mismatches.length ? { mismatches, sourceField } : null
  };
}

function createBatchFolderOnlyForDialog() {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31 = (sh.getRange('C31').getDisplayValue() || '').trim();
  
  if (!c31) {
    return { status: 'error', message: 'C31 is empty.' };
  }
  
  const parts = parseNameParts(c31);
  const domain = parts.product;
  const locale = parts.locale;
  const batch = parts.batch;
  const year = detectTargetYear(c31);

  if (!domain || !locale || !batch) {
    // ✅ ADD THIS:
    logAnalyticsEventSafe_({
      action: 'Folder Only',
      product: domain || '',
      locale: locale || '',
      batchId: batch || '',
      success: false,
      error: 'Could not parse product / locale / batch ID from C31'
    });
    
    return { status: 'error', message: 'Could not parse product / locale / batch ID from C31.' };
  }

  const base = resolveBaseFolderForProductLocale(domain, locale);
  if (!base) {
    // ✅ ADD THIS:
    logAnalyticsEventSafe_({
      action: 'Folder Only',
      product: domain || '',
      locale: locale || '',
      batchId: batch || '',
      success: false,
      error: 'Could not find base folder for ' + domain + ' / ' + locale
    });
    
    return { status: 'error', message: 'Could not find base folder for ' + domain + ' / ' + locale };
  }

  // ✅ Check for batch ID duplicates using the same logic as Standard Setup
  const fullName = sanitizeBatchName(c31).clean;
  const dupByBatch = getDuplicateBatchByIdPayload_Flexible_(base, year, fullName);
  
  if (dupByBatch && dupByBatch.status === 'exists') {
    return {
      status: 'batchDup',
      batchId: dupByBatch.batchId,
      existingFolderId: dupByBatch.existing.folderId || '',
      existingFolderUrl: dupByBatch.existing.folderUrl || '',
      yearFolderUrl: toDriveUrl_(pickScanContainer_(base, year)) || '',
      yearFolderId: pickScanContainer_(base, year).getId() || '',
      meta: { domain, locale, year, c31Name: fullName }
    };
  }

  // No conflict → create folder
  const container = getOrCreateYearContainerFlexible_(base, year);
  const folder = container.createFolder(c31);
  folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  
  // ✅ NEW: Create subfolders
  getOrCreateSubfolder(folder, 'Ad Notes');
  getOrCreateSubfolder(folder, 'Work Files');

  const result = {
    status: 'created',
    url: '',
    adsUrl: folder.getUrl(),
    adNotesUrl: '',
    message: 'Folder created successfully.'
  };

  // ✅ NEW: Try to find original ad for iterations (even though we're not creating a doc)
  try {
    if (c31.toUpperCase().includes('ITR')) {
      const c16 = sh.getRange('C16').getDisplayValue().trim();
      const parsed = parseOriginalAdId(c16);
      
      if (parsed.hasOrig) {
        const origLocale = parsed.origId.split('_')[0];
        const baseFolder = resolveBaseFolderForProductLocale(parsed.product, origLocale);
        
        if (baseFolder) {
          const searchResult = findOriginalAdFile(parsed.origId, parsed.product, parsed.adType, baseFolder);
          
          if (searchResult.status === 'found') {
            let fileUrl;
            if (typeof searchResult.file.getName === 'function') {
              fileUrl = searchResult.file.getUrl();
            } else {
              fileUrl = searchResult.file.url || searchResult.file.fileUrl || '';
            }
            
            if (fileUrl) {
              result.originalAdUrl = fileUrl;
              result.originalAdName = (typeof searchResult.file.getName === 'function') 
                ? searchResult.file.getName() 
                : (searchResult.file.name || 'Original Ad');
              Logger.log('✅ Found original ad for Folder Only: ' + result.originalAdName);
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('⚠️ Could not find original ad for Folder Only: ' + e);
  }

  // ✅ Save to record (including original ad if found)
  const recordData = {
    docUrl: '',
    adsUrl: result.adsUrl,
    adNotesUrl: ''
  };
  
  if (result.originalAdUrl) {
    recordData.originalAdUrl = result.originalAdUrl;
    recordData.originalAdName = result.originalAdName;
  }

  saveLastCreatedRecord_('folder', recordData);

  // ✅ ADD THIS:
  logAnalyticsEventSafe_({
    action: 'Folder Only',
    product: domain,
    locale: locale,
    batchId: batch,
    success: true
  });

  return result;
}

function createAdNotesAutoRoute() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();

  try {
    const fullName = sh.getRange('C31').getDisplayValue().trim();
    if (!fullName) throw new Error('C31 is empty');

    // detect domain + locale from C31
    const p31 = parseNameParts(fullName);
    const domain = parseDomainFromC31(fullName);
    const locale = p31.locale;
    if (!domain || !locale) throw new Error('Could not detect product or locale from C31.');

    // use direct map OR auto-discover sibling locale folder; do NOT create when missing
    const baseFolder = resolveBaseFolderForProductLocale(domain, locale);
    if (!baseFolder) {
      throw new Error(
        '🟡 Missing base folder for ' + domain + ' / ' + locale +
        '. Add it in PRODUCT_LOCALE_FOLDERS or create a folder named "' + locale +
        '" under the product’s Ads folder, then rerun.'
      );
    }

    // locale → year → <C31> → Ad Notes
    const yearFolder = getOrCreateYearFolder(baseFolder, fullName); // <-- pass fullName!
    const batchFolder = getOrCreateSubfolder(yearFolder, fullName);
    const adNotesFolder = getOrCreateSubfolder(batchFolder, 'Ad Notes');

    // create the doc inside Ad Notes (default: populate ad name)
    const url = createAdNotesCoreOpt(adNotesFolder.getId(), true);

    // open
    const html = '<script>window.open("' + url + '","_blank");google.script.host.close();</script>';
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(10).setHeight(10), 'Opening…');

  } catch (e) {
    ui.alert('Create Ad Notes (Standard Setup) failed:\n' + (e && e.message ? e.message : String(e)));
  }
}

function getAutoBaseFolderFor(domain, locale) {
  const byProduct = PRODUCT_LOCALE_FOLDERS[domain];
  if (byProduct && byProduct[locale]) return byProduct[locale];

  // optional fallback if you kept LOCALE_BASE_FOLDERS
  if (typeof LOCALE_BASE_FOLDERS !== 'undefined' && LOCALE_BASE_FOLDERS[locale]) {
    return LOCALE_BASE_FOLDERS[locale];
  }
  return '';
}

function buildCreateDialogHtml(initialDiagJson) {
  const sh = SpreadsheetApp.getActiveSheet();
  const c31 = sh.getRange('C31').getDisplayValue() || '';
  
  return `
<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <title>Create</title>
    <style>
  @keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}
  @keyframes fadeIn {
    from { opacity: 0; transform: translateY(-10px); }
    to { opacity: 1; transform: translateY(0); }
  }
  @keyframes slideUp {
    from { opacity: 0; transform: translateY(20px); }
    to { opacity: 1; transform: translateY(0); }
  }

  :root{
    --shadow-sm: 0 2px 6px rgba(16,17,20,.08);
    --bg:#fff; --fg:#101114; --sub:#5c667a; --muted:#7a8294;
    --card:#fff; --border:#e6e9ee; --shadow:0 6px 24px rgba(16,17,20,.08);
    --primary:#1a73e8; --primary-ink:#fff; --primary-ghost:#eaf2fe;
    --success:#16a34a; --success-light:#f0fdf4; --success-border:#bbf7d0;
    --danger:#ef4444;
    --ok-bg:#ecfdf3; --ok-bd:#c6f6d5; --ok-ink:#14532d;
    --radius:14px;
  }
  *{box-sizing:border-box}
  html,body{height:100%}
  body{font-family:Inter,system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:0;padding:18px;background:var(--bg);color:var(--fg);overflow:auto}
  h1{margin:6px 0 6px;font-size:23px;font-weight:800}
  p.lead{margin:0 0 12px;color:var(--sub);font-size:13px}
  .mini{font-size:12px;color:var(--muted)}
  .link{color:var(--primary);cursor:pointer;text-decoration:underline;font-size:14px}

  .alert{display:none;margin:10px 0 8px;padding:12px 14px;border-radius:12px;border:1px solid transparent;font-size:13px;line-height:1.5}
  .alert.show{display:block;animation:slideUp 0.3s ease;}
  .alert.ok{background:var(--ok-bg);border-color:var(--ok-bd);color:var(--ok-ink)}
  .alert.soft{background:#eef6ff;border:1px solid #cfe3ff;color:#0b54a6}
  .alert.danger{background:#fef2f2;border:1px solid #fecaca;color:#991b1b}

  .choice{margin-top:10px;padding:12px;border:1px dashed var(--border);border-radius:12px}
  .choice .label{font-weight:700;margin-bottom:6px;font-size:13px}
  .rad{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:10px}
  .rad:hover{background:#fafbff}
  .rad input[type="radio"]{appearance:none;width:18px;height:18px;border-radius:50%;border:2px solid #c7cfda;outline:none;cursor:pointer;flex:0 0 auto}
  .rad input[type="radio"]:checked{border-color:var(--primary);background-image:radial-gradient(var(--primary) 0 50%,transparent 52% 100%);background-position:center;background-repeat:no-repeat}
  .rad .txt{font-size:13px;color:#222}
  .muted{font-size:12px;color:var(--muted);margin-top:6px}

  .grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin:12px 0 8px}
  .grid.four-col {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 12px;
    margin: 12px 0 8px;
  }
  .card-btn{border:1px solid var(--border);background:var(--card);color:var(--fg);border-radius:var(--radius);padding:14px 12px;text-align:center;cursor:pointer;box-shadow:var(--shadow);transition:transform .08s ease,border-color .2s ease,box-shadow .2s ease;outline:none}
  .card-btn.loading{position: relative;opacity: .92;}
  .card-btn.loading::after{
    content:'';position:absolute;right:12px;top:12px;width:16px;height:16px;
    border:2px solid #c7cfda;border-top-color:#1a73e8;border-radius:50%;
    animation:spin .9s linear infinite;
  }
  .card-btn:hover{transform:translateY(-1px)}
  .card-btn:disabled{opacity:.5;cursor:not-allowed;transform:none}
  .card-btn .title{font-size:14px;font-weight:800;margin:0 0 6px}
  .card-btn .subtitle{font-size:12px;color:var(--muted);margin:0}
  .primary{background:linear-gradient(180deg,var(--primary),calc(100% - 1px),#0f5fd6);color:var(--primary-ink);border:none}
  .primary .subtitle{color:rgba(255,255,255,.92)}

  .exist{display:none;margin-top:10px}
  .exist.show{display:block}
  .exist-card{background:#fffefa;border:1px solid #f1e8c7;border-radius:14px;padding:12px;box-shadow:0 6px 24px rgba(16,17,20,.06)}
  .exist-title{font-weight:900;margin:0 0 6px;font-size:15px}
  .exist-text{font-size:13px;color:#5c667a;margin:0 0 8px}
  .btn-line{display:flex;gap:8px;flex-wrap:wrap}
  .btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;padding:8px 12px;border-radius:10px;border:1px solid #d1d5db;background:#fff;cursor:pointer;font-weight:700;font-size:13px}
  .btn.primary{background:#16a34a;color:#fff;border:none}
  .btn.ghost{background:#fff;color:#1f2937}
  .btn.danger{background:#fee2e2;border-color:#fecaca;color:#b91c1c}
  .btn:disabled{opacity:.6;cursor:not-allowed}

  /* ✨ IMPROVED DONE BOX STYLES */
.done-card {
  display: none;
  background: linear-gradient(135deg, #f0fdf4 0%, #ecfdf5 100%);
  border: 1px solid var(--success-border);
  border-radius: 16px;
  padding: 20px;
  margin-top: 20px;
  box-shadow: 0 8px 24px rgba(22, 163, 74, 0.08);
  animation: slideUp 0.4s ease;
}
  
  .done-card.show {
    display: block;
  }

  .done-card .done-title {
    margin: 0 0 16px;
    font-weight: 700;
    font-size: 15px;
    color: #166534;
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .done-row {
    display: flex;
    flex-direction: column;
    gap: 12px;
    margin-bottom: 16px;
  }

  .pair {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 2px;
  }

  .open-btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: white;
    border: 1px solid #d1d5db;
    border-radius: 10px;
    font-weight: 600;
    font-size: 13px;
    color: #111827;
    padding: 10px 16px;
    text-decoration: none;
    transition: all 0.15s ease;
    flex: 1;
    min-width: 0;
  }

  .open-btn:hover {
    background: #f9fafb;
    border-color: #16a34a;
    transform: translateX(2px);
  }

  .copy-icon-btn {
  background: white;
  border: 1px solid #e5e7eb;
  border-radius: 10px;
  padding: 10px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  transition: all 0.15s ease;
  flex-shrink: 0;
  position: relative; /* ✅ Added for checkmark positioning */
}

  .copy-icon-btn:hover {
    background: #f0fdf4;
    border-color: #16a34a;
    transform: scale(1.05);
  }

  .copy-icon-btn:active {
    transform: scale(0.95);
  }

  /* ✨ Copy button success feedback */
  .copy-icon-btn.copied {
  background: #f0fdf4;
  border-color: #16a34a;
  animation: copySuccess 0.3s ease;
  }

  /* Change the emoji when copied */
  .copy-icon-btn.copied::after {
  content: '✅';
  position: absolute;
  font-size: 14px;
  }

  @keyframes copySuccess {
  0%, 100% { transform: scale(1); }
  50% { transform: scale(1.1); }
  }

  .clickup-section {
    border-top: 1px solid #bbf7d0;
    padding-top: 12px;
    margin-top: 8px;
  }

  .clickup-btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: white;
    border: 1px solid #d1d5db;
    border-radius: 10px;
    font-weight: 600;
    font-size: 13px;
    padding: 10px 16px;
    cursor: pointer;
    color: #111827;
    transition: all 0.15s ease;
    width: 100%;
    justify-content: center;
  }

  .clickup-btn:hover {
    background: #f9fafb;
    border-color: #7b68ee;
    transform: translateY(-1px);
  }

  .clickup-icon {
  width: 18px;
  height: 18px;
  object-fit: contain; /* ✅ Add this line to prevent stretching */
  vertical-align: middle;
}

  .copy-status {
    display: none;
    color: #16a34a;
    font-size: 12px;
    font-weight: 600;
    margin-top: 8px;
    text-align: center;
    animation: fadeIn 0.3s ease;
  }

  .copy-status.show {
    display: block;
  }

  /* ✅ Admin delete button */
  .admin-delete-btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
    border: 1px solid #fca5a5;
    border-radius: 10px;
    font-weight: 600;
    font-size: 12px;
    padding: 8px 14px;
    cursor: pointer;
    color: #991b1b;
    transition: all 0.15s ease;
    width: 100%;
    justify-content: center;
  }
  
  .admin-delete-btn:hover {
    background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%);
    border-color: #f87171;
    transform: translateY(-1px);
  }
  
  .admin-delete-btn:active {
    transform: translateY(0);
  }

    /* ✨ SEPARATE CLICKUP RENAME CARD */
  .clickup-rename-card {
    display: none;
    background: linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%);
    border: 1px solid #e9d5ff;
    border-radius: 16px;
    padding: 20px;
    margin-top: 16px;
    box-shadow: 0 4px 16px rgba(139, 92, 246, 0.08);
    animation: slideUp 0.4s ease;
  }
  
  .clickup-rename-card.show {
    display: block;
  }
  
  .clickup-rename-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 16px;
  }
  
  .clickup-header-icon {
    width: 32px;
    height: 32px;
    object-fit: contain;
  }
  
  .clickup-rename-title {
    font-size: 14px;
    font-weight: 700;
    color: #7e22ce;
    margin: 0;
  }
  
  .clickup-rename-subtitle {
    font-size: 11px;
    color: #9333ea;
    opacity: 0.8;
    margin: 2px 0 0 0;
  }
  
  .clickup-rename-input-row {
    display: flex;
    gap: 10px;
    align-items: center;
  }
  
  .clickup-input {
    flex: 1;
    padding: 11px 14px;
    border: 2px solid #e9d5ff;
    border-radius: 10px;
    font-size: 13px;
    font-family: inherit;
    background: white;
    transition: all 0.2s ease;
  }
  
  .clickup-input:focus {
    outline: none;
    border-color: #a855f7;
    box-shadow: 0 0 0 3px rgba(168, 85, 247, 0.1);
  }
  
  .clickup-input::placeholder {
    color: #c084fc;
    opacity: 0.6;
  }
  
  .clickup-rename-btn {
    padding: 11px 20px;
    background: linear-gradient(135deg, #a855f7 0%, #9333ea 100%);
    color: white;
    border: none;
    border-radius: 10px;
    font-weight: 700;
    font-size: 13px;
    cursor: pointer;
    white-space: nowrap;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    transition: all 0.2s ease;
    box-shadow: 0 2px 8px rgba(168, 85, 247, 0.25);
  }
  
  .clickup-rename-btn:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(168, 85, 247, 0.35);
  }
  
  .clickup-rename-btn:active {
    transform: translateY(0);
  }
  
  .clickup-rename-btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
    transform: none;
  }
  
  .btn-emoji {
    font-size: 14px;
  }

  /* ✨ IMPROVED REOPEN BUTTON */
  .reopen-container {
    position: absolute;
    top: 18px;
    right: 22px;
  }

  .reopen-btn {
    background: linear-gradient(135deg, #f9fafb 0%, #f3f4f6 100%);
    border: 1px solid #d1d5db;
    border-radius: 10px;
    color: #374151;
    font-size: 12px;
    font-weight: 600;
    padding: 8px 14px;
    cursor: pointer;
    transition: all 0.2s ease;
    box-shadow: 0 2px 4px rgba(0,0,0,0.04);
  }

  .reopen-btn:hover {
    background: linear-gradient(135deg, #ffffff 0%, #f9fafb 100%);
    border-color: var(--primary);
    color: var(--primary);
    transform: translateY(-1px);
    box-shadow: 0 4px 8px rgba(26, 115, 232, 0.1);
  }

  .reopen-btn:active {
    transform: translateY(0);
  }

  /* ✨ IMPROVED MODAL */
.done-modal {
  display: none;
  position: fixed;
  z-index: 9999;
  left: 0;
  top: 0;
  width: 100%;
  height: 100%;
  background: rgba(16, 17, 20, 0.1); /* ✅ Subtle backdrop - change to 0 if you want none */
  backdrop-filter: blur(2px); /* ✅ Light blur - change to none if you want none */
  align-items: center;
  justify-content: center;
  animation: fadeIn 0.2s ease;
}

  .done-modal-content {
    background: linear-gradient(135deg, #ffffff 0%, #fafbfc 100%);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(17, 24, 39, 0.2);
    padding: 0;
    width: 480px;
    max-width: 90vw;
    position: relative;
    animation: slideUp 0.3s ease;
    border: 1px solid #e5e7eb;
    overflow: hidden;
  }

  .done-modal-header {
    background: linear-gradient(135deg, #f0fdf4 0%, #ecfdf5 100%);
    border-bottom: 1px solid #bbf7d0;
    padding: 20px 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .done-modal-header .done-title {
    margin: 0;
    font-weight: 700;
    font-size: 16px;
    color: #166534;
    display: flex;
    align-items: center;
    gap: 10px;
  }

  .done-close {
    border: none;
    background: rgba(255, 255, 255, 0.8);
    font-size: 18px;
    color: #6b7280;
    cursor: pointer;
    transition: all 0.15s ease;
    width: 32px;
    height: 32px;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 0;
  }

  .done-close:hover {
    background: white;
    color: #111827;
    transform: rotate(90deg);
  }

  .done-modal-body {
    padding: 24px;
  }

  .done-modal-body .done-row {
    margin-bottom: 20px;
  }

  .done-modal-body .clickup-section {
    border-top: 1px solid #e5e7eb;
    padding-top: 16px;
    margin-top: 16px;
  }

  /* ✅ NEW: Andromeda template selection styles */
      .template-choice {
        display: none;
        margin-top: 16px;
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        border-left: 4px solid #3b82f6;
        padding: 14px 16px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(59, 130, 246, 0.08);
      }
      
      .template-choice.show {
        display: block;
      }
      
      .template-label {
        display: flex;
        align-items: center;
        gap: 8px;
        color: #1e40af;
        margin-bottom: 10px;
        font-weight: 700;
        font-size: 13px;
      }
      
      .template-rad {
        background: #fff;
        border: 1px solid #bfdbfe;
        border-radius: 8px;
        margin-bottom: 8px;
        padding: 10px;
      }
      
      .template-rad:hover {
        background: #eff6ff;
      }

</style>
  </head>
  <body>
    <div class="mini" style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
      <span><strong>Create Ad Notes</strong></span>
      <span class="link" id="refresh">Refresh consistency</span>
    </div>

    <h1>What do you want to <strong>create</strong>?</h1>
    <p class="lead">Choose a flow, then confirm how to handle the Ad Name.</p>

    <div id="loadingAnim" style="display:none;margin-bottom:8px;font-size:13px;color:#5c667a">
      <span class="spinner" style="display:inline-block;width:12px;height:12px;border:2px solid #ccc;border-top-color:#1a73e8;border-radius:50%;animation:spin 1s linear infinite;vertical-align:middle;margin-right:6px"></span>
      Checking naming consistency…
    </div>

    <div id="platformAlert" class="alert danger"></div>
    <div id="warnConsistency" class="alert soft"></div>
    <div id="okConsistency" class="alert ok"></div>

    <div class="choice" id="gate">
      <div class="label">Ad Name Options:</div>
      <label class="rad" for="optYes">
        <input type="radio" id="optYes" name="adname" />
        <div class="txt"><strong>Auto-fill Ad Name</strong> — Use the ad name from this sheet (C3 if C31 has “NEW”; C16 if it has “Itr”).</div>
      </label>
      <label class="rad" for="optNo">
        <input type="radio" id="optNo" name="adname" />
        <div class="txt"><strong>Leave Ad Name Blank</strong> — Keep the template placeholder unchanged.</div>
      </label>
      <div class="muted">Select one option to enable the actions below.</div>
    </div>

<!-- ✅ NEW: Original Ad Options (only shown for iteration ads) -->
<div class="choice" id="origAdGate" style="display:none;margin-top:16px;background:linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%);border-left:4px solid #a855f7;padding:14px 16px;border-radius:12px;box-shadow:0 2px 8px rgba(168,85,247,0.08)">
  <div class="label" style="display:flex;align-items:center;gap:8px;color:#7e22ce;margin-bottom:10px">
    <span style="font-size:16px">🔗</span>
    <strong id="origAdLabel">How should we handle the Original Ad field?</strong>
  </div>
  <label class="rad" for="origAuto" style="background:#fff;border:1px solid #e9d5ff;border-radius:8px;margin-bottom:8px;padding:10px">
    <input type="radio" id="origAuto" name="origad" checked />
    <div class="txt">
      <div style="display:flex;align-items:center;gap:6px">
        <span style="font-size:14px">🤖</span>
        <strong style="color:#7e22ce">Auto-fill</strong>
        <span style="color:#64748b">— Find and link the original ad automatically</span>
      </div>
      <div style="font-size:11px;color:#64748b;margin-top:6px;margin-left:24px">Best when iterating your own product's ads</div>
    </div>
  </label>
  <label class="rad" for="origBlank" style="background:#fff;border:1px solid #e9d5ff;border-radius:8px;padding:10px">
    <input type="radio" id="origBlank" name="origad" />
    <div class="txt">
      <div style="display:flex;align-items:center;gap:6px">
        <span style="font-size:14px">✍️</span>
        <strong style="color:#7e22ce">Leave blank</strong>
        <span style="color:#64748b">— I'll fill it manually</span>
      </div>
      <div style="font-size:11px;color:#64748b;margin-top:6px;margin-left:24px">Use this when taking inspiration from another product</div>
    </div>
  </label>
</div>

<!-- ✅ NEW: Andromeda Template Selection (only shown for New VID ads) -->
<div class="choice template-choice" id="andromedaGate">
  <div class="template-label">
    <span style="font-size:16px">🎬</span>
    <strong>Video Template Selection</strong>
  </div>
  <label class="rad template-rad" for="templateStandard">
    <input type="radio" id="templateStandard" name="videotemplate" checked />
    <div class="txt">
      <div style="display:flex;align-items:center;gap:6px">
        <span style="font-size:14px">📄</span>
        <strong style="color:#1e40af">Standard Brief</strong>
        <span style="color:#64748b">— Classic video ad template</span>
      </div>
      <div style="font-size:11px;color:#64748b;margin-top:6px;margin-left:24px">Use for most video ads</div>
    </div>
  </label>
  <label class="rad template-rad" for="templateAndromeda">
    <input type="radio" id="templateAndromeda" name="videotemplate" />
    <div class="txt">
      <div style="display:flex;align-items:center;gap:6px">
        <img src="https://registry.npmmirror.com/@lobehub/icons-static-png/latest/files/dark/meta-color.png" 
             alt="Meta" 
             style="width:16px;height:16px;object-fit:contain;">
        <strong style="color:#1e40af">Andromeda Brief</strong>
        <span style="color:#64748b">— Meta's Andromeda structure</span>
      </div>
      <div style="font-size:11px;color:#64748b;margin-top:6px;margin-left:24px">Follows Meta's Andromeda algorithm testing requirements</div>
    </div>
  </label>
</div>

<!-- 🔁 Reopen button moved out of the grid -->
<div class="reopen-container">
  <button id="reopenLastBtn" class="reopen-btn">⟳ Reopen Last Created</button>
</div>

<!-- Main actions stay clean and centered -->
<div class="grid four-col">
  <button id="bAuto" class="card-btn primary" disabled>
    <div class="title">📁 Standard Setup</div>
    <p class="subtitle">Creates Ad Notes and folders automatically</p>
  </button>
  <button id="bCustom" class="card-btn" disabled>
    <div class="title">🧭 Custom Setup</div>
    <p class="subtitle">Select Ad Notes folder manually (creates Batch folder correctly)</p>
  </button>
  <button id="bNotesOnly" class="card-btn" disabled>
    <div class="title">📝 Notes Only</div>
    <p class="subtitle">Creates the document only (no folders)</p>
  </button>
  <button id="bFolderOnly" class="card-btn" disabled>
    <div class="title">📂 Folder Only</div>
    <p class="subtitle">Just creates the batch folder (no Ad Notes)</p>
  </button>
</div>

<!-- ✅ Improved inline Done box (under the 4 buttons) -->
<div id="doneBox" class="done-card">
  <div class="done-title">Done! Your Ad Notes were created.</div>

  <div class="done-row">
    <!-- Row for Ad Notes doc -->
    <div class="pair" id="docRow" style="display:none;">
      <a id="openDocBtn" class="open-btn" href="#" target="_blank">
        📄 Open Ad Notes doc
      </a>
      <button class="copy-icon-btn" id="copyDocBtn" title="Copy link">
        📋
      </button>
    </div>

    <!-- Row for Ads folder -->
    <div class="pair" id="adsRow" style="display:none;">
      <a id="openAdsBtn" class="open-btn" href="#" target="_blank">
        🗂️ Open Ads folder
      </a>
      <button class="copy-icon-btn" id="copyAdsBtn" title="Copy link">
        📋
      </button>
    </div>
  </div>

  <div class="clickup-section">
    <button id="copyClickupBtn" class="clickup-btn">
      <img src="https://1000logos.net/wp-content/uploads/2022/06/ClickUp-Emblem.png"
     alt="ClickUp" class="clickup-icon">
      Copy for ClickUp
    </button>
    <p id="copyStatus" class="copy-status">Copied ✅</p>
  </div>
  
  <!-- ✅ Admin delete button (hidden by default) -->
  <div id="adminDeleteSection" style="display:none;border-top:1px solid #bbf7d0;padding-top:12px;margin-top:8px;">
    <button id="adminDeleteBtn" class="admin-delete-btn">
      🗑️ Delete Test Batch (Admin Only)
    </button>
  </div>
</div>

  <!-- ✅ NEW: Separate ClickUp Rename Section -->
<div id="clickupRenameBox" class="clickup-rename-card">
  <div class="clickup-rename-header">
    <img src="https://1000logos.net/wp-content/uploads/2022/06/ClickUp-Emblem.png" alt="ClickUp" class="clickup-header-icon">
    <div>
      <div class="clickup-rename-title">Auto-Rename ClickUp Task</div>
      <div class="clickup-rename-subtitle">Paste your task URL to automatically update its name</div>
    </div>
  </div>
  
  <!-- ✅ ADD THIS WARNING BOX -->
  <div style="background:#fef3c7;border:1px solid #fcd34d;border-radius:8px;padding:10px;margin:12px 0;font-size:11px;line-height:1.5">
    <div style="font-weight:700;color:#92400e;margin-bottom:4px">⚠️ Important:</div>
    <div style="color:#78350f">
      • Copy the <strong>MAIN (mother) task</strong> link, not a sub-task<br>
      • Does NOT work with <strong>Packs</strong> (tasks with sub-sub-tasks)<br>
      • Will rename: Main task + all direct sub-tasks
    </div>
  </div>

  <div class="clickup-rename-input-row">
    <input 
      type="text" 
      id="clickupLinkInput" 
      placeholder="https://app.clickup.com/t/86c6h3903 or just task ID..."
      class="clickup-input"
    />
    <button id="saveClickupLinkBtn" class="clickup-rename-btn">
      <span class="btn-emoji">🚀</span>
      <span>Rename</span>
    </button>
  </div>
</div>

<!-- ✅ IMPROVED: Auto-Fill Section (better UI) -->
<div id="clickupAutoFillBox" class="clickup-rename-card" style="display:none;margin-top:16px;background:linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%);border:1px solid #e9d5ff">
  
  <!-- ✅ Header with ClickUp icon (not a button look) -->
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
    <img src="https://1000logos.net/wp-content/uploads/2022/06/ClickUp-Emblem.png" alt="ClickUp" style="width:32px;height:32px;object-fit:contain;">
    <div>
      <div style="font-size:14px;font-weight:700;color:#7e22ce;margin:0">ClickUp Auto-Fill</div>
      <div style="font-size:11px;color:#9333ea;opacity:0.8;margin:2px 0 0 0">Renames tasks + adds Drive links to descriptions</div>
    </div>
  </div>
  
  <!-- Warning box -->
  <div style="background:#fef3c7;border:1px solid #fcd34d;border-radius:8px;padding:10px;margin-bottom:12px;font-size:11px;line-height:1.5">
    <div style="font-weight:700;color:#92400e;margin-bottom:4px">⚡ What this does:</div>
    <div style="color:#78350f">
      • Renames main task + all sub-tasks<br>
      • Adds Drive links to sub-task descriptions<br>
      • Preserves "Creative: Name" line in descriptions<br>
      • Does NOT work with Packs
    </div>
  </div>
  
  <!-- Input row -->
  <div class="clickup-rename-input-row">
    <input 
      type="text" 
      id="clickupAutoFillInput" 
      placeholder="Paste MAIN task URL here..."
      class="clickup-input"
      style="border-color:#e9d5ff"
    />
    <button id="clickupAutoFillBtn" class="clickup-rename-btn" style="background:linear-gradient(135deg, #a855f7, #9333ea);border:none">
      <span class="btn-emoji">⚡</span>
      <span>Auto-Fill</span>
    </button>
  </div>
  
  <!-- Status messages -->
  <div id="clickupAutoFillStatus" style="display:none;margin-top:12px;padding:10px;border-radius:8px;font-size:12px;line-height:1.6"></div>
</div>

    <!-- Duplicate panel -->
    <div id="existBox" class="exist">
      <div class="exist-card">
        <div class="exist-title">⚠️ A batch already exists</div>
        <div class="exist-text">We found an existing path for this <code>&lt;C31&gt;</code>.</div>
        <div class="btn-line" style="margin-top:6px">
          <button id="existPrimary" class="btn primary">Open existing doc</button>
          <button id="openAdNotes" class="btn ghost">Open Ad Notes folder</button>
          <button id="openBatch" class="btn ghost">Open batch folder</button>
          <button id="override" class="btn danger">Override (delete existing doc)</button>
          <button id="cancelExist" class="btn ghost">Cancel</button>
        </div>
        <div style="margin-top:6px">
          <span id="createAnother" class="link" style="display:none">Create another doc</span>
        </div>
      </div>
    </div>

    <!-- Duplicate Batch ID Warning -->
<div id="dupBox" style="display:none; margin-top:16px;">
  <div style="
    background:#fff8e1;
    border:1px solid #fcd34d;
    border-radius:14px;
    padding:12px;
    box-shadow:0 6px 24px rgba(16,17,20,.06);
    font-size:13px;
    line-height:1.4;
    color:#5c667a;">
    
    <p style="margin:0 0 8px;font-weight:700;color:#92400e;">
      ⚠ This Batch ID already exists in this product / locale / year.
    </p>

    <p style="margin:0 0 8px;">
      Someone already used this Batch ID folder.
      You can:
    </p>

    <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:8px;">
      <button id="dupOpenExisting" class="btn ghost">Open existing batch</button>
      <button id="dupOpenYear" class="btn ghost">See all batch IDs in this year</button>
      <button id="dupOverride" class="btn danger">
        Replace existing batch
      </button>
      <button id="dupCancel" class="btn ghost">Cancel</button>
    </div>

    <div style="font-size:12px;color:#9ca3af;">
      “Replace existing batch” will delete the current batch folder (Ads, Ad Notes, Work Files, etc.)
      and recreate it fresh with your current C31.
      The old folder will be moved to Google Drive Trash, so you can still recover it later.
    </div>
  </div>
</div>

<!-- ✅ Improved popup modal for "Reopen Last Created" -->
<div id="doneModal" class="done-modal">
  <div class="done-modal-content">
    <div class="done-modal-header">
      <div class="done-title">✅ Last created assets</div>
      <button class="done-close" id="closeDoneBtn">✕</button>
    </div>

    <div class="done-modal-body">
      <div class="done-row">
        <div class="pair" id="docRow_modal" style="display:none;">
          <a id="openDocBtn_modal" class="open-btn" href="#" target="_blank">
            📄 Open Ad Notes doc
          </a>
          <button class="copy-icon-btn" id="copyDocBtn_modal" title="Copy link">
            📋
          </button>
        </div>

        <div class="pair" id="adsRow_modal" style="display:none;">
          <a id="openAdsBtn_modal" class="open-btn" href="#" target="_blank">
            🗂️ Open Ads folder
          </a>
          <button class="copy-icon-btn" id="copyAdsBtn_modal" title="Copy link">
            📋
          </button>
        </div>
      </div>

      <div class="clickup-section">
        <button id="copyClickupBtn_modal" class="clickup-btn">
          <img src="https://1000logos.net/wp-content/uploads/2022/06/ClickUp-Emblem.png"
     alt="ClickUp" class="clickup-icon">
          Copy for ClickUp
        </button>
        <p id="copyStatus_modal" class="copy-status">Copied ✅</p>
      </div>

      <!-- ✅ Admin delete button for modal -->
      <div id="adminDeleteSection_modal" style="display:none;border-top:1px solid #e5e7eb;padding-top:16px;margin-top:16px;">
        <button id="adminDeleteBtn_modal" class="admin-delete-btn">
          🗑️ Delete Test Batch (Admin Only)
        </button>
      </div>
    </div>
  </div>
</div>




    <script>

      let lastBatchDup = null;

      // 👇 add this helper for this dialog
      function showLoading(on){
        var el = document.getElementById('loadingAnim');
        if (!el) return;
        el.style.display = on ? 'block' : 'none';
      }

      (function(){
        const $ = id => document.getElementById(id);
        const optYes = $('optYes'), optNo = $('optNo');
        const bAuto = $('bAuto'), bCustom = $('bCustom'), bNotes = $('bNotesOnly'), bFolder = $('bFolderOnly');
        
        const okConsistency = $('okConsistency'), warnConsistency = $('warnConsistency');
        const refresh = $('refresh');

        const existBox = $('existBox'), existPrimary=$('existPrimary'),
              openAdNotes=$('openAdNotes'), openBatch=$('openBatch'),
              overrideBtn=$('override'), cancelExist=$('cancelExist'),
              createAnother=$('createAnother');

const doneBox = $('doneBox'),
      openDocBtn = $('openDocBtn'),
      openAdsBtn = $('openAdsBtn');
const copyDocBtn = $('copyDocBtn'),
      copyAdsBtn = $('copyAdsBtn'),
      copyClickupBtn = $('copyClickupBtn'),
      copyStatus = $('copyStatus');

        let running = false, lastExist = null, checkDone = false, lastActionType = 'standard';
        function safe(fn){ try{ if (typeof fn === 'function') fn(); } catch(e){} }

function markBusy(btn, on){
  if (!btn) return;
  if (on) btn.classList.add('loading');
  else btn.classList.remove('loading');
}
function clearAllBusy(){
  markBusy(bAuto,false);
  markBusy(bCustom,false);
  markBusy(bNotes,false);
  markBusy(bFolder, false);
}

function recomputeDisables(){
  if (!bAuto || !bCustom || !bNotes || !bFolder) return;
  if (running){
    [bAuto, bCustom, bNotes, bFolder].forEach(el => el && (el.disabled = true));
    return;
  }
  const hasChoice = !!(optYes && optYes.checked) || !!(optNo && optNo.checked);
  const allow = hasChoice && checkDone;
  [bAuto, bCustom, bNotes, bFolder].forEach(el => el && (el.disabled = !allow));
}


        function updateGate(){ recomputeDisables(); }
        optYes && optYes.addEventListener('change', updateGate);
        optNo  && optNo .addEventListener('change', updateGate);
        recomputeDisables();

function renderConsistency(c) {
  const platformAlert = document.getElementById('platformAlert');
  if (platformAlert) platformAlert.classList.remove('show');
  if (!okConsistency || !warnConsistency) return;

  // If nothing received yet
  if (!c) {
    okConsistency.classList.remove('show');
    warnConsistency.classList.remove('show');
    if (platformAlert) platformAlert.classList.remove('show');
    return;
  }

  // ✅ NEW: TVStick/Flixy → TellyStick rename warning (HIGHEST PRIORITY)
  if (c.renameWarning) {
    warnConsistency.innerHTML = c.renameWarning.message;
    warnConsistency.classList.add('show');
    okConsistency.classList.remove('show');
    if (platformAlert) platformAlert.classList.remove('show');
    return; // Don't show other checks if there's a rename warning
  }

  // 🚫 Platform check
  if (platformAlert && c.partsC31 && c.partsC31.platform) {
    const platform = String(c.partsC31.platform || '').toUpperCase();
    if (platform && platform !== 'FB' && platform !== 'META') {
      platformAlert.innerHTML =
        '🚫 <strong>Platform not supported:</strong> This tool works only with <code>FB / Meta</code> naming. Detected <code>' +
        platform +
        '</code> in C31.';
      platformAlert.classList.add('show');
    } else {
      platformAlert.classList.remove('show');
    }
  }

  // ✅ No mismatches → all good banner
  if (!c.mismatches || !c.mismatches.length) {
    okConsistency.innerHTML =
      '✅ <strong>All good:</strong> The folder and ad naming are consistent (' +
      (c.sourceField || 'C3/C16') +
      ' matches C31).';
    okConsistency.classList.add('show');
    warnConsistency.classList.remove('show');
    return;
  }

  // ⚠️ Mismatches → info banner
  okConsistency.classList.remove('show');

  const labels = {
    product: 'Product',
    platform: 'Platform',
    locale: 'Locale',
    batch: 'Batch ID',
    adType: 'Ad Type',
  };

  let lines = '';
  c.mismatches.forEach(function (m) {
    const label = labels[m.field] || m.field;
    lines +=
      '<li style="font-weight:400;">' +
      '<span style="font-weight:600;">' + label + ':</span> ' +
      'C31 = <code>' + String(m.c31 || '—') + '</code>, ' +
      String(c.sourceField || 'C3/C16') + ' = <code>' + String(m.src || '—') + '</code>' +
      '</li>';
  });

  warnConsistency.innerHTML =
    '<p style="margin-bottom:8px;">' +
    'ℹ️ <strong>There’s a mismatch between the folder name and the ad name.</strong> ' +
    'Check the details below and fix the ad name if needed — the folders will still be created based on <code>C31</code>.' +
    '</p>' +
    '<ul>' + lines + '</ul>';

  warnConsistency.classList.add('show');
}

        // ✅ NEW: Andromeda template detection and permission check
        function checkAndromedaFeature() {
          console.log('🚀 checkAndromedaFeature() called');
          
          google.script.run
            .withSuccessHandler(function(res) {
              console.log('📡 Server response received:', res);
              
              const hasPermission = res.hasPermission;
              const isNewVid = res.isNewVid;
              
              console.log('   hasPermission:', hasPermission);
              console.log('   isNewVid:', isNewVid);
              
              const andromedaGate = document.getElementById('andromedaGate');
              console.log('   andromedaGate element:', andromedaGate);
              
              if (hasPermission && isNewVid && andromedaGate) {
                andromedaGate.classList.add('show');
                console.log('✨ Andromeda template selection shown');
              } else {
                console.log('ℹ️ Andromeda not shown');
                console.log('   Reason: permission=' + hasPermission + ', isNewVid=' + isNewVid + ', element=' + !!andromedaGate);
              }
            })
            .withFailureHandler(function(err) {
              console.log('⚠️ Failed to check Andromeda feature:', err);
            })
            .checkAndromedaFeature();
        }

        // ✅ Fetch consistency and check for special features
        function fetchConsistency(){
          console.log('🎬 fetchConsistency() called');
          showLoading(true);
          checkDone = false;
          recomputeDisables();
          
          google.script.run
            .withSuccessHandler(function(res) {
              console.log('📥 Got response from server');
              showLoading(false);
              checkDone = true;
              renderConsistency(res);
              recomputeDisables();
              
              // ✅ Check for iteration/localization
              const c31Value = res.c31Value || '';
              console.log('📥 C31 value:', c31Value);
              
              const parenMatch = c31Value.match(/\(([^)]+)\)/);
              const insideParens = parenMatch ? parenMatch[1].toUpperCase() : '';
              const isIteration = /^(ITR|ITER|LOC|LOCAL)/i.test(insideParens);
              
              console.log('🔍 Is iteration:', isIteration);
              
              const origAdGate = document.getElementById('origAdGate');
              
              if (origAdGate && isIteration) {
                origAdGate.style.display = 'block';
                console.log('✅ Showing iteration options');
                
                const origAdLabel = document.getElementById('origAdLabel');
                if (origAdLabel) {
                  const c31Upper = c31Value.toUpperCase();
                  if (c31Upper.includes('LOC')) {
                    origAdLabel.textContent = 'How should we handle the Original Ad field? (Localization)';
                  } else {
                    origAdLabel.textContent = 'How should we handle the Original Ad field? (Iteration)';
                  }
                }
                
                // Pre-fetch original ad
                google.script.run
                  .withSuccessHandler(function(prefetchRes){
                    if(prefetchRes && prefetchRes.cached){
                      console.log('✅ Original ad pre-fetched:', prefetchRes.result.status);
                    }
                  })
                  .withFailureHandler(function(e){
                    console.log('⚠️ Pre-fetch failed:', e);
                  })
                  .prefetchOriginalAdIfIteration();
              }
              
              // ✅ NEW: Show Andromeda instantly (no extra server call)
              if (res.andromeda) {
                const hasPermission = res.andromeda.hasPermission;
                const isNewVid = res.andromeda.isNewVid;
                
                console.log('✨ Andromeda check (instant):', hasPermission, isNewVid);
                
                const andromedaGate = document.getElementById('andromedaGate');
                
                if (hasPermission && isNewVid && andromedaGate) {
                  andromedaGate.classList.add('show');
                  console.log('✅ Andromeda shown instantly');
                }
              }
            })
            .withFailureHandler(function(err) {
              console.log('❌ fetchConsistency failed:', err);
              showLoading(false);
              checkDone = true;
              recomputeDisables();
            })
            .getNameConsistencyDiagnostics();
        }
        
        fetchConsistency(); // ✅ Call it once

        safe(()=>{
          const refresh=document.getElementById('refresh');
          if(!refresh)return;
          refresh.onclick = function(){
  showLoading(true);
  checkDone = false;
  recomputeDisables();
  google.script.run
    .withSuccessHandler(res => {
      showLoading(false);
      checkDone = true;
      renderConsistency(res);
      recomputeDisables();
    })
    .withFailureHandler(() => {
      showLoading(false);
      checkDone = true;
      recomputeDisables();
    })
    .getNameConsistencyDiagnostics();
};

        });

        function showExistPanel(payload){
          lastExist = payload;
          const hasDoc = !!(payload.existing.docUrl);
          const hasAdNotes = !!(payload.existing.adNotesUrl);
          const hasBatch = !!(payload.existing.folderUrl);
          if (existPrimary){
            existPrimary.textContent = hasDoc ? 'Open existing doc'
                                  : hasAdNotes ? 'Create doc in Ad Notes'
                                  : 'Create “Ad Notes” + Doc here';
          }
          openAdNotes && (openAdNotes.style.display = hasAdNotes ? 'inline-flex' : 'none');
          openBatch && (openBatch.style.display = hasBatch ? 'inline-flex' : 'none');
          createAnother && (createAnother.style.display = hasDoc ? 'inline' : 'none');
          overrideBtn && (overrideBtn.style.display = hasDoc ? 'inline-flex' : 'none');
          existBox && existBox.classList.add('show');
        }

        function handleServerError(err){
          clearAllBusy();
          running = false; recomputeDisables();
          alert((err && err.message) ? err.message : String(err || 'Something went wrong.'));
        }
        function attachCopy(btn, url){
          if (!btn || !url) return;
          btn.onclick = async () => {
            try{ await navigator.clipboard.writeText(url); btn.classList.add('copied'); setTimeout(()=>btn.classList.remove('copied'), 600); }
            catch(e){ alert('Could not copy to clipboard.'); }
          };
        }

        function handleServerResult(res){
  clearAllBusy();
  running = false; 
  recomputeDisables();

  if (!res) return;

  if (res.status === 'cancelled') return;

  if (res.status === 'error'){
    alert(res.message || 'Something went wrong.');
    return;
  }

  if (res.status === 'batchDup'){
    // show dupBox with actions
    lastBatchDup = res; // store globally so buttons know what to open/override
    if (dupBox) {
      dupBox.style.display = 'block';
    }
    return;
  }

  if (res.status === 'exists'){
    showExistPanel(res); // you already have this
    return;
  }

if (res.status === 'created') {
  const urls = {
    docUrl: res.url || '',
    adsUrl: res.adsUrl || '',
    adNotesUrl: res.adNotesUrl || res.folderUrl || '',
    originalAdUrl: res.originalAdUrl || '',      // ✅ ADD THIS
    originalAdName: res.originalAdName || ''     // ✅ ADD THIS
  };
  
  console.log('📦 Result from server:', res);
  console.log('📦 URLs object passed to render:', urls);
  
  // Show the done box with animation
  if (doneBox) {
    doneBox.classList.add('show');  // ✅ Changed from style.display = 'block'
  }
  
  // Render buttons properly
  renderDoneBoxByAction(lastActionType, urls);

  // ✅ NEW: Show ClickUp sections if user has permission
  google.script.run
    .withSuccessHandler(function(hasPermission) {
      const clickupRenameBox = document.getElementById('clickupRenameBox');
      if (hasPermission && clickupRenameBox) {
        clickupRenameBox.classList.add('show');
      }
      
      // ✅ Also show Auto-Fill box if user has that permission
      if (userHasAutoFillPermission) {
        const autoFillBox = document.getElementById('clickupAutoFillBox');
        if (autoFillBox) {
          autoFillBox.style.display = 'block';
        }
      }
    })
    .checkClickUpPermission();

  // Scroll into view for visibility
  setTimeout(function() {
    var el = document.getElementById('doneBox');
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }, 300);

  // Hide all other boxes
  if (existBox) existBox.classList.remove('show');
  const dupBox = document.getElementById('dupBox');
  if (dupBox) dupBox.style.display = 'none';
  
  return;
}


}

// === Duplicate batch warning button logic ===
const dupBox       = document.getElementById('dupBox');
const dupOpenExisting = document.getElementById('dupOpenExisting');
const dupOpenYear     = document.getElementById('dupOpenYear');
const dupOverride      = document.getElementById('dupOverride');
const dupCancel        = document.getElementById('dupCancel');

if (dupOpenExisting)
  dupOpenExisting.onclick = () => {
    if (lastBatchDup && lastBatchDup.existingFolderUrl)
      window.open(lastBatchDup.existingFolderUrl, '_blank');
  };

if (dupOpenYear)
  dupOpenYear.onclick = () => {
    if (lastBatchDup && lastBatchDup.yearFolderUrl)
      window.open(lastBatchDup.yearFolderUrl, '_blank');
  };

if (dupOverride)
  dupOverride.onclick = () => {
    if (!lastBatchDup || !lastBatchDup.existingFolderId) return;
    dupBox.style.display = 'none';
    running = true; recomputeDisables();
    markBusy(bAuto, true);
    google.script.run
      .withSuccessHandler(handleServerResult)
      .withFailureHandler(handleServerError)
      .nukeExistingBatchAndRecreate_(lastBatchDup.existingFolderId, !!(optYes && optYes.checked));
  };

if (dupCancel)
  dupCancel.onclick = () => {
    dupBox.style.display = 'none';
    running = false;
    recomputeDisables();
  };

bAuto && (bAuto.onclick = function(){
  lastActionType = 'standard';
  const f = !!(optYes && optYes.checked);
  
  // ✅ NEW: Capture original ad choice
  const origBlank = document.getElementById('origBlank');
  const origAdMode = (origBlank && origBlank.checked) ? 'blank' : 'auto';
  
  // ✅ NEW: Capture template choice
  const templateAndromeda = document.getElementById('templateAndromeda');
  const videoTemplate = (templateAndromeda && templateAndromeda.checked) ? 'andromeda' : 'standard';
  
  running = true;
  recomputeDisables();
  markBusy(bAuto, true);

  // Hide all old panels before starting
  const dupBox = document.getElementById('dupBox');
  if (dupBox) dupBox.style.display = 'none';
  if (existBox) existBox.classList.remove('show');
  if (doneBox) doneBox.classList.remove('show');

  google.script.run
    .withSuccessHandler(handleServerResult)
    .withFailureHandler(handleServerError)
    .attemptStandardSetupWithBatchCheck(f, origAdMode, videoTemplate); // ✅ Pass videoTemplate
});

bCustom && (bCustom.onclick = function(){
  lastActionType = 'custom';
  const raw = prompt('Paste a Drive FOLDER link/ID (or a FILE link/ID — parent will be used):','');
  if (!raw) return;
  const f = !!(optYes && optYes.checked);
  
  // ✅ NEW: Capture original ad choice
  const origBlank = document.getElementById('origBlank');
  const origAdMode = (origBlank && origBlank.checked) ? 'blank' : 'auto';
  
  // ✅ NEW: Capture template choice
  const templateAndromeda = document.getElementById('templateAndromeda');
  const videoTemplate = (templateAndromeda && templateAndromeda.checked) ? 'andromeda' : 'standard';
  
  running = true;
  recomputeDisables();
  markBusy(bCustom, true);

  // Hide all old panels before starting
  const dupBox = document.getElementById('dupBox');
  if (dupBox) dupBox.style.display = 'none';
  if (existBox) existBox.classList.remove('show');
  if (doneBox) doneBox.classList.remove('show');

  google.script.run
    .withSuccessHandler(handleServerResult)
    .withFailureHandler(handleServerError)
    .attemptCustomSetupWithBatchCheck(raw, f, origAdMode, videoTemplate); // ✅ Pass videoTemplate
});


bNotes && (bNotes.onclick = function(){
  lastActionType = 'notes';
  const f = !!(optYes && optYes.checked);
  
  // ✅ NEW: Capture original ad choice
  const origBlank = document.getElementById('origBlank');
  const origAdMode = (origBlank && origBlank.checked) ? 'blank' : 'auto';
  
  // ✅ NEW: Capture template choice
  const templateAndromeda = document.getElementById('templateAndromeda');
  const videoTemplate = (templateAndromeda && templateAndromeda.checked) ? 'andromeda' : 'standard';
  
  running = true;
  recomputeDisables();
  markBusy(bNotes, true);

  // ✅ Hide all old panels before starting
  const dupBox = document.getElementById('dupBox');
  if (dupBox) dupBox.style.display = 'none';
  if (existBox) existBox.classList.remove('show');
  if (doneBox) doneBox.classList.remove('show');

  google.script.run
    .withSuccessHandler(handleServerResult)
    .withFailureHandler(handleServerError)
    .createAdNotesNotesOnlyForDialog(f, origAdMode, videoTemplate); // ✅ Pass videoTemplate
});


bFolder && (bFolder.onclick = function(){
  lastActionType = 'folder';
  running = true;
  recomputeDisables();
  markBusy(bFolder, true);

  // ✅ Hide all old panels before starting
  const dupBox = document.getElementById('dupBox');
  if (dupBox) dupBox.style.display = 'none';
  if (existBox) existBox.classList.remove('show');
  if (doneBox) doneBox.classList.remove('show');

  google.script.run
    .withSuccessHandler(handleServerResult)
    .withFailureHandler(handleServerError)
    .createBatchFolderOnlyForDialog();
});

        document.addEventListener('keydown', function(e){
          if (e.key === '1' && bAuto && !bAuto.disabled) bAuto.click();
          if (e.key === '2' && bCustom && !bCustom.disabled) bCustom.click();
          if (e.key === '3' && bNotes && !bNotes.disabled) bNotes.click();
          if (e.key === 'Escape'){ google.script.host.close(); }
        });


// ✅ Modal behavior for "Reopen Last Created"
const reopenBtn = document.getElementById('reopenLastBtn');

const doneModal = document.getElementById('doneModal');
const closeDoneBtn = document.getElementById('closeDoneBtn');

if (reopenBtn) {
  reopenBtn.onclick = function() {
    google.script.run.withSuccessHandler(function(record) {
      if (!record || !record.urls) {
        alert("No recent record found.");
        return;
      }
      
      // ✅ Pass deleted flag if present
      const urls = record.urls;
      if (record.deleted) {
        urls.deleted = true;
        urls.deletedAt = record.deletedAt;
      }
      
      renderDoneBoxByAction(record.type || 'standard', urls);
      doneModal.style.display = 'flex';
      
      // ✅ NEW: Check permission AFTER showing modal
      google.script.run
        .withSuccessHandler(function(hasPermission) {
          if (!hasPermission) {
            const modalSection = document.querySelector('#doneModal #clickupLinkSection');
            if (modalSection) modalSection.style.display = 'none';
          }
        })
        .checkClickUpPermission();
        
    }).getLastCreatedRecord();
  };
}

if (closeDoneBtn) {
  closeDoneBtn.onclick = function() {
    doneModal.style.display = 'none';
  };
}

// close modal if you click outside
window.onclick = function(e) {
  if (e.target === doneModal) {
    doneModal.style.display = 'none';
  }
};

// ✅ Check if user has ClickUp permission
google.script.run
  .withSuccessHandler(function(hasPermission) {
    const clickupSection = document.getElementById('clickupLinkSection');
    if (!hasPermission && clickupSection) {
      clickupSection.style.display = 'none';
    }
  })
  .checkClickUpPermission();

// ✅ NEW: ClickUp task link functionality
const clickupLinkInput = document.getElementById('clickupLinkInput');
const saveClickupLinkBtn = document.getElementById('saveClickupLinkBtn');

if (saveClickupLinkBtn && clickupLinkInput) {
  saveClickupLinkBtn.onclick = function() {
    const url = clickupLinkInput.value.trim();
    
    if (!url) {
      alert('Please paste a ClickUp task URL or task ID.');
      return;
    }
    
    // Accept URLs OR just task IDs
    const isUrl = url.includes('clickup.com');
    const isTaskId = /^[a-zA-Z0-9]+$/.test(url);
    
    if (!isUrl && !isTaskId) {
      alert('Please enter a valid ClickUp task URL or task ID.\\n\\nExamples:\\n• https://app.clickup.com/t/86c6h3903\\n• 86c6h3903');
      return;
    }
    
    saveClickupLinkBtn.disabled = true;
    saveClickupLinkBtn.textContent = '⏳ Updating...';
    
    google.script.run
      .withSuccessHandler(function(result) {
        saveClickupLinkBtn.disabled = false;
        
        console.log('📡 Server response:', result);
        
        if (result.status === 'success') {
          if (clickupLinkInput) clickupLinkInput.value = '';
          
          saveClickupLinkBtn.textContent = '✅ Renamed!';
          saveClickupLinkBtn.style.background = '#16a34a';
          
          alert('✅ Success!\\n\\nClickUp task renamed to:\\n' + result.message);
          
          setTimeout(function() {
            saveClickupLinkBtn.textContent = '💾 Rename Task';
            saveClickupLinkBtn.style.background = '#7b68ee';
          }, 2000);
          
        } else if (result.status === 'disabled') {
          alert('⚠️ ClickUp Integration Disabled\\n\\n' + result.message + '\\n\\nCheck the script configuration.');
          saveClickupLinkBtn.textContent = '💾 Rename Task';
          
        } else {
          // Show actual server error
          alert('❌ Failed to update ClickUp task:\\n\\n' + result.message);
          saveClickupLinkBtn.textContent = '💾 Rename Task';
        }
      })
      .withFailureHandler(function(err) {
        console.error('❌ Client error:', err);
        saveClickupLinkBtn.disabled = false;
        saveClickupLinkBtn.textContent = '💾 Rename Task';
        alert('Error: ' + (err.message || String(err)));
      })
      .saveClickupLinkAndUpdateTask(url);
  };
}
     // ✅ Check Auto-Fill permission (but don't show yet - wait for batch creation)
// We store the permission for later use
var userHasAutoFillPermission = false;

google.script.run
  .withSuccessHandler(function(hasAutoFillPermission) {
    userHasAutoFillPermission = hasAutoFillPermission;
  })
  .withFailureHandler(function(err) {
    // Silent fail
  })
  .checkClickUpAutoFillPermission();

// ✅ Auto-Fill button handler
var clickupAutoFillBtn = document.getElementById('clickupAutoFillBtn');
var clickupAutoFillInput = document.getElementById('clickupAutoFillInput');

if (clickupAutoFillBtn && clickupAutoFillInput) {
  clickupAutoFillBtn.onclick = function() {
    var url = clickupAutoFillInput.value.trim();
    
    if (!url) {
      alert('Please paste a ClickUp main task URL.');
      return;
    }
    
    var isUrl = url.indexOf('clickup.com') !== -1;
    var isTaskId = /^[a-zA-Z0-9]+$/.test(url);
    
    if (!isUrl && !isTaskId) {
      alert('Please enter a valid ClickUp task URL or task ID.');
      return;
    }
    
    clickupAutoFillBtn.disabled = true;
    clickupAutoFillBtn.innerHTML = '<span class="btn-emoji">⏳</span><span>Processing...</span>';
    
    var statusDiv = document.getElementById('clickupAutoFillStatus');
    if (statusDiv) statusDiv.style.display = 'none';
    
    google.script.run
      .withSuccessHandler(function(result) {
        clickupAutoFillBtn.disabled = false;
        
        if (result && result.status === 'success') {
          if (clickupAutoFillInput) clickupAutoFillInput.value = '';
          
          clickupAutoFillBtn.innerHTML = '<span class="btn-emoji">✅</span><span>Complete!</span>';
          clickupAutoFillBtn.style.background = 'linear-gradient(135deg, #16a34a, #15803d)';
          
          if (statusDiv) {
            statusDiv.style.display = 'block';
            statusDiv.style.background = '#f0fdf4';
            statusDiv.style.border = '1px solid #bbf7d0';
            statusDiv.style.color = '#166534';
            
            var detailsText = result.details || 'Success';
            var html = '<div style="font-weight:700;margin-bottom:8px">ClickUp Auto-Fill Complete!</div>';
            html += '<div style="font-size:11px">' + detailsText + '</div>';
            
            if (result.descriptionsUpdated && result.descriptionsUpdated > 0) {
              html += '<div style="margin-top:8px;padding:8px;background:#d1fae5;border-radius:6px;font-size:11px">';
              html += 'Drive links added to ' + result.descriptionsUpdated + ' sub-task description(s)';
              html += '</div>';
            }
            
            statusDiv.innerHTML = html;
          }
          
          setTimeout(function() {
            clickupAutoFillBtn.innerHTML = '<span class="btn-emoji">⚡</span><span>Auto-Fill</span>';
            clickupAutoFillBtn.style.background = 'linear-gradient(135deg, #10b981, #059669)';
          }, 3000);
          
        } else if (result && result.status === 'disabled') {
          if (statusDiv) {
            statusDiv.style.display = 'block';
            statusDiv.style.background = '#fef3c7';
            statusDiv.style.border = '1px solid #fcd34d';
            statusDiv.style.color = '#92400e';
            statusDiv.innerHTML = 'ClickUp integration is disabled.';
          }
          clickupAutoFillBtn.innerHTML = '<span class="btn-emoji">⚡</span><span>Auto-Fill</span>';
          
        } else {
          if (statusDiv) {
            statusDiv.style.display = 'block';
            statusDiv.style.background = '#fef2f2';
            statusDiv.style.border = '1px solid #fecaca';
            statusDiv.style.color = '#991b1b';
            var errMsg = (result && result.message) ? result.message : 'Unknown error';
            statusDiv.innerHTML = errMsg;
          }
          clickupAutoFillBtn.innerHTML = '<span class="btn-emoji">⚡</span><span>Auto-Fill</span>';
        }
      })
      .withFailureHandler(function(err) {
        clickupAutoFillBtn.disabled = false;
        clickupAutoFillBtn.innerHTML = '<span class="btn-emoji">⚡</span><span>Auto-Fill</span>';
        
        var statusDiv = document.getElementById('clickupAutoFillStatus');
        if (statusDiv) {
          statusDiv.style.display = 'block';
          statusDiv.style.background = '#fef2f2';
          statusDiv.style.border = '1px solid #fecaca';
          statusDiv.style.color = '#991b1b';
          var errText = (err && err.message) ? err.message : String(err || 'Unknown error');
          statusDiv.innerHTML = 'Error: ' + errText;
        }
      })
      .saveClickupLinkAndAutoFill(url);
  };
}   
      })();
    </script>
  </body>
</html>`;
}

/* ========================================
   📊 ANALYTICS SYSTEM
   ======================================== */

/**
 * Main logging function - called after every action
 * Now handles missing user permissions gracefully
 */
function logAnalyticsEvent_(eventData) {
  if (!ANALYTICS_ENABLED) return;
  
  try {
    // ✅ Safe email access
    let currentUser = '';
    try {
      currentUser = Session.getActiveUser().getEmail();
    } catch (emailError) {
      currentUser = 'anonymous@commercecore.com';
      Logger.log('⚠️ Could not get user email, using anonymous');
    }
    
    // Skip logging for admins if TRACK_ADMINS is false
    if (!TRACK_ADMINS && ADMIN_EMAILS.includes(currentUser)) {
      return;
    }
    
    // Add user to event data
    eventData.user = currentUser;
    
    // Log to Google Sheet (detailed)
    logToSheet_(eventData);
    
    // Update aggregated stats (quick dashboard)
    updateAggregatedStats_(eventData);
    
  } catch (e) {
    // Silent fail - never interrupt user experience
    Logger.log('⚠️ Analytics logging failed: ' + e);
  }
}

/**
 * Check if current user is an admin
 * Used by popup to show admin-only features
 */
function checkIfCurrentUserIsAdmin() {
  try {
    const currentUser = Session.getActiveUser().getEmail();
    return ADMIN_EMAILS.includes(currentUser);
  } catch (e) {
    return false;
  }
}

/**
 * Queue analytics event (stores in PropertiesService)
 * Runs with user's permissions but doesn't access the sheet
 */
function logToSheet_(eventData) {
  try {
    // Get user email safely
    let userEmail = 'anonymous@commercecore.com';
    try {
      userEmail = eventData.user || Session.getActiveUser().getEmail();
    } catch (e) {
      Logger.log('⚠️ Could not get user email');
    }

    const rowData = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      action: eventData.action || '',
      product: eventData.product || '',
      locale: eventData.locale || '',
      batchId: eventData.batchId || '',
      template: eventData.template || '',
      success: eventData.success ? 'SUCCESS' : 'FAILED',
      duration: eventData.duration || '',
      error: eventData.error || '',
      additionalData: JSON.stringify(eventData.additionalData || {})
    };
    
    // Store in queue instead of writing directly
    const props = PropertiesService.getScriptProperties();
    const queueJson = props.getProperty('ANALYTICS_QUEUE') || '[]';
    const queue = JSON.parse(queueJson);
    
    queue.push(rowData);
    
    // Keep queue size manageable (max 100 events)
    if (queue.length > 100) {
      queue.shift(); // Remove oldest
    }
    
    props.setProperty('ANALYTICS_QUEUE', JSON.stringify(queue));
    
  } catch (e) {
    // Silent fail - don't interrupt user
    Logger.log('⚠️ Failed to queue analytics: ' + e);
  }
}

/**
 * Update aggregated statistics in PropertiesService
 */
function updateAggregatedStats_(eventData) {
  try {
    const props = PropertiesService.getScriptProperties();
    const statsJson = props.getProperty('ANALYTICS_STATS') || '{}';
    const stats = JSON.parse(statsJson);
    
    // Initialize structure if needed
    if (!stats.totalCreations) stats.totalCreations = 0;
    if (!stats.byUser) stats.byUser = {};
    if (!stats.byProduct) stats.byProduct = {};
    if (!stats.byAction) stats.byAction = {};
    if (!stats.byLocale) stats.byLocale = {};
    if (!stats.byTemplate) stats.byTemplate = {};
    if (!stats.errors) stats.errors = {};
    if (!stats.successCount) stats.successCount = 0;
    if (!stats.totalDuration) stats.totalDuration = 0;
    if (!stats.lastUpdated) stats.lastUpdated = '';
    
    // Update counts
    stats.totalCreations++;
    stats.lastUpdated = new Date().toISOString();
    
    const user = eventData.user || Session.getActiveUser().getEmail();
    stats.byUser[user] = (stats.byUser[user] || 0) + 1;
    
    if (eventData.product) {
      stats.byProduct[eventData.product] = (stats.byProduct[eventData.product] || 0) + 1;
    }
    
    if (eventData.action) {
      stats.byAction[eventData.action] = (stats.byAction[eventData.action] || 0) + 1;
    }
    
    if (eventData.locale) {
      stats.byLocale[eventData.locale] = (stats.byLocale[eventData.locale] || 0) + 1;
    }
    
    if (eventData.template) {
      stats.byTemplate[eventData.template] = (stats.byTemplate[eventData.template] || 0) + 1;
    }
    
    if (eventData.success) {
      stats.successCount++;
    } else if (eventData.error) {
      stats.errors[eventData.error] = (stats.errors[eventData.error] || 0) + 1;
    }
    
    if (eventData.duration) {
      stats.totalDuration += eventData.duration;
    }
    
    props.setProperty('ANALYTICS_STATS', JSON.stringify(stats));
    
  } catch (e) {
    Logger.log('⚠️ Failed to update aggregated stats: ' + e);
  }
}

/**
 * Get current analytics statistics
 */
function getAnalyticsStats_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const statsJson = props.getProperty('ANALYTICS_STATS') || '{}';
    return JSON.parse(statsJson);
  } catch (e) {
    Logger.log('⚠️ Failed to get stats: ' + e);
    return {};
  }
}

/**
 * Calculate time saved with detailed breakdown and quarterly projection
 */
function calculateTimeSaved_() {
  const stats = getAnalyticsStats_();
  const totalCreations = stats.totalCreations || 0;
  
  // Time saved per action (in seconds)
  const TIME_SAVED_BY_ACTION = {
    'Standard Setup': 90,
    'Custom Setup': 90,
    'Notes Only': 75,
    'Folder Only': 45
  };
  
  // Calculate total time saved
  let totalSecondsSaved = 0;
  const byAction = stats.byAction || {};
  
  Object.keys(byAction).forEach(action => {
    const count = byAction[action] || 0;
    const timeSavedPerAction = TIME_SAVED_BY_ACTION[action] || 90;
    totalSecondsSaved += count * timeSavedPerAction;
  });
  
  // Convert to hours:minutes:seconds
  const hours = Math.floor(totalSecondsSaved / 3600);
  const minutes = Math.floor((totalSecondsSaved % 3600) / 60);
  const seconds = totalSecondsSaved % 60;
  
  // Work days (8 hours = 1 day)
  const workDays = (totalSecondsSaved / (8 * 3600)).toFixed(1);
  
  // ✅ NEW: Quarterly projection
  const now = new Date();
  const currentMonth = now.getMonth(); // 0-11
  const currentQuarter = Math.floor(currentMonth / 3) + 1; // 1-4
  
  // Calculate days elapsed in current quarter
  const quarterStartMonth = (currentQuarter - 1) * 3;
  const quarterStart = new Date(now.getFullYear(), quarterStartMonth, 1);
  const quarterEnd = new Date(now.getFullYear(), quarterStartMonth + 3, 0);
  const totalDaysInQuarter = Math.ceil((quarterEnd - quarterStart) / (1000 * 60 * 60 * 24));
  const daysElapsedInQuarter = Math.ceil((now - quarterStart) / (1000 * 60 * 60 * 24));
  
  // Project to end of quarter
  let projectedSecondsSaved = 0;
  if (daysElapsedInQuarter > 0 && totalSecondsSaved > 0) {
    const dailyRate = totalSecondsSaved / daysElapsedInQuarter;
    projectedSecondsSaved = dailyRate * totalDaysInQuarter;
  }
  
  const projectedHours = Math.floor(projectedSecondsSaved / 3600);
  const projectedMinutes = Math.floor((projectedSecondsSaved % 3600) / 60);
  const projectedWorkDays = (projectedSecondsSaved / (8 * 3600)).toFixed(1);
  
  return {
    totalCreations,
    secondsSaved: totalSecondsSaved,
    hours,
    minutes,
    seconds,
    workDays: parseFloat(workDays),
    
    // Quarterly data
    currentQuarter,
    quarterName: 'Q' + currentQuarter,
    daysElapsedInQuarter,
    totalDaysInQuarter,
    quarterProgress: ((daysElapsedInQuarter / totalDaysInQuarter) * 100).toFixed(0),
    
    // Projections
    projectedSecondsSaved,
    projectedHours,
    projectedMinutes,
    projectedWorkDays: parseFloat(projectedWorkDays),
    
    byAction: TIME_SAVED_BY_ACTION
  };
}

/**
 * Show analytics dashboard in a dialog
 */
function showAnalyticsDashboard() {
  const stats = getAnalyticsStats_();
  const timeSaved = calculateTimeSaved_();
  
  const html = buildAnalyticsDashboardHtml_(stats, timeSaved);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Analytics Dashboard');
}

/**
 * Build the analytics dashboard HTML
 */
function buildAnalyticsDashboardHtml_(stats, timeSaved) {
  const totalCreations = stats.totalCreations || 0;
  const successRate = totalCreations > 0 
    ? ((stats.successCount || 0) / totalCreations * 100).toFixed(1) 
    : 0;
  
  // Top 5 users
  const userEntries = Object.entries(stats.byUser || {});
  userEntries.sort((a, b) => b[1] - a[1]);
  const topUsers = userEntries.slice(0, 5);
  
  // Top 5 products
  const productEntries = Object.entries(stats.byProduct || {});
  productEntries.sort((a, b) => b[1] - a[1]);
  const topProducts = productEntries.slice(0, 5);
  
  // Actions breakdown
  const actionEntries = Object.entries(stats.byAction || {});
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { 
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      padding: 24px; 
      background: #fafbfc; 
      margin: 0;
    }
    
    .container { 
      max-width: 700px; 
      margin: 0 auto; 
    }
    
    .header { 
      text-align: center; 
      margin-bottom: 32px; 
    }
    
    .header h1 { 
      margin: 0 0 8px 0; 
      color: #1a1a1a; 
      font-size: 32px; 
      font-weight: 700;
      letter-spacing: -0.5px;
    }
    
    .header p { 
      color: #6b7280; 
      font-size: 13px; 
      margin: 0;
      font-weight: 500;
    }
    
    /* ✨ Unified Card Grid */
    .stats-grid { 
      display: grid; 
      grid-template-columns: repeat(2, 1fr); 
      gap: 16px; 
      margin-bottom: 24px; 
    }
    
    .stat-card { 
      background: white; 
      border-radius: 12px; 
      padding: 24px; 
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
      border: 1px solid #e5e7eb;
      transition: all 0.2s ease;
    }
    
    .stat-card:hover {
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
      transform: translateY(-1px);
    }
    
    .stat-value { 
      font-size: 48px; 
      font-weight: 700; 
      color: #1a1a1a; 
      margin: 12px 0 8px 0;
      line-height: 1;
      letter-spacing: -1px;
    }
    
    .stat-label { 
      font-size: 12px; 
      color: #6b7280; 
      text-transform: uppercase; 
      letter-spacing: 0.8px;
      font-weight: 600;
      margin-bottom: 8px;
    }
    
    .stat-sublabel { 
      font-size: 13px; 
      color: #9ca3af; 
      margin-top: 4px;
      font-weight: 500;
    }
    
    /* ✨ Enhanced Time Display - Minimalist */
    .time-display {
      display: flex;
      align-items: baseline;
      justify-content: center;
      margin: 16px 0 12px;
      font-family: ui-monospace, 'SF Mono', Monaco, 'Cascadia Mono', 'Courier New', monospace;
      gap: 2px;
    }
    
    .time-large {
      font-size: 48px;
      font-weight: 700;
      color: #1a1a1a;
      line-height: 1;
      letter-spacing: -1px;
    }
    
    .time-unit {
      font-size: 14px;
      color: #9ca3af;
      margin-right: 6px;
      font-weight: 600;
      align-self: flex-end;
      margin-bottom: 6px;
    }
    
    .time-separator {
      font-size: 32px;
      color: #d1d5db;
      margin: 0 4px;
      font-weight: 300;
    }
    
    /* ✨ Quarterly Projection - Subtle */
    .projection-section {
      margin-top: 20px;
      padding-top: 20px;
      border-top: 1px solid #f3f4f6;
    }
    
    .projection-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }
    
    .quarter-badge {
      background: #f3f4f6;
      padding: 4px 10px;
      border-radius: 6px;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.5px;
      color: #4b5563;
    }
    
    .quarter-progress {
      font-size: 12px;
      color: #6b7280;
      font-weight: 600;
    }
    
    .projection-bar {
      height: 4px;
      background: #f3f4f6;
      border-radius: 2px;
      overflow: hidden;
      margin-bottom: 12px;
    }
    
    .projection-fill {
      height: 100%;
      background: linear-gradient(90deg, #3b82f6, #2563eb);
      border-radius: 2px;
      transition: width 0.3s ease;
    }
    
    .projection-text {
      font-size: 12px;
      color: #6b7280;
      text-align: center;
      line-height: 1.6;
    }
    
    .projection-text strong {
      color: #1a1a1a;
      font-weight: 600;
    }
    
    /* ✨ Sections */
    .section { 
      background: white; 
      border-radius: 12px; 
      padding: 24px; 
      margin-bottom: 16px; 
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
      border: 1px solid #e5e7eb;
    }
    
    .section-title { 
      font-size: 14px; 
      font-weight: 700; 
      color: #1a1a1a; 
      margin: 0 0 20px 0;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    .list-item { 
      display: flex; 
      justify-content: space-between; 
      padding: 12px 0; 
      border-bottom: 1px solid #f3f4f6;
    }
    
    .list-item:last-child { 
      border-bottom: none; 
    }
    
    .list-label { 
      color: #374151; 
      font-size: 14px;
      font-weight: 500;
    }
    
    .list-value { 
      color: #1a1a1a; 
      font-weight: 700; 
      font-size: 14px; 
    }
    
    .progress-bar { 
      height: 4px; 
      background: #f3f4f6; 
      border-radius: 2px; 
      margin-top: 8px; 
      overflow: hidden; 
    }
    
    .progress-fill { 
      height: 100%; 
      background: linear-gradient(90deg, #3b82f6, #2563eb);
      border-radius: 2px;
      transition: width 0.3s ease;
    }
    
    /* ✨ Footer */
    .footer { 
      text-align: center; 
      margin-top: 32px; 
      padding-top: 24px; 
      border-top: 1px solid #e5e7eb; 
    }
    
    .footer-link { 
      color: #3b82f6; 
      text-decoration: none; 
      font-size: 13px; 
      margin: 0 12px;
      font-weight: 600;
      transition: color 0.2s ease;
    }
    
    .footer-link:hover { 
      color: #2563eb;
      text-decoration: underline;
    }

  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 Analytics Dashboard</h1>
      <p>Last updated: ${stats.lastUpdated ? new Date(stats.lastUpdated).toLocaleString() : 'Never'}</p>
    </div>
    
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-label">Total Creations</div>
        <div class="stat-value">${totalCreations.toLocaleString()}</div>
        <div class="stat-sublabel">All-time ad notes created</div>
      </div>
      
      <div class="stat-card">
        <div class="stat-label">⏱️ TIME SAVED</div>
        <div class="time-display">
          <span class="time-large">${timeSaved.hours}</span><span class="time-unit">h</span>
          <span class="time-separator">:</span>
          <span class="time-large">${String(timeSaved.minutes).padStart(2, '0')}</span><span class="time-unit">m</span>
          <span class="time-separator">:</span>
          <span class="time-large">${String(timeSaved.seconds).padStart(2, '0')}</span><span class="time-unit">s</span>
        </div>
        <div class="stat-sublabel">${timeSaved.workDays} work days saved</div>
        
        ${timeSaved.projectedSecondsSaved > 0 ? `
        <div class="projection-section">
          <div class="projection-header">
            <span class="quarter-badge">${timeSaved.quarterName} ${new Date().getFullYear()}</span>
            <span class="quarter-progress">${timeSaved.quarterProgress}% complete</span>
          </div>
          <div class="projection-bar">
            <div class="projection-fill" style="width: ${timeSaved.quarterProgress}%"></div>
          </div>
          <div class="projection-text">
            <strong>Projected quarter end:</strong> ${timeSaved.projectedHours}h ${timeSaved.projectedMinutes}m <span style="color: #9ca3af;">(${timeSaved.projectedWorkDays} work days)</span>
          </div>
        </div>
        ` : ''}
      </div>
      
      <div class="stat-card">
        <div class="stat-label">Success Rate</div>
        <div class="stat-value">${successRate}%</div>
        <div class="stat-sublabel">${stats.successCount || 0} successful / ${totalCreations} total</div>
      </div>
      
      <div class="stat-card">
        <div class="stat-label">Active Users</div>
        <div class="stat-value">${Object.keys(stats.byUser || {}).length}</div>
        <div class="stat-sublabel">Team members using tool</div>
      </div>
    </div>
    
    <div class="section">
      <div class="section-title">👥 Top Users</div>
      ${topUsers.map(([email, count]) => {
        const percentage = (count / totalCreations * 100).toFixed(0);
        return `
          <div class="list-item">
            <span class="list-label">${email.split('@')[0]}</span>
            <span class="list-value">${count} (${percentage}%)</span>
          </div>
          <div class="progress-bar">
            <div class="progress-fill" style="width: ${percentage}%"></div>
          </div>
        `;
      }).join('')}
    </div>
    
    <div class="section">
      <div class="section-title">📦 Top Products</div>
      ${topProducts.map(([product, count]) => {
        const percentage = (count / totalCreations * 100).toFixed(0);
        return `
          <div class="list-item">
            <span class="list-label">${product}</span>
            <span class="list-value">${count} (${percentage}%)</span>
          </div>
        `;
      }).join('')}
    </div>
    
    <div class="section">
      <div class="section-title">⚡ Action Breakdown</div>
      ${actionEntries.map(([action, count]) => {
        const percentage = (count / totalCreations * 100).toFixed(0);
        return `
          <div class="list-item">
            <span class="list-label">${action}</span>
            <span class="list-value">${count} (${percentage}%)</span>
          </div>
        `;
      }).join('')}
    </div>
    
    <div class="footer">
      <a href="https://docs.google.com/spreadsheets/d/${ANALYTICS_SHEET_ID}" target="_blank" class="footer-link">
        📊 View Detailed Logs
      </a>
      <span style="color: #dadce0;">|</span>
      <a href="#" onclick="google.script.run.exportAnalyticsData(); google.script.host.close();" class="footer-link">
        📥 Export Data
      </a>
    </div>
  </div>
</body>
</html>
  `;
  
  return html;
}

/**
 * Export analytics data (opens the sheet)
 */
function exportAnalyticsData() {
  const url = 'https://docs.google.com/spreadsheets/d/' + ANALYTICS_SHEET_ID;
  const html = '<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(10).setHeight(10),
    'Opening Analytics Sheet...'
  );
}

/**
 * Share all templates with the domain (run once as admin)
 */
function shareTemplatesWithTeam() {
  // ✅ Check admin access
  let currentUser = '';
  try {
    currentUser = Session.getActiveUser().getEmail();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Permission Required', 'Please grant email access permission to use this function.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!ADMIN_EMAILS.includes(currentUser)) {
    SpreadsheetApp.getUi().alert('Admin Only', 'This function is only available to admins.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const templates = [
    { id: IMG_TEMPLATE_ID, name: 'IMG Template' },
    { id: VID_TEMPLATE_ID, name: 'VID Template' },
    { id: ITR_TEMPLATE_ID, name: 'ITR Template' },
    { id: LOC_TEMPLATE_ID, name: 'LOC Template' },
    { id: YTTOFB_TEMPLATE_ID, name: 'YTtoFB Template' }
  ];
  
  const results = [];
  
  templates.forEach(template => {
    try {
      const file = DriveApp.getFileById(template.id);
      
      // Share with domain (view only)
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
      
      results.push('✅ ' + template.name + ' - Shared with domain');
      Logger.log('✅ ' + template.name + ' shared successfully');
      
    } catch (e) {
      results.push('❌ ' + template.name + ' - Error: ' + e.message);
      Logger.log('❌ ' + template.name + ' error: ' + e);
    }
  });
  
  // Show results
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Template Sharing Results',
    results.join('\n\n'),
    ui.ButtonSet.OK
  );
}

/**
 * Check which users have used the tool and granted permissions
 * Admin-only function
 */
function checkUserAnalyticsStatus() {
  if (!ADMIN_EMAILS.includes(Session.getActiveUser().getEmail())) {
    SpreadsheetApp.getUi().alert('Admin Only', 'This function is only available to admins.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // Get analytics data
    const sheet = SpreadsheetApp.openById(ANALYTICS_SHEET_ID).getSheetByName('Events');
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('No Data', 'Analytics sheet not found or empty.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No usage data recorded yet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Extract unique users (skip header row)
    const users = new Set();
    for (let i = 1; i < data.length; i++) {
      const user = data[i][1]; // Column B = User email
      if (user) users.add(user);
    }
    
    // Build report
    let report = '📊 Analytics Authorization Status\n\n';
    report += '✅ USERS WITH ACCESS (authorized & used tool):\n';
    report += Array.from(users).map(u => '  • ' + u).join('\n');
    report += '\n\n';
    report += '📈 Total users tracked: ' + users.size + '\n';
    report += '📅 Data from: ' + new Date(data[1][0]).toLocaleDateString() + ' to ' + new Date(data[data.length-1][0]).toLocaleDateString();
    
    // Show in dialog
    const html = HtmlService.createHtmlOutput(
      '<pre style="font-family: monospace; font-size: 13px; white-space: pre-wrap;">' + 
      report + 
      '</pre>'
    ).setWidth(500).setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(html, '📊 User Analytics Status');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to check status: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Wrapper to detect permission failures
 */
function logAnalyticsEventSafe_(eventData) {
  if (!ANALYTICS_ENABLED) return;
  
  try {
    const currentUser = Session.getActiveUser().getEmail();
    
    // Skip admins if needed
    if (!TRACK_ADMINS && ADMIN_EMAILS.includes(currentUser)) {
      return;
    }
    
    logAnalyticsEvent_(eventData);
    
  } catch (e) {
    // ✅ Permission was denied or analytics sheet inaccessible
    Logger.log('⚠️ Analytics failed for user: ' + Session.getActiveUser().getEmail());
    Logger.log('   Error: ' + e.message);
    
    // Store locally that this user denied permissions
    try {
      const props = PropertiesService.getUserProperties();
      props.setProperty('ANALYTICS_DENIED', 'true');
      props.setProperty('ANALYTICS_DENIED_DATE', new Date().toISOString());
    } catch (_) {}
    
    // Don't throw error - let the tool continue working
  }
}

/**
 * Check if current user denied analytics permissions
 */
function checkMyAnalyticsStatus() {
  const props = PropertiesService.getUserProperties();
  const denied = props.getProperty('ANALYTICS_DENIED');
  const deniedDate = props.getProperty('ANALYTICS_DENIED_DATE');
  
  if (denied === 'true') {
    SpreadsheetApp.getUi().alert(
      '📊 Analytics Status',
      'You previously denied analytics permissions.\n\n' +
      'Date: ' + (deniedDate ? new Date(deniedDate).toLocaleString() : 'Unknown') + '\n\n' +
      'Your usage is not being tracked.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '📊 Analytics Status',
      'Your usage is being tracked (you authorized analytics).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Generate complete team analytics report (admin only)
 * Uses same data source as dashboard
 */
function generateTeamAnalyticsReport() {
  let currentUser = '';
  try {
    currentUser = Session.getActiveUser().getEmail();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Permission Required', 'Please grant email access permission.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!ADMIN_EMAILS.includes(currentUser)) {
    SpreadsheetApp.getUi().alert('Admin Only', 'This function is only available to admins.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // ✅ USE SAME SOURCE AS DASHBOARD
    const stats = getAnalyticsStats_();
    const byUser = stats.byUser || {};
    const usersTracked = Object.keys(byUser);
    
    if (usersTracked.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Data', 
        'No usage data recorded yet.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Build report
    let html = '<style>body{font-family:system-ui;padding:20px;background:#fafbfc}h2{margin-top:0;color:#1a1a1a}.section{background:white;padding:16px;border-radius:12px;margin:12px 0;border:1px solid #e5e7eb}h3{margin-top:0;font-size:14px;text-transform:uppercase;letter-spacing:0.5px}ul{line-height:1.8;padding-left:24px}.tracked{color:#16a34a}.count{color:#6b7280;font-size:13px}.summary{background:white;padding:20px;border-radius:12px;margin:12px 0;border:1px solid #e5e7eb;text-align:center}.stat{font-size:32px;font-weight:bold;color:#1a1a1a;margin:8px 0}</style>';
    
    html += '<h2>📊 Team Analytics Report</h2>';
    
    html += '<div class="summary">';
    html += '<div class="stat">' + usersTracked.length + '</div>';
    html += '<div style="color:#6b7280;font-size:14px">Active Users</div>';
    html += '<div style="color:#9ca3af;font-size:12px;margin-top:8px">' + (stats.totalCreations || 0) + ' total creations</div>';
    html += '</div>';
    
    html += '<div class="section">';
    html += '<h3 style="color:#16a34a">✅ Users Being Tracked (' + usersTracked.length + ')</h3>';
    html += '<ul class="tracked">';
    
    const sortedUsers = usersTracked.sort((a, b) => byUser[b] - byUser[a]);
    
    sortedUsers.forEach(email => {
      const count = byUser[email];
      const percentage = stats.totalCreations > 0 ? ((count / stats.totalCreations) * 100).toFixed(0) : 0;
      html += '<li>' + email + ' <span class="count">(' + count + ' uses • ' + percentage + '%)</span></li>';
    });
    html += '</ul>';
    html += '</div>';
    
    html += '<div class="section">';
    html += '<h3 style="color:#ea580c">⚠️ Not Yet Tracked</h3>';
    html += '<p style="font-size:13px;color:#6b7280;line-height:1.6">Users who haven\'t used the tool since analytics were added:</p>';
    html += '<ul style="color:#9ca3af;font-size:13px;line-height:1.8">';
    html += '<li>Haven\'t used the tool yet</li>';
    html += '<li>Denied permissions when prompted</li>';
    html += '<li>Have restricted account permissions</li>';
    html += '</ul>';
    html += '</div>';
    
    html += '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">';
    html += '<p style="font-size:12px;color:#6b7280;text-align:center">Last checked: ' + new Date().toLocaleString() + '</p>';
    
    const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Team Analytics Report');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to generate report: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate complete team analytics report with roster comparison (admin only)
 */
function generateCompleteTeamReport() {
  let currentUser = '';
  try {
    currentUser = Session.getActiveUser().getEmail();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Permission Required', 'Please grant email access permission.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!ADMIN_EMAILS.includes(currentUser)) {
    SpreadsheetApp.getUi().alert('Admin Only', 'This function is only available to admins.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Check if TEAM_MEMBERS is defined
  if (typeof TEAM_MEMBERS === 'undefined' || !TEAM_MEMBERS || TEAM_MEMBERS.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Setup Required',
      'Please add a TEAM_MEMBERS array to your script with all team emails.\n\nExample:\nconst TEAM_MEMBERS = [\n  "user1@commercecore.com",\n  "user2@commercecore.com"\n];',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  try {
    // Get all users who have used the tool
    const sheet = SpreadsheetApp.openById(ANALYTICS_SHEET_ID).getSheetByName('Events');
    const usersTracked = new Set();
    
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const user = data[i][1];
        if (user) usersTracked.add(user.toLowerCase());
      }
    }
    
    // Compare with roster
    const trackedList = [];
    const notTrackedList = [];
    
    TEAM_MEMBERS.forEach(email => {
      const emailLower = email.toLowerCase();
      if (usersTracked.has(emailLower)) {
        trackedList.push(email);
      } else {
        notTrackedList.push(email);
      }
    });
    
    // Build report
    let html = '<style>body{font-family:system-ui;padding:20px;background:#fafbfc}h2{margin-top:0;color:#1a1a1a}.section{background:white;padding:16px;border-radius:12px;margin:12px 0;border:1px solid #e5e7eb}h3{margin-top:0;font-size:14px;text-transform:uppercase;letter-spacing:0.5px}ul{line-height:1.8;padding-left:24px}.tracked{color:#16a34a}.unknown{color:#ea580c}.stat{font-size:24px;font-weight:bold}.summary{display:flex;gap:20px;margin:20px 0}.summary-card{flex:1;background:white;padding:16px;border-radius:12px;border:1px solid #e5e7eb;text-align:center}</style>';
    
    html += '<h2>📊 Complete Team Analytics Report</h2>';
    
    // Summary cards
    html += '<div class="summary">';
    html += '<div class="summary-card"><div class="stat" style="color:#16a34a">' + trackedList.length + '</div><div>Tracked</div></div>';
    html += '<div class="summary-card"><div class="stat" style="color:#ea580c">' + notTrackedList.length + '</div><div>Not Tracked</div></div>';
    html += '<div class="summary-card"><div class="stat">' + TEAM_MEMBERS.length + '</div><div>Total Team</div></div>';
    html += '</div>';
    
    // Tracked users
    html += '<div class="section">';
    html += '<h3 style="color:#16a34a">✅ Being Tracked (' + trackedList.length + ')</h3>';
    if (trackedList.length > 0) {
      html += '<ul class="tracked">';
      trackedList.forEach(u => html += '<li>' + u + '</li>');
      html += '</ul>';
    } else {
      html += '<p style="color:#9ca3af">No users tracked yet</p>';
    }
    html += '</div>';
    
    // Not tracked users
    html += '<div class="section">';
    html += '<h3 style="color:#ea580c">⚠️ Not Yet Tracked (' + notTrackedList.length + ')</h3>';
    if (notTrackedList.length > 0) {
      html += '<p style="font-size:13px;color:#6b7280;margin-bottom:12px">These users haven\'t used the tool since analytics were added, or denied permissions:</p>';
      html += '<ul class="unknown">';
      notTrackedList.forEach(u => html += '<li>' + u + '</li>');
      html += '</ul>';
    } else {
      html += '<p style="color:#16a34a;font-weight:600">🎉 Everyone is tracked!</p>';
    }
    html += '</div>';
    
    html += '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">';
    html += '<p style="font-size:12px;color:#6b7280;text-align:center">Last checked: ' + new Date().toLocaleString() + '</p>';
    
    const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Complete Team Report');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to generate report: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Check template permissions for current user
 */
function checkMyTemplatePermissions() {
  let user = 'Current User';
  try {
    user = Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log('⚠️ Could not get user email');
  }
  const templates = [
    { id: IMG_TEMPLATE_ID, name: 'IMG Template' },
    { id: VID_TEMPLATE_ID, name: 'VID Template' },
    { id: ITR_TEMPLATE_ID, name: 'ITR Template' },
    { id: LOC_TEMPLATE_ID, name: 'LOC Template' },
    { id: YTTOFB_TEMPLATE_ID, name: 'YTtoFB Template' }
  ];
  
  let results = '🔐 Template Permissions Check\n';
  results += 'User: ' + user + '\n\n';
  
  templates.forEach(template => {
    try {
      const file = DriveApp.getFileById(template.id);
      const name = file.getName();
      
      // Check if we can actually copy it
      try {
        const testFolder = DriveApp.getRootFolder();
        const testCopy = file.makeCopy('TEST_COPY_DELETE_ME', testFolder);
        testCopy.setTrashed(true); // Clean up immediately
        
        results += '✅ ' + template.name + '\n';
        results += '   File name: "' + name + '"\n';
        results += '   Can copy: YES ✓\n\n';
        
      } catch (copyError) {
        results += '⚠️ ' + template.name + '\n';
        results += '   File name: "' + name + '"\n';
        results += '   Can VIEW: YES\n';
        results += '   Can COPY: NO ✗\n';
        results += '   Error: ' + copyError.message + '\n\n';
      }
      
    } catch (e) {
      results += '❌ ' + template.name + '\n';
      results += '   Can access: NO\n';
      results += '   Error: ' + e.message + '\n\n';
    }
  });
  
  Logger.log(results);
  SpreadsheetApp.getUi().alert('Template Permissions Check', results, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Install analytics trigger (run once as admin)
 */
function installAnalyticsTrigger() {
  // Remove any existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'logAnalyticsQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Install new time-based trigger (runs every minute)
  ScriptApp.newTrigger('processAnalyticsQueue_')
    .timeBased()
    .everyMinutes(1)
    .create();
  
  SpreadsheetApp.getUi().alert('✅ Analytics trigger installed!');
}

/**
 * Process queued analytics events (runs as trigger with admin permissions)
 * This function runs with the script owner's permissions, not the user's
 */
function processAnalyticsQueue_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const queueJson = props.getProperty('ANALYTICS_QUEUE');
    
    if (!queueJson) return; // Nothing to process
    
    const queue = JSON.parse(queueJson);
    
    if (queue.length === 0) return; // Empty queue
    
    // Get or create Events sheet
    const ss = SpreadsheetApp.openById(ANALYTICS_SHEET_ID);
    let sheet = ss.getSheetByName('Events');
    
    if (!sheet) {
      sheet = ss.insertSheet('Events');
      sheet.appendRow([
        'Timestamp', 'User', 'Action', 'Product', 'Locale', 'Batch ID', 
        'Template', 'Success', 'Duration (s)', 'Error Message', 'Additional Data'
      ]);
      sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    
    // Write all queued events
    queue.forEach(event => {
      const row = [
        new Date(event.timestamp),
        event.user,
        event.action,
        event.product,
        event.locale,
        event.batchId,
        event.template,
        event.success,
        event.duration,
        event.error,
        event.additionalData
      ];
      sheet.appendRow(row);
    });
    
    // Clear the queue
    props.setProperty('ANALYTICS_QUEUE', '[]');
    
    Logger.log('✅ Processed ' + queue.length + ' analytics events');
    
  } catch (e) {
    Logger.log('⚠️ Failed to process analytics queue: ' + e);
    // Don't clear queue on error - will retry next minute
  }
}

/**
 * Recalculate analytics stats from the sheet (admin only)
 * Call this after manually editing the analytics sheet
 */
function recalculateAnalyticsStats() {
  if (!ANALYTICS_ENABLED) {
    SpreadsheetApp.getUi().alert('Analytics Disabled', 'Analytics is currently disabled in the script configuration.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // ✅ Admin check
  let currentUser = '';
  try {
    currentUser = Session.getActiveUser().getEmail();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Permission Required', 'Please grant email access permission.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (!ADMIN_EMAILS.includes(currentUser)) {
    SpreadsheetApp.getUi().alert('Admin Only', 'This function is only available to admins.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(ANALYTICS_SHEET_ID);
    const sheet = ss.getSheetByName('Events');
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('No Data', 'Analytics sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      // Only header row, clear stats
      const props = PropertiesService.getScriptProperties();
      props.setProperty('ANALYTICS_STATS', JSON.stringify({
        totalCreations: 0,
        byUser: {},
        byProduct: {},
        byAction: {},
        byLocale: {},
        byTemplate: {},
        errors: {},
        successCount: 0,
        totalDuration: 0,
        lastUpdated: new Date().toISOString()
      }));
      
      SpreadsheetApp.getUi().alert('✅ Stats Reset', 'Analytics stats have been reset (sheet is empty).', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Rebuild stats from scratch
    const stats = {
      totalCreations: 0,
      byUser: {},
      byProduct: {},
      byAction: {},
      byLocale: {},
      byTemplate: {},
      errors: {},
      successCount: 0,
      totalDuration: 0,
      lastUpdated: new Date().toISOString()
    };
    
    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Column mapping:
      // A: Timestamp, B: User, C: Action, D: Product, E: Locale, F: Batch ID,
      // G: Template, H: Success, I: Duration, J: Error, K: Additional Data
      
      const user = row[1] || 'unknown';
      const action = row[2] || '';
      const product = row[3] || '';
      const locale = row[4] || '';
      const template = row[6] || '';
      const success = row[7] === 'SUCCESS';
      const duration = parseFloat(row[8]) || 0;
      const error = row[9] || '';
      
      // Update counts
      stats.totalCreations++;
      
      if (user) {
        stats.byUser[user] = (stats.byUser[user] || 0) + 1;
      }
      
      if (product) {
        stats.byProduct[product] = (stats.byProduct[product] || 0) + 1;
      }
      
      if (action) {
        stats.byAction[action] = (stats.byAction[action] || 0) + 1;
      }
      
      if (locale) {
        stats.byLocale[locale] = (stats.byLocale[locale] || 0) + 1;
      }
      
      if (template) {
        stats.byTemplate[template] = (stats.byTemplate[template] || 0) + 1;
      }
      
      if (success) {
        stats.successCount++;
      } else if (error) {
        stats.errors[error] = (stats.errors[error] || 0) + 1;
      }
      
      if (duration) {
        stats.totalDuration += duration;
      }
    }
    
    // Save recalculated stats
    const props = PropertiesService.getScriptProperties();
    props.setProperty('ANALYTICS_STATS', JSON.stringify(stats));
    
    SpreadsheetApp.getUi().alert(
      '✅ Stats Recalculated',
      'Analytics stats have been rebuilt from the sheet.\n\n' +
      'Total events: ' + stats.totalCreations + '\n' +
      'Unique users: ' + Object.keys(stats.byUser).length + '\n' +
      'Success rate: ' + (stats.totalCreations > 0 ? ((stats.successCount / stats.totalCreations) * 100).toFixed(1) : 0) + '%',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to recalculate stats: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Check if current user is allowed to use ClickUp integration
 */
function checkClickUpPermission() {
  try {
    const currentUser = Session.getActiveUser().getEmail();
    return CLICKUP_ALLOWED_USERS.includes(currentUser);
  } catch (e) {
    Logger.log('⚠️ Could not check ClickUp permission: ' + e);
    return false;
  }
}

function testAdNamingConfig() {
  Logger.log('Admins: ' + JSON.stringify(ADMIN_EMAILS));
  Logger.log('ClickUp Users: ' + CLICKUP_ALLOWED_USERS.length);
  Logger.log('Andromeda Users: ' + ANDROMEDA_USERS.length);
  Logger.log('Locales: ' + JSON.stringify(LOCALES));
  Logger.log('Products in index: ' + Object.keys(PRODUCT_INDEX).length);
  Logger.log('Sample product (CamTrix): ' + JSON.stringify(PRODUCT_INDEX.CamTrix));
  Logger.log('Products with folders: ' + Object.keys(PRODUCT_LOCALE_FOLDERS).length);
  Logger.log('Sample folder (CamTrix US): ' + (PRODUCT_LOCALE_FOLDERS.CamTrix ? PRODUCT_LOCALE_FOLDERS.CamTrix.US : 'NOT FOUND'));
}

/**
 * ================================================================================
 * AD NAMING SHEET - COMPREHENSIVE CONFIG DIAGNOSTICS
 * ================================================================================
 * Run this function to test all config connections and validate data integrity.
 * Results are logged and also shown in a popup summary.
 */

function runFullConfigDiagnostics() {
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
  Logger.log('║        AD NAMING SHEET - CONFIG DIAGNOSTICS REPORT              ║');
  Logger.log('║        ' + new Date().toLocaleString() + '                            ║');
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
    log('Connection', 'Open Config Sheet', 'PASS', 'Sheet ID: ' + CONFIG_SHEET_ID);
    
    // Check all required tabs exist
    const requiredTabs = ['Products', 'Product Folders', 'Team Members', 'ClickUp Config', 'Valid Values', 'Sync Log'];
    for (const tabName of requiredTabs) {
      const tab = configSheet.getSheetByName(tabName);
      if (tab) {
        const rowCount = tab.getLastRow();
        log('Connection', `Tab: ${tabName}`, 'PASS', `${rowCount} rows`);
      } else {
        log('Connection', `Tab: ${tabName}`, 'FAIL', 'Tab not found!');
      }
    }
  } catch (e) {
    log('Connection', 'Open Config Sheet', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 2: ADMIN EMAILS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('👤 TEST 2: ADMIN EMAILS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const admins = ADMIN_EMAILS;
    if (Array.isArray(admins) && admins.length > 0) {
      log('Admins', 'Load Admin Emails', 'PASS', `${admins.length} admin(s)`);
      
      // Validate email format
      const validEmails = admins.filter(e => e && e.includes('@'));
      if (validEmails.length === admins.length) {
        log('Admins', 'Email Format Validation', 'PASS', 'All emails valid');
      } else {
        log('Admins', 'Email Format Validation', 'WARN', `${admins.length - validEmails.length} invalid email(s)`);
      }
      
      Logger.log('   Admins: ' + admins.join(', '));
    } else {
      log('Admins', 'Load Admin Emails', 'WARN', 'No admins configured');
    }
  } catch (e) {
    log('Admins', 'Load Admin Emails', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 3: TEAM MEMBERS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('👥 TEST 3: TEAM MEMBERS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const members = TEAM_MEMBERS;
    if (Array.isArray(members) && members.length > 0) {
      log('Team', 'Load Team Members', 'PASS', `${members.length} member(s)`);
      
      // Check for duplicates
      const uniqueMembers = [...new Set(members)];
      if (uniqueMembers.length === members.length) {
        log('Team', 'Duplicate Check', 'PASS', 'No duplicates');
      } else {
        log('Team', 'Duplicate Check', 'WARN', `${members.length - uniqueMembers.length} duplicate(s)`);
      }
      
      // Check commercecore domain
      const ccMembers = members.filter(e => e && e.endsWith('@commercecore.com'));
      log('Team', 'Domain Check', ccMembers.length === members.length ? 'PASS' : 'WARN', 
          `${ccMembers.length}/${members.length} @commercecore.com`);
      
      Logger.log('   Members: ' + members.slice(0, 5).join(', ') + (members.length > 5 ? '...' : ''));
    } else {
      log('Team', 'Load Team Members', 'FAIL', 'No team members found');
    }
  } catch (e) {
    log('Team', 'Load Team Members', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 4: CLICKUP ACCESS USERS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔗 TEST 4: CLICKUP ACCESS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const clickupUsers = CLICKUP_ALLOWED_USERS;
    if (Array.isArray(clickupUsers)) {
      log('ClickUp', 'Load ClickUp Users', 'PASS', `${clickupUsers.length} user(s)`);
      Logger.log('   Users: ' + clickupUsers.join(', '));
    } else {
      log('ClickUp', 'Load ClickUp Users', 'FAIL', 'Not an array');
    }
    
    // Test API key retrieval for current user
    const currentUser = Session.getActiveUser().getEmail();
    if (currentUser) {
      log('ClickUp', 'Current User Detection', 'PASS', currentUser);
      
      if (typeof getClickUpApiKeyFromConfig_ === 'function') {
        const apiKey = getClickUpApiKeyFromConfig_();
        if (apiKey && apiKey.startsWith('pk_')) {
          log('ClickUp', 'API Key Retrieval', 'PASS', 'Key found (pk_...)');
        } else if (apiKey) {
          log('ClickUp', 'API Key Retrieval', 'WARN', 'Key found but format unexpected');
        } else {
          log('ClickUp', 'API Key Retrieval', 'WARN', 'No API key for current user');
        }
      }
    } else {
      log('ClickUp', 'Current User Detection', 'WARN', 'Could not detect user email');
    }
  } catch (e) {
    log('ClickUp', 'ClickUp Access Test', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 5: ANDROMEDA USERS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🌌 TEST 5: ANDROMEDA ACCESS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const andromedaUsers = ANDROMEDA_USERS;
    if (Array.isArray(andromedaUsers)) {
      log('Andromeda', 'Load Andromeda Users', 'PASS', `${andromedaUsers.length} user(s)`);
      Logger.log('   Users: ' + andromedaUsers.join(', '));
    } else {
      log('Andromeda', 'Load Andromeda Users', 'FAIL', 'Not an array');
    }
  } catch (e) {
    log('Andromeda', 'Load Andromeda Users', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 6: LOCALES
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🌍 TEST 6: LOCALES');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const locales = LOCALES;
    if (Array.isArray(locales) && locales.length > 0) {
      log('Locales', 'Load Locales', 'PASS', `${locales.length} locale(s)`);
      Logger.log('   Locales: ' + locales.join(', '));
      
      // Check for expected core locales
      const coreLocales = ['US', 'UK', 'DE', 'CA', 'AU', 'LATAM', 'ES'];
      const missingCore = coreLocales.filter(l => !locales.includes(l));
      if (missingCore.length === 0) {
        log('Locales', 'Core Locales Present', 'PASS', 'All 7 core locales found');
      } else {
        log('Locales', 'Core Locales Present', 'WARN', `Missing: ${missingCore.join(', ')}`);
      }
    } else {
      log('Locales', 'Load Locales', 'FAIL', 'No locales found');
    }
  } catch (e) {
    log('Locales', 'Load Locales', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 7: PRODUCT INDEX
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📦 TEST 7: PRODUCT INDEX');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const products = PRODUCT_INDEX;
    if (typeof products === 'object' && Object.keys(products).length > 0) {
      const productKeys = Object.keys(products);
      log('Products', 'Load Product Index', 'PASS', `${productKeys.length} product(s)`);
      
      // Validate product structure
      let validStructure = 0;
      let missingId = [];
      let missingName = [];
      
      for (const key of productKeys) {
        const prod = products[key];
        if (prod && typeof prod === 'object') {
          validStructure++;
          if (!prod.id) missingId.push(key);
          if (!prod.name) missingName.push(key);
        }
      }
      
      log('Products', 'Structure Validation', validStructure === productKeys.length ? 'PASS' : 'WARN',
          `${validStructure}/${productKeys.length} valid structures`);
      
      if (missingId.length > 0) {
        log('Products', 'Product IDs', 'WARN', `${missingId.length} missing IDs: ${missingId.slice(0,3).join(', ')}...`);
      } else {
        log('Products', 'Product IDs', 'PASS', 'All products have IDs');
      }
      
      // Test specific products
      const testProducts = ['CamTrix', 'TellyStick', 'Guardality', 'AliveBlue', 'KatuChef'];
      for (const testProd of testProducts) {
        if (products[testProd]) {
          Logger.log(`   ${testProd}: id=${products[testProd].id}, name=${products[testProd].name}`);
        }
      }
    } else {
      log('Products', 'Load Product Index', 'FAIL', 'No products found');
    }
  } catch (e) {
    log('Products', 'Load Product Index', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 8: PRODUCT LOCALE FOLDERS
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📁 TEST 8: PRODUCT LOCALE FOLDERS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    const folders = PRODUCT_LOCALE_FOLDERS;
    if (typeof folders === 'object' && Object.keys(folders).length > 0) {
      const folderKeys = Object.keys(folders);
      log('Folders', 'Load Product Folders', 'PASS', `${folderKeys.length} product(s) with folders`);
      
      // Count total folder links
      let totalLinks = 0;
      let emptyLinks = 0;
      let invalidLinks = [];
      
      for (const prodKey of folderKeys) {
        const prodFolders = folders[prodKey];
        if (prodFolders && typeof prodFolders === 'object') {
          for (const locale of Object.keys(prodFolders)) {
            const link = prodFolders[locale];
            if (link && link.trim() !== '') {
              totalLinks++;
              if (!link.includes('drive.google.com')) {
                invalidLinks.push(`${prodKey}/${locale}`);
              }
            } else {
              emptyLinks++;
            }
          }
        }
      }
      
      log('Folders', 'Folder Link Count', 'PASS', `${totalLinks} links, ${emptyLinks} empty`);
      
      if (invalidLinks.length > 0) {
        log('Folders', 'Link Format Validation', 'WARN', `${invalidLinks.length} non-Google Drive links`);
      } else {
        log('Folders', 'Link Format Validation', 'PASS', 'All links are Google Drive URLs');
      }
      
      // Test specific folder lookups
      const testCases = [
        { product: 'CamTrix', locale: 'US' },
        { product: 'TellyStick', locale: 'DE' },
        { product: 'Guardality', locale: 'ES' },
        { product: 'AliveBlue', locale: 'LATAM' }
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
  // TEST 9: CROSS-REFERENCE VALIDATION
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔀 TEST 9: CROSS-REFERENCE VALIDATION');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Check if products in PRODUCT_INDEX have matching folders
    const productKeys = Object.keys(PRODUCT_INDEX || {});
    const folderKeys = Object.keys(PRODUCT_LOCALE_FOLDERS || {});
    
    const productsWithoutFolders = productKeys.filter(p => !folderKeys.includes(p));
    const foldersWithoutProducts = folderKeys.filter(f => !productKeys.includes(f));
    
    if (productsWithoutFolders.length === 0) {
      log('CrossRef', 'Products → Folders', 'PASS', 'All products have folder entries');
    } else {
      log('CrossRef', 'Products → Folders', 'WARN', 
          `${productsWithoutFolders.length} products without folders: ${productsWithoutFolders.slice(0,5).join(', ')}`);
    }
    
    if (foldersWithoutProducts.length === 0) {
      log('CrossRef', 'Folders → Products', 'PASS', 'All folder entries have products');
    } else {
      log('CrossRef', 'Folders → Products', 'WARN', 
          `${foldersWithoutProducts.length} orphan folder entries: ${foldersWithoutProducts.join(', ')}`);
    }
    
    // Check if ClickUp users are in team members
    const teamSet = new Set(TEAM_MEMBERS || []);
    const clickupNotInTeam = (CLICKUP_ALLOWED_USERS || []).filter(u => !teamSet.has(u));
    
    if (clickupNotInTeam.length === 0) {
      log('CrossRef', 'ClickUp Users ⊂ Team', 'PASS', 'All ClickUp users are team members');
    } else {
      log('CrossRef', 'ClickUp Users ⊂ Team', 'WARN', 
          `${clickupNotInTeam.length} ClickUp users not in team: ${clickupNotInTeam.join(', ')}`);
    }
    
    // Check if admins are in team members
    const adminsNotInTeam = (ADMIN_EMAILS || []).filter(u => !teamSet.has(u));
    if (adminsNotInTeam.length === 0) {
      log('CrossRef', 'Admins ⊂ Team', 'PASS', 'All admins are team members');
    } else {
      log('CrossRef', 'Admins ⊂ Team', 'WARN', 
          `${adminsNotInTeam.length} admins not in team: ${adminsNotInTeam.join(', ')}`);
    }
  } catch (e) {
    log('CrossRef', 'Cross-Reference Check', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 10: CACHE PERFORMANCE
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('⚡ TEST 10: CACHE PERFORMANCE');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Clear cache first
    if (typeof clearAdNamingConfigCache === 'function') {
      CacheService.getScriptCache().removeAll(['adn_admins', 'adn_clickup_users', 'adn_andromeda_users', 'adn_locales', 'adn_product_index', 'adn_folders', 'adn_team_members']);
      
      // First call (uncached)
      const uncachedStart = Date.now();
      getProductIndexFromConfig_();
      getProductLocaleFoldersFromConfig_();
      const uncachedTime = Date.now() - uncachedStart;
      
      // Second call (cached)
      const cachedStart = Date.now();
      getProductIndexFromConfig_();
      getProductLocaleFoldersFromConfig_();
      const cachedTime = Date.now() - cachedStart;
      
      log('Cache', 'Uncached Load Time', uncachedTime < 5000 ? 'PASS' : 'WARN', `${uncachedTime}ms`);
      log('Cache', 'Cached Load Time', cachedTime < 100 ? 'PASS' : 'WARN', `${cachedTime}ms`);
      log('Cache', 'Cache Speedup', cachedTime < uncachedTime ? 'PASS' : 'WARN', 
          `${Math.round(uncachedTime / Math.max(cachedTime, 1))}x faster`);
    } else {
      log('Cache', 'Cache Test', 'WARN', 'clearAdNamingConfigCache not found');
    }
  } catch (e) {
    log('Cache', 'Cache Performance Test', 'FAIL', e.message);
  }

  // ============================================================================
  // TEST 11: EDGE CASES
  // ============================================================================
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🔬 TEST 11: EDGE CASES');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  
  try {
    // Test non-existent product
    const fakeProduct = PRODUCT_INDEX['NonExistentProduct123'];
    log('EdgeCase', 'Non-existent Product', fakeProduct === undefined ? 'PASS' : 'WARN', 
        fakeProduct === undefined ? 'Returns undefined' : 'Unexpected value');
    
    // Test product with special characters
    const dermaSonic = PRODUCT_INDEX['DermaSonic+'];
    log('EdgeCase', 'Product with Special Chars (DermaSonic+)', 
        dermaSonic !== undefined ? 'PASS' : 'WARN',
        dermaSonic ? 'Found' : 'Not found - may need alias');
    
    // Test LATAM/ES aliasing in folders
    const folders = PRODUCT_LOCALE_FOLDERS;
    if (folders.CamTrix) {
      const latamFolder = folders.CamTrix.LATAM;
      const esFolder = folders.CamTrix.ES;
      if (latamFolder && esFolder && latamFolder === esFolder) {
        log('EdgeCase', 'LATAM/ES Alias Check', 'PASS', 'Same folder for both');
      } else if (latamFolder || esFolder) {
        log('EdgeCase', 'LATAM/ES Alias Check', 'WARN', 'Different or partial folders');
      } else {
        log('EdgeCase', 'LATAM/ES Alias Check', 'WARN', 'No LATAM/ES folders for CamTrix');
      }
    }
    
    // Test empty locale handling
    const emptyLocaleProduct = Object.keys(folders).find(p => {
      const f = folders[p];
      return f && Object.values(f).some(v => v === '');
    });
    log('EdgeCase', 'Empty Locale Handling', emptyLocaleProduct ? 'PASS' : 'WARN',
        emptyLocaleProduct ? `${emptyLocaleProduct} has empty locales (expected)` : 'No empty locales found');
        
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
    Logger.log('║  ✅ All critical tests passed. Review warnings if needed.       ║');
  } else {
    Logger.log('║  ⚠️  Some tests failed. Review the log above for details.       ║');
  }
  
  Logger.log('╚══════════════════════════════════════════════════════════════════╝');
  
  // Show popup summary
  const ui = SpreadsheetApp.getUi();
  const summaryMsg = `
CONFIG DIAGNOSTICS COMPLETE

✅ Passed: ${results.passed}
⚠️ Warnings: ${results.warnings}  
❌ Failed: ${results.failed}

Time: ${totalTime}ms

${results.failed === 0 ? '🎉 All critical tests passed!' : '⚠️ Check Execution Log for details.'}

View full results: View → Execution log
  `;
  
  ui.alert('Config Diagnostics', summaryMsg, ui.ButtonSet.OK);
  
  return results;
}


/**
 * Quick health check - run this for a fast status
 */
function quickConfigHealthCheck() {
  const checks = [];
  
  try {
    checks.push({ name: 'Config Sheet', ok: !!SpreadsheetApp.openById(CONFIG_SHEET_ID) });
  } catch (e) {
    checks.push({ name: 'Config Sheet', ok: false, error: e.message });
  }
  
  checks.push({ name: 'Admin Emails', ok: Array.isArray(ADMIN_EMAILS) && ADMIN_EMAILS.length > 0 });
  checks.push({ name: 'Team Members', ok: Array.isArray(TEAM_MEMBERS) && TEAM_MEMBERS.length > 0 });
  checks.push({ name: 'Locales', ok: Array.isArray(LOCALES) && LOCALES.length > 0 });
  checks.push({ name: 'Product Index', ok: typeof PRODUCT_INDEX === 'object' && Object.keys(PRODUCT_INDEX).length > 0 });
  checks.push({ name: 'Product Folders', ok: typeof PRODUCT_LOCALE_FOLDERS === 'object' && Object.keys(PRODUCT_LOCALE_FOLDERS).length > 0 });
  
  const passed = checks.filter(c => c.ok).length;
  const total = checks.length;
  
  let msg = `Quick Health Check: ${passed}/${total} passed\n\n`;
  for (const c of checks) {
    msg += `${c.ok ? '✅' : '❌'} ${c.name}${c.error ? ': ' + c.error : ''}\n`;
  }
  
  Logger.log(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${passed}/${total} checks passed`,
    passed === total ? '✅ Config OK' : '⚠️ Config Issues',
    5
  );
  
  return { passed, total, checks };
}

/**
 * Run diagnostics for Ad Naming config centralization
 */
function runAdNamingConfigDiagnostics() {
  Logger.log('');
  Logger.log('╔══════════════════════════════════════════════════════════════════╗');
  Logger.log('║       AD NAMING SHEET - CONFIG DIAGNOSTICS                       ║');
  Logger.log('╚══════════════════════════════════════════════════════════════════╝');
  Logger.log('');
  
  let passed = 0;
  let warnings = 0;
  let failed = 0;
  
  // Test 1: Templates from Config
  Logger.log('━━━ TEST 1: TEMPLATES FROM CONFIG ━━━');
  try {
    const templates = getTemplatesFromConfig_();
    if (templates && Object.keys(templates).length > 0) {
      Logger.log('✅ Templates loaded from central Config');
      Logger.log('   Count: ' + Object.keys(templates).length);
      Logger.log('   Keys: ' + Object.keys(templates).join(', '));
      passed++;
      
      // Verify each template is accessible
      const templateKeys = ['IMG', 'VID', 'ITR', 'LOC', 'YTTOFB', 'ANDROMEDA_VID'];
      for (const key of templateKeys) {
        if (templates[key]) {
          Logger.log('   ✅ ' + key + ': ' + templates[key].substring(0, 20) + '...');
        } else {
          Logger.log('   ⚠️ ' + key + ': Not found in Config');
        }
      }
    } else {
      Logger.log('⚠️ Templates not in central Config - using hardcoded fallback');
      warnings++;
    }
  } catch (e) {
    Logger.log('❌ Templates error: ' + e.message);
    failed++;
  }
  
  // Test 2: Footage Folders from Config
  Logger.log('');
  Logger.log('━━━ TEST 2: FOOTAGE FOLDERS FROM CONFIG ━━━');
  try {
    const footage = getFootageFoldersFromConfig_();
    if (footage && Object.keys(footage).length > 0) {
      Logger.log('✅ Footage Folders loaded from central Config');
      Logger.log('   Count: ' + Object.keys(footage).length + ' products');
      
      // Count how many have URLs vs empty
      let withUrl = 0;
      let empty = 0;
      for (const key of Object.keys(footage)) {
        if (footage[key] && footage[key].includes('drive.google.com')) {
          withUrl++;
        } else {
          empty++;
        }
      }
      Logger.log('   With URLs: ' + withUrl);
      Logger.log('   Empty: ' + empty);
      passed++;
    } else {
      Logger.log('⚠️ Footage Folders not in central Config - using hardcoded fallback');
      warnings++;
    }
  } catch (e) {
    Logger.log('❌ Footage Folders error: ' + e.message);
    failed++;
  }
  
  // Test 3: Template constants are set correctly
  Logger.log('');
  Logger.log('━━━ TEST 3: TEMPLATE CONSTANTS ━━━');
  try {
    const checks = [
      { name: 'IMG_TEMPLATE_ID', value: IMG_TEMPLATE_ID },
      { name: 'VID_TEMPLATE_ID', value: VID_TEMPLATE_ID },
      { name: 'ITR_TEMPLATE_ID', value: ITR_TEMPLATE_ID },
      { name: 'LOC_TEMPLATE_ID', value: LOC_TEMPLATE_ID },
      { name: 'YTTOFB_TEMPLATE_ID', value: YTTOFB_TEMPLATE_ID },
      { name: 'ANDROMEDA_VID_TEMPLATE_ID', value: ANDROMEDA_VID_TEMPLATE_ID }
    ];
    
    let allSet = true;
    for (const check of checks) {
      if (check.value && check.value.length > 20) {
        Logger.log('   ✅ ' + check.name + ': Set');
      } else {
        Logger.log('   ❌ ' + check.name + ': Missing or invalid');
        allSet = false;
      }
    }
    
    if (allSet) {
      Logger.log('✅ All template constants are set');
      passed++;
    } else {
      Logger.log('❌ Some template constants are missing');
      failed++;
    }
  } catch (e) {
    Logger.log('❌ Template constants error: ' + e.message);
    failed++;
  }
  
  // Test 4: FOLDER_LINKS is set correctly
  Logger.log('');
  Logger.log('━━━ TEST 4: FOLDER_LINKS CONSTANT ━━━');
  try {
    if (typeof FOLDER_LINKS === 'object' && Object.keys(FOLDER_LINKS).length > 0) {
      Logger.log('✅ FOLDER_LINKS is set');
      Logger.log('   Products: ' + Object.keys(FOLDER_LINKS).length);
      
      // Test a few known products
      const testProducts = ['CamTrix', 'AliveBlue', 'WellHeater'];
      for (const prod of testProducts) {
        if (FOLDER_LINKS[prod]) {
          Logger.log('   ✅ ' + prod + ': Has footage URL');
        } else {
          Logger.log('   ⚠️ ' + prod + ': No footage URL (may be intentional)');
        }
      }
      passed++;
    } else {
      Logger.log('❌ FOLDER_LINKS is empty or not set');
      failed++;
    }
  } catch (e) {
    Logger.log('❌ FOLDER_LINKS error: ' + e.message);
    failed++;
  }
  
  // Summary
  Logger.log('');
  Logger.log('╔══════════════════════════════════════════════════════════════════╗');
  Logger.log('║  ✅ PASSED: ' + passed + '   ⚠️ WARNINGS: ' + warnings + '   ❌ FAILED: ' + failed + '                      ║');
  Logger.log('╚══════════════════════════════════════════════════════════════════╝');
  
  SpreadsheetApp.getUi().alert(
    'Ad Naming Config Diagnostics',
    '✅ Passed: ' + passed + '\n⚠️ Warnings: ' + warnings + '\n❌ Failed: ' + failed + '\n\nSee Execution Log for details.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  return { passed, warnings, failed };
}

function logErrorToHub_(errorType, message) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID); // Already defined in your script
    let sheet = ss.getSheetByName('Error Log');
    
    if (!sheet) {
      sheet = ss.insertSheet('Error Log');
      sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'System', 'Error Type', 'Message', 'User']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f4cccc');
    }
    
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 5).setValues([[
      new Date().toISOString(),
      'Ad Naming Sheet',  // ← This identifies the system
      errorType,
      message.substring(0, 500), // Limit message length
      user
    ]]);
  } catch (e) {
    Logger.log('Could not log error to Hub: ' + e);
  }
}

function testErrorLogging() {
  logErrorToHub_('Test', 'This is a test error');
  Logger.log('Done - check Config sheet Error Log tab');
}

/**
 * Run this to force Drive authorization prompt
 */
function forceAuthorizeDrive() {
  // This should trigger the permissions dialog
  const root = DriveApp.getRootFolder();
  SpreadsheetApp.getUi().alert('✅ Drive access authorized!\n\nFolder: ' + root.getName() + '\n\nYou can now use Standard Setup.');
}

function testDriveAccess() {
  const ui = SpreadsheetApp.getUi();
  
  // Test 1: Basic Drive access
  try {
    const root = DriveApp.getRootFolder();
    Logger.log('✅ Root folder access: ' + root.getName());
  } catch (e) {
    ui.alert('❌ Cannot access Drive root: ' + e.message);
    return;
  }
  
  // Test 2: Check a specific product folder from config
  try {
    const folders = getProductLocaleFoldersFromConfig_();
    const firstProduct = Object.keys(folders)[0];
    const firstLocale = Object.keys(folders[firstProduct])[0];
    const folderUrl = folders[firstProduct][firstLocale];
    
    if (folderUrl) {
      const folderId = folderUrl.match(/[-\w]{25,}/)?.[0];
      if (folderId) {
        const folder = DriveApp.getFolderById(folderId);
        Logger.log('✅ Can access: ' + folder.getName());
        ui.alert('✅ Drive access works!\n\nTested folder: ' + folder.getName());
        return;
      }
    }
  } catch (e) {
    ui.alert('❌ Cannot access product folder!\n\n' + e.message + '\n\nYour colleague needs to be granted access to the Ad Notes folders in Drive.');
    return;
  }
  
  ui.alert('✅ Basic Drive access works, but no folders configured to test.');
}