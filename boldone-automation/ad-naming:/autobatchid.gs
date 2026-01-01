/*******************************************************
 * Auto Batch ID Detector & Incrementer
 * (runs only in "Kristijonas" tab)
 * Adds soft highlight while processing.
 *******************************************************/

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
      'Ad Naming Sheet',
      errorType,
      String(message).substring(0, 500),
      user
    ]]);
  } catch (e) {
    Logger.log('Could not log error to Hub: ' + e);
  }
}

function handleAutoBatchTrigger(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // ✅ Add allowed sheet tabs here
    const ALLOWED_SHEETS = getAllowedSheetsFromConfig_();

    // Run only if sheet is in the allowed list
    if (!ALLOWED_SHEETS.includes(sheetName)) return;


    const cellA1 = e.range.getA1Notation();
    const MAP = {
      'C7':  { product: 'C4',  platform: 'C5',  locale: 'C6'  },
      'C20': { product: 'C17', platform: 'C18', locale: 'C19' },
      'C35': { product: 'C32', platform: 'C33', locale: 'C34' }
    };

    // Skip if user directly edits batch cell (manual override)
    if (cellA1 in MAP) return;

    // Detect which group changed
    let targetKey = null;
    for (const [key, ref] of Object.entries(MAP)) {
      if ([key, ref.product, ref.platform, ref.locale].includes(cellA1)) {
        targetKey = key;
        break;
      }
    }
    if (!targetKey) return;

    const ref = MAP[targetKey];
    const product  = sheet.getRange(ref.product).getDisplayValue().trim();
    const platform = sheet.getRange(ref.platform).getDisplayValue().trim();
    const locale   = sheet.getRange(ref.locale).getDisplayValue().trim();
    const targetCell = sheet.getRange(targetKey);

    if (!product || !platform || !locale) {
      targetCell.setBackground('#F4CCCC').setValue('⚠️ Missing info');
      return;
    }

    // 🟡 Animated "Checking..." feedback
    targetCell.setBackground('#FFF2CC').setValue('🔄 Checking…');
    SpreadsheetApp.flush();

    const start = Date.now();
    let nextId;

    try {
      nextId = getNextBatchIdSmart_(product, platform, locale);
        } catch (innerErr) {
      const msg = String(innerErr).replace('Error: ', '').trim();
      Logger.log(`❌ [BatchID] ${msg}`);

      // 🟥 Compact 3–4 word summary in the cell
      let shortMsg = '❌ Error';
      if (/product/i.test(msg)) shortMsg = '❌ Product folder';
      else if (/platform/i.test(msg)) shortMsg = '❌ Platform folder';
      else if (/ads/i.test(msg)) shortMsg = '❌ Ads folder';
      else if (/locale/i.test(msg)) shortMsg = '❌ Locale folder';
      else if (/year/i.test(msg)) shortMsg = '❌ Year folder';
      else if (/base/i.test(msg)) shortMsg = '❌ Base folder';
      else if (/missing info/i.test(msg)) shortMsg = '❌ Missing info';

      targetCell.setBackground('#F4CCCC').setValue(shortMsg);

      // Full text still in toast
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `❌ ${msg}`,
        'Batch ID Lookup Failed',
        6
      );
      return;
    }

    // Keep the "Checking..." visible for at least 1s
    const elapsed = Date.now() - start;
    if (elapsed < 1000) Utilities.sleep(1000 - elapsed);

    // ✅ Success
    targetCell.setValue(nextId).setBackground('#D9EAD3');
    targetCell.setNote(null); // remove any leftover error notes
    SpreadsheetApp.flush();
    Utilities.sleep(600);
    targetCell.setBackground('#FFFFFF');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Next Batch ID for ${product}/${locale}: ${nextId}`,
      'Batch ID Ready',
      4
    );
    Logger.log(`✅ [BatchID] ${product}/${platform}/${locale} → ${nextId}`);

  } catch (err) {
    const msg = String(err).replace('Error: ', '').trim();
    Logger.log('❌ Auto Batch Trigger crash: ' + msg);
    const cell = e.range;
    cell.setBackground('#F4CCCC').setValue(`❌ ${msg}`.substring(0, 80));
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `❌ ${msg}`,
      'Script Error',
      8
    );
  }
}

/**
 * Finds the highest batch number under:
 * /Product/Platform/<Ads*>/Locale/<Year or main>  → returns +1
 * ⚠️ Always reads fresh from Drive (no cache)
 */
function getNextBatchIdSmart_(product, platform, locale) {
  Logger.log(`🔍 Scanning fresh for next batch ID: ${product}, ${platform}, ${locale}`);

  // 1️⃣ Find product root folder
  const root = findProductRoot_(product);
  if (!root) throw new Error(`No base folder found for product ${product}`);
  Logger.log(`✅ Product root: ${root.getName()}`);

  // 2️⃣ Find platform folder (fuzzy match, allows "Meta", "Facebook", etc.)
  const platFolder = findPlatformFolder_(root, platform);
  if (!platFolder)
    throw new Error(`Platform folder ${platform} not found under ${product}`);
  Logger.log(`✅ Platform folder: ${platFolder.getName()}`);

  // 3️⃣ Find Ads folder
  let adsFolder = findSubfolderContains_(platFolder, 'ads');
  if (!adsFolder) {
    Logger.log(`⚠️ No 'Ads' folder found under ${product}/${platform} → using platform root instead`);
    adsFolder = platFolder;
  } else {
    Logger.log(`✅ Ads folder: ${adsFolder.getName()}`);
  }

  // 🆕 3.1️⃣ Check if Ads folder contains product-specific subfolder (e.g., "Clairu_M2 Ads")
  const possibleProductSub = findSubfolderContains_(adsFolder, product);
  if (possibleProductSub) {
    Logger.log(`🔁 Going deeper: found product-level Ads folder → ${possibleProductSub.getName()}`);
    adsFolder = possibleProductSub;
  }

  // 4️⃣ Find locale folder with LATAM/ES aliasing support
  let localeFolder = null;

  // ✅ NEW: Handle LATAM/ES aliases
  if (locale === 'LATAM' || locale === 'ES') {
    Logger.log(`🔍 Looking for LATAM or ES folder (they're aliases)...`);
    
    // Try to find either LATAM or ES folder
    localeFolder = findSubfolder_(adsFolder, 'LATAM');
    if (!localeFolder) {
      localeFolder = findSubfolder_(adsFolder, 'ES');
    }
    
    if (localeFolder) {
      Logger.log(`✅ Found alias folder: ${localeFolder.getName()} (requested ${locale})`);
    }
  } else {
    // Standard lookup for other locales
    localeFolder = findSubfolder_(adsFolder, locale);
  }

  if (!localeFolder)
    throw new Error(`Locale folder ${locale} not found under ${product}/${platform}/Ads`);
  Logger.log(`✅ Locale folder: ${localeFolder.getName()}`);

  // 5️⃣ Choose either current-year folder or fallback to main
  const year = String(new Date().getFullYear());
  let searchFolder = findSubfolder_(localeFolder, year);

  // ✅ FIX: If we accidentally landed in a batch folder, go back up
  if (searchFolder) {
    const folderName = searchFolder.getName();
    // Check if this looks like a batch folder (contains locale + number pattern)
    if (/_[A-Z]{2,6}_\d+_/i.test(folderName)) {
      Logger.log(`⚠️ Landed in batch folder ${folderName}, using parent instead`);
      searchFolder = localeFolder;
    }
  }

  if (!searchFolder) searchFolder = localeFolder;

  Logger.log(`✅ Searching in: ${searchFolder.getName()}`);

  // 6️⃣ Look through subfolders for pattern like "_US_123_..."
  let maxId = 0;
  const subfolders = searchFolder.getFolders();
  const pattern = new RegExp(`_${locale}_(\\d+)(?:_|\\(|$)`, 'i');

  while (subfolders.hasNext()) {
    const f = subfolders.next();
    const name = f.getName();
    Logger.log('🔍 Folder name: ' + name);
    const match = name.match(pattern);
    if (match && match[1]) {
      const num = parseInt(match[1], 10);
      if (!isNaN(num)) {
        Logger.log('   ✅ Matched! ID: ' + num);
        maxId = Math.max(maxId, num);
      }
    }
  }

  Logger.log(`✅ Highest ID found: ${maxId}`);
  return maxId + 1;
}



// ---------- helper utilities ----------

function findPlatformFolder_(productFolder, platformName) {
  const it = productFolder.getFolders();
  const platformLower = platformName.toLowerCase();

  while (it.hasNext()) {
    const f = it.next();
    const fname = f.getName().toLowerCase();

    // Allow common Meta/Facebook variations
    const isMeta = ['meta', 'facebook', 'facebook meta', '.meta', 'meta.'].some(v =>
      fname.includes(v)
    );
    const isPlatform =
      fname.includes(platformLower) ||
      (platformLower === 'fb' && isMeta) ||
      (platformLower === 'meta' && fname.includes('fb'));


    if (isPlatform) return f;
  }

  return null;
}

/**
 * Finds subfolder with loose match (case-insensitive, punctuation-tolerant).
 * Accepts variants like ".US", "_US", "US_", "US 2025" etc.
 */
function findSubfolder_(parentFolder, name) {
  if (!parentFolder || !name) return null;
  const target = name.toLowerCase().replace(/[^a-z0-9]/g, ''); // strip punctuation
  const it = parentFolder.getFolders();

  while (it.hasNext()) {
    const f = it.next();
    const cleanName = f.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanName.includes(target)) return f;
  }
  return null;
}

/**
 * Same as above but looks for subfolder containing a keyword fragment.
 * E.g. "ads" matches "Ads", "ads folder", "Meta Ads 2025"
 */
function findSubfolderContains_(parentFolder, fragment) {
  if (!parentFolder || !fragment) return null;
  const target = fragment.toLowerCase().replace(/[^a-z0-9]/g, '');
  const it = parentFolder.getFolders();

  while (it.hasNext()) {
    const f = it.next();
    const cleanName = f.getName().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (cleanName.includes(target)) return f;
  }
  return null;
}

function findProductRoot_(productName) {
  const ROOT_ID = '151s5U5ZemSfaZ4JKXRWm_MhRiJzw0hHq';
  const root = DriveApp.getFolderById(ROOT_ID);

  // Products that live under "02 Brands"
  const BRANDS_PRODUCTS = [
    'Shapewear', 'Jewelry', 'Bryt', 'EmaWell',
    'Planet Wash', 'AuraNaturals', 'BioVitals'
  ];

  const productLower = String(productName || '').toLowerCase();

  // Decide which main category to search
  const isBrand = BRANDS_PRODUCTS.some(p =>
    productLower.includes(String(p).toLowerCase())
  );
  const categoryFolderName = isBrand ? '02 Brands' : '01 Fast Ecom';

  Logger.log(`📁 Searching for product root: ${productName}`);
  Logger.log(`🔍 Looking under: ${categoryFolderName}`);

  const categoryFolder = findSubfolderContains_(root, categoryFolderName);
  if (!categoryFolder) {
    throw new Error(`Category folder ${categoryFolderName} not found in main root`);
  }
  Logger.log(`✅ Found category folder: ${categoryFolder.getName()}`);

  // Get product meta (ID + human name) from PRODUCT_INDEX
  const meta = PRODUCT_INDEX[productName] || {};
  const idLower    = meta.id   ? meta.id.toLowerCase()   : '';
  const humanLower = meta.name ? meta.name.toLowerCase() : '';

  if (idLower) {
    Logger.log(`🎯 Expecting folder containing "${meta.id}" and something like "${meta.name}"`);
  }

  // Pass 1: prefer folders that start with the product ID (e.g. "A10_...")
  let fallback = null;
  const it = categoryFolder.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const nameLower = f.getName().toLowerCase();

    const hasId    = idLower && nameLower.indexOf(idLower) !== -1;
    const startsId = idLower && nameLower.indexOf(idLower) === 0;
    const hasHuman = humanLower && nameLower.indexOf(humanLower) !== -1;

    // Best case: starts with ID → e.g. "A10_TV antena"
    if (startsId) {
      Logger.log(`✅ Found product folder: ${f.getName()}`);
      return f;
    }

    // Keep a fallback candidate if it at least contains the ID
    if (hasId && !fallback) fallback = f;
  }

  // Pass 2: use the best fallback found by ID
  if (fallback) {
    Logger.log(`✅ Fallback product folder by ID: ${fallback.getName()}`);
    return fallback;
  }

  // Pass 3: last resort – match by product name only
  const it2 = categoryFolder.getFolders();
  while (it2.hasNext()) {
    const f = it2.next();
    const nameLower = f.getName().toLowerCase();
    if (nameLower.indexOf(productLower) !== -1) {
      Logger.log(`✅ Fallback product folder by name: ${f.getName()}`);
      return f;
    }
  }

  Logger.log(`❌ Could not find folder for "${productName}"`);
  throw new Error(`No base folder found for product ${productName}`);
}

function refreshAllBatchIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return;

  const MAP = {
    'C7':  { product: 'C4',  platform: 'C5',  locale: 'C6'  },
    'C20': { product: 'C17', platform: 'C18', locale: 'C19' },
    'C35': { product: 'C32', platform: 'C33', locale: 'C34' }
  };

  SpreadsheetApp.getActiveSpreadsheet().toast('🔄 Refreshing all Batch IDs…', 'Batch ID Update', 3);

  for (const [targetKey, ref] of Object.entries(MAP)) {
    const product  = sheet.getRange(ref.product).getDisplayValue().trim();
    const platform = sheet.getRange(ref.platform).getDisplayValue().trim();
    const locale   = sheet.getRange(ref.locale).getDisplayValue().trim();
    const targetCell = sheet.getRange(targetKey);

    if (!product || !platform || !locale) {
      targetCell.setBackground('#F4CCCC').setValue('⚠️ Missing info');
      continue;
    }

    try {
      targetCell.setBackground('#FFF2CC').setValue('🔄 Checking…');
      SpreadsheetApp.flush();
      const nextId = getNextBatchIdSmart_(product, platform, locale);
      targetCell.setValue(nextId).setBackground('#D9EAD3').setNote(null);
      SpreadsheetApp.flush();
      Utilities.sleep(600);
      targetCell.setBackground('#FFFFFF');
    } catch (err) {
      const msg = String(err).replace('Error: ', '').trim();
      Logger.log(`❌ [RefreshAll] ${msg}`);
      targetCell.setValue(`❌ ${msg.substring(0, 60)}`).setBackground('#F4CCCC');
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ All Batch IDs refreshed!', 'Done', 3);
}
// 🔄 Refresh only Batch ID at C7
function refreshBatchId1() {
  refreshSingleBatchId('C7');
}

// 🔄 Refresh only Batch ID at C20
function refreshBatchId2() {
  refreshSingleBatchId('C20');
}

// 🔄 Refresh only Batch ID at C35
function refreshBatchId3() {
  refreshSingleBatchId('C35');
}

// ⚙️ Shared function (used by all three)
function refreshSingleBatchId(targetKey) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const MAP = {
    'C7':  { product: 'C4',  platform: 'C5',  locale: 'C6'  },
    'C20': { product: 'C17', platform: 'C18', locale: 'C19' },
    'C35': { product: 'C32', platform: 'C33', locale: 'C34' }
  };
  
  const ref = MAP[targetKey];
  const product  = sheet.getRange(ref.product).getDisplayValue().trim();
  const platform = sheet.getRange(ref.platform).getDisplayValue().trim();
  const locale   = sheet.getRange(ref.locale).getDisplayValue().trim();
  const targetCell = sheet.getRange(targetKey);

  if (!product || !platform || !locale) {
    targetCell.setValue('⚠️ Missing info').setBackground('#F4CCCC');
    return;
  }

  try {
    targetCell.setBackground('#FFF2CC').setValue('🔄 Checking…');
    SpreadsheetApp.flush();
    const nextId = getNextBatchIdSmart_(product, platform, locale);
    targetCell.setValue(nextId).setBackground('#D9EAD3').setNote(null);
    SpreadsheetApp.flush();
    Utilities.sleep(600);
    targetCell.setBackground('#FFFFFF');
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Batch ID updated for ${product}`, 'Done', 3);
  } catch (err) {
    const msg = String(err).replace('Error: ', '').trim();
    targetCell.setValue(`❌ ${msg.substring(0, 60)}`).setBackground('#F4CCCC');
    SpreadsheetApp.getActiveSpreadsheet().toast(`❌ ${msg}`, 'Batch ID Error', 4);
  }
}