/**
 * ================================================================================
 * BOLD.ONE CONFIG SHEET GENERATOR
 * ================================================================================
 * 
 * HOW TO USE:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any code in Code.gs
 * 4. Paste this entire script
 * 5. Save (Ctrl+S)
 * 6. Refresh the Google Sheet
 * 7. Click the new menu: "⚙️ Config Setup" → "🚀 Create All Config Tabs"
 * 8. Grant permissions when prompted
 * 9. Wait for completion (~30 seconds)
 * 
 * ================================================================================
 */

// ✅ YOUR CONFIG SHEET ID (already filled in!)
const CONFIG_SHEET_ID = '1TO3ZRxGnkLWO2hJn9odoApMl6WsjwiKZ48s7mkl54kk';

const CONFIG_CACHE_DURATION = 300; // 5 minutes

/**
 * Get valid products from central config
 */
function getProductsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_products');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  
  const products = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === true) { // Column G = Active
      products.push(data[i][0]); // Column A = Product Key
    }
  }
  
  cache.put('dnc_products', JSON.stringify(products), CONFIG_CACHE_DURATION);
  Logger.log('✅ Loaded ' + products.length + ' products from central config');
  return products;
}

/**
 * Get valid platforms from central config
 */
function getPlatformsFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_platforms');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  
  // Find Platforms column
  const headers = data[0];
  const platformCol = headers.indexOf('Platforms');
  
  const platforms = [];
  if (platformCol >= 0) {
    for (let i = 1; i < data.length; i++) {
      const val = data[i][platformCol];
      if (val && val.toString().trim() !== '') {
        platforms.push(val.toString().trim());
      }
    }
  }
  
  cache.put('dnc_platforms', JSON.stringify(platforms), CONFIG_CACHE_DURATION);
  Logger.log('✅ Loaded ' + platforms.length + ' platforms from central config');
  return platforms;
}

/**
 * Get valid locales from central config
 */
function getLocalesFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_locales');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  
  // Find Locales column
  const headers = data[0];
  const localeCol = headers.indexOf('Locales');
  
  const locales = [];
  if (localeCol >= 0) {
    for (let i = 1; i < data.length; i++) {
      const val = data[i][localeCol];
      if (val && val.toString().trim() !== '') {
        locales.push(val.toString().trim());
      }
    }
  }
  
  cache.put('dnc_locales', JSON.stringify(locales), CONFIG_CACHE_DURATION);
  Logger.log('✅ Loaded ' + locales.length + ' locales from central config');
  return locales;
}

/**
 * Get folders to monitor from central config
 */
function getFoldersToMonitorFromConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dnc_folders');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const sheet = ss.getSheetByName('Valid Values');
  const data = sheet.getDataRange().getValues();
  
  // Find Monitor Folders column
  const headers = data[0];
  const folderCol = headers.indexOf('Monitor Folders');
  
  const folders = [];
  if (folderCol >= 0) {
    for (let i = 1; i < data.length; i++) {
      const val = data[i][folderCol];
      if (val && val.toString().trim() !== '') {
        folders.push(val.toString().trim());
      }
    }
  }
  
  cache.put('dnc_folders', JSON.stringify(folders), CONFIG_CACHE_DURATION);
  Logger.log('✅ Loaded ' + folders.length + ' monitor folders from central config');
  return folders;
}

/**
 * Clear config cache (add to menu for manual refresh)
 */
function clearDNCConfigCache() {
  CacheService.getScriptCache().removeAll(['dnc_products', 'dnc_platforms', 'dnc_locales', 'dnc_folders']);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Config cache cleared! Next run will reload from central config.', 'Cache Cleared', 3);
  Logger.log('✅ Config cache cleared');
}

/**
 * Test config connection (run this to verify everything works)
 */
function testConfigConnection() {
  try {
    Logger.log('🔍 Testing connection to central config...\n');
    
    const products = getProductsFromConfig_();
    Logger.log('✅ Products: ' + products.length + ' loaded');
    Logger.log('   Sample: ' + products.slice(0, 5).join(', '));
    
    const platforms = getPlatformsFromConfig_();
    Logger.log('\n✅ Platforms: ' + platforms.join(', '));
    
    const locales = getLocalesFromConfig_();
    Logger.log('\n✅ Locales: ' + locales.join(', '));
    
    const folders = getFoldersToMonitorFromConfig_();
    Logger.log('\n✅ Monitor Folders: ' + folders.length + ' loaded');
    
    Logger.log('\n🎉 All config loaded successfully!');
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Config connection working! Check Execution Log for details.', 'Test Passed', 5);
    
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert('❌ Config Test Failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Config Setup')
    .addItem('🚀 Create All Config Tabs', 'createAllConfigTabs')
    .addItem('🎨 Reformat All Tabs', 'reformatAllTabs')
    .addItem('🗑️ Clear & Rebuild', 'clearAndRebuild')
    .addToUi();
}

/**
 * Main function - creates all config tabs
 */
function createAllConfigTabs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Confirm
  const response = ui.alert(
    '🚀 Create Config Tabs',
    'This will create 6 configuration tabs with all your product and team data.\n\nExisting tabs with the same names will be replaced.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Show progress
  ss.toast('Creating Products tab...', '⚙️ Setup', -1);
  createProductsTab(ss);
  
  ss.toast('Creating Product Folders tab...', '⚙️ Setup', -1);
  createProductFoldersTab(ss);
  
  ss.toast('Creating Team Members tab...', '⚙️ Setup', -1);
  createTeamMembersTab(ss);
  
  ss.toast('Creating ClickUp Config tab...', '⚙️ Setup', -1);
  createClickUpConfigTab(ss);
  
  ss.toast('Creating Valid Values tab...', '⚙️ Setup', -1);
  createValidValuesTab(ss);
  
  ss.toast('Creating Sync Log tab...', '⚙️ Setup', -1);
  createSyncLogTab(ss);
  
  // Delete default Sheet1 if it exists and is empty
  try {
    const sheet1 = ss.getSheetByName('Sheet1');
    if (sheet1 && sheet1.getLastRow() === 0) {
      ss.deleteSheet(sheet1);
    }
  } catch (e) {}
  
  // Activate Products tab
  ss.getSheetByName('Products').activate();
  
  ss.toast('✅ All config tabs created successfully!', '⚙️ Setup Complete', 5);
  
  ui.alert('✅ Setup Complete!', 
    'All 6 configuration tabs have been created:\n\n' +
    '• Products (43 products)\n' +
    '• Product Folders (20 products × 12 locales)\n' +
    '• Team Members (11 members)\n' +
    '• ClickUp Config (6 settings)\n' +
    '• Valid Values (platforms, locales, aliases)\n' +
    '• Sync Log (empty, auto-populated)\n\n' +
    '📋 Sheet ID: ' + ss.getId() + '\n\n' +
    'Copy this ID - you\'ll need it for the automation scripts!',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// PRODUCTS TAB
// ============================================================================

function createProductsTab(ss) {
  let sheet = ss.getSheetByName('Products');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Products');
  
  // Headers
  const headers = ['Product Key', 'Display Name', 'Product ID', 'Product Full Name', 'Target CPA (€)', 'Product URL', 'Active', 'Notes'];
  
  // Data
  const data = [
    ['AliveBlue', 'AliveBlue', 'A35', 'Hydrogen Bottle', 70, 'https://getaliveblue.com/', true, ''],
    ['AuraNaturals', 'AuraNaturals', 'B02', 'Essential Oils', 47.5, 'https://auranaturals.com/', true, ''],
    ['BioVitals', 'BioVitals', 'B01', 'Supplements', 60, '', false, 'Inactive'],
    ['Blitron', 'Blitron', 'A15', 'Mini Magnetic Flashlight', 60, 'https://getblitron.com/', false, 'Inactive'],
    ['BRYT', 'BRYT', 'B05', 'Mushroom Coffee', 60, 'https://www.brytme.com/', false, 'Inactive'],
    ['CamTrix', 'CamTrix', 'A03', 'Mini Wifi Camera', 60, 'https://getcamtrix.com/', true, ''],
    ['Clairu', 'Clairu', 'A14', 'Ionic Air Purifier', 60, 'https://clairu.com/', false, 'Legacy - use Clairu_M2'],
    ['Clairu_M2', 'Clairu_M2', 'A14', 'Ionic Air Purifier', 60, 'https://clairu.com/', true, ''],
    ['Cleanlix', 'Cleanlix', 'A25', 'Cleaning Spray Tablets', 60, 'https://nuroclean.com/', true, ''],
    ['CleanlixMold', 'CleanlixMold', 'A33', 'Mold Remover Gel', 60, 'https://cleanlix.com/', true, ''],
    ['CleanlixPowder', 'CleanlixPowder', 'A32', 'Toilet Cleaning Powder', 60, 'https://cleanlix.com/', true, ''],
    ['DermaSonic+', 'DermaSonic+', 'A36', 'Face Hair Remover', 60, '', false, 'Inactive'],
    ['EmaWell', 'EmaWell', 'B04', 'Supplements', 60, '', false, 'Inactive'],
    ['EpiCooler', 'EpiCooler', 'A52', 'AC Cooler', 90, 'https://get-epicooler.online/', true, ''],
    ['Epibella', 'Epibella', 'A30', 'Face Epilator', 60, 'https://epibella.com/', false, 'Inactive'],
    ['Flixy', 'TellyStick', 'A43', 'TV Streaming Stick', 55, 'https://tellystick.checkoutera.com/', false, 'LEGACY - use TellyStick'],
    ['Guardality', 'Guardality', 'A39', 'Card Protector', 50, 'https://guardality.com/', true, 'Formerly ArmouredCard'],
    ['HammerDex', 'HammerDex', 'A04', 'Window Breaker', 60, 'https://breeker.com/', false, 'Inactive'],
    ['Heatoor', 'Heatoor', 'A53', 'Heatoor Heater', 100, 'https://heatoor.com/', true, ''],
    ['JayShoes', 'JayShoes', 'A28', 'Running Shoes', 60, 'https://getjayshoes.com/', false, 'Inactive'],
    ['KatuChef', 'KatuChef', 'A37', 'Cutting Board', 75, 'https://katuchef.com/', true, ''],
    ['LoveAndShine', 'LoveAndShine', 'A27', 'Necklace Gift Box', 60, 'https://lovenshine.com/', false, 'Inactive'],
    ['LoweSkin', 'LoweSkin', 'A17', 'Laser Epilator', 60, 'http://getloweskin.com/', false, 'Inactive'],
    ['LumenLight', 'LumenLight', 'A11', 'Head Flashlight', 60, 'https://getlumenlight.com/', true, ''],
    ['NuraDerma', 'NuraDerma', 'A29', 'Face Mask', 60, 'https://nuraderma.com/', false, 'Inactive'],
    ['NuraFix', 'NuraFix', 'A23', 'Nano Spray', 60, 'https://nurafix.com/', false, 'Inactive'],
    ['OmniHear', 'OmniHear', 'A18', 'Hearing Aid', 65, 'https://get-omnihear.com/', false, 'Inactive'],
    ['OriBreeze', 'OriBreeze', 'A18', 'Mini Portable Cooler', 65, 'https://oribreeze.com/', true, ''],
    ['Ozoori', 'Ozoori', 'A41', 'Ozoori', 60, 'https://get-ozoori.com/', true, ''],
    ['Phoxfer', 'Phoxfer', 'A16', 'USB Backup Stick', 60, 'https://getmemovault.com/', false, 'Inactive'],
    ['PlanetWash', 'PlanetWash', 'B03', 'Detergent Sheets', 60, 'https://tryplanetwash.com/', false, 'Inactive'],
    ['PrimaFocus', 'PrimaFocus', 'A01', 'Adjustable Glasses', 60, 'https://buyprimafocus.com/article/en', false, 'Inactive'],
    ['Repellio', 'Repellio', 'A13', 'Ultrasonic Pest Repellant', 60, 'http://repellio.com/', false, 'Inactive'],
    ['SleepZee', 'SleepZee', 'A12', 'Anti-Snoring Mouthpiece', 60, 'https://sleepzee.com/', false, 'Inactive'],
    ['TellyStick', 'TellyStick', 'A43', 'TV Streaming Stick', 55, 'https://tellystick.checkoutera.com/', true, ''],
    ['TheraWolfNeuro', 'TheraWolfNeuro', 'A47', 'NeuroBalm', 75, 'https://therawolf.com/pp/neuro-all-04/', true, ''],
    ['TheraWolfRelief', 'TheraWolfRelief', 'A46', 'ReliefBalm', 75, 'https://therawolf.com/', true, ''],
    ['TVStick', 'TellyStick', 'A43', 'TV Streaming Stick', 55, 'https://tellystick.checkoutera.com/', false, 'LEGACY - use TellyStick'],
    ['VacuumSealer', 'VacuumSealer', 'A34', 'Vacuum Sealer', 60, '', false, 'Inactive'],
    ['WaveMax', 'WaveMax', 'A10', 'TV Antenna', 47.5, 'https://getwavemax.com/pp/en/', true, ''],
    ['WellHeater', 'WellHeater', 'A08', 'Heater', 65, 'https://blumeheat.com/', true, ''],
    ['WiggyDog', 'WiggyDog', 'A51', 'WiggyDog', 40, 'https://getwiggydog.com', true, ''],
    ['Zapfie', 'Zapfie', 'A20', 'Mosquito Lamp', 60, 'https://mosqishock.com/', false, 'Inactive']
  ];
  
  // Write data
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  // Column widths
  sheet.setColumnWidth(1, 130); // Product Key
  sheet.setColumnWidth(2, 130); // Display Name
  sheet.setColumnWidth(3, 80);  // Product ID
  sheet.setColumnWidth(4, 180); // Full Name
  sheet.setColumnWidth(5, 100); // Target CPA
  sheet.setColumnWidth(6, 280); // URL
  sheet.setColumnWidth(7, 60);  // Active
  sheet.setColumnWidth(8, 180); // Notes
  
  // Add checkboxes to Active column
  const activeRange = sheet.getRange(2, 7, data.length, 1);
  activeRange.insertCheckboxes();
  
  // Center align certain columns
  sheet.getRange(2, 3, data.length, 1).setHorizontalAlignment('center'); // Product ID
  sheet.getRange(2, 5, data.length, 1).setHorizontalAlignment('center'); // Target CPA
  sheet.getRange(2, 7, data.length, 1).setHorizontalAlignment('center'); // Active
  
  // Alternate row colors
  for (let i = 0; i < data.length; i++) {
    const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
    if (i % 2 === 1) {
      rowRange.setBackground('#f8f9fa');
    }
    // Gray out inactive products
    if (data[i][6] === false) {
      rowRange.setFontColor('#9aa0a6');
    }
  }
  
  // Add filter
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
  
  // Add data validation for CPA (must be positive number)
  const cpaRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setAllowInvalid(false)
    .setHelpText('CPA must be a positive number')
    .build();
  sheet.getRange(2, 5, data.length, 1).setDataValidation(cpaRule);
}

// ============================================================================
// PRODUCT FOLDERS TAB
// ============================================================================

function createProductFoldersTab(ss) {
  let sheet = ss.getSheetByName('Product Folders');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Product Folders');
  
  // Headers
  const headers = ['Product Key', 'US', 'UK', 'DE', 'CA', 'AU', 'LATAM', 'ES', 'FR', 'IT', 'IL', 'EU', 'MIX'];
  
  // Data - folder IDs extracted from your code
  const data = [
    ['AliveBlue', '15Jv3CdIEIvQwZ2QftSG5jIzlRzFJT9BR', '1fe1ihS9epIXe804TRnzamfqM4BUCx9dG', '1-2f-2vWOTj6CKCNKYJk0vYpNMrqqusRY', '1Rgo9rP1-R2lqGGerOk8TRI_Ynpkwt9KO', '', '1z_iMoelezniujBlajafC6tywCSZrelTJ', '1z_iMoelezniujBlajafC6tywCSZrelTJ', '', '', '1LLGaEOc2PyZquHuO6R-w5ZCiacb_aXJv', '', ''],
    ['AuraNaturals', '1gPzMMeoq9gkpM12yliQZtckAwoKTwuj9', '', '', '', '', '', '', '', '', '', '', ''],
    ['CamTrix', '1h8r5iKGUFCZEaCmVbUiMh5e3qsFv6XNe', '1KvrD8QVToCtFNtuV8-1lwWacdOVVM-hm', '1zt-XAG9oI-QohzRkUaYjt7j5eg2hxCA0', '1sjHh4STdy_VkJyAmShvDHZX18a28lURI', '10_EDKWb1PFTBXRKAnFOLBFwh7hKzHDB_', '1tsXOMH2Ze78AySS4RxGlxnmtvhpZi3DQ', '1tsXOMH2Ze78AySS4RxGlxnmtvhpZi3DQ', '', '', '', '', ''],
    ['Clairu_M2', '1HHmWxdUCkplT5vaS-X1Cz6JwBpaDMREH', '', '1lKkC8TxAnPtX6xX_bZ0QiGBcDMOY9p9t', '', '', '1pn7Nih_cuok1GLCUzMZdnhTu5njSJyvt', '1pn7Nih_cuok1GLCUzMZdnhTu5njSJyvt', '', '', '', '', ''],
    ['Cleanlix', '1wjRAgl9ySvpU4V20k2WRDCnxyKC1FfPe', '1fgkdeDEb1UwLhhep_8JB8OpqoK98GRAO', '', '', '1mygxPA3JAWF56jTjZBmTBJ5kwr5uTvHY', '', '', '', '', '', '', ''],
    ['CleanlixMold', '1iAEDeVdCE47GKoBmVauc6Qe6JEX5tN6f', '1HDCe48Jkd0UDAyxYdDi7vwS6uXaUvu_L', '15sI18onKa1FhZVd2T1zpug1OvkshMckx', '', '1yuVXoQ5b8S9zFJfmrUKzZynSrZk2fuhP', '1Uid3-pbl6MBodRHhjji3BKv3TyIJJSiO', '1Uid3-pbl6MBodRHhjji3BKv3TyIJJSiO', '', '', '', '', ''],
    ['CleanlixPowder', '11V8pbXvt723DJfsozxplwzdoC9h4aoWo', '', '1-8uC_H2EMcst1igbBmdnMUT0pGx_o4J4', '', '1yHMRa1lc3OtgwhFsb8zGvfgZEGwJiU5v', '1iAWiiBq_PUsM48Yo9rZdEyzLEGQ7vaPM', '1iAWiiBq_PUsM48Yo9rZdEyzLEGQ7vaPM', '', '', '', '', ''],
    ['EpiCooler', '1aOeoLUw9LVi9tkxlourQYoq7yX1tZOTx', '', '1XkA6hqF6zfBRxX7gYxucW7J8PF8s8Lyb', '', '1IzdntHrfTPAcx_3R94FrMDjVtcOXJ8Gr', '1-7AclsqDmrmJt7AeLk04WdFfkdpjXcf9', '1-7AclsqDmrmJt7AeLk04WdFfkdpjXcf9', '', '', '', '', ''],
    ['Guardality', '1tAGgYVfnlNm4iuBQsdwxi3Gmsm7tE6iK', '', '1syzA6DF_QMCCMZ2YkCpEoH9OHa9oyqmP', '', '', '1-lGDwtMH-jPKgBx4a2QOqxkWb9rfCMac', '1-lGDwtMH-jPKgBx4a2QOqxkWb9rfCMac', '', '', '', '', '12D2E-OM8bKACCA9XB14oONUEd4-z6lWU'],
    ['Heatoor', '1735OWyLwQwx5wXIhPLrqkEUvtaTxIxx4', '124Wklv7HTwya92BvjVhhrDmGG7WuPfqS', '1DUsrSYe0t4NS7lO0ifVniqc5qHAB-YKh', '1S973oQ6t3Atpx78lSTO6htEVEWusRk7h', '1SxBkSVrNNE58FXXOSA9xNRNfR11RNClu', '', '', '', '1Su3-3Dqfpn1kQw-K8qhs7jLDLvjB99IU', '', '', ''],
    ['KatuChef', '1Eo9iTCR4w37ihpxuK_QVgdZAqhhvwK2D', '', '1k4KVtad_ji4AvGXHVn59ad9Be-dj5oWL', '', '', '1gIm9nE4NPFmlK6MvqCWfLyogXoVe85Hg', '1gIm9nE4NPFmlK6MvqCWfLyogXoVe85Hg', '', '', '', '', ''],
    ['LumenLight', '1CjNry2fguhtPRr9o5l7tQzQSuTX1XnZF', '', '17TQgEgn_BVR11FtEFnbwUQeOtt6qga-6', '', '1XyRgChqEgzTJHuAgqLySDRYH7tFYB5M_', '', '1iCtb_5Pd6T3zATzy8caowkqCxU8W_jte', '', '', '', '', ''],
    ['OriBreeze', '1_Y5GfcPxWvuZdL-7DuCuoioJhwDbYcKP', '1zScyGVTy6nme-RxbugqR8PjYZZUy7tu-', '1cspIVblJ9xM34S5rq41SWkk5Gl1Kd7O0', '', '1b1-jehcR6n7cAqYOuwrNUjfnrRbf8A3q', '19CdbeIIF1_dbe_V3-eFuZniECrDRiMre', '19CdbeIIF1_dbe_V3-eFuZniECrDRiMre', '1qqYlHxFFuus4Ji6U5gbFBK2UkTSr4YQr', '1we6yZgcNDt4hG9cb5wa7qMKSYkHhTK1k', '', '', ''],
    ['Ozoori', '14qJ_lZn_peNtl7b2q2MC-Me6ad9f3Fb8', '1vQGP5kXwGNlAH4NzHztgX5xDQfCl0-YU', '1lRkcJHKhkcVjyMV9EdkzVyEQ58urT7aC', '19wNItq4mm0eLeJRqFPM9s2r5-MABxFzq', '1cAGSPqu_PgwFEKV1v2IIVcluEVqCLJxz', '1k6RhLVvu8hoXVntJiKBcLIP6JBuQfvKo', '1k6RhLVvu8hoXVntJiKBcLIP6JBuQfvKo', '', '1hL-WwWBYyhcm3xsxHp4pXCps-_hMnM8g', '', '', ''],
    ['TellyStick', '1LTZEdfSAzsv-rYKkDd1PAqfBiWnv-aS1', '1QHpOJ0T0n61hjZQNOPO0jmuWXgqCBI1B', '14USHHtrHUZ5Ca6oqPkJxv3d13QAgnVk6', '1KBc6DAUSeWEnJqv18VqEhvXfwbp7kLHf', '1w2-v3zV_yzHH42ktKOMLEPla8_uUapX2', '1bOP2lc4Mb0_kL8CSDnu2cFd_pMIw9BYh', '1TCdUzioA_ZVxTmBU52JO9uPG7w9HwLTy', '', '', '', '', ''],
    ['TheraWolfNeuro', '1pSLBJzREOCgS-zSif81szl3lR4uM4Z-v', '', '1RyIpNP0N8kkevItG5-xF2G-nRXJI6HpB', '', '', '1EBHphXw1CaiM5l-kubs2qAaXlJWyqCpG', '1EBHphXw1CaiM5l-kubs2qAaXlJWyqCpG', '', '', '', '', ''],
    ['TheraWolfRelief', '1eN3ZVo9pMulfUohuxMVaG1oYyx_23ewC', '', '1NHgmuI0lLa10UKmzCnFhyz9m638-BNH0', '', '', '1qw7C_d8e0hEp4MK1Uy1IWx0F6E7DmS7H', '1qw7C_d8e0hEp4MK1Uy1IWx0F6E7DmS7H', '', '', '', '', ''],
    ['WaveMax', '1JqFJ-_RAMUzUYV_71rxlv5_kHUTdf67X', '', '1nehAxA-G5Avd-UB4aoB_r4fDqvIJe1md', '', '', '1Ps9Af5hbCu5iDor-G58Qk15-DEW9zrZW', '1Ps9Af5hbCu5iDor-G58Qk15-DEW9zrZW', '', '', '', '', ''],
    ['WellHeater', '1iQ15AKMAM80jdkYjw5J05mjyCN0TPEQ1', '1I0svtjYNcgHL0l6eNzQyNdi5JUQviaZv', '1Aj3ZVdiL59nU5n10MUdcX--zRzS3AgHj', '', '1I1lH1nBRaZH_yX7FZjF6lF54npbFIiDF', '', '19fJZlVmgI9f1m80bqnmlySTEmbOokJxp', '1cgQuWUStUIjzRU1QPjdvxR47NxQzGzWo', '1tr3Yz0St5WMTqlOmis7zk0XKm-CuWpmU', '', '', ''],
    ['WiggyDog', '1WSF90kHIa7sRV6iih8XflRSANMjN3p6N', '', '', '', '', '', '', '', '', '', '', '']
  ];
  
  // Write data
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // Freeze header and first column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  // Column widths
  sheet.setColumnWidth(1, 130); // Product Key
  for (let i = 2; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 280); // Locale columns - wide for folder IDs
  }
  
  // Alternate row colors
  for (let i = 0; i < data.length; i++) {
    const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
    if (i % 2 === 1) {
      rowRange.setBackground('#f8f9fa');
    }
  }
  
  // Highlight empty cells in a subtle way
  const dataRange = sheet.getRange(2, 2, data.length, headers.length - 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#fce8e6')
    .setRanges([dataRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Add note to header
  sheet.getRange(1, 1).setNote('Product Key must match the Products tab.\nFolder IDs should be just the ID, not full URLs.');
}

// ============================================================================
// TEAM MEMBERS TAB
// ============================================================================

function createTeamMembersTab(ss) {
  let sheet = ss.getSheetByName('Team Members');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Team Members');
  
  // Headers
  const headers = ['Name', 'Email', 'ClickUp API Key', 'Is Admin', 'ClickUp Access', 'Andromeda Access', 'AutoBatch Sheet Tab', 'Active', 'Added Date'];
  
  // Data
  const data = [
    ['Kristijonas', 'kristijonas.t@commercecore.com', 'pk_100652070_MF1MRV4VX1ETPN5B192QG8BJS3RETZH6', true, true, true, 'Kristijonas', true, '2024-01-01'],
    ['Patricija', 'patricija.a@commercecore.com', 'pk_82749912_TV7GRLV4GNYW71DZVN0HRM0DQ595GSXA', false, true, false, 'Patricija', true, '2024-01-01'],
    ['Linas', 'linas.razgauskas@commercecore.com', 'pk_82794540_WT7KOPM9QFBBBVX3G8PQZO6SOACP7JEZ', false, true, true, 'Linas', true, '2024-01-01'],
    ['Inga', 'inga.k@commercecore.com', 'pk_82598770_WHK3KI4WSFU2P36CRPQQLMM4TQ4D9HQI', false, true, false, '', true, '2024-01-01'],
    ['Vismantas', 'vismantas.k@commercecore.com', 'pk_100558589_T3ZOO1FXW0C07Y2ZMCSLSQF22GG1YNKJ', false, true, false, '', true, '2024-01-01'],
    ['Paulius M.', 'paulius.ma@commercecore.com', 'pk_94576831_MLQFD0TQQE2UJKEJTZ363F824DJVUCPF', false, true, true, 'Paulius M', true, '2024-01-01'],
    ['Paulius S.', 'paulius.s@commercecore.com', '', false, false, false, '', true, '2024-01-01'],
    ['Karen', 'karen.s@commercecore.com', 'pk_100781578_NW2QPNDBMR4GW1P3R9KSCO8DOJRUOCHI', false, true, false, '', true, '2024-01-01'],
    ['Ieva', 'ieva.b@commercecore.com', '', false, false, false, '', true, '2024-01-01'],
    ['Gabija', 'gabija.p@commercecore.com', '', false, false, false, '', true, '2024-01-01'],
    ['Justas', 'justas.b@commercecore.com', '', false, false, false, '', true, '2024-01-01']
  ];
  
  // Write data
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  // Column widths
  sheet.setColumnWidth(1, 100); // Name
  sheet.setColumnWidth(2, 250); // Email
  sheet.setColumnWidth(3, 350); // API Key
  sheet.setColumnWidth(4, 80);  // Is Admin
  sheet.setColumnWidth(5, 110); // ClickUp Access
  sheet.setColumnWidth(6, 120); // Andromeda Access
  sheet.setColumnWidth(7, 130); // AutoBatch Sheet Tab
  sheet.setColumnWidth(8, 60);  // Active
  sheet.setColumnWidth(9, 100); // Added Date
  
  // Add checkboxes
  sheet.getRange(2, 4, data.length, 1).insertCheckboxes(); // Is Admin
  sheet.getRange(2, 5, data.length, 1).insertCheckboxes(); // ClickUp Access
  sheet.getRange(2, 6, data.length, 1).insertCheckboxes(); // Andromeda Access
  sheet.getRange(2, 8, data.length, 1).insertCheckboxes(); // Active
  
  // Center align checkbox columns
  sheet.getRange(2, 4, data.length, 3).setHorizontalAlignment('center');
  sheet.getRange(2, 8, data.length, 1).setHorizontalAlignment('center');
  
  // Alternate row colors
  for (let i = 0; i < data.length; i++) {
    const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
    if (i % 2 === 1) {
      rowRange.setBackground('#f8f9fa');
    }
  }
  
  // Highlight missing API keys
  const apiKeyRange = sheet.getRange(2, 3, data.length, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#fce8e6')
    .setRanges([apiKeyRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Add note
  sheet.getRange(1, 3).setNote('Get API key from: https://app.clickup.com/settings/apps\n\n⚠️ Keep this tab private - contains sensitive keys!');
  sheet.getRange(1, 7).setNote('The sheet tab name in Ad Naming Sheet where this user works');
  
  // Add filter
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
}

// ============================================================================
// CLICKUP CONFIG TAB
// ============================================================================

function createClickUpConfigTab(ss) {
  let sheet = ss.getSheetByName('ClickUp Config');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('ClickUp Config');
  
  // Headers
  const headers = ['Setting', 'Value', 'Description'];
  
  // Data
  const data = [
    ['TEAM_ID', '2428192', 'ClickUp Team/Workspace ID'],
    ['WORKSPACE_ID', '2428192', 'Same as Team ID'],
    ['SPACE_ID', '90151192866', 'Performance Creative space'],
    ['FB_ADS_FOLDER_ID', '90150821271', 'FB Ads folder containing product lists'],
    ['SPRINT_FOLDER_ID', '90154797247', 'PC Sprint 2025 folder'],
    ['FB_ADS_LISTS_FOLDER_ID', '90151192866', 'Folder containing A01, A02... lists']
  ];
  
  // Write data
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#fbbc04')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  // Column widths
  sheet.setColumnWidth(1, 200); // Setting
  sheet.setColumnWidth(2, 200); // Value
  sheet.setColumnWidth(3, 300); // Description
  
  // Style data
  sheet.getRange(2, 1, data.length, 1).setFontWeight('bold'); // Setting names bold
  sheet.getRange(2, 2, data.length, 1).setFontFamily('Courier New'); // Values in monospace
  sheet.getRange(2, 3, data.length, 1).setFontColor('#5f6368'); // Descriptions gray
  
  // Alternate row colors
  for (let i = 0; i < data.length; i++) {
    const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
    if (i % 2 === 1) {
      rowRange.setBackground('#fffde7');
    }
  }
  
  // Add border
  sheet.getRange(1, 1, data.length + 1, headers.length).setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
}

// ============================================================================
// VALID VALUES TAB
// ============================================================================

function createValidValuesTab(ss) {
  let sheet = ss.getSheetByName('Valid Values');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Valid Values');
  
  // Section 1: Platforms
  sheet.getRange(1, 1).setValue('Platforms').setFontWeight('bold').setBackground('#e8f0fe');
  const platforms = ['FB', 'IG', 'TT', 'YT', 'GG'];
  for (let i = 0; i < platforms.length; i++) {
    sheet.getRange(i + 2, 1).setValue(platforms[i]);
  }
  
  // Section 2: Locales
  sheet.getRange(1, 2).setValue('Locales').setFontWeight('bold').setBackground('#e8f0fe');
  const locales = ['US', 'UK', 'DE', 'CA', 'AU', 'LATAM', 'ES', 'FR', 'IT', 'IL', 'EU', 'MIX', 'EN', 'MX', 'BR'];
  for (let i = 0; i < locales.length; i++) {
    sheet.getRange(i + 2, 2).setValue(locales[i]);
  }
  
  // Section 3: Product Aliases
  sheet.getRange(1, 4).setValue('Alias').setFontWeight('bold').setBackground('#fce8e6');
  sheet.getRange(1, 5).setValue('Maps To').setFontWeight('bold').setBackground('#fce8e6');
  
  const aliases = [
    ['ArmouredCard', 'Guardality'],
    ['Wavemax', 'WaveMax'],
    ['Flixy', 'TellyStick'],
    ['TVStick', 'TellyStick'],
    ['Clairu-M2', 'Clairu_M2'],
    ['therawolfrelief', 'TheraWolfRelief'],
    ['therawolfneuro', 'TheraWolfNeuro']
  ];
  
  for (let i = 0; i < aliases.length; i++) {
    sheet.getRange(i + 2, 4).setValue(aliases[i][0]);
    sheet.getRange(i + 2, 5).setValue(aliases[i][1]);
  }
  
  // Section 4: Monitor Folders (NEW!)
  sheet.getRange(1, 7).setValue('Monitor Folders').setFontWeight('bold').setBackground('#e6f4ea');
  sheet.getRange(1, 8).setValue('Folder Product').setFontWeight('bold').setBackground('#e6f4ea');
  
  const monitorFolders = [
    ['1uZ8nkfizmfWqmg5wOCmcwgihGqCmeTbI', 'CamTrix'],
    ['1dlDonXG_fl0USaQ9XvXVlh3M8pDio74M', 'Clairu'],
    ['1gMgRFp2pFxacgf4lIk_gsPY66sPzkQ08', 'KatuChef'],
    ['1vvu3ZbPK1BwHpf537ReysJuTh1mFHYKZ', 'TellyStick'],
    ['1_bL9apX3nWqimjr038LrWeVlUDEpU6bf', 'OriBreeze'],
    ['1xLA4kbQokrmLpI9vhsomaP1qQ5Ij0Dbl', 'EpiCooler'],
    ['1taIm60nJ6N7DaC8WFp6b6uLpdPHjT5Fv', 'WaveMax'],
    ['1lhZ3VJkxf8xIm35t6BkfFGP6n3ku0DvS', 'Guardality'],
    ['15MHoweMB84JeSYzg9pRIvXbqgZUE6clU', 'WellHeater'],
    ['1v8HVgXaqgolX-ABwbO9WjOS7nYPxKeHz', 'Heatoor'],
    ['17rI7s65Zf4e4gNqn_FVGXZZ3ZYixPNJJ', 'TheraWolfNeuro'],
    ['1ZtDXwM-drOEI9T58xJGguQRCmln7Pr1H', 'TheraWolfRelief'],
    ['15q3nVlBsUZu8yIl8zXrdvb1unRwrlwFT', 'AliveBlue'],
    ['1mG1Mr0ejq7fLlDPFTWhXj_ZdDA0U-et1', 'CleanLix'],
    ['1OA9n0Yxk5TgGcsM_ZcnODyogANmUAXS4', 'WiggyDog'],
    ['1LDDGu1o6yRCQNiCf4rwiNNV6sRdaGRlD', 'Ozoori'],
    ['1cme5S5BrJdX1iQ789UGIIGqkegtrgehp', 'LumenLight']
  ];
  
  for (let i = 0; i < monitorFolders.length; i++) {
    sheet.getRange(i + 2, 7).setValue(monitorFolders[i][0]).setFontFamily('Courier New').setFontSize(9);
    sheet.getRange(i + 2, 8).setValue(monitorFolders[i][1]).setFontColor('#5f6368');
  }
  
  // Column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 30);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(8, 120);
}

// ============================================================================
// SYNC LOG TAB
// ============================================================================

function createSyncLogTab(ss) {
  let sheet = ss.getSheetByName('Sync Log');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Sync Log');
  
  // Headers
  const headers = ['Timestamp', 'System', 'Status', 'Items Loaded', 'Duration (ms)', 'Notes'];
  
  // Write headers only
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#9e9e9e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  // Column widths
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 180); // System
  sheet.setColumnWidth(3, 80);  // Status
  sheet.setColumnWidth(4, 100); // Items Loaded
  sheet.setColumnWidth(5, 100); // Duration
  sheet.setColumnWidth(6, 300); // Notes
  
  // Add example row (will be overwritten by actual syncs)
  const exampleData = [
    [new Date(), 'Example - Config Created', '✅ Success', 'N/A', 0, 'Initial setup complete. This row will be replaced by actual sync logs.']
  ];
  sheet.getRange(2, 1, 1, headers.length).setValues(exampleData);
  sheet.getRange(2, 1, 1, headers.length).setFontColor('#9e9e9e').setFontStyle('italic');
  
  // Add note
  sheet.getRange(1, 1).setNote('This tab auto-populates when automation systems read from config.\nMost recent entries appear at the top.\nOld entries (100+) are auto-deleted.');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function reformatAllTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Reformatting not yet implemented - use Create All Config Tabs instead', '⚙️ Info', 3);
}

function clearAndRebuild() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Clear & Rebuild',
    'This will DELETE all config tabs and recreate them from scratch.\n\nAny manual changes you made will be LOST.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    createAllConfigTabs();
  }
}

/**
 * Quick function to show the sheet ID (for copying to automation scripts)
 */
function showSheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().alert('📋 Config Sheet ID', ss.getId(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function createNewConfigTabs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.alert(
    '🚀 Create New Config Tabs',
    'This will create 4 new configuration tabs:\n\n' +
    '• Slack Config\n' +
    '• Product Spreadsheets\n' +
    '• Templates\n' +
    '• Footage Folders\n\n' +
    'Existing tabs with the same names will be replaced.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  ss.toast('Creating Slack Config tab...', '⚙️ Setup', -1);
  createSlackConfigTab(ss);
  
  ss.toast('Creating Product Spreadsheets tab...', '⚙️ Setup', -1);
  createProductSpreadsheetsTab(ss);
  
  ss.toast('Creating Templates tab...', '⚙️ Setup', -1);
  createTemplatesTab(ss);
  
  ss.toast('Creating Footage Folders tab...', '⚙️ Setup', -1);
  createFootageFoldersTab(ss);
  
  ss.toast('✅ All 4 new config tabs created!', '⚙️ Setup Complete', 5);
  
  ui.alert('✅ Setup Complete!', 
    'Created 4 new configuration tabs:\n\n' +
    '• Slack Config (4 settings)\n' +
    '• Product Spreadsheets (copy your mappings here)\n' +
    '• Templates (6 template IDs)\n' +
    '• Footage Folders (20 products)\n\n' +
    '⚠️ IMPORTANT:\n' +
    '1. Copy your Slack Bot Token to Slack Config\n' +
    '2. Copy your Product Spreadsheet IDs from Atria Settings\n' +
    '3. Verify Templates and Footage Folders are correct',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// SLACK CONFIG TAB
// ============================================================================

function createSlackConfigTab(ss) {
  let sheet = ss.getSheetByName('Slack Config');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Slack Config');
  
  const headers = ['Setting', 'Value', 'Description'];
  
  const data = [
    ['BOT_TOKEN', '', 'Slack Bot OAuth Token (xoxb-...)'],
    ['WEEKLY_CHANNEL_ID', '', 'Channel ID for weekly reports'],
    ['CREATIVE_CHANNEL_ID', '', 'Channel ID for creative tests'],
    ['DRIVE_NAMING_CHANNEL_ID', '', 'Channel ID for drive naming alerts']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 350);
  
  // Style data
  sheet.getRange(2, 1, data.length, 1).setFontWeight('bold');
  sheet.getRange(2, 2, data.length, 1).setFontFamily('Courier New');
  sheet.getRange(2, 3, data.length, 1).setFontColor('#5f6368');
  
  // Add security note
  sheet.getRange(1, 1).setNote('⚠️ SENSITIVE: This tab contains API tokens. Keep access restricted!');
  
  // Highlight empty values
  const valueRange = sheet.getRange(2, 2, data.length, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#fce8e6')
    .setRanges([valueRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

// ============================================================================
// PRODUCT SPREADSHEETS TAB
// ============================================================================

function createProductSpreadsheetsTab(ss) {
  let sheet = ss.getSheetByName('Product Spreadsheets');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Product Spreadsheets');
  
  const headers = ['Product Key', 'Spreadsheet ID', 'Description'];
  
  // Example data - user will need to fill in from their Atria Settings
  const data = [
    ['AliveBlue', '', 'AliveBlue performance tracking'],
    ['CamTrix', '', 'CamTrix performance tracking'],
    ['CamTrix (Main)', '', 'CamTrix main account tracking'],
    ['Clairu_M2', '', 'Clairu M2 performance tracking'],
    ['EpiCooler', '', 'EpiCooler performance tracking'],
    ['Guardality', '', 'Guardality performance tracking'],
    ['Heatoor', '', 'Heatoor performance tracking'],
    ['OmniHear', '', 'OmniHear performance tracking'],
    ['OriBreeze', '', 'OriBreeze performance tracking'],
    ['Ozoori', '', 'Ozoori performance tracking'],
    ['TellyStick', '', 'TellyStick performance tracking'],
    ['TheraWolfNeuro', '', 'TheraWolfNeuro performance tracking'],
    ['TheraWolfRelief', '', 'TheraWolfRelief performance tracking'],
    ['WaveMax', '', 'WaveMax performance tracking'],
    ['WellHeater', '', 'WellHeater performance tracking'],
    ['WiggyDog', '', 'WiggyDog performance tracking']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#0f9d58')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 280);
  
  // Style data
  sheet.getRange(2, 1, data.length, 1).setFontWeight('bold');
  sheet.getRange(2, 2, data.length, 1).setFontFamily('Courier New');
  sheet.getRange(2, 3, data.length, 1).setFontColor('#5f6368');
  
  // Alternate row colors
  for (let i = 0; i < data.length; i++) {
    const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
    if (i % 2 === 1) {
      rowRange.setBackground('#e6f4ea');
    }
  }
  
  // Add note
  sheet.getRange(1, 1).setNote('Product Key must match labels used in Atria Reports.\nSupports labels like "CamTrix (Main)" for multi-account products.');
  sheet.getRange(1, 2).setNote('Just the Spreadsheet ID (44-character string), not the full URL.\nGet it from: https://docs.google.com/spreadsheets/d/[THIS_PART]/edit');
  
  // Add filter
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
}

// ============================================================================
// TEMPLATES TAB
// ============================================================================

function createTemplatesTab(ss) {
  let sheet = ss.getSheetByName('Templates');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Templates');
  
  const headers = ['Template Key', 'Document ID', 'Description'];
  
  const data = [
    ['IMG', '1zmyCJLrNLxOgsRperFRH319Ctaj8LPne3xwuYgPoI1w', 'Image ad brief template'],
    ['VID', '1RKXPnmirIelXitg8HVRuD8r3afDloLRqqjoD6y6AoXc', 'Video ad brief template'],
    ['ITR', '1LKlyX1j0sNVJYYH3WdGTip4ZdyTJ92Ob5tnwzfZlth8', 'Iteration brief template'],
    ['LOC', '15-YSsBz_8gDY9s2rTp02ry23ZzraYdJcGdMsy7oZY-4', 'Localization brief template'],
    ['YTTOFB', '1nA8iyXEjcG38O46EF87zSW-fyZWrQNtw48ec-yK3VMQ', 'YouTube to FB brief template'],
    ['ANDROMEDA_VID', '1LPY644EsF-bLjUwrt4oy_pZXbR6P-04CAy-yh-bAWnk', 'Andromeda video brief template']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#db4437')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 280);
  
  // Style data
  sheet.getRange(2, 1, data.length, 1).setFontWeight('bold');
  sheet.getRange(2, 2, data.length, 1).setFontFamily('Courier New').setFontSize(9);
  sheet.getRange(2, 3, data.length, 1).setFontColor('#5f6368');
  
  // Alternate row colors
  for (let i = 0; i < data.length; i++) {
    const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
    if (i % 2 === 1) {
      rowRange.setBackground('#fce8e6');
    }
  }
  
  // Add note
  sheet.getRange(1, 1).setNote('Template Key used in code: IMG_TEMPLATE_ID, VID_TEMPLATE_ID, etc.');
  sheet.getRange(1, 2).setNote('Google Doc ID (44-character string).\nAll team members need at least View access to these templates.');
}

// ============================================================================
// FOOTAGE FOLDERS TAB
// ============================================================================

function createFootageFoldersTab(ss) {
  let sheet = ss.getSheetByName('Footage Folders');
  if (sheet) ss.deleteSheet(sheet);
  
  sheet = ss.insertSheet('Footage Folders');
  
  const headers = ['Product Key', 'Footage Folder URL'];
  
  const data = [
    ['AliveBlue', 'https://drive.google.com/drive/folders/11VZ3TIR9u_R3wCcJXaYLTYq0MAv08zH6'],
    ['AuraNaturals', 'https://drive.google.com/drive/folders/1B8YmG7BWb7oUPVAasD4mr63nCetbLdlt'],
    ['CamTrix', 'https://drive.google.com/drive/folders/1jCrAaVyvV1zURcFExthu8jccoDiUs0Mk'],
    ['Clairu', 'https://drive.google.com/drive/folders/1glDIEoPCGEF13UR1tjBj4hPE7ZzFzhnk'],
    ['Clairu_M2', 'https://drive.google.com/drive/u/0/folders/1glDIEoPCGEF13UR1tjBj4hPE7ZzFzhnk'],
    ['CleanlixMold', 'https://drive.google.com/drive/folders/1B5rsMz-wdmtOFoIQ3JwsQs93mKOviBPS'],
    ['CleanlixPowder', 'https://drive.google.com/drive/folders/1g6mE8O_3SEQIqqM64bc74a2FFryPzT2_'],
    ['EpiCooler', 'https://drive.google.com/drive/folders/15qj9M0m11ShfnvgKyaoUtC4upC6Vtimd'],
    ['Guardality', 'https://drive.google.com/drive/folders/1eyEqM2iSOJMnNygTaqjIIt9U8OYpTOzv'],
    ['Heatoor', 'https://drive.google.com/drive/folders/12RpUYCeRXGN3DKAqfMw-KyItm5U2BEBq'],
    ['KatuChef', 'https://drive.google.com/drive/folders/1gTcSZlOAUC84OEpitVyOIE49bCu7D-RV'],
    ['OmniHear', 'https://drive.google.com/drive/folders/1mT7jpjXaAhjCtKOE0wc0h4RziILuM5sX'],
    ['OriBreeze', 'https://drive.google.com/drive/folders/1CYp1kWaHpCMVG19xXiiSFwgpdQC_RBRG'],
    ['Ozoori', 'https://drive.google.com/drive/folders/1pxF-P2YMeJsqOoJCe0UhexqhgCpq6TDF'],
    ['TellyStick', 'https://drive.google.com/drive/folders/1U0YKbwrwdZeTCNmfrWooKPQcomfSkOZV'],
    ['TheraWolfNeuro', 'https://drive.google.com/drive/folders/1kSYm3aIi_4K55JAVGKCNCoBjddmHvKYq'],
    ['TheraWolfRelief', 'https://drive.google.com/drive/folders/1EK3o966nv4WTsH8gQahuzxCZkgWotuHw'],
    ['WaveMax', 'https://drive.google.com/drive/folders/1wW14AT70gm2-EVK7XH0153KMxSDFlFeK'],
    ['WellHeater', 'https://drive.google.com/drive/folders/1jn6GeQgHNqOG3Un32wYPhqFDfnlsR8h6'],
    ['WiggyDog', 'https://drive.google.com/drive/folders/1AET_w4i4OPxpsHjGFn7qDqA0oQuVXmdd'],
    ['Blitron', ''],
    ['Cleanlix', ''],
    ['DermaSonic', ''],
    ['Epibella', ''],
    ['Flixy', ''],
    ['HammerDex', ''],
    ['LoveAndShine', ''],
    ['LumenLight', ''],
    ['LoweSkin', ''],
    ['NuraDerma', ''],
    ['NuraFix', ''],
    ['Phoxfer', ''],
    ['PlanetWash', ''],
    ['PrimaFocus', ''],
    ['Repellio', ''],
    ['SleepZee', ''],
    ['TVStick', ''],
    ['VacuumSealer', ''],
    ['Zapfie', ''],
    ['BioVitals', ''],
    ['BRYT', ''],
    ['EmaWell', '']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#f4b400')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 500);
  
  // Style data
  sheet.getRange(2, 1, data.length, 1).setFontWeight('bold');
  sheet.getRange(2, 2, data.length, 1).setFontFamily('Courier New').setFontSize(9);
  
  // Highlight empty URLs
  const urlRange = sheet.getRange(2, 2, data.length, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#fff3e0')
    .setRanges([urlRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  // Add note
  sheet.getRange(1, 1).setNote('Product Key must match the Products tab exactly.');
  sheet.getRange(1, 2).setNote('Full Google Drive folder URL for product footage.\nLeave empty if no footage folder exists.');
  
  // Add filter
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();
}