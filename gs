// Global Spreadsheet ID Configuration to prevent context loss errors
const SPREADSHEET_ID = "1EikuuKnQJeCLLsy8-TCxokIC_TZ2EzaKBWFluXtaBQ4";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GE Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper function to reduce repetitive spreadsheet fetching code
 * @param {string} sheetName - Name of the tab to fetch
 * @param {string} keyName - Object property key for column 1
 * @param {string} valueName - Object property key for column 2
 */
function fetchTwoColumnData(sheetName, keyName, valueName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName); 
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      [keyName]: row[0] ? row[0].toString().trim() : '',
      [valueName]: row[1] ? row[1].toString().trim() : ''
    })).filter(item => item[keyName] !== '' && item[valueName] !== '');
  } catch(e) {
    console.error(`Error fetching data from ${sheetName}: `, e);
    return [];
  }
}

// Rewritten presets using the helper function
function getSpreadsheetPresets() { return fetchTwoColumnData('Templates', 'name', 'text'); }
function getWorkNotesPresets()   { return fetchTwoColumnData('Categories', 'name', 'text'); }
function getThreadRequestsPresets(){ return fetchTwoColumnData('ThreadRequests', 'name', 'text'); }
function getPhonebookData()       { return fetchTwoColumnData('Phonebook', 'name', 'phone'); }

// Fetches data for the GE Updates Feed sidebar panel from "Updates1" tab
function getSharedTemplates() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Updates1'); 
    if (!sheet) return ["Error: 'Updates1' sheet tab missing."];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 1) return ["No updates found."];
    
    return data
      .map(row => row[0] ? row[0].toString().trim() : '')
      .filter(txt => txt !== '');
  } catch(e) {
    return ["Error synchronizing stream: " + e.toString()];
  }
}

// --- Properties Service (User-Specific Data) ---

/**
 * Helper function to save user properties with quota error handling
 */
function saveUserProperty(key, value) {
  try {
    PropertiesService.getUserProperties().setProperty(key, value);
    return { success: true };
  } catch (e) {
    console.error(`Error saving property ${key}: `, e);
    return { success: false, error: `Failed to save. You may have reached the storage limit. Error: ${e.message}` };
  }
}

function getUserNotes() { return PropertiesService.getUserProperties().getProperty('GE_SAVED_NOTES') || '[]'; }
function saveUserNotes(jsonString) { return saveUserProperty('GE_SAVED_NOTES', jsonString); }

function getQuickResponses() { return PropertiesService.getUserProperties().getProperty('GE_QUICK_RESPONSES') || '[]'; }
function saveQuickResponses(jsonString) { return saveUserProperty('GE_QUICK_RESPONSES', jsonString); }

function getCustomFolders() { return PropertiesService.getUserProperties().getProperty('GE_CUSTOM_FOLDERS') || '[]'; }
function saveCustomFolders(jsonString) { return saveUserProperty('GE_CUSTOM_FOLDERS', jsonString); }

function loadVaultData() { return PropertiesService.getUserProperties().getProperty('GE_SECURE_VAULT') || ''; }
function saveVaultData(scrambledPayload) { return saveUserProperty('GE_SECURE_VAULT', scrambledPayload); }

function clearAllUserData() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('GE_SAVED_NOTES');
  props.deleteProperty('GE_QUICK_RESPONSES');
  props.deleteProperty('GE_CUSTOM_FOLDERS');
  props.deleteProperty('GE_SECURE_VAULT');
}

/**
 * Saves call trending metadata to the "Call Trending" sheet tab safely using LockService
 */
function logCallTrendingData(trendingData) {
  const lock = LockService.getScriptLock();
  
  try {
    // Wait for up to 3 seconds for other script executions to finish modifying the sheet
    lock.waitLock(3000);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Call Trending");
    
    if (!sheet) {
      sheet = ss.insertSheet("Call Trending");
      sheet.appendRow(["Timestamp", "Call Type", "Trend Reason", "Additional Details / Notes"]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#26262b").setFontColor("#f5f5f7");
    }
    
    sheet.appendRow([
      new Date(),
      trendingData.callType,
      trendingData.reason,
      trendingData.notes
    ]);
    
    return { success: true };
  } catch (error) {
    console.error("Error logging call trending data: ", error);
    return { success: false, error: error.toString() };
  } finally {
    // Always release the lock so other executions can proceed, even if an error occurred
    lock.releaseLock();
  }
}
