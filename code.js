/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const SHEET_NAME = 'Page Speed'; // The name of your sheet
const API_KEY_PROPERTY_NAME = 'PAGESPEED_API_KEY'; // The name of the script property holding your API key
const BATCH_SIZE = 5; // The number of URLs to process in each batch. Reduced size for more data points per URL.
const NOTIFICATION_EMAIL = 'abhishek.yadav@infidigit.com'; // Add your email to get a notification on completion


/**
 * Creates the custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Automate Page Speed')
    .addItem('Start Full Update', 'startUpdateProcess')
    .addItem('Cancel Running Update', 'cancelUpdateProcess')
    .addToUi();
}

/**
 * Kicks off the entire multi-part update process.
 * This is the function the user runs from the menu.
 */
function startUpdateProcess() {
  const ui = SpreadsheetApp.getUi();

  // 1. Clean up any old triggers and properties from previous runs.
  cancelUpdateProcess();

  // 2. Initialize the state in Script Properties for all metrics.
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('currentIndex', '0');

  const metrics = ['performance', 'fcp', 'si', 'tbt', 'lcp', 'cls', 'inp'];
  const strategies = ['mobile', 'desktop'];

  for (const strategy of strategies) {
    for (const metric of metrics) {
      scriptProperties.setProperty(`${strategy}_${metric}_scores`, JSON.stringify([]));
    }
  }

  // 3. Create a trigger to run the first batch immediately.
  ScriptApp.newTrigger('processBatch')
    .timeBased()
    .after(1000) // Run after 1 second
    .create();

  ui.alert('✅ Update Started', 'The PageSpeed analysis has begun. It will fetch detailed metrics and may take 20-30 minutes. You will receive an email when it is complete. You can close this sheet.', ui.ButtonSet.OK);
}

/**
 * Processes one batch of URLs. This function is called by a trigger.
 */
function processBatch() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty(API_KEY_PROPERTY_NAME);

  // --- Load State ---
  let currentIndex = parseInt(scriptProperties.getProperty('currentIndex'), 10);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastColumn = sheet.getLastColumn();
  const allUrls = sheet.getRange(3, 3, 1, lastColumn - 2).getValues()[0];

  const startTime = new Date();

  // --- Process Batch ---
  for (let i = 0; i < BATCH_SIZE && currentIndex < allUrls.length; i++) {
    const url = allUrls[currentIndex];
    console.log(`Processing URL ${currentIndex + 1}/${allUrls.length}: ${url}`);
    
    // Fetch data for both desktop and mobile
    const mobileData = fetchPageSpeedData(url, 'mobile', apiKey);
    const desktopData = fetchPageSpeedData(url, 'desktop', apiKey);

    // Save the retrieved data into script properties
    saveDataToProperties('mobile', mobileData, scriptProperties);
    saveDataToProperties('desktop', desktopData, scriptProperties);

    currentIndex++;

    // Safety check to ensure we don't exceed the time limit within a batch
    if (new Date() - startTime > 270000) { // 4.5 minutes
      break;
    }
  }

  // --- Save State ---
  scriptProperties.setProperty('currentIndex', currentIndex);

  // --- Decide Next Step ---
  if (currentIndex < allUrls.length) {
    // If there are more URLs, set a trigger for the next batch.
    ScriptApp.newTrigger('processBatch')
      .timeBased()
      .after(60 * 1000) // Run next batch after 1 minute to respect API limits
      .create();
  } else {
    // If all URLs are processed, finalize the sheet.
    finalizeUpdate();
  }
}

/**
 * Helper to save the fetched data object into respective script properties.
 */
function saveDataToProperties(strategy, data, scriptProperties) {
  for (const key in data) {
    const propName = `${strategy}_${key}_scores`;
    const scoresArray = JSON.parse(scriptProperties.getProperty(propName));
    scoresArray.push(data[key]);
    scriptProperties.setProperty(propName, JSON.stringify(scoresArray));
  }
}


/**
 * Writes the final data to the sheet and cleans up.
 */
function finalizeUpdate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const newRowStart = lastRow + 2; // Add a blank row for separation

  const scriptProperties = PropertiesService.getScriptProperties();
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // --- Prepare Data for Writing ---
  const mobilePerf = JSON.parse(scriptProperties.getProperty('mobile_performance_scores'));
  const mobileFcp = JSON.parse(scriptProperties.getProperty('mobile_fcp_scores'));
  const mobileSi = JSON.parse(scriptProperties.getProperty('mobile_si_scores'));
  const mobileTbt = JSON.parse(scriptProperties.getProperty('mobile_tbt_scores'));
  const mobileLcp = JSON.parse(scriptProperties.getProperty('mobile_lcp_scores'));
  const mobileCls = JSON.parse(scriptProperties.getProperty('mobile_cls_scores'));
  const mobileInp = JSON.parse(scriptProperties.getProperty('mobile_inp_scores'));

  const desktopPerf = JSON.parse(scriptProperties.getProperty('desktop_performance_scores'));
  const desktopFcp = JSON.parse(scriptProperties.getProperty('desktop_fcp_scores'));
  const desktopSi = JSON.parse(scriptProperties.getProperty('desktop_si_scores'));
  const desktopTbt = JSON.parse(scriptProperties.getProperty('desktop_tbt_scores'));
  const desktopLcp = JSON.parse(scriptProperties.getProperty('desktop_lcp_scores'));
  const desktopCls = JSON.parse(scriptProperties.getProperty('desktop_cls_scores'));
  const desktopInp = JSON.parse(scriptProperties.getProperty('desktop_inp_scores'));

  const dataToWrite = [
    [currentDate, 'Mobile Performance', ...mobilePerf],
    ['', 'First Contentful Paint', ...mobileFcp],
    ['', 'Speed Index', ...mobileSi],
    ['', 'Total Blocking Time', ...mobileTbt],
    ['', 'Largest Contentful Paint', ...mobileLcp],
    ['', 'Cumulative Layout Shift', ...mobileCls],
    ['', 'Interaction to Next Paint', ...mobileCls],
    ['', 'Desktop Performance', ...desktopPerf],
    ['', 'First Contentful Paint', ...desktopFcp],
    ['', 'Speed Index', ...desktopSi],
    ['', 'Total Blocking Time', ...desktopTbt],
    ['', 'Largest Contentful Paint', ...desktopLcp],
    ['', 'Cumulative Layout Shift', ...desktopCls],
    ['', 'Interaction to Next Paint', ...mobileCls],
  ];

  // --- Write and Format ---
  sheet.getRange(newRowStart, 1, 14, 2 + mobilePerf.length).setValues(dataToWrite);
  sheet.getRange(newRowStart, 1, 1, 2).setFontWeight('bold');
  sheet.getRange(newRowStart + 6, 2, 1, 1).setFontWeight('bold');

  // Notify user via email
  if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL !== 'your.email@example.com') {
    MailApp.sendEmail(NOTIFICATION_EMAIL, 'Google Sheet PageSpeed Update Complete', 'The automated competitor PageSpeed analysis (True Sliver) has finished successfully.');
  }

  // Clean up triggers and properties
  cancelUpdateProcess();
}

/**
 * Fetches PageSpeed data and returns an object with all required metrics.
 */
function fetchPageSpeedData(url, strategy, apiKey) {
  const emptyResult = { performance: '-', fcp: '-', si: '-', tbt: '-', lcp: '-', cls: '-',inp: '_' };
  
  if (!url || !url.toString().trim().startsWith('http')) {
      return emptyResult;
  }

  const apiUrl = `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encodeURIComponent(url)}&strategy=${strategy}&key=${apiKey}&category=PERFORMANCE`;
  const options = { 'muteHttpExceptions': true, 'validateHttpsCertificates': false };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.error) {
        console.error(`API Error for ${url} (${strategy}): ${json.error.message}`);
        return emptyResult;
      }
      
      const lighthouse = json.lighthouseResult;
      const audits = lighthouse?.audits;
      if (!audits) return emptyResult;

      return {
        performance: Math.round((lighthouse.categories?.performance?.score || 0) * 100),
        fcp: audits['first-contentful-paint']?.displayValue || '-',
        si: audits['speed-index']?.displayValue || '-',
        tbt: audits['total-blocking-time']?.displayValue || '-',
        lcp: audits['largest-contentful-paint']?.displayValue || '-',
        cls: audits['cumulative-layout-shift']?.displayValue || '-',
        inp: audits['Interaction-to-Next-Paint']?.displayValue || '-'
      };
    } else {
       console.error(`HTTP Error for ${url} (${strategy}): Response code ${responseCode}`);
       return emptyResult;
    }
  } catch (e) {
    console.error(`General Error for ${url} (${strategy}): ${e.message}`);
    return emptyResult;
  }
}


/**
 * Deletes all script triggers and clears properties to stop any running process.
 */
function cancelUpdateProcess() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('currentIndex');

  const metrics = ['performance', 'fcp', 'si', 'tbt', 'lcp', 'cls', 'inp'];
  const strategies = ['mobile', 'desktop'];
  for (const strategy of strategies) {
    for (const metric of metrics) {
      scriptProperties.deleteProperty(`${strategy}_${metric}_scores`);
    }
  }
}