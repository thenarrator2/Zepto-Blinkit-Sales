/**
 * Sales Data Dashboard - Template Updater
 * 
 * This script updates your existing template sheet with new data
 * Just paste your data and click the update button!
 */

// Configuration - Modify these to match your sheet structure
const CONFIG = {
  // Data input sheets (where you paste new data)
  zeptoDataSheet: "Zepto Data Input",
  blinkitDataSheet: "Blinkit Data Input",
  
  // Your template sheet (the one that gets updated)
  templateSheet: "Template for Blinkit",
  
  // Template layout configuration
  template: {
    skuSummaryStartRow: 3,      // Where SKU summary starts in your template
    skuSummaryStartCol: 2,      // Column B (SKU names)
    
    citySummaryStartRow: 25,    // Where City summary starts in your template  
    citySummaryStartCol: 2,     // Column B (City names)
    
    weekHeaderRow: 2,           // Row where week headers go
    weekStartCol: 3,            // Column C (first week column)
    
    maxRows: 50,                // Maximum rows to clear/update
    maxCols: 20                 // Maximum columns to clear/update
  },
  
  // Column mappings for your data
  zepto: {
    skuColumn: "SKU Name",
    cityColumn: "City", 
    dateColumn: "Sales Date",
    qtyColumn: "Quantity"
  },
  blinkit: {
    skuColumn: "item_name",
    cityColumn: "city_name",
    dateColumn: "date", 
    qtyColumn: "qty"
  }
};

/**
 * Create the update button when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Sales Dashboard')
    .addItem('🔄 Update Template with New Data', 'updateTemplate')
    .addSeparator()
    .addItem('📋 Setup Instructions', 'showInstructions')
    .addItem('⚙️ Configuration', 'showConfiguration')
    .addToUi();
}

/**
 * Main function to update the template with new data
 */
function updateTemplate() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if template sheet exists
    const templateSheet = ss.getSheetByName(CONFIG.templateSheet);
    if (!templateSheet) {
      ui.alert('Error', `Template sheet "${CONFIG.templateSheet}" not found!\n\nPlease create this sheet or update the CONFIG.templateSheet name.`, ui.ButtonSet.OK);
      return;
    }
    
    // Ask which data to process
    const response = ui.alert(
      'Update Template',
      'Which data would you like to process?\n\n' +
      'YES = Zepto Data\n' +
      'NO = Blinkit Data\n' +
      'CANCEL = Both',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    let processed = false;
    
    if (response === ui.Button.YES || response === ui.Button.CANCEL) {
      // Process Zepto data
      if (processAndUpdateTemplate('zepto')) {
        processed = true;
      }
    }
    
    if (response === ui.Button.NO || response === ui.Button.CANCEL) {
      // Process Blinkit data
      if (processAndUpdateTemplate('blinkit')) {
        processed = true;
      }
    }
    
    if (processed) {
      ui.alert('Success!', 'Template updated successfully!\n\nCheck your "' + CONFIG.templateSheet + '" sheet for the updated summaries.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Process data and update template
 */
function processAndUpdateTemplate(platform) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(CONFIG.templateSheet);
  
  try {
    // Get the data input sheet
    const dataSheetName = platform === 'zepto' ? CONFIG.zeptoDataSheet : CONFIG.blinkitDataSheet;
    const dataSheet = ss.getSheetByName(dataSheetName);
    
    if (!dataSheet) {
      SpreadsheetApp.getUi().alert('Warning', `Data sheet "${dataSheetName}" not found.\n\nPlease create this sheet and paste your ${platform} data there.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
    
    // Get and process the data
    const rawData = getSheetData(dataSheet, CONFIG[platform]);
    
    if (rawData.length === 0) {
      SpreadsheetApp.getUi().alert('Warning', `No data found in "${dataSheetName}" sheet.\n\nPlease paste your ${platform} data and try again.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
    
    // Process data into summaries
    const summaries = createSummaries(rawData);
    
    // Update the template sheet
    updateTemplateSheet(templateSheet, summaries);
    
    return true;
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error processing ${platform} data: ` + error.toString());
    return false;
  }
}

/**
 * Get data from input sheet
 */
function getSheetData(sheet, config) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return []; // No data rows
  
  const headers = data[0];
  
  // Find column indices
  const skuIndex = headers.indexOf(config.skuColumn);
  const cityIndex = headers.indexOf(config.cityColumn);
  const dateIndex = headers.indexOf(config.dateColumn);
  const qtyIndex = headers.indexOf(config.qtyColumn);
  
  if (skuIndex === -1 || cityIndex === -1 || dateIndex === -1 || qtyIndex === -1) {
    throw new Error(`Column not found. Expected: ${config.skuColumn}, ${config.cityColumn}, ${config.dateColumn}, ${config.qtyColumn}`);
  }
  
  // Extract and clean data
  const cleanData = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[skuIndex] && row[cityIndex] && row[dateIndex] && row[qtyIndex]) {
      cleanData.push({
        sku: row[skuIndex].toString(),
        city: row[cityIndex].toString(),
        date: new Date(row[dateIndex]),
        qty: Number(row[qtyIndex]) || 0
      });
    }
  }
  
  return cleanData;
}

/**
 * Create summaries from raw data
 */
function createSummaries(data) {
  // Add week ranges to data
  data.forEach(row => {
    row.weekRange = getWeekRange(row.date);
  });
  
  // Get unique weeks and sort them
  const weeks = [...new Set(data.map(row => row.weekRange))].sort();
  
  // Create SKU summary
  const skuData = {};
  data.forEach(row => {
    if (!skuData[row.sku]) {
      skuData[row.sku] = {};
      weeks.forEach(week => skuData[row.sku][week] = 0);
    }
    skuData[row.sku][row.weekRange] += row.qty;
  });
  
  // Create City summary
  const cityData = {};
  data.forEach(row => {
    if (!cityData[row.city]) {
      cityData[row.city] = {};
      weeks.forEach(week => cityData[row.city][week] = 0);
    }
    cityData[row.city][row.weekRange] += row.qty;
  });
  
  return {
    weeks: weeks,
    skuData: skuData,
    cityData: cityData
  };
}

/**
 * Update the template sheet with new data
 */
function updateTemplateSheet(sheet, summaries) {
  const { weeks, skuData, cityData } = summaries;
  const config = CONFIG.template;
  
  // Clear existing data areas
  clearTemplateAreas(sheet);
  
  // Update week headers
  if (weeks.length > 0) {
    const weekHeaders = weeks.map(week => [week]);
    sheet.getRange(config.weekHeaderRow, config.weekStartCol, 1, weeks.length).setValues([weeks]);
  }
  
  // Update SKU summary
  updateSummarySection(sheet, skuData, weeks, config.skuSummaryStartRow, config.skuSummaryStartCol, config.weekStartCol);
  
  // Update City summary  
  updateSummarySection(sheet, cityData, weeks, config.citySummaryStartRow, config.citySummaryStartCol, config.weekStartCol);
  
  // Format the updated areas
  formatTemplate(sheet, weeks.length);
}

/**
 * Clear existing data in template areas
 */
function clearTemplateAreas(sheet) {
  const config = CONFIG.template;
  
  // Clear SKU summary area
  sheet.getRange(
    config.skuSummaryStartRow, 
    config.skuSummaryStartCol, 
    config.maxRows, 
    config.maxCols
  ).clearContent();
  
  // Clear City summary area
  sheet.getRange(
    config.citySummaryStartRow, 
    config.citySummaryStartCol, 
    config.maxRows, 
    config.maxCols
  ).clearContent();
  
  // Clear week headers
  sheet.getRange(
    config.weekHeaderRow, 
    config.weekStartCol, 
    1, 
    config.maxCols
  ).clearContent();
}

/**
 * Update a summary section (SKU or City)
 */
function updateSummarySection(sheet, data, weeks, startRow, nameCol, dataStartCol) {
  const items = Object.keys(data).sort();
  
  items.forEach((item, index) => {
    const row = startRow + index;
    
    // Set item name
    sheet.getRange(row, nameCol).setValue(item);
    
    // Set weekly quantities
    weeks.forEach((week, weekIndex) => {
      const col = dataStartCol + weekIndex;
      const value = data[item][week] || 0;
      sheet.getRange(row, col).setValue(value);
    });
  });
}

/**
 * Format the template after updating
 */
function formatTemplate(sheet, numWeeks) {
  const config = CONFIG.template;
  
  // Format week headers
  if (numWeeks > 0) {
    sheet.getRange(config.weekHeaderRow, config.weekStartCol, 1, numWeeks)
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  // Format summary section headers (assuming they exist)
  sheet.getRange(config.skuSummaryStartRow - 1, config.skuSummaryStartCol, 1, numWeeks + 1)
    .setFontWeight('bold');
    
  sheet.getRange(config.citySummaryStartRow - 1, config.citySummaryStartCol, 1, numWeeks + 1)
    .setFontWeight('bold');
}

/**
 * Get week range string (Monday to Sunday)
 */
function getWeekRange(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(d.setDate(diff));
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  const formatDate = (date) => {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${date.getDate()}-${months[date.getMonth()]}`;
  };
  
  return `${formatDate(monday)} - ${formatDate(sunday)}`;
}

/**
 * Show setup instructions
 */
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Instructions',
    'How to use this Template Updater:\n\n' +
    '1. Create input sheets: "' + CONFIG.zeptoDataSheet + '" and "' + CONFIG.blinkitDataSheet + '"\n' +
    '2. Make sure your template sheet "' + CONFIG.templateSheet + '" exists\n' +
    '3. Paste your Excel data into the input sheets\n' +
    '4. Click "📊 Sales Dashboard" → "🔄 Update Template with New Data"\n' +
    '5. Your template will be updated automatically!\n\n' +
    'Expected columns:\n' +
    '• Zepto: ' + CONFIG.zepto.skuColumn + ', ' + CONFIG.zepto.cityColumn + ', ' + CONFIG.zepto.dateColumn + ', ' + CONFIG.zepto.qtyColumn + '\n' +
    '• Blinkit: ' + CONFIG.blinkit.skuColumn + ', ' + CONFIG.blinkit.cityColumn + ', ' + CONFIG.blinkit.dateColumn + ', ' + CONFIG.blinkit.qtyColumn,
    ui.ButtonSet.OK
  );
}

/**
 * Show configuration options
 */
function showConfiguration() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Configuration',
    'Current Configuration:\n\n' +
    'Input Sheets:\n' +
    '• Zepto Data: "' + CONFIG.zeptoDataSheet + '"\n' +
    '• Blinkit Data: "' + CONFIG.blinkitDataSheet + '"\n\n' +
    'Template Sheet: "' + CONFIG.templateSheet + '"\n\n' +
    'Template Layout:\n' +
    '• SKU Summary starts at Row ' + CONFIG.template.skuSummaryStartRow + ', Column ' + CONFIG.template.skuSummaryStartCol + '\n' +
    '• City Summary starts at Row ' + CONFIG.template.citySummaryStartRow + ', Column ' + CONFIG.template.citySummaryStartCol + '\n' +
    '• Week headers at Row ' + CONFIG.template.weekHeaderRow + ', starting Column ' + CONFIG.template.weekStartCol + '\n\n' +
    'To modify: Edit the CONFIG object in Apps Script editor.',
    ui.ButtonSet.OK
  );
} 