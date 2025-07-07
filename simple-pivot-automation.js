/**
 * Simple Pivot Table Automation
 * 
 * This script automatically creates pivot tables from your raw data
 * and updates your template - no manual pivot table creation needed!
 */

// Configuration
const CONFIG = {
  // Raw data sheets (where you paste your Excel data)
  zeptoRawData: "Zepto Raw Data",
  blinkitRawData: "Blinkit Raw Data",
  
  // Your template sheet
  templateSheet: "Template for Blinkit",
  
  // Template positions (adjust to match your template)
  skuSectionStartRow: 3,
  citySectionStartRow: 25,
  weekHeaderRow: 2,
  firstDataCol: 2,      // Column B for names
  firstWeekCol: 3,      // Column C for first week
  
  // Column names in your raw data
  zepto: {
    sku: "SKU Name",
    city: "City",
    date: "Sales Date", 
    qty: "Quantity"
  },
  blinkit: {
    sku: "item_name",
    city: "city_name",
    date: "date",
    qty: "qty"
  }
};

/**
 * Create menu when sheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Auto Pivot')
    .addItem('📊 Create Pivot Tables & Update Template', 'createPivotsAndUpdate')
    .addSeparator()
    .addItem('📋 Instructions', 'showInstructions')
    .addToUi();
}

/**
 * Main function - creates pivot tables and updates template
 */
function createPivotsAndUpdate() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Ask which data to process
    const response = ui.alert(
      'Auto Pivot Tables',
      'Which data would you like to process?\n\n' +
      'YES = Zepto Data\n' +
      'NO = Blinkit Data\n' +
      'CANCEL = Both',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    let processed = false;
    
    if (response === ui.Button.YES || response === ui.Button.CANCEL) {
      if (processData('zepto')) processed = true;
    }
    
    if (response === ui.Button.NO || response === ui.Button.CANCEL) {
      if (processData('blinkit')) processed = true;
    }
    
    if (processed) {
      ui.alert('✅ Success!', 'Pivot tables created and template updated!\n\nCheck your "' + CONFIG.templateSheet + '" sheet.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Error', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Process data for a specific platform
 */
function processData(platform) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Get the raw data sheet
  const rawDataSheetName = platform === 'zepto' ? CONFIG.zeptoRawData : CONFIG.blinkitRawData;
  let rawDataSheet = ss.getSheetByName(rawDataSheetName);
  
  // If exact name not found, look for similar sheets
  if (!rawDataSheet) {
    const allSheets = ss.getSheets();
    const platformName = platform === 'zepto' ? 'zepto' : 'blinkit';
    const similarSheets = allSheets.filter(sheet => 
      sheet.getName().toLowerCase().includes(platformName) || 
      sheet.getName().toLowerCase().includes(platformName.substring(0, 4))
    );
    
    if (similarSheets.length === 0) {
      ui.alert('⚠️ Warning', `No ${platform} data sheet found.\n\nPlease create a sheet named "${rawDataSheetName}" or make sure you have a sheet with "${platformName}" in the name.`, ui.ButtonSet.OK);
      return false;
    } else if (similarSheets.length === 1) {
      rawDataSheet = similarSheets[0];
      ui.alert('📋 Info', `Using sheet: "${rawDataSheet.getName()}" for ${platform} data.`, ui.ButtonSet.OK);
    } else {
      // Multiple sheets found - let user choose
      const sheetNames = similarSheets.map(sheet => sheet.getName());
      
      if (similarSheets.length === 2) {
        const choice = ui.alert(
          `Multiple ${platform} sheets found`,
          `Found these sheets:\n\n${sheetNames.join('\n')}\n\nWould you like to:\nYES = Use "${sheetNames[0]}"\nNO = Use "${sheetNames[1]}"\nCANCEL = Combine both sheets`,
          ui.ButtonSet.YES_NO_CANCEL
        );
        
        if (choice === ui.Button.YES) {
          rawDataSheet = similarSheets[0];
        } else if (choice === ui.Button.NO) {
          rawDataSheet = similarSheets[1];
        } else {
          // Combine both sheets
          return processMultipleSheets(similarSheets, platform);
        }
      } else {
        // More than 2 sheets - offer combine option
        const choice = ui.alert(
          `Multiple ${platform} sheets found`,
          `Found ${similarSheets.length} sheets:\n\n${sheetNames.slice(0, 5).join('\n')}${sheetNames.length > 5 ? '\n...' : ''}\n\nWould you like to:\nYES = Use "${sheetNames[0]}" only\nNO = Combine ALL ${platform} sheets\nCANCEL = Skip this platform`,
          ui.ButtonSet.YES_NO_CANCEL
        );
        
        if (choice === ui.Button.CANCEL) {
          return false;
        } else if (choice === ui.Button.YES) {
          rawDataSheet = similarSheets[0];
        } else {
          // Combine all sheets
          return processMultipleSheets(similarSheets, platform);
        }
      }
      
      ui.alert('📋 Selected', `Using sheet: "${rawDataSheet.getName()}" for ${platform} data.`, ui.ButtonSet.OK);
    }
  }
  
  // Get raw data
  const data = rawDataSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert('⚠️ Warning', `No data found in "${rawDataSheet.getName()}".\n\nPlease paste your ${platform} data and try again.`, ui.ButtonSet.OK);
    return false;
  }
  
  // Process the data
  const summaries = createPivotSummaries(data, platform);
  
  // Update template
  updateTemplate(summaries);
  
  return true;
}

/**
 * Process multiple sheets and combine their data
 */
function processMultipleSheets(sheets, platform) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    let combinedData = [];
    let headerRow = null;
    
    // Process each sheet
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const data = sheet.getDataRange().getValues();
      
      if (data.length < 2) {
        ui.alert('⚠️ Warning', `Sheet "${sheet.getName()}" has no data. Skipping...`, ui.ButtonSet.OK);
        continue;
      }
      
      // Use first sheet's headers as the standard
      if (headerRow === null) {
        headerRow = data[0];
        combinedData.push(headerRow);
      } else {
        // Verify headers match (optional - for data consistency)
        const currentHeaders = data[0];
        const headerMismatch = headerRow.some((header, index) => 
          header !== currentHeaders[index]
        );
        
        if (headerMismatch) {
          const proceed = ui.alert(
            '⚠️ Header Mismatch',
            `Sheet "${sheet.getName()}" has different column headers.\n\nContinue anyway?\n\nYES = Continue (may cause errors)\nNO = Skip this sheet\nCANCEL = Stop processing`,
            ui.ButtonSet.YES_NO_CANCEL
          );
          
          if (proceed === ui.Button.CANCEL) {
            return false;
          } else if (proceed === ui.Button.NO) {
            continue;
          }
        }
      }
      
      // Add data rows (skip header)
      for (let j = 1; j < data.length; j++) {
        combinedData.push(data[j]);
      }
    }
    
    if (combinedData.length < 2) {
      ui.alert('⚠️ Warning', `No valid data found in any ${platform} sheets.`, ui.ButtonSet.OK);
      return false;
    }
    
    // Show summary of combined data
    const sheetNames = sheets.map(s => s.getName()).join(', ');
    ui.alert(
      '📊 Data Combined', 
      `Successfully combined data from:\n\n${sheetNames}\n\nTotal rows: ${combinedData.length - 1} (excluding header)`, 
      ui.ButtonSet.OK
    );
    
    // Process the combined data
    const summaries = createPivotSummaries(combinedData, platform);
    
    // Update template
    updateTemplate(summaries);
    
    return true;
    
  } catch (error) {
    ui.alert('❌ Error', `Error combining sheets: ${error.toString()}`, ui.ButtonSet.OK);
    return false;
  }
}

/**
 * Create pivot table summaries from raw data
 */
function createPivotSummaries(data, platform) {
  const config = CONFIG[platform];
  const headers = data[0];
  
  // Find column indices
  const skuIndex = headers.indexOf(config.sku);
  const cityIndex = headers.indexOf(config.city);
  const dateIndex = headers.indexOf(config.date);
  const qtyIndex = headers.indexOf(config.qty);
  
  if (skuIndex === -1 || cityIndex === -1 || dateIndex === -1 || qtyIndex === -1) {
    throw new Error(`Missing columns. Expected: ${config.sku}, ${config.city}, ${config.date}, ${config.qty}`);
  }
  
  // Process data rows
  const processedData = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows or rows with missing critical data
    if (!row || row.length === 0) continue;
    
    const skuValue = row[skuIndex];
    const cityValue = row[cityIndex];
    const dateValue = row[dateIndex];
    const qtyValue = row[qtyIndex];
    
    // Check if all required fields have valid data
    if (!skuValue || !cityValue || !dateValue || 
        skuValue === '' || cityValue === '' || dateValue === '' ||
        skuValue === null || cityValue === null || dateValue === null) {
      continue; // Skip invalid rows
    }
    
    // Try to parse the date
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) {
      console.log(`Skipping row ${i+1}: Invalid date "${dateValue}"`);
      continue; // Skip rows with invalid dates
    }
    
    // Parse quantity - default to 0 if invalid
    let qty = 0;
    if (qtyValue !== null && qtyValue !== undefined && qtyValue !== '') {
      const parsedQty = Number(qtyValue);
      if (!isNaN(parsedQty)) {
        qty = parsedQty;
      }
    }
    
    const weekRange = getWeekRange(date);
    
    processedData.push({
      sku: skuValue.toString().trim(),
      city: cityValue.toString().trim(),
      weekRange: weekRange,
      qty: qty
    });
  }
  
  // Create pivot summaries
  const weeks = [...new Set(processedData.map(row => row.weekRange))].sort();
  
  // SKU pivot
  const skuPivot = {};
  processedData.forEach(row => {
    if (!skuPivot[row.sku]) {
      skuPivot[row.sku] = {};
      weeks.forEach(week => skuPivot[row.sku][week] = 0);
    }
    skuPivot[row.sku][row.weekRange] += row.qty;
  });
  
  // City pivot
  const cityPivot = {};
  processedData.forEach(row => {
    if (!cityPivot[row.city]) {
      cityPivot[row.city] = {};
      weeks.forEach(week => cityPivot[row.city][week] = 0);
    }
    cityPivot[row.city][row.weekRange] += row.qty;
  });
  
  return {
    weeks: weeks,
    skuPivot: skuPivot,
    cityPivot: cityPivot
  };
}

/**
 * Update the template with pivot data
 */
function updateTemplate(summaries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(CONFIG.templateSheet);
  
  if (!templateSheet) {
    throw new Error(`Template sheet "${CONFIG.templateSheet}" not found!`);
  }
  
  const { weeks, skuPivot, cityPivot } = summaries;
  
  // Clear existing data
  clearTemplateData(templateSheet, weeks.length);
  
  // Update week headers
  if (weeks.length > 0) {
    const weekRange = templateSheet.getRange(CONFIG.weekHeaderRow, CONFIG.firstWeekCol, 1, weeks.length);
    weekRange.setValues([weeks]);
    weekRange.setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
  }
  
  // Update SKU section
  const skuItems = Object.keys(skuPivot).sort();
  updateSection(templateSheet, skuItems, skuPivot, weeks, CONFIG.skuSectionStartRow);
  
  // Update City section
  const cityItems = Object.keys(cityPivot).sort();
  updateSection(templateSheet, cityItems, cityPivot, weeks, CONFIG.citySectionStartRow);
}

/**
 * Clear existing template data
 */
function clearTemplateData(sheet, numWeeks) {
  const maxRows = 50;
  const maxCols = Math.max(numWeeks + 2, 10);
  
  // Clear SKU section
  const skuRange = sheet.getRange(CONFIG.skuSectionStartRow, CONFIG.firstDataCol, maxRows, maxCols);
  skuRange.clearContent();
  
  // Clear City section
  const cityRange = sheet.getRange(CONFIG.citySectionStartRow, CONFIG.firstDataCol, maxRows, maxCols);
  cityRange.clearContent();
  
  // Clear week headers
  const weekRange = sheet.getRange(CONFIG.weekHeaderRow, CONFIG.firstWeekCol, 1, maxCols);
  weekRange.clearContent().clearFormat();
}

/**
 * Update a section (SKU or City) in the template
 */
function updateSection(sheet, items, pivotData, weeks, startRow) {
  items.forEach((item, index) => {
    const row = startRow + index;
    
    // Set item name - ensure it's not undefined
    const itemName = item && item !== 'undefined' ? item : 'Unknown';
    sheet.getRange(row, CONFIG.firstDataCol).setValue(itemName);
    
    // Set weekly quantities
    weeks.forEach((week, weekIndex) => {
      const col = CONFIG.firstWeekCol + weekIndex;
      
      // Ensure we have valid data
      let value = 0;
      if (pivotData[item] && pivotData[item][week] !== undefined && pivotData[item][week] !== null) {
        value = pivotData[item][week];
        // Make sure it's a number
        if (isNaN(value)) {
          value = 0;
        }
      }
      
      sheet.getRange(row, col).setValue(value);
    });
  });
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
 * Show instructions
 */
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '📋 How to Use Auto Pivot',
    'Simple 3-step process:\n\n' +
    '1. Create sheets: "' + CONFIG.zeptoRawData + '" and "' + CONFIG.blinkitRawData + '"\n' +
    '2. Paste your raw Excel data into these sheets\n' +
    '3. Click "🔄 Auto Pivot" → "📊 Create Pivot Tables & Update Template"\n\n' +
    'That\'s it! No more manual pivot table creation.\n\n' +
    'Your "' + CONFIG.templateSheet + '" will be automatically updated with:\n' +
    '• Week headers\n' +
    '• SKU summary (quantities by week)\n' +
    '• City summary (quantities by week)\n\n' +
    'Expected columns:\n' +
    '• Zepto: ' + CONFIG.zepto.sku + ', ' + CONFIG.zepto.city + ', ' + CONFIG.zepto.date + ', ' + CONFIG.zepto.qty + '\n' +
    '• Blinkit: ' + CONFIG.blinkit.sku + ', ' + CONFIG.blinkit.city + ', ' + CONFIG.blinkit.date + ', ' + CONFIG.blinkit.qty,
    ui.ButtonSet.OK
  );
} 