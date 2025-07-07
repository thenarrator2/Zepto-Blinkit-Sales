/**
 * Sales Data Dashboard - Google Apps Script
 * 
 * This script processes Zepto and Blinkit sales data directly in Google Sheets
 * No external hosting or credentials needed - everything runs in Google Sheets
 */

// Configuration - Users can modify these if their Excel format is different
const CONFIG = {
  zepto: {
    sheetName: "Zepto Sales",
    skuColumn: "SKU Name",
    cityColumn: "City", 
    dateColumn: "Sales Date",
    qtyColumn: "Quantity"
  },
  blinkit: {
    sheetName: "Blinkit Sales",
    skuColumn: "item_name",
    cityColumn: "city_name",
    dateColumn: "date", 
    qtyColumn: "qty"
  },
  // Analytics configuration
  analytics: {
    lowStockThreshold: 50, // Alert if weekly sales < 50 units
    topPerformersCount: 10, // Show top 10 performers
    revenueEstimate: { // Estimated revenue per unit (can be customized)
      defaultPrice: 100, // Default price per unit in INR
      zepto: 120, // Average price per unit for Zepto
      blinkit: 110 // Average price per unit for Blinkit
    },
    alertThresholds: {
      lowSales: 25, // Alert if any SKU-City combo < 25 units in a week
      highSales: 1000 // Highlight if any SKU-City combo > 1000 units in a week
    }
  }
};

// Removed the generic processSalesData function - now using direct buttons

/**
 * Process Zepto sales data
 */
function processZeptoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get Zepto data
    const zeptoSheet = ss.getSheetByName(CONFIG.zepto.sheetName);
    if (!zeptoSheet) {
      throw new Error(`Sheet "${CONFIG.zepto.sheetName}" not found. Please create it and paste your Zepto data.`);
    }
    
    const data = getSheetData(zeptoSheet, CONFIG.zepto);
    
    if (data.length === 0) {
      throw new Error('No data found in Zepto sheet');
    }
    
    // Process and create summaries
    const summaries = createSummaries(data, 'Zepto');
    
    // Update sheets
    updateSummarySheets(ss, summaries, 'Zepto');
    
    SpreadsheetApp.getUi().alert('Success!', 'Zepto data processed successfully!\n\nCheck the following sheets:\n- Zepto - SKU Summary\n- Zepto - City Summary\n- Zepto - SKU-City Summary', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error processing Zepto data: ' + error.toString());
  }
}

/**
 * Process Blinkit sales data
 */
function processBlinkitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get Blinkit data
    const blinkitSheet = ss.getSheetByName(CONFIG.blinkit.sheetName);
    if (!blinkitSheet) {
      throw new Error(`Sheet "${CONFIG.blinkit.sheetName}" not found. Please create it and paste your Blinkit data.`);
    }
    
    const data = getSheetData(blinkitSheet, CONFIG.blinkit);
    
    if (data.length === 0) {
      throw new Error('No data found in Blinkit sheet');
    }
    
    // Process and create summaries
    const summaries = createSummaries(data, 'Blinkit');
    
    // Update sheets
    updateSummarySheets(ss, summaries, 'Blinkit');
    
    SpreadsheetApp.getUi().alert('Success!', 'Blinkit data processed successfully!\n\nCheck the following sheets:\n- Blinkit - SKU Summary\n- Blinkit - City Summary\n- Blinkit - SKU-City Summary', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error processing Blinkit data: ' + error.toString());
  }
}

/**
 * Get data from a sheet with proper column mapping
 */
function getSheetData(sheet, config) {
  const data = sheet.getDataRange().getValues();
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
      console.log(`Skipping row ${i+1}: Missing required data`);
      continue;
    }
    
    // Validate the date
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) {
      console.log(`Skipping row ${i+1}: Invalid date "${dateValue}"`);
      continue;
    }
    
    // Parse quantity
    let qty = 0;
    if (qtyValue !== null && qtyValue !== undefined && qtyValue !== '') {
      const parsedQty = Number(qtyValue);
      if (!isNaN(parsedQty)) {
        qty = parsedQty;
      }
    }
    
    cleanData.push({
      sku: skuValue.toString().trim(),
      city: cityValue.toString().trim(),
      date: date,
      qty: qty
    });
  }
  
  return cleanData;
}

/**
 * Create summaries from data
 */
function createSummaries(data, platform) {
  // Add week ranges and filter out invalid dates
  const validData = [];
  data.forEach(row => {
    const weekRange = getWeekRange(row.date);
    if (weekRange !== null) {
      row.weekRange = weekRange;
      validData.push(row);
    } else {
      console.log(`Skipping row with invalid date: ${row.date}`);
    }
  });
  
  if (validData.length === 0) {
    throw new Error('No valid data found after date validation');
  }
  
  // Get unique weeks (filter out any null/undefined values)
  const weeks = [...new Set(validData.map(row => row.weekRange))].filter(week => week !== null && week !== undefined).sort();
  
  // Create SKU summary
  const skuSummary = createPivotTable(validData, 'sku', weeks);
  
  // Create City summary
  const citySummary = createPivotTable(validData, 'city', weeks);
  
  // Create SKU-City summary
  validData.forEach(row => {
    row.skuCity = `${row.city} - ${row.sku}`;
  });
  const skuCitySummary = createPivotTable(validData, 'skuCity', weeks);
  
  return {
    sku: { data: skuSummary, weeks: weeks },
    city: { data: citySummary, weeks: weeks },
    skuCity: { data: skuCitySummary, weeks: weeks }
  };
}

/**
 * Create pivot table data
 */
function createPivotTable(data, groupBy, weeks) {
  const result = {};
  
  // Initialize result structure
  data.forEach(row => {
    const key = row[groupBy];
    if (!result[key]) {
      result[key] = {};
      weeks.forEach(week => {
        result[key][week] = 0;
      });
    }
  });
  
  // Aggregate data
  data.forEach(row => {
    const key = row[groupBy];
    const week = row.weekRange;
    result[key][week] += row.qty;
  });
  
  return result;
}

/**
 * Get week range string (Monday to Sunday)
 */
function getWeekRange(date) {
  // Validate the input date
  const d = new Date(date);
  if (isNaN(d.getTime())) {
    console.log(`Invalid date encountered: ${date}`);
    return null; // Return null for invalid dates
  }
  
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Adjust for Sunday
  const monday = new Date(d.setDate(diff));
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  const formatDate = (date) => {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const day = date.getDate();
    const month = date.getMonth();
    
    // Validate day and month
    if (isNaN(day) || isNaN(month) || month < 0 || month > 11) {
      return 'Invalid';
    }
    
    return `${day}-${months[month]}`;
  };
  
  const mondayStr = formatDate(monday);
  const sundayStr = formatDate(sunday);
  
  // Check if formatting was successful
  if (mondayStr === 'Invalid' || sundayStr === 'Invalid') {
    return null;
  }
  
  return `${mondayStr} - ${sundayStr}`;
}

/**
 * Update summary sheets with data
 */
function updateSummarySheets(ss, summaries, platform) {
  const types = ['sku', 'city', 'skuCity'];
  const names = ['SKU Summary', 'City Summary', 'SKU-City Summary'];
  
  types.forEach((type, index) => {
    const sheetName = `${platform} - ${names[index]}`;
    const summary = summaries[type];
    
    // Create or get sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Prepare data for sheet
    const weeks = summary.weeks;
    const headers = ['Item', ...weeks];
    const rows = [headers];
    
    // Add data rows
    Object.keys(summary.data).forEach(key => {
      const row = [key];
      weeks.forEach(week => {
        row.push(summary.data[key][week]);
      });
      rows.push(row);
    });
    
    // Write to sheet
    if (rows.length > 1) {
      sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
      
      // Format headers
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285f4')
        .setFontColor('white')
        .setFontWeight('bold');
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, headers.length);
    }
  });
}

/**
 * Generate Low Stock Alerts
 */
function generateLowStockAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const alerts = [];
    const platforms = ['Zepto', 'Blinkit'];
    
    platforms.forEach(platform => {
      const skuSheet = ss.getSheetByName(`${platform} - SKU Summary`);
      if (skuSheet) {
        const data = skuSheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          const lastWeekIndex = headers.length - 1; // Most recent week
          
          for (let i = 1; i < data.length; i++) {
            const sku = data[i][0];
            const lastWeekSales = data[i][lastWeekIndex] || 0;
            
            if (lastWeekSales < CONFIG.analytics.lowStockThreshold) {
              alerts.push({
                platform: platform,
                sku: sku,
                lastWeekSales: lastWeekSales,
                week: headers[lastWeekIndex]
              });
            }
          }
        }
      }
    });
    
    // Create alerts sheet
    createAlertsSheet(ss, alerts, 'Low Stock Alerts');
    
    if (alerts.length > 0) {
      SpreadsheetApp.getUi().alert('Low Stock Alert!', 
        `Found ${alerts.length} SKUs with low sales (< ${CONFIG.analytics.lowStockThreshold} units).\n\nCheck the "Low Stock Alerts" sheet for details.`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('No Low Stock Issues', 
        'All SKUs are performing above the threshold!', 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating alerts: ' + error.toString());
  }
}

/**
 * Generate Top Performers Report
 */
function generateTopPerformers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const performers = [];
    const platforms = ['Zepto', 'Blinkit'];
    
    platforms.forEach(platform => {
      const skuSheet = ss.getSheetByName(`${platform} - SKU Summary`);
      if (skuSheet) {
        const data = skuSheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          
          for (let i = 1; i < data.length; i++) {
            const sku = data[i][0];
            let totalSales = 0;
            
            // Sum all weeks (skip first column which is SKU name)
            for (let j = 1; j < data[i].length; j++) {
              totalSales += data[i][j] || 0;
            }
            
            performers.push({
              platform: platform,
              sku: sku,
              totalSales: totalSales
            });
          }
        }
      }
    });
    
    // Sort by total sales and take top performers
    performers.sort((a, b) => b.totalSales - a.totalSales);
    const topPerformers = performers.slice(0, CONFIG.analytics.topPerformersCount);
    
    // Create top performers sheet
    createTopPerformersSheet(ss, topPerformers);
    
    SpreadsheetApp.getUi().alert('Top Performers Generated!', 
      `Top ${CONFIG.analytics.topPerformersCount} performing SKUs analysis is ready.\n\nCheck the "Top Performers" sheet.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating top performers: ' + error.toString());
  }
}

/**
 * Generate Performance Trends Analysis
 */
function generatePerformanceTrends() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const trends = [];
    const platforms = ['Zepto', 'Blinkit'];
    
    platforms.forEach(platform => {
      const skuSheet = ss.getSheetByName(`${platform} - SKU Summary`);
      if (skuSheet) {
        const data = skuSheet.getDataRange().getValues();
        if (data.length > 1 && data[0].length > 3) { // Need at least 2 weeks of data
          const headers = data[0];
          
          for (let i = 1; i < data.length; i++) {
            const sku = data[i][0];
            const weekData = [];
            
            // Get last 4 weeks of data (or available weeks)
            const startWeek = Math.max(1, headers.length - 4);
            for (let j = startWeek; j < headers.length; j++) {
              weekData.push(data[i][j] || 0);
            }
            
            if (weekData.length >= 2) {
              const trend = calculateTrend(weekData);
              trends.push({
                platform: platform,
                sku: sku,
                trend: trend,
                weeks: weekData,
                weekHeaders: headers.slice(startWeek)
              });
            }
          }
        }
      }
    });
    
    // Create trends sheet
    createTrendsSheet(ss, trends);
    
    SpreadsheetApp.getUi().alert('Trends Analysis Complete!', 
      `Performance trends analysis is ready.\n\nCheck the "Performance Trends" sheet for insights.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating trends: ' + error.toString());
  }
}

/**
 * Generate City Performance Analysis
 */
function generateCityPerformance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const cityPerformance = [];
    const platforms = ['Zepto', 'Blinkit'];
    
    platforms.forEach(platform => {
      const citySheet = ss.getSheetByName(`${platform} - City Summary`);
      if (citySheet) {
        const data = citySheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          
          for (let i = 1; i < data.length; i++) {
            const city = data[i][0];
            let totalSales = 0;
            let avgWeeklySales = 0;
            let weekCount = 0;
            
            // Calculate totals and averages
            for (let j = 1; j < data[i].length; j++) {
              const sales = data[i][j] || 0;
              totalSales += sales;
              if (sales > 0) weekCount++;
            }
            
            avgWeeklySales = weekCount > 0 ? totalSales / weekCount : 0;
            
            cityPerformance.push({
              platform: platform,
              city: city,
              totalSales: totalSales,
              avgWeeklySales: Math.round(avgWeeklySales),
              activeWeeks: weekCount
            });
          }
        }
      }
    });
    
    // Sort by total sales
    cityPerformance.sort((a, b) => b.totalSales - a.totalSales);
    
    // Create city performance sheet
    createCityPerformanceSheet(ss, cityPerformance);
    
    SpreadsheetApp.getUi().alert('City Performance Analysis Complete!', 
      `City performance analysis is ready.\n\nCheck the "City Performance" sheet.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating city performance: ' + error.toString());
  }
}

/**
 * Generate Revenue Analysis
 */
function generateRevenueAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const revenueData = [];
    const platforms = ['Zepto', 'Blinkit'];
    
    platforms.forEach(platform => {
      const skuSheet = ss.getSheetByName(`${platform} - SKU Summary`);
      if (skuSheet) {
        const data = skuSheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          const pricePerUnit = CONFIG.analytics.revenueEstimate[platform.toLowerCase()] || CONFIG.analytics.revenueEstimate.defaultPrice;
          
          for (let i = 1; i < data.length; i++) {
            const sku = data[i][0];
            let totalSales = 0;
            
            // Sum all weeks
            for (let j = 1; j < data[i].length; j++) {
              totalSales += data[i][j] || 0;
            }
            
            const estimatedRevenue = totalSales * pricePerUnit;
            
            revenueData.push({
              platform: platform,
              sku: sku,
              totalUnits: totalSales,
              pricePerUnit: pricePerUnit,
              estimatedRevenue: estimatedRevenue
            });
          }
        }
      }
    });
    
    // Sort by revenue
    revenueData.sort((a, b) => b.estimatedRevenue - a.estimatedRevenue);
    
    // Create revenue sheet
    createRevenueSheet(ss, revenueData);
    
    const totalRevenue = revenueData.reduce((sum, item) => sum + item.estimatedRevenue, 0);
    
    SpreadsheetApp.getUi().alert('Revenue Analysis Complete!', 
      `Total estimated revenue: ₹${totalRevenue.toLocaleString()}\n\nCheck the "Revenue Analysis" sheet for detailed breakdown.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating revenue analysis: ' + error.toString());
  }
}

/**
 * Helper function to calculate trend
 */
function calculateTrend(weekData) {
  if (weekData.length < 2) return 'Insufficient Data';
  
  const first = weekData[0];
  const last = weekData[weekData.length - 1];
  
  if (first === 0 && last === 0) return 'No Sales';
  if (first === 0) return 'New Entry';
  
  const change = ((last - first) / first) * 100;
  
  if (change > 10) return 'Growing';
  if (change < -10) return 'Declining';
  return 'Stable';
}

/**
 * Create alerts sheet
 */
function createAlertsSheet(ss, alerts, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  const headers = ['Platform', 'SKU', 'Last Week Sales', 'Week', 'Alert Level'];
  const rows = [headers];
  
  alerts.forEach(alert => {
    let alertLevel = 'Medium';
    if (alert.lastWeekSales < CONFIG.analytics.alertThresholds.lowSales) {
      alertLevel = 'High';
    }
    
    rows.push([
      alert.platform,
      alert.sku,
      alert.lastWeekSales,
      alert.week,
      alertLevel
    ]);
  });
  
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    
    // Format headers
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#ff6b6b')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Color code alert levels
    for (let i = 2; i <= rows.length; i++) {
      const alertLevel = sheet.getRange(i, 5).getValue();
      if (alertLevel === 'High') {
        sheet.getRange(i, 1, 1, headers.length).setBackground('#ffe6e6');
      }
    }
    
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Create top performers sheet
 */
function createTopPerformersSheet(ss, performers) {
  const sheetName = 'Top Performers';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  const headers = ['Rank', 'Platform', 'SKU', 'Total Sales', 'Performance Level'];
  const rows = [headers];
  
  performers.forEach((performer, index) => {
    let level = 'Good';
    if (index < 3) level = 'Excellent';
    else if (index < 7) level = 'Very Good';
    
    rows.push([
      index + 1,
      performer.platform,
      performer.sku,
      performer.totalSales,
      level
    ]);
  });
  
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    
    // Format headers
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4caf50')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Color code top 3
    for (let i = 2; i <= Math.min(4, rows.length); i++) {
      sheet.getRange(i, 1, 1, headers.length).setBackground('#e8f5e8');
    }
    
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Create trends sheet
 */
function createTrendsSheet(ss, trends) {
  const sheetName = 'Performance Trends';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  const headers = ['Platform', 'SKU', 'Trend', 'Recent Weeks Sales', 'Recommendation'];
  const rows = [headers];
  
  trends.forEach(trend => {
    let recommendation = 'Monitor';
    if (trend.trend === 'Growing') recommendation = 'Increase Stock';
    else if (trend.trend === 'Declining') recommendation = 'Review Strategy';
    else if (trend.trend === 'New Entry') recommendation = 'Track Performance';
    
    rows.push([
      trend.platform,
      trend.sku,
      trend.trend,
      trend.weeks.join(', '),
      recommendation
    ]);
  });
  
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    
    // Format headers
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2196f3')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Color code trends
    for (let i = 2; i <= rows.length; i++) {
      const trendValue = sheet.getRange(i, 3).getValue();
      if (trendValue === 'Growing') {
        sheet.getRange(i, 3).setBackground('#c8e6c9');
      } else if (trendValue === 'Declining') {
        sheet.getRange(i, 3).setBackground('#ffcdd2');
      }
    }
    
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Create city performance sheet
 */
function createCityPerformanceSheet(ss, cityData) {
  const sheetName = 'City Performance';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  const headers = ['Platform', 'City', 'Total Sales', 'Avg Weekly Sales', 'Active Weeks', 'Performance Rating'];
  const rows = [headers];
  
  cityData.forEach((city, index) => {
    let rating = 'Average';
    if (index < 5) rating = 'Top Performer';
    else if (index < 10) rating = 'Good';
    else if (city.totalSales < 100) rating = 'Needs Attention';
    
    rows.push([
      city.platform,
      city.city,
      city.totalSales,
      city.avgWeeklySales,
      city.activeWeeks,
      rating
    ]);
  });
  
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    
    // Format headers
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#9c27b0')
      .setFontColor('white')
      .setFontWeight('bold');
    
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Create revenue sheet
 */
function createRevenueSheet(ss, revenueData) {
  const sheetName = 'Revenue Analysis';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  const headers = ['Platform', 'SKU', 'Total Units Sold', 'Price Per Unit (₹)', 'Estimated Revenue (₹)', 'Revenue Share (%)'];
  const rows = [headers];
  
  const totalRevenue = revenueData.reduce((sum, item) => sum + item.estimatedRevenue, 0);
  
  revenueData.forEach(item => {
    const revenueShare = totalRevenue > 0 ? ((item.estimatedRevenue / totalRevenue) * 100).toFixed(2) : 0;
    
    rows.push([
      item.platform,
      item.sku,
      item.totalUnits,
      item.pricePerUnit,
      Math.round(item.estimatedRevenue),
      revenueShare + '%'
    ]);
  });
  
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    
    // Format headers
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#ff9800')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Format currency columns
    sheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('₹#,##0');
    
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Create menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Sales Data Dashboard')
    .addItem('🟢 Process Zepto Data', 'processZeptoData')
    .addItem('🔵 Process Blinkit Data', 'processBlinkitData')
    .addSeparator()
    .addSubMenu(ui.createMenu('📈 Analytics & Alerts')
      .addItem('⚠️ Low Stock Alerts', 'generateLowStockAlerts')
      .addItem('🏆 Top Performers', 'generateTopPerformers')
      .addItem('📊 Performance Trends', 'generatePerformanceTrends')
      .addItem('📍 City Performance', 'generateCityPerformance')
      .addItem('💰 Revenue Analysis', 'generateRevenueAnalysis'))
    .addSeparator()
    .addItem('📋 Setup Instructions', 'showInstructions')
    .addItem('⚙️ Configuration', 'showConfiguration')
    .addToUi();
}

/**
 * Show setup instructions
 */
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Instructions',
    'How to use this Sales Data Dashboard:\n\n' +
    '📊 BASIC PROCESSING:\n' +
    '1. Create sheets named "Zepto Sales" and "Blinkit Sales"\n' +
    '2. Paste your Excel data into these sheets\n' +
    '3. Use the menu buttons to process data\n' +
    '4. Results appear in new summary sheets\n\n' +
    '📈 ANALYTICS & ALERTS:\n' +
    '• Low Stock Alerts: Find SKUs with declining sales\n' +
    '• Top Performers: Identify best-selling products\n' +
    '• Performance Trends: Track growth/decline patterns\n' +
    '• City Performance: Compare sales by location\n' +
    '• Revenue Analysis: Estimate revenue by product\n\n' +
    'Expected columns:\n' +
    '• Zepto: SKU Name, City, Sales Date, Quantity\n' +
    '• Blinkit: item_name, city_name, date, qty\n\n' +
    'Run basic processing first, then use analytics features!',
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
    'Customize settings in Extensions > Apps Script:\n\n' +
    '📊 DATA COLUMNS:\n' +
    '• Zepto: ' + CONFIG.zepto.skuColumn + ', ' + CONFIG.zepto.cityColumn + ', ' + CONFIG.zepto.dateColumn + ', ' + CONFIG.zepto.qtyColumn + '\n' +
    '• Blinkit: ' + CONFIG.blinkit.skuColumn + ', ' + CONFIG.blinkit.cityColumn + ', ' + CONFIG.blinkit.dateColumn + ', ' + CONFIG.blinkit.qtyColumn + '\n\n' +
    '📈 ANALYTICS SETTINGS:\n' +
    '• Low Stock Threshold: ' + CONFIG.analytics.lowStockThreshold + ' units/week\n' +
    '• Top Performers Count: ' + CONFIG.analytics.topPerformersCount + '\n' +
    '• Zepto Price/Unit: ₹' + CONFIG.analytics.revenueEstimate.zepto + '\n' +
    '• Blinkit Price/Unit: ₹' + CONFIG.analytics.revenueEstimate.blinkit + '\n' +
    '• Low Sales Alert: < ' + CONFIG.analytics.alertThresholds.lowSales + ' units\n' +
    '• High Sales Alert: > ' + CONFIG.analytics.alertThresholds.highSales + ' units\n\n' +
    'Edit CONFIG object in script to customize these values.',
    ui.ButtonSet.OK
  );
} 