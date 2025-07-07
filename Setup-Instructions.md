# Sales Data Dashboard - Google Apps Script Setup

## 🎉 Super Simple Solution!

This Google Apps Script solution runs **directly inside Google Sheets** - no external hosting, no credentials, no complicated setup!

---

## 📋 One-Time Setup (5 minutes)

### Step 1: Create a New Google Sheet
1. Go to [sheets.google.com](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it "Sales Data Dashboard"

### Step 2: Add the Script
1. In your Google Sheet, go to **Extensions > Apps Script**
2. Delete the default code in the editor
3. Copy and paste the entire code from `google-apps-script.js`
4. Click **Save** (💾)
5. Name your project "Sales Data Dashboard"

### Step 3: Set Permissions
1. Click **Run** (▶️) once to authorize the script
2. Click **Review permissions**
3. Choose your Google account
4. Click **Advanced** → **Go to Sales Data Dashboard (unsafe)**
5. Click **Allow**

### Step 4: Refresh Your Sheet
1. Go back to your Google Sheet
2. Refresh the page (F5 or Ctrl+R)
3. You should see a new menu: **📊 Sales Data Dashboard**

---

## 🚀 Daily Usage (30 seconds)

### For Zepto Data:
1. **Create a sheet** named "Zepto Sales"
2. **Paste your Excel data** into this sheet
3. **Go to menu**: 📊 Sales Data Dashboard → 🔄 Process Zepto Data
4. **Done!** Check the new summary sheets

### For Blinkit Data:
1. **Create a sheet** named "Blinkit Sales"  
2. **Paste your Excel data** into this sheet
3. **Go to menu**: 📊 Sales Data Dashboard → 🔄 Process Blinkit Data
4. **Done!** Check the new summary sheets

---

## 📊 What You Get

The script automatically creates these sheets:
- **Zepto - SKU Summary**: Weekly quantities by product
- **Zepto - City Summary**: Weekly quantities by city
- **Zepto - SKU-City Summary**: Weekly quantities by city-product combination
- **Blinkit - SKU Summary**: Weekly quantities by product
- **Blinkit - City Summary**: Weekly quantities by city  
- **Blinkit - SKU-City Summary**: Weekly quantities by city-product combination

---

## ⚙️ Configuration (If Your Excel Format is Different)

### Expected Column Names:
- **Zepto**: SKU Name, City, Sales Date, Quantity
- **Blinkit**: item_name, city_name, date, qty

### If Your Columns Are Different:
1. Go to **Extensions > Apps Script**
2. Find the `CONFIG` object at the top
3. Update the column names to match your data
4. Save and refresh your sheet

### Example:
```javascript
const CONFIG = {
  zepto: {
    sheetName: "Zepto Sales",
    skuColumn: "Product Name",      // Changed from "SKU Name"
    cityColumn: "Location",         // Changed from "City"
    dateColumn: "Sale Date",        // Changed from "Sales Date"
    qtyColumn: "Quantity Sold"      // Changed from "Quantity"
  },
  // ... rest of config
};
```

---

## ✅ Benefits of This Solution

### For You:
- **No hosting costs** - runs in Google Sheets
- **No maintenance** - Google handles everything
- **No credentials** - uses Google's built-in security
- **Always available** - works wherever Google Sheets works

### For Your Client:
- **Super simple** - just paste data and click menu
- **No technical skills** - familiar Google Sheets interface
- **Instant results** - summaries appear in seconds
- **Always up-to-date** - processes fresh data every time

---

## 🆘 Troubleshooting

### "Sheet not found" error:
- Make sure you created sheets named exactly "Zepto Sales" and "Blinkit Sales"

### "Column not found" error:
- Check that your column names match the expected format
- Use the Configuration menu option to see current settings

### Menu not showing:
- Refresh your Google Sheet
- Make sure you saved the script and authorized it

### Script not working:
- Go to Extensions > Apps Script
- Check the execution log for detailed error messages

---

## 🎯 Perfect For:

- **Non-technical users** who want automation
- **Small businesses** that don't want hosting costs
- **Teams** that already use Google Sheets
- **Quick solutions** that need to work immediately

This is the **simplest possible solution** - everything runs inside Google Sheets with zero external dependencies! 