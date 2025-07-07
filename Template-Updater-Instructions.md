# Sales Data Template Updater - Setup Instructions

## 🎯 Perfect Solution for Updating Existing Templates!

This Google Apps Script **updates your existing template sheet** instead of creating new sheets. Just paste your data and click a button!

---

## 📋 Quick Setup (5 minutes)

### Step 1: Prepare Your Google Sheet Structure
1. **Keep your existing template sheet** (e.g., "Template for Blinkit")
2. **Create 2 new input sheets**:
   - "Zepto Data Input" 
   - "Blinkit Data Input"

### Step 2: Add the Script
1. Go to **Extensions > Apps Script**
2. Delete the default code
3. Copy and paste the code from `google-apps-script-template-updater.js`
4. **Save** and authorize the script
5. **Refresh your Google Sheet**

### Step 3: Configure for Your Template
Update the CONFIG section to match your template layout:

```javascript
const CONFIG = {
  // Your sheet names
  templateSheet: "Template for Blinkit",  // Your actual template sheet name
  
  // Template layout (adjust these numbers to match your template)
  template: {
    skuSummaryStartRow: 3,      // Where SKU summary starts (Row 3)
    skuSummaryStartCol: 2,      // SKU names in Column B
    
    citySummaryStartRow: 25,    // Where City summary starts (Row 25)  
    citySummaryStartCol: 2,     // City names in Column B
    
    weekHeaderRow: 2,           // Week headers in Row 2
    weekStartCol: 3,            // First week column (Column C)
  }
};
```

---

## 🚀 Daily Usage (30 seconds)

### For Zepto Data:
1. **Paste your Excel data** into "Zepto Data Input" sheet
2. **Click menu**: 📊 Sales Dashboard → 🔄 Update Template with New Data
3. **Choose "YES"** for Zepto data
4. **Done!** Your template is updated instantly

### For Blinkit Data:
1. **Paste your Excel data** into "Blinkit Data Input" sheet  
2. **Click menu**: 📊 Sales Dashboard → 🔄 Update Template with New Data
3. **Choose "NO"** for Blinkit data
4. **Done!** Your template is updated instantly

### For Both Together:
1. **Paste data** into both input sheets
2. **Click menu**: 📊 Sales Dashboard → 🔄 Update Template with New Data
3. **Choose "CANCEL"** to process both
4. **Done!** Your template shows combined weekly summaries

---

## 📊 What Happens

### ✅ **Keeps Your Template Format**
- Uses your existing template sheet
- Preserves your formatting and structure
- Only updates the data areas

### ✅ **Smart Data Placement**
- **SKU Summary**: Updates the SKU section with weekly totals
- **City Summary**: Updates the City section with weekly totals  
- **Week Headers**: Automatically adds week ranges (e.g., "28-Oct - 3-Nov")

### ✅ **Clean Updates**
- Clears old data before adding new
- Maintains consistent formatting
- Sorts SKUs and Cities alphabetically

---

## ⚙️ Customization

### If Your Template Layout is Different:

1. **Go to Extensions > Apps Script**
2. **Find the CONFIG object**
3. **Update the row and column numbers** to match your template:

```javascript
template: {
  skuSummaryStartRow: 5,      // Change if SKUs start in different row
  citySummaryStartRow: 30,    // Change if Cities start in different row
  weekHeaderRow: 1,           // Change if week headers are in different row
  weekStartCol: 4,            // Change if first week is in different column
}
```

### If Your Column Names are Different:

```javascript
zepto: {
  skuColumn: "Product Name",    // Instead of "SKU Name"
  cityColumn: "Location",       // Instead of "City"
  dateColumn: "Sale Date",      // Instead of "Sales Date"
  qtyColumn: "Units Sold"       // Instead of "Quantity"
}
```

---

## 📋 Sheet Structure You Need

### Input Sheets (where you paste data):
- **"Zepto Data Input"**: Paste your Zepto Excel data here
- **"Blinkit Data Input"**: Paste your Blinkit Excel data here

### Template Sheet (gets updated automatically):
- **"Template for Blinkit"**: Your existing template (keeps all formatting)

---

## 🎯 Perfect for Your Use Case Because:

### ✅ **No New Sheets Created**
- Updates your existing template
- Preserves all your formatting
- Maintains your familiar structure

### ✅ **One-Click Updates** 
- Just paste data and click button
- No manual copying or formatting
- Instant weekly summaries

### ✅ **Separate Processing**
- Process Zepto and Blinkit independently
- Or combine both in one update
- Flexible workflow

### ✅ **Professional Results**
- Clean, formatted output
- Consistent week ranges
- Sorted data presentation

---

## 🆘 Troubleshooting

### "Template sheet not found":
- Check that your template sheet name matches CONFIG.templateSheet
- Make sure the sheet exists in your spreadsheet

### "Data sheet not found":  
- Create "Zepto Data Input" and "Blinkit Data Input" sheets
- Make sure you're pasting data into the correct input sheets

### Data not appearing in right place:
- Check the row/column numbers in CONFIG.template
- Count rows/columns in your template to get correct positions

### Week headers in wrong place:
- Adjust CONFIG.template.weekHeaderRow and weekStartCol
- Make sure they point to where you want week headers

---

## 📞 Support

Check the built-in help:
- **📊 Sales Dashboard → 📋 Setup Instructions**: Quick reminder
- **📊 Sales Dashboard → ⚙️ Configuration**: Shows current settings

This solution gives you **exactly what you wanted**: 
- ✅ Updates existing template (no new sheets)
- ✅ Simple button click to update
- ✅ Paste data and process instantly
- ✅ Preserves your template formatting 