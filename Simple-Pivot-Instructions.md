# 🔄 Simple Pivot Table Automation

**Problem Solved:** Eliminates the manual step of creating pivot tables and pasting them into your template.

## What This Does

Instead of manually:
1. Creating pivot tables from raw data
2. Copying the pivot results 
3. Pasting into your template

Now you just:
1. Paste raw data
2. Click one button
3. Template is automatically updated!

## Setup (One-Time)

### Step 1: Install the Script
1. Open your Google Sheet with the template
2. Go to **Extensions** → **Apps Script**
3. Delete any existing code
4. Copy the entire contents of `simple-pivot-automation.js`
5. Paste it into the Apps Script editor
6. Click **Save** (💾)
7. Close the Apps Script tab

### Step 2: Create Data Input Sheets
In your Google Sheet, create these new sheets:
- `Zepto Raw Data`
- `Blinkit Raw Data`

## How to Use

### Weekly Workflow
1. **Paste Raw Data**
   - Copy your Excel data from Zepto → paste into "Zepto Raw Data" sheet
   - Copy your Excel data from Blinkit → paste into "Blinkit Raw Data" sheet
   
2. **Run Auto Pivot**
   - Click the **🔄 Auto Pivot** menu (appears at top after setup)
   - Select **📊 Create Pivot Tables & Update Template**
   - Choose which data to process (or process both)
   
3. **Done!**
   - Your template is automatically updated with pivot summaries
   - Week headers are added automatically
   - SKU and City sections are populated

## Expected Data Format

### Zepto Data Columns:
- `SKU Name`
- `City` 
- `Sales Date`
- `Quantity`

### Blinkit Data Columns:
- `item_name`
- `city_name`
- `date`
- `qty`

## What Gets Updated

The script automatically creates pivot table summaries and updates:
- **Week headers** (Monday-Sunday format like "28-Oct - 3-Nov")
- **SKU section** with quantities by week
- **City section** with quantities by week

## Configuration

If your template layout is different, you can adjust these settings in the script:

```javascript
const CONFIG = {
  // Template positions
  skuSectionStartRow: 3,      // Where SKU data starts
  citySectionStartRow: 25,    // Where City data starts  
  weekHeaderRow: 2,           // Where week headers go
  firstDataCol: 2,            // Column B for names
  firstWeekCol: 3,            // Column C for first week
  
  // Your template sheet name
  templateSheet: "Template for Blinkit"
};
```

## Benefits

✅ **No more manual pivot tables**  
✅ **Consistent formatting**  
✅ **Automatic week calculations**  
✅ **Error handling**  
✅ **One-click updates**  
✅ **Preserves template structure**

## Troubleshooting

**"Sheet not found" error:**
- Make sure you created the input sheets with exact names:
  - `Zepto Raw Data`
  - `Blinkit Raw Data`

**"Missing columns" error:**
- Check that your data has the expected column headers
- Column names must match exactly (case-sensitive)

**Template not updating:**
- Verify your template sheet is named `Template for Blinkit`
- Check the row/column settings in CONFIG match your template layout

## Support

If you need to adjust the template positions or column names, edit the `CONFIG` section at the top of the script in Apps Script editor.

---

**This solution eliminates your manual pivot table workflow while keeping everything simple and automated!** 🎯 