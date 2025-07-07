# Sales Data Automation with Zapier

## 🚀 **Super Simple Solution with Zapier!**

Using Zapier makes this **completely automated** - no coding, no manual buttons, just pure automation!

---

## 🎯 **How This Works**

### **Current Manual Process:**
1. Get Zepto/Blinkit Excel files
2. Upload to Google Sheets
3. Process data manually
4. Update template

### **New Automated Process:**
1. **Email Excel files** to a specific address OR **Upload to Google Drive**
2. **Zapier automatically processes** and updates your Google Sheet template
3. **Done!** - No manual work needed

---

## 📋 **Setup Options (Choose One)**

## **Option 1: Email-Based Automation** ⭐ *Recommended*

### **Workflow:**
1. **Email trigger**: Client emails Excel files to `sales-data@yourdomain.com`
2. **Zapier extracts** Excel attachments
3. **Zapier processes** data and updates Google Sheets template
4. **Zapier sends** confirmation email with updated sheet link

### **Benefits:**
- **Super simple** for client - just email files
- **Works from anywhere** - phone, computer, etc.
- **No Google Drive/folder management**
- **Automatic notifications** when processing is complete

---

## **Option 2: Google Drive Automation**

### **Workflow:**
1. **Drive trigger**: Client uploads Excel files to specific Google Drive folder
2. **Zapier detects** new files
3. **Zapier processes** data and updates Google Sheets template
4. **Zapier moves** processed files to "Completed" folder

### **Benefits:**
- **Familiar Google Drive** interface
- **Easy file organization**
- **No email setup required**

---

## 📊 **Detailed Zapier Setup**

### **Zap Components:**

#### **Trigger** (Choose one):
- **Email Parser** - Receives Excel files via email
- **Google Drive** - Watches for new files in folder

#### **Actions:**
1. **Extract/Download** Excel files
2. **Code by Zapier (Python)** - Process Excel data into summaries
3. **Google Sheets** - Update your template sheet
4. **Email/Slack** - Send completion notification

---

## 🔧 **Step-by-Step Setup**

### **Step 1: Create Email Parser (Option 1) or Drive Folder (Option 2)**

#### **For Email Parser:**
1. Go to [parser.zapier.com](https://parser.zapier.com)
2. Create new mailbox: `sales-data-parser`
3. Get email address (e.g., `sales@robot.zapier.com`)
4. Send sample email with Excel attachments to set up parsing

#### **For Google Drive:**
1. Create folder: `Sales Data Input`
2. Create subfolders: `Zepto`, `Blinkit`, `Processed`

### **Step 2: Create the Zap**

1. **Trigger**: Email Parser (new email) OR Google Drive (new file)
2. **Action 1**: Code by Zapier (Python) - Process data
3. **Action 2**: Google Sheets - Update rows
4. **Action 3**: Email - Send notification

### **Step 3: Python Code for Processing**

```python
import pandas as pd
import io
from datetime import datetime, timedelta

def process_sales_data(file_content, platform):
    """Process Excel data and return summaries"""
    
    # Configuration
    config = {
        'zepto': {
            'sheet': 'Zepto Sales',
            'sku_col': 'SKU Name',
            'city_col': 'City',
            'date_col': 'Sales Date',
            'qty_col': 'Quantity'
        },
        'blinkit': {
            'sheet': 'Blinkit Sales',
            'sku_col': 'item_name',
            'city_col': 'city_name', 
            'date_col': 'date',
            'qty_col': 'qty'
        }
    }
    
    # Read Excel file
    df = pd.read_excel(io.BytesIO(file_content), sheet_name=config[platform]['sheet'])
    
    # Standardize columns
    df = df.rename(columns={
        config[platform]['sku_col']: 'sku',
        config[platform]['city_col']: 'city',
        config[platform]['date_col']: 'date',
        config[platform]['qty_col']: 'qty'
    })
    
    # Convert date and create week ranges
    df['date'] = pd.to_datetime(df['date'])
    df['week_start'] = df['date'] - pd.to_timedelta(df['date'].dt.dayofweek, unit='d')
    df['week_end'] = df['week_start'] + pd.Timedelta(days=6)
    df['week_range'] = df['week_start'].dt.strftime('%d-%b') + ' - ' + df['week_end'].dt.strftime('%d-%b')
    
    # Create summaries
    sku_summary = df.groupby(['sku', 'week_range'])['qty'].sum().unstack(fill_value=0)
    city_summary = df.groupby(['city', 'week_range'])['qty'].sum().unstack(fill_value=0)
    
    return {
        'sku_summary': sku_summary.to_dict(),
        'city_summary': city_summary.to_dict(),
        'weeks': list(sku_summary.columns)
    }

# Main Zapier function
def main(input_data):
    file_content = input_data.get('file_content')  # From email attachment or Drive file
    filename = input_data.get('filename', '').lower()
    
    # Determine platform from filename
    platform = 'zepto' if 'zepto' in filename else 'blinkit'
    
    # Process data
    results = process_sales_data(file_content, platform)
    
    return {
        'platform': platform,
        'summaries': results,
        'processed_at': datetime.now().isoformat()
    }
```

---

## 🎯 **Client Workflow Options**

### **Email Workflow** (Simplest):
1. **Client emails Excel files** to: `sales-data@yourparser.zapier.com`
2. **Subject**: "Zepto Weekly Sales" or "Blinkit Weekly Sales"
3. **Attach Excel files**
4. **Send** → Automation runs automatically
5. **Receive confirmation email** with updated sheet link

### **Google Drive Workflow**:
1. **Client uploads** Excel files to Google Drive folder
2. **Names files**: `Zepto_2024_Week_45.xlsx` or `Blinkit_2024_Week_45.xlsx`
3. **Upload** → Automation runs automatically
4. **File moves** to "Processed" folder when complete

---

## 📈 **Advanced Features You Can Add**

### **Smart Notifications:**
- **Slack alerts** when processing completes
- **Email summaries** with key metrics
- **Error notifications** if files can't be processed

### **Data Validation:**
- **Check file formats** before processing
- **Validate column names** and data types
- **Send error reports** for malformed data

### **Multiple Templates:**
- **Route to different sheets** based on file names
- **Update multiple templates** simultaneously
- **Archive old data** automatically

### **Reporting:**
- **Weekly summary emails** with trends
- **Dashboard links** in notifications
- **Performance metrics** tracking

---

## 💰 **Cost Comparison**

### **Zapier Costs:**
- **Free Plan**: 100 tasks/month (might be enough for weekly updates)
- **Starter Plan**: $19.99/month for 750 tasks (definitely enough)
- **Professional Plan**: $49/month for advanced features

### **vs. Other Solutions:**
- **Streamlit hosting**: $0-20/month + maintenance time
- **Google Apps Script**: $0 but manual button clicking
- **Custom development**: $500-2000 + ongoing support

**Zapier = Set it once, works forever!**

---

## 🎉 **Why Zapier is Perfect Here**

### ✅ **Zero Maintenance**
- No servers to manage
- No code to maintain  
- No version updates needed

### ✅ **User-Friendly**
- Client just emails files or uploads to Drive
- No technical knowledge required
- Works from any device

### ✅ **Reliable**
- Enterprise-grade infrastructure
- Automatic retries if something fails
- Built-in error handling

### ✅ **Scalable**
- Handles increasing data volumes
- Add new features easily
- Integrate with other tools

---

## 🚀 **Quick Start Recommendation**

### **Phase 1: Basic Email Automation** (1 hour setup)
1. Create Email Parser for file reception
2. Set up basic Zap with Python processing
3. Update Google Sheets template
4. Send confirmation emails

### **Phase 2: Enhanced Features** (if needed later)
- Add data validation
- Create error handling
- Set up advanced notifications
- Add reporting features

---

## 📞 **Implementation Plan**

### **I can help you:**
1. **Set up the Zapier workflows**
2. **Write the Python processing code**
3. **Configure the Google Sheets integration**
4. **Test with sample data**
5. **Provide client instructions**

### **Your client gets:**
- **Email address** to send Excel files to
- **Simple instructions**: "Email your files here"
- **Automatic updates** to their Google Sheet
- **Zero technical complexity**

---

**Would you like me to start with the Email Parser setup or Google Drive automation?** The email option is probably simpler for most clients since they can send files from anywhere! 