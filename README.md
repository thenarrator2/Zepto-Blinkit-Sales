# Sales Data Dashboard

A Streamlit web app that automates the processing of Zepto and Blinkit sales data to generate weekly summaries by SKU and City.

## Features

- 📁 Upload Excel files from Zepto and Blinkit
- 📊 Automatic data processing using pandas
- 📈 Weekly summary by SKU 
- 🏙️ Weekly summary by City
- 📥 Download results as Excel file
- 🌐 Web-based interface (no technical skills required)

## How to Use

1. Visit the deployed app (link will be provided after deployment)
2. Upload your Zepto Sales Excel file (should contain "Zepto Sales" sheet)
3. Upload your Blinkit Sales Excel file (should contain "Blinkit Sales" sheet)
4. View the automatically generated summaries
5. Download the Excel summary file

## Local Setup (Optional)

If you want to run this locally:

```bash
# Clone the repository
git clone <your-repo-url>
cd <repo-name>

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

## Deployment to Streamlit Cloud

### Step 1: Create GitHub Repository
1. Create a new repository on GitHub
2. Upload these files:
   - `app.py`
   - `requirements.txt` 
   - `README.md`

### Step 2: Deploy to Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click "New app"
4. Select your repository
5. Set main file path to `app.py`
6. Click "Deploy"

Your app will be live at: `https://your-username-repo-name-main-abc123.streamlit.app`

## File Structure Expected

### Zepto Sales Sheet
- **Sheet Name**: "Zepto Sales"
- **Columns**: SKU Name, City, Sales Date, Quantity

### Blinkit Sales Sheet  
- **Sheet Name**: "Blinkit Sales"
- **Columns**: item_name, city_name, date, qty

## Output

The app generates two summaries:
1. **SKU Summary**: Quantity sold by each SKU for each week
2. **City Summary**: Quantity sold by each city for each week

Week ranges are automatically calculated as Monday-Sunday periods in "DD-MMM - DD-MMM" format.

## Technical Details

- Built with Streamlit and pandas
- Processes Excel files in-browser
- No data is stored on servers (privacy-friendly)
- Automatic week calculation and data aggregation
- Error handling for malformed data

## Support

This is a standalone application that requires no technical support once deployed. Users simply upload files and download results. 