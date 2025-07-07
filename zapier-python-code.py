# Zapier Python Code for Sales Data Processing
# This code runs in "Code by Zapier" action

import pandas as pd
import io
import base64
from datetime import datetime, timedelta
import json

def main(input_data):
    """
    Main function that Zapier calls
    Input: Email attachment or Google Drive file content
    Output: Processed data ready for Google Sheets
    """
    
    try:
        # Get file content and metadata from Zapier
        file_content = input_data.get('file_content')  # Base64 encoded file
        filename = input_data.get('filename', '').lower()
        
        # Decode base64 file content
        if isinstance(file_content, str):
            file_content = base64.b64decode(file_content)
        
        # Determine platform from filename or subject
        platform = determine_platform(filename, input_data.get('subject', ''))
        
        # Process the Excel file
        summaries = process_excel_file(file_content, platform)
        
        # Format data for Google Sheets update
        sheets_data = format_for_google_sheets(summaries, platform)
        
        return {
            'success': True,
            'platform': platform,
            'processed_at': datetime.now().isoformat(),
            'sheets_data': sheets_data,
            'summary_stats': {
                'total_skus': len(summaries['sku_data']),
                'total_cities': len(summaries['city_data']),
                'weeks_covered': len(summaries['weeks'])
            }
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'processed_at': datetime.now().isoformat()
        }

def determine_platform(filename, subject):
    """Determine if file is Zepto or Blinkit based on filename/subject"""
    text = f"{filename} {subject}".lower()
    
    if 'zepto' in text:
        return 'zepto'
    elif 'blinkit' in text:
        return 'blinkit'
    else:
        # Default fallback - could also throw error
        return 'zepto'

def process_excel_file(file_content, platform):
    """Process Excel file and create summaries"""
    
    # Configuration for different platforms
    config = {
        'zepto': {
            'sheet_name': 'Zepto Sales',
            'columns': {
                'sku': 'SKU Name',
                'city': 'City',
                'date': 'Sales Date', 
                'qty': 'Quantity'
            }
        },
        'blinkit': {
            'sheet_name': 'Blinkit Sales',
            'columns': {
                'sku': 'item_name',
                'city': 'city_name',
                'date': 'date',
                'qty': 'qty'
            }
        }
    }
    
    platform_config = config[platform]
    
    # Read Excel file
    try:
        # Try to read with specified sheet name first
        df = pd.read_excel(
            io.BytesIO(file_content), 
            sheet_name=platform_config['sheet_name']
        )
    except:
        # Fallback: read first sheet if named sheet doesn't exist
        df = pd.read_excel(io.BytesIO(file_content), sheet_name=0)
    
    # Standardize column names
    column_mapping = platform_config['columns']
    df = df.rename(columns=column_mapping)
    
    # Clean and validate data
    required_cols = ['sku', 'city', 'date', 'qty']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")
    
    # Remove rows with missing critical data
    df = df.dropna(subset=required_cols)
    
    # Convert data types
    df['date'] = pd.to_datetime(df['date'])
    df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
    
    # Create week ranges
    df['week_start'] = df['date'] - pd.to_timedelta(df['date'].dt.dayofweek, unit='d')
    df['week_end'] = df['week_start'] + pd.Timedelta(days=6)
    df['week_range'] = df['week_start'].dt.strftime('%d-%b') + ' - ' + df['week_end'].dt.strftime('%d-%b')
    
    # Get unique weeks and sort them
    weeks = sorted(df['week_range'].unique())
    
    # Create SKU summary
    sku_pivot = df.pivot_table(
        index='sku',
        columns='week_range', 
        values='qty',
        aggfunc='sum',
        fill_value=0
    )
    
    # Create City summary
    city_pivot = df.pivot_table(
        index='city',
        columns='week_range',
        values='qty', 
        aggfunc='sum',
        fill_value=0
    )
    
    # Convert to dictionaries for easier handling
    sku_data = {}
    for sku in sku_pivot.index:
        sku_data[sku] = {week: sku_pivot.loc[sku, week] for week in weeks}
    
    city_data = {}
    for city in city_pivot.index:
        city_data[city] = {week: city_pivot.loc[city, week] for week in weeks}
    
    return {
        'weeks': weeks,
        'sku_data': sku_data,
        'city_data': city_data,
        'raw_data_count': len(df)
    }

def format_for_google_sheets(summaries, platform):
    """Format processed data for Google Sheets update"""
    
    weeks = summaries['weeks']
    sku_data = summaries['sku_data'] 
    city_data = summaries['city_data']
    
    # Template configuration (adjust these based on your template)
    template_config = {
        'week_header_row': 2,
        'week_start_col': 3,  # Column C
        'sku_summary_start_row': 3,
        'sku_summary_start_col': 2,  # Column B  
        'city_summary_start_row': 25,
        'city_summary_start_col': 2,  # Column B
        'clear_rows': 50,  # How many rows to clear
        'clear_cols': 20   # How many columns to clear
    }
    
    # Prepare data for Google Sheets batch update
    updates = []
    
    # 1. Clear existing data areas first
    # Week headers
    updates.append({
        'range': f'Sheet1!C{template_config["week_header_row"]}:T{template_config["week_header_row"]}',
        'values': [['']*18]  # Clear week header row
    })
    
    # Clear SKU area
    for i in range(template_config['clear_rows']):
        row = template_config['sku_summary_start_row'] + i
        updates.append({
            'range': f'Sheet1!B{row}:T{row}',
            'values': [['']*19]
        })
    
    # Clear City area  
    for i in range(template_config['clear_rows']):
        row = template_config['city_summary_start_row'] + i
        updates.append({
            'range': f'Sheet1!B{row}:T{row}',
            'values': [['']*19]
        })
    
    # 2. Add week headers
    if weeks:
        week_row = template_config['week_header_row']
        week_values = [weeks + ['']*(18-len(weeks))]  # Pad to 18 columns
        updates.append({
            'range': f'Sheet1!C{week_row}:T{week_row}',
            'values': week_values
        })
    
    # 3. Add SKU summary data
    sku_items = sorted(sku_data.keys())
    for i, sku in enumerate(sku_items):
        row = template_config['sku_summary_start_row'] + i
        
        # SKU name in column B
        updates.append({
            'range': f'Sheet1!B{row}',
            'values': [[sku]]
        })
        
        # Weekly quantities starting from column C
        quantities = [sku_data[sku].get(week, 0) for week in weeks]
        quantities += [0]*(18-len(quantities))  # Pad to 18 columns
        updates.append({
            'range': f'Sheet1!C{row}:T{row}',
            'values': [quantities]
        })
    
    # 4. Add City summary data
    city_items = sorted(city_data.keys())
    for i, city in enumerate(city_items):
        row = template_config['city_summary_start_row'] + i
        
        # City name in column B
        updates.append({
            'range': f'Sheet1!B{row}',
            'values': [[city]]
        })
        
        # Weekly quantities starting from column C
        quantities = [city_data[city].get(week, 0) for week in weeks]
        quantities += [0]*(18-len(quantities))  # Pad to 18 columns
        updates.append({
            'range': f'Sheet1!C{row}:T{row}',
            'values': [quantities]
        })
    
    return {
        'updates': updates,
        'platform': platform,
        'weeks_count': len(weeks),
        'sku_count': len(sku_items),
        'city_count': len(city_items)
    }

# For testing purposes (not used in Zapier)
if __name__ == "__main__":
    # Test data structure
    test_input = {
        'filename': 'zepto_sales_data.xlsx',
        'subject': 'Weekly Zepto Sales Data',
        'file_content': 'base64_encoded_excel_file_content_here'
    }
    
    result = main(test_input)
    print(json.dumps(result, indent=2, default=str)) 