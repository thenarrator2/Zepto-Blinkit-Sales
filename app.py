import streamlit as st
import pandas as pd
import io
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

def process_data(zepto_file, blinkit_file, config):
    """Process the uploaded Excel files and return summary data"""
    
    # Read the uploaded files with user-defined sheet names
    zepto = pd.read_excel(zepto_file, sheet_name=config['zepto_sheet'])
    blinkit = pd.read_excel(blinkit_file, sheet_name=config['blinkit_sheet'])
    
    # Add platform identifier
    zepto['platform'] = 'Zepto'
    blinkit['platform'] = 'Blinkit'
    
    # Standardize column names for merging using user-defined mappings
    zepto_clean = zepto.rename(columns={
        config['zepto_sku_col']: "item_name",
        config['zepto_city_col']: "city_name", 
        config['zepto_date_col']: "date",
        config['zepto_qty_col']: "qty"
    })
    
    blinkit_clean = blinkit.rename(columns={
        config['blinkit_sku_col']: "item_name",
        config['blinkit_city_col']: "city_name",
        config['blinkit_date_col']: "date", 
        config['blinkit_qty_col']: "qty"
    })
    
    # Combine both datasets
    data = pd.concat([
        zepto_clean[["item_name", "city_name", "date", "qty", "platform"]],
        blinkit_clean[["item_name", "city_name", "date", "qty", "platform"]]
    ])
    
    # Convert date to datetime and clean data
    data["date"] = pd.to_datetime(data["date"], errors='coerce')
    data = data.dropna(subset=['date', 'item_name', 'city_name'])
    data['qty'] = pd.to_numeric(data['qty'], errors='coerce').fillna(0)
    
    # Create week ranges
    data["week_start"] = data["date"] - pd.to_timedelta(data["date"].dt.dayofweek, unit='d')
    data["week_end"] = data["week_start"] + pd.Timedelta(days=6)
    data["week_range"] = data["week_start"].dt.strftime('%d-%b') + " - " + data["week_end"].dt.strftime('%d-%b')
    
    # Create pivot tables for SKU summary
    sku_summary = pd.pivot_table(
        data,
        index="item_name",
        columns="week_range", 
        values="qty",
        aggfunc="sum",
        fill_value=0
    )
    
    # Create pivot tables for City summary  
    city_summary = pd.pivot_table(
        data,
        index="city_name",
        columns="week_range",
        values="qty", 
        aggfunc="sum",
        fill_value=0
    )
    
    # Create combined SKU + City summary
    data["sku_city"] = data["city_name"] + " - " + data["item_name"]
    sku_city_summary = pd.pivot_table(
        data,
        index="sku_city",
        columns="week_range",
        values="qty",
        aggfunc="sum", 
        fill_value=0
    )
    
    return sku_summary, city_summary, sku_city_summary, data

def create_excel_download(sku_summary, city_summary, sku_city_summary, analytics_data=None):
    """Create an Excel file with all summaries for download"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sku_summary.to_excel(writer, sheet_name='SKU Summary')
        city_summary.to_excel(writer, sheet_name='City Summary')
        sku_city_summary.to_excel(writer, sheet_name='SKU-City Summary')
        
        # Add analytics sheets if available
        if analytics_data:
            if 'low_stock' in analytics_data:
                analytics_data['low_stock'].to_excel(writer, sheet_name='Low Stock Alerts', index=False)
            if 'top_performers' in analytics_data:
                analytics_data['top_performers'].to_excel(writer, sheet_name='Top Performers', index=False)
            if 'trends' in analytics_data:
                analytics_data['trends'].to_excel(writer, sheet_name='Performance Trends', index=False)
            if 'city_performance' in analytics_data:
                analytics_data['city_performance'].to_excel(writer, sheet_name='City Performance', index=False)
            if 'revenue_analysis' in analytics_data:
                analytics_data['revenue_analysis'].to_excel(writer, sheet_name='Revenue Analysis', index=False)
    
    output.seek(0)
    return output

def analyze_low_stock(data, threshold=50):
    """Identify SKUs with low sales in recent weeks"""
    # Group by platform, SKU, and week
    weekly_sales = data.groupby(['platform', 'item_name', 'week_range'])['qty'].sum().reset_index()
    
    # Get the most recent week
    recent_weeks = weekly_sales['week_range'].unique()
    if len(recent_weeks) == 0:
        return pd.DataFrame()
    
    most_recent = sorted(recent_weeks)[-1]
    recent_sales = weekly_sales[weekly_sales['week_range'] == most_recent]
    
    # Find low stock items
    low_stock = recent_sales[recent_sales['qty'] < threshold].copy()
    low_stock['alert_level'] = low_stock['qty'].apply(
        lambda x: 'High' if x < 25 else 'Medium'
    )
    
    return low_stock[['platform', 'item_name', 'qty', 'week_range', 'alert_level']].rename(columns={
        'item_name': 'SKU',
        'qty': 'Last Week Sales',
        'week_range': 'Week'
    })

def analyze_top_performers(data, top_n=10):
    """Identify top performing SKUs across all platforms"""
    # Calculate total sales per SKU per platform
    total_sales = data.groupby(['platform', 'item_name'])['qty'].sum().reset_index()
    total_sales = total_sales.sort_values('qty', ascending=False)
    
    # Add performance level
    total_sales['rank'] = range(1, len(total_sales) + 1)
    total_sales['performance_level'] = total_sales['rank'].apply(
        lambda x: 'Excellent' if x <= 3 else 'Very Good' if x <= 7 else 'Good'
    )
    
    return total_sales.head(top_n).rename(columns={
        'item_name': 'SKU',
        'qty': 'Total Sales'
    })

def analyze_trends(data):
    """Analyze sales trends for each SKU"""
    # Get weekly sales by SKU and platform
    weekly_sales = data.groupby(['platform', 'item_name', 'week_range'])['qty'].sum().reset_index()
    
    trends = []
    for (platform, sku), group in weekly_sales.groupby(['platform', 'item_name']):
        if len(group) >= 2:
            group = group.sort_values('week_range')
            recent_weeks = group.tail(4)  # Last 4 weeks
            
            if len(recent_weeks) >= 2:
                first_week = recent_weeks.iloc[0]['qty']
                last_week = recent_weeks.iloc[-1]['qty']
                
                if first_week == 0 and last_week == 0:
                    trend = 'No Sales'
                elif first_week == 0:
                    trend = 'New Entry'
                else:
                    change = ((last_week - first_week) / first_week) * 100
                    if change > 10:
                        trend = 'Growing'
                    elif change < -10:
                        trend = 'Declining'
                    else:
                        trend = 'Stable'
                
                recommendation = {
                    'Growing': 'Increase Stock',
                    'Declining': 'Review Strategy',
                    'Stable': 'Monitor',
                    'New Entry': 'Track Performance',
                    'No Sales': 'Consider Removal'
                }.get(trend, 'Monitor')
                
                trends.append({
                    'Platform': platform,
                    'SKU': sku,
                    'Trend': trend,
                    'Recent Weeks Sales': ', '.join(map(str, recent_weeks['qty'].tolist())),
                    'Recommendation': recommendation
                })
    
    return pd.DataFrame(trends)

def analyze_city_performance(data):
    """Analyze performance by city"""
    city_stats = data.groupby(['platform', 'city_name']).agg({
        'qty': ['sum', 'mean', 'count']
    }).reset_index()
    
    city_stats.columns = ['Platform', 'City', 'Total Sales', 'Avg Sales', 'Records Count']
    city_stats = city_stats.sort_values('Total Sales', ascending=False)
    
    # Add performance rating
    city_stats['rank'] = range(1, len(city_stats) + 1)
    city_stats['Performance Rating'] = city_stats.apply(
        lambda x: 'Top Performer' if x['rank'] <= 5 
                 else 'Good' if x['rank'] <= 10 
                 else 'Needs Attention' if x['Total Sales'] < 100 
                 else 'Average', axis=1
    )
    
    return city_stats[['Platform', 'City', 'Total Sales', 'Avg Sales', 'Performance Rating']]

def analyze_revenue(data, zepto_price=120, blinkit_price=110):
    """Analyze estimated revenue by platform and SKU"""
    revenue_data = data.groupby(['platform', 'item_name'])['qty'].sum().reset_index()
    
    revenue_data['price_per_unit'] = revenue_data['platform'].map({
        'Zepto': zepto_price,
        'Blinkit': blinkit_price
    })
    
    revenue_data['estimated_revenue'] = revenue_data['qty'] * revenue_data['price_per_unit']
    
    # Calculate revenue share
    total_revenue = revenue_data['estimated_revenue'].sum()
    revenue_data['revenue_share'] = (revenue_data['estimated_revenue'] / total_revenue * 100).round(2)
    
    revenue_data = revenue_data.sort_values('estimated_revenue', ascending=False)
    
    return revenue_data.rename(columns={
        'platform': 'Platform',
        'item_name': 'SKU',
        'qty': 'Total Units Sold',
        'price_per_unit': 'Price Per Unit (₹)',
        'estimated_revenue': 'Estimated Revenue (₹)',
        'revenue_share': 'Revenue Share (%)'
    })

def create_charts(data):
    """Create various charts for the dashboard"""
    charts = {}
    
    # 1. Sales by Platform
    platform_sales = data.groupby('platform')['qty'].sum().reset_index()
    charts['platform_pie'] = px.pie(
        platform_sales, 
        values='qty', 
        names='platform',
        title='Total Sales by Platform',
        color_discrete_sequence=['#1f77b4', '#ff7f0e']
    )
    
    # 2. Weekly Trends
    weekly_trends = data.groupby(['week_range', 'platform'])['qty'].sum().reset_index()
    charts['weekly_trends'] = px.line(
        weekly_trends,
        x='week_range',
        y='qty',
        color='platform',
        title='Weekly Sales Trends by Platform',
        markers=True
    )
    charts['weekly_trends'].update_xaxes(tickangle=45)
    
    # 3. Top Cities
    city_sales = data.groupby('city_name')['qty'].sum().reset_index()
    city_sales = city_sales.sort_values('qty', ascending=True).tail(10)
    charts['top_cities'] = px.bar(
        city_sales,
        x='qty',
        y='city_name',
        orientation='h',
        title='Top 10 Cities by Sales Volume',
        color='qty',
        color_continuous_scale='Blues'
    )
    
    # 4. Top SKUs
    sku_sales = data.groupby('item_name')['qty'].sum().reset_index()
    sku_sales = sku_sales.sort_values('qty', ascending=True).tail(10)
    charts['top_skus'] = px.bar(
        sku_sales,
        x='qty',
        y='item_name',
        orientation='h',
        title='Top 10 SKUs by Sales Volume',
        color='qty',
        color_continuous_scale='Greens'
    )
    
    return charts

# Streamlit App
st.set_page_config(
    page_title="Sales Data Dashboard", 
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales Data Dashboard")
st.markdown("### Upload your Zepto and Blinkit sales data to generate automated summaries")

# Configuration section (collapsible)
with st.expander("⚙️ Configuration (Click to adjust if your Excel format is different)", expanded=False):
    st.markdown("**Adjust these settings if your Excel files have different sheet names or column names:**")
    
    col_config1, col_config2, col_config3 = st.columns(3)
    
    with col_config1:
        st.subheader("Zepto File Settings")
        zepto_sheet = st.text_input("Sheet Name", value="Zepto Sales", key="zepto_sheet")
        zepto_sku_col = st.text_input("SKU Column Name", value="SKU Name", key="zepto_sku")
        zepto_city_col = st.text_input("City Column Name", value="City", key="zepto_city")
        zepto_date_col = st.text_input("Date Column Name", value="Sales Date", key="zepto_date")
        zepto_qty_col = st.text_input("Quantity Column Name", value="Quantity", key="zepto_qty")
    
    with col_config2:
        st.subheader("Blinkit File Settings")
        blinkit_sheet = st.text_input("Sheet Name", value="Blinkit Sales", key="blinkit_sheet")
        blinkit_sku_col = st.text_input("SKU Column Name", value="item_name", key="blinkit_sku")
        blinkit_city_col = st.text_input("City Column Name", value="city_name", key="blinkit_city")
        blinkit_date_col = st.text_input("Date Column Name", value="date", key="blinkit_date")
        blinkit_qty_col = st.text_input("Quantity Column Name", value="qty", key="blinkit_qty")
    
    with col_config3:
        st.subheader("Analytics Settings")
        low_stock_threshold = st.number_input("Low Stock Alert Threshold", value=50, min_value=1, help="Alert if weekly sales below this number")
        top_performers_count = st.number_input("Top Performers Count", value=10, min_value=1, max_value=50, help="Number of top performers to show")
        zepto_price = st.number_input("Zepto Price per Unit (₹)", value=120, min_value=1, help="Estimated price per unit for revenue calculation")
        blinkit_price = st.number_input("Blinkit Price per Unit (₹)", value=110, min_value=1, help="Estimated price per unit for revenue calculation")
        low_sales_alert = st.number_input("High Priority Alert Threshold", value=25, min_value=1, help="High priority alert if sales below this")
    
    st.info("💡 **Tip**: If your Excel files have different column names, just update them above and the app will automatically adapt!")

# Store configuration
config = {
    'zepto_sheet': zepto_sheet,
    'zepto_sku_col': zepto_sku_col,
    'zepto_city_col': zepto_city_col,
    'zepto_date_col': zepto_date_col,
    'zepto_qty_col': zepto_qty_col,
    'blinkit_sheet': blinkit_sheet,
    'blinkit_sku_col': blinkit_sku_col,
    'blinkit_city_col': blinkit_city_col,
    'blinkit_date_col': blinkit_date_col,
    'blinkit_qty_col': blinkit_qty_col,
    'low_stock_threshold': low_stock_threshold,
    'top_performers_count': top_performers_count,
    'zepto_price': zepto_price,
    'blinkit_price': blinkit_price,
    'low_sales_alert': low_sales_alert
}

st.markdown("---")

# Create two columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Upload Zepto Sales Data")
    zepto_file = st.file_uploader(
        "Choose Zepto Excel file",
        type=['xlsx', 'xls'],
        key="zepto"
    )
    
with col2:
    st.subheader("📁 Upload Blinkit Sales Data") 
    blinkit_file = st.file_uploader(
        "Choose Blinkit Excel file",
        type=['xlsx', 'xls'],
        key="blinkit"
    )

# Process files when both are uploaded
if zepto_file is not None and blinkit_file is not None:
    try:
        with st.spinner("Processing your data..."):
            sku_summary, city_summary, sku_city_summary, raw_data = process_data(zepto_file, blinkit_file, config)
        
        st.success("✅ Data processed successfully!")
        
        # Display summary statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total SKUs", len(sku_summary))
        with col2:
            st.metric("Total Cities", len(city_summary))
        with col3:
            st.metric("SKU-City Combinations", len(sku_city_summary))
        with col4:
            st.metric("Total Records", len(raw_data))
        
        # Create main tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Overview & Charts", 
            "📈 Basic Reports", 
            "⚠️ Analytics & Alerts",
            "💰 Revenue Analysis", 
            "📥 Download Data"
        ])
        
        # Generate analytics data
        with st.spinner("Generating analytics..."):
            low_stock_data = analyze_low_stock(raw_data, config['low_stock_threshold'])
            top_performers_data = analyze_top_performers(raw_data, config['top_performers_count'])
            trends_data = analyze_trends(raw_data)
            city_performance_data = analyze_city_performance(raw_data)
            revenue_data = analyze_revenue(raw_data, config['zepto_price'], config['blinkit_price'])
            charts = create_charts(raw_data)
        
        # Tab 1: Overview & Charts
        with tab1:
            st.subheader("📊 Sales Overview")
            
            # Key metrics row
            total_revenue = revenue_data['Estimated Revenue (₹)'].sum()
            avg_weekly_sales = raw_data.groupby('week_range')['qty'].sum().mean()
            
            metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
            with metric_col1:
                st.metric("💰 Total Estimated Revenue", f"₹{total_revenue:,.0f}")
            with metric_col2:
                st.metric("📈 Avg Weekly Sales", f"{avg_weekly_sales:.0f} units")
            with metric_col3:
                zepto_sales = raw_data[raw_data['platform'] == 'Zepto']['qty'].sum()
                st.metric("🟢 Zepto Total Sales", f"{zepto_sales:,.0f} units")
            with metric_col4:
                blinkit_sales = raw_data[raw_data['platform'] == 'Blinkit']['qty'].sum()
                st.metric("🔵 Blinkit Total Sales", f"{blinkit_sales:,.0f} units")
            
            # Charts
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                st.plotly_chart(charts['platform_pie'], use_container_width=True)
                st.plotly_chart(charts['top_cities'], use_container_width=True)
            
            with chart_col2:
                st.plotly_chart(charts['weekly_trends'], use_container_width=True)
                st.plotly_chart(charts['top_skus'], use_container_width=True)
        
        # Tab 2: Basic Reports
        with tab2:
            sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📈 SKU Summary", "🏙️ City Summary", "🔗 SKU-City Summary"])
            
            with sub_tab1:
                st.subheader("Quantity Sold by SKU for Each Week")
                st.dataframe(sku_summary, use_container_width=True)
            
            with sub_tab2:
                st.subheader("Quantity Sold by City for Each Week")
                st.dataframe(city_summary, use_container_width=True)
                
            with sub_tab3:
                st.subheader("Quantity Sold by City-SKU Combination for Each Week")
                st.dataframe(sku_city_summary, use_container_width=True)
        
        # Tab 3: Analytics & Alerts
        with tab3:
            alert_col1, alert_col2 = st.columns(2)
            
            with alert_col1:
                st.subheader("⚠️ Low Stock Alerts")
                if len(low_stock_data) > 0:
                    st.error(f"🚨 Found {len(low_stock_data)} SKUs with low sales!")
                    
                    # Color code the dataframe
                    def color_alerts(val):
                        if val == 'High':
                            return 'background-color: #ffebee'
                        elif val == 'Medium':
                            return 'background-color: #fff3e0'
                        return ''
                    
                    styled_alerts = low_stock_data.style.applymap(color_alerts, subset=['alert_level'])
                    st.dataframe(styled_alerts, use_container_width=True)
                else:
                    st.success("✅ No low stock alerts - all SKUs performing well!")
                
                st.subheader("🏆 Top Performers")
                def color_performance(val):
                    if val == 'Excellent':
                        return 'background-color: #e8f5e8'
                    elif val == 'Very Good':
                        return 'background-color: #f3f8ff'
                    return ''
                
                styled_performers = top_performers_data.style.applymap(color_performance, subset=['performance_level'])
                st.dataframe(styled_performers, use_container_width=True)
            
            with alert_col2:
                st.subheader("📊 Performance Trends")
                if len(trends_data) > 0:
                    def color_trends(val):
                        if val == 'Growing':
                            return 'background-color: #c8e6c9'
                        elif val == 'Declining':
                            return 'background-color: #ffcdd2'
                        elif val == 'New Entry':
                            return 'background-color: #e1f5fe'
                        return ''
                    
                    styled_trends = trends_data.style.applymap(color_trends, subset=['Trend'])
                    st.dataframe(styled_trends, use_container_width=True)
                else:
                    st.info("Need at least 2 weeks of data for trend analysis")
                
                st.subheader("📍 City Performance")
                def color_city_performance(val):
                    if val == 'Top Performer':
                        return 'background-color: #e8f5e8'
                    elif val == 'Needs Attention':
                        return 'background-color: #ffebee'
                    return ''
                
                styled_cities = city_performance_data.style.applymap(color_city_performance, subset=['Performance Rating'])
                st.dataframe(styled_cities, use_container_width=True)
        
        # Tab 4: Revenue Analysis
        with tab4:
            st.subheader("💰 Revenue Analysis")
            
            # Revenue metrics
            revenue_col1, revenue_col2, revenue_col3 = st.columns(3)
            with revenue_col1:
                zepto_revenue = revenue_data[revenue_data['Platform'] == 'Zepto']['Estimated Revenue (₹)'].sum()
                st.metric("🟢 Zepto Revenue", f"₹{zepto_revenue:,.0f}")
            with revenue_col2:
                blinkit_revenue = revenue_data[revenue_data['Platform'] == 'Blinkit']['Estimated Revenue (₹)'].sum()
                st.metric("🔵 Blinkit Revenue", f"₹{blinkit_revenue:,.0f}")
            with revenue_col3:
                avg_revenue_per_sku = revenue_data['Estimated Revenue (₹)'].mean()
                st.metric("📊 Avg Revenue/SKU", f"₹{avg_revenue_per_sku:,.0f}")
            
            # Revenue chart
            revenue_chart = px.bar(
                revenue_data.head(15),
                x='Estimated Revenue (₹)',
                y='SKU',
                color='Platform',
                orientation='h',
                title='Top 15 SKUs by Estimated Revenue',
                color_discrete_sequence=['#1f77b4', '#ff7f0e']
            )
            st.plotly_chart(revenue_chart, use_container_width=True)
            
            # Revenue table
            st.subheader("📋 Detailed Revenue Breakdown")
            st.dataframe(revenue_data, use_container_width=True)
        
        # Tab 5: Download Data
        with tab5:
            st.subheader("📥 Download Options")
            
            # Prepare analytics data for download
            analytics_data = {
                'low_stock': low_stock_data,
                'top_performers': top_performers_data,
                'trends': trends_data,
                'city_performance': city_performance_data,
                'revenue_analysis': revenue_data
            }
            
            # Create comprehensive Excel file
            excel_file = create_excel_download(sku_summary, city_summary, sku_city_summary, analytics_data)
            
            col_download1, col_download2 = st.columns(2)
            with col_download1:
                st.download_button(
                    label="📥 Download Complete Report (Excel)",
                    data=excel_file,
                    file_name=f"sales_complete_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Downloads Excel file with all summaries and analytics"
                )
                
                st.info("""
                **Complete Report includes:**
                - SKU Summary
                - City Summary  
                - SKU-City Summary
                - Low Stock Alerts
                - Top Performers
                - Performance Trends
                - City Performance
                - Revenue Analysis
                """)
            
            with col_download2:
                # Individual CSV downloads
                st.markdown("**Individual Reports:**")
                
                if len(low_stock_data) > 0:
                    csv_low_stock = low_stock_data.to_csv(index=False)
                    st.download_button(
                        "⚠️ Low Stock Alerts (CSV)", 
                        csv_low_stock, 
                        f"low_stock_alerts_{datetime.now().strftime('%Y%m%d')}.csv"
                    )
                
                csv_top_performers = top_performers_data.to_csv(index=False)
                st.download_button(
                    "🏆 Top Performers (CSV)", 
                    csv_top_performers, 
                    f"top_performers_{datetime.now().strftime('%Y%m%d')}.csv"
                )
                
                csv_revenue = revenue_data.to_csv(index=False)
                st.download_button(
                    "💰 Revenue Analysis (CSV)", 
                    csv_revenue, 
                    f"revenue_analysis_{datetime.now().strftime('%Y%m%d')}.csv"
                )
        
    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        
        # Show available sheets and columns to help user configure
        st.info("**Debug Information:**")
        try:
            zepto_sheets = pd.ExcelFile(zepto_file).sheet_names
            blinkit_sheets = pd.ExcelFile(blinkit_file).sheet_names
            
            col_debug1, col_debug2 = st.columns(2)
            with col_debug1:
                st.write("**Zepto file sheets:**", zepto_sheets)
                if len(zepto_sheets) > 0:
                    sample_zepto = pd.read_excel(zepto_file, sheet_name=zepto_sheets[0], nrows=0)
                    st.write("**Zepto columns:**", list(sample_zepto.columns))
            
            with col_debug2:
                st.write("**Blinkit file sheets:**", blinkit_sheets)
                if len(blinkit_sheets) > 0:
                    sample_blinkit = pd.read_excel(blinkit_file, sheet_name=blinkit_sheets[0], nrows=0)
                    st.write("**Blinkit columns:**", list(sample_blinkit.columns))
                    
            st.info("💡 Use the configuration section above to match your file structure!")
        except:
            st.info("Please check your file format and configuration settings above.")

else:
    st.info("👆 Please upload both Excel files to get started")

# Instructions
with st.expander("📋 Instructions"):
    st.markdown("""
    **How to use this dashboard:**
    
         1. **Upload Files**: Upload your Zepto and Blinkit Excel files
     
     2. **Configure (if needed)**: If your files have different sheet names or column names, click the "⚙️ Configuration" section above to adjust settings
     
     3. **Process**: The app will automatically process and display:
        - SKU Summary: Quantity sold by SKU for each week
        - City Summary: Quantity sold by City for each week
        - SKU-City Summary: Quantity sold by City-SKU combination for each week
     
     4. **Download**: Get an Excel file with all 3 summaries
     
     **Default Expected Format:**
     - **Zepto**: Sheet "Zepto Sales" with columns: SKU Name, City, Sales Date, Quantity
     - **Blinkit**: Sheet "Blinkit Sales" with columns: item_name, city_name, date, qty
     
     **Note**: The app automatically groups sales data by weeks (Monday to Sunday) and creates date ranges in "DD-MMM - DD-MMM" format.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit and pandas for automated sales reporting*") 