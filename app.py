import zipfile
import io
import pandas as pd
import re
from datetime import datetime
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Developer Timesheet Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def detect_headers(df):
    """
    Dynamically detect Date and Hours columns in the dataframe
    Returns: (header_row, date_col, hours_col)
    """
    date_col, hours_col, header_row = None, None, None
    
    # Search through first 10 rows for headers
    for r in range(min(10, len(df))):
        row_has_date = False
        row_has_hours = False
        temp_date_col = None
        temp_hours_col = None
        
        for c in range(len(df.columns)):
            if r >= len(df) or c >= len(df.columns):
                continue
                
            cell_val = str(df.iloc[r, c]).lower().strip()
            
            # Look for date-related headers
            if any(keyword in cell_val for keyword in ['date', 'day', 'datum', 'fecha']):
                temp_date_col = c
                row_has_date = True
            
            # Look for hours-related headers
            if any(keyword in cell_val for keyword in ['hour', 'time', 'work', 'duration', 'hrs', 'horas']):
                temp_hours_col = c
                row_has_hours = True
        
        # If we found both in the same row, that's our header row
        if row_has_date and row_has_hours:
            header_row = r
            date_col = temp_date_col
            hours_col = temp_hours_col
            break
    
    return header_row, date_col, hours_col

def extract_developer_name(filename):
    """
    Extract developer name from filename - takes first word before * or first word before _timesheet
    Examples:
    - prashanth*_reddy_timesheet.xlsx -> Prashanth
    - prashanth_reddy_timesheet.xlsx -> Prashanth (first word only)
    - john_doe_timesheet.xlsx -> John (first word only)
    - mary*smith*timesheet.xlsx -> Mary
    """
    # Remove file extension and path
    base_name = filename.split('/')[-1]  # Remove path if present
    base_name = base_name.replace('.xlsx', '').replace('.xls', '')
    
    # First, check if there's an asterisk - if so, take everything before the first *
    if '*' in base_name:
        name = base_name.split('*')[0]
        return name.strip().title()
    
    # If no asterisk, always take just the first word (before first underscore)
    # This handles cases like prashanth_reddy_timesheet -> prashanth
    first_words = base_name.split('_')[0].strip()
    first_word = first_words.split(' ')[0].strip()
    # Clean up the first word
    if first_word and first_word.isalpha():
        return first_word.title()
    
    # Fallback for edge cases
    return "Unknown"

def clean_and_parse_date(date_val):
    """
    Try to parse various date formats
    """
    if pd.isna(date_val):
        return None
    
    date_str = str(date_val).strip()
    
    # Common date formats to try
    date_formats = [
        '%Y-%m-%d',
        '%m/%d/%Y',
        '%d/%m/%Y',
        '%m-%d-%Y',
        '%d-%m-%Y',
        '%Y/%m/%d',
        '%B %d, %Y',
        '%b %d, %Y',
        '%d %B %Y',
        '%d %b %Y',
    ]
    
    for fmt in date_formats:
        try:
            return pd.to_datetime(date_str, format=fmt)
        except:
            continue
    
    # Try pandas' flexible parsing
    try:
        return pd.to_datetime(date_str, errors='coerce')
    except:
        return None

def extract_data_from_xlsx(file_bytes, file_name):
    """
    Extract timesheet data from XLSX file
    """
    try:
        # Read the Excel file
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')
        
        if df.empty:
            return None, f"Empty file: {file_name}"
        
        # Detect headers
        header_row, date_col, hours_col = detect_headers(df)
        
        if header_row is None or date_col is None or hours_col is None:
            return None, f"Could not detect date/hours columns in: {file_name}"
        
        # Extract data starting from row after header
        data_start_row = header_row + 1
        if data_start_row >= len(df):
            return None, f"No data rows found after header in: {file_name}"
        
        # Get the data columns
        date_data = df.iloc[data_start_row:, date_col]
        hours_data = df.iloc[data_start_row:, hours_col]
        
        # Create dataframe
        data = pd.DataFrame({
            'Date': date_data,
            'Hours': hours_data
        })
        
        # Clean and parse dates
        data['Date'] = data['Date'].apply(clean_and_parse_date)
        
        # Clean and parse hours
        data['Hours'] = pd.to_numeric(data['Hours'], errors='coerce')
        
        # Remove rows where date or hours are invalid
        data = data.dropna(subset=['Date', 'Hours'])
        data = data[data['Hours'] > 0]  # Remove zero or negative hours
        
        if data.empty:
            return None, f"No valid date/hours data found in: {file_name}"
        
        # Extract developer name
        dev_name = extract_developer_name(file_name)
        data['Developer'] = dev_name
        
        # Add month/year for aggregation
        data['YearMonth'] = data['Date'].dt.to_period('M')
        data['Year'] = data['Date'].dt.year
        data['Month'] = data['Date'].dt.month
        data['MonthName'] = data['Date'].dt.month_name()
        
        return data, None
        
    except Exception as e:
        return None, f"Error processing {file_name}: {str(e)}"

def is_valid_excel_file(filename):
    """
    Check if the file is a valid Excel file (not macOS metadata)
    """
    # Skip macOS metadata files
    if filename.startswith('__MACOSX/') or '._' in filename:
        return False
    
    # Skip hidden files and directories
    if filename.startswith('.') or '/.DS_Store' in filename:
        return False
    
    # Must end with xlsx or xls
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return False
    
    # Skip directories
    if filename.endswith('/'):
        return False
    
    return True

def process_zip_file(uploaded_file):
    """
    Process ZIP file containing XLSX timesheets
    """
    all_data = []
    error_messages = []
    processed_files = []
    skipped_files = []
    
    try:
        with zipfile.ZipFile(uploaded_file) as zip_file:
            # Filter for valid Excel files only
            all_files = zip_file.namelist()
            xlsx_files = [f for f in all_files if is_valid_excel_file(f)]
            
            # Track skipped files for user info
            skipped_files = [f for f in all_files if f.endswith(('.xlsx', '.xls')) and not is_valid_excel_file(f)]
            
            if not xlsx_files:
                return pd.DataFrame(), ["No valid XLSX files found in ZIP"], [], skipped_files
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, file_name in enumerate(xlsx_files):
                status_text.text(f"Processing: {file_name}")
                
                try:
                    with zip_file.open(file_name) as xlsx_file:
                        file_bytes = xlsx_file.read()
                        data, error = extract_data_from_xlsx(file_bytes, file_name)
                        
                        if data is not None:
                            all_data.append(data)
                            processed_files.append(file_name)
                        else:
                            error_messages.append(error)
                            
                except Exception as e:
                    error_messages.append(f"Could not read {file_name}: {str(e)}")
                
                progress_bar.progress((i + 1) / len(xlsx_files))
            
            progress_bar.empty()
            status_text.empty()
            
    except Exception as e:
        return pd.DataFrame(), [f"Error reading ZIP file: {str(e)}"], [], []
    
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        return combined_df, error_messages, processed_files, skipped_files
    else:
        return pd.DataFrame(), error_messages, processed_files, skipped_files

def create_monthly_summary(df):
    """
    Create monthly summary of hours worked
    """
    summary = df.groupby(['Developer', 'Year', 'Month', 'MonthName', 'YearMonth']).agg({
        'Hours': 'sum',
        'Date': 'count'  # Number of working days
    }).reset_index()
    
    summary.rename(columns={'Date': 'WorkingDays'}, inplace=True)
    summary['YearMonthStr'] = summary['YearMonth'].dt.strftime('%Y-%m')
    summary['MonthYear'] = summary['MonthName'] + ' ' + summary['Year'].astype(str)
    
    return summary.sort_values(['Year', 'Month', 'Developer'])

def main():
    st.title("📊 Developer Timesheet Analyzer")
    st.markdown("Upload a ZIP file containing XLSX timesheet files to analyze monthly working hours per developer.")
    
    # Sidebar for filters
    st.sidebar.header("📁 File Upload")
    uploaded_file = st.sidebar.file_uploader(
        "Choose ZIP file containing timesheets", 
        type="zip",
        help="Upload a ZIP file containing XLSX timesheet files. Each file should contain date and hours columns."
    )
    
    if uploaded_file is not None:
        st.sidebar.success(f"Uploaded: {uploaded_file.name}")
        
        # Process the file
        with st.spinner("Processing timesheet files..."):
            df, errors, processed_files, skipped_files = process_zip_file(uploaded_file)
        
        # Show processing results
        if processed_files:
            st.success(f"Successfully processed {len(processed_files)} files")
            with st.expander("📄 Processed Files", expanded=False):
                for file in processed_files:
                    st.write(f"✅ {file}")
        
        if skipped_files:
            st.info(f"Skipped {len(skipped_files)} metadata/system files")
            with st.expander("⏭️ Skipped Files", expanded=False):
                for file in skipped_files:
                    st.write(f"⏭️ {file} (macOS metadata or hidden file)")
        
        if errors:
            st.warning(f"Encountered {len(errors)} errors")
            with st.expander("⚠️ Processing Errors", expanded=False):
                for error in errors:
                    st.write(f"❌ {error}")
        
        if not df.empty:
            # Create monthly summary
            monthly_summary = create_monthly_summary(df)
            
            # Sidebar filters
            st.sidebar.header("🎛️ Filters")
            
            # Developer filter
            developers = sorted(df['Developer'].unique())
            selected_developers = st.sidebar.multiselect(
                "Select Developers",
                developers,
                default=developers,
                help="Choose which developers to display"
            )
            
            # Date range filter
            min_date = df['Date'].min()
            max_date = df['Date'].max()
            date_range = st.sidebar.date_input(
                "Date Range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            
            # Apply filters
            if len(date_range) == 2:
                filtered_df = df[
                    (df['Developer'].isin(selected_developers)) &
                    (df['Date'] >= pd.Timestamp(date_range[0])) &
                    (df['Date'] <= pd.Timestamp(date_range[1]))
                ]
                filtered_summary = create_monthly_summary(filtered_df)
            else:
                filtered_df = df[df['Developer'].isin(selected_developers)]
                filtered_summary = monthly_summary[monthly_summary['Developer'].isin(selected_developers)]
            
            # Main content
            if not filtered_df.empty:
                # Key metrics
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    total_hours = filtered_df['Hours'].sum()
                    st.metric("Total Hours", f"{total_hours:,.1f}")
                
                with col2:
                    avg_monthly = filtered_summary['Hours'].mean()
                    st.metric("Avg Monthly Hours", f"{avg_monthly:.1f}")
                
                with col3:
                    active_developers = filtered_df['Developer'].nunique()
                    st.metric("Active Developers", active_developers)
                
                with col4:
                    months_covered = filtered_summary['YearMonth'].nunique()
                    st.metric("Months Covered", months_covered)
                
                # Tabs for different views
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Monthly Summary", "📈 Visualizations", "📅 Detailed Data", "📋 Raw Data"])
                
                with tab1:
                    st.subheader("Monthly Hours Summary")
                    
                    # Pivot table for better display
                    pivot_summary = filtered_summary.pivot_table(
                        index='Developer',
                        columns='MonthYear',
                        values='Hours',
                        fill_value=0
                    )
                    
                    # Calculate monthly totals for each column
                    monthly_totals = pivot_summary.sum()
                    
                    # Add totals row to pivot table
                    pivot_with_totals = pivot_summary.copy()
                    pivot_with_totals.loc['TOTAL HOURS'] = monthly_totals
                    
                    st.dataframe(
                        pivot_with_totals,
                        use_container_width=True,
                        column_config={
                            col: st.column_config.NumberColumn(
                                col,
                                format="%.1f"
                            ) for col in pivot_with_totals.columns
                        }
                    )
                    
                    # Show monthly totals in a more prominent way
                    st.subheader("📊 Monthly Totals Summary")
                    
                    # Create columns for monthly totals display
                    months = list(monthly_totals.index)
                    if len(months) <= 6:
                        cols = st.columns(len(months))
                        for i, (month, total_hours) in enumerate(monthly_totals.items()):
                            with cols[i]:
                                st.metric(
                                    label=month,
                                    value=f"{total_hours:.1f} hrs",
                                    help=f"Total hours worked by all developers in {month}"
                                )
                    else:
                        # If more than 6 months, display in rows of 6
                        for i in range(0, len(months), 6):
                            chunk_months = months[i:i+6]
                            cols = st.columns(len(chunk_months))
                            for j, month in enumerate(chunk_months):
                                total_hours = monthly_totals[month]
                                with cols[j]:
                                    st.metric(
                                        label=month,
                                        value=f"{total_hours:.1f} hrs",
                                        help=f"Total hours worked by all developers in {month}"
                                    )
                    
                    # Monthly totals chart
                    monthly_totals_df = pd.DataFrame({
                        'Month': monthly_totals.index,
                        'Total Hours': monthly_totals.values
                    })
                    
                    fig_totals = px.bar(
                        monthly_totals_df,
                        x='Month',
                        y='Total Hours',
                        title="Total Hours per Month (All Developers)",
                        labels={'Total Hours': 'Total Hours Worked'},
                        color='Total Hours',
                        color_continuous_scale='Blues'
                    )
                    fig_totals.update_layout(
                        xaxis_tickangle=-45,
                        height=400,
                        showlegend=False
                    )
                    # Add value labels on bars
                    fig_totals.update_traces(
                        texttemplate='%{y:.1f}',
                        textposition='outside'
                    )
                    st.plotly_chart(fig_totals, use_container_width=True)
                    
                    # Download button for summary
                    csv = filtered_summary.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Monthly Summary as CSV",
                        data=csv,
                        file_name=f"monthly_timesheet_summary_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
                
                with tab2:
                    st.subheader("Visual Analysis")
                    
                    # Monthly hours bar chart
                    fig_bar = px.bar(
                        filtered_summary,
                        x='MonthYear',
                        y='Hours',
                        color='Developer',
                        title="Monthly Hours by Developer",
                        labels={'Hours': 'Hours Worked', 'MonthYear': 'Month'},
                        barmode='group'
                    )
                    fig_bar.update_layout(xaxis_tickangle=-45, height=500)
                    st.plotly_chart(fig_bar, use_container_width=True)
                    
                    # Line chart showing trends
                    fig_line = px.line(
                        filtered_summary,
                        x='YearMonthStr',
                        y='Hours',
                        color='Developer',
                        title="Monthly Hours Trend",
                        labels={'Hours': 'Hours Worked', 'YearMonthStr': 'Month'},
                        markers=True
                    )
                    fig_line.update_layout(height=500)
                    st.plotly_chart(fig_line, use_container_width=True)
                    
                    # Heatmap
                    if len(developers) > 1 and len(filtered_summary['YearMonthStr'].unique()) > 1:
                        heatmap_data = filtered_summary.pivot_table(
                            index='Developer',
                            columns='YearMonthStr',
                            values='Hours',
                            fill_value=0
                        )
                        
                        fig_heatmap = px.imshow(
                            heatmap_data,
                            title="Hours Heatmap by Developer and Month",
                            color_continuous_scale="Blues",
                            labels={'color': 'Hours'}
                        )
                        fig_heatmap.update_layout(height=400)
                        st.plotly_chart(fig_heatmap, use_container_width=True)
                
                with tab3:
                    st.subheader("Detailed Monthly Data")
                    
                    # Sort by date descending
                    detailed_view = filtered_summary.sort_values(['Year', 'Month'], ascending=[False, False])
                    
                    st.dataframe(
                        detailed_view[['Developer', 'MonthYear', 'Hours', 'WorkingDays']],
                        use_container_width=True,
                        column_config={
                            "Hours": st.column_config.NumberColumn(
                                "Hours Worked",
                                help="Total hours worked in the month",
                                format="%.1f"
                            ),
                            "WorkingDays": st.column_config.NumberColumn(
                                "Working Days",
                                help="Number of days worked in the month"
                            )
                        }
                    )
                
                with tab4:
                    st.subheader("Raw Timesheet Data")
                    
                    # Show individual daily entries
                    raw_display = filtered_df[['Developer', 'Date', 'Hours', 'MonthName', 'Year']].sort_values('Date', ascending=False)
                    
                    st.dataframe(
                        raw_display,
                        use_container_width=True,
                        column_config={
                            "Date": st.column_config.DateColumn(
                                "Date",
                                format="YYYY-MM-DD"
                            ),
                            "Hours": st.column_config.NumberColumn(
                                "Hours",
                                format="%.2f"
                            )
                        }
                    )
                    
                    # Download raw data
                    raw_csv = raw_display.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Raw Data as CSV",
                        data=raw_csv,
                        file_name=f"raw_timesheet_data_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
            
            else:
                st.warning("No data matches the selected filters.")
        
        else:
            st.error("No valid timesheet data could be extracted from the uploaded ZIP file.")
    
    else:
        # Show instructions when no file is uploaded
        st.info("👆 Please upload a ZIP file containing XLSX timesheet files to get started.")
        
        with st.expander("📖 Instructions", expanded=True):
            st.markdown("""
            ### How to use this analyzer:
            
            1. **Prepare your timesheet files:**
               - Each developer should have their own XLSX file
               - Files should be named like: `DeveloperName_timesheet.xlsx`
               - Each file should contain columns for dates and hours worked
            
            2. **Create a ZIP file:**
               - Put all XLSX timesheet files into a ZIP archive
               - Upload the ZIP file using the sidebar
            
            3. **Supported formats:**
               - The system automatically detects date and hours columns
               - Supports various date formats (YYYY-MM-DD, MM/DD/YYYY, etc.)
               - Hours should be numeric values
            
            4. **View results:**
               - Monthly summaries with interactive filtering
               - Visual charts showing trends and comparisons
               - Downloadable CSV reports
            
            ### Example file structure:
            ```
            timesheets.zip
            ├── John_Smith_timesheet.xlsx
            ├── Jane_Doe_timesheet.xlsx
            └── Mike_Johnson_timesheet.xlsx
            ```
            """)

if __name__ == "__main__":
    main()