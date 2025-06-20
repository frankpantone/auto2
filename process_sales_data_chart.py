import pandas as pd
import datetime
import xlsxwriter
from collections import Counter
import numpy as np
import os
import sys

# Display current working directory for troubleshooting
print(f"Current working directory: {os.getcwd()}")
print(f"Python version: {sys.version}")
print(f"Platform: {sys.platform}")

# Check if required file exists
if not os.path.exists('raw.csv'):
    print("❌ Error: 'raw.csv' file not found in the current directory.")
    print(f"Please ensure 'raw.csv' is in: {os.getcwd()}")
    sys.exit(1)

# Read the CSV file with encoding specification for cross-platform compatibility
try:
    # Try UTF-8 first (most common)
    df = pd.read_csv('raw.csv', encoding='utf-8')
except UnicodeDecodeError:
    try:
        # If UTF-8 fails, try Windows-1252 (common on Windows)
        df = pd.read_csv('raw.csv', encoding='windows-1252')
        print("Note: CSV file read using Windows-1252 encoding")
    except UnicodeDecodeError:
        # If both fail, try latin-1 (universal fallback)
        df = pd.read_csv('raw.csv', encoding='latin-1')
        print("Note: CSV file read using latin-1 encoding")
except FileNotFoundError:
    print("Error: raw.csv file not found. Please ensure the file exists in the current directory.")
    sys.exit(1)
except Exception as e:
    print(f"Error reading CSV file: {e}")
    sys.exit(1)

# Get the most common customer name for the filename
customer_name = "Unknown_Customer"
if 'Customer Business Name' in df.columns:
    # Get the most common customer name (excluding empty/null values)
    customer_names = df['Customer Business Name'].dropna()
    customer_names = customer_names[customer_names.str.strip() != '']
    if len(customer_names) > 0:
        most_common_customer = customer_names.value_counts().index[0]
        # Clean the customer name for filename (remove Windows-incompatible characters)
        # Windows forbidden characters: < > : " | ? * / \
        forbidden_chars = '<>:"|?*/' + '\\'
        customer_name = "".join(c for c in most_common_customer if c.isalnum() or c in (' ', '-', '_'))
        customer_name = customer_name.strip().replace(' ', '_')
        # Ensure filename isn't too long (Windows has 260 character path limit)
        if len(customer_name) > 50:
            customer_name = customer_name[:50]

# Create date stamp for filename
date_stamp = datetime.datetime.now().strftime('%Y%m%d')

# Create the output filename with customer name and date stamp
output_filename = f'sales_report_{customer_name}_{date_stamp}.xlsx'
print(f"Creating file: {output_filename}")

# Pre-process specific date columns for better Excel compatibility
date_columns = ['Created Date', 'Accepted by Carrier Date', 'Carrier Pickup Scheduled At', 'Delivery Completed At', 'Adjusted delivery date']
for col in date_columns:
    if col in df.columns:
        # Convert string dates to datetime objects before writing to Excel
        df[col] = pd.to_datetime(df[col], errors='coerce', format='%m/%d/%y')
        # Format back to string for display but now Excel will recognize as dates
        df[col] = df[col].dt.strftime('%m/%d/%Y')

# Check if Vehicle delivery date is missing and populate with Adjusted delivery date where needed
df['Vehicle delivery date'] = df.apply(
    lambda row: row['Adjusted delivery date'] if pd.isna(row['Vehicle delivery date']) or str(row['Vehicle delivery date']).strip() == '' else row['Vehicle delivery date'],
    axis=1
)

# Standardize date format
def standardize_date(date_val):
    if pd.isna(date_val) or str(date_val).strip() == '':
        return 'Unknown'
    
    date_str = str(date_val).strip()
    
    # If it's already in a date format with slashes, parse and reformat consistently
    if '/' in date_str:
        parts = date_str.split('/')
        if len(parts) == 3:
            try:
                month = int(parts[0])
                day = int(parts[1])
                year = parts[2]
                # Handle 2-digit years and convert to 4-digit
                if len(year) == 2:
                    year = '20' + year  # Assuming all dates are in the 2000s
                # Always format as MM/DD/YYYY with consistent leading zeros
                return f"{month:02d}/{day:02d}/{year}"
            except ValueError:
                # If conversion fails, return as is
                pass
    
    # If it doesn't match expected format, just return as is
    return date_str

# Apply standardization
df['Vehicle delivery date'] = df['Vehicle delivery date'].apply(standardize_date)

# Ensure numeric columns are properly converted before writing to Excel
columns_to_convert = {
    'Carrier Price Per Vehicle': 'float',
    'Tariff Per Vehicle': 'float',
    'Total Carrier Price': 'float',
    'Customer Payment Total Tariff': 'float'
}

for col, dtype in columns_to_convert.items():
    if col in df.columns:
        # For numeric columns, ensure they're properly converted
        if dtype == 'float':
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# Handle ZIP codes - extract 5-digit portion and standardize format
for zip_col in ['Pickup ZIP', 'Delivery ZIP']:
    if zip_col in df.columns:
        def standardize_zip_code(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0  # Return 0 for empty values
            
            value_str = str(value).strip()
            
            # Handle ZIP+4 format (e.g., "33172-2749" -> "33172")
            if '-' in value_str:
                zip_part = value_str.split('-')[0]
                if zip_part.isdigit():
                    return int(zip_part.zfill(5))  # Ensure 5 digits with leading zeros if needed
                else:
                    return 0  # Invalid format
            
            # Handle regular numeric ZIP codes
            if value_str.isdigit():
                if len(value_str) <= 5:
                    # Pad 4-digit ZIPs to 5 digits (e.g., "6096" -> "06096")
                    return int(value_str.zfill(5))
                else:
                    # Truncate if longer than 5 digits and take first 5
                    return int(value_str[:5])
            
            # For non-numeric values, return 0 (this handles state abbreviations, etc.)
            return 0
        
        df[zip_col] = df[zip_col].apply(standardize_zip_code)

# For columns containing currency with $ symbol, preprocess them
if 'Total Price per Mile' in df.columns:
    df['Total Price per Mile'] = df['Total Price per Mile'].apply(
        lambda x: float(str(x).replace('$', '').replace(',', '')) if isinstance(x, str) and '$' in x else x
    )
    df['Total Price per Mile'] = pd.to_numeric(df['Total Price per Mile'], errors='coerce').fillna(0)

# Create Excel workbook using xlsxwriter for charting capabilities
try:
    workbook = xlsxwriter.Workbook(output_filename, {'nan_inf_to_errors': True})
except PermissionError:
    print(f"Error: Cannot create {output_filename}. Please close the file if it's currently open.")
    sys.exit(1)
except Exception as e:
    print(f"Error creating Excel workbook: {e}")
    sys.exit(1)

# Create pivot table on a new sheet FIRST
pivot_sheet = workbook.add_worksheet('Pivot Table')

# Create the main data worksheet SECOND
worksheet = workbook.add_worksheet('Sales Data')

# Write the dataframe to the worksheet
# First write headers
for col, header in enumerate(df.columns):
    worksheet.write(0, col, header)

# Apply the same city/state normalization to the dataframe for Sales Data consistency
df_normalized = df.copy()

# Define state mappings for normalization (same as used in analysis)
state_mappings = {
    'ALABAMA': 'AL', 'ALASKA': 'AK', 'ARIZONA': 'AZ', 'ARKANSAS': 'AR', 'CALIFORNIA': 'CA',
    'COLORADO': 'CO', 'CONNECTICUT': 'CT', 'DELAWARE': 'DE', 'FLORIDA': 'FL', 'GEORGIA': 'GA',
    'HAWAII': 'HI', 'IDAHO': 'ID', 'ILLINOIS': 'IL', 'INDIANA': 'IN', 'IOWA': 'IA',
    'KANSAS': 'KS', 'KENTUCKY': 'KY', 'LOUISIANA': 'LA', 'MAINE': 'ME', 'MARYLAND': 'MD',
    'MASSACHUSETTS': 'MA', 'MICHIGAN': 'MI', 'MINNESOTA': 'MN', 'MISSISSIPPI': 'MS',
    'MISSOURI': 'MO', 'MONTANA': 'MT', 'NEBRASKA': 'NE', 'NEVADA': 'NV', 'NEW HAMPSHIRE': 'NH',
    'NEW JERSEY': 'NJ', 'NEW MEXICO': 'NM', 'NEW YORK': 'NY', 'NORTH CAROLINA': 'NC',
    'NORTH DAKOTA': 'ND', 'OHIO': 'OH', 'OKLAHOMA': 'OK', 'OREGON': 'OR', 'PENNSYLVANIA': 'PA',
    'RHODE ISLAND': 'RI', 'SOUTH CAROLINA': 'SC', 'SOUTH DAKOTA': 'SD', 'TENNESSEE': 'TN',
    'TEXAS': 'TX', 'UTAH': 'UT', 'VERMONT': 'VT', 'VIRGINIA': 'VA', 'WASHINGTON': 'WA',
    'WEST VIRGINIA': 'WV', 'WISCONSIN': 'WI', 'WYOMING': 'WY'
}

# Enhanced function to normalize and consolidate city names
def normalize_city_name(city_name):
    if not isinstance(city_name, str) or not city_name.strip():
        return city_name
    
    normalized = city_name.strip().title()
    normalized = ' '.join(normalized.split())  # Remove extra whitespace
    
    # Standard abbreviation normalizations
    normalized = normalized.replace('Saint ', 'St. ').replace('Saint-', 'St. ')
    normalized = normalized.replace(' Of ', ' of ')
    normalized = normalized.replace('Ft. ', 'Fort ').replace('Ft ', 'Fort ')
    normalized = normalized.replace('Mt. ', 'Mount ').replace('Mt ', 'Mount ')
    normalized = normalized.replace(' Beach', ' Bch').replace(' Heights', ' Hts')
    
    # Consolidate common city variations
    # San Francisco area consolidations (handle before general replacements)
    san_francisco_handled = False
    
    if normalized in ['S. San Francisco', 'So. San Francisco']:
        normalized = 'South San Francisco'
        san_francisco_handled = True
    elif normalized in ['N. San Francisco', 'No. San Francisco']:
        normalized = 'North San Francisco'
        san_francisco_handled = True
    elif normalized in ['E. San Francisco']:
        normalized = 'East San Francisco'
        san_francisco_handled = True
    elif normalized in ['W. San Francisco']:
        normalized = 'West San Francisco'
        san_francisco_handled = True
    elif normalized in ['San Francisco', 'South San Francisco', 'North San Francisco', 'East San Francisco', 'West San Francisco']:
        # Already properly named, don't change
        san_francisco_handled = True
    
    # Apply general directional consolidations (but skip if already processed above)
    if not san_francisco_handled:
        # Other common consolidations for non-SF cities
        if normalized.startswith('N. '):
            normalized = normalized.replace('N. ', 'North ', 1)
        elif normalized.startswith('S. '):
            normalized = normalized.replace('S. ', 'South ', 1)
        elif normalized.startswith('E. '):
            normalized = normalized.replace('E. ', 'East ', 1)
        elif normalized.startswith('W. '):
            normalized = normalized.replace('W. ', 'West ', 1)
        elif normalized.startswith('Ne. '):
            normalized = normalized.replace('Ne. ', 'Northeast ', 1)
        elif normalized.startswith('Nw. '):
            normalized = normalized.replace('Nw. ', 'Northwest ', 1)
        elif normalized.startswith('Se. '):
            normalized = normalized.replace('Se. ', 'Southeast ', 1)
        elif normalized.startswith('Sw. '):
            normalized = normalized.replace('Sw. ', 'Southwest ', 1)
    
    # Standardize directional abbreviations at the end (only when they are the last word)
    if normalized.endswith(' N'):
        normalized = normalized[:-2] + ' North'
    elif normalized.endswith(' S'):
        normalized = normalized[:-2] + ' South'
    elif normalized.endswith(' E'):
        normalized = normalized[:-2] + ' East'
    elif normalized.endswith(' W'):
        normalized = normalized[:-2] + ' West'
    elif normalized.endswith(' Ne'):
        normalized = normalized[:-3] + ' Northeast'
    elif normalized.endswith(' Nw'):
        normalized = normalized[:-3] + ' Northwest'
    elif normalized.endswith(' Se'):
        normalized = normalized[:-3] + ' Southeast'
    elif normalized.endswith(' Sw'):
        normalized = normalized[:-3] + ' Southwest'
    
    # Los Angeles area consolidations
    if 'Los Angeles' in normalized and ('Downtown' in normalized or 'Dtla' in normalized.upper()):
        normalized = 'Los Angeles'
    
    # Las Vegas area consolidations - keep North Las Vegas separate as it's a different city
    # (North Las Vegas is actually a separate municipality from Las Vegas)
    
    # Portland consolidations - check for state to differentiate OR/ME
    if normalized == 'Portland':
        normalized = 'Portland'  # Will be differentiated by state
    
    # Remove redundant words
    normalized = normalized.replace('City of ', '').replace('Township of ', '')
    normalized = normalized.replace(' City', '').replace(' Town', '')
    
    return normalized

# Function to normalize state names
def normalize_state_name(state_name):
    if not isinstance(state_name, str) or not state_name.strip():
        return state_name
    
    normalized = state_name.strip().upper()
    return state_mappings.get(normalized, normalized)

# Apply normalization to pickup and delivery cities/states
if 'Pickup City' in df_normalized.columns:
    df_normalized['Pickup City'] = df_normalized['Pickup City'].apply(normalize_city_name)
if 'Pickup State' in df_normalized.columns:
    df_normalized['Pickup State'] = df_normalized['Pickup State'].apply(normalize_state_name)
if 'Delivery City' in df_normalized.columns:
    df_normalized['Delivery City'] = df_normalized['Delivery City'].apply(normalize_city_name)
if 'Delivery State' in df_normalized.columns:
    df_normalized['Delivery State'] = df_normalized['Delivery State'].apply(normalize_state_name)

# Write normalized data
for row_idx, row in df_normalized.iterrows():
    for col_idx, value in enumerate(row):
        # Handle NaN values
        if pd.isna(value):
            worksheet.write(row_idx + 1, col_idx, '')
        else:
            worksheet.write(row_idx + 1, col_idx, value)

# Define formats using xlsxwriter
header_format = workbook.add_format({
    'bold': True,
    'font_size': 12,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#D6EAF8',
    'border': 1
})

date_format = workbook.add_format({
    'num_format': 'mm/dd/yyyy',
    'align': 'center',
    'valign': 'vcenter'
})

currency_format = workbook.add_format({
    'num_format': '$#,##0.00',
    'align': 'center',
    'valign': 'vcenter'
})

currency_format_0 = workbook.add_format({
    'num_format': '$#,##0',
    'align': 'center',
    'valign': 'vcenter'
})

number_format = workbook.add_format({
    'num_format': '0',
    'align': 'center',
    'valign': 'vcenter'
})

zip_format = workbook.add_format({
    'num_format': '00000',  # Number format with leading zeros (5-digit ZIP codes)
    'align': 'center',
    'valign': 'vcenter'
})

# Add a general centered format for columns without special formatting
general_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter'
})

# Apply header formatting with equal column spacing based on longest header
headers_list = list(df.columns)

# Set fixed column width for all columns
equal_column_width = 34  # Fixed width of 34 characters for all columns

max_header_length = max(len(str(header)) for header in headers_list)
print(f"Longest header: {max_header_length} characters")
print(f"Setting all columns to: {equal_column_width} characters wide")

for col_idx, header in enumerate(headers_list):
    # Set equal column width for all columns based on longest header
    worksheet.set_column(col_idx, col_idx, equal_column_width)
    
    # Apply header formatting
    worksheet.write(0, col_idx, header, header_format)

# Set row height for headers
worksheet.set_row(0, 36)

# Apply column formatting based on column names
column_formats = {}
for col_idx, header in enumerate(headers_list):
    if header == 'Created Date':
        column_formats[col_idx] = date_format
    elif header == 'Carrier Price Per Vehicle':
        column_formats[col_idx] = currency_format
    elif header == 'Vehicle delivery date':
        column_formats[col_idx] = date_format
    elif header == 'Tariff Per Vehicle':
        column_formats[col_idx] = currency_format_0
    elif header == 'Total Carrier Price':
        column_formats[col_idx] = currency_format_0
    elif header == 'Total Price per Mile':
        column_formats[col_idx] = currency_format
    elif header == 'Customer Payment Total Tariff':
        column_formats[col_idx] = currency_format_0
    elif header == 'Accepted by Carrier Date':
        column_formats[col_idx] = date_format
    elif header == 'Carrier Pickup Scheduled At':
        column_formats[col_idx] = date_format
    elif header == 'Pickup ZIP':
        column_formats[col_idx] = zip_format
    elif header == 'Delivery Completed At':
        column_formats[col_idx] = date_format
    elif header == 'Adjusted delivery date':
        column_formats[col_idx] = date_format
    elif header == 'Delivery ZIP':
        column_formats[col_idx] = zip_format

# Apply formatting to all data columns while preserving column width
for col_idx in range(len(headers_list)):
    if col_idx in column_formats:
        # Apply special format to specific columns
        worksheet.set_column(col_idx, col_idx, equal_column_width, column_formats[col_idx])
    else:
        # Apply general centered format to columns without special formatting
        worksheet.set_column(col_idx, col_idx, equal_column_width, general_format)

# Freeze panes at A2 (freeze the header row)
worksheet.freeze_panes(1, 0)

# Enable filters for all columns
last_row = len(df)
last_col = len(df.columns) - 1
worksheet.autofilter(0, 0, last_row, last_col)

# Use the normalized dataframe for all subsequent analysis to ensure consistency
df = df_normalized

# Ensure numeric columns are properly converted
df['Tariff Per Vehicle'] = pd.to_numeric(df['Tariff Per Vehicle'], errors='coerce').fillna(0)
df['Total Carrier Price'] = pd.to_numeric(df['Total Carrier Price'], errors='coerce').fillna(0)

# Sort the dates before grouping
def extract_date_components(date_str):
    if pd.isna(date_str) or date_str == 'Unknown':
        return (9999, 12, 31)  # Put Unknown at the end
    
    if '/' in str(date_str):
        parts = str(date_str).split('/')
        if len(parts) == 3:
            try:
                month = int(parts[0])
                day = int(parts[1])
                year = int(parts[2])
                return (year, month, day)
            except ValueError:
                pass
    
    return (9999, 12, 30)

# Apply standardization again to ensure all dates have the same format
df['Vehicle delivery date'] = df['Vehicle delivery date'].apply(standardize_date)

# Add a sorting key column
df['date_sort_key'] = df['Vehicle delivery date'].apply(lambda x: extract_date_components(x))

# Sort the dataframe
df = df.sort_values('date_sort_key')

# Calculate data for pivot table
pivot_data = df.groupby('Vehicle delivery date', as_index=False, dropna=False, observed=True).agg({
    'VIN #': 'count',
    'Tariff Per Vehicle': 'sum',
    'Total Carrier Price': 'sum'
})

# Sort the pivot data by the date components
pivot_data['date_sort_key'] = pivot_data['Vehicle delivery date'].apply(lambda x: extract_date_components(x))
pivot_data = pivot_data.sort_values('date_sort_key')
pivot_data = pivot_data.drop('date_sort_key', axis=1)

print(f"\nPivot table - date ranges:")
print(f"First date: {pivot_data['Vehicle delivery date'].iloc[0]}")
print(f"Last date: {pivot_data['Vehicle delivery date'].iloc[-1]}")
print(f"Number of unique dates in pivot: {len(pivot_data)}")

# Pivot table sheet was already created above as the first tab

# Add pivot table headers
headers = ['Vehicle delivery date', 'VIN #', 'Tariff Per Vehicle', 'Total Carrier Price', 
           'Margin $', 'Margin %', 'Avg Margin Per Unit']

# Set column widths and write headers
for col, header in enumerate(headers):
    pivot_sheet.set_column(col, col, 20)
    pivot_sheet.write(0, col, header, header_format)

# Set header row height
pivot_sheet.set_row(0, 24)

# Define formats for pivot table
pivot_date_format = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'center'})
pivot_currency_format_0 = workbook.add_format({'num_format': '$#,##0', 'align': 'center'})
pivot_currency_format_2 = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center'})
pivot_percent_format = workbook.add_format({'num_format': '0.0%', 'align': 'center'})
pivot_number_format = workbook.add_format({'align': 'center'})

# Initialize variables to calculate grand totals
grand_total_vin = 0
grand_total_tariff = 0
grand_total_carrier_price = 0
grand_total_margin = 0

# Add calculated columns
row_num = 1
for index, row in pivot_data.iterrows():
    # Vehicle delivery date
    pivot_sheet.write(row_num, 0, row['Vehicle delivery date'], pivot_date_format)
    
    # VIN # count
    vin_count = row['VIN #']
    grand_total_vin += vin_count
    pivot_sheet.write(row_num, 1, vin_count, pivot_number_format)
    
    # Tariff Per Vehicle sum
    tariff = row['Tariff Per Vehicle']
    grand_total_tariff += tariff
    pivot_sheet.write(row_num, 2, tariff, pivot_currency_format_2)
    
    # Total Carrier Price sum
    carrier_price = row['Total Carrier Price']
    grand_total_carrier_price += carrier_price
    pivot_sheet.write(row_num, 3, carrier_price, pivot_currency_format_2)
    
    # Calculate Margin $
    margin_dollars = tariff - carrier_price
    grand_total_margin += margin_dollars
    pivot_sheet.write(row_num, 4, margin_dollars, pivot_currency_format_2)
    
    # Calculate Margin %
    margin_percent = margin_dollars / tariff if tariff != 0 else 0
    pivot_sheet.write(row_num, 5, margin_percent, pivot_percent_format)
    
    # Calculate Avg Margin Per Unit
    avg_margin_per_unit = margin_dollars / vin_count if vin_count != 0 else 0
    pivot_sheet.write(row_num, 6, avg_margin_per_unit, pivot_currency_format_2)
    
    row_num += 1

# Add Grand Total row
grand_total_row = row_num
overall_margin_percent = grand_total_margin / grand_total_tariff if grand_total_tariff != 0 else 0
overall_avg_margin_per_unit = grand_total_margin / grand_total_vin if grand_total_vin != 0 else 0

# Define grand total format
grand_total_format = workbook.add_format({
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2
})

grand_total_currency_0 = workbook.add_format({
    'num_format': '$#,##0',
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2
})

grand_total_currency_2 = workbook.add_format({
    'num_format': '$#,##0.00',
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2
})

grand_total_percent = workbook.add_format({
    'num_format': '0.0%',
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2
})

# Add grand total row
pivot_sheet.write(grand_total_row, 0, "Grand Total", grand_total_format)
pivot_sheet.write(grand_total_row, 1, grand_total_vin, grand_total_format)
pivot_sheet.write(grand_total_row, 2, grand_total_tariff, grand_total_currency_2)
pivot_sheet.write(grand_total_row, 3, grand_total_carrier_price, grand_total_currency_2)
pivot_sheet.write(grand_total_row, 4, grand_total_margin, grand_total_currency_2)
pivot_sheet.write(grand_total_row, 5, overall_margin_percent, grand_total_percent)
pivot_sheet.write(grand_total_row, 6, overall_avg_margin_per_unit, grand_total_currency_2)

# Freeze panes on the pivot table headers
pivot_sheet.freeze_panes(1, 0)

# Create Top Cities Sheet
city_sheet = workbook.add_worksheet('Top Cities')

# Create separate dictionaries to track pickup and delivery cities with states and prices
pickup_cities = {}
delivery_cities = {}
pickup_city_total_price = {}  # To track total price for average calculation
delivery_city_total_price = {}
pickup_city_vehicle_count = {}  # To track total vehicles for each pickup city
delivery_city_vehicle_count = {}  # To track total vehicles for each delivery city
pickup_city_carriers = {}  # To track carriers for each pickup city
delivery_city_carriers = {}  # To track carriers for each delivery city

# State mappings already defined above for consistency

# Count occurrences and track prices for each city-state combination
for index, row in df.iterrows():
    pickup_city = row.get('Pickup City', '')
    pickup_state = row.get('Pickup State', '')
    delivery_city = row.get('Delivery City', '')
    delivery_state = row.get('Delivery State', '')
    carrier_price = row.get('Carrier Price Per Vehicle', 0)
    carrier_name = row.get('Carrier Name', '')
    
    # Convert carrier price to numeric, handle any non-numeric values
    try:
        carrier_price = float(carrier_price) if pd.notna(carrier_price) else 0
    except (ValueError, TypeError):
        carrier_price = 0
    
    # Clean carrier name
    carrier_name = str(carrier_name).strip() if pd.notna(carrier_name) else 'Unknown'
    
    # Skip blank or invalid city names
    if isinstance(pickup_city, str) and pickup_city.strip() and isinstance(pickup_state, str) and pickup_state.strip():
        # Use enhanced normalization function for consolidation
        pickup_city = normalize_city_name(pickup_city)
        pickup_state = normalize_state_name(pickup_state)
        
        location_key = f"{pickup_city}, {pickup_state}"
        pickup_cities[location_key] = pickup_cities.get(location_key, 0) + 1
        
        # Track total price and vehicle count for average calculation
        if location_key not in pickup_city_total_price:
            pickup_city_total_price[location_key] = 0
            pickup_city_vehicle_count[location_key] = 0
        pickup_city_total_price[location_key] += carrier_price
        pickup_city_vehicle_count[location_key] += 1
        
        # Track carriers for most used carrier calculation
        if location_key not in pickup_city_carriers:
            pickup_city_carriers[location_key] = {}
        pickup_city_carriers[location_key][carrier_name] = pickup_city_carriers[location_key].get(carrier_name, 0) + 1
    
    if isinstance(delivery_city, str) and delivery_city.strip() and isinstance(delivery_state, str) and delivery_state.strip():
        # Use enhanced normalization function for consolidation
        delivery_city = normalize_city_name(delivery_city)
        delivery_state = normalize_state_name(delivery_state)
        
        location_key = f"{delivery_city}, {delivery_state}"
        delivery_cities[location_key] = delivery_cities.get(location_key, 0) + 1
        
        # Track total price and vehicle count for average calculation
        if location_key not in delivery_city_total_price:
            delivery_city_total_price[location_key] = 0
            delivery_city_vehicle_count[location_key] = 0
        delivery_city_total_price[location_key] += carrier_price
        delivery_city_vehicle_count[location_key] += 1
        
        # Track carriers for most used carrier calculation
        if location_key not in delivery_city_carriers:
            delivery_city_carriers[location_key] = {}
        delivery_city_carriers[location_key][carrier_name] = delivery_city_carriers[location_key].get(carrier_name, 0) + 1

# Sort cities by count (most frequent first)
pickup_city_ranking = sorted(pickup_cities.items(), key=lambda x: x[1], reverse=True)
delivery_city_ranking = sorted(delivery_cities.items(), key=lambda x: x[1], reverse=True)

# Display city statistics for verification
total_pickups_check = sum(count for _, count in pickup_city_ranking)
total_deliveries_check = sum(count for _, count in delivery_city_ranking)

print(f"\n📍 City Analysis Results:")
print(f"   • Total unique pickup locations: {len(pickup_city_ranking)}")
print(f"   • Total unique delivery locations: {len(delivery_city_ranking)}")
print(f"   • Total pickup shipments: {total_pickups_check:,}")
print(f"   • Total delivery shipments: {total_deliveries_check:,}")
print(f"   • Top 5 pickup locations:")
for i, (location, count) in enumerate(pickup_city_ranking[:5], 1):
    percentage = (count / total_pickups_check) * 100 if total_pickups_check > 0 else 0
    print(f"     {i}. {location}: {count} shipments ({percentage:.2f}%) - 1-decimal format")

# Debug: Check for San Francisco variations after consolidation
sf_related = [(loc, count) for loc, count in pickup_city_ranking if 'francisco' in loc.lower()]
if sf_related:
    print(f"   • San Francisco related pickup locations after consolidation:")
    for location, count in sf_related:
        percentage = (count / total_pickups_check) * 100 if total_pickups_check > 0 else 0
        print(f"     - {location}: {count} shipments ({percentage:.2f}%)")

# Debug: Check for other potential duplicates (cities with similar names)
print(f"   • Checking for potential remaining duplicates:")
location_frequency = {}
for loc, count in pickup_city_ranking:
    city_state = loc.split(', ')
    if len(city_state) == 2:
        city, state = city_state
        key = f"{city}_{state}"
        if key in location_frequency:
            location_frequency[key] += count
            print(f"     - DUPLICATE FOUND: {loc} (adding {count} to existing)")
        else:
            location_frequency[key] = count

# Check for base name duplicates (different directional variations)
all_cities = [loc.split(', ')[0] for loc, _ in pickup_city_ranking]
city_base_names = {}
for city in all_cities:
    base = city.replace('North ', '').replace('South ', '').replace('East ', '').replace('West ', '').replace('Northeast ', '').replace('Northwest ', '').replace('Southeast ', '').replace('Southwest ', '')
    if base not in city_base_names:
        city_base_names[base] = []
    city_base_names[base].append(city)

potential_duplicates = {base: cities for base, cities in city_base_names.items() if len(cities) > 1}
if potential_duplicates:
    print(f"   • Cities with directional variations:")
    for base, cities in list(potential_duplicates.items())[:5]:  # Show first 5
        unique_cities = list(set(cities))  # Remove duplicates from the list
        if len(unique_cities) > 1:
            city_counts = []
            for city in unique_cities:
                count = next((count for loc, count in pickup_city_ranking if loc.split(', ')[0] == city), 0)
                if count > 0:
                    city_counts.append((city, count))
            if city_counts:
                total_count = sum(count for _, count in city_counts)
                print(f"     - {base}: {[city for city, _ in city_counts]} (Total: {total_count} shipments)")
else:
    print(f"     - No significant directional duplicates found")


print(f"   • Bottom 5 pickup locations (showing small percentages):")
for i, (location, count) in enumerate(pickup_city_ranking[-5:], start=len(pickup_city_ranking)-4):
    percentage_decimal = (count / total_pickups_check) if total_pickups_check > 0 else 0
    percentage_display = percentage_decimal * 100
    # Check which format would be used
    format_used = "2-decimal" if round(percentage_display, 1) == 0.0 and percentage_decimal > 0 else "1-decimal"
    print(f"     {i}. {location}: {count} shipments ({percentage_display:.3f}%) - {format_used} format")

# Set up column formats for Top Cities
city_header_format = workbook.add_format({
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#D6EAF8'
})

city_data_format = workbook.add_format({
    'font_size': 12,
    'align': 'left'
})

city_number_format = workbook.add_format({
    'font_size': 12,
    'align': 'center'
})

city_percent_format_1 = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'num_format': '0.0%'  # Standard 1 decimal place format
})

city_percent_format_2 = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'num_format': '0.00%'  # 2 decimal places for very small percentages
})

city_currency_format = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'num_format': '$#,##0.00'
})

# Create hyperlink format for count cells
city_hyperlink_format = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'color': 'blue',
    'underline': 1,
    'num_format': '0'  # Ensure it displays as a whole number
})

# Set column widths
city_sheet.set_column(0, 0, 8)   # Rank
city_sheet.set_column(1, 1, 28)  # Pickup Location
city_sheet.set_column(2, 2, 12)  # Count
city_sheet.set_column(3, 3, 12)  # Percentage
city_sheet.set_column(4, 4, 20)  # Avg. Cost per VIN (increased from 15)
city_sheet.set_column(5, 5, 50)  # Most Used Carrier (increased for better visibility)
city_sheet.set_column(6, 6, 4)   # Spacer
city_sheet.set_column(7, 7, 8)   # Rank
city_sheet.set_column(8, 8, 28)  # Delivery Location
city_sheet.set_column(9, 9, 12)  # Count
city_sheet.set_column(10, 10, 12) # Percentage
city_sheet.set_column(11, 11, 20) # Avg. Cost per VIN (increased from 15)
city_sheet.set_column(12, 12, 50) # Most Used Carrier (increased for better visibility)

# Add "Top Pickup Cities" title
city_sheet.merge_range('A1:F1', 'Top Pickup Locations', city_header_format)

# Add "Top Delivery Cities" title
city_sheet.merge_range('H1:M1', 'Top Delivery Locations', city_header_format)

# Add pickup table headers
pickup_headers = ['Rank', 'Pickup Location', 'Count', 'Percentage', 'Avg. Cost Per VIN', 'Most Used Carrier']
for col, header in enumerate(pickup_headers):
    city_sheet.write(1, col, header, city_header_format)

# Add delivery table headers
delivery_headers = ['Rank', 'Delivery Location', 'Count', 'Percentage', 'Avg. Cost Per VIN', 'Most Used Carrier']
for col, header in enumerate(delivery_headers):
    city_sheet.write(1, col + 7, header, city_header_format)

# Calculate total pickups and deliveries for percentage calculation
total_pickups = sum(count for _, count in pickup_city_ranking)
total_deliveries = sum(count for _, count in delivery_city_ranking)

# Create Orders Detail sheet FIRST to populate row maps for hyperlinks
# Pre-calculate row positions for hyperlinks
location_row_map = {}  # Maps location to starting row for hyperlinks
carrier_row_map = {}   # Maps carrier to starting row for hyperlinks

# Create Orders Detail sheet for drill-down functionality
orders_detail_sheet = workbook.add_worksheet('Orders Detail')

# Set up formats for Orders Detail sheet
detail_header_format = workbook.add_format({
    'bold': True,
    'font_size': 12,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#D6EAF8',
    'border': 1
})

detail_section_format = workbook.add_format({
    'bold': True,
    'font_size': 14,
    'align': 'left',
    'bg_color': '#E8F4FD',
    'border': 1
})

detail_data_format = workbook.add_format({
    'font_size': 10,
    'align': 'center',
    'border': 1
})

# Add alternating row format for better readability
detail_data_alt_format = workbook.add_format({
    'font_size': 10,
    'align': 'center',
    'border': 1,
    'bg_color': '#F8F9FA'
})

# Add currency formats for specific columns
detail_currency_format = workbook.add_format({
    'font_size': 10,
    'align': 'center',
    'border': 1,
    'num_format': '$#,##0.00'
})

detail_currency_alt_format = workbook.add_format({
    'font_size': 10,
    'align': 'center',
    'border': 1,
    'bg_color': '#F8F9FA',
    'num_format': '$#,##0.00'
})

# Define key columns for the detail view (data column names)
key_columns = ['VIN #', 'Vehicle Info', 'Pickup City', 'Pickup State', 'Delivery City', 'Delivery State', 
               'Carrier Name', 'Carrier Price Per Vehicle', 'Vehicle delivery date', 
               'Distance', 'Tariff Per Vehicle']

# Define display headers (what shows in the Excel header)
display_headers = ['VIN #', 'Vehicle Info', 'Pickup City', 'Pickup State', 'Delivery City', 'Delivery State', 
                   'Carrier Name', 'Carrier Cost Per Vehicle', 'Vehicle delivery date', 
                   'Distance (Miles)', 'Tariff Per Vehicle']

# Set column widths for Orders Detail sheet (increased for better visibility)
detail_col_widths = [18, 30, 25, 12, 25, 12, 40, 22, 20, 15, 22]
for i, width in enumerate(detail_col_widths):
    orders_detail_sheet.set_column(i, i, width)

# Write main title and navigation info
navigation_format = workbook.add_format({
    'bold': True,
    'font_size': 16,
    'align': 'center',
    'bg_color': '#E8F4FD',
    'border': 1
})

orders_detail_sheet.write(0, 0, "📋 ORDERS DETAIL - DRILL-DOWN DATA", navigation_format)
orders_detail_sheet.merge_range(0, 0, 0, len(display_headers)-1, "📋 ORDERS DETAIL - DRILL-DOWN DATA", navigation_format)

# Add navigation guide
guide_format = workbook.add_format({
    'font_size': 10,
    'align': 'center',
    'italic': True,
    'bg_color': '#F8F9FA',
    'border': 1
})

orders_detail_sheet.write(1, 0, "💡 Use Ctrl+F to quickly find specific locations or carriers", guide_format)
orders_detail_sheet.merge_range(1, 0, 1, len(display_headers)-1, "💡 Use Ctrl+F to quickly find specific locations or carriers", guide_format)

# Write column headers using display headers
for i, header in enumerate(display_headers):
    orders_detail_sheet.write(3, i, header, detail_header_format)

# Set header row height
orders_detail_sheet.set_row(0, 25)
orders_detail_sheet.set_row(1, 20)
orders_detail_sheet.set_row(3, 30)

current_row = 5

# Add pickup cities sections with enhanced formatting
main_section_format = workbook.add_format({
    'bold': True,
    'font_size': 18,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'border': 2,
    'top': 3,
    'bottom': 3
})

# Add extra spacing for better hyperlink positioning
current_row += 2
orders_detail_sheet.write(current_row, 0, "📍 PICKUP LOCATIONS", main_section_format)
orders_detail_sheet.merge_range(current_row, 0, current_row, len(display_headers)-1, "📍 PICKUP LOCATIONS", main_section_format)
orders_detail_sheet.set_row(current_row, 35)
current_row += 3

for location, count in pickup_city_ranking:
    # Add invisible anchor row above section for better hyperlink positioning
    orders_detail_sheet.write(current_row, 0, "", workbook.add_format({'font_size': 1}))
    current_row += 1
    
    # Add section header
    orders_detail_sheet.write(current_row, 0, f"Pickup: {location} ({count} VINs shipped)", detail_section_format)
    orders_detail_sheet.merge_range(current_row, 0, current_row, len(display_headers)-1, f"Pickup: {location} ({count} VINs shipped)", detail_section_format)
    current_row += 1
    
    # Store hyperlink target to the first order row below header
    location_row_map[f"pickup_{location}"] = current_row
    
    # Filter and add orders for this pickup location (data is already normalized)
    pickup_city, pickup_state = location.split(', ')
    location_orders = df[(df['Pickup City'] == pickup_city) & 
                        (df['Pickup State'] == pickup_state)]
    
    for order_idx, (_, order) in enumerate(location_orders.iterrows()):
        # Alternate row colors for better readability
        for i, col in enumerate(key_columns):
            value = order.get(col, 'N/A')
            # Apply currency formatting to columns 7 (Carrier Cost Per Vehicle) and 10 (Tariff Per Vehicle)
            if i in [7, 10]:  # Currency columns
                row_format = detail_currency_alt_format if order_idx % 2 == 0 else detail_currency_format
            else:
                row_format = detail_data_alt_format if order_idx % 2 == 0 else detail_data_format
            orders_detail_sheet.write(current_row, i, value, row_format)
        current_row += 1
    
    current_row += 1  # Add spacing

# Add delivery cities sections with enhanced formatting
# Add extra spacing for better hyperlink positioning
current_row += 2
orders_detail_sheet.write(current_row, 0, "🚚 DELIVERY LOCATIONS", main_section_format)
orders_detail_sheet.merge_range(current_row, 0, current_row, len(display_headers)-1, "🚚 DELIVERY LOCATIONS", main_section_format)
orders_detail_sheet.set_row(current_row, 35)
current_row += 3

for location, count in delivery_city_ranking:
    # Add invisible anchor row above section for better hyperlink positioning
    orders_detail_sheet.write(current_row, 0, "", workbook.add_format({'font_size': 1}))
    current_row += 1
    
    # Add section header
    orders_detail_sheet.write(current_row, 0, f"Delivery: {location} ({count} VINs shipped)", detail_section_format)
    orders_detail_sheet.merge_range(current_row, 0, current_row, len(display_headers)-1, f"Delivery: {location} ({count} VINs shipped)", detail_section_format)
    current_row += 1
    
    # Store hyperlink target to the first order row below header
    location_row_map[f"delivery_{location}"] = current_row
    
    # Filter and add orders for this delivery location (data is already normalized)
    delivery_city, delivery_state = location.split(', ')
    location_orders = df[(df['Delivery City'] == delivery_city) & 
                        (df['Delivery State'] == delivery_state)]
    
    for order_idx, (_, order) in enumerate(location_orders.iterrows()):
        # Alternate row colors for better readability
        for i, col in enumerate(key_columns):
            value = order.get(col, 'N/A')
            # Apply currency formatting to columns 7 (Carrier Cost Per Vehicle) and 10 (Tariff Per Vehicle)
            if i in [7, 10]:  # Currency columns
                row_format = detail_currency_alt_format if order_idx % 2 == 0 else detail_currency_format
            else:
                row_format = detail_data_alt_format if order_idx % 2 == 0 else detail_data_format
            orders_detail_sheet.write(current_row, i, value, row_format)
        current_row += 1
    
    current_row += 1  # Add spacing

# Calculate carrier metrics for later use
# Convert date columns to datetime for delivery time calculation
for date_col in ['Accepted by Carrier Date', 'Delivery Completed At']:
    if date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

# Calculate delivery time in days
df['Delivery Time (Days)'] = np.nan
mask = (~df['Accepted by Carrier Date'].isna()) & (~df['Delivery Completed At'].isna())
df.loc[mask, 'Delivery Time (Days)'] = (df.loc[mask, 'Delivery Completed At'] - df.loc[mask, 'Accepted by Carrier Date']).dt.days

# Ensure Carrier Price Per Vehicle is numeric
df['Carrier Price Per Vehicle'] = pd.to_numeric(df['Carrier Price Per Vehicle'], errors='coerce').fillna(0)

# Ensure Distance is numeric
df['Distance'] = pd.to_numeric(df['Distance'], errors='coerce').fillna(0)

# Calculate Price per Mile for each record
df['Price per Mile'] = np.where(df['Distance'] > 0, df['Carrier Price Per Vehicle'] / df['Distance'], 0)

# Group by carrier
carrier_metrics = df.groupby('Carrier Name').agg({
    'VIN #': 'count',
    'Carrier Price Per Vehicle': 'mean',
    'Delivery Time (Days)': lambda x: x[x >= 0].mean(),  # Only include non-negative delivery times
    'Distance': 'mean',  # Average distance in miles
    'Price per Mile': 'mean'  # Average price per mile
}).reset_index()

# Sort by VIN count (descending)
carrier_metrics = carrier_metrics.sort_values('VIN #', ascending=False).reset_index(drop=True)

# Add carrier sections to Orders Detail with enhanced formatting
# Add extra spacing for better hyperlink positioning
current_row += 2
orders_detail_sheet.write(current_row, 0, "🚛 CARRIERS", main_section_format)
orders_detail_sheet.merge_range(current_row, 0, current_row, len(display_headers)-1, "🚛 CARRIERS", main_section_format)
orders_detail_sheet.set_row(current_row, 35)
current_row += 3

for _, carrier_data in carrier_metrics.iterrows():
    carrier_name = carrier_data['Carrier Name']
    vin_count = carrier_data['VIN #']
    
    # Add multiple invisible anchor rows for better carrier hyperlink positioning
    anchor_start = current_row
    for i in range(3):  # Add 3 anchor rows to position section header higher in view
        orders_detail_sheet.write(current_row, 0, "", workbook.add_format({'font_size': 1}))
        current_row += 1
    
    # Store hyperlink target to the first anchor row for optimal positioning
    carrier_row_map[carrier_name] = anchor_start
    
    # Add section header
    orders_detail_sheet.write(current_row, 0, f"Carrier: {carrier_name} ({vin_count} VINs shipped)", detail_section_format)
    orders_detail_sheet.merge_range(current_row, 0, current_row, len(display_headers)-1, f"Carrier: {carrier_name} ({vin_count} VINs shipped)", detail_section_format)
    current_row += 1
    
    # Filter and add orders for this carrier
    carrier_orders = df[df['Carrier Name'] == carrier_name]
    
    for order_idx, (_, order) in enumerate(carrier_orders.iterrows()):
        # Alternate row colors for better readability
        for i, col in enumerate(key_columns):
            value = order.get(col, 'N/A')
            # Apply currency formatting to columns 7 (Carrier Cost Per Vehicle) and 10 (Tariff Per Vehicle)
            if i in [7, 10]:  # Currency columns
                row_format = detail_currency_alt_format if order_idx % 2 == 0 else detail_currency_format
            else:
                row_format = detail_data_alt_format if order_idx % 2 == 0 else detail_data_format
            orders_detail_sheet.write(current_row, i, value, row_format)
        current_row += 1
    
    current_row += 5  # Add extra spacing for carrier sections at bottom

# Add extra blank rows at the end to ensure proper hyperlink positioning for carriers
for i in range(10):
    orders_detail_sheet.write(current_row, 0, "", workbook.add_format({'font_size': 1}))
    current_row += 1

# Restore frozen panes for better readability
orders_detail_sheet.freeze_panes(4, 0)

# NOW create the Top Cities sheet with correct hyperlink targets
# Add pickup cities data (all locations)
for idx, (location, count) in enumerate(pickup_city_ranking, start=1):
    row_idx = idx + 1  # Start data from row 2
    
    # Rank
    city_sheet.write(row_idx, 0, idx, city_number_format)
    
    # Location (City, State)
    city_sheet.write(row_idx, 1, location, city_data_format)
    
    # Count - with hyperlink using formula approach
    target_row = location_row_map.get(f'pickup_{location}', 1)
    hyperlink_formula = f'=HYPERLINK("#\'Orders Detail\'!A{target_row}",{count})'
    city_sheet.write_formula(row_idx, 2, hyperlink_formula, city_hyperlink_format)
    
    # Percentage - use conditional formatting based on value
    percentage = (count / total_pickups) if total_pickups > 0 else 0
    # Use 2 decimal places if 1 decimal would show as 0.0%, otherwise use 1 decimal
    if round(percentage * 100, 1) == 0.0 and percentage > 0:
        city_sheet.write(row_idx, 3, percentage, city_percent_format_2)
    else:
        city_sheet.write(row_idx, 3, percentage, city_percent_format_1)
    
    # Average Price
    avg_price = pickup_city_total_price[location] / pickup_city_vehicle_count[location] if location in pickup_city_total_price and pickup_city_vehicle_count[location] > 0 else 0
    city_sheet.write(row_idx, 4, avg_price, city_currency_format)
    
    # Most Used Carrier
    if location in pickup_city_carriers and pickup_city_carriers[location]:
        most_used_carrier = max(pickup_city_carriers[location], key=pickup_city_carriers[location].get)
    else:
        most_used_carrier = 'Unknown'
    city_sheet.write(row_idx, 5, most_used_carrier, city_data_format)

# Add delivery cities data (all locations)
for idx, (location, count) in enumerate(delivery_city_ranking, start=1):
    row_idx = idx + 1  # Start data from row 2
    
    # Rank
    city_sheet.write(row_idx, 7, idx, city_number_format)
    
    # Location (City, State)
    city_sheet.write(row_idx, 8, location, city_data_format)
    
    # Count - with hyperlink using formula approach
    target_row = location_row_map.get(f'delivery_{location}', 1)
    hyperlink_formula = f'=HYPERLINK("#\'Orders Detail\'!A{target_row}",{count})'
    city_sheet.write_formula(row_idx, 9, hyperlink_formula, city_hyperlink_format)
    
    # Percentage - use conditional formatting based on value
    percentage = (count / total_deliveries) if total_deliveries > 0 else 0
    # Use 2 decimal places if 1 decimal would show as 0.0%, otherwise use 1 decimal
    if round(percentage * 100, 1) == 0.0 and percentage > 0:
        city_sheet.write(row_idx, 10, percentage, city_percent_format_2)
    else:
        city_sheet.write(row_idx, 10, percentage, city_percent_format_1)
    
    # Average Price
    avg_price = delivery_city_total_price[location] / delivery_city_vehicle_count[location] if location in delivery_city_total_price and delivery_city_vehicle_count[location] > 0 else 0
    city_sheet.write(row_idx, 11, avg_price, city_currency_format)
    
    # Most Used Carrier
    if location in delivery_city_carriers and delivery_city_carriers[location]:
        most_used_carrier = max(delivery_city_carriers[location], key=delivery_city_carriers[location].get)
    else:
        most_used_carrier = 'Unknown'
    city_sheet.write(row_idx, 12, most_used_carrier, city_data_format)

# Freeze the headers
city_sheet.freeze_panes(2, 0)

# Create Carrier Metrics Sheet
carrier_sheet = workbook.add_worksheet('Carrier Metrics')

# Set up formats for Carrier Metrics
carrier_header_format = workbook.add_format({
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#D6EAF8',
    'text_wrap': True
})

carrier_data_format = workbook.add_format({
    'font_size': 12,
    'align': 'left'
})

carrier_number_format = workbook.add_format({
    'font_size': 12,
    'align': 'center'
})

carrier_currency_format = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'num_format': '$#,##0.00'
})

carrier_decimal_format = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'num_format': '0.0'
})

# Create hyperlink format for carrier count cells
carrier_hyperlink_format = workbook.add_format({
    'font_size': 12,
    'align': 'center',
    'color': 'blue',
    'underline': 1,
    'num_format': '0'  # Ensure it displays as a whole number
})

# Set column widths
carrier_sheet.set_column(0, 0, 12)  # Rank
carrier_sheet.set_column(1, 1, 45)  # Carrier Name
carrier_sheet.set_column(2, 2, 25)  # Total VINs Shipped
carrier_sheet.set_column(3, 3, 25)  # Avg. Cost Per VIN
carrier_sheet.set_column(4, 4, 30)  # Avg. Delivery Time (Days)
carrier_sheet.set_column(5, 5, 25)  # Avg. Distance (Miles)
carrier_sheet.set_column(6, 6, 25)  # Avg. Cost per Mile

# Add headers
headers = ['Rank', 'Carrier Name', 'Total VINs Shipped', 'Avg. Cost Per VIN', 'Avg. Delivery Time (Days)', 'Avg. Distance (Miles)', 'Avg. Cost per Mile']
for col, header in enumerate(headers):
    carrier_sheet.write(0, col, header, carrier_header_format)

# Set header row height
carrier_sheet.set_row(0, 45)

# Add data to carrier metrics sheet
for row_idx, (index, data) in enumerate(carrier_metrics.iterrows(), start=1):
    # Rank
    rank = row_idx
    carrier_sheet.write(row_idx, 0, rank, carrier_number_format)
    
    # Carrier Name
    carrier_name = data['Carrier Name']
    carrier_sheet.write(row_idx, 1, carrier_name, carrier_data_format)
    
    # Total VINs Shipped - with hyperlink using formula approach
    vin_count = data['VIN #']
    target_row = carrier_row_map.get(carrier_name, 1)
    hyperlink_formula = f'=HYPERLINK("#\'Orders Detail\'!A{target_row}",{vin_count})'
    carrier_sheet.write_formula(row_idx, 2, hyperlink_formula, carrier_hyperlink_format)
    
    # Avg. Cost Per VIN
    avg_price = data['Carrier Price Per Vehicle']
    carrier_sheet.write(row_idx, 3, avg_price, carrier_currency_format)
    
    # Avg. Delivery Time (Days)
    delivery_time = data['Delivery Time (Days)']
    if pd.isna(delivery_time):
        carrier_sheet.write(row_idx, 4, "N/A", carrier_number_format)
    else:
        carrier_sheet.write(row_idx, 4, delivery_time, carrier_decimal_format)
    
    # Avg. Distance (Miles)
    avg_distance = data['Distance']
    if pd.isna(avg_distance):
        carrier_sheet.write(row_idx, 5, "N/A", carrier_number_format)
    else:
        carrier_sheet.write(row_idx, 5, avg_distance, carrier_decimal_format)
    
    # Avg. Cost per Mile
    price_per_mile = data['Price per Mile']
    if pd.isna(price_per_mile):
        carrier_sheet.write(row_idx, 6, "N/A", carrier_number_format)
    else:
        carrier_sheet.write(row_idx, 6, price_per_mile, carrier_currency_format)

# Freeze the header row
carrier_sheet.freeze_panes(1, 0)

# Add filter to the header row
max_row = len(carrier_metrics)
carrier_sheet.autofilter(0, 0, max_row, 6)

# Add heat map conditional formatting to the "Avg. Cost per Mile" column (column 6)
# Only apply to data rows (row 1 to max_row, excluding header and grand total)
if max_row > 0:
    carrier_sheet.conditional_format(1, 6, max_row, 6, {
        'type': '3_color_scale',
        'min_color': '#63BE7B',  # Green for lowest price per mile (best value)
        'mid_color': '#FFEB84',  # Yellow for medium price per mile
        'max_color': '#F8696B'   # Red for highest price per mile (most expensive)
    })

# Add Grand Total row for Carrier Metrics
grand_total_row = max_row + 1

# Calculate grand totals using weighted averages (weighted by VIN count)
total_carriers = len(carrier_metrics)
total_vins_shipped = carrier_metrics['VIN #'].sum()

# Weighted average for Avg. Cost Per VIN (Column D)
# Formula: Sum(Price × VIN_Count) / Sum(VIN_Count)
price_weighted_sum = (carrier_metrics['Carrier Price Per Vehicle'] * carrier_metrics['VIN #']).sum()
overall_avg_price = price_weighted_sum / total_vins_shipped if total_vins_shipped > 0 else 0

# Weighted average for Avg. Delivery Time (Column E) - excluding NaN values
valid_delivery_mask = ~carrier_metrics['Delivery Time (Days)'].isna()
if valid_delivery_mask.any():
    delivery_time_weighted_sum = (carrier_metrics.loc[valid_delivery_mask, 'Delivery Time (Days)'] * 
                                 carrier_metrics.loc[valid_delivery_mask, 'VIN #']).sum()
    delivery_time_total_vins = carrier_metrics.loc[valid_delivery_mask, 'VIN #'].sum()
    overall_avg_delivery_time = delivery_time_weighted_sum / delivery_time_total_vins if delivery_time_total_vins > 0 else None
else:
    overall_avg_delivery_time = None

# Weighted average for Avg. Distance (Column F) - excluding NaN values
valid_distance_mask = ~carrier_metrics['Distance'].isna()
if valid_distance_mask.any():
    distance_weighted_sum = (carrier_metrics.loc[valid_distance_mask, 'Distance'] * 
                           carrier_metrics.loc[valid_distance_mask, 'VIN #']).sum()
    distance_total_vins = carrier_metrics.loc[valid_distance_mask, 'VIN #'].sum()
    overall_avg_distance = distance_weighted_sum / distance_total_vins if distance_total_vins > 0 else None
else:
    overall_avg_distance = None

# Weighted average for Avg. Cost per Mile - excluding NaN values
valid_price_per_mile_mask = ~carrier_metrics['Price per Mile'].isna()
if valid_price_per_mile_mask.any():
    price_per_mile_weighted_sum = (carrier_metrics.loc[valid_price_per_mile_mask, 'Price per Mile'] * 
                                 carrier_metrics.loc[valid_price_per_mile_mask, 'VIN #']).sum()
    price_per_mile_total_vins = carrier_metrics.loc[valid_price_per_mile_mask, 'VIN #'].sum()
    overall_avg_price_per_mile = price_per_mile_weighted_sum / price_per_mile_total_vins if price_per_mile_total_vins > 0 else None
else:
    overall_avg_price_per_mile = None

# Define grand total formats for carrier metrics
carrier_grand_total_format = workbook.add_format({
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2,
    'left': 1,
    'right': 1
})

carrier_grand_total_currency = workbook.add_format({
    'num_format': '$#,##0.00',
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2,
    'left': 1,
    'right': 1
})

carrier_grand_total_decimal = workbook.add_format({
    'num_format': '0.0',
    'bold': True,
    'font_size': 14,
    'align': 'center',
    'bg_color': '#D6EAF8',
    'top': 2,
    'bottom': 2,
    'left': 1,
    'right': 1
})

# Add grand total row
carrier_sheet.write(grand_total_row, 0, "Grand Total", carrier_grand_total_format)
carrier_sheet.write(grand_total_row, 1, f"{total_carriers} Total Carriers", carrier_grand_total_format)
carrier_sheet.write(grand_total_row, 2, total_vins_shipped, carrier_grand_total_format)
carrier_sheet.write(grand_total_row, 3, overall_avg_price, carrier_grand_total_currency)

if overall_avg_delivery_time is not None:
    carrier_sheet.write(grand_total_row, 4, overall_avg_delivery_time, carrier_grand_total_decimal)
else:
    carrier_sheet.write(grand_total_row, 4, "N/A", carrier_grand_total_format)

if overall_avg_distance is not None:
    carrier_sheet.write(grand_total_row, 5, overall_avg_distance, carrier_grand_total_decimal)
else:
    carrier_sheet.write(grand_total_row, 5, "N/A", carrier_grand_total_format)

if overall_avg_price_per_mile is not None:
    carrier_sheet.write(grand_total_row, 6, overall_avg_price_per_mile, carrier_grand_total_currency)
else:
    carrier_sheet.write(grand_total_row, 6, "N/A", carrier_grand_total_format)



# Create a hidden data sheet for chart data
chart_data_sheet = workbook.add_worksheet('_ChartData')
chart_data_sheet.hide()

# Write chart data headers
chart_data_sheet.write(0, 0, 'Vehicle delivery date')
chart_data_sheet.write(0, 1, 'Tariff Per Vehicle')
chart_data_sheet.write(0, 2, 'Total Carrier Price')
chart_data_sheet.write(0, 3, 'Margin $')
chart_data_sheet.write(0, 4, 'Margin %')
chart_data_sheet.write(0, 5, 'Avg Margin Per Unit')

# Write chart data
for idx, (index, row) in enumerate(pivot_data.iterrows(), start=1):
    chart_data_sheet.write(idx, 0, row['Vehicle delivery date'])
    
    # Calculate all metrics
    tariff = row['Tariff Per Vehicle']
    carrier_price = row['Total Carrier Price']
    vin_count = row['VIN #']
    margin_dollars = tariff - carrier_price
    margin_percent = margin_dollars / tariff if tariff != 0 else 0
    avg_margin_per_unit = margin_dollars / vin_count if vin_count != 0 else 0
    
    chart_data_sheet.write(idx, 1, tariff)
    chart_data_sheet.write(idx, 2, carrier_price)
    chart_data_sheet.write(idx, 3, margin_dollars)
    chart_data_sheet.write(idx, 4, margin_percent)
    chart_data_sheet.write(idx, 5, avg_margin_per_unit)

last_row = len(pivot_data)

# Create Dashboard with Multiple Charts (LAST TAB)
dashboard_sheet = workbook.add_worksheet('Charts Dashboard')

# Chart 1: Tariff Per Vehicle (Top Left)
tariff_chart = workbook.add_chart({'type': 'line'})
tariff_chart.add_series({
    'name': 'Tariff Per Vehicle',
    'categories': ['_ChartData', 1, 0, last_row, 0],
    'values': ['_ChartData', 1, 1, last_row, 1],
    'marker': {'type': 'circle', 'size': 5},
    'line': {'width': 2.5, 'color': '#2E75B6'},
})
tariff_chart.set_title({'name': 'Tariff Per Vehicle Trend', 'name_font': {'size': 16, 'bold': True}})
tariff_chart.set_x_axis({
    'name': 'Delivery Date',
    'num_format': 'mm/dd',
    'name_font': {'size': 12},
    'num_font': {'size': 10, 'rotation': 45},
})
tariff_chart.set_y_axis({
    'name': 'Amount ($)',
    'num_format': '$#,##0',
    'name_font': {'size': 12},
    'num_font': {'size': 10},
})
tariff_chart.set_legend({'position': 'top', 'font': {'size': 11}})
tariff_chart.set_size({'width': 720, 'height': 432})
dashboard_sheet.insert_chart('A2', tariff_chart)

# Chart 2: Total Carrier Price (Top Right)
carrier_chart = workbook.add_chart({'type': 'line'})
carrier_chart.add_series({
    'name': 'Total Carrier Price',
    'categories': ['_ChartData', 1, 0, last_row, 0],
    'values': ['_ChartData', 1, 2, last_row, 2],
    'marker': {'type': 'circle', 'size': 5},
    'line': {'width': 2.5, 'color': '#C65911'},
})
carrier_chart.set_title({'name': 'Total Carrier Price Trend', 'name_font': {'size': 16, 'bold': True}})
carrier_chart.set_x_axis({
    'name': 'Delivery Date',
    'num_format': 'mm/dd',
    'name_font': {'size': 12},
    'num_font': {'size': 10, 'rotation': 45},
})
carrier_chart.set_y_axis({
    'name': 'Amount ($)',
    'num_format': '$#,##0',
    'name_font': {'size': 12},
    'num_font': {'size': 10},
})
carrier_chart.set_legend({'position': 'top', 'font': {'size': 11}})
carrier_chart.set_size({'width': 720, 'height': 432})
dashboard_sheet.insert_chart('L2', carrier_chart)

# Chart 3: Margin $ (Bottom Left)
margin_dollar_chart = workbook.add_chart({'type': 'line'})
margin_dollar_chart.add_series({
    'name': 'Margin $',
    'categories': ['_ChartData', 1, 0, last_row, 0],
    'values': ['_ChartData', 1, 3, last_row, 3],
    'marker': {'type': 'circle', 'size': 5},
    'line': {'width': 2.5, 'color': '#70AD47'},
})
margin_dollar_chart.set_title({'name': 'Margin $ Trend', 'name_font': {'size': 16, 'bold': True}})
margin_dollar_chart.set_x_axis({
    'name': 'Delivery Date',
    'num_format': 'mm/dd',
    'name_font': {'size': 12},
    'num_font': {'size': 10, 'rotation': 45},
})
margin_dollar_chart.set_y_axis({
    'name': 'Amount ($)',
    'num_format': '$#,##0',
    'name_font': {'size': 12},
    'num_font': {'size': 10},
})
margin_dollar_chart.set_legend({'position': 'top', 'font': {'size': 11}})
margin_dollar_chart.set_size({'width': 720, 'height': 432})
dashboard_sheet.insert_chart('A30', margin_dollar_chart)

# Chart 4: Margin % (Bottom Center)
margin_percent_chart = workbook.add_chart({'type': 'line'})
margin_percent_chart.add_series({
    'name': 'Margin %',
    'categories': ['_ChartData', 1, 0, last_row, 0],
    'values': ['_ChartData', 1, 4, last_row, 4],
    'marker': {'type': 'circle', 'size': 5},
    'line': {'width': 2.5, 'color': '#5B9BD5'},
})
margin_percent_chart.set_title({'name': 'Margin % Trend', 'name_font': {'size': 16, 'bold': True}})
margin_percent_chart.set_x_axis({
    'name': 'Delivery Date',
    'num_format': 'mm/dd',
    'name_font': {'size': 12},
    'num_font': {'size': 10, 'rotation': 45},
})
margin_percent_chart.set_y_axis({
    'name': 'Percentage',
    'num_format': '0.0%',
    'name_font': {'size': 12},
    'num_font': {'size': 10},
})
margin_percent_chart.set_legend({'position': 'top', 'font': {'size': 11}})
margin_percent_chart.set_size({'width': 720, 'height': 432})
dashboard_sheet.insert_chart('L30', margin_percent_chart)

# Chart 5: Avg Margin Per Unit (Bottom Center)
avg_margin_chart = workbook.add_chart({'type': 'line'})
avg_margin_chart.add_series({
    'name': 'Avg Margin Per Unit',
    'categories': ['_ChartData', 1, 0, last_row, 0],
    'values': ['_ChartData', 1, 5, last_row, 5],
    'marker': {'type': 'circle', 'size': 5},
    'line': {'width': 2.5, 'color': '#A5A5A5'},
})
avg_margin_chart.set_title({'name': 'Avg Margin Per Unit Trend', 'name_font': {'size': 16, 'bold': True}})
avg_margin_chart.set_x_axis({
    'name': 'Delivery Date',
    'num_format': 'mm/dd',
    'name_font': {'size': 12},
    'num_font': {'size': 10, 'rotation': 45},
})
avg_margin_chart.set_y_axis({
    'name': 'Amount ($)',
    'num_format': '$#,##0.00',
    'name_font': {'size': 12},
    'num_font': {'size': 10},
})
avg_margin_chart.set_legend({'position': 'top', 'font': {'size': 11}})
avg_margin_chart.set_size({'width': 720, 'height': 432})
dashboard_sheet.insert_chart('F58', avg_margin_chart)

# Add dashboard title
dashboard_title_format = workbook.add_format({
    'bold': True,
    'font_size': 18,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#D6EAF8'
})
dashboard_sheet.merge_range('A1:Y1', 'Sales Analytics Dashboard - Pivot Table Trends', dashboard_title_format)
dashboard_sheet.set_row(0, 30)

# Freeze the header row
dashboard_sheet.freeze_panes(1, 0)

# Close the workbook with error handling
try:
    workbook.close()
    print(f"\n✅ Excel file '{output_filename}' created successfully!")
except Exception as e:
    print(f"Warning: Error closing workbook: {e}")
    print(f"File '{output_filename}' may still have been created successfully.")

print("\n📊 Report contents:")
print("  ✓ Sales Data (formatted)")
print("  ✓ Pivot Table (with calculations)")
print("  ✓ Charts Dashboard (5 trend charts)")
print("  ✓ Top Cities Analysis (with clickable drill-down)")
print("  ✓ Carrier Metrics (with weighted averages & clickable drill-down)")
print("  ✓ Orders Detail (comprehensive drill-down data)")
print("\n📈 Charts Dashboard includes:")
print("  • Tariff Per Vehicle Trend")
print("  • Total Carrier Price Trend") 
print("  • Margin $ Trend")
print("  • Margin % Trend")
print("  • Avg Margin Per Unit Trend") 
print("\n🎯 Key Features:")
print("   • Grand totals use weighted averages for accurate metrics")
print("   • Clickable count fields link to detailed order views")
print("   • Complete drill-down functionality for data exploration") 