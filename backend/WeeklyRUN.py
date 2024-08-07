import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import date

# Define directories
WEEKLYDROP_DIR = 'WEEKLYDROP'
WEEKLYCOMPLETE_DIR = 'WEEKLYCOMPLETE'

def process_weekly_csv():
    # Ensure WEEKLYDROP directory exists and is accessible
    if not os.path.exists(WEEKLYDROP_DIR):
        raise Exception(f"The WEEKLYDROP directory '{WEEKLYDROP_DIR}' does not exist.")
    if not os.path.isdir(WEEKLYDROP_DIR):
        raise Exception(f"'{WEEKLYDROP_DIR}' is not a directory.")

    # Ensure WEEKLYCOMPLETE directory exists
    if not os.path.exists(WEEKLYCOMPLETE_DIR):
        os.makedirs(WEEKLYCOMPLETE_DIR)

    # Delete existing files in WEEKLYCOMPLETE directory
    for file in os.listdir(WEEKLYCOMPLETE_DIR):
        file_path = os.path.join(WEEKLYCOMPLETE_DIR, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")

    # Check for files in WEEKLYDROP
    files = [f for f in os.listdir(WEEKLYDROP_DIR) if f.lower().endswith('.csv')]

    # Handle cases based on the number of CSV files
    if len(files) == 0:
        raise Exception("No CSV file found in the WEEKLYDROP directory.")
    elif len(files) > 1:
        raise Exception(f"Multiple CSV files found in the WEEKLYDROP directory: {', '.join(files)}. Please ensure only one CSV file is present.")

    # Get the file name
    file_name = files[0]

    # Read the CSV file
    try:
        df = pd.read_csv(os.path.join(WEEKLYDROP_DIR, file_name))
    except Exception as e:
        raise Exception(f"Error reading the CSV file: {str(e)}")

    # Check if required columns exist
    required_columns = ['Vendor', 'Product', 'Package ID', 'Room', 'Available', 'Expiration date']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise Exception(f"The following required columns are missing: {', '.join(missing_columns)}")

    # Keep only the required columns
    df = df[required_columns]

    # Remove rows where 'Room' is 'Backstock'
    df = df[df['Room'].str.lower() != 'backstock']

    # Remove rows where 'Expiration date' is 0
    df = df[df['Expiration date'] != 0]

    # Remove rows where 'Product' contains 'gear' or 'battery' (case-insensitive)
    df = df[~df['Product'].str.contains('gear|battery', case=False, na=False)]

    # Extract last 4 digits of 'Package ID' and create a new column for sorting
    df['Sort_ID'] = df['Package ID'].astype(str).str[-4:].astype(int)

    # Sort the DataFrame based on the last 4 digits
    df = df.sort_values('Sort_ID')

    # Update 'Package ID' to show only last 4 digits
    df['Package ID'] = df['Package ID'].astype(str).str[-4:]

    # Drop the temporary sorting column
    df = df.drop('Sort_ID', axis=1)

    # Add new columns
    df['Exp Date Conf.'] = ''  # Column G
    df['Fulfillment'] = ''     # Column H
    df['Vault'] = ''           # Column I
    df['Sold'] = ''            # Column J
    df['Total'] = ''           # Column K
    df['✓'] = ''               # Column L with checkmark symbol

    # Reorder columns
    new_order = ['Vendor', 'Product', 'Package ID', 'Room', 'Available', 'Expiration date',
                 'Exp Date Conf.', 'Fulfillment', 'Vault', 'Sold', 'Total', '✓']
    df = df[new_order]

    # Create a new workbook and select the active sheet
    wb = Workbook()
    ws = wb.active

    # Write the DataFrame to the worksheet
    for r_idx, r in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(r)
        for c_idx, cell in enumerate(ws[r_idx], 1):
            # Make all text bold
            cell.font = Font(bold=True)
            # Add thin borders to all cells
            cell.border = Border(left=Side(style='thin'),
                                 right=Side(style='thin'),
                                 top=Side(style='thin'),
                                 bottom=Side(style='thin'))
            # Align columns A and B to the left, others to the center
            if c_idx in [1, 2]:  # Columns A and B
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Set the orientation to landscape
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    # Set the title rows to repeat on each page
    ws.print_title_rows = '1:1'  # This will repeat the first row on each page

    # Add custom header and footer
    current_date = date.today().strftime("%m/%d/%Y")
    ws.oddHeader.left.text = "Discrepancies: _________________"
    ws.oddHeader.center.text = f"Physical Count {current_date}"
    ws.oddHeader.right.text = "Name + Badge: ___________________"
    ws.oddFooter.center.text = "&P"  # Page number

    # Adjust page setup for better fit
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 0 means it will use as many pages as needed vertically

    # Save the workbook
    output_file = os.path.join(WEEKLYCOMPLETE_DIR, 'WeeklyCount_' + os.path.splitext(file_name)[0] + '.xlsx')
    wb.save(output_file)

    print(f"Weekly Count Excel file has been saved to {output_file}")

    # Code to delete the original file from WEEKLYDROP directory
    try:
        os.remove(os.path.join(WEEKLYDROP_DIR, file_name))
        print(f"Original file {file_name} has been deleted from {WEEKLYDROP_DIR}")
    except Exception as e:
        print(f"Error deleting original file {file_name}: {e}")

    return output_file  # Return the path of the processed file

# The __main__ check is removed as this will now be called from app.py
