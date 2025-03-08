import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import date

# Define directories
WEEKLYDROP_DIR = 'WEEKLYDROP'
WEEKLYCOMPLETE_DIR = 'WEEKLYCOMPLETE'

def process_weekly_file(file_path):
    # Ensure WEEKLYCOMPLETE directory exists
    if not os.path.exists(WEEKLYCOMPLETE_DIR):
        os.makedirs(WEEKLYCOMPLETE_DIR)

    # Delete existing files in WEEKLYCOMPLETE directory
    for file in os.listdir(WEEKLYCOMPLETE_DIR):
        file_path_to_delete = os.path.join(WEEKLYCOMPLETE_DIR, file)
        try:
            if os.path.isfile(file_path_to_delete):
                os.unlink(file_path_to_delete)
        except Exception as e:
            print(f"Error deleting file {file_path_to_delete}: {e}")

    # Get the file name
    file_name = os.path.basename(file_path)

    # Read the CSV file
    try:
        df = pd.read_csv(file_path)
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
    df['Exp Date Conf.'] = '' # Column G
    df['Fulfillment'] = '' # Column H
    df['Vaults'] = '' # Column I
    df['Sold'] = '' # Column J
    df['Total'] = '' # Column K
    df['✓'] = '' # Column L with checkmark symbol

    # Reorder columns
    new_order = ['Vendor', 'Product', 'Package ID', 'Room', 'Available', 'Expiration date',
                 'Exp Date Conf.', 'Fulfillment', 'Vaults', 'Sold', 'Total', '✓']
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
            if c_idx in [1, 2]: # Columns A and B
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Get the company name from the first row of the 'Vendor' column
    company_name = df['Vendor'].iloc[0] if not df.empty else "Unknown Company"

    # Set the orientation to landscape
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    # Set the title rows to repeat on each page
    ws.print_title_rows = '1:1' # This will repeat the first row on each page

    # Add custom header and footer
    current_date = date.today().strftime("%m/%d/%Y")
    ws.oddHeader.left.text = "Discrepancies: _________________"
    ws.oddHeader.center.text = f"{company_name} Physical Count {current_date}"
    ws.oddHeader.right.text = "Name + Badge: ___________________"
    ws.oddFooter.center.text = "&P" # Page number

    # Adjust page setup for better fit
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0 # 0 means it will use as many pages as needed vertically

    # Save the workbook
    output_file = os.path.join(WEEKLYCOMPLETE_DIR, f'{company_name}_{os.path.splitext(file_name)[0]}.xlsx')
    wb.save(output_file)
    print(f"Weekly Count Excel file has been saved to {output_file}")

    # Code to delete the original file
    try:
        os.remove(file_path)
        print(f"Original file {file_name} has been deleted")
    except Exception as e:
        print(f"Error deleting original file {file_name}: {e}")

    return os.path.basename(output_file)  # Return just the filename

    
