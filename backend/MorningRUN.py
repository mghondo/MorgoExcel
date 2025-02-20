import os
import openpyxl
from datetime import datetime
from openpyxl.styles import Border, Side, Font, PatternFill
from openpyxl.worksheet.page import PageMargins

# Use the current working directory instead of __file__
lib_folder = os.getcwd()
file_render_folder = os.path.join(lib_folder, 'MORNINGDROP')
file_complete_folder = os.path.join(lib_folder, 'MORNINGCOMPLETE')

# Ensure both folders exist
os.makedirs(file_render_folder, exist_ok=True)
os.makedirs(file_complete_folder, exist_ok=True)

def process_excel(input_path):
    current_date = datetime.now().strftime("%m.%d.%Y")
    output_filename = f"{current_date}.MorningCount.xlsx"
    output_path = os.path.join(file_complete_folder, output_filename)

    try:
        print(f"Input file: {input_path}")
        print(f"Output file: {output_path}")

        # Load the workbook and sheet
        workbook = openpyxl.load_workbook(input_path)
        sheet = workbook.active

        # Delete the first four rows
        sheet.delete_rows(1, 4)

        # Delete columns I and J (9th and 10th columns)
        sheet.delete_cols(9, 2)

        # Cut column H and paste it before column A
        col_h_values = [cell.value for cell in sheet['H']]
        sheet.delete_cols(8)  # Delete column H
        sheet.insert_cols(1)  # Insert a new column at position A
        for i, value in enumerate(col_h_values, start=1):
            sheet.cell(row=i, column=1, value=value)  # Paste values into the new column A

        # Delete the 6th column
        sheet.delete_cols(6)

        # Collect rows and identify rows to delete
        rows_to_keep = []
        for row in sheet.iter_rows(min_row=1, values_only=False):  # Start from the first row
            if row[4].value != "Gear":  # 5th column is at index 4 (0-based index)
                rows_to_keep.append(row)

        # Create a new workbook and sheet
        new_workbook = openpyxl.Workbook()
        new_sheet = new_workbook.active

        # Extract the header row
        header_row = rows_to_keep[0]

        # Repeat the header row at the top of each page
        new_sheet.print_title_rows = '1:1'

        # Exclude the header row from the rows to be sorted
        data_rows = rows_to_keep[1:]

        # Modify the 4th column to retain only the last 4 digits
        for row in data_rows:
            cell = row[3]  # 4th column (index 3)
            if cell.value and isinstance(cell.value, str):
                cell.value = cell.value[-4:]

        # Sort the data rows by column A (index 0) and then by column D (index 3) in a case-insensitive manner
        sorted_data_rows = sorted(data_rows, key=lambda x: (x[0].value.lower(), int(x[3].value)))

        # Append the header row to the new sheet
        new_sheet.append([cell.value for cell in header_row])

        # Append the sorted data rows to the new sheet
        for row in sorted_data_rows:
            new_sheet.append([cell.value for cell in row])

        # Rename the 4th column to "Last 4 #s"
        new_sheet.cell(row=1, column=4, value="Last 4 #s")
        new_sheet.cell(row=1, column=2, value="Product Name")
        new_sheet.cell(row=1, column=8, value="Fulfillment")
        new_sheet.cell(row=1, column=9, value="Vault")
        new_sheet.cell(row=1, column=10, value="Quarantine")
        new_sheet.cell(row=1, column=11, value="Total")
        new_sheet.cell(row=1, column=12, value="\u2713")  # Unicode for checkmark

        # Add custom header and footer
        new_sheet.oddHeader.left.text = "Discrepancies: _____________________________"
        new_sheet.oddHeader.center.text = f"Morning Count {current_date}"
        new_sheet.oddHeader.right.text = "Name + Badge: _____________________________"
        new_sheet.oddFooter.center.text = "&P"  # Page number

        # Set margins and orientation
        new_sheet.page_margins = PageMargins(left=0, right=0)
        new_sheet.page_setup.orientation = 'landscape'

        # Apply borders, bold text, and font size 16 to all cells
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))
        for row in new_sheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.font = Font(bold=True, size=10)

        # Freeze the top row
        new_sheet.freeze_panes = 'A2'

        # Apply a different style to the header row
        for cell in new_sheet[1]:
            cell.font = Font(bold=True, size=12)
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

        # Ensure the sheet is set to print gridlines
        new_sheet.print_gridlines = True

        # Save the modified workbook
        new_workbook.save(output_path)
        print(f"Processed file saved as: {output_filename}")
        print(f"File saved to: {output_path}")

    except Exception as e:
        print(f"An error occurred while processing the Excel file: {str(e)}")
        print(f"Input file path: {input_path}")
        print(f"Current working directory: {os.getcwd()}")
        raise  # Re-raise the exception for debugging

def print_file_preview(file_path, num_lines=10):
    with open(file_path, 'rb') as f:
        print(f"First {num_lines} lines of the file:")
        for i, line in enumerate(f):
            if i >= num_lines:
                break
            print(line)

# Main execution
if __name__ == "__main__":
    try:
        # List all files in the FileRenderFolder
        files = [f for f in os.listdir(file_render_folder) if f.endswith('.xlsx')]
        print(f"Files found in {file_render_folder}: {files}")

        # Check if there is exactly one file
        if len(files) != 1:
            raise Exception(f"There should be exactly one .xlsx file in {file_render_folder}. Found: {len(files)}")

        # Process the file
        input_filename = files[0]
        input_path = os.path.join(file_render_folder, input_filename)

        # Print file preview
        print_file_preview(input_path)

        # Process the Excel file
        process_excel(input_path)

        # Remove the original file
        os.remove(input_path)
        print(f"Original file {input_filename} deleted.")

    except Exception as e:
        print(f"An error occurred in the main execution: {str(e)}")
        print(f"Current working directory: {os.getcwd()}")
        print(f"Contents of {file_render_folder}:")
        print(os.listdir(file_render_folder))

print("Processing complete.")
