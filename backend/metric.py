import os
import csv
import openpyxl
from openpyxl.utils import get_column_letter
import datetime

def process_metric_file(input_file, output_file):
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    # Read the CSV file
    with open(input_file, 'r', newline='') as csvfile:
        csv_reader = csv.reader(csvfile)
        for row in csv_reader:
            ws.append(row)

    # Move data from column S to T
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=19, max_col=19):
        for cell in row:
            ws.cell(row=cell.row, column=20, value=cell.value)

    # Delete columns S, P to K, and H to C
    ws.delete_cols(19)  # Delete column S
    ws.delete_cols(11, 6)  # Delete columns K to P
    ws.delete_cols(3, 6)  # Delete columns C to H

    if ws['B1'].value == "Item":
        ws['B1'].value = "Description"
        ws['B1'].alignment = openpyxl.styles.Alignment(wrap_text=True)  # Enable text wrapping

    # Change the title of column C from "Quantity" to "METRIC\nQuantity"
    if ws['C1'].value == "Quantity":
        ws['C1'].value = "METRIC\nQuantity"
        ws['C1'].alignment = openpyxl.styles.Alignment(wrap_text=True)  # Enable text wrapping

    # Delete the last three columns (E, F, G)
    ws.delete_cols(ws.max_column - 2, 3)

    # Add new titles to cells starting with E
    new_titles = ['COUNT', 'Grams per', 'Dutchie Qty', 'Dutchie \n Location', 'Last 4 of \n pkd ID', 'Dutchie Product', 'Dutchie Vendor']
    for index, title in enumerate(new_titles, start=5):  # Start from column E (5)
        cell = ws.cell(row=1, column=index)
        cell.value = title
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

    # Save the workbook as XLSX
    wb.save(output_file)

def clear_output_directory(output_dir):
    # Delete all files in the output directory except 'temp'
    for filename in os.listdir(output_dir):
        if filename != 'temp':
            file_path = os.path.join(output_dir, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

def main():
    input_dir = "METRIC-IN"
    output_dir = "METRIC-OUT"

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    else:
        # Clear the output directory
        clear_output_directory(output_dir)

    print("Cleared METRIC-OUT directory (except 'temp' file)")

    # Get current date in YYYY-MM-DD format
    current_date = datetime.date.today().strftime("%Y-%m-%d")

    # Process all CSV files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.csv'):
            input_path = os.path.join(input_dir, filename)
            output_filename = f'METRIC-{current_date}.xlsx'
            output_path = os.path.join(output_dir, output_filename)
            process_metric_file(input_path, output_path)
            print(f"Processed: {filename} -> {output_filename}")

if __name__ == "__main__":
    main()
