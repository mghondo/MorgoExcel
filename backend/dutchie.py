import os
import csv
import openpyxl
import datetime

def process_dutchie_file(input_file, output_file):
    try:
        # Create a new workbook
        wb = openpyxl.Workbook()
        ws = wb.active

        # Read the CSV file
        with open(input_file, 'r', newline='') as csvfile:
            csv_reader = csv.reader(csvfile)
            for row in csv_reader:
                ws.append(row)

        # Insert a new column A with title "Column"
        ws.insert_cols(1)
        ws['A1'] = "Count"

        # Copy columns B and C (originally A and B) and insert them to the right of column D (originally C)
        ws.insert_cols(5, 2)
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=5).value = ws.cell(row=row, column=2).value  # Copy column B to E
            ws.cell(row=row, column=6).value = ws.cell(row=row, column=3).value  # Copy column C to F

        # Delete columns B and C (originally A and B)
        ws.delete_cols(2, 2)

        # Save the workbook as XLSX
        wb.save(output_file)
        print(f"Successfully saved: {output_file}")
    except Exception as e:
        print(f"Error processing file {input_file}: {str(e)}")

def clear_output_directory(output_dir):
    for filename in os.listdir(output_dir):
        if filename != 'temp':
            file_path = os.path.join(output_dir, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                    print(f"Deleted: {file_path}")
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

def main():
    input_dir = "DUTCHIE-IN"
    output_dir = "DUTCHIE-OUT"

    print(f"Input directory: {os.path.abspath(input_dir)}")
    print(f"Output directory: {os.path.abspath(output_dir)}")

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Clear the output directory
    clear_output_directory(output_dir)
    print("Cleared DUTCHIE-OUT directory (except 'temp' file)")

    # Get current date in YYYY-MM-DD format
    current_date = datetime.date.today().strftime("%Y-%m-%d")

    # Process all CSV files in the input directory
    files_processed = 0
    for filename in os.listdir(input_dir):
        if filename.endswith('.csv'):
            input_path = os.path.join(input_dir, filename)
            output_filename = f'DUTCHIE-{current_date}.xlsx'
            output_path = os.path.join(output_dir, output_filename)
            process_dutchie_file(input_path, output_path)
            print(f"Processed: {filename} -> {output_filename}")
            files_processed += 1

    if files_processed == 0:
        print("No CSV files found in the input directory.")
    else:
        print(f"Total files processed: {files_processed}")

if __name__ == "__main__":
    main()
