import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from decorator import calculate_time

@calculate_time
def process_excel_file(input_file, output_file):
    # Load workbook
    wb = load_workbook(input_file)
    ws = wb.active

    # Find the main header row by looking for consistent data patterns
    main_header_row = None
    max_consistent_cols = 0
    
    for row in range(1, ws.max_row + 1):
        row_values = []
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            # Get actual value from merged cells
            value = None
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    value = ws.cell(merged_range.min_row, merged_range.min_col).value
                    break
            if value is None:
                value = cell.value
            row_values.append(str(value) if value is not None else "")

        # Count non-empty cells in this row
        non_empty = sum(1 for val in row_values if val.strip())
        if non_empty > max_consistent_cols:
            main_header_row = row
            max_consistent_cols = non_empty

    # Create a dictionary to store header values for each column
    column_headers = {}
    
    # Initialize column headers dictionary
    for col in range(1, ws.max_column + 1):
        column_headers[col] = []

    # Process cells row by row up to main_header_row
    for row in range(1, main_header_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            value = None
            is_merged = False
            
            # Check if cell is part of a merged range
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    is_merged = True
                    if cell.coordinate == merged_range.start_cell.coordinate:
                        value = ws.cell(merged_range.min_row, merged_range.min_col).value
                    break
            
            if not is_merged:
                value = cell.value
                
            if value is not None and str(value).strip():
                column_headers[col].append(str(value))

    # Combine headers for each column
    final_headers = []
    for col in range(1, ws.max_column + 1):
        # Preserve original order by not filtering
        header_parts = column_headers[col]
        if header_parts:
            final_headers.append(" ".join(header_parts))
        else:
            final_headers.append("")

    # Read the data portion of the Excel file
    df = pd.read_excel(input_file, header=None, skiprows=main_header_row)
    
    # Set the combined headers
    df.columns = final_headers
    
    # Remove completely empty rows
    df_cleaned = df.dropna(how='all')
    
    # Reset the index after removing rows
    df_cleaned = df_cleaned.reset_index(drop=True)
    
    # Save the cleaned DataFrame to a new Excel file
    df_cleaned.to_excel(output_file, index=False)
    print(f"File processed and saved as: {output_file}")

    # Print the headers for verification
    print("\nColumn headers:")
    for header in final_headers:
        print(f"- {header}")

    return df_cleaned

if __name__ == "__main__":
    input_file = "inventory_copy.xlsx"  # Replace with your input file name
    output_file = "inventory_cleaned.xlsx"  # Replace with your desired output file name

    try:
        process_excel_file(input_file, output_file)
    except Exception as e:
        print(f"An error occurred: {str(e)}")
