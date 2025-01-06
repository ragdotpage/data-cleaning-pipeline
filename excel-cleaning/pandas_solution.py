import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from decorator import calculate_time

@calculate_time
def process_excel_file(input_file, output_file):
    # Load workbook to handle merged cells
    wb = load_workbook(input_file)
    ws = wb.active

    # Find the row containing the main headers
    main_header_row = None
    main_headers = ["Part or Item Number", "Item Description"]
    
    for row in range(1, ws.max_row + 1):
        values = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
        if any(header in str(value) for value in values for header in main_headers):
            main_header_row = row
            break

    if main_header_row is None:
        raise ValueError("Could not find main headers in the file")

    # Initialize lists to store the consolidated data for each column
    consolidated_data = [[] for _ in range(ws.max_column)]
    
    # Process rows before main headers to consolidate unrelated data
    for row in range(1, main_header_row):
        row_values = []
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            value = cell.value if cell.value is not None else ""
            
            # Check if cell is part of a merged range
            is_merged = False
            merged_value = value
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    is_merged = True
                    merged_value = ws.cell(merged_range.min_row, merged_range.min_col).value
                    break
            
            row_values.append(str(merged_value if is_merged else value).strip())
        
        # Add non-empty values to their respective columns
        for col, value in enumerate(row_values):
            if value and value != "nan":
                consolidated_data[col].append(value)

    # Create the final processed data
    processed_data = []
    
    # Add consolidated headers
    header_row = []
    for col in range(ws.max_column):
        header = " ".join(consolidated_data[col])
        if header:
            header_row.append(header + " " + str(ws.cell(main_header_row, col + 1).value))
        else:
            header_row.append(str(ws.cell(main_header_row, col + 1).value))
    processed_data.append(header_row)
    
    # Add the actual data rows
    for row in range(main_header_row + 1, ws.max_row + 1):
        row_data = []
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            value = cell.value if cell.value is not None else ""
            row_data.append(str(value).strip())
        if any(value.strip() for value in row_data):  # Only add non-empty rows
            processed_data.append(row_data)

    # Create DataFrame and save
    df = pd.DataFrame(processed_data[1:], columns=processed_data[0])
    
    # Remove any completely empty rows
    df_cleaned = df.dropna(how='all')
    
    # Reset the index after removing rows
    df_cleaned = df_cleaned.reset_index(drop=True)
    
    # Save the cleaned DataFrame to a new Excel file
    df_cleaned.to_excel(output_file, index=False)
    print(f"File processed and saved as: {output_file}")
    
    # Print the headers for verification
    print("\nColumn headers:")
    for header in processed_data[0]:
        print(f"- {header}")

    return df_cleaned


if __name__ == "__main__":
    input_file = "inventory_copy.xlsx"  # Replace with your input file name
    output_file = "inventory_cleaned.xlsx"  # Replace with your desired output file name

    try:
        process_excel_file(input_file, output_file)
    except Exception as e:
        print(f"An error occurred: {str(e)}")
