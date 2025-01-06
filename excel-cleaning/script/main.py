# main.py
from excel_utils import (
    load_excel_worksheet,
    initialize_column_headers,
    process_merged_cells,
    process_non_merged_cells,
    combine_headers,
    process_dataframe,
    save_dataframe,
    print_headers
)


def process_excel_file(file_path: str) -> None:
    """
    Process an Excel file: handle merged cells, clean data, and update headers.

    Args:
        file_path: Path to the Excel file to process
    """
    try:
        # Load worksheet
        worksheet = load_excel_worksheet(file_path)

        # Initialize and process headers
        column_headers = initialize_column_headers(worksheet)
        max_row = process_merged_cells(worksheet, column_headers)
        process_non_merged_cells(worksheet, column_headers, max_row)

        # Create final headers and process data
        final_headers = combine_headers(column_headers, worksheet.max_column)
        df_cleaned = process_dataframe(file_path, max_row, final_headers)

        # Save and report results
        save_dataframe(df_cleaned, file_path)
        print_headers(final_headers, file_path)

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        raise


if __name__ == "__main__":
    file_path = "inventory.xlsx"  # Replace with your file path
    process_excel_file(file_path)
