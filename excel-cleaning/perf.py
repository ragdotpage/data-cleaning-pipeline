from pandas_solution import process_excel_file
from polar_solution import process_excel_file_polars
from decorator import calculate_time

process_excel_file("inventory.xlsx", "inventory_cleaned_pandas.xlsx")
process_excel_file_polars("inventory.xlsx", "inventory_cleaned_polars.xlsx")
