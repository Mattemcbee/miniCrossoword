from openpyxl import Workbook

# Create a new workbook
wb = Workbook()

# Get the active worksheet
ws = wb.active

# Write "Hello" to cell A1
ws['A1'] = "Hello"

# Save the workbook
excel_file_name = 'hello.xlsx'
try:
    wb.save(excel_file_name)
    print(f"Excel file '{excel_file_name}' has been created with 'Hello' written in cell A1.")
except Exception as e:
    print(f"Error saving Excel file: {e}")
finally:
    # Close the workbook
    wb.close()
