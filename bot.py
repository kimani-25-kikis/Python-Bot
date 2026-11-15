import openpyxl
import pandas as pd

# This is the filename the bot will look for.
# Make sure it matches your actual file EXACTLY.
FILE_NAME = "Sales.xlsx"

def inspect_my_data():
    print(f"--- Opening {FILE_NAME} ---")
    
    try:
        # 1. Load the workbook (the whole file)
        wb = openpyxl.load_workbook(FILE_NAME)
        print(f"Tabs found in this file: {wb.sheetnames}")
        
        # 2. Load just the 'Order' sheet using pandas
        print("\n--- Looking at the 'Order' sheet ---")
        df = pd.read_excel(FILE_NAME, sheet_name="Order")
        
        # 3. Print the columns and first 3 rows
        print("\nHere are the columns the bot sees:")
        print(list(df.columns))
        
        print("\nHere is the first 3 rows of data:")
        print(df.head(3))
        
    except FileNotFoundError:
        print(f"\n❌ ERROR: I couldn't find a file named '{FILE_NAME}' in this folder.")
        print("Check that the file is in the left sidebar and the name matches exactly.")

if __name__ == "__main__":
    inspect_my_data()