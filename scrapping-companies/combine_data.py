import pandas as pd

# --- Configuration ---
FILE_DATA1 = 'data1.xlsx'  # Your main file (the default)
FILE_DATA2 = 'data2.xlsx'  # The file with new data
OUTPUT_FILE = 'data_combined.xlsx' # The final output file

# Column name to use for matching
MATCH_COLUMN = 'Company Name'
# ---------------------

def main():
    print(f"Reading '{FILE_DATA1}' (default) and '{FILE_DATA2}' (new data)...")
    try:
        # Read both Excel files (assuming data is on the first sheet)
        df1 = pd.read_excel(FILE_DATA1, sheet_name=0)
        df2 = pd.read_excel(FILE_DATA2, sheet_name=0)
    except FileNotFoundError as e:
        print(f"Error: Could not find file {e.filename}.")
        print("Please make sure 'data1.xlsx' and 'data2.xlsx' are in the same directory.")
        return
    except Exception as e:
        print(f"Error reading Excel files: {e}")
        return

    # --- Data Cleaning ---
    # This is critical. It removes extra spaces from the matching column.
    # e.g., "Google " will become "Google"
    try:
        df1[MATCH_COLUMN] = df1[MATCH_COLUMN].astype(str).str.strip()
        df2[MATCH_COLUMN] = df2[MATCH_COLUMN].astype(str).str.strip()
        print("Successfully cleaned company names (removed extra spaces).")
    except KeyError:
        print(f"Error: The match column '{MATCH_COLUMN}' was not found in one of the files.")
        print("Please ensure both files have this exact column name.")
        return
    # ---------------------

    print("Merging matching companies to 'data1'...")
    # 1. Merge data1 and data2, keeping ALL rows from data1 (how='left')
    # This adds columns from data2 to data1 where 'Company Name' matches.
    merged_sheet1 = df1.merge(
        df2, 
        on=MATCH_COLUMN, 
        how='left', 
        suffixes=('_data1', '_data2') # This renames 'Rank' to 'Rank_data1' and 'Rank_data2'
    )
    
    # 2. Find additional companies (in data2 but NOT in data1)
    print("Finding additional companies from 'data2'...")
    
    # We find which data2 companies are NOT in the data1 list.
    additional_companies = df2[~df2[MATCH_COLUMN].isin(df1[MATCH_COLUMN])]

    # 3. Write both DataFrames to a new Excel file
    print(f"Writing results to '{OUTPUT_FILE}'...")
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            # Write the main merged data
            merged_sheet1.to_excel(writer, sheet_name='Sheet1_Combined', index=False)
            
            # Write the additional companies to a new sheet
            additional_companies.to_excel(writer, sheet_name='Additional_Companies', index=False)
        
        print(f"\n✅ Success! Data combined and saved to '{OUTPUT_FILE}'.")
        print("\nWhat this script did:")
        print(f"- 'Sheet1_Combined': Contains all companies from '{FILE_DATA1}'.")
        print("  - Where a match was found, 'Rank_data2' has the new rank.")
        print("  - Where no match was found, 'Rank_data2' will be blank (NaN).")
        print(f"- 'Additional_Companies': Contains all companies from '{FILE_DATA2}' that were NOT in '{FILE_DATA1}'.")
        
    except Exception as e:
        print(f"Error writing Excel file: {e}")

if __name__ == "__main__":
    main()