import pandas as pd
from bs4 import BeautifulSoup

# --- Configuration ---
HTML_FILE_TO_READ = 'index.html'
OUTPUT_EXCEL_FILE = 'extracted_companies.xlsx'
COLUMN_NAME = 'Company Name'
# ---------------------

def extract_company_names(html_content):
    """
    Parses the HTML and extracts company names based on the h2 tag.
    """
    # Initialize BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find all <h2> tags that match the specific attribute
    company_tags = soup.find_all(
        'h2', 
        attrs={'data-dd-action-name': 'Exhibitor name'}
    )
    
    # Extract the text from each tag and clean it
    company_names = [tag.get_text().strip() for tag in company_tags]
    
    return company_names

def main():
    """
    Main function to read the HTML, extract data, and save to Excel.
    """
    print(f"Reading HTML file: '{HTML_FILE_TO_READ}'...")
    try:
        # Read the content of the local HTML file
        with open(HTML_FILE_TO_READ, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except FileNotFoundError:
        print(f"Error: The file '{HTML_FILE_TO_READ}' was not found.")
        print("Please make sure the HTML file is in the same directory.")
        return
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return

    # Extract the data
    company_names = extract_company_names(html_content)

    if not company_names:
        print("No company names were found. The output file will be empty.")
    else:
        print(f"Successfully found {len(company_names)} companies.")

    # Create a pandas DataFrame
    df = pd.DataFrame(company_names, columns=[COLUMN_NAME])

    # Save the DataFrame to an Excel file
    try:
        df.to_excel(OUTPUT_EXCEL_FILE, index=False)
        print(f"✅ Success! Data saved to '{OUTPUT_EXCEL_FILE}'.")
    except Exception as e:
        print(f"An error occurred while saving the Excel file: {e}")

if __name__ == "__main__":
    main()