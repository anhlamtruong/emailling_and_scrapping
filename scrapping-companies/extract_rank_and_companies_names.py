import pandas as pd
from bs4 import BeautifulSoup

# --- Configuration ---
HTML_FILE_TO_READ = 'index.html'
OUTPUT_EXCEL_FILE = 'extracted_ranks_and_names.xlsx'
COLUMN_NAMES = ['Rank', 'Company Name']
# ---------------------

def extract_company_data(html_content):
    """
    Parses the HTML and extracts rank and company name for each company row.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find the table body that contains the company rows
    # Assuming there's only one relevant tbody
    table_body = soup.find('tbody') 
    if not table_body:
        print("Error: Could not find the <tbody> tag.")
        return []

    # Find all table rows within the tbody that represent a company
    company_rows = table_body.find_all('tr', class_='company')
    
    extracted_data = []
    
    for row in company_rows:
        rank = "N/A"
        company_name = "N/A"
        
        # Find the rank cell and extract the rank number
        rank_cell = row.find('td', class_='rank')
        if rank_cell:
            rank_div = rank_cell.find('div', class_='first-line')
            if rank_div:
                rank = rank_div.get_text(strip=True)
        
        # Find the name cell and extract the company name from the <a> tag
        name_cell = row.find('td', class_='name')
        if name_cell:
            name_link = name_cell.find('a')
            if name_link:
                company_name = name_link.get_text(strip=True)
                
        # Only add if we found a valid company name (avoid empty rows)
        if company_name != "N/A":
            extracted_data.append({'Rank': rank, 'Company Name': company_name})
            
    return extracted_data

def main():
    """
    Main function to read the HTML, extract data, and save to Excel.
    """
    print(f"Reading HTML file: '{HTML_FILE_TO_READ}'...")
    try:
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
    company_data = extract_company_data(html_content)

    if not company_data:
        print("No company data found. The output file will be empty.")
    else:
        print(f"Successfully found data for {len(company_data)} companies.")

    # Create a pandas DataFrame
    df = pd.DataFrame(company_data, columns=COLUMN_NAMES)

    # Save the DataFrame to an Excel file
    try:
        df.to_excel(OUTPUT_EXCEL_FILE, index=False)
        print(f"✅ Success! Data saved to '{OUTPUT_EXCEL_FILE}'.")
    except Exception as e:
        print(f"An error occurred while saving the Excel file: {e}")

if __name__ == "__main__":
    main()