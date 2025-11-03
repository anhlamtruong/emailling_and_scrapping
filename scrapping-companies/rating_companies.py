import pandas as pd
from google import genai
import time
# --- Configuration ---
# It's recommended to use an environment variable for security.
API_KEY = "AIzaSyD9RbJbZdNanc-vk-FKycthnZBeLc-SfOI" 
INPUT_FILE = './extracted_companies.xlsx'
OUTPUT_FILE = 'company_ratings.xlsx'
COMPANY_COLUMN_NAME = 'Company Name'

# --- Setup Gemini API ---
try:
    client = genai.Client(api_key=API_KEY)
    print("Successfully configured Gemini API.")
except Exception as e:
    print(f"Error configuring Gemini API: {e}")
    exit()

# --- Main Application Logic ---
def get_company_rating(company_name):
    """
    Asks Gemini to rate a single company and returns the response.
    """
    # This prompt is designed to be token-efficient and get a structured response.
    prompt = (
        f"Rate the company '{company_name}' for a software engineering career on a scale of 1 to 10. "
        "Respond with only the rating and a brief, one-sentence explanation. "
        "Your entire response must always be in the format: Rating: [number]/10. Explanation: [text with only one sentence]."
    )
    
    try:
        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        # Simple parsing based on the requested format
        parts = response.text.split('. Explanation: ')
        rating = parts[0].replace('Rating: ', '').strip()
        explanation = parts[1].strip() if len(parts) > 1 else "No explanation provided."
        return rating, explanation
    except Exception as e:
        print(f"  - Could not get rating for {company_name}. Error: {e}")
        return "Error", str(e)

def main():
    """
    Main function to read companies, get ratings, and save results.
    """
    print(f"Reading companies from '{INPUT_FILE}'...")
    try:
        # Reads the first sheet of the Excel file
        df = pd.read_excel(INPUT_FILE, sheet_name=0) 
        if COMPANY_COLUMN_NAME not in df.columns:
            print(f"Error: Column '{COMPANY_COLUMN_NAME}' not found in the Excel file.")
            return
    except FileNotFoundError:
        print(f"Error: The file '{INPUT_FILE}' was not found.")
        return
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return

    companies = df[COMPANY_COLUMN_NAME].dropna().unique().tolist()
    print(f"Found {len(companies)} unique companies to rate.")

    results = []
    for i, company in enumerate(companies):
        print(f"[{i+1}/{len(companies)}] Rating '{company}'...")
        rating, explanation = get_company_rating(company)
        results.append({
            'Company': company,
            'Rating': rating,
            'Explanation': explanation
        })
        time.sleep(1) 

    print("\nAll companies have been rated. Saving results...")
    results_df = pd.DataFrame(results)
    results_df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Success! Results saved to '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    main()