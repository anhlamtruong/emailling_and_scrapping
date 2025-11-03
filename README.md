# Automatic Email Sender

This project is designed to send automatic emails using information from an Excel sheet. It uses Google email services, and the password in the `config.py` file is the Google App Password.

## Features

- Scrapes and validates data from various sources.
- Combines and processes data into a structured format.
- Sends personalized emails using templates.

## Prerequisites

1. **Python**: Ensure Python 3.8 or higher is installed on your system.
2. **Google App Password**: Generate an App Password for your Google account. This is required for authentication.
3. **Dependencies**: Install the required Python packages listed in `requirements.txt`.

## Installation

1. Clone the repository:

   ```bash
   git clone <repository-url>
   cd scapping-and-validate-data
   ```

2. Create a virtual environment (optional but recommended):

   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Setup

1. **Set up the `config` file**:

   - Edit a config.py file in the sending_email directory.

## Usage

1. **Send Emails**:

   - Run the `main.py` script from the root directory:
     ```bash
     python main.py
     ```

2. **Templates**:

   - Email templates are located in the `sending_email/templates` directory. Modify these templates as needed.

## File Structure

- `scrapping-companies/`: Contains scripts for scraping and processing data.
- `sending_email/`: Contains scripts for sending emails and managing templates.
- `requirements.txt`: Lists the Python dependencies.
- `.env`: Stores sensitive information like email credentials.

## Notes

- **Google App Password**: Ensure you use an App Password instead of your Google account password for security.
- **Excel File**: Ensure the Excel file is formatted correctly to match the script's requirements.
- **Error Handling**: Check the logs for any errors during execution.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contributing

Feel free to fork this repository and submit pull requests for improvements or bug fixes.

## Contact

For any questions or issues, please contact [your-email]@gmail.com.
