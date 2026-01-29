# Excel File Data Filler Script

This script reads an Excel file, finds empty cells, and fills them with data from various sources (Google Sheets, APIs, databases, etc.).

## Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Prepare Your Excel File
- Place your Excel file in the same directory as the script
- Update the `EXCEL_FILE` variable in the script to match your filename
- Example structure:
  ```
  | Name    | Email | Phone |
  |---------|-------|-------|
  | John    |       |       |
  | Jane    |       |       |
  | Bob     |       |       |
  ```

## Usage

### Basic Usage
```bash
python excel_uploader.py
```

### Using Google Sheets as Data Source

1. **Set up Google API credentials:**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project
   - Enable the Google Sheets API
   - Create a service account
   - Download the credentials JSON file and save as `credentials.json`

2. **Update the script:**
   - Modify the `lookup_from_google_sheets()` function with your sheet ID
   - Update the `example_lookup_function()` to call Google Sheets

3. **Run the script:**
   ```bash
   python excel_uploader.py
   ```

## Customization

### Custom Data Lookup Function

Edit the `example_lookup_function()` to fit your data source:

```python
def example_lookup_function(row_num: int, row_data: dict) -> str:
    # row_data contains all columns for current row
    # Example: row_data = {'A': 'John', 'B': None, 'C': None}
    
    # Get value from column A (Name)
    name = row_data.get('A')
    
    # Look up data (customize this)
    if name == 'John':
        return 'john@example.com'
    elif name == 'Jane':
        return 'jane@example.com'
    
    return None
```

### Using a Database

```python
import sqlite3

def lookup_from_database(row_data: dict) -> str:
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()
    
    name = row_data.get('A')
    cursor.execute("SELECT email FROM users WHERE name = ?", (name,))
    result = cursor.fetchone()
    
    conn.close()
    return result[0] if result else None
```

### Using an API

```python
import requests

def lookup_from_api(row_data: dict) -> str:
    name = row_data.get('A')
    
    response = requests.get(f'https://api.example.com/users?name={name}')
    if response.status_code == 200:
        data = response.json()
        return data.get('email')
    
    return None
```

## Configuration

Edit these variables in the script:

```python
EXCEL_FILE = "data.xlsx"        # Your input Excel file
OUTPUT_FILE = "data_filled.xlsx" # Output file with filled data
TARGET_COLUMN = "B"              # Column letter with empty cells (A, B, C, etc.)
```

## Output

The script will:
1. Load the Excel file
2. Find empty cells in the specified column
3. Look up data for each empty cell
4. Fill the cells with found data
5. Save to the output file

## Example Output
```
✓ Successfully loaded Excel file: data.xlsx
✓ Found 3 empty cells in column B

📊 Looking up data for empty cells...
  Filled cell (2, 2) with: Data for John
  Filled cell (3, 2) with: Data for Jane
  Filled cell (4, 2) with: Data for Bob
✓ Successfully filled 3 cells

💾 Saving updated Excel file...
✓ Successfully saved Excel file: data_filled.xlsx

✅ Complete! Saved to: data_filled.xlsx
```

## Troubleshooting

- **"No module named 'openpyxl'"**: Run `pip install -r requirements.txt`
- **Google API errors**: Ensure `credentials.json` is in the script directory
- **File not found**: Check that `EXCEL_FILE` variable points to correct path
- **No empty cells found**: Verify target column letter is correct

## Features

- ✅ Read Excel files (.xlsx, .xls)
- ✅ Identify empty cells automatically
- ✅ Fill data from multiple sources (Google Sheets, APIs, databases)
- ✅ Preserve original Excel formatting
- ✅ Support for all columns
- ✅ Detailed logging and progress feedback
