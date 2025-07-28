# TSSR - CSV to Excel Converter Functionality

## Overview

This Flask web application allows users to upload a CSV file, processes the data, and returns a formatted Excel (.xlsx) file with enhanced features and formatting.

## Key Features

### File Upload & Processing
- **File Upload:** Accepts CSV files up to 35MB in size
- **Encoding Support:** Handles both UTF-8 and Latin-1 encoding
- **Chunked Processing:** Reads large CSV files in chunks for memory efficiency
- **Error Handling:** Comprehensive error handling with user-friendly messages

### Data Processing
- **Date Conversion:** Converts specific columns to datetime format
- **Email Processing:** Extracts username (`Em`) and domain (`Dom`) from "Email address" column
- **Formula Columns:** Adds two calculated columns:
  - `EmCap#`: Excel formula to count capital letters in `Em`
  - `EmNum#`: Excel formula to count numeric digits in `Em`

### Column Management
- **Predefined Order:** Columns are reordered according to specification:
  1. Panelist id, Created, Date last answer, Date last sent
  2. Last login, Last updated, Date of birth, Postal code
  3. Em, Dom, Email address, EmCap#, EmNum#
  4. Gender, Postal name, First name, Last name
  5. Phone number, Address, Cell phone number
  6. Member id, Panel id, Panelist status, Invitations
  7. Ranking, Recruitment source, Year of birth
  8. Current points balance, Days between mail outs
- **Dynamic Handling:** Extra columns are appended at the end
- **Empty Column Removal:** Automatically removes completely empty columns

### Excel Formatting
- **Header Alignment:** All column headers are left-aligned
- **Column Widths:** Custom widths for optimal display:
  - Key columns: 11.5 Excel units
  - Em: 23 units, Dom: 19 units
  - Email address: 33 units
  - EmCap#/EmNum#: 6 units each
- **Date Formatting:** All date columns formatted as `yyyy-mm-dd`
- **Header Cleanup:** Removes "Date " prefix from column headers

### Conditional Formatting
- **Individual Formatting (Green→Yellow→Red):**
  - Panelist id, Date of birth, Postal code, EmCap#, EmNum#
- **Collective Formatting (Red→Yellow→Green):**
  - Created, Last login, Last updated, Date last answer, Date last sent

### Data Sorting
- **Multi-level Sorting:** Sheet is sorted in sequence:
  1. By "Em" column
  2. By "Panelist id" column  
  3. By "Dom" column

### Sheet Features
- **Frozen Headers:** First row is frozen for easy navigation
- **Auto-filtering:** Applied to all columns for data filtering
- **Multiple Sheets:** Creates additional "Date-Dom Pivot" sheet

### Pivot Analysis Sheet
- **Pivot Table:** Shows counts of Created dates vs Dom values
- **Top Domains:** Displays top 20 Dom values, aggregates others as "Others"
- **Sorting:** Rows by date (latest first), columns by total counts
- **Totals:** Includes row and column totals with highlighting
- **Conditional Formatting:** Yellow (highest) to white (lowest) color scale
- **Column Management:** "Others" column positioned before "Row Total"

### Security & Performance
- **File Cleanup:** Automatically deletes files older than 1 hour
- **Memory Efficient:** In-memory processing without permanent storage
- **File Validation:** Only accepts CSV files
- **Size Limits:** 35MB maximum file size

### Output Features
- **Timestamped Filenames:** Original filename plus timestamp
- **Download Ready:** Automatic download trigger in browser
- **Excel Compatibility:** Full .xlsx format with formulas and formatting

## Technical Implementation

### Dependencies
- Flask: Web framework
- Pandas: Data processing
- OpenPyXL: Excel file generation
- Gunicorn: Production server

### Deployment
- **Google App Engine Ready:** Configured for GAE deployment
- **Static File Handling:** Proper static file configuration
- **Environment Variables:** PORT configuration support

## Usage Workflow

1. User uploads CSV file via web interface
2. File is validated and processed in chunks
3. Data transformation and column reordering
4. Excel file generation with formatting
5. Pivot sheet creation and analysis
6. Automatic download with timestamped filename
7. Error handling and user feedback throughout process

## Error Handling

- File size validation
- Encoding detection and fallback
- Missing column detection
- Processing error recovery
- User-friendly error messages
