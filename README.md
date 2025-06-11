### Step 1: Set Up Your Environment

1. **Install Required Packages**:
   Make sure you have Flask and Pandas installed, along with OpenPyXL for Excel file handling. You can install these packages using pip:

   ```bash
   pip install Flask pandas openpyxl
   ```

2. **Create the Project Structure**:
   Create a directory for your project and navigate into it:

   ```bash
   mkdir flask_csv_to_excel
   cd flask_csv_to_excel
   ```

3. **Create the Flask Application**:
   Create a file named `app.py` in your project directory.

### Step 2: Write the Flask Application Code

Here’s a simple implementation of the Flask application:

```python
from flask import Flask, request, send_file, render_template
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill

app = Flask(__name__)

# Ensure the uploads directory exists
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']
    
    if file.filename == '':
        return "No selected file"
    
    if file and file.filename.endswith('.csv'):
        # Save the uploaded CSV file
        csv_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(csv_path)

        # Convert CSV to DataFrame
        df = pd.read_csv(csv_path)

        # Perform column shuffling (example: shuffle columns)
        df = df.sample(frac=1, axis=1)

        # Save DataFrame to Excel
        excel_path = os.path.join(UPLOAD_FOLDER, 'output.xlsx')
        df.to_excel(excel_path, index=False)

        # Apply conditional formatting
        apply_conditional_formatting(excel_path)

        return send_file(excel_path, as_attachment=True)

    return "Invalid file format. Please upload a CSV file."

def apply_conditional_formatting(excel_path):
    wb = Workbook()
    ws = wb.active

    # Load the existing Excel file
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Example: Apply conditional formatting to the first column
    fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    for row in range(2, ws.max_row + 1):  # Skip header
        if ws.cell(row=row, column=1).value > 100:  # Example condition
            ws.cell(row=row, column=1).fill = fill

    # Save the modified Excel file
    wb.save(excel_path)

if __name__ == '__main__':
    app.run(debug=True)
```

### Step 3: Create the HTML Template

Create a folder named `templates` in your project directory and create a file named `index.html` inside it:

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CSV to Excel Converter</title>
</head>
<body>
    <h1>Upload CSV File</h1>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".csv" required>
        <input type="submit" value="Upload">
    </form>
</body>
</html>
```

### Step 4: Run the Application

1. **Run the Flask Application**:
   In your terminal, run the following command:

   ```bash
   python app.py
   ```

2. **Access the Application**:
   Open your web browser and go to `http://127.0.0.1:5000/`. You should see the upload form.

### Step 5: Test the Application

1. Upload a CSV file.
2. The application will convert it to Excel format, shuffle the columns, apply conditional formatting, and return the updated file for download.

### Notes

- This is a basic implementation. You can enhance it by adding error handling, logging, and more complex formatting options.
- Make sure to handle file cleanup (deleting uploaded files) as needed to avoid filling up your server with temporary files.
- You can customize the conditional formatting logic based on your requirements.