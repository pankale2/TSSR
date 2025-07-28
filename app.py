import os
import shutil
from datetime import datetime, timedelta
from io import StringIO, BytesIO
from flask import Flask, render_template, request, send_file, redirect, flash
import pandas as pd
from werkzeug.utils import secure_filename
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Add upload folder configuration
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

MAX_FILE_SIZE = 35 * 1024 * 1024  # 35MB

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'csv'

def cleanup_old_files():
    """Delete files older than 1 hour from uploads folder"""
    threshold = datetime.now() - timedelta(hours=1)
    for filename in os.listdir(UPLOAD_FOLDER):
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.getctime(filepath) < threshold.timestamp():
            try:
                os.remove(filepath)
            except Exception as e:
                print(f"Error deleting {filepath}: {e}")

@app.route('/', methods=['GET', 'POST'])
def index():
    # Cleanup old files on each request
    cleanup_old_files()
    
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            # File size check
            file.seek(0, os.SEEK_END)
            file_length = file.tell()
            file.seek(0)
            if file_length > MAX_FILE_SIZE:
                flash(f'File too large. Please upload a file smaller than {MAX_FILE_SIZE // (1024*1024)}MB.')
                return redirect(request.url)
            
            try:
                # Use file.stream directly for pandas
                try:
                    csv_data = StringIO(file.stream.read().decode("UTF-8"))
                except UnicodeDecodeError:
                    file.stream.seek(0)
                    csv_data = StringIO(file.stream.read().decode("latin-1"))  # Try latin-1 encoding if UTF-8 fails
                
                chunk_size = 10000  # Adjust based on memory constraints
                df_chunks = []
                for chunk in pd.read_csv(csv_data, sep=';', low_memory=False, chunksize=chunk_size):
                    # Convert date columns to datetime
                    date_columns = [
                        'Created', 'Date last answer', 'Date last sent', 'Last login', 'Last updated', 'Date of birth'
                    ]
                    for col in date_columns:
                        if col in chunk.columns:
                            chunk[col] = pd.to_datetime(chunk[col], errors='coerce', dayfirst=False)

                    # Create 'Em' and 'Dom' columns from 'Email address'
                    if 'Email address' in chunk.columns:
                        emails = chunk['Email address'].astype(str).fillna('')
                        chunk['Em'] = emails.str.split('@').str[0]
                        chunk['Dom'] = emails.str.split('@').str[1].fillna('')
                        # Add placeholder columns for EmCap# and EmNum#
                        chunk['EmCap#'] = ''
                        chunk['EmNum#'] = ''
                    else:
                        flash('No "Email address" column found in CSV.')
                        return redirect(request.url)

                    df_chunks.append(chunk)
                df = pd.concat(df_chunks, ignore_index=True)

                desired_order = [
                    'Panelist id', 'Created', 'Date last answer', 'Date last sent', 'Last login', 'Last updated',
                    'Date of birth', 'Postal code', 'Em', 'Dom', 'Email address', 'EmCap#', 'EmNum#',
                    'Gender', 'Postal name', 'First name', 'Last name', 'Phone number', 'Address',
                    'Cell phone number', 'Member id', 'Panel id', 'Panelist status', 'Invitations',
                    'Ranking', 'Recruitment source', 'Year of birth', 'Current points balance', 'Days between mail outs'
                ]
                columns_to_use = [col for col in desired_order if col in df.columns]
                columns_to_use += [col for col in df.columns if col not in columns_to_use]
                df = df.reindex(columns=columns_to_use)

                # No need to move EmCap# and EmNum# after Em, as they are now after Email address in desired_order

                # --- Place static EmCap# and EmNum# population here, after reindexing ---
                if 'Em' in df.columns and 'EmCap#' in df.columns and 'EmNum#' in df.columns:
                    def count_capitals(s):
                        return sum(1 for c in str(s) if c.isupper())
                    def count_digits(s):
                        return sum(1 for c in str(s) if c.isdigit())
                    df['EmCap#'] = df['Em'].apply(count_capitals)
                    df['EmNum#'] = df['Em'].apply(count_digits)
                # --- End static population ---

                # Create Excel file in memory
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl',
                                    datetime_format='yyyy-mm-dd',
                                    date_format='yyyy-mm-dd') as writer:
                    df.to_excel(writer, index=False)
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']

                    # Set all column header texts to left-aligned
                    from openpyxl.styles import Alignment
                    for cell in worksheet[1]:
                        cell.alignment = Alignment(horizontal='left')

                    # Optionally, set number format for date columns in Excel
                    from openpyxl.styles import numbers
                    date_columns = [
                        'Created', 'Date last answer', 'Date last sent', 'Last login', 'Last updated', 'Date of birth'
                    ]
                    for col in date_columns:
                        if col in df.columns:
                            col_idx = df.columns.get_loc(col) + 1
                            for row in range(2, len(df) + 2):
                                cell = worksheet.cell(row=row, column=col_idx)
                                cell.number_format = 'yyyy-mm-dd'

                    # Define the color scale rules
                    individual_color_rule = ColorScaleRule(
                        start_type='min',
                        start_color='FF63BE7B',  # Green
                        mid_type='percentile',
                        mid_value=50,
                        mid_color='FFFEB64F',  # Yellow
                        end_type='max',
                        end_color='FFF8696B'    # Red
                    )
                    collective_color_rule = ColorScaleRule(
                        start_type='min',
                        start_color='FFF8696B',  # Red
                        mid_type='percentile',
                        mid_value=50,
                        mid_color='FFFEB64F',  # Yellow
                        end_type='max',
                        end_color='FF63BE7B'    # Green
                    )

                    # Apply individual formatting after header is defined
                    header = [cell.value for cell in worksheet[1]]
                    individual_columns = ['Panelist id', 'Date of birth', 'Postal code', 'EmCap#', 'EmNum#']
                    for col in individual_columns:
                        if col in header:
                            col_idx = header.index(col) + 1
                            col_letter = get_column_letter(col_idx)
                            col_range = f'{col_letter}2:{col_letter}{len(df) + 1}'
                            worksheet.conditional_formatting.add(col_range, individual_color_rule)

                    # Apply collective formatting
                    collective_columns = ['Created', 'Last login', 'Last updated', 'Date last answer', 'Date last sent']
                    ranges = []
                    for col in collective_columns:
                        if col in header:
                            col_idx = header.index(col) + 1
                            col_letter = get_column_letter(col_idx)
                            ranges.append(f'{col_letter}2:{col_letter}{len(df) + 1}')
                    if ranges:
                        worksheet.conditional_formatting.add(' '.join(ranges), collective_color_rule)

                    # Remove the word "Date " from any column headers if present
                    for cell in worksheet[1]:
                        if cell.value and isinstance(cell.value, str) and cell.value.startswith("Date "):
                            cell.value = cell.value.replace("Date ", "", 1)

                    # Remove empty columns from the Excel sheet before finalizing
                    for col in reversed(range(1, worksheet.max_column + 1)):
                        col_letter = get_column_letter(col)
                        if all(worksheet.cell(row=row, column=col).value in (None, '', float('nan')) for row in range(2, worksheet.max_row + 1)):
                            worksheet.delete_cols(col)

                    # --- SET COLUMN WIDTHS AFTER ALL MODIFICATIONS ---
                    header = [cell.value for cell in worksheet[1]]
                    
                    # Use renamed column headers after "Date " removal
                    columns_to_resize = [
                        'Panelist id', 'Created', 'last answer', 'last sent',
                        'Last login', 'Last updated', 'of birth', 'Postal code'
                    ]
                    width_units = 11.5
                    for col in columns_to_resize:
                        if col in header:
                            col_idx = header.index(col) + 1
                            col_letter = get_column_letter(col_idx)
                            worksheet.column_dimensions[col_letter].width = width_units

                    custom_widths = {
                        'Em': 23,
                        'Dom': 19,
                        'Email address': 33,
                        'EmCap#': 6,
                        'EmNum#': 6
                    }
                    for col, width in custom_widths.items():
                        if col in header:
                            col_idx = header.index(col) + 1
                            col_letter = get_column_letter(col_idx)
                            worksheet.column_dimensions[col_letter].width = width

                    # Sort the entire sheet 3 times: Em, then Panelist id, then Dom - MOVED TO END
                    def sort_worksheet(ws, sort_col_name):
                        header = [cell.value for cell in ws[1]]
                        if sort_col_name not in header:
                            return
                        sort_col_idx = header.index(sort_col_name) + 1
                        data = []
                        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
                            data.append(row)
                        data.sort(key=lambda x: (x[sort_col_idx-1] is None, x[sort_col_idx-1]))
                        for i, row in enumerate(data, start=2):
                            for j, value in enumerate(row, start=1):
                                ws.cell(row=i, column=j, value=value)

                    sort_worksheet(worksheet, "Em")
                    sort_worksheet(worksheet, "Panelist id")
                    sort_worksheet(worksheet, "Dom")

                    # Freeze the first row (must be last)
                    worksheet.freeze_panes = 'A2'

                    # Apply autofilter to all columns (must be last)
                    worksheet.auto_filter.ref = worksheet.dimensions

                    # --- Add Date-Dom Pivot Sheet ---
                    df_pivot = df.copy()
                    if 'Created' in df_pivot.columns:
                        df_pivot['Created'] = pd.to_datetime(df_pivot['Created'], errors='coerce').dt.date
                    else:
                        df_pivot['Created'] = pd.NaT

                    # Get top 20 Dom values by count, aggregate others
                    dom_counts = df_pivot['Dom'].value_counts()
                    top_doms = dom_counts.nlargest(20).index.tolist()
                    df_pivot['Dom_pivot'] = df_pivot['Dom'].where(df_pivot['Dom'].isin(top_doms), 'Others')

                    # Build pivot table
                    pivot = pd.pivot_table(
                        df_pivot,
                        index='Created',
                        columns='Dom_pivot',
                        values='Panelist id',
                        aggfunc='count',
                        fill_value=0
                    )

                    # Sort rows by date descending
                    pivot = pivot.sort_index(ascending=False)

                    # Sort columns by total counts descending, but keep "Others" at the end
                    col_totals = pivot.sum(axis=0)
                    sorted_cols = col_totals.sort_values(ascending=False).index.tolist()
                    
                    # Move "Others" to the end if it exists
                    if 'Others' in sorted_cols:
                        sorted_cols.remove('Others')
                        sorted_cols.append('Others')
                    
                    pivot = pivot[sorted_cols]

                    # Add Row Total
                    pivot['Row Total'] = pivot.sum(axis=1)

                    # Add Column Total row
                    pivot.loc['Column Total'] = pivot.sum(axis=0)

                    # Write to new sheet
                    pivot_sheetname = "Date-Dom Pivot"
                    pivot.to_excel(writer, sheet_name=pivot_sheetname)
                    pivot_ws = writer.sheets[pivot_sheetname]

                    # Format Created column as date
                    from openpyxl.styles import PatternFill
                    created_col_idx = 1
                    for row in range(2, pivot_ws.max_row):
                        cell = pivot_ws.cell(row=row, column=created_col_idx)
                        cell.number_format = 'yyyy-mm-dd'

                    # Conditional formatting (yellow for highest, white for lowest, exclude totals)
                    last_data_row = pivot_ws.max_row - 1
                    last_data_col = pivot_ws.max_column - 1
                    if last_data_row > 1 and last_data_col > 1:
                        data_range = f"{get_column_letter(2)}2:{get_column_letter(last_data_col)}{last_data_row}"
                        pivot_color_rule = ColorScaleRule(
                            start_type='min',
                            start_color='FFFFFF',  # White
                            end_type='max',
                            end_color='FFFF00'     # Yellow
                        )
                        pivot_ws.conditional_formatting.add(data_range, pivot_color_rule)

                    # Set header fill for totals
                    total_fill = PatternFill(start_color="FFFEB64F", end_color="FFFEB64F", fill_type="solid")
                    # Row Total header
                    pivot_ws.cell(row=1, column=pivot_ws.max_column).fill = total_fill
                    # Column Total row
                    for col in range(1, pivot_ws.max_column + 1):
                        pivot_ws.cell(row=pivot_ws.max_row, column=col).fill = total_fill

                    # Set column widths for pivot sheet
                    for col in range(1, pivot_ws.max_column + 1):
                        pivot_ws.column_dimensions[get_column_letter(col)].width = 14

                # End of writer context - Excel file is now complete

                # Ensure buffer is at start and not truncated
                excel_buffer.seek(0)

                # Generate filename with proper timestamp
                from datetime import datetime as dt
                original_name = os.path.splitext(secure_filename(file.filename))[0]
                timestamp = dt.now().strftime('%Y%m%d_%H%M%S')
                download_name = f"{original_name}_{timestamp}.xlsx"

                return send_file(
                    excel_buffer,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name=download_name
                )
            except Exception as e:
                flash(f'Error processing file: {str(e)}')
                return redirect(request.url)
    return render_template('index.html')

if __name__ == '__main__':
    # Only open browser once, avoid duplicate tabs
    import webbrowser
    import threading

    def open_browser_once():
        # Only open if not already opened (avoid multiple triggers)
        webbrowser.open_new('http://127.0.0.1:8080/')

    threading.Timer(1.0, open_browser_once).start()
    app.run(host='127.0.0.1', port=8080, debug=True, use_reloader=False)