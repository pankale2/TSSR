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
                    else:
                        flash('No "Email address" column found in CSV.')
                        return redirect(request.url)

                    df_chunks.append(chunk)
                
                df = pd.concat(df_chunks, ignore_index=True)

                # Remove empty columns
                df = df.dropna(axis=1, how='all')

                desired_order = [
                    'Panelist id', 'Created', 'Date last answer', 'Date last sent', 'Last login', 'Last updated',
                    'Date of birth', 'Em', 'Dom', 'Email address', 'First name', 'Gender', 'Last name',
                    'Phone number', 'Postal code', 'Postal name', 'Address', 'Cell phone number', 'Member id',
                    'Panel id', 'Panelist status', 'Invitations', 'Ranking', 'Recruitment source', 'Year of birth',
                    'Current points balance', 'Days between mail outs'
                ]
                columns_to_use = [col for col in desired_order if col in df.columns]
                columns_to_use += [col for col in df.columns if col not in columns_to_use]
                df = df.reindex(columns=columns_to_use)

                # Create Excel file in memory
                excel_buffer = BytesIO()
                try:
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

                        # Set column width for specified columns to ~77px (11 Excel units)
                        columns_to_resize = [
                            'Panelist id', 'Created', 'Date last answer', 'Date last sent',
                            'Last login', 'Last updated', 'Date of birth'
                        ]
                        width_units = 11
                        for col in columns_to_resize:
                            if col in df.columns:
                                col_idx = df.columns.get_loc(col) + 1
                                col_letter = get_column_letter(col_idx)
                                worksheet.column_dimensions[col_letter].width = width_units

                        # Set custom widths for Em, Dom, Email address
                        custom_widths = {
                            'Em': 23,    # ≈ 160px
                            'Dom': 19,   # ≈ 130px
                            'Email address': 33  # ≈ 230px
                        }
                        for col, width in custom_widths.items():
                            if col in df.columns:
                                col_idx = df.columns.get_loc(col) + 1
                                col_letter = get_column_letter(col_idx)
                                worksheet.column_dimensions[col_letter].width = width

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

                        # Apply individual formatting
                        individual_columns = ['Panelist id', 'Created', 'Date of birth']
                        for col in individual_columns:
                            if col in df.columns:
                                col_idx = df.columns.get_loc(col) + 1
                                col_letter = get_column_letter(col_idx)
                                col_range = f'{col_letter}2:{col_letter}{len(df) + 1}'
                                worksheet.conditional_formatting.add(col_range, individual_color_rule)

                        # Apply collective formatting
                        collective_columns = ['Last login', 'Last updated', 'Date last answer', 'Date last sent']
                        ranges = []
                        for col in collective_columns:
                            if col in df.columns:
                                col_idx = df.columns.get_loc(col) + 1
                                col_letter = get_column_letter(col_idx)
                                ranges.append(f'{col_letter}2:{col_letter}{len(df) + 1}')
                        if ranges:
                            worksheet.conditional_formatting.add(' '.join(ranges), collective_color_rule)

                        # Freeze the first row (must be last)
                        worksheet.freeze_panes = 'A2'

                        # Apply autofilter to all columns (must be last)
                        worksheet.auto_filter.ref = worksheet.dimensions
                except Exception as excel_err:
                    flash(f'Error writing to Excel: {str(excel_err)}')
                    return redirect(request.url)

                excel_buffer.seek(0)
                # Use secure_filename but strip trailing underscores and extension
                base_filename = os.path.splitext(secure_filename(file.filename))[0].rstrip('_')
                download_name = f"{base_filename}.xlsx"
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
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))