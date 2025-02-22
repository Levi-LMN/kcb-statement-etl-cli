from flask import Flask, render_template, request, send_file
import tabula
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create uploads folder if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def create_summary(df):
    """Create summary statistics from the transaction data"""
    summary = {
        'Total Money In': df['Money In'].sum(),
        'Total Money Out': df['Money Out'].sum(),
        'Net Movement': df['Money In'].sum() + df['Money Out'].sum(),  # Changed: just add Money Out
        'Number of Transactions': len(df),
        'Number of Incoming Transactions': df['Money In'].notna().sum(),
        'Number of Outgoing Transactions': df['Money Out'].notna().sum(),
        'Average Transaction In': df['Money In'].mean(),
        'Average Transaction Out': df['Money Out'].mean(),
        'Largest Transaction In': df['Money In'].max(),
        'Largest Transaction Out': df['Money Out'].min(),  # Changed: use min since it's the most negative
        'First Transaction Date': df['Transaction Date'].iloc[0],
        'Last Transaction Date': df['Transaction Date'].iloc[-1],
    }

    # Convert to DataFrame for easy Excel writing
    summary_df = pd.DataFrame(list(summary.items()), columns=['Metric', 'Value'])
    return summary_df


def create_daily_totals(df):
    """Create daily totals from the transaction data"""
    # Convert Transaction Date to datetime if it's not already
    df['Transaction Date'] = pd.to_datetime(df['Transaction Date'])

    # Create daily totals
    daily_totals = df.groupby('Transaction Date').agg({
        'Money In': 'sum',
        'Money Out': 'sum',
        'Transaction Details': 'count'  # Count of transactions per day
    }).reset_index()

    # Calculate net movement for each day (changed: just add Money Out)
    daily_totals['Net Movement'] = daily_totals['Money In'] + daily_totals['Money Out']

    return daily_totals


def create_monthly_totals(df):
    """Create monthly totals and analysis from the transaction data"""
    # Convert Transaction Date to datetime if it's not already
    df['Transaction Date'] = pd.to_datetime(df['Transaction Date'])

    # Add month column for grouping
    df['Month'] = df['Transaction Date'].dt.strftime('%B %Y')

    # Calculate monthly totals
    monthly_totals = df.groupby('Month').agg({
        'Money In': 'sum',
        'Money Out': 'sum',
        'Transaction Details': 'count',
    }).reset_index()

    # Calculate additional metrics (changed: just add Money Out)
    monthly_totals['Net Movement'] = monthly_totals['Money In'] + monthly_totals['Money Out']
    # Changed: use absolute values for Money Out in average calculation
    monthly_totals['Average Transaction Value'] = (monthly_totals['Money In'] + monthly_totals['Money Out'].abs()) / \
                                                monthly_totals['Transaction Details']

    # Calculate month-over-month growth
    monthly_totals['Money In Growth %'] = monthly_totals['Money In'].pct_change() * 100
    monthly_totals['Money Out Growth %'] = monthly_totals['Money Out'].pct_change() * 100
    monthly_totals['Transaction Count Growth %'] = monthly_totals['Transaction Details'].pct_change() * 100

    # Calculate running totals
    monthly_totals['Running Total In'] = monthly_totals['Money In'].cumsum()
    monthly_totals['Running Total Out'] = monthly_totals['Money Out'].cumsum()
    monthly_totals['Running Net Movement'] = monthly_totals['Running Total In'] + monthly_totals['Running Total Out']  # Changed: just add

    # Calculate grand totals
    grand_totals = pd.DataFrame([{
        'Month': 'GRAND TOTAL',
        'Money In': monthly_totals['Money In'].sum(),
        'Money Out': monthly_totals['Money Out'].sum(),
        'Transaction Details': monthly_totals['Transaction Details'].sum(),
        'Net Movement': monthly_totals['Net Movement'].sum(),
        # Changed: use absolute values for Money Out in average calculation
        'Average Transaction Value': (monthly_totals['Money In'].sum() + monthly_totals['Money Out'].abs().sum()) /
                                   monthly_totals['Transaction Details'].sum(),
        'Money In Growth %': None,
        'Money Out Growth %': None,
        'Transaction Count Growth %': None,
        'Running Total In': monthly_totals['Money In'].sum(),
        'Running Total Out': monthly_totals['Money Out'].sum(),
        'Running Net Movement': monthly_totals['Net Movement'].sum()
    }])

    # Combine monthly totals with grand total
    final_monthly_totals = pd.concat([monthly_totals, grand_totals], ignore_index=True)

    return final_monthly_totals


def process_pdf(pdf_path):
    # Read all tables from all pages
    tables = tabula.read_pdf(
        pdf_path,
        pages='all',
        multiple_tables=True,
        lattice=True,
        guess=False,
        pandas_options={'header': None}
    )

    processed_tables = []

    for table in tables:
        if len(table.columns) >= 6:
            table.columns = ['Transaction Date', 'Value Date', 'Transaction Details',
                           'Money Out', 'Money In', 'Ledger Balance', 'Bank Reference Number']

            table = table.dropna(how='all')
            table = table[table['Transaction Date'].str.contains(r'\d{2}\.\d{2}\.\d{4}', na=False)]

            processed_tables.append(table)

    if processed_tables:
        final_df = pd.concat(processed_tables, ignore_index=True)
        final_df = final_df[final_df['Transaction Date'] != 'Transaction Date']

        # Sort by transaction date
        final_df['Transaction Date'] = pd.to_datetime(final_df['Transaction Date'], format='%d.%m.%Y')
        final_df = final_df.sort_values('Transaction Date')

        # Clean up monetary columns
        for col in ['Money Out', 'Money In', 'Ledger Balance']:
            final_df[col] = final_df[col].replace('[\$,]', '', regex=True)
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce')

        return final_df

    return None


def apply_excel_formatting(writer, df, summary_df, daily_totals_df, monthly_totals_df):
    """Apply enhanced Excel formatting to all sheets"""
    workbook = writer.book

    # Create format objects
    formats = {
        'money': workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'right'
        }),
        'percentage': workbook.add_format({
            'num_format': '0.0"%"',
            'align': 'right'
        }),
        'date': workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'align': 'center'
        }),
        'header': workbook.add_format({
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        }),
        'monthly': workbook.add_format({
            'bold': True,
            'bg_color': '#DCE6F1',
            'border': 1,
            'num_format': '#,##0.00'
        }),
        'grand_total': workbook.add_format({
            'bold': True,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'border': 1,
            'num_format': '#,##0.00'
        }),
        'text': workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True
        })
    }

    # Format Transactions sheet
    trans_worksheet = writer.sheets['Transactions']
    trans_worksheet.set_column('A:B', 12, formats['date'])
    trans_worksheet.set_column('C:C', 50, formats['text'])
    trans_worksheet.set_column('D:F', 15, formats['money'])
    trans_worksheet.set_column('G:G', 20, formats['text'])
    trans_worksheet.freeze_panes(1, 0)
    trans_worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    # Format Summary sheet
    summary_worksheet = writer.sheets['Summary']
    summary_worksheet.set_column('A:A', 30, formats['text'])
    summary_worksheet.set_column('B:B', 20, formats['money'])

    # Format Daily Totals sheet
    daily_worksheet = writer.sheets['Daily Totals']
    daily_worksheet.set_column('A:A', 12, formats['date'])
    daily_worksheet.set_column('B:E', 15, formats['money'])
    daily_worksheet.freeze_panes(1, 0)

    # Format Monthly Analysis sheet
    monthly_worksheet = writer.sheets['Monthly Analysis']
    monthly_worksheet.set_column('A:A', 15, formats['text'])
    monthly_worksheet.set_column('B:F', 15, formats['money'])
    monthly_worksheet.set_column('G:I', 15, formats['percentage'])
    monthly_worksheet.set_column('J:L', 18, formats['money'])
    monthly_worksheet.freeze_panes(1, 0)

    # Add conditional formatting for positive/negative values
    positive_format = workbook.add_format({'font_color': 'green', 'num_format': '#,##0.00'})
    negative_format = workbook.add_format({'font_color': 'red', 'num_format': '#,##0.00'})

    # Apply to Net Movement columns
    daily_worksheet.conditional_format('E2:E1048576', {
        'type': 'cell',
        'criteria': '>',
        'value': 0,
        'format': positive_format
    })
    daily_worksheet.conditional_format('E2:E1048576', {
        'type': 'cell',
        'criteria': '<',
        'value': 0,
        'format': negative_format
    })

    monthly_worksheet.conditional_format('E2:E1048576', {
        'type': 'cell',
        'criteria': '>',
        'value': 0,
        'format': positive_format
    })
    monthly_worksheet.conditional_format('E2:E1048576', {
        'type': 'cell',
        'criteria': '<',
        'value': 0,
        'format': negative_format
    })

    # Write headers
    for sheet, data in {
        'Transactions': df,
        'Summary': summary_df,
        'Daily Totals': daily_totals_df,
        'Monthly Analysis': monthly_totals_df
    }.items():
        worksheet = writer.sheets[sheet]
        for col_num, value in enumerate(data.columns.values):
            worksheet.write(0, col_num, value, formats['header'])

    # Format grand total row in Monthly Analysis
    last_row = len(monthly_totals_df)
    monthly_worksheet.set_row(last_row - 1, None, formats['grand_total'])

    return writer


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)

        try:
            # Process the PDF
            df = process_pdf(pdf_path)

            if df is not None:
                # Create summary and totals
                summary_df = create_summary(df)
                daily_totals_df = create_daily_totals(df)
                monthly_totals_df = create_monthly_totals(df)

                # Format Transaction Date back to string for Excel output
                df['Transaction Date'] = df['Transaction Date'].dt.strftime('%d.%m.%Y')

                # Save to Excel
                excel_filename = f"{filename.rsplit('.', 1)[0]}.xlsx"
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)

                # Create Excel writer object
                writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')

                # Write to different sheets
                df.to_excel(writer, index=False, sheet_name='Transactions')
                summary_df.to_excel(writer, index=False, sheet_name='Summary')
                daily_totals_df.to_excel(writer, index=False, sheet_name='Daily Totals')
                monthly_totals_df.to_excel(writer, index=False, sheet_name='Monthly Analysis')

                # Apply enhanced formatting
                writer = apply_excel_formatting(writer, df, summary_df, daily_totals_df, monthly_totals_df)

                # Close the Excel writer
                writer.close()

                # Clean up PDF file
                os.remove(pdf_path)

                # Return the Excel file
                return send_file(excel_path,
                               as_attachment=True,
                               download_name=excel_filename,
                               mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

            return 'Error processing PDF - No valid tables found'
        except Exception as e:
            return f'Error processing PDF: {str(e)}'
        finally:
            # Clean up uploaded file if it still exists
            if os.path.exists(pdf_path):
                os.remove(pdf_path)

    return 'Invalid file type'


if __name__ == '__main__':
    app.run(debug=True)