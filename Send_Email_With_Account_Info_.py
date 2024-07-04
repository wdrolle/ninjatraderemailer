import os
import time
import smtplib
import xlsxwriter
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, timedelta
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Email credentials
USERNAME = os.getenv('EUSERNAME')
PASSWORD = os.getenv('PASSWORD')

# Email settings
RECIPIENT_EMAIL = 'w.d.rolle@gmail.com'
EMAIL_SUBJECT = 'Profits Report'

# Paths to the HTML template and CSS file
HTML_TEMPLATE_PATH = r'C:\Users\Administrator\.csv\apps\templates\email_template.html'
CSS_FILE_PATH = r'C:\Users\Administrator\.csv\apps\static\assets\css\styles.css'

# Excel file path
EXCEL_FILE_PATH = r'C:\Users\Administrator\.csv\profits\profits_report.xlsx'
ACCOUNT_CSV_PATH = r'C:\Users\Administrator\.csv\profits\account_data.csv'

# Directory to watch
WATCH_DIRECTORY = r'C:\Users\Administrator\.csv\profits'

# Update headers and column names
daily_headers = [
    'Time', 'Account', 'Account Balance', 'Profit', 'Net Change', 'Unrealized P/L', 
    'Total Cash Balance', 'Realized P/L', 'Gross Realized P/L', 'Buying Power', 
    'Cash Value', 'Commission', 'Fee'
]

# Generate HTML table from DataFrame
def generate_html_table(df):
    html = ""
    for i, row in df.iterrows():
        html += "<tr>"
        for col, val in row.items():
            if col in ['Account Balance', 'Unrealized P/L', 'Total Cash Balance', 'Realized P/L', 'Gross Realized P/L', 'Buying Power', 'Cash Value', 'Commission', 'Fee']:
                html += f"<td>${val:,.2f}</td>"
            elif col in ['Profit', 'Net Change']:
                html += f"<td>{val:.2f}%</td>"
            else:
                html += f"<td>{val}</td>"
        html += "</tr>"
    return html

# Load account data from CSV
def load_account_data():
    try:
        df = pd.read_csv(ACCOUNT_CSV_PATH)
        print("Data loaded from CSV:")
        print(df.head())
        df['Time'] = pd.to_datetime(df['Time'])
        df.sort_values(by='Time', ascending=False, inplace=True)
        df['Time'] = df['Time'].dt.strftime('%m/%d/%y %I:%M:%S %p')
        df.fillna(0, inplace=True)  # Fill NaN values with 0
        
        # Ensure all necessary columns are present
        necessary_columns = ['Account Balance', 'Net Change', 'Unrealized P/L', 'Total Cash Balance', 'Realized P/L', 'Gross Realized P/L', 'Buying Power', 'Cash Value', 'Commission', 'Fee', 'Profit']
        for col in necessary_columns:
            if col not in df.columns:
                df[col] = 0

        return df
    except Exception as e:
        print(f"Error loading account data: {e}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error

# Generate the Excel file
def generate_excel_file(data):
    if data.empty:
        print("No data to process. DataFrame is empty.")
        return

    print("Generating Excel file")
    workbook = xlsxwriter.Workbook(EXCEL_FILE_PATH)
    header_format = workbook.add_format({'bold': True, 'bg_color': '#0000FF', 'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'border_color': '#D9D9D9'})
    currency_format = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'})
    percentage_format = workbook.add_format({'num_format': '0.00%', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'})
    general_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'})
    border_format = workbook.add_format({'border_color': '#D9D9D9', 'border': 1})

    def apply_conditional_formatting(sheet, range, type='currency'):
        format_red = workbook.add_format({'bg_color': '#FFA48F', 'num_format': '$#,##0.00', 'border': 1, 'border_color': '#D9D9D9'})
        if type == 'percent':
            format_red = workbook.add_format({'bg_color': '#FFA48F', 'num_format': '0.00%', 'border': 1, 'border_color': '#D9D9D9'})
        sheet.conditional_format(range, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': format_red})
        sheet.conditional_format(range, {'type': 'cell', 'criteria': '==', 'value': 0, 'format': workbook.add_format({'num_format': '""', 'border': 1, 'border_color': '#D9D9D9'})})

    def apply_custom_conditional_formatting(sheet, row):
        account = data.loc[row, 'Account']
        balance = data.loc[row, 'Account Balance']
        if (account.startswith('Sim') or account.startswith('APEX')) and balance < 100000:
            sheet.write(row + 1, 2, balance, workbook.add_format({'bg_color': '#FFA48F', 'num_format': '$#,##0.00', 'border': 1, 'border_color': '#D9D9D9'}))
        else:
            sheet.write(row + 1, 2, balance, currency_format)

    # Daily sheet
    daily_sheet = workbook.add_worksheet('Daily')
    daily_sheet.hide_gridlines(2)
    daily_sheet.write_row('A1', daily_headers, header_format)
    daily_sheet.set_column('A:A', 20)
    daily_sheet.set_column('B:N', 14)

    for index, row in data.iterrows():
        daily_sheet.write_row(f'A{index+2}', [0 if pd.isna(row[header]) else row[header] for header in daily_headers], border_format)
        for col_num, header in enumerate(daily_headers):
            if header in ['Account Balance', 'Unrealized P/L', 'Total Cash Balance', 'Realized P/L', 'Gross Realized P/L', 'Buying Power', 'Cash Value', 'Commission', 'Fee']:
                daily_sheet.write(index + 1, col_num, row[header], currency_format)
            elif header in ['Profit', 'Net Change']:
                daily_sheet.write(index + 1, col_num, row[header], percentage_format)
            else:
                daily_sheet.write(index + 1, col_num, row[header], general_format)

        apply_custom_conditional_formatting(daily_sheet, index)

    apply_conditional_formatting(daily_sheet, 'C2:C{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'D2:D{}'.format(len(data)+1), 'percent')
    apply_conditional_formatting(daily_sheet, 'E2:E{}'.format(len(data)+1), 'percent')
    apply_conditional_formatting(daily_sheet, 'F2:F{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'G2:G{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'H2:H{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'I2:I{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'J2:J{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'K2:K{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'L2:L{}'.format(len(data)+1), 'currency')
    apply_conditional_formatting(daily_sheet, 'M2:M{}'.format(len(data)+1), 'currency')

    # Function to get the start and end dates of the current week (Monday to Friday)
    def get_current_week():
        today = datetime.now()
        start_of_week = today - timedelta(days=today.weekday())
        end_of_week = start_of_week + timedelta(days=4)
        return start_of_week.strftime('%m-%d-%y'), end_of_week.strftime('%m-%d-%y')

    # Function to get the current month
    def get_current_month():
        return datetime.now().strftime('%B')

    # Function to get the start and end months of the current quarter
    def get_current_quarter():
        current_month = datetime.now().month
        if current_month in [1, 2, 3]:
            return 'January - March'
        elif current_month in [4, 5, 6]:
            return 'April - June'
        elif current_month in [7, 8, 9]:
            return 'July - September'
        else:
            return 'October - December'

    # Function to get the current year
    def get_current_year():
        return datetime.now().strftime('%Y')

    # Dynamically generate the periods dictionary
    start_of_week, end_of_week = get_current_week()
    periods = {
        'Weekly': f'{start_of_week} - {end_of_week}',
        'Monthly': get_current_month(),
        'Quarterly': get_current_quarter(),
        'Yearly': get_current_year()
    }

    for period, name in periods.items():
        sheet = workbook.add_worksheet(period)
        headers = ['Time', 'Account', 'Account Balance', 'Net Change', 'Profit']
        sheet.write_row('A1', headers, header_format)
        sheet.set_column('A:A', 20)
        sheet.set_column('B:E', 14)

        prev_balance = data['Account Balance'].iloc[0] if not data.empty else 0  # Ensure prev_balance is set correctly
        for index, row in data.iterrows():
            change_percent = (row['Account Balance'] - prev_balance) / prev_balance if prev_balance != 0 else 0
            prev_balance = row['Account Balance']
            total_profit = row['Profit']  # or some other calculation if needed
            sheet.write_row(f'A{index+2}', [name, row['Account'], row['Account Balance'], change_percent, total_profit], border_format)
            for col_num, header in enumerate(headers):
                if header in ['Account Balance', 'Profit']:
                    sheet.write(index + 1, col_num, row[header], currency_format)
                elif header == 'Net Change':
                    sheet.write(index + 1, col_num, change_percent, percentage_format)

        apply_conditional_formatting(sheet, 'C2:C{}'.format(len(data)+1), 'currency')
        apply_conditional_formatting(sheet, 'D2:D{}'.format(len(data)+1), 'percent')
        apply_conditional_formatting(sheet, 'E2:E{}'.format(len(data)+1))

    workbook.close()
    print("Excel file generated successfully.")

# Send email with attachment
def send_email():
    try:
        msg = MIMEMultipart()
        msg['From'] = USERNAME
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = EMAIL_SUBJECT

        # Load data from the Excel file to generate the email body
        dfs = {
            'Daily': pd.read_excel(EXCEL_FILE_PATH, sheet_name='Daily').head(5),
            'Weekly': pd.read_excel(EXCEL_FILE_PATH, sheet_name='Weekly').head(5),
            'Monthly': pd.read_excel(EXCEL_FILE_PATH, sheet_name='Monthly').head(5),
            'Quarterly': pd.read_excel(EXCEL_FILE_PATH, sheet_name='Quarterly').head(5),
            'Yearly': pd.read_excel(EXCEL_FILE_PATH, sheet_name='Yearly').head(5)
        }

        # List of columns to consider for deduplication (excluding 'Time')
        columns_to_check = ['Account', 'Account Balance', 'Profit', 'Net Change', 'Unrealized P/L', 
                            'Total Cash Balance', 'Realized P/L', 'Gross Realized P/L', 
                            'Buying Power', 'Cash Value', 'Commission', 'Fee']

        # Remove duplicates for each DataFrame
        for key in dfs.keys():
            available_columns = [col for col in columns_to_check if col in dfs[key].columns]
            dfs[key] = dfs[key].drop_duplicates(subset=available_columns)

        # Extract DataFrames back from the dictionary
        df_daily = dfs['Daily']
        df_weekly = dfs['Weekly']
        df_monthly = dfs['Monthly']
        df_quarterly = dfs['Quarterly']
        df_yearly = dfs['Yearly']

        # Summary of today's change in Profit and Loss
        today = datetime.now().strftime('%m/%d/%y')
        todays_data = df_daily[df_daily['Time'].str.startswith(today)]
        if not todays_data.empty:
            summary_lines = ["<strong>Today's Change in Percentage by Account</strong><br>"]
            accounts = todays_data['Account'].unique()
            for account in accounts:
                account_data = todays_data[todays_data['Account'] == account]
                if len(account_data) > 1:
                    start_balance = account_data.iloc[-1]['Account Balance']
                    end_balance = account_data.iloc[0]['Account Balance']
                    pnl_change = ((end_balance - start_balance) / start_balance) * 100
                    previous_day_data = df_daily[(df_daily['Account'] == account) & (~df_daily['Time'].str.startswith(today))]
                    if not previous_day_data.empty:
                        previous_day_end_balance = previous_day_data.iloc[0]['Account Balance']
                        previous_day_change = ((start_balance - previous_day_end_balance) / previous_day_end_balance) * 100
                        summary_lines.append(f"Account {account}: {pnl_change:.2f}% change from yesterday's {previous_day_change:.2f}% at the end of the day. This change reflects a {pnl_change - previous_day_change:.2f}% change.<br>")
                    else:
                        summary_lines.append(f"Account {account}: {pnl_change:.2f}% change today.<br>")
            summary_text = "".join(summary_lines)
        else:
            previous_data = df_daily[~df_daily['Time'].str.startswith(today)]
            if not previous_data.empty:
                last_change_date = previous_data.iloc[0]['Time']
                last_change_balance = previous_data.iloc[0]['Account Balance']
                summary_text = f"No change in percentage was made today.<br>The last change was on {last_change_date} with a balance of ${last_change_balance:,.2f} that reflected a {previous_data.iloc[0]['Net Change']:,.2f}%.<br>Between {previous_data.iloc[-1]['Time']} and {today}, no change in percentage was made."
            else:
                summary_text = 'No data available for today to show percentage differences between days.'

        # Read the HTML template file
        with open(HTML_TEMPLATE_PATH, 'r') as file:
            html_template = file.read()

        # Read the CSS file
        with open(CSS_FILE_PATH, 'r') as file:
            css_content = file.read()

        # Replace placeholders with actual data
        email_body = html_template.replace(
            '{{ styles }}', css_content
        ).replace(
            '{{ summary }}', summary_text
        ).replace(
            '{{ daily_table }}', generate_html_table(df_daily.sort_values(by='Time', ascending=True))
        ).replace(
            '{{ weekly_table }}', generate_html_table(df_weekly.sort_values(by='Time', ascending=True))
        ).replace(
            '{{ monthly_table }}', generate_html_table(df_monthly.sort_values(by='Time', ascending=True))
        ).replace(
            '{{ quarterly_table }}', generate_html_table(df_quarterly.sort_values(by='Time', ascending=True))
        ).replace(
            '{{ yearly_table }}', generate_html_table(df_yearly.sort_values(by='Time', ascending=True))
        ).replace(
            '{{ year }}', str(datetime.now().year)
        )

        msg.attach(MIMEText(email_body, 'html'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(EXCEL_FILE_PATH, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(EXCEL_FILE_PATH)}"')
        msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(USERNAME, PASSWORD)
            server.sendmail(USERNAME, RECIPIENT_EMAIL, msg.as_string())
        print("Email sent successfully.")
    except Exception as e:
        print(f"Error sending email: {e}")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(EXCEL_FILE_PATH, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(EXCEL_FILE_PATH)}"')
        msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(USERNAME, PASSWORD)
            server.sendmail(USERNAME, RECIPIENT_EMAIL, msg.as_string())
        print("Email sent successfully.")
    except Exception as e:
        print(f"Error sending email: {e}")


# Watchdog event handler
class Watcher(FileSystemEventHandler):
    def on_modified(self, event):
        if not event.is_directory and event.src_path.endswith('.csv'):
            print(f"Detected change in file: {event.src_path}")  # Debug statement
            data = load_account_data()
            if not data.empty:
                generate_excel_file(data)
                send_email()
                print(f"Email sent for file: {event.src_path}")  # Debug statement
            else:
                print("No data found to process and send email.")

if __name__ == "__main__":
    event_handler = Watcher()
    observer = Observer()
    observer.schedule(event_handler, path=WATCH_DIRECTORY, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()