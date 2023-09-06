import win32com.client as win32
import os
import pandas as pd
from bs4 import BeautifulSoup

print("Starting script...")

# Initialize Outlook
print("Initializing Outlook...")
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Inbox folder (6)
print(f"Found {len(inbox.Items)} emails in inbox.")

# Path to your Outlook OST file
ost_file_path = r"path to OST File Here"

# Read the last processed index if available
try:
    with open('last_processed_index.txt', 'r') as index_file:
        last_processed_index = int(index_file.read())
        print(f"Read last processed index: {last_processed_index}")
except FileNotFoundError:
    last_processed_index = 0
    print("Last processed index file not found. Starting from 0.")

csv_folder = 'csv_data/'
os.makedirs(csv_folder, exist_ok=True)  # Create folder for CSV files
print(f"CSV folder: {csv_folder}")

def should_skip_row(row):
    return 'work authorization start date' in row or 'work authorization end date' in row

def get_tax_term(mail):
    tax_terms = ['c2c', 'w2', '1099']
    # Check if any of the tax terms are in the subject or body of the email
    if any(term in mail.Subject.lower() or term in mail.Body.lower() for term in tax_terms):
        return [term for term in tax_terms if term in mail.Subject.lower() or term in mail.Body.lower()][0]
    # Check if any of the tax terms are in the tables of the email
    if mail.BodyFormat == 2:  # HTML format
        soup = BeautifulSoup(mail.HTMLBody, 'html.parser')
        for table in soup.find_all('table'):
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all(['td', 'th'])  # Include table headers (th) as well
                row_data = [cell.get_text(strip=True).lower().replace('\n', ' ') for cell in cells]
                if any(term in row_data for term in tax_terms):
                    return [term for term in tax_terms if term in row_data][0]
    return None

# Loop through emails in the OST file
for mail_index, mail in enumerate(inbox.Items):
    print(f"Processing email {mail_index + 1} of {len(inbox.Items)}...")
    if mail_index < last_processed_index:
        continue

    tables = []
    
    # Extract tables from email body
    if mail.BodyFormat == 2:  # HTML format
        print("Email is in HTML format. Parsing tables...")
        soup = BeautifulSoup(mail.HTMLBody, 'html.parser')
        for table in soup.find_all('table'):
            rows = table.find_all('tr')
            table_data = []
            for row in rows:
                cells = row.find_all(['td', 'th'])  # Include table headers (th) as well
                row_data = [cell.get_text(strip=True).lower().replace('\n', ' ') for cell in cells]
                if should_skip_row(row_data):
                    continue
                table_data.append(row_data)
            tables.append(pd.DataFrame(table_data))
            print(f"Found table with shape {tables[-1].shape}")
    
    tax_term = get_tax_term(mail)
    if tax_term:
        print(f"Found tax term '{tax_term}'")
        for table_df in tables:
            new_row = ['tax term', tax_term] + [None] * (len(table_df.columns) - 2)
            table_df.loc[len(table_df)] = new_row

    
    for table_index, table_df in enumerate(tables):
        df_filename = f'{csv_folder}table_{mail_index}_{table_index}.csv'
        table_df.to_csv(df_filename, index=False)
        print(f"Saved table to CSV file: {df_filename}")
    
    # Update last processed index
    with open('last_processed_index.txt', 'w') as index_file:
        index_file.write(str(mail_index + 1))
        print(f"Updated last processed index to {mail_index + 1}")

print("Script completed.")
