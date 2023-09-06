import mailbox
from bs4 import BeautifulSoup
import pandas as pd
import os
import chardet

# Path to your mbox file
mbox_file = 'path to mbox file here'

# CSV folder
csv_folder = 'csv_data/'
os.makedirs(csv_folder, exist_ok=True)  # Create folder for CSV files

def should_skip_row(row):
    return 'work authorization start date' in row or 'work authorization end date' in row or 'work   authorization start date' in row or 'work   authorization end date' in row or 'i 9 status' in row or 'i 9 completion date' in row

def get_tax_term(message):
    tax_terms = ['c2c', 'w2', '1099']
    # Check if any of the tax terms are in the subject or body of the email
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            if content_type == 'text/html':
                payload = part.get_payload(decode=True)
                encoding = chardet.detect(payload)['encoding']
                payload = payload.decode(encoding)
                soup = BeautifulSoup(payload, 'html.parser')
                break
    else:
        payload = message.get_payload(decode=True)
        encoding = chardet.detect(payload)['encoding']
        payload = payload.decode(encoding)
        soup = BeautifulSoup(payload, 'html.parser')
    
    if any(term in message['subject'].lower() or term in soup.get_text().lower() for term in tax_terms):
        return [term for term in tax_terms if term in message['subject'].lower() or term in soup.get_text().lower()][0]
    # Check if any of the tax terms are in the tables of the email
    for table in soup.find_all('table'):
        rows = table.find_all('tr')
        for row in rows:
            cells = row.find_all(['td', 'th'])  # Include table headers (th) as well
            row_data = [cell.get_text(strip=True).lower() for cell in cells]
            if any(term in row_data for term in tax_terms):
                return [term for term in tax_terms if term in row_data][0]
    return None

def has_passthrough(message):
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            if content_type == 'text/html':
                payload = part.get_payload(decode=True)
                encoding = chardet.detect(payload)['encoding']
                payload = payload.decode(encoding)
                soup = BeautifulSoup(payload, 'html.parser')
                break
    else:
        payload = message.get_payload(decode=True)
        encoding = chardet.detect(payload)['encoding']
        payload = payload.decode(encoding)
        soup = BeautifulSoup(payload, 'html.parser')
    
    return 'passthrough' in soup.get_text().lower()

# Set the path to the file that keeps track of the last processed email
last_processed_email_path = 'last_processed.txt'

# Read the last processed email from the file
last_processed_email_index = None
if os.path.exists(last_processed_email_path):
    with open(last_processed_email_path, 'r') as f:
        last_processed_email_index = int(f.read().strip())

# Set last_processed_email_index to 0 if it's None
if last_processed_email_index is None:
    last_processed_email_index = 0

# Open the mbox file
mbox = mailbox.mbox(mbox_file)

# Loop through emails in the mbox starting from the last processed email
for mail_index in range(last_processed_email_index, len(mbox)):
    message = mbox[mail_index]
    print(f"Processing email {mail_index + 1} of {len(mbox)}")
    print('found mail')
    tables = []
    
    # Extract tables from email body
    if message.is_multipart():
        for part in message.walk():
            content_type = part.get_content_type()
            if content_type == 'text/html':
                payload = part.get_payload(decode=True)
                encoding = chardet.detect(payload)['encoding']
                payload = payload.decode(encoding)
                soup = BeautifulSoup(payload, 'html.parser')
                break
    else:
        payload = message.get_payload(decode=True)
        encoding = chardet.detect(payload)['encoding']
        payload = payload.decode(encoding)
        soup = BeautifulSoup(payload, 'html.parser')
    
    all_tables = soup.find_all('table')
    print(f"Found {len(all_tables)} tables")
    for table_index, table in enumerate(all_tables):
        print(f"Processing table {table_index + 1}")
        rows = table.find_all('tr')
        table_data = []
        for row_index, row in enumerate(rows):
            print(f"Processing row {row_index + 1}")
            print('found table')
            cells = row.find_all(['td', 'th'])  # Include table headers (th) as well
            row_data = [cell.get_text(strip=True).lower().replace('\n', ' ').replace('<br>', ' ') for cell in cells]
            if should_skip_row(row_data):
                continue
            table_data.append(row_data)
        tables.append(pd.DataFrame(table_data))
    
    tax_term = get_tax_term(message)
    if tax_term:
        print(f"Found tax term: {tax_term}")
        for table_df in tables:
            new_row = ['tax term', tax_term] + [None] * (len(table_df.columns) - 2)
            table_df.loc[len(table_df)] = new_row
    
    passthrough = has_passthrough(message)
    print(f"Found passthrough: {passthrough}")
    for table_df in tables:
        new_row = ['passthrough', 'yes' if passthrough else 'no'] + [None] * (len(table_df.columns) - 2)
        table_df.loc[len(table_df)] = new_row
    
    for table_index, table_df in enumerate(tables):
        df_filename = f'{csv_folder}table_{mail_index}_{table_index}.csv'
        table_df.to_csv(df_filename, index=False)
        print(f"Saved CSV file: {df_filename}")
    
    # Update the last processed email
    with open(last_processed_email_path, 'w') as f:
        f.write(str(mail_index))
        
        
print('Script Complete')
