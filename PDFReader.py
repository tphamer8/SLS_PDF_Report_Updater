import fitz
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os


def delete_duplicates(spreadsheet):
    sheet = spreadsheet.worksheet('Main')
    all_rows = sheet.get_all_values()

    header = all_rows[0]
    data = all_rows[1:]

    # Track by URL: {url: [(row_index, row)]}
    url_map = {}
    # Track exact row duplicates: {tuple(row): [row_index]}
    row_map = {}

    for i, row in enumerate(data, start=2):  # sheet rows start at 1, data starts at row 2
        url = row[0].strip()
        status = row[3].strip() if len(row) > 3 else ""

        # Group by URL
        url_map.setdefault(url, []).append((i, row))

        # Group by exact row content
        row_tuple = tuple(cell.strip() for cell in row)
        row_map.setdefault(row_tuple, []).append(i)

    rows_to_delete = set()

    # Handle "Remediated" vs non-Remediated duplicates (by URL)
    for url, entries in url_map.items():
        if len(entries) > 1:
            has_remediated = any(r[1][3].strip().lower() == "remediated" for r in entries if len(r[1]) > 3)
            if has_remediated:
                for row_index, row in entries:
                    if len(row) < 4 or row[3].strip().lower() != "remediated":
                        rows_to_delete.add(row_index)

    # Handle exact duplicates
    for row_vals, indices in row_map.items():
        if len(indices) > 1:
            # Keep one, delete the rest
            rows_to_delete.update(indices[1:])

    # Delete from bottom to top
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    print(f"✅ Deleted {len(rows_to_delete)} duplicate rows.")

# Write data to sheet
def write_to_sheet(spreedsheet, rows):
    main = spreadsheet.worksheet('Main')
    data_to_append = [[url, path, int(clicks), "Pending"] for url, path, clicks in rows]
    for row in data_to_append:
        main.append_row(row, value_input_option='USER_ENTERED')

# Read the PDF
def extract_pdf_report_data(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()

    # Extract rows
    pattern = r"(https://law\.stanford\.edu[^\s]+)\s+(/[^\s]+)\s+(\d+)"
    matches = re.findall(pattern, text)
    return matches

# Authenticate Google Sheets
def authenticate_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client

# Open the spreadsheet
def open_spreadsheet(client, spreadsheet_name):
    spreadsheet = client.open(spreadsheet_name)
    return spreadsheet

# Main
if __name__ == '__main__':
    client = authenticate_google_sheet()
    spreadsheet_name = 'PDF Remediate Report List'  # Replace with your spreadsheet name
    spreadsheet = open_spreadsheet(client, spreadsheet_name)

    pdf_path = input("Enter the name of the PDF report file (with extension): ").strip()

    if not os.path.isfile(pdf_path):
        print(f"❌ File not found: {pdf_path}")
        exit(1)

    rows = extract_pdf_report_data(pdf_path) # Replace with PDF file name
    write_to_sheet(spreadsheet, rows)

    os.remove(pdf_path)
    print(f"✅ Uploaded data and deleted: {pdf_path}")
    delete_duplicates(spreadsheet)
    print(f"✅ Deleted Duplicate Entries")


