import os

from dotenv import load_dotenv
from googleapiclient.discovery import build

from .plaid_client import PlaidClient
from .sheets_client import SheetsClient


def _load_tokens():
    token_value = os.getenv('PLAID_ACCESS_TOKEN') or ''
    return [token.strip() for token in token_value.split(',') if token.strip()]


def build_label_map(plaid_client, access_tokens):
    labels = {}
    for token in access_tokens:
        labels.update(plaid_client.get_account_label_map(token))
    return labels


def main():
    load_dotenv('.env')

    client_id = os.getenv('PLAID_CLIENT_ID')
    secret = os.getenv('PLAID_SECRET')
    env = os.getenv('PLAID_ENV', 'sandbox')
    spreadsheet_id = os.getenv('GOOGLE_SPREADSHEET_ID')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', 'Transactions').strip("'")
    access_tokens = _load_tokens()

    if not all([client_id, secret, spreadsheet_id, access_tokens]):
        print('Missing Plaid or Google Sheets configuration in .env.')
        return

    plaid_client = PlaidClient(client_id, secret, env)
    labels = build_label_map(plaid_client, access_tokens)
    if not labels:
        print('No Plaid account labels were resolved. Nothing to migrate.')
        return

    sheets_client = SheetsClient(spreadsheet_id, sheet_name)
    creds = sheets_client.authenticate()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    range_name = f"'{sheet_name}'!F2:F"
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    rows = result.get('values', [])
    if not rows:
        print('No Account column values found. Nothing to migrate.')
        return

    changed = 0
    updated_rows = []
    for row in rows:
        current = str(row[0]).strip() if row else ''
        updated = labels.get(current, current)
        if updated != current:
            changed += 1
        updated_rows.append([updated])

    if changed == 0:
        print('No raw Plaid account IDs matched current linked accounts. Nothing changed.')
        return

    sheet.values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body={'values': updated_rows}
    ).execute()
    print(f'Updated {changed} Account cell(s) in {sheet_name}!F:F.')


if __name__ == '__main__':
    main()
