import logging
import os
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from .sync_utils import coerce_sheet_date_value, extract_existing_ids

logger = logging.getLogger(__name__)

class SheetsClient:
    def __init__(self, spreadsheet_id, sheet_name='Transactions'):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']
        self.token_path = os.path.join(os.path.dirname(__file__), '../../token.json')
        self.credentials_path = os.path.join(os.path.dirname(__file__), '../../credentials.json')

    def authenticate(self):
        """Authenticates with Google Sheets API using OAuth 2.0."""
        creds = None
        if os.path.exists(self.token_path):
            try:
                creds = Credentials.from_authorized_user_file(self.token_path, self.scopes)
            except Exception as e:
                logger.error(f"Error loading token.json: {e}. Please check the file or delete it to re-authenticate.")
                creds = None 

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    logger.info("Refreshing Google API token.")
                    creds.refresh(Request())
                except Exception as e:
                    logger.error(f"Error refreshing token: {e}. Re-authentication required.")
                    if os.path.exists(self.token_path):
                        os.remove(self.token_path)
                    flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, self.scopes)
                    creds = flow.run_local_server(port=0) 
            else:
                logger.info("Google credentials not found or invalid, starting authentication flow...")
                if not os.path.exists(self.credentials_path):
                     logger.error(f"Credentials file not found at: {self.credentials_path}")
                     raise FileNotFoundError(f"Missing credentials.json at {self.credentials_path}")
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, self.scopes)
                    creds = flow.run_local_server(port=0)
                except Exception as e:
                    logger.error(f"Failed to run authentication flow: {e}")
                    raise

            # Save the credentials for the next run
            try:
                with open(self.token_path, 'w') as token:
                    token.write(creds.to_json())
                logger.info("Google credentials saved to token.json.")
            except Exception as e:
                logger.error(f"Failed to save token.json: {e}")

        return creds

    def append_data(self, creds, values):
        """Appends data to the configured Google Sheet."""
        if not values:
            logger.info("No new transaction data to append.")
            return True

        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()

            range_to_append = f"'{self.sheet_name}'!A:G"

            body = {
                'values': values
            }
            
            result = sheet.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=range_to_append,
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()

            logger.info(f"{result.get('updates').get('updatedCells')} cells appended to sheet.")
            return True

        except HttpError as err:
            logger.error(f"Google Sheets API Error: {err}")
            return False
        except Exception as e:
            logger.error(f"An unexpected error occurred updating Google Sheets: {e}")
            return False

    def get_existing_ids(self, creds, id_column_index=0):
        """Reads the ID column from the sheet to prevent duplicates."""
        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()

            range_name = f"'{self.sheet_name}'!A2:A"
            result = sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])
            if not values:
                return set()

            return extract_existing_ids(values)

        except HttpError as err:
            logger.error(f"Google Sheets API Error: {err}")
            return set()
        except Exception as e:
            logger.error(f"An unexpected error occurred reading IDs: {e}")
            return set()

    def get_latest_transaction_date(self, creds):
        """Reads the Date column (B) to find the most recent transaction date."""
        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()

            range_name = f"'{self.sheet_name}'!B2:B"
            result = sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueRenderOption='UNFORMATTED_VALUE',
                dateTimeRenderOption='SERIAL_NUMBER'
            ).execute()

            values = result.get('values', [])
            if not values:
                return None

            dates = [coerce_sheet_date_value(row[0]) for row in values if row]
            dates = [date for date in dates if date]
            if not dates:
                return None

            return max(dates).isoformat()

        except HttpError as err:
            logger.error(f"Google Sheets API Error: {err}")
            return None
        except Exception as e:
            logger.error(f"An unexpected error occurred reading latest date: {e}")
            return None
