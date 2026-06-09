import os
import datetime
import logging
from dotenv import load_dotenv

from .plaid_client import PlaidClient
from .sheets_client import SheetsClient
from .processor import TransactionProcessor
from .sync_utils import determine_sync_window

# Setup logging
log_file_path = os.path.join(os.path.dirname(__file__), '../../automation.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def main():
    """Main function to orchestrate the automation."""
    logger.info("--- Starting Bank Transaction Sync ---")

    # 1. Load configuration
    load_dotenv() 
    
    PLAID_CLIENT_ID = os.getenv('PLAID_CLIENT_ID')
    PLAID_SECRET = os.getenv('PLAID_SECRET')
    PLAID_ENV = os.getenv('PLAID_ENV', 'sandbox')
    PLAID_ACCESS_TOKEN = os.getenv('PLAID_ACCESS_TOKEN')
    GOOGLE_SPREADSHEET_ID = os.getenv('GOOGLE_SPREADSHEET_ID')
    GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME', 'Transactions')

    if not all([PLAID_CLIENT_ID, PLAID_SECRET, PLAID_ACCESS_TOKEN, GOOGLE_SPREADSHEET_ID]):
        logger.error("Missing critical configuration in .env file. Exiting.")
        return

    # Split tokens by comma to support multiple institutions
    access_tokens = [t.strip() for t in PLAID_ACCESS_TOKEN.split(',') if t.strip()]
    logger.info(f"Detected {len(access_tokens)} institution(s) to sync.")
    logger.info(
        "Sync target -> spreadsheet_id=%s | sheet_name=%s | plaid_env=%s",
        GOOGLE_SPREADSHEET_ID,
        GOOGLE_SHEET_NAME,
        PLAID_ENV,
    )

    # 2. Initialize clients
    plaid_client = PlaidClient(PLAID_CLIENT_ID, PLAID_SECRET, PLAID_ENV)
    sheets_client = SheetsClient(GOOGLE_SPREADSHEET_ID, GOOGLE_SHEET_NAME)
    processor = TransactionProcessor()

    # 3. Authenticate Google Sheets
    logger.info("Authenticating with Google Sheets...")
    try:
        google_creds = sheets_client.authenticate()
    except Exception as e:
        logger.error(f"Failed to authenticate with Google Sheets: {e}")
        return

    # 4. Determine Date Range (Incremental Sync)
    latest_date_str = sheets_client.get_latest_transaction_date(google_creds)
    logger.info("Latest detected transaction date -> %s", latest_date_str or "none")

    start_date, end_date, sync_mode = determine_sync_window(latest_date_str)
    logger.info(
        "Chosen sync mode -> %s | window=%s -> %s",
        sync_mode,
        start_date.isoformat(),
        end_date.isoformat(),
    )
    if sync_mode == 'incremental':
        logger.info(f"Incremental sync detected. Last transaction was {latest_date_str}. Starting from {start_date.isoformat()}")
    else:
        if latest_date_str:
            logger.warning(f"Could not parse latest date '{latest_date_str}'. Defaulting to 2 years.")
        logger.info("No previous transactions found. Performing deep 2-year history sync.")

    # 5. Retrieve Existing Transaction IDs (for deduplication)
    logger.info("Reading existing transaction IDs from Google Sheet...")
    existing_ids = sheets_client.get_existing_ids(google_creds)
    logger.info("Loaded %s existing transaction IDs from Google Sheet.", len(existing_ids))
    if existing_ids and sync_mode == 'bootstrap':
        logger.warning(
            "Existing IDs were found while sync mode is bootstrap. Verify the target sheet date column formatting and spreadsheet target."
        )

    # 6. Retrieve Transactions from all Plaid tokens
    all_new_data = []
    total_tokens = len(access_tokens)
    for index, token in enumerate(access_tokens, start=1):
        logger.info(f"Retrieving transactions for institution {index} of {total_tokens}...")
        account_labels = plaid_client.get_account_label_map(token)
        logger.info("Resolved %s account label(s) for institution %s.", len(account_labels), index)
        transactions = plaid_client.get_transactions(token, start_date, end_date)

        if transactions is None:
            logger.error(f"Failed to retrieve transactions for institution {index}. Skipping.")
            continue

        # 7. Parse Transactions
        logger.info("Parsing data for institution...")
        parsed_data = processor.parse_plaid_transactions(transactions, existing_ids, account_labels)
        all_new_data.extend(parsed_data)
        existing_ids.update(row[0] for row in parsed_data)

    # 8. Update Google Sheet
    if not all_new_data:
        logger.info("No new transactions found across all institutions. Sheet is up to date.")
    else:
        logger.info(f"Appending {len(all_new_data)} total new transactions to Google Sheet...")
        success = sheets_client.append_data(google_creds, all_new_data) 
        if success:
            logger.info("Google Sheet update successful.")
        else:
            logger.error("Failed to update Google Sheet.")

    logger.info("--- Bank Transaction Sync Finished ---")

if __name__ == '__main__':
    main()
