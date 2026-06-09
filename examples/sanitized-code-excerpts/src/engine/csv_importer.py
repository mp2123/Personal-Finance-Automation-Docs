import argparse
import csv
import hashlib
import os
import re
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

from .sheets_client import SheetsClient


DEFAULT_MANUAL_INCOME_CATEGORY = 'Income > Manual Income'
DEFAULT_MANUAL_INCOME_ACCOUNT = 'Manual Income'
TRANSACTION_WIDTH = 7


class ManualIncomeImportError(Exception):
    pass


def parse_money(value):
    text = str(value or '').strip()
    if not text:
        raise ManualIncomeImportError('Amount is required.')
    negative = text.startswith('(') and text.endswith(')')
    normalized = text.replace('$', '').replace(',', '').replace('+', '').strip()
    if negative:
        normalized = '-' + normalized.strip('()')
    try:
        return round(float(normalized), 2)
    except ValueError as exc:
        raise ManualIncomeImportError(f'Invalid amount: {value}') from exc


def normalize_date(value):
    text = re.sub(r'\s+', ' ', str(value or '').strip())
    if not text:
        raise ManualIncomeImportError('Date is required.')

    formats = ['%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y', '%b %d %Y', '%B %d %Y']
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    raise ManualIncomeImportError(f'Invalid date: {value}')


def normalize_label(value):
    return re.sub(r'\s+', ' ', str(value or '').strip())


def normalize_id_component(value):
    return normalize_label(value).lower()


def is_forbidden_income_category(category):
    normalized = normalize_id_component(category).replace('_', ' ')
    forbidden_patterns = [
        'transfer in',
        'transfer out',
        'account transfer',
        'credit card payment',
        'loan payments',
        'internal transfer',
        'payment thank you',
    ]
    return any(pattern in normalized for pattern in forbidden_patterns)


def generate_manual_id(date, name, amount, account, category):
    normalized_date = normalize_date(date)
    normalized_amount = f'{parse_money(amount):.2f}'
    raw = '|'.join([
        normalized_date,
        normalize_id_component(name),
        normalized_amount,
        normalize_id_component(account),
        normalize_id_component(category),
    ])
    digest = hashlib.sha256(raw.encode('utf-8')).hexdigest()[:18]
    return 'manual_income_' + digest


def row_get(row, key):
    lowered = {str(k or '').strip().lower(): v for k, v in row.items()}
    return lowered.get(key, '')


def validate_required_columns(fieldnames):
    normalized = {str(name or '').strip().lower() for name in (fieldnames or [])}
    missing = [name for name in ['date', 'name', 'amount'] if name not in normalized]
    if missing:
        raise ManualIncomeImportError('Missing required CSV column(s): ' + ', '.join(missing))


def parse_manual_income_csv(
    file_path,
    default_account=DEFAULT_MANUAL_INCOME_ACCOUNT,
    default_category=DEFAULT_MANUAL_INCOME_CATEGORY,
    allow_negative=False,
):
    path = Path(file_path)
    if not path.exists():
        raise ManualIncomeImportError(f'CSV file not found: {path}')

    rows = []
    with path.open(newline='', encoding='utf-8-sig') as handle:
        reader = csv.DictReader(handle)
        validate_required_columns(reader.fieldnames)
        for line_number, raw_row in enumerate(reader, start=2):
            date = normalize_date(row_get(raw_row, 'date'))
            name = normalize_label(row_get(raw_row, 'name'))
            if not name:
                raise ManualIncomeImportError(f'Line {line_number}: name is required.')
            amount = parse_money(row_get(raw_row, 'amount'))
            if amount < 0 and not allow_negative:
                raise ManualIncomeImportError(f'Line {line_number}: negative manual income rows are rejected by default.')
            category = normalize_label(row_get(raw_row, 'category')) or default_category
            account = normalize_label(row_get(raw_row, 'account')) or default_account
            if is_forbidden_income_category(category):
                raise ManualIncomeImportError(f'Line {line_number}: category looks transfer/payment-like: {category}')

            transaction_id = generate_manual_id(date, name, amount, account, category)
            rows.append([
                transaction_id,
                date,
                name,
                amount,
                category,
                account,
                'FALSE',
            ])

    return rows


def dedupe_import_rows(rows, existing_ids=None):
    existing_ids = set(existing_ids or set())
    seen = set()
    new_rows = []
    skipped_existing = 0
    skipped_batch_duplicates = 0

    for row in rows:
        transaction_id = row[0]
        if transaction_id in existing_ids:
            skipped_existing += 1
            continue
        if transaction_id in seen:
            skipped_batch_duplicates += 1
            continue
        seen.add(transaction_id)
        new_rows.append(row)

    return {
        'new_rows': new_rows,
        'skipped_existing': skipped_existing,
        'skipped_batch_duplicates': skipped_batch_duplicates,
    }


def summarize_import(total_rows, new_rows, skipped_existing, skipped_batch_duplicates, dry_run):
    new_row_count = len(new_rows) if isinstance(new_rows, list) else int(new_rows or 0)
    return {
        'total_rows': int(total_rows or 0),
        'new_rows': new_row_count,
        'skipped_existing': int(skipped_existing or 0),
        'skipped_batch_duplicates': int(skipped_batch_duplicates or 0),
        'dry_run': bool(dry_run),
    }


def run_manual_income_import(
    file_path,
    account=DEFAULT_MANUAL_INCOME_ACCOUNT,
    spreadsheet_id=None,
    sheet_name='Transactions',
    dry_run=True,
    confirm=False,
    sheets_client_factory=SheetsClient,
):
    if not dry_run and not confirm:
        raise ManualIncomeImportError('Manual income import requires --confirm before appending rows.')

    rows = parse_manual_income_csv(file_path, default_account=account)
    client = sheets_client_factory(spreadsheet_id or '', sheet_name)
    creds = client.authenticate()
    existing_ids = client.get_existing_ids(creds)
    deduped = dedupe_import_rows(rows, existing_ids)
    new_rows = deduped['new_rows']
    appended = False

    if not dry_run and new_rows:
        appended = bool(client.append_data(creds, new_rows))

    return {
        'appended': appended,
        'rows': new_rows,
        'summary': summarize_import(
            total_rows=len(rows),
            new_rows=new_rows,
            skipped_existing=deduped['skipped_existing'],
            skipped_batch_duplicates=deduped['skipped_batch_duplicates'],
            dry_run=dry_run,
        ),
    }


def format_import_result(result):
    summary = result['summary']
    lines = [
        'Manual Income Import',
        f"Mode: {'dry run' if summary['dry_run'] else 'confirmed append'}",
        f"Parsed rows: {summary['total_rows']}",
        f"New rows: {summary['new_rows']}",
        f"Skipped existing IDs: {summary['skipped_existing']}",
        f"Skipped batch duplicates: {summary['skipped_batch_duplicates']}",
    ]
    if summary['dry_run']:
        lines.append('No rows were appended. Re-run with --confirm to append new manual income rows.')
    elif result['appended']:
        lines.append('Append complete. Refresh Dashboard & Visuals in Google Sheets.')
    else:
        lines.append('No rows appended.')
    return '\n'.join(lines)


def parse_args():
    parser = argparse.ArgumentParser(description='Safely import explicit manual income rows into Transactions!A:G.')
    parser.add_argument('--type', choices=['manual-income'], required=True, help='Import mode. Only manual-income is supported in this safe importer.')
    parser.add_argument('--file', required=True, help='CSV file to import.')
    parser.add_argument('--account', default=DEFAULT_MANUAL_INCOME_ACCOUNT, help='Default account label when the CSV account cell is blank.')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--dry-run', action='store_true', help='Parse and dedupe without appending rows.')
    mode.add_argument('--confirm', action='store_true', help='Append new rows after parsing and dedupe.')
    return parser.parse_args()


def main():
    args = parse_args()
    load_dotenv()
    spreadsheet_id = os.getenv('GOOGLE_SPREADSHEET_ID')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', 'Transactions').strip("'")
    dry_run = not args.confirm

    try:
        result = run_manual_income_import(
            args.file,
            account=args.account,
            spreadsheet_id=spreadsheet_id,
            sheet_name=sheet_name,
            dry_run=dry_run,
            confirm=args.confirm,
        )
    except ManualIncomeImportError as exc:
        print(f'Manual income import failed: {exc}')
        return 1

    print(format_import_result(result))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
