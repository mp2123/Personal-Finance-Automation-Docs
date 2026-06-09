import csv
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .doctor import repo_root


DEMO_TRANSACTIONS_FILE = Path('sample_data/demo_transactions.csv')
REQUIRED_TRANSACTION_COLUMNS = [
    'Transaction ID',
    'Date',
    'Name',
    'Amount',
    'Category',
    'Account',
    'Pending',
]


class DemoDataError(Exception):
    pass


@dataclass(frozen=True)
class DemoTransaction:
    transaction_id: str
    date: str
    name: str
    amount: float
    category: str
    account: str
    pending: str


def parse_demo_transactions(path=None):
    path = Path(path or repo_root() / DEMO_TRANSACTIONS_FILE)
    if not path.exists():
        raise DemoDataError(f'Demo transaction fixture not found: {path}')

    with path.open(newline='') as handle:
        reader = csv.DictReader(handle)
        missing = [column for column in REQUIRED_TRANSACTION_COLUMNS if column not in (reader.fieldnames or [])]
        if missing:
            raise DemoDataError('Demo transaction fixture missing required column(s): ' + ', '.join(missing))

        rows = []
        for index, raw in enumerate(reader, start=2):
            transaction_id = str(raw.get('Transaction ID') or '').strip()
            name = str(raw.get('Name') or '').strip()
            account = str(raw.get('Account') or '').strip()
            category = str(raw.get('Category') or '').strip()
            if not transaction_id.startswith('demo_'):
                raise DemoDataError(f'Demo row {index} transaction ID must start with demo_.')
            if not is_visibly_demo_row(name, account, transaction_id):
                raise DemoDataError(f'Demo row {index} must be visibly labeled as demo/sample data.')
            try:
                normalized_date = date.fromisoformat(str(raw.get('Date') or '').strip()).isoformat()
            except ValueError as exc:
                raise DemoDataError(f'Demo row {index} has an invalid ISO date.') from exc
            try:
                amount = float(str(raw.get('Amount') or '').replace(',', '').strip())
            except ValueError as exc:
                raise DemoDataError(f'Demo row {index} has an invalid amount.') from exc

            rows.append(DemoTransaction(
                transaction_id=transaction_id,
                date=normalized_date,
                name=name,
                amount=round(amount, 2),
                category=category or 'Uncategorized',
                account=account or 'Demo Account',
                pending=str(raw.get('Pending') or 'FALSE').strip().upper(),
            ))

    return rows


def is_visibly_demo_row(name, account, transaction_id):
    joined = ' '.join([name or '', account or '', transaction_id or '']).lower()
    return 'demo' in joined or 'sample only' in joined


def summarize_demo_data(path=None, limit=8):
    rows = parse_demo_transactions(path)
    total_income = round(sum(row.amount for row in rows if row.amount > 0), 2)
    total_spend = round(sum(abs(row.amount) for row in rows if row.amount < 0), 2)
    net_cashflow = round(total_income - total_spend, 2)
    savings_rate = round((net_cashflow / total_income) * 100, 1) if total_income > 0 else None

    account_totals = defaultdict(float)
    category_totals = defaultdict(float)
    for row in rows:
        if row.amount >= 0:
            continue
        account_totals[row.account] += abs(row.amount)
        category_totals[primary_category(row.category)] += abs(row.amount)

    return {
        'ok': True,
        'synthetic': True,
        'source': str(Path(path or repo_root() / DEMO_TRANSACTIONS_FILE)),
        'warning': 'Demo Mode uses committed synthetic fixtures only. It is not connected to the live Google Sheet.',
        'summary': {
            'total_income': total_income,
            'total_spend': total_spend,
            'net_cashflow': net_cashflow,
            'savings_rate': savings_rate,
            'transaction_count': len(rows),
        },
        'top_accounts': serialize_totals(account_totals),
        'top_categories': serialize_totals(category_totals),
        'rows': [serialize_demo_transaction(row) for row in rows[:limit]],
    }


def primary_category(category):
    return str(category or 'Uncategorized').split('>')[0].strip() or 'Uncategorized'


def serialize_totals(totals):
    return [
        {'name': name, 'total': round(total, 2)}
        for name, total in sorted(totals.items(), key=lambda item: item[1], reverse=True)
    ]


def serialize_demo_transaction(row):
    return {
        'transaction_id': row.transaction_id,
        'date': row.date,
        'name': row.name,
        'amount': row.amount,
        'category': row.category,
        'account': row.account,
        'pending': row.pending,
    }
