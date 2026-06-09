import os

from dotenv import load_dotenv

from .account_labels import build_account_label
from .env_tokens import split_access_tokens
from .plaid_client import PlaidClient


def _load_tokens():
    token_value = os.getenv('PLAID_ACCESS_TOKEN') or ''
    return [token.strip() for token in token_value.split(',') if token.strip()]


def collect_linked_accounts(env=None, client_factory=PlaidClient):
    env = env or os.environ
    client_id = env.get('PLAID_CLIENT_ID')
    secret = env.get('PLAID_SECRET')
    plaid_env = env.get('PLAID_ENV', 'sandbox')
    tokens = split_access_tokens(env.get('PLAID_ACCESS_TOKEN'))

    missing = []
    if not client_id:
        missing.append('PLAID_CLIENT_ID')
    if not secret:
        missing.append('PLAID_SECRET')
    if not tokens:
        missing.append('PLAID_ACCESS_TOKEN')
    if missing:
        return {
            'ok': False,
            'plaid_env': plaid_env,
            'item_count': 0,
            'items': [],
            'error': 'Missing required key(s): ' + ', '.join(missing),
        }

    plaid_client = client_factory(client_id, secret, plaid_env)
    items = []
    errors = []

    for index, token in enumerate(tokens, start=1):
        try:
            item = plaid_client.get_item(token)
            institution_id = item.get('institution_id')
            institution_name = plaid_client.get_institution_name(institution_id)
            accounts = plaid_client.get_accounts(token)
        except Exception as exc:
            errors.append({
                'item_index': index,
                'error': str(exc),
            })
            continue

        account_rows = []
        for account in accounts or []:
            account_rows.append({
                'label': build_account_label(institution_name, account),
                'type': account.get('type') or 'unknown',
                'subtype': account.get('subtype') or 'unknown',
            })

        items.append({
            'item_index': index,
            'institution_name': institution_name,
            'account_count': len(account_rows),
            'accounts': account_rows,
        })

    return {
        'ok': not errors,
        'plaid_env': plaid_env,
        'item_count': len(tokens),
        'items': items,
        'errors': errors,
    }


def main():
    load_dotenv('.env')

    result = collect_linked_accounts()

    if not result.get('ok') and result.get('error'):
        print(result['error'])
        return

    print(f"Plaid env: {result['plaid_env']}")
    print(f"Linked Plaid item count: {result['item_count']}")

    for item in result.get('items', []):
        print('')
        print(f"{item['item_index']}. {item['institution_name']}")
        if not item.get('accounts'):
            print('  - No accounts returned for this item.')
            continue

        for account in item['accounts']:
            print(f"  - {account['label']} | {account['type']}/{account['subtype']}")

    for error in result.get('errors', []):
        print('')
        print(f"Item {error['item_index']} failed: {error['error']}")


if __name__ == '__main__':
    main()
