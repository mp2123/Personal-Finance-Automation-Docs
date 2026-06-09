import re


def clean_account_name(value):
    """Returns a readable account/card name from Plaid metadata."""
    text = str(value or '').replace('\ufffd', '').replace('®', '')
    text = re.sub(r'\.{2,}\s*\d{2,4}$', '', text)
    text = re.sub(r'\s+', ' ', text).strip(' -')
    if text.isupper():
        text = text.title()
    return text or 'Unknown Account'


def build_account_label(institution_name, account):
    """Builds the label stored in the Transactions Account column."""
    account = account or {}
    name = clean_account_name(
        account.get('official_name') or account.get('name') or 'Unknown Account'
    )
    institution = clean_account_name(institution_name or 'Unknown Institution')
    mask = str(account.get('mask') or '').strip()

    label = f"{institution} - {name}"
    if mask:
        label += f" ending {mask}"
    return label


def account_id_from_transaction(transaction):
    """Reads account_id from either a dict or a Plaid model object."""
    if isinstance(transaction, dict):
        return transaction.get('account_id')
    return getattr(transaction, 'account_id', None)
