import argparse
import html
import json
import os
import re
import subprocess
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime

from dotenv import load_dotenv

from .appscript_deploy import (
    AppScriptDeployError,
    format_deploy_report,
    run_deploy_plan,
)
from .csv_importer import (
    DEFAULT_MANUAL_INCOME_ACCOUNT,
    ManualIncomeImportError,
    run_manual_income_import,
)
from .demo_data import DemoDataError, summarize_demo_data
from .doctor import (
    PLAID_OAUTH_STATUS_URL,
    QUICKSTART_PYTHON_DIR,
    collect_checks,
    repo_root,
)
from .env_tokens import split_access_tokens
from .list_linked_accounts import collect_linked_accounts


DEFAULT_HOST = '127.0.0.1'
DEFAULT_PORT = 8790
RUNTIME_STATE = {
    'last_sync': None,
    'last_doctor': None,
    'last_appscript_deploy': None,
    'last_import': None,
}


class ControlCenterError(Exception):
    pass


def load_control_env(root=None):
    root = Path(root or repo_root())
    load_dotenv(root / '.env')
    return dict(os.environ)


def build_config_status(env):
    required = ['PLAID_CLIENT_ID', 'PLAID_SECRET', 'PLAID_ENV', 'PLAID_ACCESS_TOKEN', 'GOOGLE_SPREADSHEET_ID']
    missing = [key for key in required if not env.get(key)]
    tokens = split_access_tokens(env.get('PLAID_ACCESS_TOKEN'))
    return {
        'plaid_env': env.get('PLAID_ENV') or 'missing',
        'token_count': len(tokens),
        'required_keys_present': not missing,
        'missing_keys': missing,
        'apps_script_deploy_configured': bool(env.get('GOOGLE_APPS_SCRIPT_ID')),
    }


def build_sheet_status(env):
    spreadsheet_id = env.get('GOOGLE_SPREADSHEET_ID') or ''
    sheet_name = env.get('GOOGLE_SHEET_NAME') or 'Transactions'
    return {
        'configured': bool(spreadsheet_id),
        'can_open': bool(spreadsheet_id),
        'sheet_name': sheet_name,
        'masked_spreadsheet_id': ('...' + spreadsheet_id[-4:]) if spreadsheet_id else '',
        'target_label': 'Configured Google Sheet' if spreadsheet_id else 'Missing GOOGLE_SPREADSHEET_ID',
    }


def build_sheet_open_url(env):
    spreadsheet_id = env.get('GOOGLE_SPREADSHEET_ID') or ''
    if not spreadsheet_id:
        raise ControlCenterError('GOOGLE_SPREADSHEET_ID is missing.')
    return f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit'


def build_appscript_redeploy_checklist():
    return '\n'.join([
        'Apps Script redeploy checklist:',
        '1. Open the bound Apps Script project from the Google Sheet.',
        '2. Replace Code.gs with src/appscript/Code.gs from this repo.',
        '3. Replace Sidebar.html with src/appscript/Sidebar.html from this repo.',
        '4. Save the Apps Script project.',
        '5. Reload the Google Sheet.',
        '6. Run Bank Automation -> Refresh Dashboard & Visuals.',
    ])


def build_quickstart_repair_command():
    return '\n'.join([
        'cd /Users/michaelpanico/Desktop/quickstart/python',
        '/bin/rm -rf -- ./venv',
        'python3 -m venv venv',
        './venv/bin/python -m pip install --upgrade pip',
        './venv/bin/python -m pip install -r requirements.txt',
        './venv/bin/python server.py',
        '',
        'Then start the frontend in another Terminal:',
        'cd /Users/michaelpanico/Desktop/quickstart/frontend',
        'npm start',
    ])


def build_us_bank_guidance():
    return '\n'.join([
        'U.S. Bank connection status:',
        'Plaid currently reports INSTITUTION_REGISTRATION_REQUIRED for U.S. Bank on this client.',
        'That means Plaid Production/OAuth institution registration is blocking the connection.',
        'Do not keep retrying U.S. Bank until the Plaid Dashboard shows registration is ready.',
        f'Check status: {PLAID_OAUTH_STATUS_URL}',
        '',
        'When registration is ready, run:',
        '.venv/bin/python -m src.engine.connect_bank --institution-note "U.S. Bank"',
    ])


def build_manual_income_import_guidance():
    return '\n'.join([
        'Manual income import:',
        'Use this when payroll/checking income is not available through Plaid yet, but you want savings rate to become real.',
        '',
        'Supported CSV columns:',
        'date,name,amount,category,account,notes',
        '',
        'Dry run:',
        '.venv/bin/python -m src.engine.csv_importer --type manual-income --file src/imports/income.csv --account "Manual Income" --dry-run',
        '',
        'Confirmed append:',
        '.venv/bin/python -m src.engine.csv_importer --type manual-income --file src/imports/income.csv --account "Manual Income" --confirm',
        '',
        'Manual income rows must be positive. They append to Transactions!A:G and use the default category Income > Manual Income when category is blank.',
        'After confirming an import, reload the Google Sheet and run Bank Automation -> Refresh Dashboard & Visuals.',
    ])


def build_self_serve_setup_commands():
    return [
        {
            'label': 'Start Control Center',
            'command': 'cd /Users/michaelpanico/Desktop/DevBase/active_projects/Danny_Bank_Automation\n.venv/bin/python -m src.engine.control_center',
            'detail': 'Open the local control center at 127.0.0.1:8790.',
        },
        {
            'label': 'Run Release Smoke Check',
            'command': 'scripts/release_smoke_check.sh',
            'detail': 'Validate tests, Apps Script syntax, demo data, packaging preflights, and secret-safe output.',
        },
        {
            'label': 'Check Apps Script Deploy',
            'command': '.venv/bin/python -m src.engine.appscript_deploy --dry-run',
            'detail': 'Compare repo Apps Script files with the bound Google Sheet project.',
        },
        {
            'label': 'Dry Run Manual Income',
            'command': '.venv/bin/python -m src.engine.csv_importer --type manual-income --file src/imports/income.csv --account "Manual Income" --dry-run',
            'detail': 'Preview local income CSV rows without appending to Google Sheets.',
        },
    ]


def build_trusted_tester_checklist():
    return '\n'.join([
        'Trusted Tester Checklist - self-serve local beta',
        '',
        'Before sharing:',
        '1. Run scripts/release_smoke_check.sh.',
        '2. Verify Demo Mode is visibly synthetic.',
        '3. Verify Setup Readiness has no blocking items on your machine.',
        '4. Verify Apps Script deploy dry-run is clean.',
        '5. Confirm manual income import is dry-run only unless real positive income is reviewed.',
        '',
        'Tester expectations:',
        '1. This is a local beta, not App Store-style software.',
        '2. Data stays on the tester machine and in the tester Google account.',
        '3. Do not share bank credentials, one-time codes, tokens, keys, or raw .env files.',
        '4. Bank coverage depends on Plaid and OAuth institution availability.',
        '5. Outputs are personal finance analytics, not tax, legal, investment, credit, or regulated financial advice.',
        '',
        'After testing:',
        '1. Ask what setup step was confusing.',
        '2. Ask what warning felt scary or unclear.',
        '3. Ask whether the control center made the next action obvious.',
        '4. Record issues in docs/beta_rehearsal_report_template.md.',
    ])


def build_redacted_diagnostics(status_payload=None, env=None):
    status_payload = status_payload or {}
    env = env or {}
    accounts = status_payload.get('accounts') or {}
    account_items = accounts.get('items') or []
    account_count = sum(len(item.get('accounts') or []) for item in account_items)
    payload = {
        'product': 'Danny Bank local beta',
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'project_root': '[local path redacted]',
        'config': status_payload.get('config') or {},
        'sheet': status_payload.get('sheet') or {},
        'readiness': {
            'can_sync': (status_payload.get('readiness') or {}).get('can_sync'),
            'has_blocking_items': (status_payload.get('readiness') or {}).get('has_blocking_items'),
            'recommended_next_step': (status_payload.get('readiness') or {}).get('recommended_next_step'),
        },
        'doctor': status_payload.get('doctor') or {'checks': []},
        'accounts': {
            'institution_count': len(account_items),
            'account_count': account_count,
            'errors': accounts.get('errors') or [],
        },
        'next_actions': status_payload.get('next_actions') or [],
    }
    return redact_local_paths_payload(mask_sensitive_payload(payload, env))


def build_demo_payload(root=None):
    try:
        return summarize_demo_data(Path(root or repo_root()) / 'sample_data' / 'demo_transactions.csv')
    except DemoDataError as exc:
        return {
            'ok': False,
            'synthetic': True,
            'error': str(exc),
            'warning': 'Demo Mode is unavailable until sample_data/demo_transactions.csv is valid.',
            'summary': {},
            'top_accounts': [],
            'top_categories': [],
            'rows': [],
        }


def record_runtime_event(state, event_type, payload):
    state[f'last_{event_type}'] = {
        'timestamp': datetime.now().isoformat(timespec='seconds'),
        **payload,
    }
    return state[f'last_{event_type}']


def mask_sensitive_text(text, env=None):
    env = env or {}
    masked = str(text or '')
    sensitive_values = []
    for key in [
        'PLAID_CLIENT_ID',
        'PLAID_SECRET',
        'GOOGLE_SPREADSHEET_ID',
        'GEMINI_API_KEY',
        'GOOGLE_APPS_SCRIPT_ID',
    ]:
        value = env.get(key)
        if value:
            sensitive_values.append(str(value))
    sensitive_values.extend(split_access_tokens(env.get('PLAID_ACCESS_TOKEN')))

    for value in sorted(set(sensitive_values), key=len, reverse=True):
        if len(value) >= 4:
            masked = masked.replace(value, '[masked]')
    return masked


def mask_sensitive_payload(value, env=None):
    if isinstance(value, dict):
        return {key: mask_sensitive_payload(item, env) for key, item in value.items()}
    if isinstance(value, list):
        return [mask_sensitive_payload(item, env) for item in value]
    if isinstance(value, tuple):
        return tuple(mask_sensitive_payload(item, env) for item in value)
    if isinstance(value, str):
        return mask_sensitive_text(value, env)
    return value


def redact_local_paths(value):
    text = str(value or '')
    text = re.sub(r'/Users/[^\s,;:]+', '[local path redacted]', text)
    text = re.sub(r'/private/var/[^\s,;:]+', '[local path redacted]', text)
    return text


def redact_local_paths_payload(value):
    if isinstance(value, dict):
        return {key: redact_local_paths_payload(item) for key, item in value.items()}
    if isinstance(value, list):
        return [redact_local_paths_payload(item) for item in value]
    if isinstance(value, tuple):
        return tuple(redact_local_paths_payload(item) for item in value)
    if isinstance(value, str):
        return redact_local_paths(value)
    return value


def check_results_to_dicts(checks):
    return [
        {
            'name': check.name,
            'status': check.status,
            'detail': check.detail,
        }
        for check in checks
    ]


def build_doctor_payload(root, env, skip_network=True):
    checks = collect_checks(Path(root), env, QUICKSTART_PYTHON_DIR, skip_network=skip_network)
    return {
        'ok': not any(check.status == 'FAIL' for check in checks),
        'checks': check_results_to_dicts(checks),
    }


def build_account_guidance(accounts_payload):
    accounts = []
    for item in (accounts_payload or {}).get('items', []):
        accounts.extend(item.get('accounts') or [])

    has_income_source = any(
        account.get('type') in ('depository', 'investment') and account.get('subtype') not in ('credit card', 'loan')
        for account in accounts
    )
    if has_income_source:
        return {
            'income_status': 'potential_income_source_present',
            'detail': 'A non-credit account is linked, so verified income may be available if payroll/deposits are present.',
        }
    return {
        'income_status': 'no_verified_income_source',
        'detail': 'Savings rate remains N/A until a checking/payroll income account is linked or income rows are imported.',
    }


def readiness_step(status, title, detail, action_label, blocking):
    return {
        'status': status,
        'title': title,
        'detail': detail,
        'action_label': action_label,
        'blocking': bool(blocking),
    }


def get_doctor_check(doctor_payload, name):
    for check in (doctor_payload or {}).get('checks', []):
        if check.get('name') == name:
            return check
    return None


def build_readiness(root=None, env=None, doctor_payload=None, accounts_payload=None):
    root = Path(root or repo_root())
    env = env or {}
    doctor_payload = doctor_payload or {'checks': []}
    accounts_payload = accounts_payload or {'items': []}
    config = build_config_status(env)
    sheet = build_sheet_status(env)
    account_guidance = build_account_guidance(accounts_payload)

    env_exists = (root / '.env').exists()
    credentials_exists = (root / 'credentials.json').exists()
    token_exists = (root / 'token.json').exists()
    python_ready = (root / '.venv' / 'bin' / 'python').exists() or (root / 'venv' / 'bin' / 'python').exists()
    plaid_keys = all(env.get(key) for key in ['PLAID_CLIENT_ID', 'PLAID_SECRET', 'PLAID_ENV'])
    plaid_tokens = config['token_count'] > 0
    linked_accounts_available = bool((accounts_payload or {}).get('items'))

    google_sheet_check = get_doctor_check(doctor_payload, 'Google Sheets')
    if google_sheet_check:
        sheet_reachable_status = 'ready' if google_sheet_check.get('status') == 'PASS' else 'missing'
        sheet_reachable_detail = google_sheet_check.get('detail') or 'Google Sheet check completed.'
        sheet_reachable_blocking = google_sheet_check.get('status') == 'FAIL'
    elif sheet['configured']:
        sheet_reachable_status = 'ready'
        sheet_reachable_detail = 'Sheet ID is configured. Run Doctor when you need a live reachability check.'
        sheet_reachable_blocking = False
    else:
        sheet_reachable_status = 'missing'
        sheet_reachable_detail = 'GOOGLE_SPREADSHEET_ID is missing.'
        sheet_reachable_blocking = True

    steps = [
        readiness_step(
            'ready' if python_ready else 'warning',
            'Local Python environment',
            'Local virtualenv is ready.' if python_ready else 'The launcher can create or repair .venv when opened.',
            'Open Control Center',
            False,
        ),
        readiness_step(
            'ready' if env_exists else 'missing',
            'Create local .env',
            '.env exists.' if env_exists else 'Create .env from .env.example and fill in local Plaid/Google settings.',
            'Create .env',
            not env_exists,
        ),
        readiness_step(
            'ready' if plaid_keys else 'missing',
            'Add Plaid API keys',
            'Plaid client ID, secret, and env are configured.' if plaid_keys else 'Add PLAID_CLIENT_ID, PLAID_SECRET, and PLAID_ENV to .env.',
            'Add Plaid keys',
            not plaid_keys,
        ),
        readiness_step(
            'ready' if credentials_exists else 'missing',
            'Add Google OAuth credentials',
            'credentials.json exists.' if credentials_exists else 'Place Google OAuth credentials.json in the repo root.',
            'Add credentials.json',
            not credentials_exists,
        ),
        readiness_step(
            'ready' if token_exists else ('warning' if credentials_exists else 'missing'),
            'Authorize Google Sheets',
            'token.json exists.' if token_exists else ('Google OAuth token can be created from credentials.json on first Sheets auth.' if credentials_exists else 'Add credentials.json before Google auth can run.'),
            'Authorize Google',
            False if credentials_exists else True,
        ),
        readiness_step(
            'ready' if sheet['configured'] else 'missing',
            'Configure Google Sheet',
            'GOOGLE_SPREADSHEET_ID is configured.' if sheet['configured'] else 'Add GOOGLE_SPREADSHEET_ID to .env.',
            'Add Sheet ID',
            not sheet['configured'],
        ),
        readiness_step(
            sheet_reachable_status,
            'Reach Google Sheet',
            sheet_reachable_detail,
            'Run Doctor',
            sheet_reachable_blocking,
        ),
        readiness_step(
            'ready' if env.get('GOOGLE_APPS_SCRIPT_ID') else 'warning',
            'Apps Script deploy path',
            'GOOGLE_APPS_SCRIPT_ID is configured for API deploys.' if env.get('GOOGLE_APPS_SCRIPT_ID') else 'API deploy helper is not configured; manual Apps Script paste fallback remains available.',
            'Configure Apps Script ID',
            False,
        ),
        readiness_step(
            'ready' if plaid_tokens else 'missing',
            'Connect a bank',
            f'{config["token_count"]} Plaid access token(s) configured.' if plaid_tokens else 'No Plaid access token is configured yet.',
            'Connect a bank',
            not plaid_tokens,
        ),
        readiness_step(
            'ready' if linked_accounts_available else ('warning' if plaid_tokens else 'missing'),
            'Load linked accounts',
            'Linked account metadata is available.' if linked_accounts_available else ('Plaid token exists, but account metadata has not loaded in this snapshot.' if plaid_tokens else 'Connect a bank before linked accounts are available.'),
            'List Linked Accounts',
            False if plaid_tokens else True,
        ),
        readiness_step(
            'ready' if account_guidance['income_status'] == 'potential_income_source_present' else 'warning',
            'Savings rate needs income',
            account_guidance['detail'],
            'Import manual income',
            False,
        ),
    ]

    blocking_steps = [step for step in steps if step['blocking'] and step['status'] != 'ready']
    warning_steps = [step for step in steps if not step['blocking'] and step['status'] != 'ready']
    recommended = (blocking_steps or warning_steps or [None])[0]
    can_sync = env_exists and plaid_keys and credentials_exists and sheet['configured'] and plaid_tokens

    return {
        'can_sync': bool(can_sync),
        'has_blocking_items': bool(blocking_steps),
        'recommended_next_step': recommended,
        'steps': steps,
    }


def parse_sync_summary(output, ok):
    text = output or ''
    appended_match = re.search(r'Appending\s+(\d+)\s+total new transactions', text)
    cells_match = re.search(r'(\d+)\s+cells appended', text)
    retrieval_count = len(re.findall(r'Retrieving transactions for institution', text))
    no_new = 'No new transactions found across all institutions' in text

    if not ok:
        status = 'failed'
    elif no_new:
      status = 'up_to_date'
    else:
      status = 'completed'

    new_transactions = int(appended_match.group(1)) if appended_match else (0 if no_new else None)
    cells_appended = int(cells_match.group(1)) if cells_match else 0
    steps = []
    if 'Authenticating with Google Sheets' in text:
        steps.append({'label': 'Google Sheets authentication', 'status': 'done'})
    if 'Reading existing transaction IDs' in text:
        steps.append({'label': 'Ledger dedupe scan', 'status': 'done'})
    if retrieval_count:
        steps.append({'label': 'Plaid transaction retrieval', 'status': 'done', 'detail': f'{retrieval_count} institution(s) checked'})
    if new_transactions is not None:
        steps.append({'label': 'New transaction detection', 'status': 'done', 'detail': f'{new_transactions} new row(s)'})
    if cells_appended:
        steps.append({'label': 'Google Sheet append', 'status': 'done', 'detail': f'{cells_appended} cell(s) appended'})
    if no_new:
        steps.append({'label': 'Sheet status', 'status': 'done', 'detail': 'Already up to date'})
    if not steps:
        steps.append({'label': 'Command completed' if ok else 'Command failed', 'status': 'done' if ok else 'failed'})

    return {
        'status': status,
        'new_transactions': new_transactions,
        'cells_appended': cells_appended,
        'steps': steps,
    }


def build_next_actions(doctor_payload=None, accounts_payload=None, sync_result=None, import_result=None):
    doctor_payload = doctor_payload or {'checks': []}
    accounts_payload = accounts_payload or {'items': []}
    actions = []

    summary = (sync_result or {}).get('summary') or {}
    if number_like(summary.get('new_transactions')) > 0:
        actions.append({
            'priority': 'high',
            'title': 'Refresh the Dashboard',
            'detail': 'A sync appended new rows. In Google Sheets, run Bank Automation -> Refresh Dashboard & Visuals.',
        })

    import_summary = (import_result or {}).get('summary') or {}
    if import_result and import_result.get('appended') and number_like(import_summary.get('new_rows')) > 0:
        actions.append({
            'priority': 'high',
            'title': 'Refresh the Dashboard',
            'detail': 'Manual income rows were appended. Refresh Dashboard & Visuals in Google Sheets.',
        })

    for check in doctor_payload.get('checks', []):
        name = check.get('name', '')
        if name == 'quickstart venv' and check.get('status') == 'WARN':
            actions.append({
                'priority': 'medium',
                'title': 'Repair Quickstart fallback when needed',
                'detail': 'Quickstart has stale path markers. Use the Quickstart repair command before using fallback bank linking.',
            })
        if name == 'Plaid OAuth blockers':
            actions.append({
                'priority': 'medium',
                'title': 'Wait for Plaid OAuth registration',
                'detail': 'Plaid OAuth institutions such as U.S. Bank stay blocked until Plaid registration is approved.',
            })
        if name == 'Apps Script deploy config' and check.get('status') == 'WARN':
            actions.append({
                'priority': 'medium',
                'title': 'Apps Script deploy helper not configured',
                'detail': 'Add GOOGLE_APPS_SCRIPT_ID to .env for Apps Script deploy checks, or keep using the manual redeploy checklist.',
            })

    guidance = build_account_guidance(accounts_payload)
    if guidance['income_status'] == 'no_verified_income_source':
        actions.append({
            'priority': 'low',
            'title': 'Savings rate needs income',
            'detail': guidance['detail'],
        })

    if not actions:
        actions.append({
            'priority': 'low',
            'title': 'No immediate action',
            'detail': 'System checks are clean and no new sync follow-up is required.',
        })
    return actions


def number_like(value):
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def build_status_payload(root=None, env=None, runtime_state=None):
    root = Path(root or repo_root())
    env = env or load_control_env(root)
    doctor_payload = build_doctor_payload(root, env, skip_network=True)
    accounts_payload = collect_linked_accounts(env)
    readiness = build_readiness(root, env, doctor_payload, accounts_payload)
    runtime_state = RUNTIME_STATE if runtime_state is None else runtime_state
    return {
        'project_root': str(root),
        'config': build_config_status(env),
        'sheet': build_sheet_status(env),
        'doctor': doctor_payload,
        'accounts': accounts_payload,
        'account_guidance': build_account_guidance(accounts_payload),
        'readiness': readiness,
        'runtime': runtime_state,
        'next_actions': build_next_actions(
            doctor_payload=doctor_payload,
            accounts_payload=accounts_payload,
            sync_result=runtime_state.get('last_sync'),
            import_result=runtime_state.get('last_import'),
        ),
        'demo': build_demo_payload(root),
        'appscript': {
            'deploy_helper_configured': bool(env.get('GOOGLE_APPS_SCRIPT_ID')),
            'manual_deploy_required': not bool(env.get('GOOGLE_APPS_SCRIPT_ID')),
            'checklist': build_appscript_redeploy_checklist(),
        },
        'blockers': [{
            'name': 'U.S. Bank OAuth registration',
            'status': 'blocked_by_plaid',
            'code': 'INSTITUTION_REGISTRATION_REQUIRED',
            'url': PLAID_OAUTH_STATUS_URL,
            'detail': 'Wait for Plaid Production/OAuth institution registration before retrying U.S. Bank.',
        }],
    }


def run_sync_command(confirm=False, root=None, env=None, runner=subprocess.run):
    if not confirm:
        raise ControlCenterError('Sync requires explicit browser confirmation because it can append rows to Google Sheets.')

    root = Path(root or repo_root())
    env = env or load_control_env(root)
    proc = runner(
        [sys.executable, '-m', 'src.engine.main'],
        cwd=str(root),
        text=True,
        capture_output=True,
        timeout=180,
    )
    output = '\n'.join(part for part in [getattr(proc, 'stdout', ''), getattr(proc, 'stderr', '')] if part)
    masked_output = mask_sensitive_text(output.strip(), env)
    ok = getattr(proc, 'returncode', 1) == 0
    return {
        'ok': ok,
        'returncode': getattr(proc, 'returncode', 1),
        'output': masked_output,
        'summary': parse_sync_summary(masked_output, ok),
    }


def run_doctor_command(root=None, env=None, skip_network=True):
    root = Path(root or repo_root())
    env = env or load_control_env(root)
    return build_doctor_payload(root, env, skip_network=skip_network)


def normalize_import_path(root, requested_path):
    root = Path(root or repo_root()).resolve()
    imports_dir = (root / 'src' / 'imports').resolve()
    raw_path = str(requested_path or '').strip() or 'src/imports/income.csv'
    candidate = Path(raw_path)
    path = candidate.resolve() if candidate.is_absolute() else (root / candidate).resolve()

    try:
        path.relative_to(imports_dir)
    except ValueError as exc:
        raise ControlCenterError('Manual income import files must be under src/imports/.') from exc

    if path.suffix.lower() != '.csv':
        raise ControlCenterError('Manual income import file must be a .csv file under src/imports/.')

    return path


def serialize_import_row(row):
    return {
        'transaction_id': row[0],
        'date': row[1],
        'name': row[2],
        'amount': row[3],
        'category': row[4],
        'account': row[5],
        'pending': row[6],
    }


def format_manual_income_import_output(result):
    summary = result.get('summary') or {}
    lines = [
        'Manual Income Import',
        'Mode: dry run' if summary.get('dry_run') else 'Mode: confirmed append',
        f"Parsed rows: {summary.get('total_rows', 0)}",
        f"New rows: {summary.get('new_rows', 0)}",
        f"Skipped existing IDs: {summary.get('skipped_existing', 0)}",
        f"Skipped batch duplicates: {summary.get('skipped_batch_duplicates', 0)}",
    ]
    rows = result.get('rows') or []
    if rows:
        lines.append('')
        lines.append('Rows:')
        for row in rows[:12]:
            item = serialize_import_row(row)
            lines.append(
                f"- {item['date']} | {item['name']} | ${float(item['amount']):.2f} | "
                f"{item['category']} | {item['account']} | {item['transaction_id']}"
            )
    if summary.get('dry_run'):
        lines.append('')
        lines.append('No rows were appended. Review the rows, then use Confirm Manual Income Import if they are correct.')
    elif result.get('appended'):
        lines.append('')
        lines.append('Append complete. Refresh Dashboard & Visuals in Google Sheets.')
    else:
        lines.append('')
        lines.append('No rows appended.')
    return '\n'.join(lines)


def run_manual_income_import_command(
    root=None,
    env=None,
    file_path='src/imports/income.csv',
    account=DEFAULT_MANUAL_INCOME_ACCOUNT,
    dry_run=True,
    confirm=False,
    import_runner=run_manual_income_import,
):
    if not dry_run and not confirm:
        raise ControlCenterError('Manual income import requires explicit browser confirmation before appending rows.')

    root = Path(root or repo_root())
    env = env or load_control_env(root)
    path = normalize_import_path(root, file_path)
    try:
        result = import_runner(
            file_path=path,
            account=account or DEFAULT_MANUAL_INCOME_ACCOUNT,
            spreadsheet_id=env.get('GOOGLE_SPREADSHEET_ID'),
            sheet_name=(env.get('GOOGLE_SHEET_NAME') or 'Transactions').strip("'"),
            dry_run=dry_run,
            confirm=confirm,
        )
    except ManualIncomeImportError as exc:
        return {
            'ok': False,
            'error': str(exc),
            'output': mask_sensitive_text('Manual income import failed: ' + str(exc), env),
        }

    output = format_manual_income_import_output(result)
    return {
        'ok': True,
        'appended': bool(result.get('appended')),
        'summary': result.get('summary') or {},
        'rows': [serialize_import_row(row) for row in (result.get('rows') or [])],
        'output': mask_sensitive_text(output, env),
        'next_action': 'Refresh Dashboard & Visuals in Google Sheets.' if result.get('appended') else '',
    }


def run_appscript_dry_run(root=None, env=None, deploy_runner=run_deploy_plan):
    root = Path(root or repo_root())
    env = env or load_control_env(root)
    try:
        report = deploy_runner(env, dry_run=True, root=root)
        output = format_deploy_report(report, env)
        return {
            'ok': bool(report.get('ok')),
            'report': mask_sensitive_payload(report, env),
            'output': mask_sensitive_text(output, env),
        }
    except AppScriptDeployError as exc:
        masked_error = mask_sensitive_text(str(exc), env)
        output = masked_error + '\n\n' + build_appscript_redeploy_checklist()
        return {
            'ok': False,
            'error': masked_error,
            'output': mask_sensitive_text(output, env),
        }
    except Exception as exc:
        masked_error = mask_sensitive_text(str(exc), env)
        output = masked_error + '\n\n' + build_appscript_redeploy_checklist()
        return {
            'ok': False,
            'error': masked_error,
            'output': mask_sensitive_text(output, env),
        }


def run_appscript_deploy(confirm=False, root=None, env=None, deploy_runner=run_deploy_plan):
    if not confirm:
        raise ControlCenterError('Apps Script deploy requires explicit browser confirmation because it overwrites the bound script project.')

    root = Path(root or repo_root())
    env = env or load_control_env(root)
    try:
        report = deploy_runner(env, dry_run=False, root=root, confirmed=True)
        output = format_deploy_report(report, env)
        return {
            'ok': bool(report.get('ok')),
            'report': mask_sensitive_payload(report, env),
            'output': mask_sensitive_text(output, env),
        }
    except AppScriptDeployError as exc:
        masked_error = mask_sensitive_text(str(exc), env)
        output = masked_error + '\n\n' + build_appscript_redeploy_checklist()
        return {
            'ok': False,
            'error': masked_error,
            'output': mask_sensitive_text(output, env),
        }
    except Exception as exc:
        masked_error = mask_sensitive_text(str(exc), env)
        output = masked_error + '\n\n' + build_appscript_redeploy_checklist()
        return {
            'ok': False,
            'error': masked_error,
            'output': mask_sensitive_text(output, env),
        }


def render_control_center_html():
    title = 'Danny Bank Control Center'
    escaped_checklist = html.escape(build_appscript_redeploy_checklist())
    escaped_guidance = html.escape(build_us_bank_guidance())
    escaped_quickstart = html.escape(build_quickstart_repair_command())
    escaped_income_import = html.escape(build_manual_income_import_guidance())
    escaped_tester_checklist = html.escape(build_trusted_tester_checklist())
    setup_commands_json = json.dumps(build_self_serve_setup_commands()).replace('`', '\\`').replace('</', '<\\/')
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <style>
    :root {{
      --bg: #f6f8fb;
      --panel: #ffffff;
      --text: #111827;
      --muted: #64748b;
      --border: #d5dbe5;
      --blue: #2563eb;
      --green: #15803d;
      --amber: #b45309;
      --red: #b91c1c;
      --navy: #0f172a;
      --cyan: #0e7490;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
    }}
    header {{
      background: var(--navy);
      color: #fff;
      padding: 18px 28px;
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: center;
    }}
    h1 {{ font-size: 20px; margin: 0; }}
    main {{ max-width: 1240px; margin: 0 auto; padding: 22px; }}
    h2 {{ font-size: 14px; margin: 0 0 10px; color: #0f172a; }}
    h3 {{ font-size: 13px; margin: 0 0 8px; color: #334155; }}
    .topline {{ display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }}
    .pill {{
      border-radius: 999px;
      padding: 5px 9px;
      font-size: 11px;
      font-weight: 800;
      border: 1px solid var(--border);
      background: #fff;
      color: #334155;
    }}
    .pill.pass {{ color: var(--green); border-color: #bbf7d0; background: #f0fdf4; }}
    .pill.warn {{ color: var(--amber); border-color: #fde68a; background: #fffbeb; }}
    .pill.fail {{ color: var(--red); border-color: #fecaca; background: #fef2f2; }}
    .safety-tag {{
      display: inline-block;
      margin-left: 7px;
      border-radius: 999px;
      padding: 2px 6px;
      font-size: 10px;
      font-weight: 850;
      line-height: 1.2;
      vertical-align: middle;
      border: 1px solid rgba(255,255,255,0.65);
      background: rgba(255,255,255,0.16);
      color: inherit;
    }}
    button.secondary .safety-tag {{
      border-color: #bfdbfe;
      background: #eff6ff;
      color: var(--blue);
    }}
    button.warning .safety-tag {{
      border-color: #fed7aa;
      background: #fff7ed;
      color: var(--amber);
    }}
    .layout {{ display: grid; grid-template-columns: 1.5fr 1fr; gap: 16px; align-items: start; }}
    .start-here {{
      border: 1px solid #bae6fd;
      border-left: 5px solid var(--cyan);
      background: #f0f9ff;
    }}
    .next-step-banner {{
      border: 1px solid #bbf7d0;
      border-left: 5px solid var(--green);
      background: #f0fdf4;
    }}
    .next-step-banner.blocked {{
      border-color: #fecaca;
      border-left-color: var(--red);
      background: #fef2f2;
    }}
    .next-step-banner.warning {{
      border-color: #fde68a;
      border-left-color: var(--amber);
      background: #fffbeb;
    }}
    .grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-bottom: 16px; }}
    .card {{
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 14px;
      min-height: 106px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 14px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }}
    .metric {{ font-size: 23px; font-weight: 850; margin: 5px 0; color: #0f172a; }}
    .muted {{ color: var(--muted); font-size: 12px; line-height: 1.45; }}
    .actions {{ display: flex; flex-wrap: wrap; gap: 10px; margin: 18px 0; }}
    .action-group {{
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 14px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }}
    details.action-group summary {{
      cursor: pointer;
      font-weight: 850;
      color: #0f172a;
      margin-bottom: 8px;
    }}
    button {{
      border: 1px solid var(--blue);
      background: var(--blue);
      color: #fff;
      border-radius: 6px;
      padding: 10px 12px;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      font-size: 13px;
    }}
    button.secondary {{ background: #fff; color: var(--blue); }}
    button.warning {{ background: var(--amber); border-color: var(--amber); }}
    button:disabled {{ opacity: 0.65; cursor: not-allowed; }}
    input[type="text"] {{
      width: 100%;
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 9px 10px;
      font-size: 13px;
      color: var(--text);
      background: #fff;
    }}
    label {{
      display: block;
      font-size: 11px;
      font-weight: 800;
      color: #334155;
      margin-bottom: 5px;
    }}
    .form-grid {{
      display: grid;
      grid-template-columns: minmax(0, 1.5fr) minmax(180px, 0.5fr);
      gap: 10px;
      align-items: end;
    }}
    pre {{
      background: var(--navy);
      color: #e2e8f0;
      padding: 14px;
      border-radius: 8px;
      overflow: auto;
      white-space: pre-wrap;
      min-height: 160px;
      margin: 0;
    }}
    .list {{ display: grid; gap: 8px; }}
    .row {{
      border: 1px solid #e2e8f0;
      border-radius: 7px;
      padding: 10px;
      background: #fff;
    }}
    .row-title {{ font-weight: 800; font-size: 13px; color: #0f172a; }}
    .row-detail {{ color: var(--muted); font-size: 12px; margin-top: 3px; line-height: 1.45; }}
    .account-card {{
      border-left: 4px solid var(--cyan);
    }}
    .timeline {{ display: grid; gap: 7px; }}
    .step {{
      display: grid;
      grid-template-columns: 18px 1fr;
      gap: 8px;
      align-items: start;
      font-size: 12px;
    }}
    .dot {{
      width: 10px;
      height: 10px;
      border-radius: 999px;
      margin-top: 4px;
      background: var(--green);
    }}
    .dot.failed {{ background: var(--red); }}
    .empty {{ color: var(--muted); font-size: 12px; padding: 10px 0; }}
    .readiness {{
      border: 1px solid var(--border);
      border-left: 5px solid var(--amber);
      background: #fff;
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 16px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }}
    .readiness.ready {{ border-left-color: var(--green); }}
    .readiness-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; margin-top: 10px; }}
    .demo-banner {{
      border: 1px solid #bfdbfe;
      border-left: 5px solid var(--blue);
      background: #eff6ff;
    }}
    .mini-grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin: 10px 0; }}
    .mini-card {{
      border: 1px solid #dbeafe;
      border-radius: 7px;
      background: #fff;
      padding: 10px;
    }}
    .mini-value {{ font-size: 18px; font-weight: 850; color: #0f172a; margin-top: 4px; }}
    @media (max-width: 900px) {{
      .layout {{ grid-template-columns: 1fr; }}
      .grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .form-grid {{ grid-template-columns: 1fr; }}
      .mini-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      header {{ align-items: flex-start; flex-direction: column; }}
    }}
    @media (max-width: 560px) {{
      .grid {{ grid-template-columns: 1fr; }}
      .mini-grid {{ grid-template-columns: 1fr; }}
      main {{ padding: 14px; }}
      .actions button {{ width: 100%; }}
    }}
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Danny Bank Control Center</h1>
      <div class="muted">Local-only finance operations for sync, health, accounts, and setup guidance.</div>
    </div>
    <div class="topline" id="headerBadges"></div>
  </header>
  <main>
    <section class="panel start-here" id="startHerePanel">
      <h2>Start Here</h2>
      <div class="muted">Use this local beta as a self-serve control center. Follow readiness, run dry checks before write actions, and share only redacted diagnostics when asking for help.</div>
      <div class="actions" style="margin:12px 0 0;">
        <button class="secondary" onclick="showText(`{escaped_tester_checklist}`)">Trusted Tester Checklist <span class="safety-tag">Safe To Click</span></button>
        <button class="secondary" onclick="copySetupCommands()">Copy Setup Commands <span class="safety-tag">Safe To Click</span></button>
        <button class="secondary" onclick="copyRedactedDiagnostics()">Copy Redacted Diagnostics <span class="safety-tag">Safe To Click</span></button>
      </div>
    </section>
    <section id="recommendedNextStepPanel"></section>
    <section id="readinessPanel"></section>
    <section class="grid" id="cards"></section>
    <section class="panel demo-banner" id="demoModePanel"></section>
    <section class="action-group">
      <h2>Primary Actions</h2>
      <div class="muted">Start with read-only checks. Buttons marked as write actions require confirmation before changing Google Sheets.</div>
      <div class="actions">
        <button onclick="runDoctor()">Run Doctor <span class="safety-tag">Safe To Click</span></button>
        <button onclick="listAccounts()">List Linked Accounts <span class="safety-tag">Safe To Click</span></button>
        <button class="secondary" onclick="checkAppsScriptDeploy()">Check Apps Script Deploy Status <span class="safety-tag">Safe To Click</span></button>
        <button class="secondary" onclick="openSheet()">Open Google Sheet <span class="safety-tag">Safe To Click</span></button>
        <button id="runSyncButton" class="warning" onclick="runSync()">Run Sync Now <span class="safety-tag">Writes To Google Sheet</span></button>
        <button class="secondary" onclick="showText(`{escaped_income_import}`)">Manual Income Import Guide <span class="safety-tag">Safe To Click</span></button>
      </div>
    </section>
    <details class="action-group">
      <summary>Advanced Tools</summary>
      <div class="muted">Use these when a dry run, fallback deploy, Plaid OAuth blocker, or setup repair step is specifically needed.</div>
      <div class="actions">
        <button class="secondary" onclick="deployAppsScript()">Deploy Apps Script</button>
        <button class="secondary" onclick="showText(`{escaped_guidance}`)">Connect a Bank / OAuth Help</button>
        <button class="secondary" onclick="showText(`{escaped_quickstart}`)">Quickstart Repair Command</button>
        <button class="secondary" onclick="copyChecklist()">Copy Apps Script Redeploy Checklist</button>
      </div>
    </details>
    <section class="panel">
      <h2>Manual Income Import</h2>
      <div class="form-grid">
        <div>
          <label for="manualIncomePath">CSV path</label>
          <input type="text" id="manualIncomePath" value="src/imports/income.csv">
        </div>
        <div>
          <label for="manualIncomeAccount">Account</label>
          <input type="text" id="manualIncomeAccount" value="Manual Income">
        </div>
      </div>
      <div class="actions" style="margin:12px 0 0;">
        <button class="secondary" onclick="dryRunManualIncomeImport()">Dry Run Manual Income Import <span class="safety-tag">Safe To Click</span></button>
        <button class="warning" onclick="confirmManualIncomeImport()">Confirm Manual Income Import <span class="safety-tag">Writes To Google Sheet</span></button>
      </div>
      <div class="muted">Imports are restricted to repo-local CSV files under src/imports/. Dry run first; confirmed import can append rows to Google Sheets.</div>
    </section>
    <section class="layout">
      <div>
        <section class="panel">
          <h2>Next Actions</h2>
          <div class="list" id="nextActions"></div>
        </section>
        <section class="panel">
          <h2>Sync Progress</h2>
          <div class="timeline" id="syncTimeline"></div>
        </section>
        <section class="panel">
          <h2>Linked Accounts</h2>
          <div class="list" id="accounts"></div>
        </section>
      </div>
      <aside>
        <section class="panel">
          <h2>Health Warnings</h2>
          <div class="list" id="warnings"></div>
        </section>
        <section class="panel">
          <h2>Last Activity</h2>
          <div class="list" id="activity"></div>
        </section>
        <section class="panel">
          <h2>Command Output</h2>
          <pre id="output">Loading status...</pre>
        </section>
      </aside>
    </section>
  </main>
  <script>
    const checklist = `{escaped_checklist}`;
    const setupCommands = {setup_commands_json};
    let latestStatus = null;

    function setOutput(value) {{
      document.getElementById('output').textContent = typeof value === 'string' ? value : JSON.stringify(value, null, 2);
    }}

    function showText(value) {{
      const parser = new DOMParser();
      setOutput(parser.parseFromString(value, 'text/html').documentElement.textContent);
    }}

    async function api(path, options) {{
      const response = await fetch(path, options);
      const payload = await response.json();
      if (!response.ok) {{
        throw new Error(payload.error || 'Request failed');
      }}
      return payload;
    }}

    function badge(text, kind) {{
      return `<span class="pill ${{kind || ''}}">${{text}}</span>`;
    }}

    function formatMoney(value) {{
      const number = Number(value || 0);
      return number.toLocaleString(undefined, {{ style: 'currency', currency: 'USD' }});
    }}

    function renderCards(status) {{
      const cards = [
        ['Plaid', status.config.plaid_env, `${{status.config.token_count}} configured token(s)`],
        ['Google Sheet', status.sheet.configured ? 'Configured' : 'Missing', status.sheet.sheet_name],
        ['Linked Accounts', `${{status.accounts.items.length}} institution(s)`, status.account_guidance.detail],
        ['Doctor', status.doctor.ok ? 'No failures' : 'Needs attention', `${{status.doctor.checks.length}} check(s)`]
      ];
      document.getElementById('cards').innerHTML = cards.map(card => `
        <article class="card">
          <h3>${{card[0]}}</h3>
          <div class="metric">${{card[1]}}</div>
          <div class="muted">${{card[2]}}</div>
        </article>
      `).join('');
    }}

    function renderDemoMode(status) {{
      const demo = status.demo || {{}};
      const summary = demo.summary || {{}};
      if (!demo.ok) {{
        document.getElementById('demoModePanel').innerHTML = `
          <h2>Demo Mode - synthetic data only <span class="safety-tag">Safe To Click</span></h2>
          <div class="muted">${{demo.warning || 'Demo data is unavailable.'}}</div>
        `;
        return;
      }}
      const metricCards = [
        ['Demo Income', formatMoney(summary.total_income)],
        ['Demo Spend', formatMoney(summary.total_spend)],
        ['Demo Net', formatMoney(summary.net_cashflow)],
        ['Demo Savings Rate', `${{summary.savings_rate}}%`]
      ];
      const accountCards = (demo.top_accounts || []).slice(0, 3).map(item => `
        <div class="row">
          <div class="row-title">${{item.name}}</div>
          <div class="row-detail">Demo spend: ${{formatMoney(item.total)}}</div>
        </div>
      `).join('');
      const categoryCards = (demo.top_categories || []).slice(0, 3).map(item => `
        <div class="row">
          <div class="row-title">${{item.name}}</div>
          <div class="row-detail">Demo spend: ${{formatMoney(item.total)}}</div>
        </div>
      `).join('');
      document.getElementById('demoModePanel').innerHTML = `
        <h2>Demo Mode - synthetic data only <span class="safety-tag">Safe To Click</span></h2>
        <div class="muted">${{demo.warning}}</div>
        <div class="mini-grid">
          ${{metricCards.map(card => `
            <div class="mini-card">
              <div class="muted">${{card[0]}}</div>
              <div class="mini-value">${{card[1]}}</div>
            </div>
          `).join('')}}
        </div>
        <div class="layout" style="grid-template-columns:1fr 1fr; gap:10px;">
          <div>
            <h3>Sample Accounts</h3>
            <div class="list">${{accountCards || '<div class="empty">No demo accounts.</div>'}}</div>
          </div>
          <div>
            <h3>Sample Categories</h3>
            <div class="list">${{categoryCards || '<div class="empty">No demo categories.</div>'}}</div>
          </div>
        </div>
      `;
    }}

    function renderReadiness(status) {{
      const readiness = status.readiness || {{}};
      const recommended = readiness.recommended_next_step;
      const visibleSteps = (readiness.steps || []).filter(step => step.status !== 'ready').slice(0, 6);
      const panelClass = readiness.has_blocking_items ? 'readiness' : 'readiness ready';
      document.getElementById('readinessPanel').innerHTML = `
        <section class="${{panelClass}}">
          <h2>Setup Readiness</h2>
          <div class="muted">${{readiness.can_sync ? 'Required sync setup is ready.' : 'Complete the blocking setup items before running sync.'}}</div>
          ${{recommended ? `<div class="row" style="margin-top:10px;"><div class="row-title">Recommended next step: ${{recommended.title}}</div><div class="row-detail">${{recommended.detail}}</div></div>` : ''}}
          <div class="readiness-grid">
            ${{visibleSteps.map(step => `
              <div class="row">
                <div class="row-title">${{badge(step.status, step.blocking ? 'fail' : 'warn')}} ${{step.title}}</div>
                <div class="row-detail">${{step.detail}}</div>
              </div>
            `).join('')}}
          </div>
        </section>
      `;
      const syncButton = document.getElementById('runSyncButton');
      if (syncButton) {{
        syncButton.disabled = !readiness.can_sync;
        syncButton.title = readiness.can_sync ? '' : 'Complete setup readiness items before syncing.';
      }}
    }}

    function renderRecommendedNextStep(status) {{
      const readiness = status.readiness || {{}};
      const recommended = readiness.recommended_next_step;
      const classes = ['panel', 'next-step-banner'];
      if (readiness.has_blocking_items) {{
        classes.push('blocked');
      }} else if (recommended && recommended.status !== 'ready') {{
        classes.push('warning');
      }}
      const title = recommended ? recommended.title : 'No immediate action';
      const detail = recommended ? recommended.detail : 'System checks are clean and no local follow-up is required.';
      const actionLabel = recommended && recommended.action_label ? recommended.action_label : 'Continue testing';
      document.getElementById('recommendedNextStepPanel').innerHTML = `
        <section class="${{classes.join(' ')}}">
          <h2>Recommended Next Step</h2>
          <div class="row-title">${{title}}</div>
          <div class="row-detail">${{detail}}</div>
          <div class="muted" style="margin-top:8px;">Suggested action: ${{actionLabel}}</div>
        </section>
      `;
    }}

    function renderHeader(status) {{
      const warningCount = status.doctor.checks.filter(check => check.status === 'WARN').length;
      const failCount = status.doctor.checks.filter(check => check.status === 'FAIL').length;
      document.getElementById('headerBadges').innerHTML = [
        badge('Local only', 'pass'),
        badge(status.doctor.ok ? 'No failures' : `${{failCount}} failure(s)`, status.doctor.ok ? 'pass' : 'fail'),
        badge(`${{warningCount}} warning(s)`, warningCount ? 'warn' : 'pass')
      ].join('');
    }}

    function renderWarnings(status) {{
      const warnings = status.doctor.checks.filter(check => check.status !== 'PASS').concat(status.blockers || []);
      document.getElementById('warnings').innerHTML = warnings.length ? warnings.map(item => `
        <div class="row">
          <div class="row-title">${{item.name || item.code}}</div>
          <div class="row-detail">${{item.detail}}</div>
        </div>
      `).join('') : '<div class="empty">No health warnings.</div>';
    }}

    function renderActions(status) {{
      document.getElementById('nextActions').innerHTML = (status.next_actions || []).map(action => `
        <div class="row">
          <div class="row-title">${{badge(action.priority, action.priority === 'high' ? 'fail' : action.priority === 'medium' ? 'warn' : 'pass')}} ${{action.title}}</div>
          <div class="row-detail">${{action.detail}}</div>
        </div>
      `).join('');
    }}

    function renderAccounts(status) {{
      const items = status.accounts.items || [];
      if (!items.length) {{
        document.getElementById('accounts').innerHTML = '<div class="empty">No linked account data loaded.</div>';
        return;
      }}
      document.getElementById('accounts').innerHTML = items.map(item => `
        <div class="row account-card">
          <div class="row-title">${{item.institution_name}}</div>
          <div class="row-detail">${{(item.accounts || []).map(account => `${{account.label}} (${{
            account.type
          }}/${{account.subtype}})`).join('<br>')}}</div>
        </div>
      `).join('');
    }}

    function renderTimeline(sync) {{
      const summary = sync && sync.summary ? sync.summary : null;
      if (!summary) {{
        document.getElementById('syncTimeline').innerHTML = '<div class="empty">No sync has run from this control center session yet.</div>';
        return;
      }}
      document.getElementById('syncTimeline').innerHTML = summary.steps.map(step => `
        <div class="step">
          <span class="dot ${{step.status === 'failed' ? 'failed' : ''}}"></span>
          <div><strong>${{step.label}}</strong><div class="row-detail">${{step.detail || step.status}}</div></div>
        </div>
      `).join('');
    }}

    function renderActivity(status) {{
      const runtime = status.runtime || {{}};
      const rows = [];
      if (runtime.last_sync) {{
        rows.push(`<div class="row"><div class="row-title">Last sync</div><div class="row-detail">${{runtime.last_sync.timestamp}} · ${{runtime.last_sync.summary.status}}</div></div>`);
      }}
      if (runtime.last_doctor) {{
        rows.push(`<div class="row"><div class="row-title">Last doctor</div><div class="row-detail">${{runtime.last_doctor.timestamp}} · ${{runtime.last_doctor.ok ? 'No failures' : 'Needs attention'}}</div></div>`);
      }}
      if (runtime.last_appscript_deploy) {{
        rows.push(`<div class="row"><div class="row-title">Last Apps Script deploy check</div><div class="row-detail">${{runtime.last_appscript_deploy.timestamp}} · ${{runtime.last_appscript_deploy.ok ? 'Ready' : 'Needs setup'}}</div></div>`);
      }}
      if (runtime.last_import) {{
        rows.push(`<div class="row"><div class="row-title">Last manual income import</div><div class="row-detail">${{runtime.last_import.timestamp}} · ${{runtime.last_import.appended ? 'Rows appended' : 'Dry run / no append'}}</div></div>`);
      }}
      document.getElementById('activity').innerHTML = rows.length ? rows.join('') : '<div class="empty">No actions yet this session.</div>';
    }}

    function renderAll(status) {{
      latestStatus = status;
      renderHeader(status);
      renderRecommendedNextStep(status);
      renderReadiness(status);
      renderCards(status);
      renderDemoMode(status);
      renderWarnings(status);
      renderActions(status);
      renderAccounts(status);
      renderTimeline(status.runtime.last_sync);
      renderActivity(status);
    }}

    async function loadStatus(writeOutput = true) {{
      const status = await api('/api/status');
      renderAll(status);
      const lines = ['Project root: ' + status.project_root, '', 'Doctor snapshot:'];
      status.doctor.checks.forEach(check => {{
        lines.push(`[${{check.status}}] ${{check.name}}: ${{check.detail}}`);
      }});
      if (writeOutput) {{
        setOutput(lines.join('\\n'));
      }}
    }}

    async function runDoctor() {{
      setOutput('Running doctor...');
      const result = await api('/api/doctor', {{ method: 'POST' }});
      setOutput(result.checks.map(check => `[${{check.status}}] ${{check.name}}: ${{check.detail}}`).join('\\n'));
      await loadStatus(false);
    }}

    async function listAccounts() {{
      setOutput('Loading linked accounts...');
      const result = await api('/api/linked-accounts');
      const lines = [`Plaid env: ${{result.plaid_env}}`, `Linked Plaid item count: ${{result.item_count}}`];
      (result.items || []).forEach(item => {{
        lines.push('', `${{item.item_index}}. ${{item.institution_name}}`);
        (item.accounts || []).forEach(account => lines.push(`  - ${{account.label}} | ${{account.type}}/${{account.subtype}}`));
      }});
      (result.errors || []).forEach(error => lines.push('', `Item ${{error.item_index}} failed: ${{error.error}}`));
      setOutput(lines.join('\\n'));
      await loadStatus(false);
    }}

    async function runSync() {{
      if (latestStatus && latestStatus.readiness && !latestStatus.readiness.can_sync) {{
        setOutput('Sync is blocked until required setup readiness items are complete.');
        return;
      }}
      if (!confirm('Run a live sync now? This can append new rows to Google Sheets.')) {{
        return;
      }}
      const syncButton = document.getElementById('runSyncButton');
      syncButton.disabled = true;
      setOutput('Running sync...\\n\\nThis may take a few seconds while Plaid and Google Sheets respond.');
      const result = await api('/api/sync', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify({{ confirm: true }})
      }});
      syncButton.disabled = latestStatus && latestStatus.readiness ? !latestStatus.readiness.can_sync : false;
      const summaryLines = [
        result.ok ? 'Sync completed.' : 'Sync failed.',
        `Status: ${{result.summary.status}}`,
        `New transactions: ${{result.summary.new_transactions ?? 'unknown'}}`,
        `Cells appended: ${{result.summary.cells_appended}}`,
        '',
        result.output || 'No output.'
      ];
      setOutput(summaryLines.join('\\n'));
      await loadStatus(false);
    }}

    async function openSheet() {{
      const result = await api('/api/open-sheet', {{ method: 'POST' }});
      setOutput(result.message);
    }}

    async function checkAppsScriptDeploy() {{
      setOutput('Checking Apps Script deploy status...');
      const result = await api('/api/appscript/dry-run', {{ method: 'POST' }});
      setOutput(result.output || result);
      await loadStatus(false);
    }}

    async function deployAppsScript() {{
      if (!confirm('Deploy local Code.gs and Sidebar.html to the bound Apps Script project? Reload the Sheet and refresh visuals after this finishes.')) {{
        return;
      }}
      setOutput('Deploying Apps Script...');
      const result = await api('/api/appscript/deploy', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify({{ confirm: true }})
      }});
      setOutput(result.output || result);
      await loadStatus(false);
    }}

    function manualIncomePayload(confirmValue) {{
      return {{
        file_path: document.getElementById('manualIncomePath').value,
        account: document.getElementById('manualIncomeAccount').value,
        confirm: confirmValue === true
      }};
    }}

    async function dryRunManualIncomeImport() {{
      setOutput('Running manual income dry run...');
      const result = await api('/api/import/manual-income/dry-run', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify(manualIncomePayload(false))
      }});
      setOutput(result.output || result);
      await loadStatus(false);
    }}

    async function confirmManualIncomeImport() {{
      if (!confirm('Append reviewed manual income rows to Google Sheets? Run a dry run first.')) {{
        return;
      }}
      setOutput('Appending manual income rows...');
      const result = await api('/api/import/manual-income/confirm', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify(manualIncomePayload(true))
      }});
      setOutput(result.output || result);
      await loadStatus(false);
    }}

    async function copyChecklist() {{
      const parser = new DOMParser();
      const text = parser.parseFromString(checklist, 'text/html').documentElement.textContent;
      await navigator.clipboard.writeText(text);
      setOutput(text + '\\n\\nCopied to clipboard.');
    }}

    async function copySetupCommands() {{
      const text = setupCommands.map(item => `${{item.label}}\\n${{item.command}}\\n${{item.detail}}`).join('\\n\\n');
      await navigator.clipboard.writeText(text);
      setOutput(text + '\\n\\nCopied to clipboard.');
    }}

    async function copyRedactedDiagnostics() {{
      const result = await api('/api/diagnostics/redacted');
      const text = JSON.stringify(result, null, 2);
      await navigator.clipboard.writeText(text);
      setOutput(text + '\\n\\nCopied redacted diagnostics to clipboard.');
    }}

    loadStatus().catch(error => setOutput(error.message));
  </script>
</body>
</html>"""


class ControlCenterHandler(BaseHTTPRequestHandler):
    root = repo_root()

    def log_message(self, format, *args):
        return

    def _env(self):
        return load_control_env(self.root)

    def _send_json(self, payload, status=200):
        body = json.dumps(payload, indent=2).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get('Content-Length') or 0)
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode('utf-8') or '{}')

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/':
            body = render_control_center_html().encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == '/favicon.ico':
            self.send_response(204)
            self.end_headers()
            return
        if path == '/api/status':
            self._send_json(build_status_payload(self.root, self._env()))
            return
        if path == '/api/demo/status':
            self._send_json(build_demo_payload(self.root))
            return
        if path == '/api/trusted-tester/checklist':
            self._send_json({'text': build_trusted_tester_checklist()})
            return
        if path == '/api/diagnostics/redacted':
            env = self._env()
            self._send_json(build_redacted_diagnostics(build_status_payload(self.root, env), env))
            return
        if path == '/api/linked-accounts':
            self._send_json(collect_linked_accounts(self._env()))
            return
        if path == '/api/instructions/us-bank':
            self._send_json({'text': build_us_bank_guidance()})
            return
        if path == '/api/instructions/appscript':
            self._send_json({'text': build_appscript_redeploy_checklist()})
            return
        if path == '/api/instructions/quickstart':
            self._send_json({'text': build_quickstart_repair_command()})
            return
        if path == '/api/instructions/manual-income':
            self._send_json({'text': build_manual_income_import_guidance()})
            return
        self._send_json({'error': 'Not found'}, status=404)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            if path == '/api/doctor':
                payload = run_doctor_command(self.root, self._env(), skip_network=True)
                record_runtime_event(RUNTIME_STATE, 'doctor', payload)
                self._send_json(payload)
                return
            if path == '/api/sync':
                payload = self._read_json()
                readiness = build_readiness(self.root, self._env(), build_doctor_payload(self.root, self._env(), skip_network=True), collect_linked_accounts(self._env()))
                if not readiness['can_sync']:
                    recommended = readiness.get('recommended_next_step') or {}
                    raise ControlCenterError('Sync is blocked until setup is ready. Next step: ' + (recommended.get('title') or 'complete setup readiness items'))
                result = run_sync_command(
                    confirm=payload.get('confirm') is True,
                    root=self.root,
                    env=self._env(),
                )
                record_runtime_event(RUNTIME_STATE, 'sync', result)
                self._send_json(result)
                return
            if path == '/api/open-sheet':
                webbrowser.open(build_sheet_open_url(self._env()))
                self._send_json({'ok': True, 'message': 'Google Sheet open request sent to the local browser.'})
                return
            if path == '/api/appscript/dry-run':
                result = run_appscript_dry_run(root=self.root, env=self._env())
                record_runtime_event(RUNTIME_STATE, 'appscript_deploy', result)
                self._send_json(result)
                return
            if path == '/api/appscript/deploy':
                payload = self._read_json()
                result = run_appscript_deploy(
                    confirm=payload.get('confirm') is True,
                    root=self.root,
                    env=self._env(),
                )
                record_runtime_event(RUNTIME_STATE, 'appscript_deploy', result)
                self._send_json(result)
                return
            if path == '/api/import/manual-income/dry-run':
                payload = self._read_json()
                result = run_manual_income_import_command(
                    root=self.root,
                    env=self._env(),
                    file_path=payload.get('file_path') or 'src/imports/income.csv',
                    account=payload.get('account') or DEFAULT_MANUAL_INCOME_ACCOUNT,
                    dry_run=True,
                    confirm=False,
                )
                record_runtime_event(RUNTIME_STATE, 'import', result)
                self._send_json(result)
                return
            if path == '/api/import/manual-income/confirm':
                payload = self._read_json()
                result = run_manual_income_import_command(
                    root=self.root,
                    env=self._env(),
                    file_path=payload.get('file_path') or 'src/imports/income.csv',
                    account=payload.get('account') or DEFAULT_MANUAL_INCOME_ACCOUNT,
                    dry_run=False,
                    confirm=payload.get('confirm') is True,
                )
                record_runtime_event(RUNTIME_STATE, 'import', result)
                self._send_json(result)
                return
            self._send_json({'error': 'Not found'}, status=404)
        except ControlCenterError as exc:
            self._send_json({'error': str(exc)}, status=400)
        except Exception as exc:
            self._send_json({'error': str(exc)}, status=500)


def parse_args():
    parser = argparse.ArgumentParser(description='Run the local Danny Bank control center.')
    parser.add_argument('--host', default=DEFAULT_HOST, help='Bind host. Use 127.0.0.1 for local-only access.')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, help='Bind port.')
    parser.add_argument('--no-open', action='store_true', help='Do not open the browser automatically.')
    return parser.parse_args()


def main():
    args = parse_args()
    if args.host not in ('127.0.0.1', 'localhost'):
        print('Refusing to bind to a non-local host. Use 127.0.0.1.')
        return 2

    ControlCenterHandler.root = repo_root()
    server = ThreadingHTTPServer((args.host, args.port), ControlCenterHandler)
    url = f'http://{args.host}:{args.port}/'
    print(f'Danny Bank Control Center: {url}')
    print('Press Ctrl+C to stop.')
    if not args.no_open:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('')
        print('Control center stopped.')
    finally:
        server.server_close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
