import argparse
import importlib
import os
import subprocess
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv
from googleapiclient.discovery import build

from .env_tokens import split_access_tokens
from .plaid_client import PlaidClient
from .sheets_client import SheetsClient


CURRENT_ROOT = Path('/Users/michaelpanico/Desktop/DevBase/active_projects/Danny_Bank_Automation')
QUICKSTART_PYTHON_DIR = Path('/Users/michaelpanico/Desktop/quickstart/python')
OLD_MAC_PRO_PATH = '/Users/michael_s_panico'
PLAID_OAUTH_STATUS_URL = 'https://dashboard.plaid.com/activity/status/oauth-institutions'


@dataclass
class CheckResult:
    name: str
    status: str
    detail: str


def repo_root():
    override = os.environ.get('DANNY_BANK_HOME')
    if override:
        return Path(override).expanduser().resolve()
    return Path(__file__).resolve().parents[2]


def format_result(result):
    return f'[{result.status}] {result.name}: {result.detail}'


def check_required_env(env):
    required = ['PLAID_CLIENT_ID', 'PLAID_SECRET', 'PLAID_ENV', 'PLAID_ACCESS_TOKEN', 'GOOGLE_SPREADSHEET_ID']
    missing = [key for key in required if not env.get(key)]
    if missing:
        return CheckResult('env', 'FAIL', 'Missing required key(s): ' + ', '.join(missing))

    token_count = len(split_access_tokens(env.get('PLAID_ACCESS_TOKEN')))
    return CheckResult('env', 'PASS', f'Required keys present; {token_count} Plaid access token(s) configured.')


def check_python_imports():
    modules = ['dotenv', 'googleapiclient', 'google_auth_oauthlib', 'plaid', 'requests']
    missing = []
    for module in modules:
        try:
            importlib.import_module(module)
        except Exception:
            missing.append(module)
    if missing:
        return CheckResult('python imports', 'FAIL', 'Missing import(s): ' + ', '.join(missing))
    return CheckResult('python imports', 'PASS', 'Required runtime packages import successfully.')


def quickstart_venv_findings(venv_dir):
    venv = Path(venv_dir)
    findings = []
    if not venv.exists():
        return ['Quickstart venv is missing; rebuild it before using Quickstart fallback.']

    files_to_scan = [
        venv / 'pyvenv.cfg',
        venv / 'bin' / 'activate',
        venv / 'bin' / 'pip',
        venv / 'bin' / 'pip3',
    ]
    for path in files_to_scan:
        if not path.exists():
            continue
        try:
            if OLD_MAC_PRO_PATH in path.read_text(errors='ignore'):
                findings.append(f'{path} contains old Mac Pro path {OLD_MAC_PRO_PATH}.')
        except OSError as exc:
            findings.append(f'Could not read {path}: {exc}')

    python_bin = venv / 'bin' / 'python'
    if not python_bin.exists():
        findings.append(f'{python_bin} does not exist.')
    elif python_bin.is_symlink():
        target = python_bin.resolve(strict=False)
        if not target.exists():
            findings.append(f'{python_bin} points to missing interpreter {target}.')

    return findings


def check_quickstart_venv(quickstart_python_dir=QUICKSTART_PYTHON_DIR):
    findings = quickstart_venv_findings(Path(quickstart_python_dir) / 'venv')
    if findings:
        return CheckResult('quickstart venv', 'WARN', ' '.join(findings))
    return CheckResult('quickstart venv', 'PASS', 'Quickstart venv does not contain known stale-path markers.')


def check_appscript_syntax(root):
    code_path = Path(root) / 'src' / 'appscript' / 'Code.gs'
    if not code_path.exists():
        return CheckResult('Apps Script syntax', 'FAIL', f'Missing {code_path}.')

    try:
        proc = subprocess.run(
            ['node', '--check', '--input-type=commonjs'],
            input=code_path.read_text(),
            text=True,
            capture_output=True,
            timeout=20,
        )
    except FileNotFoundError:
        return CheckResult('Apps Script syntax', 'WARN', 'node is not installed; syntax check skipped.')
    except Exception as exc:
        return CheckResult('Apps Script syntax', 'FAIL', f'node syntax check failed to run: {exc}')

    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or 'node returned non-zero').strip().splitlines()[-1]
        return CheckResult('Apps Script syntax', 'FAIL', detail)
    return CheckResult('Apps Script syntax', 'PASS', 'Code.gs passes node syntax check.')


def check_appscript_deploy_config(env):
    if env.get('GOOGLE_APPS_SCRIPT_ID'):
        return CheckResult('Apps Script deploy config', 'PASS', 'GOOGLE_APPS_SCRIPT_ID is configured for API-based deploy checks.')
    return CheckResult(
        'Apps Script deploy config',
        'WARN',
        'GOOGLE_APPS_SCRIPT_ID is missing. Use manual paste deploy fallback, or add the bound Apps Script project ID to enable the deploy helper.',
    )


def check_plaid_accounts(env):
    tokens = split_access_tokens(env.get('PLAID_ACCESS_TOKEN'))
    if not tokens:
        return CheckResult('Plaid accounts', 'FAIL', 'No Plaid access tokens configured.')
    client = PlaidClient(env.get('PLAID_CLIENT_ID'), env.get('PLAID_SECRET'), env.get('PLAID_ENV', 'sandbox'))

    institution_names = []
    failed = 0
    for token in tokens:
        item = client.get_item(token)
        institution_id = item.get('institution_id')
        name = client.get_institution_name(institution_id)
        accounts = client.get_accounts(token)
        if not accounts:
            failed += 1
            continue
        institution_names.append(name)

    if failed:
        return CheckResult('Plaid accounts', 'WARN', f'{failed} Plaid item(s) returned no accounts; found: {", ".join(institution_names) or "none"}.')
    return CheckResult('Plaid accounts', 'PASS', f'Linked institutions reachable: {", ".join(institution_names)}.')


def check_google_sheet(env):
    spreadsheet_id = env.get('GOOGLE_SPREADSHEET_ID')
    if not spreadsheet_id:
        return CheckResult('Google Sheets', 'FAIL', 'GOOGLE_SPREADSHEET_ID is missing.')
    try:
        creds = SheetsClient(spreadsheet_id, env.get('GOOGLE_SHEET_NAME', 'Transactions')).authenticate()
        service = build('sheets', 'v4', credentials=creds)
        result = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields='properties.title',
        ).execute()
        title = result.get('properties', {}).get('title') or spreadsheet_id
        return CheckResult('Google Sheets', 'PASS', f'Reachable spreadsheet: {title}.')
    except Exception as exc:
        return CheckResult('Google Sheets', 'FAIL', f'Could not reach configured spreadsheet: {exc}')


def plaid_oauth_blocker_note():
    return CheckResult(
        'Plaid OAuth blockers',
        'WARN',
        'U.S. Bank and other OAuth-gated institutions may require Plaid Production/OAuth institution registration. Check ' + PLAID_OAUTH_STATUS_URL,
    )


def collect_checks(root, env, quickstart_python_dir, skip_network=False):
    env_check = check_required_env(env)
    checks = [
        env_check,
        check_python_imports(),
        check_appscript_syntax(root),
        check_appscript_deploy_config(env),
        check_quickstart_venv(quickstart_python_dir),
        plaid_oauth_blocker_note(),
    ]
    if skip_network:
        checks.append(CheckResult('network checks', 'WARN', 'Skipped Plaid and Google network checks.'))
    elif env_check.status == 'FAIL':
        checks.append(CheckResult('network checks', 'WARN', 'Skipped Plaid and Google network checks because required .env keys are missing.'))
    else:
        checks.append(check_plaid_accounts(env))
        checks.append(check_google_sheet(env))
    return checks


def parse_args():
    parser = argparse.ArgumentParser(description='Run local setup and integration health checks.')
    parser.add_argument('--skip-network', action='store_true', help='Skip Plaid and Google Sheets network checks.')
    parser.add_argument('--quickstart-python-dir', default=str(QUICKSTART_PYTHON_DIR), help='Path to Plaid Quickstart python directory.')
    return parser.parse_args()


def main():
    args = parse_args()
    root = repo_root()
    env_path = root / '.env'
    load_dotenv(env_path)
    env = dict(os.environ)

    print('Danny Bank Automation Doctor')
    print(f'Project root: {root}')
    print('')

    checks = collect_checks(root, env, Path(args.quickstart_python_dir), skip_network=args.skip_network)
    for result in checks:
        print(format_result(result))

    failed = [result for result in checks if result.status == 'FAIL']
    if failed:
        print('')
        print(f'{len(failed)} failing check(s) need attention.')
        return 1

    print('')
    print('Doctor finished without failing checks.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
