import argparse
import hashlib
import json
import os
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from .doctor import repo_root


APPS_SCRIPT_SCOPE = 'https://www.googleapis.com/auth/script.projects'
POST_DEPLOY_NEXT_STEP = 'Reload the Google Sheet, then run Bank Automation -> Refresh Dashboard & Visuals.'
MANAGED_FILES = {
    'Code': {
        'path': Path('src/appscript/Code.gs'),
        'type': 'SERVER_JS',
    },
    'Sidebar': {
        'path': Path('src/appscript/Sidebar.html'),
        'type': 'HTML',
    },
}


class AppScriptDeployError(Exception):
    pass


def short_hash(value):
    return hashlib.sha256(str(value or '').encode('utf-8')).hexdigest()[:12]


def load_local_sources(root=None):
    root = Path(root or repo_root())
    sources = {}
    for name, config in MANAGED_FILES.items():
        path = root / config['path']
        if not path.exists():
            raise AppScriptDeployError(f'Missing local Apps Script file: {path}')
        sources[name] = {
            'name': name,
            'type': config['type'],
            'source': path.read_text(),
        }

    manifest_path = root / 'src/appscript/appsscript.json'
    if manifest_path.exists():
        sources['appsscript'] = {
            'name': 'appsscript',
            'type': 'JSON',
            'source': manifest_path.read_text(),
        }
    return sources


def remote_file_map(remote_files):
    return {item.get('name'): item for item in remote_files or [] if item.get('name')}


def unmanaged_remote_files(remote_files):
    managed = set(MANAGED_FILES.keys()) | {'appsscript'}
    return sorted(name for name in remote_file_map(remote_files) if name not in managed)


def build_update_payload(local_sources, remote_files, allow_unmanaged=False):
    remote = remote_file_map(remote_files)
    unmanaged = unmanaged_remote_files(remote_files)
    if unmanaged and not allow_unmanaged:
        raise AppScriptDeployError(
            'Remote Apps Script project contains unmanaged file(s): ' +
            ', '.join(unmanaged) +
            '. Re-run with --allow-unmanaged only if these files should be preserved.'
        )

    files = []
    for name in MANAGED_FILES:
        source = local_sources[name]
        files.append({
            'name': name,
            'type': source['type'],
            'source': source['source'],
        })

    if 'appsscript' in local_sources:
        files.append(local_sources['appsscript'])
    elif 'appsscript' in remote:
        manifest = remote['appsscript']
        files.append({
            'name': 'appsscript',
            'type': manifest.get('type') or 'JSON',
            'source': manifest.get('source') or '{}',
        })

    if allow_unmanaged:
        for name in unmanaged:
            files.append(remote[name])

    return {'files': files}


def compare_managed_files(local_sources, remote_files):
    remote = remote_file_map(remote_files)
    comparison = {}
    for name in MANAGED_FILES:
        local_source = local_sources[name]['source']
        remote_source = (remote.get(name) or {}).get('source')
        if remote_source is None:
            status = 'remote_missing'
        elif remote_source == local_source:
            status = 'unchanged'
        else:
            status = 'changed'
        comparison[name] = {
            'status': status,
            'local_hash': short_hash(local_source),
            'remote_hash': short_hash(remote_source) if remote_source is not None else None,
            'type': local_sources[name]['type'],
        }
    return comparison


def manual_deploy_fallback_message():
    return '\n'.join([
        'Manual Apps Script fallback:',
        '1. Open the bound Apps Script project from the Google Sheet.',
        '2. Replace Code.gs with src/appscript/Code.gs from this repo.',
        '3. Replace Sidebar.html with src/appscript/Sidebar.html from this repo.',
        '4. Save, reload the Sheet, then run Bank Automation -> Refresh Dashboard & Visuals.',
    ])


def authenticate_apps_script(root=None):
    root = Path(root or repo_root())
    token_path = root / 'token_appscript.json'
    credentials_path = root / 'credentials.json'
    scopes = [APPS_SCRIPT_SCOPE]
    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not credentials_path.exists():
                raise AppScriptDeployError(f'Missing credentials.json at {credentials_path}')
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), scopes)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json())

    return creds


def build_script_service(root=None):
    creds = authenticate_apps_script(root)
    return build('script', 'v1', credentials=creds)


def run_deploy_plan(env, dry_run=True, service_factory=None, root=None, allow_unmanaged=False, confirmed=False):
    root = Path(root or repo_root())
    script_id = env.get('GOOGLE_APPS_SCRIPT_ID') or ''
    if not script_id:
        return {
            'ok': False,
            'reason': 'missing_script_id',
            'message': 'GOOGLE_APPS_SCRIPT_ID is missing. Add the bound Apps Script project ID to .env to use deploy helper.',
            'fallback': manual_deploy_fallback_message(),
            'next_step': manual_deploy_fallback_message(),
        }

    if not dry_run and not confirmed:
        raise AppScriptDeployError('Apps Script deployment requires explicit confirmation.')

    local_sources = load_local_sources(root)
    service = service_factory() if service_factory else build_script_service(root)
    remote_payload = service.projects().getContent(scriptId=script_id).execute()
    remote_files = remote_payload.get('files', [])
    comparison = compare_managed_files(local_sources, remote_files)
    payload = build_update_payload(local_sources, remote_files, allow_unmanaged=allow_unmanaged)

    report = {
        'ok': True,
        'script_id': script_id,
        'dry_run': dry_run,
        'updated': False,
        'comparison': comparison,
        'unmanaged_remote_files': unmanaged_remote_files(remote_files),
        'managed_files': list(MANAGED_FILES.keys()),
        'next_step': POST_DEPLOY_NEXT_STEP,
    }
    if dry_run:
        return report

    service.projects().updateContent(scriptId=script_id, body=payload).execute()
    report['updated'] = True
    return report


def mask_report_text(text, env):
    masked = str(text)
    for key in ['GOOGLE_APPS_SCRIPT_ID', 'GOOGLE_SPREADSHEET_ID', 'PLAID_SECRET', 'PLAID_ACCESS_TOKEN', 'GEMINI_API_KEY']:
        value = env.get(key)
        if value and len(str(value)) >= 4:
            masked = masked.replace(str(value), '[masked]')
    return masked


def format_deploy_report(report, env=None):
    env = env or {}
    lines = ['Apps Script Deployment Helper']
    lines.append(f'Required scope: {APPS_SCRIPT_SCOPE}')
    if not report.get('ok'):
        lines.append(report.get('message') or 'Deployment helper is not ready.')
        if report.get('fallback'):
            lines.append('')
            lines.append(report['fallback'])
        return mask_report_text('\n'.join(lines), env)

    lines.append(f"Script ID: {report.get('script_id')}")
    lines.append('Mode: dry run' if report.get('dry_run') else 'Mode: push')
    lines.append('')
    lines.append('Managed files:')
    for name, info in (report.get('comparison') or {}).items():
        lines.append(f"- {name}: {info['status']} local={info['local_hash']} remote={info.get('remote_hash') or 'missing'}")
    unmanaged = report.get('unmanaged_remote_files') or []
    if unmanaged:
        lines.append('')
        lines.append('Unmanaged remote files: ' + ', '.join(unmanaged))
    lines.append('')
    lines.append('Next step: ' + report.get('next_step', POST_DEPLOY_NEXT_STEP))
    return mask_report_text('\n'.join(lines), env)


def format_deploy_error(error, env=None):
    env = env or {}
    lines = [
        'Apps Script deployment failed: ' + mask_report_text(str(error), env),
        '',
        manual_deploy_fallback_message(),
    ]
    return mask_report_text('\n'.join(lines), env)


def parse_args():
    parser = argparse.ArgumentParser(description='Deploy local Apps Script files to the bound Apps Script project.')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--dry-run', action='store_true', help='Show what would change without updating Apps Script.')
    mode.add_argument('--push', action='store_true', help='Update the remote Apps Script project.')
    parser.add_argument('--yes', action='store_true', help='Confirm push non-interactively.')
    parser.add_argument('--allow-unmanaged', action='store_true', help='Preserve unmanaged remote files instead of refusing.')
    return parser.parse_args()


def main():
    args = parse_args()
    root = repo_root()
    load_dotenv(root / '.env')
    env = dict(os.environ)
    dry_run = not args.push

    if args.push and not args.yes:
        answer = input('Deploy local Code.gs and Sidebar.html to the configured Apps Script project? Type YES: ')
        if answer != 'YES':
            print('Deployment canceled.')
            return 1

    try:
        report = run_deploy_plan(
            env,
            dry_run=dry_run,
            root=root,
            allow_unmanaged=args.allow_unmanaged,
            confirmed=(args.yes or not dry_run),
        )
        print(format_deploy_report(report, env))
        return 0 if report.get('ok') or report.get('reason') == 'missing_script_id' else 1
    except AppScriptDeployError as exc:
        print(format_deploy_error(exc, env))
        return 1
    except Exception as exc:
        print(format_deploy_error(exc, env))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
