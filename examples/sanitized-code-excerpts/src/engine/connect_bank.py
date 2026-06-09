import argparse
import json
import os
import socket
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from dotenv import load_dotenv
from plaid.exceptions import ApiException as PlaidApiException

from .account_labels import build_account_label
from .env_tokens import append_plaid_access_token, mask_token
from .plaid_client import PlaidClient


DEFAULT_PORT = 8765
OAUTH_STATUS_URL = 'https://dashboard.plaid.com/activity/status/oauth-institutions'


class PlaidInstitutionRegistrationRequired(RuntimeError):
    """Raised when Plaid blocks an OAuth institution pending registration."""

    def __init__(self, institution='', message='', link_session_id=''):
        self.institution = institution or 'the selected institution'
        self.message = message or 'Plaid institution registration is required.'
        self.link_session_id = link_session_id or ''
        super().__init__(self.message)


def _repo_root():
    return Path(__file__).resolve().parents[2]


def _find_free_port(preferred_port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        result = sock.connect_ex(('127.0.0.1', preferred_port))
    if result != 0:
        return preferred_port

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(('127.0.0.1', 0))
        return sock.getsockname()[1]


def _read_request_body(handler):
    length = int(handler.headers.get('Content-Length') or 0)
    return handler.rfile.read(length).decode('utf-8') if length else ''


def _link_page(link_token):
    token_json = json.dumps(link_token)
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Danny Bank Automation - Connect Bank</title>
  <script src="https://cdn.plaid.com/link/v2/stable/link-initialize.js"></script>
  <style>
    body {{
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #0d1117;
      color: #f8fafc;
    }}
    main {{
      width: min(560px, calc(100vw - 32px));
      border: 1px solid #334155;
      border-radius: 8px;
      padding: 28px;
      background: #111827;
    }}
    h1 {{
      margin: 0 0 10px;
      font-size: 22px;
    }}
    p {{
      color: #cbd5e1;
      line-height: 1.5;
    }}
    button {{
      border: 0;
      border-radius: 6px;
      padding: 12px 16px;
      font-weight: 700;
      color: white;
      background: #2563eb;
      cursor: pointer;
    }}
    #status {{
      margin-top: 14px;
      color: #93c5fd;
      font-size: 14px;
    }}
  </style>
</head>
<body>
  <main>
    <h1>Connect a Plaid institution</h1>
    <p>Use Plaid Link to connect the U.S. Bank credit card. After Link succeeds, return to Terminal to review the account preview before anything is saved.</p>
    <button id="open-link">Open Plaid Link</button>
    <div id="status">Ready.</div>
  </main>
  <script>
    const statusEl = document.getElementById('status');
    const handler = Plaid.create({{
      token: {token_json},
      onSuccess: async (public_token, metadata) => {{
        statusEl.textContent = 'Link succeeded. Sending public token to local connector...';
        await fetch('/public_token', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{ public_token, metadata }})
        }});
        statusEl.textContent = 'Done. Return to Terminal.';
      }},
      onEvent: async (eventName, metadata) => {{
        await fetch('/event', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{ eventName, metadata }})
        }});
      }},
      onExit: async (err, metadata) => {{
        statusEl.textContent = err ? 'Plaid Link exited with an error. Return to Terminal.' : 'Plaid Link was closed. Return to Terminal.';
        await fetch('/exit', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{ error: err, metadata }})
        }});
      }}
    }});
    document.getElementById('open-link').onclick = () => handler.open();
    window.addEventListener('load', () => setTimeout(() => handler.open(), 500));
  </script>
</body>
</html>"""


def _make_handler(state, link_token):
    class ConnectBankHandler(BaseHTTPRequestHandler):
        def log_message(self, format, *args):
            return

        def _send_text(self, status, body, content_type='text/plain'):
            payload = body.encode('utf-8')
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def do_GET(self):
            parsed = urlparse(self.path)
            if parsed.path == '/':
                self._send_text(200, _link_page(link_token), 'text/html')
                return
            if parsed.path == '/oauth-return':
                self._send_text(200, _link_page(link_token), 'text/html')
                return
            self._send_text(404, 'Not found')

        def do_POST(self):
            parsed = urlparse(self.path)
            body = _read_request_body(self)
            data = {}
            if body:
                try:
                    data = json.loads(body)
                except json.JSONDecodeError:
                    data = {key: values[0] for key, values in parse_qs(body).items()}

            if parsed.path == '/public_token':
                state['public_token'] = data.get('public_token') or ''
                state['metadata'] = data.get('metadata') or {}
                state['done'].set()
                self._send_text(200, 'OK')
                return

            if parsed.path == '/exit':
                state['exit'] = data
                state['done'].set()
                self._send_text(200, 'OK')
                return

            if parsed.path == '/event':
                event_name = data.get('eventName') or 'UNKNOWN_EVENT'
                metadata = data.get('metadata') or {}
                institution = metadata.get('institution_name') or metadata.get('institution_id') or ''
                error_code = metadata.get('error_code') or ''
                error_message = metadata.get('error_message') or ''
                link_session_id = metadata.get('link_session_id') or ''
                view_name = metadata.get('view_name') or ''
                details = ' | '.join(part for part in [view_name, institution, error_code, error_message, link_session_id] if part)
                print(f'Plaid Link event: {event_name}' + (f' | {details}' if details else ''))
                if error_code:
                    state['errors'].append({
                        'event_name': event_name,
                        'institution': institution,
                        'error_code': error_code,
                        'error_message': error_message,
                        'link_session_id': link_session_id,
                    })
                self._send_text(200, 'OK')
                return

            self._send_text(404, 'Not found')

    return ConnectBankHandler


def _run_link_flow(link_token, port, no_open, timeout_seconds):
    state = {
        'public_token': '',
        'metadata': {},
        'exit': None,
        'errors': [],
        'done': threading.Event(),
    }
    handler = _make_handler(state, link_token)
    server = ThreadingHTTPServer(('127.0.0.1', port), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    url = f'http://127.0.0.1:{port}/'

    print(f'Local Plaid Link page: {url}')
    if not no_open:
        webbrowser.open(url)
    else:
        print('Open the URL above in your browser.')

    state['done'].wait(timeout_seconds)
    server.shutdown()
    server.server_close()

    if state['public_token']:
        return state['public_token'], state['metadata']

    if state['exit']:
        registration_error = find_registration_required_error_(state)
        if registration_error:
            raise PlaidInstitutionRegistrationRequired(
                institution=registration_error.get('institution'),
                message=registration_error.get('error_message'),
                link_session_id=registration_error.get('link_session_id'),
            )
        raise RuntimeError('Plaid Link was closed before a bank was connected.')

    raise TimeoutError('Timed out waiting for Plaid Link to finish.')


def find_registration_required_error_(state):
    """Returns Plaid registration error metadata from Link state if present."""
    for event in state.get('errors') or []:
        if event.get('error_code') == 'INSTITUTION_REGISTRATION_REQUIRED':
            return event

    exit_data = state.get('exit') or {}
    candidates = [
        exit_data.get('error') if isinstance(exit_data, dict) else {},
        exit_data.get('metadata') if isinstance(exit_data, dict) else {},
    ]
    for candidate in candidates:
        if not isinstance(candidate, dict):
            continue
        if candidate.get('error_code') == 'INSTITUTION_REGISTRATION_REQUIRED':
            return {
                'institution': candidate.get('institution_name') or candidate.get('institution_id') or '',
                'error_code': candidate.get('error_code'),
                'error_message': candidate.get('error_message') or candidate.get('display_message') or '',
                'link_session_id': candidate.get('link_session_id') or '',
            }
    return None


def format_registration_required_guidance_(error):
    """Builds user-facing guidance for Plaid OAuth institution registration blockers."""
    institution = getattr(error, 'institution', '') or 'the selected institution'
    message = getattr(error, 'message', '') or 'Plaid requires additional OAuth institution registration before this institution can be connected.'
    link_session_id = getattr(error, 'link_session_id', '') or ''
    lines = [
        '',
        'Plaid blocked this connection because institution registration is required.',
        f'Institution: {institution}',
        f'Reason: {message}',
    ]
    if link_session_id:
        lines.append(f'Link session ID: {link_session_id}')
    lines.extend([
        '',
        'Next step: wait for Plaid Production/OAuth institution approval, or complete the required registration steps in Plaid Dashboard:',
        f'  {OAUTH_STATUS_URL}',
        '',
        '.env was not changed.',
    ])
    return '\n'.join(lines)


def _print_account_preview(description):
    institution_name = description.get('institution_name') or 'Unknown Institution'
    accounts = description.get('accounts') or []
    print('')
    print(f'Institution: {institution_name}')
    if not accounts:
        print('  - No accounts returned by Plaid.')
        return

    for account in accounts:
        label = build_account_label(institution_name, account)
        account_type = account.get('type') or 'unknown'
        subtype = account.get('subtype') or 'unknown'
        print(f'  - {label} | {account_type}/{subtype}')


def _confirm_append():
    answer = input('\nAppend this new Plaid access token to .env? Type YES to confirm: ')
    return answer.strip() == 'YES'


def parse_args():
    parser = argparse.ArgumentParser(description='Connect a new Plaid institution and append its access token after confirmation.')
    parser.add_argument('--institution-note', default='', help='Human note for terminal output, such as "U.S. Bank".')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, help='Preferred local port for the browser flow.')
    parser.add_argument('--no-open', action='store_true', help='Print the local URL instead of opening the browser automatically.')
    parser.add_argument('--dry-run', action='store_true', help='Complete Link and preview accounts, but do not modify .env.')
    parser.add_argument('--redirect-uri', default='', help='Optional Plaid OAuth redirect URI registered in Plaid Dashboard.')
    parser.add_argument('--timeout-seconds', type=int, default=600, help='Seconds to wait for Plaid Link to finish.')
    return parser.parse_args()


def main():
    args = parse_args()
    root = _repo_root()
    env_path = root / '.env'
    load_dotenv(env_path)

    client_id = os.getenv('PLAID_CLIENT_ID')
    secret = os.getenv('PLAID_SECRET')
    env = os.getenv('PLAID_ENV', 'sandbox')
    if not client_id or not secret:
        print('Missing PLAID_CLIENT_ID or PLAID_SECRET in .env.')
        return 1

    note = f' for {args.institution_note}' if args.institution_note else ''
    print(f'Creating Plaid Link token{note}...')
    print(f'Plaid env: {env}')
    print('Plaid products: transactions')

    plaid_client = PlaidClient(client_id, secret, env)
    port = _find_free_port(args.port)
    redirect_uri = args.redirect_uri.strip()

    try:
        link_token = plaid_client.create_transactions_link_token(
            client_user_id='danny-bank-automation-local',
            redirect_uri=redirect_uri or None,
        )
        if not link_token:
            print('Plaid did not return a link token.')
            return 1

        public_token, metadata = _run_link_flow(
            link_token=link_token,
            port=port,
            no_open=args.no_open,
            timeout_seconds=args.timeout_seconds,
        )
        institution = metadata.get('institution') or {}
        if institution.get('name'):
            print(f"Plaid Link completed: {institution.get('name')}")

        exchange = plaid_client.exchange_public_token(public_token)
        access_token = exchange.get('access_token')
        if not access_token:
            print('Plaid did not return an access token.')
            return 1

        description = plaid_client.describe_access_token(access_token)
        _print_account_preview(description)
        print(f'\nNew token received: {mask_token(access_token)}')

        if args.dry_run:
            print('Dry run enabled. .env was not changed.')
            return 0

        if not _confirm_append():
            print('.env was not changed.')
            return 1

        result = append_plaid_access_token(env_path, access_token)
        if not result['changed'] and result['reason'] == 'duplicate':
            print('This access token is already present in .env. Nothing changed.')
        else:
            print('Appended new Plaid access token to .env.')

        print('')
        print('Next commands:')
        print('  .venv/bin/python -m src.engine.list_linked_accounts')
        print('  .venv/bin/python -m src.engine.main')
        print('Then refresh the Google Sheet: Bank Automation -> Refresh Dashboard & Visuals')
        return 0
    except PlaidApiException as exc:
        print(f'Plaid API error: {exc.body}')
        return 1
    except PlaidInstitutionRegistrationRequired as exc:
        print(format_registration_required_guidance_(exc))
        return 1
    except KeyboardInterrupt:
        print('\nCanceled. .env was not changed.')
        return 1
    except Exception as exc:
        print(f'Connection failed: {exc}')
        print('.env was not changed.')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
