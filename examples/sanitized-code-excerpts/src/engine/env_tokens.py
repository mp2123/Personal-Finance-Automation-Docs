from pathlib import Path


TOKEN_KEY = 'PLAID_ACCESS_TOKEN'


def split_access_tokens(value):
    """Returns clean Plaid access tokens from a comma-separated env value."""
    return [token.strip() for token in str(value or '').split(',') if token.strip()]


def mask_token(token):
    """Masks an access token for logs and terminal output."""
    value = str(token or '')
    if len(value) <= 14:
        return '<masked>'
    return value[:8] + '...' + value[-6:]


def append_plaid_access_token(env_path, new_token):
    """Appends a Plaid access token to .env while preserving unrelated lines."""
    path = Path(env_path)
    token = str(new_token or '').strip()
    if not token:
        raise ValueError('No Plaid access token was provided.')

    original = path.read_text() if path.exists() else ''
    lines = original.splitlines()
    has_trailing_newline = original.endswith('\n') or original == ''
    token_line_index = None
    current_tokens = []

    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith(TOKEN_KEY + '='):
            token_line_index = index
            current_tokens = split_access_tokens(line.split('=', 1)[1])
            break

    if token in current_tokens:
        return {
            'changed': False,
            'reason': 'duplicate',
            'tokens': current_tokens,
        }

    updated_tokens = current_tokens + [token]
    updated_line = TOKEN_KEY + '=' + ','.join(updated_tokens)

    if token_line_index is None:
        if lines and lines[-1].strip():
            lines.append(updated_line)
        elif lines:
            lines[-1] = updated_line
        else:
            lines = [updated_line]
    else:
        lines[token_line_index] = updated_line

    output = '\n'.join(lines)
    if has_trailing_newline:
        output += '\n'
    path.write_text(output)

    return {
        'changed': True,
        'reason': 'appended',
        'tokens': updated_tokens,
    }
