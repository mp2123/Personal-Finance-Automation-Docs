import argparse
import json
from pathlib import Path

from dotenv import load_dotenv

from .control_center import (
    build_redacted_diagnostics,
    build_status_payload,
    load_control_env,
)
from .doctor import repo_root


def build_diagnostics_payload(root=None, env=None, status_builder=build_status_payload):
    root = Path(root or repo_root())
    load_dotenv(root / '.env')
    env = env or load_control_env(root)
    status = status_builder(root=root, env=env, runtime_state={})
    return build_redacted_diagnostics(status, env)


def format_diagnostics_json(payload):
    return json.dumps(payload, indent=2, sort_keys=True)


def parse_args():
    parser = argparse.ArgumentParser(description='Print a redacted Danny Bank diagnostics packet for trusted beta support.')
    parser.add_argument('--output', help='Optional path to write the redacted diagnostics JSON.')
    return parser.parse_args()


def main():
    args = parse_args()
    payload = build_diagnostics_payload()
    output = format_diagnostics_json(payload)

    if args.output:
        path = Path(args.output).expanduser()
        path.write_text(output + '\n')
        print(f'Redacted diagnostics written to {path}')
        return 0

    print(output)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
