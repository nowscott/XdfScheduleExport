#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ -d .venv ]]; then
  source .venv/bin/activate
elif [[ -d venv ]]; then
  source venv/bin/activate
else
  python3 -m venv .venv
  source .venv/bin/activate
  python -m pip install -r requirements.txt
  python -m playwright install chromium
fi

python main.py "$@"
