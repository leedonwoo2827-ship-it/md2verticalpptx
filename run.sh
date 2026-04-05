#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ ! -d "$SCRIPT_DIR/.venv" ]; then
    echo "[ERROR] Virtual environment not found. Run install.sh first."
    exit 1
fi

source "$SCRIPT_DIR/.venv/bin/activate"
python -m md2pptx "$@"
