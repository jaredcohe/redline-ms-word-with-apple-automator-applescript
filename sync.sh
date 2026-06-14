#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BUNDLE="$HOME/Library/Services/Redline with Word.workflow/Contents/Resources"

if [ ! -d "$BUNDLE" ]; then
    echo "Error: workflow bundle not found."
    echo "Run install.sh first to create it."
    exit 1
fi

cp "$SCRIPT_DIR/clean_redline.py" "$BUNDLE/"
cp "$SCRIPT_DIR/normalize_docx.py" "$BUNDLE/"
echo "Scripts synced to workflow bundle."
