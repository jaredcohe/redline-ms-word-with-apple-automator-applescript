#!/bin/bash
set -euo pipefail

WORKFLOW_NAME="Redline with Word"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APPLESCRIPT="$SCRIPT_DIR/applescript-v2"
WORKFLOW_DIR="$HOME/Library/Services/${WORKFLOW_NAME}.workflow"
CONTENTS_DIR="$WORKFLOW_DIR/Contents"
LSREG="/System/Library/Frameworks/CoreServices.framework/Versions/A/Frameworks/LaunchServices.framework/Versions/A/Support/lsregister"

if [ ! -f "$APPLESCRIPT" ]; then
    echo "Error: applescript-v2 not found in $SCRIPT_DIR"
    exit 1
fi

echo "Installing '${WORKFLOW_NAME}'..."

echo "Running tests..."
(cd "$SCRIPT_DIR" && python3 -m unittest tests/test_clean_redline.py -q) || { echo "Tests failed — aborting install."; exit 1; }
echo ""

# Remove any previous version
if [ -d "$WORKFLOW_DIR" ]; then
    rm -rf "$WORKFLOW_DIR"
fi
mkdir -p "$CONTENTS_DIR"

# Generate Info.plist and document.wflow from the AppleScript source
python3 "$SCRIPT_DIR/build_workflow.py" "$APPLESCRIPT" "$CONTENTS_DIR"

# Register with Launch Services so Finder picks it up (warning about Spotlight is harmless)
"$LSREG" -f "$WORKFLOW_DIR" 2>/dev/null || true

# Refresh the Services menu immediately (no logout required)
/System/Library/CoreServices/pbs -update 2>/dev/null || true

echo ""
echo "Done! '${WORKFLOW_NAME}' is now available in Finder."
echo ""
echo "ONE-TIME SETUP:"
echo "  The first time you run it, macOS will ask:"
echo "  \"Automator wants to control Microsoft Word\" -> click Allow"
echo ""
echo "  If your files are stored on Google Drive or iCloud Drive:"
echo "  System Settings > Privacy & Security > Full Disk Access > enable Automator"
echo ""
echo "USAGE:"
echo "  1. Select exactly 2 .docx files in Finder"
echo "  2. Right-click > Quick Actions > ${WORKFLOW_NAME}"
echo "  3. A .redline.docx file will appear next to the revised document"
