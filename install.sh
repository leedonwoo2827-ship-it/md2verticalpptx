#!/usr/bin/env bash
set -e

echo "============================================"
echo "  md2pptx Installer (Linux/macOS)"
echo "============================================"
echo ""
echo "WARNING: md2pptx uses Windows COM automation (PowerPoint)."
echo "  Full functionality requires Windows + Microsoft PowerPoint."
echo "  This installer sets up the Python environment only."
echo ""

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 not found. Install Python 3.10+ first."
    exit 1
fi

PYVER=$(python3 --version 2>&1)
echo "[OK] $PYVER"

# Create venv
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
fi

# Activate and install
echo "Installing dependencies..."
source .venv/bin/activate
pip install --upgrade pip >/dev/null 2>&1

# Install without comtypes (Windows-only)
pip install "rich>=13.0" "python-pptx>=0.6.23"
pip install -e .

echo ""
echo "============================================"
echo "  Installation complete!"
echo "============================================"
echo ""
echo "Usage:"
echo "  ./run.sh <body.md> -t <templates_dir> [-o output.pptx]"
echo ""
echo "Note: COM-based slide building requires Windows + PowerPoint."
