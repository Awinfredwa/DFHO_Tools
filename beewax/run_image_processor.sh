#!/bin/bash

echo ""
echo "===================================================="
echo "   BEE WAX CLOTH IMAGE PROCESSOR"
echo "===================================================="
echo ""
echo "Starting the image processing script..."
echo ""

# Try to run with python3 first, then python
if command -v python3 &> /dev/null; then
    python3 "Bee Wax Cloth Image Change Script.py"
elif command -v python &> /dev/null; then
    python "Bee Wax Cloth Image Change Script.py"
else
    echo ""
    echo "ERROR: Python is not installed or not found in PATH"
    echo ""
    echo "Please install Python:"
    echo "macOS (with Homebrew): brew install python"
    echo "Ubuntu/Debian: sudo apt install python3 python3-pip"
    echo "CentOS/RHEL: sudo yum install python3 python3-pip"
    echo ""
    echo "Required packages:"
    echo "pip3 install Pillow python-pptx"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo ""
echo "Script completed!"
read -p "Press Enter to exit..." 