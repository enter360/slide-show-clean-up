#!/bin/bash

# Go to the project folder
cd "$(dirname "$0")"

# Install dependencies if needed
/usr/local/bin/python3 -m pip install -q python-pptx pypdf pdfplumber Pillow

# Run the script
/usr/local/bin/python3 remove_blank_slides.py

echo ""
echo "Press any key to close..."
read -n 1
