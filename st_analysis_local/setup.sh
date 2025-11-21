#!/bin/bash

echo "======================================"
echo "Setting up Sales Analyzer"
echo "======================================"

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install requirements
echo "Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "======================================"
echo "Setup Complete!"
echo "======================================"
echo ""
echo "To use the analyzer:"
echo "1. Activate virtual environment: source venv/bin/activate"
echo "2. Place your .xls file in this directory"
echo "3. Run: python local_analyzer.py"
echo ""
