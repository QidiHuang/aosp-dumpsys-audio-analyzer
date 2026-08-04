#!/bin/bash
# Check if python3 is installed
if ! command -v python3 &> /dev/null
then
    echo "python3 could not be found. Please install Python 3.6 or newer."
    echo "On Ubuntu/Debian: sudo apt-get install python3 python3-tk"
    echo "On macOS (using brew): brew install python-tk"
else
    # Run the analyzer
    python3 dumpsys_analyzer.py
fi
