#!/bin/bash

# Define the virtual environment directory
VENV_DIR="venv"

# Check if the virtual environment exists
if [ ! -d "$VENV_DIR" ]; then
    echo "Virtual environment not found. Creating..."
    python3 -m venv "$VENV_DIR"
    
    echo "Activating virtual environment..."
    source "$VENV_DIR/bin/activate"
    
    echo "Installing dependencies..."
    pip install -r requirements.txt
else
    echo "Virtual environment found. Activating..."
    source "$VENV_DIR/bin/activate"
fi

# Run the Marpit Slide Editor
echo "Starting Marpit Slide Editor..."
python marpit_slide_editor.py
