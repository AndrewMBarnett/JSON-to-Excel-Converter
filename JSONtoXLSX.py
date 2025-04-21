#!/usr/bin/env python3

import os
import sys
import subprocess
import importlib
import json
import argparse

# ─────────────────────────────────────────────────────────────────────────────
# macOS dialog helpers using AppleScript
# ─────────────────────────────────────────────────────────────────────────────

def macos_prompt(message, title="Python Script"):
    script = f'display dialog "{message}" with title "{title}" buttons {{"No", "Yes"}} default button "Yes"'
    try:
        result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        return "yes" in result.stdout.lower()
    except Exception as e:
        print(f"Failed to show dialog: {e}")
        return False

def macos_choose_file():
    script = '''
    set theFile to choose file with prompt "Select your input JSON file"
    POSIX path of theFile
    '''
    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        return result.stdout.strip()
    except Exception as e:
        print(f"Error choosing file: {e}")
        return None

def macos_choose_folder():
    script = '''
    set theFolder to choose folder with prompt "Select a folder to save the Excel file"
    POSIX path of theFolder
    '''
    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        return result.stdout.strip()
    except Exception as e:
        print(f"Error choosing folder: {e}")
        return None

# ─────────────────────────────────────────────────────────────────────────────
# Package check and install
# ─────────────────────────────────────────────────────────────────────────────

def check_and_install(package_name, description=""):
    try:
        importlib.import_module(package_name)
    except ImportError:
        should_install = False
        if sys.platform == "darwin":
            should_install = macos_prompt(
                f"'{package_name}' is not installed and is required for this script.\n{description}\nWould you like to install it?",
                title="JSON to Excel"
            )
        else:
            response = input(f"'{package_name}' is not installed.\n{description}\nInstall it now? (y/n): ").strip().lower()
            should_install = response in ['y', 'yes']
        if should_install:
            print(f"Installing {package_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        else:
            print(f"Cannot continue without '{package_name}'. Exiting.")
            sys.exit(1)

check_and_install("pandas", "Pandas helps manipulate and analyze data in Python, offering powerful data structures and easy file handling.")
check_and_install("openpyxl", "OpenPyXL handles the creation and manipulation of Excel files, enabling .xlsx output.")

import pandas as pd
import openpyxl

# ─────────────────────────────────────────────────────────────────────────────
# Parse CLI arguments
# ─────────────────────────────────────────────────────────────────────────────

parser = argparse.ArgumentParser(description="Convert JSON to Excel")
parser.add_argument('--input', '-i', help='Path to input JSON file')
parser.add_argument('--output', '-o', help='Path to output Excel file')
args = parser.parse_args()

input_file = args.input
output_file = args.output

# ─────────────────────────────────────────────────────────────────────────────
# Interactive fallbacks
# ─────────────────────────────────────────────────────────────────────────────

if not input_file:
    if sys.platform == 'darwin':
        input_file = macos_choose_file()
    else:
        input_file = input("Enter path to input JSON file: ").strip()

if not input_file or not os.path.isfile(input_file):
    print("❌ Invalid or missing input file. Exiting.")
    sys.exit(1)

if not output_file:
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    if sys.platform == 'darwin':
        output_folder = macos_choose_folder()
        if not output_folder:
            output_folder = os.getcwd()
    else:
        output_folder = os.getcwd()
    output_file = os.path.join(output_folder, f"{base_name}.xlsx")

# ─────────────────────────────────────────────────────────────────────────────
# Convert JSON to Excel
# ─────────────────────────────────────────────────────────────────────────────

try:
    with open(input_file, 'r') as f:
        data = json.load(f)
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"✅ Excel file created: {output_file}")
except Exception as e:
    print(f"❌ Failed to convert JSON to Excel: {e}")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# Open output folder
# ─────────────────────────────────────────────────────────────────────────────

folder_path = os.path.dirname(os.path.abspath(output_file))

if sys.platform == 'darwin':
    os.system(f"open '{folder_path}'")
elif sys.platform == 'linux':
    subprocess.run(['xdg-open', folder_path])
elif sys.platform == 'win32':
    os.startfile(folder_path)
