#!/usr/bin/env python3

import os
import sys
import subprocess
import importlib
import json

# ─────────────────────────────────────────────────────────────────────────────
# macOS dialog helpers using AppleScript
# ─────────────────────────────────────────────────────────────────────────────

import subprocess

def macos_prompt(message, title="Python Script"):
    """
    Displays a Yes/No dialog using osascript on macOS.

    Args:
        message (str): The message to display in the dialog.
        title (str): The title of the dialog. Default is "Python Script".

    Returns:
        bool: True if "Yes" is selected, False otherwise.
    """
    script = f'display dialog "{message}" with title "{title}" buttons {{"No", "Yes"}} default button "Yes"'
    try:
        result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        return "yes" in result.stdout.lower()
    except Exception as e:
        print(f"Failed to show dialog: {e}")
        return False

def macos_choose_file():
    """
    Returns a file path using macOS file picker.

    Returns:
        str: File path selected using macOS file picker.
            None if an error occurs during the file selection process.
    """
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

import subprocess

def macos_choose_folder():
    """
    Returns a folder path using macOS folder picker.

    Returns:
        str: Folder path selected using macOS folder picker.
            None if an error occurs during the folder selection process.
    """
    script = '''
    set theFolder to choose folder with prompt "Select a location to save the Excel file"
    POSIX path of theFolder
    '''
    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        return result.stdout.strip()
    except Exception as e:
        print(f"Error choosing folder: {e}")
        return None

# ─────────────────────────────────────────────────────────────────────────────
# Package check and install if missing
# ─────────────────────────────────────────────────────────────────────────────

def check_and_install(package_name):
    """
    Check if a package is installed and install it if necessary.

    Args:
        package_name (str): The name of the package to check and install.

    Raises:
        Exception: If the package cannot be installed.

    Returns:
        None
    """
    try:
        importlib.import_module(package_name)
    except ImportError:
        if sys.platform == "darwin":
            # Prompt the user on macOS for package installation
            should_install = macos_prompt(
                f"'{package_name}' is not installed and is required for this script.\nPandas helps manipulate and analyze data in Python, offering powerful data structures and easy file handling.\nWould you like to install it?",
                title="JSON to Excel"
            )
        else:
            response = input(f"'{package_name}' is not installed and is required for this script. \nOpenPyXL handles the creation and manipulation of Excel files, enabling you to read/write data to.xlsx format. \nWould you like to install it?: ").strip().lower()
            should_install = response in ['yes', 'y']
        
        if should_install:
            print(f"Installing {package_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        else:
            print(f"Cannot continue without '{package_name}'. Exiting.")
            sys.exit(1)

# Ensure dependencies are installed
check_and_install("pandas")
check_and_install("openpyxl")

# Import pandas and openpyxl after confirming they’re installed
import pandas as pd
import openpyxl 

def check_and_install(package_name):
    """
    Check if a package is installed and install it if necessary.

    Args:
        package_name (str): The name of the package to check and install.

    Raises:
        Exception: If the package cannot be installed.

    Returns:
        None
    """
    try:
        importlib.import_module(package_name)
    except ImportError:
        if sys.platform == "darwin":
            # Prompt the user on macOS for package installation
            should_install = macos_prompt(
                f"'{package_name}' is not installed and is required for this script.\nPandas helps manipulate and analyze data in Python, offering powerful data structures and easy file handling.\nWould you like to install it?",
                title="JSON to Excel"
            )
        else:
            response = input(f"'{package_name}' is not installed and is required for this script. \nOpenPyXL handles the creation and manipulation of Excel files, enabling you to read/write data to.xlsx format. \nWould you like to install it?: ").strip().lower()
            should_install = response in ['yes', 'y']
        
        if should_install:
            print(f"Installing {package_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        else:
            print(f"Cannot continue without '{package_name}'. Exiting.")
            sys.exit(1)
# ─────────────────────────────────────────────────────────────────────────────
# Get input/output file paths
# ─────────────────────────────────────────────────────────────────────────────

input_file = ""
output_file = ""

if not input_file and sys.platform == "darwin":
    input_file = macos_choose_file()

if not input_file:
    input_file = input("Please enter your input JSON file path: ")

if not output_file and sys.platform == "darwin":
    output_folder = macos_choose_folder()
    if output_folder:
        base_name = os.path.splitext(os.path.basename(input_file))[0]  # Get the base name without the extension
        output_file = os.path.join(output_folder, f"{base_name}.xlsx")  # Append .xlsx

if not output_file:
    base_name = os.path.splitext(os.path.basename(input_file))[0]  # Get the base name without the extension
    output_file = f"{base_name}.xlsx"  # Append .xlsx

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
    print("📂 Opening output folder...")
    os.system(f"open '{folder_path}'")
elif sys.platform == 'linux':
    os.system(f"xdg-open '{folder_path}'")
elif sys.platform == 'win32':
    os.startfile(folder_path)
