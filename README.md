# JSON to Excel Converter Script

This Python script is designed to convert JSON data into an Excel file format (.xlsx). It works on macOS, Linux, and Windows, offering a simple and interactive way to convert JSON data into a well-organized Excel spreadsheet.

## Features

- **macOS Integration**: Interactive file and folder picker dialogs using AppleScript.
- **Dependency Installation**: Automatically installs required Python packages (`pandas` and `openpyxl`) if not already installed.
- **Cross-Platform**: Works on macOS, Linux, and Windows with minimal setup.
- **Easy Conversion**: Converts any JSON file into a `.xlsx` Excel file.

## Requirements

- Python 3.x
- `pandas`: A Python package used for data manipulation and analysis.
- `openpyxl`: A Python package for reading and writing Excel (xlsx) files.

- *The script will ask if you would like to install it before running*

## Installation

Before running the script, make sure you have Python 3.x installed on your machine.

- *The script will ask if you would like to install it before running*

1. Clone or download this repository.
2. Install the required Python packages (if they are not already installed) by running:

```bash
python3 -m pip install pandas openpyxl
```

## Usage

### Step-by-step guide:

1. **Run the script**:
   On macOS, you can run it directly from the terminal:

```bash
python3 json_to_excel.py
```

2. **Choose your input file**:
   - For macOS: A dialog will prompt you to select your input JSON file.
   - For other platforms: You'll be asked to manually provide the path to your JSON file.

3. **Choose the output folder**:
   - For macOS: A dialog will prompt you to select the folder where you want to save the resulting Excel file.
   - For other platforms: The script will automatically save the Excel file in the current working directory.

4. **Conversion**:
   - The script will read the input JSON file, convert it into a pandas DataFrame, and write it to an Excel file using `openpyxl`.

5. **Output**:
   - The script will print a confirmation message indicating the location of the saved `.xlsx` file.
   - The output folder will be opened automatically (macOS only).

### Example:

1. Input: `data.json`
2. Output: `data.xlsx`

## Script Flow

1. **macOS Dialog Helpers**: The script uses AppleScript to prompt the user for file/folder selections (only available on macOS).
2. **Package Installation**: If `pandas` or `openpyxl` are missing, the script will prompt the user to install them.
3. **JSON to Excel Conversion**: The script loads the input JSON, converts it into a pandas DataFrame, and then exports it to Excel.
4. **Opening Output Folder**: After conversion, the script opens the output folder on macOS, Linux, or Windows.

## Troubleshooting

- **Missing Dependencies**: If you encounter an error related to missing dependencies (e.g., `pandas` or `openpyxl`), the script will prompt you to install them automatically.
- **Invalid File Paths**: Ensure that the input JSON file and the chosen output folder are valid paths.

