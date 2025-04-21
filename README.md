# JSON to Excel Converter Script

This Python script converts JSON data into an Excel file format (`.xlsx`). It works on **macOS, Linux, and Windows**, offering both interactive and command-line options for a smooth and flexible experience.

## 🔧 Features

- ✅ **Cross-Platform Support**: Works seamlessly on macOS, Linux, and Windows.
- 🧰 **Optional Command-Line Arguments**: Supports `--input` and `--output` flags to skip interactive prompts.
- 🍎 **macOS Integration**: Uses native AppleScript dialogs for selecting files and folders.
- 📦 **Automatic Dependency Installation**: Installs required Python packages (`pandas`, `openpyxl`) if not already installed.
- 📁 **Smart Defaults**: If no output is specified, saves the Excel file in the same location as the JSON file.
- 📂 **Auto Open Output Folder**: Opens the output folder automatically after conversion on all platforms.

## 📦 Requirements

- Python 3.x
- `pandas`: For working with tabular data.
- `openpyxl`: For Excel file creation and editing.

> ✨ **No need to manually install dependencies** — the script will offer to install them for you.

## 📥 Installation

1. Clone or download this repository:
    ```bash
    git clone https://github.com/your-repo/json-to-excel.git
    cd json-to-excel
    ```

2. (Optional) Manually install requirements:
    ```bash
    python3 -m pip install pandas openpyxl
    ```

## 🚀 Usage

### Option 1: **Interactive Mode (No flags)**

Run the script and follow the prompts:

```bash
python3 json_to_excel.py
```

- **macOS**: Native dialogs appear for file and folder selection.
- **Other Platforms**: Prompts will appear in the terminal.

### Option 2: **Command-Line Mode**

```bash
python3 json_to_excel.py --input /path/to/input.json --output /path/to/output.xlsx
```

This bypasses prompts and runs headlessly.

### Example

```bash
python3 json_to_excel.py --input ./data/myfile.json
```

Output will be saved as `./data/myfile.xlsx`.

## 🔄 Script Flow

1. **Checks Dependencies**: Installs `pandas` and `openpyxl` if needed.
2. **Handles Input**: Reads from command-line or prompts user.
3. **Converts JSON**: Uses `pandas` to turn JSON into an Excel spreadsheet.
4. **Saves Output**: Exports `.xlsx` file and opens output folder.

## 🧩 Troubleshooting

- **Missing Dependencies**: The script will install them if you approve.
- **Invalid Paths**: Ensure the input file exists and is valid JSON.
- **macOS-only Dialogs**: File/folder dialogs are only available on macOS. Other platforms fall back to terminal prompts.
