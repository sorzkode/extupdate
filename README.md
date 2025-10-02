[![CodeQL](https://github.com/sorzkode/extupdate/actions/workflows/codeql.yml/badge.svg)](https://github.com/sorzkode/extupdate/actions/workflows/codeql.yml)
[[MIT Licence](https://en.wikipedia.org/wiki/MIT_License)]


![alt text](https://raw.githubusercontent.com/sorzkode/extupdate/master/assets/extupgit.png)

# EXTUPDATE - Excel Extension Manager

EXTUPDATE is a tool for bulk conversion of Microsoft Excel file extensions. Built with tkinter, it converts Excel files across directories with backup creation, recursive processing, and conversion history tracking.

## Features

- **Batch conversion** - Convert multiple Excel files at once
- **tkinter GUI** - Clean interface with dropdown menus
- **Backup creation** - Optional backup before conversion
- **Recursive processing** - Include subfolders in operations
- **File preview** - View files with metadata (size, modification date)
- **Conversion history** - Track operations with timestamps
- **Compatibility warnings** - Alerts for potential functionality loss
- **Progress tracking** - Real-time progress bar
- **Menu system** - Access to history, settings, and help

Standardize file extensions, convert legacy formats, or prepare files for specific applications.

Give it a try and streamline your file extension updates today!

## Example

![alt text](https://raw.githubusercontent.com/sorzkode/extupdate/master/assets/example.png)

## Installation

### Option 1: Download Executable (No Python Required)
Download the latest `EXTUPDATE.exe` from the [Releases](https://github.com/sorzkode/extupdate/releases) page and run directly.

### Option 2: Install with Python
**Install directly from GitHub:**
```bash
pip install git+https://github.com/sorzkode/extupdate.git
```

**Or install locally:**
1. Download/clone the repository
2. Navigate to the project directory
3. Run:
```bash
pip install .
```

## Requirements

- **For executable**: None - runs on any Windows machine
- **For Python version**: Python 3.8+ with tkinter (included with most Python installations)

## Usage

**Executable:**
```bash
# Simply run the downloaded file
EXTUPDATE.exe
```

**Python installation:**
```bash
extupdate
```

**Run directly from source:**
```bash
python extupdate.py
```
Once the script is running:
```
  1. Click "Select Folder" to choose a directory containing Excel files
  2. Select the current extension from the "Current Extension" dropdown
  3. Select the target extension from the "Convert to" dropdown
  4. Optional: Enable backup creation and/or recursive search
  5. Review files in the preview list
  6. Click "Update Extensions" to start conversion
```

Additional features:
```
  * View conversion history from the Tools menu
  * Check compatibility warnings before converting
  * Monitor progress with the real-time progress bar
  * Create backups automatically before conversion
```

Things to note:
```
  * Changing Excel file extensions may cause functionality loss depending on format compatibility
  * Enable backup creation when converting between incompatible formats
  * Use recursive search to include files in subfolders
```





