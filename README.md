[![CodeQL](https://github.com/sorzkode/extupdate/actions/workflows/codeql.yml/badge.svg)](https://github.com/sorzkode/extupdate/actions/workflows/codeql.yml)
[[MIT Licence](https://en.wikipedia.org/wiki/MIT_License)]


![alt text](https://raw.githubusercontent.com/sorzkode/extupdate/master/assets/extupgit.png)

# Ext Update

Ext Update is a convenient tool that allows you to quickly update or change the file extensions of Microsoft Excel files in a directory of your choosing. With just a few simple steps, you can modify multiple file extensions at once, saving you time and effort.

## Features

- Easy installation and usage
- Supports bulk extension updates
- Intuitive user interface with dropdown menus
- Provides warnings for potential loss of functionality when changing certain file extensions

Whether you need to standardize file extensions or convert files to a different format, Ext Update simplifies the process and helps you manage your Excel files more efficiently.

Give it a try and streamline your file extension updates today!

## Example

![alt text](https://raw.githubusercontent.com/sorzkode/extupdate/master/assets/example.png)

## Installation

To install Ext Update, follow these steps:

1. Download the Ext Update package from GitHub.
2. Open a terminal and change the directory (cd) to the script directory.
3. Run the following command to install the package:
  ```
  pip install -e .
  ```
  This will install the Ext Update package locally.

Note: Installation is not required to run the script, but you will need to ensure that the following requirements are met.

## Requirements

The install above should take care of requirments.

Alternatively you can run: pip install -r requirements.txt

  [[Python 3](https://www.python.org/downloads/)]

  [[PySimpleGUI module](https://pypi.org/project/PySimpleGUI/)]

  [[tkinter](https://docs.python.org/3/library/tkinter.html)] :: Linux Users

## Usage

If Ext Update is installed you can use the following command syntax:
```
python -m extupdate
```
Otherwise you can run the script directly by changing directory (cd) in a terminal of your choice to the Ext Update directory and using the following syntax:
```
python extupdate.py
```
Once the script is initiated: 
```
  1. Click the "Select Folder" button.
  2. Select the extension you want to change from the "Current Extension" dropdown.
  3. Select what you want to change the extension to from the "Updated Extension" dropdown.
  4. Select the "Update Extensions" button to execute.
```

Things to note:
```
  * Changing Excel file extensions may cause a loss of functionality depending on which extension types you change.
  * If you aren't confident / don't know the differences between file types and their functionality - you may want to make copies of your files first.
```





