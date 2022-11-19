[![CodeQL](https://github.com/sorzkode/extupdate/actions/workflows/codeql.yml/badge.svg)](https://github.com/sorzkode/extupdate/actions/workflows/codeql.yml)
[[MIT Licence](https://en.wikipedia.org/wiki/MIT_License)]


![alt text](https://raw.githubusercontent.com/sorzkode/extupdate/master/assets/extupgit.png)

# Ext Update

Quickly update / change all Microsoft Excel file extensions in a directory of your choosing.

## Example

![alt text](https://raw.githubusercontent.com/sorzkode/extupdate/master/assets/example.png)

## Installation

Download from Github, changedir (cd) to the script directory and run the following:
```
pip install -e .
```
*This will install the ExtUpdate package locally 

Installation isn't required to run the script but you will need to ensure the requirements below are met.

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





