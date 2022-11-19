#!/usr/bin/env python3

'''
███████ ██   ██ ████████ ██    ██ ██████  ██████   █████  ████████ ███████ 
██       ██ ██     ██    ██    ██ ██   ██ ██   ██ ██   ██    ██    ██      
█████     ███      ██    ██    ██ ██████  ██   ██ ███████    ██    █████   
██       ██ ██     ██    ██    ██ ██      ██   ██ ██   ██    ██    ██      
███████ ██   ██    ██     ██████  ██      ██████  ██   ██    ██    ███████ 
                                                                           
                                                                                                                                                                                    
Update/Change Excel File Extensions.
-
Author:
sorzkode
sorzkode@proton.me
https://github.com/sorzkode

MIT License
Copyright (c) 2022 sorzkode
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''
# Dependencies
import os                                       
from tkinter import filedialog                  
from tkinter import *
from tkinter.font import ITALIC, BOLD                                                           
import pathlib
import PySimpleGUI as sg

# PySimpleGUI version info
psversion = sg.version

# GUI window theme
sg.theme('Default1')

# Application menu
app_menu = [['&Help', ['&Usage', '&About']],]

# Directory selection
def select_folder():
    dirselect = filedialog.askdirectory()
    # Error handling if no selection
    if not dirselect:
        sg.popup_cancel('No folder selected', grab_anywhere=True, keep_on_top=True)
        return
    return dirselect

# All GUI elements
layout = [[sg.Menu(app_menu, tearoff=False, key='-MENU-')],
          [sg.Image(filename='assets\extuplogo.png', key='-LOGO-')],
          [sg.Button('Select Folder', font=('Lucida', 12, BOLD), pad=(5,15)),
          sg.In('Select Folder Path...', size=60, font=('Lucida', 11, ITALIC), text_color='Gray', readonly=True, enable_events=True, key='-FOLDER-')],
          [sg.Text('Curent Extension:', font=('Lucida', 14, BOLD)), 
          sg.OptionMenu(values=['.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xlt', '.xml'], default_value='.xls', key='-CURRENTEXT-', disabled=True),
          sg.Text('Updated Extension:', font=('Lucida', 14, BOLD)),
          sg.OptionMenu(values=['.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xlt', '.xml'], default_value='.xlsx', key='-UPDATEDEXT-', disabled=True)], 
          [sg.Button('Update Extensions', font=('Lucida', 12, BOLD), pad=(5,15), disabled=True), 
          sg.Button('Clear All', font=('Lucida', 12, BOLD), pad=(5,15), disabled=True),  
          sg.Button('Exit', font=('Lucida', 12, BOLD), pad=(5,15))]]
 
# Calls the main window / application
window = sg.Window('EXTUPDATE', layout, resizable=True, icon='assets\extup.ico', grab_anywhere=True, keep_on_top=True)

# Event loops when buttons are pressed / actions are taken in the app
while True:
    event, values = window.read()

# Window closed event
    if event == sg.WIN_CLOSED or event == 'Exit':
        break        

# File selection event
    if event == 'Select Folder':
        try:
            selected_path = select_folder()                                            
            window['-FOLDER-'].update(selected_path, text_color='Black')                     
            window['Clear All'].update(disabled=False)                    
            window['-CURRENTEXT-'].update(disabled=False)
            window['-UPDATEDEXT-'].update(disabled=False)
            window['Update Extensions'].update(disabled=False)
        except: sg.popup('Select a valid file path...try again', keep_on_top=True)                   

# Update Extensions
    if event == 'Update Extensions':
        bools_list = []
        files = os.listdir(selected_path)
        current_extension = values['-CURRENTEXT-']
        updated_extension = values['-UPDATEDEXT-']
        if current_extension == updated_extension:
            sg.popup('Extensions must not be the same...try again', keep_on_top=True)
            continue
        for file in files:
            filenames = pathlib.Path(str(file)).stem
            bools = file.endswith(current_extension)
            bools_list.append(bools)
            if not any(bools_list):
                sg.popup(f'No {current_extension} files in that directory...try again', keep_on_top=True)
                break
            elif file.endswith(current_extension):
                os.rename(
                    os.path.join(selected_path, filenames + current_extension), 
                    os.path.join(selected_path, filenames + updated_extension)
                    )
                window['Clear All'].update(disabled=True)
                window['Update Extensions'].update(disabled=True)
                window['-FOLDER-'].update('Success...try another folder', disabled=True)
                window['-CURRENTEXT-'].update(disabled=True)
                window['-UPDATEDEXT-'].update(disabled=True)

# Clear all button
    if event == 'Clear All':
        window['-FOLDER-'].update('Cleared...Select a new folder', disabled=True)
        window['Clear All'].update(disabled=True)
        window['Update Extensions'].update(disabled=True)
        window['-CURRENTEXT-'].update(disabled=True)
        window['-UPDATEDEXT-'].update(disabled=True)

# About menu selection
    if event == 'About':
        sg.popup( 
        'Update/Change Excel Extensions.',
        '',
        'Author: sorzkode',
        'Website: https://github.com/sorzkode',
        'License: MIT',
        '',
        'Files will be overwritten/changed when the script is executed',
        'Changing extension types may cause your file to lose some functionality',
        'Unless you are confident in the differences between file types, it is recommended to backup any files you plan on changing',
        '',
        f'PySimpleGUI Version: {psversion}',
        '', 
        grab_anywhere=True, keep_on_top=True, title='About')

# Usage menu selection
    if event == 'Usage':
        sg.popup( 
        'Follow these basic steps:',
        '',
        '1. Click the "Select Folder" button',
        '2. Select the extension you want to change from the "Current Extension" dropdown',
        '3. Select what you want to change the extension to from the "Updated Extension" dropdown',
        '4. Select the "Update Extensions" button to execute',
        '',
        grab_anywhere=True, keep_on_top=True, title='Usage')

window.close()