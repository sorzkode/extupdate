#!/usr/bin/env python3
#%%
"""
███████ ██   ██ ████████ ██    ██ ██████  ██████   █████  ████████ ███████ 
██       ██ ██     ██    ██    ██ ██   ██ ██   ██ ██   ██    ██    ██      
█████     ███      ██    ██    ██ ██████  ██   ██ ███████    ██    █████   
██       ██ ██     ██    ██    ██ ██      ██   ██ ██   ██    ██    ██      
███████ ██   ██    ██     ██████  ██      ██████  ██   ██    ██    ███████ 
                                                                           
                                                                                                                                                                                    
Update/Change Excel File Extensions.
-
Author:
Mister Riley
sorzkode@proton.me
https://github.com/sorzkode

Co-Author:
Mister Cohen
https://github.com/bcherb2

MIT License
Copyright (c) 2024 Mister Riley
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""
# Dependencies
import os
from tkinter import filedialog
from tkinter.font import ITALIC, BOLD
import PySimpleGUI as sg

# PySimpleGUI version info
psversion = sg.version

# GUI window theme
sg.theme("Default1")

# Application menu
app_menu = [
    ["&Help", ["&Usage", "&About"]],
]

# Font variables
font_title = ("Lucida", 14, BOLD)
font_button = ("Lucida", 12, BOLD)
font_text = ("Lucida", 11, ITALIC)

# Directory selection
def select_folder():
    dirselect = filedialog.askdirectory()
    # Error handling if no selection
    if not dirselect:
        sg.popup_cancel("No folder selected", grab_anywhere=True, keep_on_top=True)
        return None
    return dirselect

# All GUI elements
layout = [
    [sg.Menu(app_menu, tearoff=False, key="-MENU-")],
    [sg.Image(filename="assets\extuplogo.png", key="-LOGO-")],
    [
        sg.Button("Select Folder", font=font_button, pad=(5, 15)),
        sg.In(
            "Select Folder Path...",
            size=60,
            font=font_text,
            text_color="Gray",
            readonly=True,
            enable_events=True,
            key="-FOLDER-",
        ),
    ],
    [
        sg.Text("Curent Extension:", font=font_title),
        sg.OptionMenu(
            values=[
                ".xls",
                ".xlsx",
                ".xlsm",
                ".xlsb",
                ".xltx",
                ".xltm",
                ".xlt",
                ".xml",
            ],
            default_value=".xls",
            key="-CURRENTEXT-",
            disabled=True,
        ),
        sg.Text("Updated Extension:", font=font_title),
        sg.OptionMenu(
            values=[
                ".xls",
                ".xlsx",
                ".xlsm",
                ".xlsb",
                ".xltx",
                ".xltm",
                ".xlt",
                ".xml",
            ],
            default_value=".xlsx",
            key="-UPDATEDEXT-",
            disabled=True,
        ),
    ],
    [
        sg.Button(
            "Update Extensions", font=font_button, pad=(5, 15), disabled=True
        ),
        sg.Button("Clear All", font=font_button, pad=(5, 15), disabled=True),
        sg.Button("Exit", font=font_button, pad=(5, 15)),
    ],
]

class ext_window:
    def __init__(self, window):
        self.window = sg.Window(
            "EXTUPDATE",
            layout,
            resizable=True,
            icon="assets\extup.ico",
            grab_anywhere=True,
            keep_on_top=True,
        )
        self.selected_path = None
        self.start()

    def clear_window_defaults(self) -> None:
        self.window["Clear All"].update(disabled=False)
        self.window["-CURRENTEXT-"].update(disabled=False)
        self.window["-UPDATEDEXT-"].update(disabled=False)
        self.window["Update Extensions"].update(disabled=False)

    def select_folder_gui(self) -> None:
        try:
            self.selected_path = select_folder()
            if self.selected_path:
                self.clear_window_defaults()
                self.window["-FOLDER-"].update(self.selected_path, text_color="Black")
        except Exception as e:
            sg.popup(f"Error: {e}", keep_on_top=True)

    def find_files_in_path(self) -> list[str]:
        if not self.selected_path:
            return []
        # Find all files in selected path
        files = [
            f
            for f in os.listdir(self.selected_path)
            if os.path.isfile(os.path.join(self.selected_path, f))
        ]
        # Filter out files with the current extension
        current_ext = self.values["-CURRENTEXT-"]
        files = [f for f in files if f.endswith(current_ext)]
        return files

    def update_entensions_gui(self) -> None:
        if not self.selected_path:
            sg.popup("No folder selected", keep_on_top=True)
            return

        current_extension = self.values["-CURRENTEXT-"]
        updated_extension = self.values["-UPDATEDEXT-"]
        files = self.find_files_in_path()
        try:
            if current_extension == updated_extension:
                sg.popup("Extensions must not be the same...try again", keep_on_top=True)
                return
            if not any(files):
                sg.popup(
                    f"No {current_extension} files in that directory...try again",
                    keep_on_top=True,
                )
                return
            for file in files:
                file_path = os.path.join(self.selected_path, file)
                new_file_path = file_path.replace(current_extension, updated_extension)
                os.rename(file_path, new_file_path)
            sg.popup(f"Updated {len(files)} files", keep_on_top=True)
        except Exception as e:
            sg.popup(f"Error: {e}", keep_on_top=True)

    def about_gui(self):
        sg.popup(
            "Update/Change Excel Extensions.",
            "",
            "Author: sorzkode",
            "Website: https://github.com/sorzkode",
            "License: MIT",
            "",
            "Files will be overwritten/changed when the script is executed",
            "Changing extension types may cause your file to lose some functionality",
            "Unless you are confident in the differences between file types, it is recommended to backup any files you plan on changing",
            "",
            f"PySimpleGUI Version: {psversion}",
            "",
            grab_anywhere=True,
            keep_on_top=True,
            title="About",
        )

    def usage_gui(self):
        sg.popup(
            "Follow these basic steps:",
            "",
            '1. Click the "Select Folder" button',
            '2. Select the extension you want to change from the "Current Extension" dropdown',
            '3. Select what you want to change the extension to from the "Updated Extension" dropdown',
            '4. Select the "Update Extensions" button to execute',
            "",
            grab_anywhere=True,
            keep_on_top=True,
            title="Usage",
        )

    def start(self):
        # Event loops when buttons are pressed / actions are taken in the app
        while True:
            self.event, self.values = self.window.read()

            # Window closed event
            if self.event == sg.WIN_CLOSED or self.event == "Exit":
                break

            match self.event:
                case "Select Folder":
                    self.select_folder_gui()
                case "Update Extensions":
                    self.update_entensions_gui()
                case "Clear All":
                    self.clear_window_defaults()
                    self.window["-FOLDER-"].update(
                        "Cleared...Select a new folder", disabled=True
                    )
                    self.window["Update Extensions"].update(disabled=True)
                case "About":
                    self.about_gui()
                case "Usage":
                    self.usage_gui()

        self.window.close()

if __name__ == "__main__":
    ext_window(sg)