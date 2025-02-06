""" The following macro function must be created in Excel and assigned to each button on the GLOBL sheet. Yes, GLOBL.

Sub returnButton()
    RunPython "from src.globalUtils import return_button; return_button()"
End Sub
"""

import xlwings as xw

def return_button():
    # Activate the MENU sheet
    menu_sheet = xw.Book.caller().sheets["MENU"]
    menu_sheet.activate()