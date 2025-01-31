import xlwings as xw

def return_button():
    # Activate the MENU sheet
    menu_sheet = xw.Book.caller().sheets["MENU"]
    menu_sheet.activate()