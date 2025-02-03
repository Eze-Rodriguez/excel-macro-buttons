""" The following macro function must be created in Excel and assigned to each button on the MENU sheet

Sub menuButtons()
    Dim btn As Shape
    Set btn = ActiveSheet.Shapes(Application.Caller) ' Gets the button that was clicked

    ' Calls Python, passing the button name
    RunPython "from src.menu.menuButtons import handle_menu_buttons; handle_menu_buttons('" & btn.Name & "')"
End Sub
"""

import xlwings as xw

# Dictionary to map button names to sheet names
BUTTON_TO_SHEET = {
    "FACTURAR": "FACTURACION",
    "CTACTE CLIENTES": "CTACTE",
    "CTACTE PROVEEDORES": "CTACTEPROV",
}


def handle_menu_buttons(button_name):
    try:
        # Get the active workbook and sheet
        wb = xw.Book.caller()
        sheet = wb.sheets.active

        # Get the button text
        button = sheet.shapes[button_name]
        button_text = button.text.strip()  # Removes extra whitespace

        # Map the button text to the corresponding sheet name
        target_sheet_name = BUTTON_TO_SHEET.get(
            button_text, button_text
        )  # Uses the same name if it's not in the dictionary

        # Check if the sheet exists before activating it
        if target_sheet_name in [sh.name for sh in wb.sheets]:
            target_sheet = wb.sheets[target_sheet_name]
            target_sheet.activate()
            target_sheet.range("A1").select()  # Selects cell A1
            print(f"Se activó la hoja: {target_sheet_name}")
        else:
            print(f"Error: La hoja '{target_sheet_name}' no existe.")

    except Exception as e:
        print(f"Error al manejar el botón '{button_name}': {e}")