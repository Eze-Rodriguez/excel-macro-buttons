import xlwings as xw
import tkinter as tk
from tkinter import messagebox


def liquidar_button(button):
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    # Get the source sheet (ENTRADA) where the values will be retrieve and the target sheet (FACTURACION) where the values will be pasted
    entrada_sheet = wb.sheets["ENTRADA"]
    facturacion_sheet = wb.sheets["FACTURACION"]

    # Get row number of the button pressed
    active_button = sheet.shapes[button]
    button_row = active_button.top_left_cell.row

    # Retrieve values from ENTRADA sheet
    date = entrada_sheet.range(f"B{button_row}").value
    provider = entrada_sheet.sheet.range(f"C{button_row}").value
    article = entrada_sheet.range(f"D{button_row}").value
    quantity = entrada_sheet.range(f"E{button_row}").value
    price = entrada_sheet.range(f"F{button_row}").value

    # Paste values to FACTURACION sheet
    if not facturacion_sheet.range("B2").value:  # Pastes if date field is empty
        facturacion_sheet.range("B2").value = date

    facturacion_sheet.range("C5:D5").value = provider
    if textUp:  # Makes a row break if the text above is not empty
        # Salto de linea
        facturacion_sheet.range("").value = article
        facturacion_sheet.range("").value = quantity
        facturacion_sheet.range("").value = price

    # Displays a message box
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo(f"Fila {button_row}: Añadida satisfactoriamente", f"La entrada en fila N° {button_row} fue añadida satisfactoriamente a la liquidación")  # Info message box
