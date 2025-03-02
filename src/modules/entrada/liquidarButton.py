""" The following macro function must be created in Excel and assigned to each button on the ENTRADA sheet

Sub liquidateMerch()
    Dim btn As Shape
    Set btn = ActiveSheet.Shapes(Application.Caller) ' Gets the button that was clicked

    ' Calls Python, passing the button name
    RunPython "from src.modules.entrada.liquidarButton import liquidar_button; liquidar_button('" & btn.Name & "')"
End Sub
"""

import xlwings as xw
import tkinter as tk
from tkinter import messagebox
from datetime import date


def liquidar_button(button):
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    # Get the source sheet (ENTRADA) where the values will be retrieve and the target sheet (LIQUIDACIONES) where the values will be pasted
    entrada_sheet = wb.sheets["ENTRADA"]
    liquidaciones_sheet = wb.sheets["LIQUIDACIONES"]

    # Get row number of the button pressed
    active_button = sheet.shapes[button]
    button_row = active_button.api.TopLeftCell.Row

    provider = entrada_sheet.range(f"C{button_row}").value

    if provider == liquidaciones_sheet.range("C5:D5").value[0]:
        # Retrieve values from ENTRADA sheet
        article = entrada_sheet.range(f"D{button_row}").value
        quantity = entrada_sheet.range(f"E{button_row}").value
        price = entrada_sheet.range(f"F{button_row}").value
        entrada_sheet.range(f"G{button_row}").value = liquidaciones_sheet.range(
            "D4"
        ).value

        liquidaciones_sheet.api.Rows(13).Insert(Shift=-4121)  # Inserts new empty row
        # Paste values to LIQUIDACIONES sheet
        liquidaciones_sheet.range("B2").value = date.today().strftime("%m/%d/%Y")
        liquidaciones_sheet.range("C5:D5").value = provider
        liquidaciones_sheet.range("B13").value = date.today().strftime("%m/%d/%Y")
        liquidaciones_sheet.range("C13").value = quantity
        liquidaciones_sheet.range("D13").value = article
        liquidaciones_sheet.range("E13").value = price
        liquidaciones_sheet.range("F13").formula = f"=C13*E13"

        # Change the font color in ENTRADA to green
        entrada_sheet.range(f"B{button_row}:G{button_row}").color = (
            10,
            150,
            10,
        )  # Green

        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            f"Fila {button_row}: Añadida satisfactoriamente",
            f"La entrada en fila N° {button_row} fue añadida satisfactoriamente a la liquidación.",
        )  # Info message box
    else:
        # Displays a message box
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            f"ERROR: proveedor no coincide",
            f"El proveedor de la entrada en fila N° {button_row} no coincide con el proveedor de la liquidación.",
        )  # Info message box
