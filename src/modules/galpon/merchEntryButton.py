import time
import xlwings as xw
import tkinter as tk
from tkinter import messagebox
from datetime import date


def entry_button():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    # Get the source sheet (GALPON) where the values will be retrieve and the target sheet (ENTRADA) where the values will be pasted
    galpon_sheet = wb.sheets["GALPON"]
    entrada_sheet = wb.sheets["ENTRADA"]

    last_row = (
        galpon_sheet.range("B" + str(galpon_sheet.cells.last_cell.row)).end("up").row
    )
    entry_field = galpon_sheet.range("C4").value

    if entry_field:
        for row in range(4, last_row):
            # Retrieve values from GALPON sheet
            provider = galpon_sheet.range(f"C{row}").value
            article = galpon_sheet.range(f"E{row}").value
            quantity = galpon_sheet.range(f"F{row}").value

            # Inserts new empty row
            entrada_sheet.api.Rows(3).Insert(Shift=-4121)

            # Paste values to ENTRADA sheet
            entrada_sheet.range("B3").value = date.today().strftime(
                "%m/%d/%Y"
            )  # Set the current local date
            entrada_sheet.range("C3").value = provider
            entrada_sheet.range("D3").value = article
            entrada_sheet.range("E3").value = quantity

            liquidar_cell = entrada_sheet.range(
                "H3"
            )  # Retrieve cell where will be the button "Liquidar"
            X, Y = liquidar_cell.left, liquidar_cell.top  # Cell coordinates
            # Create the button "Liquidar" (Shape)
            liquidar_shape = entrada_sheet.api.Shapes.AddShape(1, X, Y, 215, 30)
            liquidar_shape.TextFrame.Characters().Text = "LIQUIDAR"
            liquidar_shape.TextFrame.HorizontalAlignment = -4108
            liquidar_shape.TextFrame.VerticalAlignment = -4108
            liquidar_shape.Placement = 2
            liquidar_shape.OnAction = "liquidateMerch"

            # Delete added entries
            galpon_sheet.range(f"B{row}:F{row}").delete()
            time.sleep(1)  # Wait 1 sec

        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            f"Entradas: añadidas satisfactoriamente",
            f"Se han añadido satisfactoriamente todas las entradas del galpon.",
        )  # Info message box
    else:
        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            f"Atención: No existen entradas",
            f"No se han encontrado nuevas entradas para agregar.",
        )  # Info message box
