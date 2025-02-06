import xlwings as xw
import tkinter as tk
from tkinter import messagebox
from datetime import date


def entry_button():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    # Get the source sheet (GALPON) where the values will be retrieve and the target sheet (ENTRADA) where the values will be pasted
    galpon_sheet = wb.sheets["GALPON"]
    entrada_sheet = wb.sheets["ENTRADA"]

    # Find the last row with data in column C
    last_row = galpon_sheet.range("B" + str(galpon_sheet.cells.last_cell.row)).end("up").row

    # Ensure there are entries to process
    if last_row >= 4:
        for row in range(last_row, 3, -1):  # Iterate from bottom to top
            provider = galpon_sheet.range(f"C{row}").value
            article = galpon_sheet.range(f"E{row}").value
            quantity = galpon_sheet.range(f"F{row}").value

            # Proceed only if all required fields have values
            entrada_sheet.api.Rows(3).Insert(Shift=-4121)  # Insert a new row at position 3

            # Insert values into the new row in ENTRADA sheet
            entrada_sheet.range("B3").value = date.today().strftime("%m/%d/%Y")
            entrada_sheet.range("C3").value = provider
            entrada_sheet.range("D3").value = article
            entrada_sheet.range("E3").value = quantity

            # Retrieve the cell where the "LIQUIDAR" button will be placed
            liquidar_cell = entrada_sheet.range("H3")
            x, y = liquidar_cell.left, liquidar_cell.top  # Get cell coordinates

            # Create the "LIQUIDAR" button (Shape)
            liquidar_shape = entrada_sheet.api.Shapes.AddShape(1, x, y, 215, 30)
            liquidar_shape.TextFrame.Characters().Text = "LIQUIDAR"
            liquidar_shape.TextFrame.HorizontalAlignment = -4108  # Center text horizontally
            liquidar_shape.TextFrame.VerticalAlignment = -4108  # Center text vertically
            liquidar_shape.Placement = 1
            liquidar_shape.OnAction = "liquidateMerch"  # Assign macro

            # Delete processed row in GALPON sheet
            galpon_sheet.range(f"B{row}:F{row}").delete(shift="up")

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
