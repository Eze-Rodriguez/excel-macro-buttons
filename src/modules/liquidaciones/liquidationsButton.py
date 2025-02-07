""" The following macro functions must be created in Excel and assigned to each button on the PAGOS sheet

Sub saveResetLiquidationButton()
    RunPython "from src.modules.liquidaciones.liquidationsButton import save_reset_button; save_reset_button()"
End Sub
"""

import math
import xlwings as xw
import tkinter as tk
from datetime import date
from tkinter import messagebox


def check_rows():
    # Get the active workbook and sheet
    liquidaciones_sheet = xw.Book.caller().sheets["LIQUIDACIONES"]
    # Detect the footer row by searching for the "TOTAL LIQUIDACIÓN" text
    footer_text = "TOTAL LIQUIDACION"
    for row in range(50, 12, -1):  # Search all rows
        if (
            liquidaciones_sheet.range(f"D{row}").value == footer_text
        ):  # Adjust column if needed
            footer_row = row
            break

    # Define the start row for the table (below the header)
    start_row = 13  # Row where the table data starts (below the header)
    columns_to_check = [
        "B",
        "C",
        "D",
        "E",
        "F",
    ]  # Table columns to validate (Fecha, Cantidad, etc.)

    # Iterate through rows from bottom to top (to avoid skipping rows when deleting)
    for row in range(footer_row - 1, start_row - 1, -1):
        # Check if the row contains data in the specified columns
        has_data = any(
            liquidaciones_sheet.range(f"{col}{row}").value for col in columns_to_check
        )

        if has_data:
            return row


def new_liquidation():
    # Get the active workbook and sheet
    liquidaciones_sheet = xw.Book.caller().sheets["LIQUIDACIONES"]

    row = check_rows()
    if isinstance(row, (int, float, complex)):
        # If data exists in the row, delete the entire row
        liquidaciones_sheet.range(f"B13:F{row}").delete(shift="up")

        # Get current liquidation number, ensuring it's a number (default to 0 if empty)
        current_liquidation = liquidaciones_sheet.range("D4").value or 0
        new_liquidation = liquidaciones_sheet.range("D4").value = (
            current_liquidation + 1
        )
        liquidaciones_sheet.range("B2").value = date.today().strftime("%m/%d/%Y")
        liquidaciones_sheet.range("C5:D5").value = None

        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            "Liquidaciones: nueva liquidación creada",
            f"Se ha creado la liquidación número N° {math.floor(new_liquidation)}.",
        )  # Info message box


def save_liquidation():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    liquidaciones_sheet = wb.sheets["LIQUIDACIONES"]
    provider_account_sheet = wb.sheets["CTACTEPROV"]

    row = check_rows()
    if isinstance(row, (int, float, complex)):
        provider_account_sheet.api.Rows(4).Insert(
            Shift=-4121
        )  # Insert a new row at position 4

        # Retrieve values from LIQUIDACIONES sheet
        date = liquidaciones_sheet.range(f"B2").value
        provider_name = liquidaciones_sheet.range(f"C5").value
        liquidation_number = liquidaciones_sheet.range(f"D4").value
        liquidation_total = liquidaciones_sheet.range(f"F{row}").value

        # Paste values to CTACTEPROV sheet
        provider_account_sheet.range(f"B4").value = date
        provider_account_sheet.range(f"C4").value = provider_name
        provider_account_sheet.range(f"D4").value = liquidation_number
        provider_account_sheet.range(f"E4").value = liquidation_total
        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            "Liquidaciones: liquidación guardada",
            f"Se ha guardado la liquidación número N° {math.floor(liquidation_number)} satisfactoriamente.",
        )  # Info message box

    else:
        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            "Atención: liquidación vacía",
            f"No se han encontrado articulos para liquidar.",
        )  # Info message box


# MAIN FUNCTION
def save_reset_button():
    save_liquidation(),  # First step is save the liquidation
    new_liquidation()  # Second step is delete all data and create a new liquidation
