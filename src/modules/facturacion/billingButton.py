""" The following macro functions must be created in Excel and assigned to each button on the PAGOS sheet

Sub saveResetbillingButton()
    RunPython "from src.modules.facturacion.billingButton import save_reset_button; save_reset_button()"
End Sub
"""

import math
import xlwings as xw
import tkinter as tk
from datetime import date
from tkinter import messagebox


def check_rows():
    # Get the active workbook and sheet
    billing_sheet = xw.Book.caller().sheets["FACTURACION"]
    for row in range(16, 31):  # Search for products in all rows
        if billing_sheet.range(f"C{row}").value > 0:
            rows_to_bill = []
            rows_to_bill.append(row)
    return rows_to_bill


def new_billing():
    # Get the active workbook and sheet
    billing_sheet = xw.Book.caller().sheets["FACTURACION"]

    rows_to_bill = check_rows()
    for row in rows_to_bill:
        if isinstance(row, (int, float, complex)):
            try:
                # If data exists in the row, reset fields in entire row

                # Get current billing number, ensuring it's a number (default to 0 if empty)
                current_billing = billing_sheet.range("D4").value or 0
                new_billing = billing_sheet.range("D4").value = current_billing + 1
                billing_sheet.range("C2").value = date.today().strftime("%m/%d/%Y")
                billing_sheet.range(f"C{row}").value = 0

                # Displays a message box
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                messagebox.showinfo(
                    "Facturacion: nueva factura creada",
                    f"Se ha creado la factura número N° {math.floor(new_billing)}.",
                )  # Info message box
            except ValueError:
                print("Oops! Se esperaban un array de números. Ponerse en contacto con el administrador...")

def save_billing():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    billing_sheet = wb.sheets["FACTURACION"]
    x_sheet = wb.sheets["PROMEDIOVENTAS"]

    row = check_rows()
    if isinstance(row, (int, float, complex)):
        provider_account_sheet.api.Rows(4).Insert(
            Shift=-4121
        )  # Insert a new row at position 4

        # Retrieve values from FACTURACION sheet
        date = billing_sheet.range(f"B2").value
        provider_name = billing_sheet.range(f"C5").value
        billing_number = billing_sheet.range(f"D4").value
        billing_total = billing_sheet.range(f"F{row}").value

        # Paste values to CTACTEPROV sheet
        provider_account_sheet.range(f"B4").value = date
        provider_account_sheet.range(f"C4").value = provider_name
        provider_account_sheet.range(f"D4").value = billing_number
        provider_account_sheet.range(f"E4").value = billing_total
        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            "Facturacion: liquidación guardada",
            f"Se ha guardado la liquidación número N° {math.floor(billing_number)} satisfactoriamente.",
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
    save_billing(),  # First step is save the billing
    new_billing()  # Second step is delete all data and create a new billing
