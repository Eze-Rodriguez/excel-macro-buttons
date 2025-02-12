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
    rows_to_bill = []
    for row in range(17, 31):  # Search for products in all rows
        if billing_sheet.range(f"C{row}").value != None:
            rows_to_bill.append(row)
    return rows_to_bill


def new_billing():
    # Get the active workbook and sheet
    billing_sheet = xw.Book.caller().sheets["FACTURACION"]

    rows_to_bill = check_rows()
    if len(rows_to_bill) != 0:
        # Get current billing number, ensuring it's a number (default to 0 if empty)
        current_billing = billing_sheet.range("D4").value or 0
        new_billing = billing_sheet.range("D4").value = current_billing + 1
        billing_sheet.range("C2").value = date.today().strftime("%m/%d/%Y")
        
        # Delete all data from the rows to bill
        for row in rows_to_bill:
            billing_sheet.range(f"C{row}").value = "=BUSCARV($D$6;GALPON!$L$4:$AC$24;5;FALSO)" # Falta terminar
            billing_sheet.range(f"G{row}").value = 0

        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            "Facturacion: nueva factura creada",
            f"Se ha creado la factura N° {math.floor(new_billing)}.",
        )  # Info message box


def save_billing():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    billing_sheet = wb.sheets["FACTURACION"]
    account_sheet = wb.sheets["CTACTE"]

    row = check_rows()
    if len(row) != 0:
        if isinstance(row, (int, float, complex)):
            account_sheet.api.Rows(3).Insert(
                Shift=-4121
            )  # Insert a new row at position 3

            # Retrieve values from FACTURACION sheet
            date = billing_sheet.range(f"B2").value
            client_name = billing_sheet.range(f"C5").value
            billing_number = billing_sheet.range(f"D4").value
            billing_total = billing_sheet.range(f"F{row}").value

            # Paste values to PROMEDIO VENTAS sheet
            account_sheet.range(f"B3").value = date
            account_sheet.range(f"C3").value = client_name
            account_sheet.range(f"D3").value = billing_number
            account_sheet.range(f"E3").value = billing_total
            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "Facturacion: factura guardada",
                f"Se ha guardado la factura N° {math.floor(billing_number)} satisfactoriamente.",
            )  # Info message box
    else:
        # Displays a message box
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Atención: facturación vacía",
            f"No se han encontrado articulos para facturar.",
        )  # Info message box


# MAIN FUNCTION
def save_reset_button():
    save_billing(),  # First step is save the billing
    new_billing()  # Second step is delete all data and create a new billing
