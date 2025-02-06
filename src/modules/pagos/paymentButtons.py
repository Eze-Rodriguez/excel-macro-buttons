""" The following macro functions must be created in Excel and assigned to each button on the PAGOS sheet

Sub clientNewReceiptButton()
    RunPython "from src.modules.pagos.paymentButton import new_receipt; new_receipt('client')"
End Sub

Sub providerNewReceiptButton()
    RunPython "from src.modules.pagos.paymentButton import new_receipt; new_receipt('provider')"
End Sub
"""

import math
import xlwings as xw
import tkinter as tk
from tkinter import messagebox


def new_receipt(paymentType):
    # Get the active workbook and sheet
    pagos_sheet = xw.Book.caller().sheets["PAGOS"]

    match paymentType:
        case "client":
            # Get current value, ensuring it's a number (default to 0 if empty)
            current_value = pagos_sheet.range("E3").value or 0

            # Increment the value
            new_value = pagos_sheet.range("E3").value = current_value + 1

            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "Clientes: nuevo recibo disponible",
                f"Recibo N°{math.floor(new_value)} disponible para clientes.",
            )  # Info message box
        case "provider":
            # Get current value, ensuring it's a number (default to 0 if empty)
            current_value = pagos_sheet.range("E10").value or 0

            # Increment the value
            new_value = pagos_sheet.range("E10").value = current_value + 1

            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "Proveedores: nuevo recibo disponible",
                f"Recibo N°{math.floor(new_value)} disponible para proveedores.",
            )  # Info message box
        case _:
            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "ERROR: No es posible crear nuevo recibo",
                f"No se puedo crear un nuevo recibo de pagos.",
            )  # Info message box


def save_payment(paymentType):
    wb = xw.Book.caller()
    pagos_sheet = wb.sheets["PAGOS"]
    client_account_sheet = wb.sheets["CTACTE"]
    provider_account_sheet = wb.sheets["CTACTEPROV"]

    match paymentType:
        case "client":
            # Retrieve values from PAGOS sheet
            receipt = pagos_sheet.range("E3").value
            .
            .
            .

            client_account_sheet.api.Rows(4).Insert(Shift=-4121)  # Insert a new row at position 3

            # Insert values into the new row in CTACTE sheet
            client_account_sheet.range("E3").value = receipt
            .
            .
            .
            
            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "Clientes: pago guardado correctamente :)",
                f"Recibo de pago N°{math.floor(receipt)} fue guardado satisfactoriamente.",
            )  # Info message box
        case "provider":
            return "provider"
        case _:
            return "default"