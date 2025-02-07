""" The following macro functions must be created in Excel and assigned to each button on the PAGOS sheet

Sub clientNewReceiptButton()
    RunPython "from src.modules.pagos.paymentButtons import new_receipt; new_receipt('client')"
End Sub

Sub providerNewReceiptButton()
    RunPython "from src.modules.pagos.paymentButtons import new_receipt; new_receipt('provider')"
End Sub


Sub clientSavePaymentButton()
    RunPython "from src.modules.pagos.paymentButtons import save_payment; save_payment('client')"
End Sub

Sub providerSavePaymentButton()
    RunPython "from src.modules.pagos.paymentButtons import save_payment; save_payment('provider')"
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
            date = pagos_sheet.range("B3").value
            client = pagos_sheet.range("D3").value
            receipt = pagos_sheet.range("E3").value
            method = pagos_sheet.range("G3").value
            amount = pagos_sheet.range("H3").value

            client_account_sheet.api.Rows(4).Insert(
                Shift=-4121
            )  # Insert a new row at position 4

            # Insert values into the new row in CTACTE sheet
            client_account_sheet.range("B4").value = date
            client_account_sheet.range("C4").value = client
            client_account_sheet.range("F4").value = receipt
            client_account_sheet.range("G4").value = method
            client_account_sheet.range("H4").value = amount

            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "Clientes: pago guardado correctamente :)",
                f"Recibo de pago N°{math.floor(receipt)} fue guardado satisfactoriamente.",
            )  # Info message box

        case "provider":
            # Retrieve values from PAGOS sheet
            date = pagos_sheet.range("B10").value
            client = pagos_sheet.range("D10").value
            receipt = pagos_sheet.range("E10").value
            method = pagos_sheet.range("G10").value
            amount = pagos_sheet.range("H10").value

            provider_account_sheet.api.Rows(4).Insert(
                Shift=-4121
            )  # Insert a new row at position 4

            # Insert values into the new row in CTACTE sheet
            provider_account_sheet.range("B4").value = date
            provider_account_sheet.range("C4").value = client
            provider_account_sheet.range("F4").value = receipt
            provider_account_sheet.range("G4").value = method
            provider_account_sheet.range("H4").value = amount

            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                "Clientes: pago guardado correctamente :)",
                f"Recibo de pago N°{math.floor(receipt)} fue guardado satisfactoriamente.",
            )  # Info message box

        case _:
            # Displays a message box
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showinfo(
                f"ERROR: Fallo al guardar pagos",
                f"No se han podido realizar el guardado de los nuevos pagos.",
            )  # Info message box
