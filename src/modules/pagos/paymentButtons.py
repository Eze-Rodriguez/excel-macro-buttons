"""The following macro functions must be created in Excel and assigned to each button on the PAGOS sheet

Sub clientSaveResetPaymentButton()
    RunPython "from src.modules.pagos.paymentButtons import save_reset_button; save_reset_button('client')"
End Sub

Sub providerSaveResetPaymentButton()
    RunPython "from src.modules.pagos.paymentButtons import save_reset_button; save_reset_button('provider')"
End Sub

"""

import math
from datetime import date
import tkinter as tk
from tkinter import messagebox
import xlwings as xw

# Create a global Tkinter root instance
root = tk.Tk()
root.withdraw()  # Hide the root window
root = tk.Tk()
root.withdraw()  # Hide the root window


def check_payments(paymentType):
    # Get the active workbook and sheet
    pagos_sheet = xw.Book.caller().sheets["PAGOS"]

    columns_to_reset = {
        "client": ["B3", "C3", "F3", "H3"],
        "provider": ["B10", "C10", "F10", "H10"],
    }
    if all(
        pagos_sheet.range(cell).value not in (None, 0)
        for cell in columns_to_reset[paymentType]
    ):
        return True
    else:
        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showerror(
            "ERROR: No se puede guardar el pago",
            f"Por favor, complete todos los campos antes de guardar el pago.",
        )  # Info message box
        root.destroy()
        return False


def new_payment(paymentType):
    # Get the active workbook and sheet
    pagos_sheet = xw.Book.caller().sheets["PAGOS"]
    if check_payments(paymentType):
        match paymentType:
            case "client":
                # Get current value, ensuring it's a number (default to 0 if empty)
                current_value = pagos_sheet.range("E3").value or 0

                # Increment the value
                new_value = pagos_sheet.range("E3").value = current_value + 1

                # Reset the payment fields
                pagos_sheet.range(f"B3").value = date.today().strftime("%m/%d/%Y")
                pagos_sheet.range(f"C3").value = 0
                pagos_sheet.range(f"F3").value = 0
                pagos_sheet.range(f"H3").value = 0

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

                # Reset the payment fields
                pagos_sheet.range(f"B10").value = date.today().strftime("%m/%d/%Y")
                pagos_sheet.range(f"C10").value = 0
                pagos_sheet.range(f"F10").value = 0
                pagos_sheet.range(f"H10").value = 0

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
                messagebox.showerror(
                    "ERROR: No es posible crear nuevo recibo",
                    f"No se puedo crear un nuevo recibo de pagos.",
                )  # Info message box


def save_payment(paymentType):
    wb = xw.Book.caller()
    pagos_sheet = wb.sheets["PAGOS"]
    client_account_sheet = wb.sheets["CTACTECLI"]
    provider_account_sheet = wb.sheets["CTACTEPROV"]

    if check_payments(paymentType):
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

                # Insert values into the new row in CTACTECLI sheet
                client_account_sheet.range("B4").value = date
                client_account_sheet.range("C4").value = client
                client_account_sheet.range("F4").value = receipt
                client_account_sheet.range("G4").value = method
                client_account_sheet.range("H4").value = amount

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

                # Insert values into the new row in CTACTEPROV sheet
                provider_account_sheet.range("B4").value = date
                provider_account_sheet.range("C4").value = client
                provider_account_sheet.range("F4").value = receipt
                provider_account_sheet.range("G4").value = method
                provider_account_sheet.range("H4").value = amount

                messagebox.showinfo(
                    "Proveedores: pago guardado correctamente :)",
                    f"Recibo de pago N°{math.floor(receipt)} fue guardado satisfactoriamente.",
                )  # Info message box
                root.destroy()
            case _:
                # Displays a message box
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                messagebox.showerror(
                    f"ERROR: Fallo al guardar pagos",
                    f"No se han podido realizar el guardado de los nuevos pagos.",
                )  # Info message box


def new_receipt(paymentType):
    wb = xw.Book.caller()
    pagos_sheet = wb.sheets["PAGOS"]
    receipt_sheet = wb.sheets["RECIBOS"]

    if check_payments(paymentType):
        match paymentType:
            case "client":
                # Retrieve values from PAGOS sheet
                payment_date = pagos_sheet.range("B3").value
                client = pagos_sheet.range("D3").value
                receipt = pagos_sheet.range("E3").value
                method = pagos_sheet.range("G3").value
                amount = pagos_sheet.range("H3").value

                # Insert values into the respective row in RECIBOS sheet
                receipt_sheet.range("A1").value = date.today().strftime("%m/%d/%Y")

                receipt_sheet.range("B3:C3").value = receipt
                receipt_sheet.range("B4:C4").value = ""

                receipt_sheet.range("C9:E9").value = payment_date
                receipt_sheet.range("C10:E10").value = client
                receipt_sheet.range("C12:E12").value = amount
                receipt_sheet.range("C13:E13").value = method
                receipt_sheet.range("C14:E14").value = ""

                # Displays a message box
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                messagebox.showinfo(
                    "Clientes: nuevo recibo listo para imprimir",
                    f"Recibo N°{math.floor(receipt)} para clientes listo para imprimir.",
                )  # Info message box

            case "provider":
                # Retrieve values from PAGOS sheet
                payment_date = pagos_sheet.range("B10").value
                client = pagos_sheet.range("D10").value
                receipt = pagos_sheet.range("E10").value
                method = pagos_sheet.range("G10").value
                amount = pagos_sheet.range("H10").value

                # Insert values into the respective row in RECIBOS sheet
                receipt_sheet.range("A1").value = date.today().strftime("%m/%d/%Y")

                receipt_sheet.range("B3:C3").value = receipt
                receipt_sheet.range("B4:C4").value = ""

                receipt_sheet.range("C9:E9").value = payment_date
                receipt_sheet.range("C10:E10").value = ""
                receipt_sheet.range("C12:E12").value = amount
                receipt_sheet.range("C13:E13").value = method
                receipt_sheet.range("C14:E14").value = client

                # Displays a message box
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                messagebox.showinfo(
                    "Proveedores: nuevo recibo listo para imprimir",
                    f"Recibo N°{math.floor(receipt)} para proveedores listo para imprimir.",
                )  # Info message box
            case _:
                # Displays a message box
                root.withdraw()  # Hide the root window
                messagebox.showerror(
                    f"ERROR: Fallo al guardar pagos",
                    f"No se han podido realizar el guardado de los nuevos pagos.",
                )  # Info message box


# MAIN FUNCTION
def save_reset_button(paymentType):
    """
    Executes a sequence of operations for handling payments:
    1. Saves the payment details to the corresponding sheet.
    2. Creates a new receipt for the payment.
    3. Resets the payment fields for a new transaction.

    Args:
        paymentType (str): Specifies the type of payment, either 'client' or 'provider'.

    Returns:
        None
    """
    save_payment(paymentType)  # Save the payment details to the corresponding sheet
    new_receipt(paymentType)  # Generate a new receipt for the payment
    new_payment(paymentType)  # Reset the payment fields for a new transaction
