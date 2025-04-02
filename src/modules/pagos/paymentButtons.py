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

# Crear una instancia global de Tkinter root
root = tk.Tk()
root.withdraw()  # Ocultar la ventana principal


def check_payments(paymentType):
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
        messagebox.showerror(
            "ERROR: No se puede guardar el pago",
            "Por favor, complete todos los campos antes de guardar el pago.",
        )
        return False


def save_reset_button(paymentType):
    """
    Ejecuta una secuencia de operaciones para manejar pagos:
    1. Guarda los detalles del pago en la hoja correspondiente.
    2. Crea un nuevo recibo para el pago.
    3. Resetea los campos de pago para una nueva transacción.

    Args:
        paymentType (str): Especifica el tipo de pago, ya sea 'client' o 'provider'.

    Returns:
        None
    """
    messages = []  # Lista para acumular mensajes

    if check_payments(paymentType):
        # Guardar el pago
        messages.append(save_payment(paymentType))

        # Crear un nuevo recibo
        messages.append(new_receipt(paymentType))

        # Resetear los campos de pago
        messages.append(new_payment(paymentType))

        # Mostrar una sola ventana emergente con toda la información
        messagebox.showinfo(
            "Operación completada",
            "\n".join(messages),
        )


def save_payment(paymentType):
    wb = xw.Book.caller()
    pagos_sheet = wb.sheets["PAGOS"]
    client_account_sheet = wb.sheets["CTACTECLI"]
    provider_account_sheet = wb.sheets["CTACTEPROV"]

    if paymentType == "client":
        # Recuperar valores de la hoja PAGOS
        date = pagos_sheet.range("B3").value
        client = pagos_sheet.range("D3").value
        receipt = pagos_sheet.range("E3").value
        method = pagos_sheet.range("G3").value
        amount = pagos_sheet.range("H3").value

        client_account_sheet.api.Rows(4).Insert(Shift=-4121)

        # Insertar valores en la nueva fila de la hoja CTACTECLI
        client_account_sheet.range("B4").value = date
        client_account_sheet.range("C4").value = client
        client_account_sheet.range("F4").value = receipt
        client_account_sheet.range("G4").value = method
        client_account_sheet.range("H4").value = amount

        return (
            f"Clientes: Recibo de pago N°{math.floor(receipt)} guardado correctamente."
        )

    elif paymentType == "provider":
        # Recuperar valores de la hoja PAGOS
        date = pagos_sheet.range("B10").value
        client = pagos_sheet.range("D10").value
        receipt = pagos_sheet.range("E10").value
        method = pagos_sheet.range("G10").value
        amount = pagos_sheet.range("H10").value

        provider_account_sheet.api.Rows(4).Insert(Shift=-4121)

        # Insertar valores en la nueva fila de la hoja CTACTEPROV
        provider_account_sheet.range("B4").value = date
        provider_account_sheet.range("C4").value = client
        provider_account_sheet.range("F4").value = receipt
        provider_account_sheet.range("G4").value = method
        provider_account_sheet.range("H4").value = amount

        return f"Proveedores: Recibo de pago N°{math.floor(receipt)} guardado correctamente."


def new_receipt(paymentType):
    wb = xw.Book.caller()
    pagos_sheet = wb.sheets["PAGOS"]
    receipt_sheet = wb.sheets["RECIBOS"]

    if paymentType == "client":
        # Recuperar valores de la hoja PAGOS
        payment_date = pagos_sheet.range("B3").value
        client = pagos_sheet.range("D3").value
        receipt = pagos_sheet.range("E3").value
        method = pagos_sheet.range("G3").value
        amount = pagos_sheet.range("H3").value

        # Insertar valores en la hoja RECIBOS
        receipt_sheet.range("A1").value = date.today().strftime("%m/%d/%Y")
        receipt_sheet.range("B3:C3").value = receipt
        receipt_sheet.range("B4:C4").value = client
        receipt_sheet.range("C9:E9").value = payment_date
        receipt_sheet.range("C10:E10").value = client
        receipt_sheet.range("C12:E12").value = amount
        receipt_sheet.range("C13:E13").value = method
        receipt_sheet.range("C14:E14").value = "Carlos Olivieri"

        return f"Clientes: Recibo N°{math.floor(receipt)} listo para imprimir."

    elif paymentType == "provider":
        # Recuperar valores de la hoja PAGOS
        payment_date = pagos_sheet.range("B10").value
        client = pagos_sheet.range("D10").value
        receipt = pagos_sheet.range("E10").value
        method = pagos_sheet.range("G10").value
        amount = pagos_sheet.range("H10").value

        # Insertar valores en la hoja RECIBOS
        receipt_sheet.range("A1").value = date.today().strftime("%m/%d/%Y")
        receipt_sheet.range("B3:C3").value = receipt
        receipt_sheet.range("B4:C4").value = client
        receipt_sheet.range("C9:E9").value = payment_date
        receipt_sheet.range("C10:E10").value = "Carlos Olivieri"
        receipt_sheet.range("C12:E12").value = amount
        receipt_sheet.range("C13:E13").value = method
        receipt_sheet.range("C14:E14").value = client

        return f"Proveedores: Recibo N°{math.floor(receipt)} listo para imprimir."


def new_payment(paymentType):
    pagos_sheet = xw.Book.caller().sheets["PAGOS"]

    if paymentType == "client":
        current_value = pagos_sheet.range("E3").value or 0
        new_value = pagos_sheet.range("E3").value = current_value + 1

        pagos_sheet.range("B3").value = date.today().strftime("%m/%d/%Y")
        pagos_sheet.range("C3").value = 0
        pagos_sheet.range("F3").value = 0
        pagos_sheet.range("H3").value = 0

        return f"Clientes: Nuevo recibo N°{math.floor(new_value)} disponible."

    elif paymentType == "provider":
        current_value = pagos_sheet.range("E10").value or 0
        new_value = pagos_sheet.range("E10").value = current_value + 1

        pagos_sheet.range("B10").value = date.today().strftime("%m/%d/%Y")
        pagos_sheet.range("C10").value = 0
        pagos_sheet.range("F10").value = 0
        pagos_sheet.range("H10").value = 0

        return f"Proveedores: Nuevo recibo N°{math.floor(new_value)} disponible."
