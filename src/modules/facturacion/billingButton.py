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


def save_sales_average():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    billing_sheet = wb.sheets["FACTURACION"]
    sales_sheet = wb.sheets["PROMEDIO VENTAS"]

    rows_to_save = check_rows()
    if len(rows_to_save) != 0:
        # Get the date and client name from FACTURACION sheet
        date = billing_sheet.range("C2").value
        client_name = billing_sheet.range(f"D8").value
        
        for row in rows_to_save:
            sales_sheet.api.Rows(3).Insert(
                Shift=-4121
            )  # Insert a new row at position 3

            # Paste each row values to PROMEDIO VENTAS sheet
            # date
            sales_sheet.range("B3").value = date
            # client_name
            sales_sheet.range("C3").value = client_name
            # quantity
            sales_sheet.range("D3").value = billing_sheet.range(f"D{row}").value
            # article
            sales_sheet.range("E3").value = billing_sheet.range(f"D{row}").value
            # price
            sales_sheet.range("F3").value = billing_sheet.range(f"D{row}").value
            # article_total
            sales_sheet.range("G3").value = billing_sheet.range(f"D{row}").value


def save_billing():
    # Get the active workbook and sheet
    wb = xw.Book.caller()
    billing_sheet = wb.sheets["FACTURACION"]
    account_sheet = wb.sheets["CTACTE"]

    row = check_rows()
    if len(row) != 0:
        account_sheet.api.Rows(4).Insert(Shift=-4121)  # Insert a new row at position 4

        # Retrieve values from FACTURACION sheet
        client_name = billing_sheet.range("D8").value
        billing_number = billing_sheet.range("D4").value
        billing_total = billing_sheet.range("I34").value

        # Paste values to CTACTE sheet
        account_sheet.range(f"B4").value = date.today().strftime("%m/%d/%Y")
        account_sheet.range(f"C4").value = client_name
        account_sheet.range(f"D4").value = billing_number
        account_sheet.range(f"E4").value = billing_total

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


def new_billing():
    # Get the active workbook and sheet
    billing_sheet = xw.Book.caller().sheets["FACTURACION"]

    rows_to_bill = check_rows()
    if len(rows_to_bill) != 0:
        # Get current billing number, ensuring it's a number (default to 0 if empty)
        current_billing = billing_sheet.range("D4").value or 0
        new_billing = billing_sheet.range("D4").value = current_billing + 1
        billing_sheet.range("C2").value = date.today().strftime("%m/%d/%Y")

        # Reset order number
        billing_sheet.range(f"D5").value = 0
        # Delete all data from the rows to bill
        for row in rows_to_bill:
            #############################################################################################################
            # W.I.P Falta agregar el reseteo de los valores de las celdas cantidad desde la hoja GALPON
            #############################################################################################################
            billing_sheet.range(f"G{row}").value = 0

        # Displays a message box
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        messagebox.showinfo(
            "Facturacion: nueva factura creada",
            f"Se ha creado la factura N° {math.floor(new_billing)}.",
        )  # Info message box


# MAIN FUNCTION
def save_reset_button():
    save_billing(),  # First step is save the billing
    save_sales_average(),  # Second step is save the sales average
    new_billing()  # Third step is delete all data and create a new billing
