from xlwings import Book


def goToFacturacion():
    # References the book when the Python function is called from Excel via RunPython.
    book = Book.caller()
    # Define FACTURACION sheet and activate it
    facturacionSheet = book.sheets["FACTURACION"]
    facturacionSheet.activate()

    # Select a specific cell (Range)
    facturacionSheet.range("A1").select()
