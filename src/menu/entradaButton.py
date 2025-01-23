from xlwings import Book


def goToEntrada():
    # References the book when the Python function is called from Excel via RunPython.
    book = Book.caller()
    # Define ENTRADA sheet and activate it
    entradaSheet = book.sheets["ENTRADA"]
    entradaSheet.activate()

    # Select a specific cell (Range)
    entradaSheet.range("A1").select()
