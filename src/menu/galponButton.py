from xlwings import Book


def goToGalpon():
    # References the book when the Python function is called from Excel via RunPython.
    book = Book.caller()
    # Define GALPON sheet and activate it
    galponSheet = book.sheets["GALPON"]
    galponSheet.activate()

    # Select a specific cell (Range)
    galponSheet.range("A1").select()
