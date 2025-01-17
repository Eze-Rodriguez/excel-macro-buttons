# Importar solo una funcion
from modules.function1 import consoleLog
funcion = consoleLog("pepe")
print(funcion)

# Importar todo el archivo
import modules.function1
funcion1 = modules.function1.consoleLog("eusebio")
print(funcion1)