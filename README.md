# 📋 Excel Macro buttons - Finance Business App
1ra Etapa: Mejoras en los codigos Macro, organización y limpieza del archivo.

2da Etapa: Proceso de transformación y desarrollo del Archivo Excel y funcionalidades a una App funcional que cumple estandares de calidad.

# ⚙️ Detalles Técnicos 
Desarrollaremos codigos de Excel Macros para que realicen distintas funciones necesarias para las finanzas de un negocio.

**Lenguaje de programación elegido:** Python 

**Bibliotecas utilizadas:** xlwings

## Instalación de xlwings en Windows 
### Requisitos Previos

Antes de comenzar con la instalación, asegúrate de cumplir con los siguientes requisitos:

- Tener instalado Microsoft Excel (versión 2016 o superior recomendada).

- Tener instalado Python (versión 3.7 o superior recomendada).

- Acceso a la línea de comandos (cmd o PowerShell).

### Paso 1: Verificar la instalación de Python

Abre la línea de comandos y ejecuta el siguiente comando para verificar si Python está instalado:

```sh
python --version
```

Si Python no está instalado, descárgalo e instálalo desde Python.org. Durante la instalación, asegúrate de marcar la opción "Add Python to PATH".

### Paso 2: Instalar xlwings

Ejecuta el siguiente comando en la línea de comandos para instalar xlwings:
```sh
pip install xlwings
```
Si deseas instalar una versión específica, usa:
```sh
pip install xlwings==0.X.X
```
(reemplaza 0.X.X con la versión deseada).

### Paso 3: Habilitar xlwings en Excel

Después de instalar xlwings, es necesario habilitar el complemento en Excel. Para ello, ejecuta:
```sh
xlwings addin install
```
Este comando instalará el complemento de xlwings en Excel.

Si Excel estaba abierto, ciérralo y ábrelo nuevamente.

### Paso 4: Habilitar Macros en Excel

Para que xlwings funcione correctamente, es necesario habilitar macros en Excel:

- Abre Excel y ve a Archivo > Opciones.

- En el panel izquierdo, selecciona Centro de confianza > Configuración del Centro de confianza.

- Ve a Configuración de macros y selecciona Habilitar todas las macros (puede ser menos restrictivo según tus necesidades).

- Guarda los cambios y cierra Excel.

### Paso 5: Configurar xlwings

Ejecuta el siguiente comando para generar un archivo de configuración en tu proyecto:

xlwings quickstart mi_proyecto

Este comando creará una carpeta mi_proyecto con los archivos necesarios para comenzar a trabajar con xlwings.

### Paso 6: Probar la Instalación

Para verificar que xlwings está funcionando correctamente, abre Python y ejecuta:
```sh
import xlwings as xw
wb = xw.Book()
sht = wb.sheets[0]
sht.range("A1").value = "Hola, xlwings!"
```

Si todo está configurado correctamente, Excel debería abrirse con el mensaje "Hola, xlwings!" en la celda A1.

### Paso 7: Solución de Problemas

Si xlwings no funciona correctamente, prueba las siguientes soluciones:

Asegúrate de que Python y xlwings estén instalados correctamente con:
```sh
python -m pip show xlwings
```

Si xlwings no reconoce Excel, intenta reinstalar el complemento con:
```sh
xlwings addin install
```

Si el error persiste, revisa la documentación oficial en xlwings.org.

¡Listo! Ahora tienes xlwings instalado y listo para usar en Windows.



# 🔗 Contacto
[![linkedin](https://img.shields.io/badge/linkedin-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ezequiel-matias-rodriguez-094b76188/)
[![portfolio](https://img.shields.io/badge/my_portfolio-000?style=for-the-badge&logo=ko-fi&logoColor=white)](https://github.com/Eze-Rodriguez)