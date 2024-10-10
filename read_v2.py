import pandas as pd
import win32com.client

# Especifica la ruta del archivo Excel
ruta_archivo = r'C:\Users\hmrodrig\Downloads\legalizacion.xlsm'

# Inicializar la aplicación de Excel
excel = win32com.client.Dispatch("Excel.Application")

# Hacer que Excel sea visible (opcional)
excel.Visible = False

# Abrir el archivo .xlsm
workbook = excel.Workbooks.Open(r'C:\Users\hmrodrig\Downloads\legalizacion.xlsm')

# Seleccionar la hoja (por nombre o índice)
sheet = workbook.Sheets('SELLO')

# Modificar una celda
sheet.Cells(2, 36).Value = 'Prueba de observacions'  # Cambiar el valor de la celda AJ2
sheet.Range("AJ3").Value = 'Prueba observacion con otro metodo' # Cambiar el valor de la celda AJ3

# Guardar los cambios
workbook.Save()

# Cerrar el libro y la aplicación de Excel
workbook.Close(SaveChanges=True)
excel.Quit()
print(f"Archivo modificado guardado en: {ruta_archivo}")
