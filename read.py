import pandas as pd

# Especifica la ruta del archivo Excel
ruta_archivo = r'C:\Users\hmrodrig\Downloads\legalizacion.xlsm'

# Lee el archivo Excel
df = pd.read_excel(ruta_archivo)

# se hace cambio sobre la columna observaciones banco para añadir texto adicional
df['OBSERVACIONES BANCO'] = " fecha de pago incorrecta"

# Guardar los cambios en un nuevo archivo Excel
nueva_ruta_archivo = r'C:\Users\hmrodrig\Downloads\document_modificado2.xlsx'
df.to_excel(nueva_ruta_archivo, index=False)

print(f"Archivo modificado guardado en: {nueva_ruta_archivo}")
