import pdfplumber
import re
import os
from datetime import datetime

# Mapeo de meses en español a su formato numérico
meses = {
    'Enero': '01', 'Febrero': '02', 'Marzo': '03', 'Abril': '04',
    'Mayo': '05', 'Junio': '06', 'Julio': '07', 'Agosto': '08',
    'Septiembre': '09', 'Octubre': '10', 'Noviembre': '11', 'Diciembre': '12'
}

def procesar_valor_pagado(valor_str):
    if ',' in valor_str:
        valor_str = valor_str.split('.')[0]  # Eliminar lo que esté después del punto
        valor_str = valor_str.replace(',', '')  # Eliminar las comas
    else:
        valor_str = valor_str.replace('.', '')  # Eliminar el punto si no hay coma
    
    return valor_str

# Función para extraer datos del PDF
def extraer_datos_pdf(ruta_pdf):
    resultados = {}

    # Patrones de búsqueda
    patron_factura = r'Factura\s*No\s*:\s*(\d+)'  
    patron_nro_factura = r'Nro\.?\s*de\s*factura\s*:\s*(\d+)'  
    patron_valor = r'(Valor Total del Pago|total a pagar|Total a Pagar|pago total|Valor pagado)\s*:?\s*\$?\s*([\d{1,3}(?:,\d{3})*(?:\.\d{2})?]+)'  
    patron_fecha_hora = r'\d{1,2}\s+de\s+(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+de\s+\d{4}'
    patron_fecha_envio = r'Fecha de envío del pago\s*:\s*(\d{2}-\d{2}-\d{4})'

    # Abrir y procesar el PDF
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()

            factura_pagada = re.search(patron_factura, texto, flags=re.IGNORECASE)
            if factura_pagada:
                resultados['factura_pagada'] = factura_pagada.group(1)

            nro_factura = re.search(patron_nro_factura, texto, flags=re.IGNORECASE)
            if nro_factura:
                resultados['nro_factura'] = nro_factura.group(1)

            valor_pagado = re.search(patron_valor, texto, flags=re.IGNORECASE)
            if valor_pagado:
                valor_str = valor_pagado.group(2)
                resultados['valor_pagado'] = procesar_valor_pagado(valor_str)

            fecha_hora_transaccion = re.search(patron_fecha_hora, texto, flags=re.IGNORECASE)
            if fecha_hora_transaccion:
                fecha_hora_str = fecha_hora_transaccion.group(0)
                try:
                    dia, mes, anio = re.search(r'(\d{1,2})\s+de\s+([A-Za-z]+)\s+de\s+(\d{4})', fecha_hora_str).groups()
                    mes_num = meses[mes.capitalize()]  
                    resultados['fecha_hora_transaccion'] = f"{anio}{mes_num}{dia.zfill(2)}"
                except ValueError as e:
                    print(f"Error al procesar la fecha: {fecha_hora_str}. Error: {e}")

            fecha_envio_pago = re.search(patron_fecha_envio, texto, flags=re.IGNORECASE)
            if fecha_envio_pago:
                fecha_envio_str = fecha_envio_pago.group(1)
                dia, mes, anio = fecha_envio_str.split('-')
                resultados['fecha_envio_pago'] = f"{anio}{mes}{dia}"

    return resultados

# Función para procesar un solo PDF y comparar los datos extraídos con los parámetros
def procesar_pdf_con_parametros(nombre_pdf, num_factura, valor_pagado, fecha_envio_pago):
    ruta_pdf = os.path.join(os.getcwd(), nombre_pdf)
    if not os.path.exists(ruta_pdf):
        print(f"El archivo {nombre_pdf} no se encontró en la carpeta actual.")
        return

    resultados_pdf = extraer_datos_pdf(ruta_pdf)
    coincidencias = []

    coincidencias.append(1 if resultados_pdf.get('nro_factura') == num_factura else 0)
    coincidencias.append(1 if resultados_pdf.get('valor_pagado') == valor_pagado else 0)

    if (resultados_pdf.get('fecha_envio_pago') == fecha_envio_pago or 
        resultados_pdf.get('fecha_hora_transaccion') == fecha_envio_pago):
        coincidencias.append(1)
    else:
        coincidencias.append(0)

    return {
        'archivo': nombre_pdf,
        'coincidencias': coincidencias
    }

# Llamada a la función con los parámetros de búsqueda
nombre_pdf = '1 C-266.pdf'  # Nombre del archivo PDF
num_factura = '151117492-4'  # Número de comprobante a buscar
valor_pagado = '47110'  # Valor pagado a buscar
fecha_envio_pago = '20240805'  # Fecha de envío a buscar en formato AAAAMMDD

resultado = procesar_pdf_con_parametros(nombre_pdf, num_factura, valor_pagado, fecha_envio_pago)

# Imprimir los resultados
if resultado:
    print("\nResultado de la extracción de datos:")
    print(f"Archivo: {resultado['archivo']}, Coincidencias: {resultado['coincidencias']}")
