import pandas as pd
import warnings
import re
import os
import pdfplumber
from datetime import datetime


warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


meses = {
    'Enero': '01', 'Febrero': '02', 'Marzo': '03', 'Abril': '04',
    'Mayo': '05', 'Junio': '06', 'Julio': '07', 'Agosto': '08',
    'Septiembre': '09', 'Octubre': '10', 'Noviembre': '11', 'Diciembre': '12'
}

def procesar_valor_pagado(valor_str):
    if ',' in valor_str:
        valor_str = valor_str.split('.')[0]  
        valor_str = valor_str.replace(',', '')  
    else:
        valor_str = valor_str.replace('.', '')  
    
    return valor_str

def extraer_datos_pdf(ruta_pdf):
    resultados = {}

    
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
    
    print("resultadoss en pdf",resultados)

    return resultados

def procesar_pdf_con_parametros(nombre_pdf, num_factura, valor_pagado, fecha_envio_pago):
    ruta_pdf = os.path.join(os.getcwd(), nombre_pdf + '.pdf')
    if not os.path.exists(ruta_pdf):
        # print(f"El archivo {nombre_pdf} no se encontró en la carpeta actual.")
        return

    resultados_pdf = extraer_datos_pdf(ruta_pdf)
    coincidencias = []

    coincidencias.append(1 if str(resultados_pdf.get('nro_factura')) == str(num_factura) else 0)
    coincidencias.append(1 if str(resultados_pdf.get('valor_pagado')) == str(valor_pagado) else 0)

    if (str(resultados_pdf.get('fecha_envio_pago')) == str(fecha_envio_pago) or 
        str(resultados_pdf.get('fecha_hora_transaccion')) == str(fecha_envio_pago)):
        coincidencias.append(1)
    else:
        coincidencias.append(0)

    return {
        'archivo': nombre_pdf,
        'coincidencias': coincidencias
    }


def agrupar_y_sumar_total(archivo_excel, columnas_a_extraer, hoja):
    # Cargar el archivo Excel especificando la hoja
    df = pd.read_excel(archivo_excel, sheet_name=hoja)
    
    # Limpiar los nombres de las columnas para eliminar espacios y saltos de línea
    df.columns = df.columns.str.strip().str.replace('\n', '', regex=False)

    # Filtrar las columnas que necesitamos
    df_filtrado = df[columnas_a_extraer]
    
    # Reemplazar valores no numéricos en 'F_PAGO AAAAMMDD' con 0
    df_filtrado['F_PAGO AAAAMMDD'] = pd.to_numeric(df_filtrado['F_PAGO AAAAMMDD'], errors='coerce').fillna(0).astype(int)
    
    # Agrupar por 'NUMERO DE COMPROBANTE' y sumar la columna 'TOTAL'
    df_agrupado = df_filtrado.groupby('NUMERO DE COMPROBANTE').agg({
        '# FACTURA PAGADA': 'first',  
        'F_PAGO AAAAMMDD': 'first',   
        'TOTAL': 'sum'               
    }).reset_index()

    return df_agrupado

archivo_excel = 'SELLO LEGALIZACION ADMINISTRACION ANTICIPO 26 PARTE 6 JURIDICOS.xlsm'

columnas = ['NUMERO DE COMPROBANTE', '# FACTURA PAGADA', 'F_PAGO AAAAMMDD', 'TOTAL']

resultado = agrupar_y_sumar_total(archivo_excel, columnas, hoja='SELLO')


def enviar_datos(num_comprobante, num_factura, valor_pagado, fecha_envio_pago):
    resultado = procesar_pdf_con_parametros(num_comprobante, num_factura, valor_pagado, fecha_envio_pago)
    
    
    if resultado is None:
        # print(f"Archivo PDF no encontrado para el comprobante {num_comprobante}.")
        return  
    else:
        print(f"Comprobante de excel: {num_comprobante}, Factura: {num_factura}, Valor Pagado: {valor_pagado}, Fecha de Envío: {fecha_envio_pago}")

    
    
    print("\nResultado de la extracción de datos:")
    print(f"Archivo: {resultado['archivo']}, Coincidencias: {resultado['coincidencias']}")


for _, row in resultado.iterrows():
    enviar_datos(
        num_comprobante=row['NUMERO DE COMPROBANTE'],
        num_factura=row['# FACTURA PAGADA'],
        valor_pagado=row['TOTAL'],
        fecha_envio_pago=row['F_PAGO AAAAMMDD']
    )









