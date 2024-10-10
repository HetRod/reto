import pandas as pd
import warnings
import re
import os
import pdfplumber
import re
from datetime import datetime


warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


meses = {
    'Enero': '01', 'Febrero': '02', 'Marzo': '03', 'Abril': '04',
    'Mayo': '05', 'Junio': '06', 'Julio': '07', 'Agosto': '08',
    'Septiembre': '09', 'Octubre': '10', 'Noviembre': '11', 'Diciembre': '12'
}


def cargar_patrones(archivo_txt):
    patrones = {}
    with open(archivo_txt, 'r') as f:
        for linea in f:
            nombre_patron, regex_patron = linea.strip().split('=')
            patrones[nombre_patron] = re.compile(regex_patron.strip()[2:-1])  # Remover r'' del patrón
    return patrones


def procesar_valor_pagado(valor_str):
    if ',' in valor_str:
        valor_str = valor_str.split('.')[0]  
        valor_str = valor_str.replace(',', '')  
    else:
        valor_str = valor_str.replace('.', '')  
    
    return valor_str

def extraer_datos_pdf(ruta_pdf, patrones):
    resultados = {}
    
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()

            # Buscar los valores usando los patrones cargados
            factura_pagada = patrones['patron_factura'].search(texto)
            if factura_pagada:
                resultados['factura_pagada'] = factura_pagada.group(1)

            nro_factura = patrones['patron_nro_factura'].search(texto)
            if nro_factura:
                resultados['nro_factura'] = nro_factura.group(1)

            valor_pagado = patrones['patron_valor'].search(texto)
            if valor_pagado:
                resultados['valor_pagado'] = procesar_valor_pagado(valor_pagado.group(2))

            fecha_hora_transaccion = patrones['patron_fecha_hora'].search(texto)
            if fecha_hora_transaccion:
                fecha_hora_str = fecha_hora_transaccion.group(0)
                dia, mes, anio = re.search(r'(\d{1,2})\s+de\s+([A-Za-z]+)\s+de\s+(\d{4})', fecha_hora_str).groups()
                mes_num = meses[mes.capitalize()]  
                resultados['fecha_hora_transaccion'] = f"{anio}{mes_num}{dia.zfill(2)}"

            fecha_envio_pago = patrones['patron_fecha_envio'].search(texto)
            if fecha_envio_pago:
                fecha_envio_str = fecha_envio_pago.group(1)
                dia, mes, anio = fecha_envio_str.split('-')
                resultados['fecha_envio_pago'] = f"{anio}{mes}{dia}"

    return resultados

def procesar_pdf_con_parametros(nombre_pdf, num_factura, valor_pagado, fecha_envio_pago):
    ruta_pdf = os.path.join(os.getcwd(), nombre_pdf + '.pdf')
    if not os.path.exists(ruta_pdf):
        # print(f"El archivo {nombre_pdf} no se encontró en la carpeta actual.")
        return

    patrones = cargar_patrones('patrones.txt')
    resultados_pdf = extraer_datos_pdf(ruta_pdf, patrones)
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
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# Función para procesar y actualizar la columna OBSERVACIONES
def actualizar_observaciones(archivo_excel, hoja, nombre_pdf, coincidencias):
    # Cargar el archivo Excel especificando la hoja
    df = pd.read_excel(archivo_excel, sheet_name=hoja)
    
    # Limpiar los nombres de las columnas para eliminar espacios y saltos de línea
    df.columns = df.columns.str.strip().str.replace('\n', '', regex=False)
    
    # Iterar por cada fila y buscar donde el NUMERO DE COMPROBANTE coincide con el nombre del PDF
    for i, row in df.iterrows():
        if str(row['NUMERO DE COMPROBANTE']) == nombre_pdf:
            # Si coincide, escribimos las coincidencias en la columna OBSERVACIONES
            df.at[i, 'OBSERVACIONES'] = str(coincidencias)
    
    # Guardar el archivo Excel actualizado
    df.to_excel(archivo_excel, sheet_name=hoja, index=False)
    print(f"Observaciones actualizadas para {nombre_pdf} con coincidencias {coincidencias}")

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# Función para procesar y actualizar la columna OBSERVACIONES
def actualizar_observaciones(archivo_excel, hoja, nombre_pdf, coincidencias):
    # Cargar el archivo Excel especificando la hoja
    df = pd.read_excel(archivo_excel, sheet_name=hoja)
    
    # Limpiar los nombres de las columnas para eliminar espacios y saltos de línea
    df.columns = df.columns.str.strip().str.replace('\n', '', regex=False)

    if coincidencias == [0, 0, 0]:
        coincidencias = 'Ninguna coincidencia'
    elif coincidencias == [1, 0, 0]:
        coincidencias = 'No coincide el valor pagado ni la fecha de envío'
    elif coincidencias == [0, 1, 0]:
        coincidencias = 'No coincide el número de factura ni la fecha de envío'
    elif coincidencias == [0, 0, 1]:
        coincidencias = 'No coincide el número de factura ni el valor pagado'
    elif coincidencias == [1, 1, 0]:
        coincidencias = 'No coincide la fecha de envío'
    elif coincidencias == [1, 0, 1]:
        coincidencias = 'No coincide el valor pagado'
    elif coincidencias == [0, 1, 1]:
        coincidencias = 'No coincide el número de factura'
    elif coincidencias == [1, 1, 1]:
        coincidencias = ''
    
    
    # Iterar por cada fila y buscar donde el NUMERO DE COMPROBANTE coincide con el nombre del PDF
    for i, row in df.iterrows():
        if str(row['NUMERO DE COMPROBANTE']) == nombre_pdf:
            # Si coincide, escribimos las coincidencias en la columna OBSERVACIONES
            df.at[i, 'OBSERVACIONES BANCO'] = str(coincidencias)
    
    # Guardar el archivo Excel actualizado
    df.to_excel(archivo_excel, sheet_name=hoja, index=False)
    print(f"Observaciones actualizadas para {nombre_pdf} con coincidencias {coincidencias}")

# Adaptación de la función enviar_datos para incluir la actualización de observaciones
def enviar_datos(num_comprobante, num_factura, valor_pagado, fecha_envio_pago, archivo_excel, hoja):
    resultado = procesar_pdf_con_parametros(num_comprobante, num_factura, valor_pagado, fecha_envio_pago)
    
    if resultado is None:
        # print(f"Archivo PDF no encontrado para el comprobante {num_comprobante}.")
        return  
    else:
        print(f"Comprobante de excel: {num_comprobante}, Factura: {num_factura}, Valor Pagado: {valor_pagado}, Fecha de Envío: {fecha_envio_pago}")
    
    coincidencias = resultado['coincidencias']
    
    # Actualizar las observaciones en el archivo Excel
    actualizar_observaciones(archivo_excel, hoja, num_comprobante, coincidencias)

# Procesar cada fila del Excel
for _, row in resultado.iterrows():
    enviar_datos(
        num_comprobante=row['NUMERO DE COMPROBANTE'],
        num_factura=row['# FACTURA PAGADA'],
        valor_pagado=row['TOTAL'],
        fecha_envio_pago=row['F_PAGO AAAAMMDD'],
        archivo_excel=archivo_excel,  # Pasar el archivo Excel
        hoja='SELLO'                  # Pasar la hoja correspondiente
    )
