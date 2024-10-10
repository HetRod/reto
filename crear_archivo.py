import pandas as pd
import warnings
import re
import os
import pdfplumber
import re
import sys
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

            nro_factura_electronica = patrones['patron_nro_factura_electronica'].search(texto)
            if nro_factura_electronica:
                resultados['nro_factura_electronica'] = nro_factura_electronica.group(1)

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
        return [nombre_pdf, 0, 0, 0]

    patrones = cargar_patrones('patrones.txt')
    resultados_pdf = extraer_datos_pdf(ruta_pdf, patrones)
    coincidencias = []


    coincidencias.append(nombre_pdf)
    if (str(resultados_pdf.get('nro_factura_electronica')) == str(num_factura).strip() or 
        str(resultados_pdf.get('nro_factura')) == str(num_factura).strip()):
        coincidencias.append(1)
    else:
        coincidencias.append(0)
    coincidencias.append(1 if str(resultados_pdf.get('valor_pagado')) == str(valor_pagado) else 0)

    if (str(resultados_pdf.get('fecha_envio_pago')) == str(fecha_envio_pago) or 
        str(resultados_pdf.get('fecha_hora_transaccion')) == str(fecha_envio_pago)):
        coincidencias.append(1)
    else:
        coincidencias.append(0)



    print(f"Coincidencias para el comprobante excel {nombre_pdf}: {num_factura} - {valor_pagado} - {fecha_envio_pago}")
    print(f"Coincidencias para el comprobante PDF {nombre_pdf}: {resultados_pdf.get('nro_factura')} - {resultados_pdf.get('valor_pagado')} - {resultados_pdf.get('fecha_envio_pago')} - {resultados_pdf.get('fecha_hora_transaccion')}")
    print(f"Coincidencias: {coincidencias}")
    return coincidencias
    


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

# Función para procesar y actualizar la columna OBSERVACIONES
def actualizar_observaciones(archivo_excel, hoja, nombre_pdf, coincidencias):
    # Cargar el archivo Excel especificando la hoja
    df = pd.read_excel(archivo_excel, sheet_name=hoja)
    
    # Limpiar los nombres de las columnas para eliminar espacios y saltos de línea
    df.columns = df.columns.str.strip().str.replace('\n', '', regex=False)

    

    if coincidencias[-3:] == [0, 0, 0]:
        observacion = 'Ninguna coincidencia'
    elif coincidencias[-3:] == [1, 0, 0]:
        observacion = 'No coincide el valor pagado ni la fecha de envío'
    elif coincidencias[-3:] == [0, 1, 0]:
        observacion = 'No coincide el número de factura ni la fecha de envío'
    elif coincidencias[-3:] == [0, 0, 1]:
        observacion = 'No coincide el número de factura ni el valor pagado'
    elif coincidencias[-3:] == [1, 1, 0]:
        observacion = 'No coincide la fecha de envío'
    elif coincidencias[-3:] == [1, 0, 1]:
        observacion = 'No coincide el valor pagado'
    elif coincidencias[-3:] == [0, 1, 1]:
        observacion = 'No coincide el número de factura'
    elif coincidencias[-3:] == [1, 1, 1]:
        observacion = ''
    
    
    # Iterar por cada fila y buscar donde el NUMERO DE COMPROBANTE coincide con el nombre del PDF
    for i, row in df.iterrows():
        if str(row['NUMERO DE COMPROBANTE']) == nombre_pdf:
            # Si coincide, escribimos las coincidencias en la columna OBSERVACIONES
            df.at[i, 'OBSERVACIONES BANCO'] = str(observacion)
    
    # Guardar el archivo Excel actualizado
    df.to_excel(archivo_excel, sheet_name=hoja, index=False)

# Adaptación de la función enviar_datos para incluir la actualización de observaciones
def enviar_datos(num_comprobante, num_factura, valor_pagado, fecha_envio_pago, archivo_excel, hoja):
    resultado = procesar_pdf_con_parametros(num_comprobante, num_factura, valor_pagado, fecha_envio_pago)
    
    if resultado is None:
        # print(f"Archivo PDF no encontrado para el comprobante {num_comprobante}.")
        return  
    # else:
    #     print(f"Comprobante de excel: {num_comprobante}, Factura: {num_factura}, Valor Pagado: {valor_pagado}, Fecha de Envío: {fecha_envio_pago} ")
    
    
    # print(armar_html(resultado))

    
    # Actualizar las observaciones en el archivo Excel
    actualizar_observaciones(archivo_excel, hoja, num_comprobante, resultado)

    return resultado

def crear_archivo_html(nombre_archivo, contenido):
    with open(nombre_archivo, 'w') as archivo:
        archivo.write(contenido) 

def armar_html(data):
    data_html = ""
    for sub_arreglo in data:
        data_html += f'{{ comprobante: "{sub_arreglo[0]}", factura:"{sub_arreglo[1]}", fecha:"{sub_arreglo[2]}", valor:"{sub_arreglo[3]}" }},\n'
    if data_html.endswith(",\n"):
        data_html = data_html[:-2] + "\n"
    return data_html

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python crear_archivo.py <nombre_archivo>")
    else:
        archivo_excel = sys.argv[1]

#parametro de entrada por input
# archivo_excel = 'SELLO LEGALIZACION ADMINISTRACION ANTICIPO 26 PARTE 6 JURIDICOS.xlsm'

        columnas = ['NUMERO DE COMPROBANTE', '# FACTURA PAGADA', 'F_PAGO AAAAMMDD', 'TOTAL']

        resultado = agrupar_y_sumar_total(archivo_excel, columnas, hoja='SELLO')
        resultados_totales =[]
        # Procesar cada fila del Excel
        for _, row in resultado.iterrows():
            resultados_totales.append(
            enviar_datos(
                num_comprobante=row['NUMERO DE COMPROBANTE'],
                num_factura=row['# FACTURA PAGADA'],
                valor_pagado=row['TOTAL'],
                fecha_envio_pago=row['F_PAGO AAAAMMDD'],
                archivo_excel=archivo_excel,  # Pasar el archivo Excel
                hoja='SELLO'                  # Pasar la hoja correspondiente
            )
        )



        contenido_html = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <meta http-equiv="X-UA-Compatible" content="ie=edge">
            <title>Legalización Gastos </title>
        
            <!-- Add icon library -->
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        
            <!--== Favicon ==-->
            <link rel="shortcut icon" href="https://raw.githubusercontent.com/HetRod/prueba/refs/heads/master/favicon.png" type="image/x-icon" />
        
            <!--== Google Fonts ==-->
            <link href="https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,500;0,700;0,900;1,300;1,400&amp;display=swap" rel="stylesheet">
        
            <style>
                .header-wrapper {
                    padding: 20px 0;
                }
                
                .pagination {
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    margin-top: 20px;
                    margin-bottom: 20px;
                    padding: 10px;
                    background-color: #f1f1f1;
                    border: 1px solid #e0e0e0;
                    border-radius: 8px;
                    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                }
        
                .pagination button {
                    margin: 0 10px;
                    padding: 8px 12px;
                    border: none;
                    border-radius: 5px;
                    background-color: #007bff;
                    color: white;
                    cursor: pointer;
                    transition: background-color 0.3s;
                }
        
                .pagination button:hover:not(:disabled) {
                    background-color: #0056b3;
                }
        
                .pagination button:disabled {
                    opacity: 0.5;
                    cursor: not-allowed;
                }
        
                .pagination span {
                    font-weight: bold;
                    margin: 0 10px;
                }
        
            
                img {
                    max-width: 50%;
                    vertical-align: middle;
                }
        
                        
                body {
                    color: #555555;
                    font-size: 16px;
                    font-family: "Roboto", sans-serif;
                    font-weight: 400;
                    line-height: 1.6;
                    margin: 0;
                }
                
            
        
                .badge-item__body {
                        -webkit-flex-basis: 100%;
                        -ms-flex-preferred-size: 100%;
                        flex-basis: 100%;
                        width: 100%;
                        text-align: center;
                    }
        
                .badge-item__body {
                    -webkit-flex-basis: calc(100% - 250px);
                    -ms-flex-preferred-size: calc(100% - 250px);
                    flex-basis: calc(100% - 250px);
                    width: calc(100% - 250px);
                    display: -webkit-box;
                    display: -webkit-flex;
                    display: -ms-flexbox;
                    display: flex;
                    -webkit-align-self: center;
                    -ms-flex-item-align: center;
                    align-self: center;
                    padding: 18px 20px 18px 50px;
                    color: #414141;
                    font-size: 14px;
                }
        
                .container,
                .container-fluid,
                .container-lg,
                .container-md,
                .container-sm,
                .container-xl,
                .container-xxl {
                    --bs-gutter-x: 1.875rem;
                }
        
                @media screen and (min-width: 1680px) {
                    .container {
                        max-width: 1620px;
                    }
                }
        
                .justify-content-between{
                    justify-content:space-between!important
                }
        
                .row{
                    --bs-gutter-x:1.5rem;
                    --bs-gutter-y:0;
                    display:flex;
                    flex-wrap:wrap;
                    margin-top:calc(-1 * var(--bs-gutter-y));
                    margin-right:calc(-.5 * var(--bs-gutter-x));
                    margin-left:calc(-.5 * var(--bs-gutter-x))
                }
        
                .align-items-center{
                    align-items:center!important
                }
        
                .col-lg-3{
                    flex:0 0 auto;
                    width:25%
                }
        
                @media only screen and (max-width: 767.98px) {
                    .logo-wrap {
                        margin-bottom: 15px;
                        margin-left: 10 px;
                    }
                }
        
                @media only screen and (min-width: 768px) and (max-width: 991.98px) {
                    .logo-wrap {
                        text-align: center;
                        margin-bottom: 20px;
                        margin-left: 10 px;
                    }
                }
        
                .badge-tracker-item:not(:last-child) {
                    margin-bottom: 60px;
                }
        
                .col-lg-6{
                    flex:0 0 auto;
                    width:50%
                }
        
                .body-dark .main-content-wrapper .title-top h2 {
                    color: #FFFFFF;
                }
        
                .mtn-25 {
                    margin: auto;
                    text-align: -webkit-center;
                }
        
                .tracker-block {
                    background-color: #FFFFFF;
                    display: -webkit-box;
                    display: -webkit-flex;
                    display: -ms-flexbox;
                    display: flex;
                    -webkit-box-align: center;
                    -webkit-align-items: center;
                    -ms-flex-align: center;
                    align-items: center;
            
                }
        
            
                .tracker-block--5 .track-item__title,
                .tracker-block--5 .track-item__no,
                .tracker-block--5 .track-item__new {
                    color: #FFFFFF;
                    font-size: 13px;
                }
        
                .tracker-block--5 .track-item__no {
                    font-size: 40px;
                    font-weight: 700;
                    margin: auto;
                }
        
                .tracker-block--5 .track-item__new {
                    margin-bottom: 0;
                }
        
                .tracker-block--5.bg-yellow {
                    background-color: #FFC260;
                    width: 40%;
                
                }
                .col-sm-5{
                    flex:0 0 auto;
                    width:41.66666667%
                }
        
                .last-update-wrap {
                    color: #4D4D4D;
                    font-size: 14px;
                    font-style: italic;
                    font-weight: 300;
                    text-align: center;
                }
        
                @media only screen and (min-width: 768px) and (max-width: 991.98px) {
                    .last-update-wrap {
                        text-align: left;
                    }
                }
        
                @media only screen and (max-width: 767.98px) {
                    .last-update-wrap {
                        text-align: left;
                    }
                }
        
                .mb-0{
                    margin-bottom:0!important
                }
        
                .main-content-wrapper {
                    background-color: #F7F8FC;
                }
        
                @media only screen and (max-width: 767.98px) {
                    .main-content-wrapper {
                        padding: 60px 0;
                    }
                }
        
                .col-xl-9{
                    flex:0 0 auto;
                    width:75%
                }
        
                .m-auto{
                    margin:auto!important
                }
        
        
                .badge-item__title {
                    background-color: #b0bfe9;
                    color: #FFFFFF;
                    padding: 16px 35px;
                    -webkit-flex-basis: 250px;
                    -ms-flex-preferred-size: 250px;
                    flex-basis: 250px;
                    width: 250px;
                    text-align: center;
                }
        
                .badge-item__block {
                    position: relative;
                    padding-left: 40px;
                    margin-left: 40px;
                }
        
                
                .badge-item {
                    background-color: #f9f9f9;
                    border: 1px solid #e0e0e0;
                    border-radius: 8px;
                    width: 100%;
                    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                    display:block;
                }
        
                .badge-item__block:before {
                    background-color: #7AEEC2;
                    content: "";
                    position: absolute;
                    left: 0;
                    top: -3px;
                    height: 30px;
                    width: 1px;
                }
        
        
                .badge-item {
                    background-color: #DFFFF3;
                    display: -webkit-box;
                    display: -webkit-flex;
                    display: -ms-flexbox;
                    display: flex;
                    -webkit-align-content: center;
                    -ms-flex-line-pack: center;
                    align-content: center;
                    -webkit-box-pack: center;
                    -webkit-justify-content: center;
                    -ms-flex-pack: center;
                    justify-content: center;
                    margin-top: 5px;
                }
        
        
                .badge-bg-green {
                    background-color: #F1FCFC;
                }
        
                .title-top {
                    margin-bottom: 20px;
                    text-align: center;
                }
        
        
                .badge-content-wrapper {
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                }
        
        
            </style>
        </head>
        
        <body>
            <!--== Start Header Wrapper ==-->
            <header class="header-wrapper">
                <div class="container">
                    <div class="row justify-content-between align-items-center">
                        <div class="col-lg-3">
                            <div class="logo-wrap">
                                <a href="index.html"><img src="https://raw.githubusercontent.com/HetRod/prueba/refs/heads/master/logo.png" alt="logo" /></a>
                            </div>
                        </div>
        
                        <div class="badge-tracker-item col-lg-9">
                            <div class="title-top">
                                <h3>Porcentaje de Validación del Sello</h3>
                            </div>
        
                            <!-- falta la clase headline-round-content y worldwide-stats -->
                            <div class="headline-round-content mtn-25 worldwide-stats">
                                <div class="tracker-block tracker-block--5 bg-yellow">
                                    <h3 class="track-item__no infected">50%</h3>            
                                </div>
                            </div>
                        </div>
        
                        <div class="col-sm-5 col-lg-3">
                            <div class="last-update-wrap">
                                <p class="mb-0">Fecha: <span id="time" class="last-update"></span></p>
                            </div>
                        </div>
                    </div>
                </div>
            </header>
        
            <!--== Start Main Content Wrapper ==-->
            <main class="main-content-wrapper">
                <div class="badge-tracker-wrap">
                    <div class="container">
                        <div class="row">
                            <div class="col-xl-9 m-auto">
                                <div class="badge-tracker-item">
                                    <div class="title-top">
                                        <h2>Reporte del Sello</h2>
                                    </div>
        
                                    <div class="badge-content-wrapper" id="badgeContainer">
                                        <!-- Sección de badges generados dinámicamente -->
                                    </div>
                                    
                                    <!-- Sección de paginación -->
                                    <div class="pagination" id="paginationContainer">
                                        <button id="prevPage" onclick="changePage(-1)">Anterior</button>
                                        <span id="pageInfo">Page 1 of 4</span>
                                        <button id="nextPage" onclick="changePage(1)">Siguiente</button>
                                    </div>
                                    
                                    <script>
                                        const data = [ """ + armar_html(resultados_totales) + """
            
        
                                        ];
                                    
                                        const itemsPerPage = 8; // Cambia este número para ajustar la cantidad de filas por página
                                        let currentPage = 1;
                                    
                                        function renderBadges() {
                                            const container = document.getElementById('badgeContainer');
                                            container.innerHTML = ''; // Limpiar contenido previo
                                            const start = (currentPage - 1) * itemsPerPage;
                                            const end = start + itemsPerPage;
                                    
                                            data.slice(start, end).forEach(item => {
                                                const badge = document.createElement('div');
                                                badge.className = `badge-item worldwide-stats badge-bg-green`;
                                                badge.innerHTML = `
                                                    <div class="badge-item__title">
                                                    <a href="/ruta_comprobante_1" target="_blank"> <p>${item.comprobante}</p></a>
                                                    </div>
                                                    <div class="badge-item__body">
                                                        <div class="badge-item__block">
                                                            <p>Factura: <br><i class="fa ${item.factura === '1' ? 'fa-check-circle-o' : 'fa-times-circle-o'} fa-2x" aria-hidden="true" style="color: ${item.factura === '1' ? 'green' : 'red'};"></i> </p>
                                                        </div>
                                                        <div class="badge-item__block">
                                                            <p>Fecha: <br><i class="fa ${item.fecha === '1' ? 'fa-check-circle-o' : 'fa-times-circle-o'} fa-2x" aria-hidden="true" style="color: ${item.fecha === '1' ? 'green' : 'red'};"></i> </p>
                                                        </div>
        
                                                        <div class="badge-item__block">
                                                            <p>Valor: <br><i class="fa ${item.valor === '1' ? 'fa-check-circle-o' : 'fa-times-circle-o'} fa-2x" aria-hidden="true" style="color: ${item.valor === '1' ? 'green' : 'red'};"></i> </p>
                                                        </div>
                                                    </div>
                                                `;
                                                container.appendChild(badge);
                                            });
                                    
                                            document.getElementById('pageInfo').innerText = `Page ${currentPage} of ${Math.ceil(data.length / itemsPerPage)}`;
                                            updatePaginationButtons();
                                        }
                                    
                                        function changePage(direction) {
                                            const totalPages = Math.ceil(data.length / itemsPerPage);
                                            currentPage += direction;
                                            if (currentPage < 1) currentPage = 1;
                                            if (currentPage > totalPages) currentPage = totalPages;
                                            renderBadges();
                                        }
                                    
                                        function updatePaginationButtons() {
                                            document.getElementById('prevPage').disabled = currentPage === 1;
                                            document.getElementById('nextPage').disabled = currentPage === Math.ceil(data.length / itemsPerPage);
                                        }
                                    
                                        // Inicializar renderizado
                                        renderBadges();
                                    </script>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </main>
            <!--== End Main Content Wrapper ==-->
        
            <script>
                const d = new Date();
                // Formatear la fecha
                const formattedDate = d.toLocaleDateString('es-ES', {
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric'
                });
                document.getElementById("time").textContent = formattedDate;
            </script>
        
        </body>
        </html>
        """
        
        crear_archivo_html('index.html', contenido_html)