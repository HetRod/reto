import pandas as pd
import sys
import pdfplumber
import re
import os
import openpyxl


RUTA = "/Users/jmcardor/Library/CloudStorage/OneDrive-SharedLibraries-GrupoBancolombia/Prueba Bintec - Documentos/Bintec_HAMJ/Sello/sello 35-1/"

def crear_archivo_html(nombre_archivo, contenido):
    with open(nombre_archivo, 'w') as archivo:
        archivo.write(contenido) 

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
            nro_factura = None
            for patron_factura in patrones['patrones_factura']:
                match = patron_factura.search(texto)
                if match:
                    nro_factura = match.group(1)  # Guardar el primer valor encontrado
                    break  # Salir del bucle una vez que encontremos el número de factura
            
            if nro_factura:
                resultados['nro_factura'] = nro_factura

            # Buscar otros valores como valor pagado
            valor_pagado = patrones['patron_valor'].search(texto)
            if valor_pagado:
                resultados['valor_pagado'] = procesar_valor_pagado(valor_pagado.group(2))

            # Buscar fecha y hora de la transacción
            fecha_hora_transaccion = patrones['patron_fecha_hora'].search(texto)
            if fecha_hora_transaccion:
                fecha_hora_str = fecha_hora_transaccion.group(0)
                dia, mes, anio = re.search(r'(\d{1,2})\s+de\s+([A-Za-z]+)\s+de\s+(\d{4})', fecha_hora_str).groups()
                mes_num = patrones['meses'][mes.capitalize()]  # Convertir el mes a número
                resultados['fecha_hora_transaccion'] = f"{anio}{mes_num}{dia.zfill(2)}"

            # Buscar fecha de envío de pago
            fecha_envio_pago = patrones['patron_fecha_envio'].search(texto)
            if fecha_envio_pago:
                fecha_envio_str = fecha_envio_pago.group(1)
                dia, mes, anio = fecha_envio_str.split('-')
                resultados['fecha_envio_pago'] = f"{anio}{mes}{dia}"

    return resultados

def procesar_pdf_con_parametros(nombre_pdf, num_factura, valor_pagado, fecha_envio_pago):
    ruta_pdf = RUTA + "soportes/" + nombre_pdf + ".pdf"


    patrones = {
        'patrones_factura': [
            re.compile(r'Factura\s*No\s*:\s*(\d+)'),  
            re.compile(r'(?:FACTURA.*|SERVICIOS.*|Factura electrónicadeventa.*)?No\.?\s*(\d{6,}-\d+|\d{6,})'), 
            re.compile(r'(?:Factura\s+elect\.\s+de\s+venta:?\s*|Factura\s+electrónicadesventa:?\s*)(\d{6,})'), 
            re.compile(r'CUENTA\s+DE\s+COBRO\s+No:\s*-\s*(\d+)'),
            re.compile(r'Nro\.?de\s*factura:?\.?\s*(\d{13,})'), 
            re.compile(r'No\.?\s*de\s*factura\s*(\d+)'), 
            re.compile(r'Nro\.?Doc\.?:?\s*([A-Z]+\d+)'), 
            re.compile(r'Factura\s*No\s*:\s*(\d+)'),  
            re.compile(r'No\.?\s*([A-Z]+\d+)'),
            re.compile(r'Factura\s+Electronica\s+De\s+Venta\s+No\s+FVEO\s+No\.?\s*(\d+)'),
            re.compile(r'Cuenta\s+de\s+Cobro\s+(\d+)'),
            re.compile(r'Cuenta\s+de\s+Cobro:\s*CA\s*(\d+)'),
            re.compile(r'FE\s*No\.\s*(\d+)'),
            re.compile(r'CUENTA\s*DE\s*COBRO\s*(\w+)'),
            re.compile(r'No\.?\s*-\s*(\d+)'),
            re.compile(r'No\.?\s*ZF\s*-\s*(\d+)'),
            re.compile(r'FACTURA\s*ELECTRONICA\s*DE\s*VENTA\s*(\w+)'),
            re.compile(r'FE\s*(\d+)'),
            re.compile(r'Mo\.\s*(\w+\d+)'),
            re.compile(r'CUENTA\s+DE\s+COBRO\s*No\.?\s*(\d+)'),
            re.compile(r'Nro\.?\s*Doc\.?:?\s*([A-Z]+\d+)'),
            re.compile(r'Cupon\s*(\d+)')
        ],
        'patron_valor': re.compile(r'(Valor Total del Pago|total a pagar|Total a Pagar|pago total|Valor pagado)\s*:?\s*\$?\s*([\d{1,3}(?:,\d{3})*(?:\.\d{2})?]+)'),
        'patron_fecha_hora': re.compile(r'\d{1,2}\s+de\s+(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+de\s+\d{4}'),
        'patron_fecha_envio': re.compile(r'Fecha de envío del pago\s*:\s*(\d{2}-\d{2}-\d{4})'),
        'meses': {
            'Enero': '01', 'Febrero': '02', 'Marzo': '03', 'Abril': '04',
            'Mayo': '05', 'Junio': '06', 'Julio': '07', 'Agosto': '08',
            'Septiembre': '09', 'Octubre': '10', 'Noviembre': '11', 'Diciembre': '12'
        }
    }

    if os.path.isfile(ruta_pdf):
        resultados_pdf = extraer_datos_pdf(ruta_pdf, patrones)
    else:
        return [nombre_pdf, 0, 0, 0]
    
    coincidencias = []
    coincidencias.append(nombre_pdf)
    coincidencias.append(1 if str(resultados_pdf.get('nro_factura')) == str(num_factura).strip() else 0)
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


def procesar_archivo_excel(nombre_archivo):
    try:
        # Leer el archivo de Excel
        df = pd.read_excel(nombre_archivo, sheet_name="SELLO")

        df.columns = df.columns.str.replace('\n', '')
        # Organizar el DataFrame por 'NUMERO DE COMPROBANTE'
        df = df.sort_values(by='NUMERO DE COMPROBANTE')
        # Eliminar filas donde 'NUMERO DE COMPROBANTE' sea vacío
        df = df[df['NUMERO DE COMPROBANTE'].notna()]
        arreglo_valores =[]
        # Recorrer el DataFrame e imprimir la columna H
        comprobante_anterior = ""
        rows = []
        for index, row in df.iterrows(): 
            if comprobante_anterior == row['NUMERO DE COMPROBANTE']:
                arreglo_valores[-1][4].append(index + 2)
                arreglo_valores[-1][3] += row['TOTAL'] 
            else:
                rows.append(index + 2)
                
                valor = [row['NUMERO DE COMPROBANTE'],row['# FACTURA PAGADA'],int(row['F_PAGO AAAAMMDD']),row['TOTAL'] ,rows]
                arreglo_valores.append(valor)
            rows = []
            comprobante_anterior = row['NUMERO DE COMPROBANTE']
        print(arreglo_valores)
        return arreglo_valores
    except Exception as e:
        print(f"Error al procesar el archivo: {e}")
        return None

def mapear_observacion(resultado):
    if resultado[-3:] == [0, 0, 0]:
        return 'Ninguna coincidencia'
    elif resultado[-3:] == [1, 0, 0]:
        return 'No coincide el valor pagado ni la fecha de envío'
    elif resultado[-3:] == [0, 1, 0]:
        return 'No coincide el número de factura ni la fecha de envío'
    elif resultado[-3:] == [0, 0, 1]:
        return 'No coincide el número de factura ni el valor pagado'
    elif resultado[-3:] == [1, 1, 0]:
        return 'No coincide la fecha de envío'
    elif resultado[-3:] == [1, 0, 1]:
        return 'No coincide el valor pagado'
    elif resultado[-3:] == [0, 1, 1]:
        return 'No coincide el número de factura'
    elif resultado[-3:] == [1, 1, 1]:
        return ''



if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python crear_archivo.py <nombre_archivo>")
    else:
        archivo_excel = sys.argv[1]
        arreglo_valores = procesar_archivo_excel(RUTA + archivo_excel)
        print(arreglo_valores)
        data_html = ""
        wb = openpyxl.load_workbook(RUTA + archivo_excel, keep_vba=True, data_only=True)
        ws = wb['SELLO']
        porcentaje_x_item = 100 / len(arreglo_valores)/3
        porcentaje_validado = 0
        for valor in arreglo_valores:
            resultado_validaciones = procesar_pdf_con_parametros(valor[0],valor[1],valor[3],valor[2])
            porcentaje_validado += (porcentaje_x_item * resultado_validaciones[1]) + (porcentaje_x_item * resultado_validaciones[2]) + (porcentaje_x_item * resultado_validaciones[3])
            print(f"Porcentaje validado: {porcentaje_validado}")
            data_html += f'{{ comprobante: "{resultado_validaciones[0]}", factura:"{resultado_validaciones[1]}", fecha:"{resultado_validaciones[3]}", valor:"{resultado_validaciones[2]}" }},\n'
            for row in valor[4]:
                ws['AJ' + str(row)] = mapear_observacion(resultado_validaciones)
        wb.save(RUTA + archivo_excel)

        porcentaje_str = "{:.2f}".format(porcentaje_validado)
        if data_html.endswith(",\n"):
            data_html = data_html[:-2] + "\n"

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
                                    <h3 class="track-item__no infected">""" + porcentaje_str + """</h3>            
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
                                        const data = [ """ + data_html + """
            
        
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
                                                    <a href=" """ + RUTA + """soportes""" + """/${item.comprobante}""" + """.pdf """ + """"  target="_blank"> <p>${item.comprobante}</p></a>
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

        reporte = RUTA + 'Reporte.html'
        crear_archivo_html(reporte, contenido_html)