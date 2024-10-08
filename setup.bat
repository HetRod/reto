@echo off

:: Verificar si Python está instalado
echo Verificando si Python está instalado...
where python
if %ERRORLEVEL% NEQ 0 (
    echo Python no está instalado. Procediendo a instalar Python...

    :: Instalar Python usando winget
    winget install Python.Python.3
    if %ERRORLEVEL% NEQ 0 (
        echo Error: No se pudo instalar Python. Por favor, instala Python manualmente.
        exit /b 1
    )
) else (
    echo Python ya está instalado.
)

echo Verificando instalación de pip...

:: Verificar si pip está instalado
where pip
if %ERRORLEVEL% NEQ 0 (
    echo pip no está instalado. Procediendo a instalar pip...

    :: Instalar pip en Windows
    python -m ensurepip --upgrade
    python -m pip install --upgrade pip
) else (
    echo pip ya está instalado.
)

echo Instalando librerías necesarias...

:: Instalar las librerías usando pip
pip install pandas
pip install pdfplumber

echo Instalación de librerías completada.

:: Ejecutar el archivo read.py
echo Ejecutando el archivo read.py...
"C:\Program Files\Python312\python.exe" "C:\Users\hmrodrig\Downloads\read.py"

:: Pausar para que puedas ver si el script fue ejecutado correctamente
echo.
echo Ejecución completada. Esperando 10 segundos antes de cerrar...
timeout /t 10

:: Redirigir la salida a un archivo de log
"C:\Program Files\Python312\python.exe" "C:\Users\hmrodrig\Downloads\read.py" > output.log 2>&1

