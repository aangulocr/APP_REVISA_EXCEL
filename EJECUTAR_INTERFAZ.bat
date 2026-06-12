@echo off
chcp 65001 >nul 2>&1
title Auditor de Excel - Interfaz Gráfica de Usuario

echo ==========================================================
echo   AUDITOR DE EXCEL - PANEL DE CONTROL GRÁFICO
echo ==========================================================
echo.

:: Cambiar al directorio donde esta el .bat
cd /d "%~dp0"

:: Verificar que Python este instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo         Instalalo desde https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: Configurar entorno virtual si no existe
set VENV_DIR=.venv
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Creando entorno virtual en %VENV_DIR%...
    python -m venv %VENV_DIR%
)

:: Activar entorno virtual
call %VENV_DIR%\Scripts\activate.bat

:: Asegurar que pip este disponible
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Instalando pip en el entorno virtual...
    powershell -Command "Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -OutFile get-pip.py"
    python get-pip.py --quiet
    del get-pip.py
)

:: Instalar dependencias si existe requirements.txt
if exist "requirements.txt" (
    echo Verificando/instalando dependencias...
    python -m pip install -q -r requirements.txt
)

:: Verificar que gui_server.py exista
if not exist "gui_server.py" (
    echo [ERROR] No se encontro el archivo gui_server.py.
    echo         Asegurate de que este archivo .bat este en la misma
    echo         carpeta que gui_server.py
    echo.
    pause
    exit /b 1
)

echo Iniciando servidor de interfaz grafica...
echo.

python gui_server.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Ocurrió un problema al ejecutar la interfaz gráfica.
    echo.
    pause
)
