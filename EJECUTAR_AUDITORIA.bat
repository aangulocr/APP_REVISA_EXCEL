@echo off
chcp 65001 >nul 2>&1
title Auditor de Excel - Revision de Trabajos

echo ==========================================================
echo   AUDITOR DE EXCEL - Comparacion de Fidelidad Total
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

:: Verificar que main.py exista
if not exist "main.py" (
    echo [ERROR] No se encontro el archivo main.py
    echo         Asegurate de que este archivo .bat este en la misma
    echo         carpeta que main.py
    echo.
    pause
    exit /b 1
)

echo Iniciando auditoria...
echo.

python main.py

echo.
echo ==========================================================
if %errorlevel% equ 0 (
    echo   Auditoria completada exitosamente.
) else (
    echo   La auditoria finalizo con errores (codigo: %errorlevel%).
)
echo ==========================================================
echo.
pause
