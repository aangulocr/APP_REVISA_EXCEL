#!/bin/bash

# Cambiar al directorio donde está el script
cd "$(dirname "$0")" || exit

echo "=========================================================="
echo "  AUDITOR DE EXCEL - PANEL DE CONTROL GRÁFICO"
echo "=========================================================="
echo ""

# Verificar qué comando de python usar (python3 o python)
if command -v python3 >/dev/null 2>&1; then
    PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_CMD="python"
else
    echo "[ERROR] Python no está instalado o no está en el PATH."
    echo "        Instálalo desde https://www.python.org/downloads/"
    echo ""
    read -p "Presiona Enter para continuar..."
    exit 1
fi

# Configurar entorno virtual si no existe o está roto
VENV_DIR=".venv"
if [ ! -d "$VENV_DIR" ] || [ ! -f "$VENV_DIR/bin/python" ]; then
    echo "Creando entorno virtual en $VENV_DIR..."
    rm -rf "$VENV_DIR"
    $PYTHON_CMD -m venv "$VENV_DIR"
fi

# Activar entorno virtual
if [ -f "$VENV_DIR/bin/activate" ]; then
    source "$VENV_DIR/bin/activate"
else
    echo "[ERROR] No se pudo activar el entorno virtual."
    read -p "Presiona Enter para continuar..."
    exit 1
fi

# Instalar dependencias si existe requirements.txt
if [ -f "requirements.txt" ]; then
    echo "Verificando dependencias..."
    python -m pip install --upgrade pip -q
    python -m pip install -q -r requirements.txt
fi

# Verificar que gui_server.py exista
if [ ! -f "gui_server.py" ]; then
    echo "[ERROR] No se encontró el archivo gui_server.py"
    echo "        Asegúrate de que este archivo .sh esté en la misma"
    echo "        carpeta que gui_server.py"
    echo ""
    read -p "Presiona Enter para continuar..."
    exit 1
fi

echo "Iniciando servidor de interfaz gráfica..."
echo ""

python gui_server.py
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo "[ERROR] Ocurrió un problema al ejecutar la interfaz gráfica."
    read -p "Presiona Enter para continuar..."
fi

deactivate
