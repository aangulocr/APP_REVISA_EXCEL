#!/bin/bash

# Cambiar al directorio donde está el script
cd "$(dirname "$0")" || exit

echo "=========================================================="
echo "  AUDITOR DE EXCEL - Comparacion de Fidelidad Total"
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

# Asegurar que pip esté instalado y funcione
echo "Verificando/instalando dependencias..."
if ! python -m pip --version >/dev/null 2>&1; then
    echo "Pip no encontrado en el entorno virtual. Intentando instalarlo..."
    curl -sS https://bootstrap.pypa.io/get-pip.py -o get-pip.py
    python get-pip.py --quiet
    rm get-pip.py
fi
python -m pip install --upgrade pip -q

# Instalar dependencias si existe requirements.txt
if [ -f "requirements.txt" ]; then
    python -m pip install -q -r requirements.txt
fi

# Verificar que main.py exista
if [ ! -f "main.py" ]; then
    echo "[ERROR] No se encontró el archivo main.py"
    echo "        Asegúrate de que este archivo .sh esté en la misma"
    echo "        carpeta que main.py"
    echo ""
    read -p "Presiona Enter para continuar..."
    exit 1
fi

echo "Iniciando auditoría..."
echo ""

python main.py
EXIT_CODE=$?

echo ""
echo "=========================================================="
if [ $EXIT_CODE -eq 0 ]; then
    echo "  Auditoría completada exitosamente."
else
    echo "  La auditoría finalizó con errores (código: $EXIT_CODE)."
fi
echo "=========================================================="
echo ""

deactivate

# Pausa final similar a "pause" en Windows
read -p "Presiona Enter para salir..."
