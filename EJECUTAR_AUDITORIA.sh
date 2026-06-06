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

# Verificar que main.py exista
if [ ! -f "main.py" ]; then
    echo "[ERROR] No se encontró el archivo main.py"
    echo "        Asegúrate de que este archivo .sh esté en la misma"
    echo "        carpeta que main.py"
    echo ""
    read -p "Presiona Enter para continuar..."
    exit 1
fi

echo "Iniciando auditoría con $PYTHON_CMD..."
echo ""

$PYTHON_CMD main.py
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

# Pausa final similar a "pause" en Windows
read -p "Presiona Enter para salir..."
