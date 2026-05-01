# ============================================================================
# config.py — Configuración central del auditor de Excel
# ============================================================================
"""
Contiene todas las constantes y rutas que el script necesita.
Modificar aquí para adaptar a distintos entornos.
"""

import os

# --- Rutas principales ---
# Directorio base del proyecto (donde vive este archivo)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Archivo maestro (plantilla de referencia)
PLANTILLA_PATH = os.path.join(BASE_DIR, "PLANTILLA.xlsx")

# Carpeta con los trabajos de los estudiantes
TRABAJOS_DIR = os.path.join(BASE_DIR, "TRABAJOS_ESTUDIANTES")

# Carpeta de salida para los archivos revisados
REVISADOS_DIR = os.path.join(BASE_DIR, "REVISADOS")

# Archivo CSV de resumen
LOG_NOTAS_PATH = os.path.join(BASE_DIR, "LOG_NOTAS.csv")

# --- Colores de retroalimentación ---
# Rojo claro para celdas con errores (ARGB hex)
COLOR_ERROR_FILL = "FFFFC7CE"       # Rojo claro de fondo
COLOR_ERROR_FONT = "FF9C0006"       # Rojo oscuro de fuente (opcional)

# --- Extensiones válidas ---
EXTENSIONES_VALIDAS = (".xlsx",)

# --- Mensajes de error estándar ---
MENSAJES = {
    "valor":               "❌ Valor incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "formula":             "❌ Fórmula incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "funcion":             "❌ Función utilizada incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "color_fuente":        "❌ Color de fuente incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "color_relleno":       "❌ Color de relleno incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "borde":               "❌ Estilo de borde incorrecto en lado '{lado}'. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "alineacion_h":        "❌ Alineación horizontal incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "alineacion_v":        "❌ Alineación vertical incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "wrap_text":           "❌ Ajuste de texto (wrap) incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "numero_formato":      "❌ Formato numérico incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "fuente_nombre":       "❌ Nombre de fuente incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "fuente_tamano":       "❌ Tamaño de fuente incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "fuente_negrita":      "❌ Negrita incorrecta. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "fuente_cursiva":      "❌ Cursiva incorrecta. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "fuente_subrayado":    "❌ Subrayado incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "tabla_faltante":      "❌ Tabla de Excel faltante: '{nombre}'",
    "tabla_rango":         "❌ Rango de tabla incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "tabla_estilo":        "❌ Estilo de tabla incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "validacion_faltante": "❌ Validación de datos faltante en rango '{rango}'",
    "validacion_config":   "❌ Validación de datos incorrecta. Esperada: '{esperado}' | Encontrada: '{encontrado}'",
    "formato_cond":        "❌ Regla de formato condicional faltante o incorrecta en rango '{rango}'",
    "hoja_faltante":       "❌ Hoja faltante: '{nombre}'",
    "merge_faltante":      "❌ Rango combinado faltante: '{rango}'",
    "merge_extra":         "❌ Rango combinado extra (no está en plantilla): '{rango}'",
    "ancho_columna":       "❌ Ancho de columna incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
    "alto_fila":           "❌ Alto de fila incorrecto. Esperado: '{esperado}' | Encontrado: '{encontrado}'",
}
