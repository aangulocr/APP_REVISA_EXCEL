# ============================================================================
# comparador_celdas.py — Comparación celda a celda (valor, fórmula, formato)
# ============================================================================
"""
Módulo encargado de la comparación granular entre celdas de la plantilla
y celdas del estudiante. Detecta diferencias en:
  - Valor resultante
  - Fórmula exacta
  - Funciones utilizadas
  - Color de fuente (RGB)
  - Color de relleno
  - Estilos de borde (estilo y grosor por cada lado)
  - Alineación (horizontal, vertical, wrap_text)
  - Formato numérico
  - Propiedades de fuente (nombre, tamaño, negrita, cursiva, subrayado)
"""

import re
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from config import COLOR_ERROR_FILL, MENSAJES


# Relleno rojo claro para marcar errores
RELLENO_ERROR = PatternFill(
    start_color=COLOR_ERROR_FILL,
    end_color=COLOR_ERROR_FILL,
    fill_type="solid"
)


def _normalizar_valor(valor):
    """Normaliza un valor para comparación (elimina espacios extra, etc.)."""
    if valor is None:
        return None
    if isinstance(valor, str):
        return valor.strip()
    return valor


# Diccionario de traducción: funciones en español → inglés (formato .xlsx interno)
# El formato .xlsx SIEMPRE almacena funciones en inglés internamente.
# Excel las muestra traducidas según el idioma del usuario, pero openpyxl
# las lee/escribe en inglés.
FUNCIONES_ES_A_EN = {
    "SUMA": "SUM",
    "PROMEDIO": "AVERAGE",
    "CONTAR": "COUNT",
    "CONTARA": "COUNTA",
    "CONTAR.SI": "COUNTIF",
    "CONTAR.SI.CONJUNTO": "COUNTIFS",
    "SUMAR.SI": "SUMIF",
    "SUMAR.SI.CONJUNTO": "SUMIFS",
    "BUSCARV": "VLOOKUP",
    "BUSCARH": "HLOOKUP",
    "SI": "IF",
    "SI.ERROR": "IFERROR",
    "Y": "AND",
    "O": "OR",
    "NO": "NOT",
    "VERDADERO": "TRUE",
    "FALSO": "FALSE",
    "CONCATENAR": "CONCATENATE",
    "TEXTO": "TEXT",
    "HOY": "TODAY",
    "AHORA": "NOW",
    "AÑO": "YEAR",
    "MES": "MONTH",
    "DIA": "DAY",
    "REDONDEAR": "ROUND",
    "ENTERO": "INT",
    "POTENCIA": "POWER",
    "RAIZ": "SQRT",
    "ABS": "ABS",
    "MAX": "MAX",
    "MIN": "MIN",
    "MAYUSC": "UPPER",
    "MINUSC": "LOWER",
    "LARGO": "LEN",
    "IZQUIERDA": "LEFT",
    "DERECHA": "RIGHT",
    "EXTRAE": "MID",
    "ENCONTRAR": "FIND",
    "SUSTITUIR": "SUBSTITUTE",
    "INDICE": "INDEX",
    "COINCIDIR": "MATCH",
    "DESREF": "OFFSET",
    "FILA": "ROW",
    "COLUMNA": "COLUMN",
    "TRANSPONER": "TRANSPOSE",
    "ELEGIR": "CHOOSE",
    "ALEATORIO": "RAND",
    "RESIDUO": "MOD",
    "PRODUCTO": "PRODUCT",
    "PROMEDIO.SI": "AVERAGEIF",
    "PROMEDIO.SI.CONJUNTO": "AVERAGEIFS",
    "K.ESIMO.MAYOR": "LARGE",
    "K.ESIMO.MENOR": "SMALL",
    "ORDENAR": "SORT",
    "FILTRAR": "FILTER",
    "UNICO": "UNIQUE",
    "TIPO": "TYPE",
    "ESERROR": "ISERROR",
    "ESNUMERO": "ISNUMBER",
    "ESTEXTO": "ISTEXT",
}

# Crear diccionario inverso: inglés → español
FUNCIONES_EN_A_ES = {v: k for k, v in FUNCIONES_ES_A_EN.items()}


def _normalizar_formula(formula):
    """
    Normaliza una fórmula de Excel para comparación justa.
    
    Maneja:
    - Prefijos internos de Excel: _xlfn., _xlfn._xlws.
    - Operador de intersección implícita: @
    - Mayúsculas/minúsculas
    - Espacios extra
    
    Ejemplo:
        '=_xlfn._xlws.SUM(A1:A10)' → '=SUM(A1:A10)'
        '=@AVERAGE(B1:B5)' → '=AVERAGE(B1:B5)'
    """
    if not formula or not isinstance(formula, str):
        return formula
    
    f = formula.strip().upper()
    
    # Eliminar prefijos internos de Excel que openpyxl puede leer
    f = f.replace("_XLFN._XLWS.", "")
    f = f.replace("_XLFN.", "")
    
    # Eliminar el operador @ de intersección implícita
    # (aparece después del = en Excel 365+)
    if f.startswith("=@"):
        f = "=" + f[2:]
    # También puede aparecer dentro de la fórmula
    f = f.replace("(@", "(")
    
    return f


def _extraer_funciones(formula):
    """
    Extrae las funciones de Excel usadas en una fórmula.
    Normaliza primero para eliminar prefijos _xlfn.
    Ejemplo: '=SUM(A1:A10)+AVERAGE(B1:B5)' -> ['AVERAGE', 'SUM']
    """
    if not formula or not isinstance(formula, str):
        return []
    # Normalizar primero
    formula_limpia = _normalizar_formula(formula)
    # Buscar patrones de funciones: NOMBRE_FUNCION(
    patron = r'([A-Za-záéíóúñÁÉÍÓÚÑ_][A-Za-z0-9áéíóúñÁÉÍÓÚÑ_.]*)\s*\('
    funciones = re.findall(patron, formula_limpia)
    return sorted(set(f.upper() for f in funciones))


def _obtener_color_rgb(color_obj):
    """
    Obtiene el color RGB de un objeto Color de openpyxl.
    Retorna el string RGB o 'Sin color' si no está definido.
    """
    if color_obj is None:
        return "Sin color"
    # Si tiene un tema, retornar el índice del tema
    if color_obj.type == "theme":
        return f"Tema:{color_obj.theme}+Tint:{color_obj.tint}"
    # Si tiene RGB definido
    if color_obj.rgb and color_obj.rgb != "00000000":
        return str(color_obj.rgb)
    # Si tiene índice
    if color_obj.indexed is not None:
        return f"Indexed:{color_obj.indexed}"
    return "Sin color"


def _obtener_color_relleno(fill):
    """Obtiene el color de relleno de una celda."""
    if fill is None or fill.fill_type is None:
        return "Sin relleno"
    if fill.fgColor:
        return _obtener_color_rgb(fill.fgColor)
    return "Sin relleno"


def _comparar_bordes(borde_plantilla, borde_estudiante):
    """
    Compara los bordes de dos celdas (left, right, top, bottom, diagonal).
    Retorna una lista de mensajes de error.
    """
    errores = []
    lados = {
        "izquierdo": ("left", borde_plantilla.left, borde_estudiante.left),
        "derecho": ("right", borde_plantilla.right, borde_estudiante.right),
        "superior": ("top", borde_plantilla.top, borde_estudiante.top),
        "inferior": ("bottom", borde_plantilla.bottom, borde_estudiante.bottom),
        "diagonal": ("diagonal", borde_plantilla.diagonal, borde_estudiante.diagonal),
    }
    for nombre_lado, (_, lado_p, lado_e) in lados.items():
        estilo_p = lado_p.style if lado_p else None
        estilo_e = lado_e.style if lado_e else None
        if estilo_p != estilo_e:
            errores.append(
                MENSAJES["borde"].format(
                    lado=nombre_lado,
                    esperado=estilo_p or "ninguno",
                    encontrado=estilo_e or "ninguno"
                )
            )
    return errores


def comparar_celda(celda_plantilla, celda_estudiante, ws_plantilla_data=None):
    """
    Compara una celda de la plantilla con su homóloga del estudiante.
    MODO: Solo fórmulas (ignora formato y valores constantes).

    Retorna:
        tuple: (es_correcto: bool, lista_errores: list[str])
    """
    errores = []

    # ----------------------------------------------------------------
    # 1. COMPARACIÓN EXCLUSIVA DE FÓRMULAS
    # ----------------------------------------------------------------
    valor_p = celda_plantilla.value
    valor_e = celda_estudiante.value

    # Solo auditamos si la PLANTILLA tiene una fórmula.
    # Si la plantilla tiene un valor constante, ignoramos la celda.
    es_formula_p = isinstance(valor_p, str) and valor_p.startswith("=")
    
    if es_formula_p:
        es_formula_e = isinstance(valor_e, str) and valor_e.startswith("=")
        
        # Comparar la fórmula normalizada (sin _xlfn., @, etc.)
        formula_p_raw = valor_p.strip()
        formula_e_raw = valor_e.strip() if es_formula_e else str(valor_e) if valor_e is not None else ""
        formula_p_norm = _normalizar_formula(formula_p_raw)
        formula_e_norm = _normalizar_formula(formula_e_raw)
        
        if formula_p_norm != formula_e_norm:
            errores.append(
                MENSAJES["formula"].format(
                    esperado=formula_p_raw,
                    encontrado=formula_e_raw or "(vacío)"
                )
            )
        
        # Comparar funciones utilizadas
        func_p = _extraer_funciones(formula_p_raw)
        func_e = _extraer_funciones(formula_e_raw)
        if func_p != func_e:
            errores.append(
                MENSAJES["funcion"].format(
                    esperado=", ".join(func_p) if func_p else "(ninguna)",
                    encontrado=", ".join(func_e) if func_e else "(ninguna)"
                )
            )
    else:
        # Si no es fórmula, la damos por correcta automáticamente (no resta puntos)
        return True, []

    # ----------------------------------------------------------------
    # RESULTADO
    # ----------------------------------------------------------------
    es_correcto = len(errores) == 0
    return es_correcto, errores


def marcar_celda_con_error(celda, errores):
    """
    Marca una celda del estudiante como errónea:
      - Cambia el color de fondo a rojo claro.
      - Inserta un comentario con la lista de errores encontrados.

    Parámetros:
        celda: Celda del libro del estudiante.
        errores: Lista de strings con los mensajes de error.
    """
    # Aplicar relleno rojo claro
    celda.fill = RELLENO_ERROR

    # Crear comentario con todos los errores
    texto_comentario = "\n".join(errores)
    celda.comment = Comment(
        text=texto_comentario,
        author="Auditor Excel"
    )
    # Ajustar tamaño del comentario para que sea legible
    celda.comment.width = 400
    celda.comment.height = 150 + (len(errores) * 30)
