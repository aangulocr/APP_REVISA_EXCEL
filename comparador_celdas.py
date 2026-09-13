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
import openpyxl.utils
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


# Funciones de agregación donde los argumentos son conmutativos
# y donde un rango (ej: J3:J8) equivale a la lista de sus celdas individuales
FUNCIONES_AGREGACION_CONMUTATIVAS = {
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "MEDIAN"
}


def _expandir_argumentos(args_str):
    """
    Toma un string de argumentos de una función, permite comas o punto y coma,
    y expande cualquier rango (ej. 'J3:J8' -> ['J3', 'J4', 'J5', 'J6', 'J7', 'J8']).
    Elimina signos '$' para comparación lógica neutral.
    Retorna la lista ordenada de celdas/términos expandidos.
    """
    if not args_str:
        return []
    # Normalizar separador ; a ,
    norm_str = args_str.replace(";", ",")
    partes = [p.strip() for p in norm_str.split(",") if p.strip()]
    expandidos = []
    for parte in partes:
        p_clean = parte.replace("$", "").strip()
        if ":" in p_clean:
            try:
                min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(p_clean)
                if min_col and min_row and max_col and max_row:
                    for c in range(min_col, max_col + 1):
                        col_let = openpyxl.utils.get_column_letter(c)
                        for r in range(min_row, max_row + 1):
                            expandidos.append(f"{col_let}{r}")
                    continue
            except Exception:
                pass
        expandidos.append(p_clean)
    return sorted(expandidos)


def _analizar_formula_agregacion(formula):
    """
    Determina si una fórmula es una llamada a función de agregación conmutativa.
    Ejemplo: '=AVERAGE(H5:I5)' -> ('AVERAGE', ['H5', 'I5'])
             '=SUMA(J3;J4;J5)' -> ('SUM', ['J3', 'J4', 'J5'])
    Retorna (nombre_funcion_en_ingles, lista_ordenada_de_celdas) o (None, []).
    """
    if not formula or not isinstance(formula, str):
        return None, []
    f = formula.strip().upper()
    m = re.match(r"^=([A-Z0-9_\.]+)\s*\((.*)\)$", f)
    if not m:
        return None, []
    func_raw = m.group(1).strip()
    args_str = m.group(2).strip()

    # Si hay paréntesis dentro de los argumentos (funciones anidadas), no aplanar
    if "(" in args_str or ")" in args_str:
        return None, []

    # Traducir función de español a inglés si aplica
    func_en = FUNCIONES_ES_A_EN.get(func_raw, func_raw)
    if func_en in FUNCIONES_AGREGACION_CONMUTATIVAS:
        celdas = _expandir_argumentos(args_str)
        return func_en, celdas

    return None, []


def _analizar_operacion_cadena(formula, operador):
    """
    Extrae los operandos de una cadena simple con un único operador (+ o *).
    Ejemplo: '=J3+J4+J5' con operador '+' -> ['J3', 'J4', 'J5']
             '=D5+E5' con operador '+' -> ['D5', 'E5']
    Elimina signos '$' y retorna la lista ordenada.
    """
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return []
    cuerpo = formula[1:].strip().upper()
    if "(" in cuerpo or ")" in cuerpo:
        return []

    # Verificar que no contenga operadores no deseados
    otros_ops = ["-", "/", "*"] if operador == "+" else ["-", "/", "+"]
    if any(op in cuerpo for op in otros_ops):
        return []

    if operador not in cuerpo:
        return []

    partes = [p.strip().replace("$", "") for p in cuerpo.split(operador) if p.strip()]
    return sorted(partes)


def _formulas_conmutativas_equivalentes(f1_norm, f2_norm):
    """
    Verifica si dos fórmulas normalizadas son equivalentes bajo
    conmutatividad de suma (+) o multiplicación (*).
    Conserva compatibilidad para multiplicaciones con porcentajes (ej. =C3*3% vs =3%*C3).
    """
    if not f1_norm or not f2_norm:
        return False
    if not f1_norm.startswith("=") or not f2_norm.startswith("="):
        return False

    cuerpo1 = f1_norm[1:].strip()
    cuerpo2 = f2_norm[1:].strip()

    if "(" in cuerpo1 or "(" in cuerpo2:
        return False

    for operador in ("+", "*"):
        otros_ops = ["-", "/"] if operador == "+" else ["-", "/", "+"]
        if operador == "*":
            otros_ops = ["-", "/", "+"]
        elif operador == "+":
            otros_ops = ["-", "/", "*"]

        tiene_op1 = operador in cuerpo1
        tiene_op2 = operador in cuerpo2

        if not tiene_op1 or not tiene_op2:
            continue

        tiene_otros1 = any(op in cuerpo1 for op in otros_ops)
        tiene_otros2 = any(op in cuerpo2 for op in otros_ops)

        if tiene_otros1 or tiene_otros2:
            continue

        operandos1 = sorted(part.strip() for part in cuerpo1.split(operador))
        operandos2 = sorted(part.strip() for part in cuerpo2.split(operador))

        if operandos1 == operandos2:
            return True

    return False


def _formulas_son_equivalentes(f1_norm, f2_norm):
    """
    Evalúa si dos fórmulas normalizadas son matemáticamente equivalentes,
    cubriendo:
      1. Separadores intercambiables (',' y ';').
      2. Expansión de rangos vs celdas desglosadas (ej. H5:I5 == H5;I5 o J3:J8 == J3;J4;...;J8).
      3. Equivalencia entre SUM(...) y suma directa A+B+C...
      4. Equivalencia entre PRODUCT(...) y multiplicación A*B*C...
      5. Conmutatividad pura (+ y *).
      6. Funciones en español vs inglés (SUMA == SUM, PROMEDIO == AVERAGE).
    """
    if not f1_norm or not f2_norm:
        return False
    if f1_norm == f2_norm:
        return True

    # 1. Comparar como funciones de agregación (SUM, AVERAGE, etc.)
    func1, celdas1 = _analizar_formula_agregacion(f1_norm)
    func2, celdas2 = _analizar_formula_agregacion(f2_norm)

    if func1 and func2:
        if func1 == func2 and celdas1 == celdas2:
            return True

    # 2. Equivalencia SUM(...) vs cadena de sumas A+B+C...
    if func1 == "SUM":
        sumandos2 = _analizar_operacion_cadena(f2_norm, "+")
        if sumandos2 and celdas1 == sumandos2:
            return True
    if func2 == "SUM":
        sumandos1 = _analizar_operacion_cadena(f1_norm, "+")
        if sumandos1 and celdas2 == sumandos1:
            return True

    # 3. Equivalencia PRODUCT(...) vs cadena de multiplicaciones A*B*C...
    if func1 == "PRODUCT":
        factores2 = _analizar_operacion_cadena(f2_norm, "*")
        if factores2 and celdas1 == factores2:
            return True
    if func2 == "PRODUCT":
        factores1 = _analizar_operacion_cadena(f1_norm, "*")
        if factores1 and celdas2 == factores1:
            return True

    # 4. Ambas son cadenas de sumas directas (+ conmutativo, sin '$')
    sumandos1 = _analizar_operacion_cadena(f1_norm, "+")
    sumandos2 = _analizar_operacion_cadena(f2_norm, "+")
    if sumandos1 and sumandos2 and sumandos1 == sumandos2:
        return True

    # 5. Ambas son cadenas de multiplicaciones directas (* conmutativo, sin '$')
    factores1 = _analizar_operacion_cadena(f1_norm, "*")
    factores2 = _analizar_operacion_cadena(f2_norm, "*")
    if factores1 and factores2 and factores1 == factores2:
        return True

    # 6. Conmutatividad con operandos literales (ej. C3*3% vs 3%*C3)
    if _formulas_conmutativas_equivalentes(f1_norm, f2_norm):
        return True

    return False


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
            # Antes de marcar error, verificar si son equivalentes
            # (conmutatividad, rangos desglosados, suma directa vs SUM, etc.)
            if not _formulas_son_equivalentes(formula_p_norm, formula_e_norm):
                errores.append(
                    MENSAJES["formula"].format(
                        esperado=formula_p_raw,
                        encontrado=formula_e_raw or "(vacío)"
                    )
                )
                # Comparar funciones utilizadas solo si las fórmulas NO son equivalentes
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
