# ==============================================================================
# ARCHIVO: matcher_hojas.py
# DESCRIPCIÓN: Emparejamiento inteligente de hojas de cálculo entre la plantilla
#              del docente y los libros entregados por los estudiantes.
# ==============================================================================

import re
import unicodedata
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Any


def normalizar_texto(texto: str) -> str:
    """
    Normaliza una cadena:
    - Remueve tildes y diacríticos.
    - Convierte a minúsculas.
    - Reemplaza signos de puntuación por espacios.
    - Elimina espacios repetidos y extremos.
    """
    if not texto:
        return ""
    nfd = unicodedata.normalize("NFD", str(texto))
    sin_tildes = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    limpio = re.sub(r"[^\w\s]", " ", sin_tildes).lower()
    return " ".join(limpio.split())


def es_rubrica(nombre: str) -> bool:
    """Verifica si el nombre de la hoja corresponde a una rúbrica."""
    norm = normalizar_texto(nombre)
    return norm in ("rubrica", "rubricas")


def extraer_numero_actividad(nombre: str) -> Optional[int]:
    """
    Extrae el número identificador de la actividad al inicio del nombre.
    Ejemplos:
        '1. Ordenar Datos' -> 1
        '6. SI.CONJUNTO y Gráficos' -> 6
        'Actividad 3' -> 3
        'Ejercicio 4 - Filtros' -> 4
    """
    if not nombre:
        return None
    m = re.match(r"^\s*(?:hoja|actividad|ejercicio|p)?\s*(\d+)", nombre, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None


def obtener_huella_encabezados(ws) -> set:
    """
    Extrae las palabras clave de los encabezados y textos fijos
    de las primeras 5 filas y 15 columnas de la hoja (openpyxl Worksheet).
    """
    palabras = set()
    if ws is None:
        return palabras

    try:
        max_r = min(ws.max_row or 5, 5)
        max_c = min(ws.max_column or 15, 15)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = ws.cell(row=r, column=c).value
                if val and isinstance(val, str):
                    norm = normalizar_texto(val)
                    for token in norm.split():
                        if len(token) >= 3:  # Palabras significativas
                            palabras.add(token)
    except Exception:
        pass

    return palabras


def emparejar_hojas(
    hojas_plantilla: List[str],
    hojas_estudiante: List[str],
    wb_plantilla=None,
    wb_estudiante=None
) -> Dict[str, Dict[str, Any]]:
    """
    Realiza un emparejamiento inteligente multinivel entre las hojas de la plantilla
    y las hojas del estudiante.

    Retorna un diccionario:
    {
        nombre_plantilla: {
            "hoja_estudiante": str | None,
            "metodo": "exacto" | "normalizado" | "prefijo_numérico" | "fuzzy" | "contenido" | "no_encontrada",
            "score": float
        }
    }
    """
    # Hojas disponibles en el estudiante (excluyendo rúbrica)
    disponibles = [h for h in hojas_estudiante if not es_rubrica(h)]
    mapa = {}

    # Filtrar hojas requeridas de la plantilla (excluyendo rúbrica)
    plantilla_req = [h for h in hojas_plantilla if not es_rubrica(h)]

    # -------------------------------------------------------------
    # FASE 1: Coincidencia Exacta
    # -------------------------------------------------------------
    for p in plantilla_req:
        if p in disponibles:
            mapa[p] = {
                "hoja_estudiante": p,
                "metodo": "exacto",
                "score": 1.0
            }
            disponibles.remove(p)

    # -------------------------------------------------------------
    # FASE 2: Coincidencia Normalizada (sin tildes, minúsculas, espacios)
    # -------------------------------------------------------------
    for p in plantilla_req:
        if p in mapa:
            continue
        p_norm = normalizar_texto(p)
        match_e = None
        for e in disponibles:
            if normalizar_texto(e) == p_norm:
                match_e = e
                break
        if match_e:
            mapa[p] = {
                "hoja_estudiante": match_e,
                "metodo": "normalizado",
                "score": 0.95
            }
            disponibles.remove(match_e)

    # -------------------------------------------------------------
    # FASE 3: Coincidencia por Prefijo / Número de Actividad
    # -------------------------------------------------------------
    for p in plantilla_req:
        if p in mapa:
            continue
        num_p = extraer_numero_actividad(p)
        if num_p is not None:
            candidatos = [
                e for e in disponibles
                if extraer_numero_actividad(e) == num_p
            ]
            if len(candidatos) == 1:
                e_sel = candidatos[0]
                mapa[p] = {
                    "hoja_estudiante": e_sel,
                    "metodo": f"prefijo_numérico ({num_p})",
                    "score": 0.90
                }
                disponibles.remove(e_sel)
            elif len(candidatos) > 1:
                p_norm = normalizar_texto(p)
                mejor_e = None
                mejor_ratio = -1.0
                for c in candidatos:
                    ratio = SequenceMatcher(None, p_norm, normalizar_texto(c)).ratio()
                    if ratio > mejor_ratio:
                        mejor_ratio = ratio
                        mejor_e = c
                if mejor_e:
                    mapa[p] = {
                        "hoja_estudiante": mejor_e,
                        "metodo": f"prefijo_numérico_desempate ({num_p})",
                        "score": 0.85
                    }
                    disponibles.remove(mejor_e)

    # -------------------------------------------------------------
    # FASE 4: Similitud Textual Difusa (Fuzzy Matching)
    # -------------------------------------------------------------
    for p in plantilla_req:
        if p in mapa:
            continue
        p_norm = normalizar_texto(p)
        mejor_e = None
        mejor_score = 0.0

        for e in disponibles:
            e_norm = normalizar_texto(e)
            ratio = SequenceMatcher(None, p_norm, e_norm).ratio()

            # También verificar si una cadena contiene a la otra
            if p_norm and e_norm and (p_norm in e_norm or e_norm in p_norm):
                ratio = max(ratio, 0.75)

            if ratio > mejor_score and ratio >= 0.60:
                mejor_score = ratio
                mejor_e = e

        if mejor_e:
            mapa[p] = {
                "hoja_estudiante": mejor_e,
                "metodo": f"fuzzy ({mejor_score:.2f})",
                "score": mejor_score
            }
            disponibles.remove(mejor_e)

    # -------------------------------------------------------------
    # FASE 5: Huella de Contenido (Headers de columnas)
    # -------------------------------------------------------------
    if wb_plantilla is not None and wb_estudiante is not None:
        for p in plantilla_req:
            if p in mapa:
                continue

            try:
                if p in wb_plantilla.sheetnames:
                    huella_p = obtener_huella_encabezados(wb_plantilla[p])
                    if len(huella_p) >= 3:
                        mejor_e = None
                        mejor_jaccard = 0.0

                        for e in disponibles:
                            if e in wb_estudiante.sheetnames:
                                huella_e = obtener_huella_encabezados(wb_estudiante[e])
                                if huella_e:
                                    inter = len(huella_p.intersection(huella_e))
                                    union = len(huella_p.union(huella_e))
                                    jaccard = inter / union if union > 0 else 0
                                    if jaccard > mejor_jaccard and jaccard >= 0.40:
                                        mejor_jaccard = jaccard
                                        mejor_e = e

                        if mejor_e:
                            mapa[p] = {
                                "hoja_estudiante": mejor_e,
                                "metodo": f"contenido ({mejor_jaccard:.2f})",
                                "score": mejor_jaccard
                            }
                            disponibles.remove(mejor_e)
            except Exception:
                pass

    # -------------------------------------------------------------
    # Hojas no encontradas
    # -------------------------------------------------------------
    for p in plantilla_req:
        if p not in mapa:
            mapa[p] = {
                "hoja_estudiante": None,
                "metodo": "no_encontrada",
                "score": 0.0
            }

    return mapa
