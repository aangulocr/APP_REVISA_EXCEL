# ============================================================================
# main.py — Punto de entrada principal del Auditor de Excel
# ============================================================================
"""
Script principal que:
  1. Valida la existencia de la PLANTILLA.xlsx y la carpeta TRABAJOS_ESTUDIANTES.
  2. Itera por todos los archivos .xlsx de la carpeta.
  3. Ejecuta la auditoría completa para cada archivo.
  4. Guarda los archivos revisados en la subcarpeta /REVISADOS.
  5. Genera el archivo LOG_NOTAS.csv con el resumen final.

Uso:
    python main.py
    python main.py --plantilla "ruta/a/PLANTILLA.xlsx" --trabajos "ruta/a/TRABAJOS_ESTUDIANTES"
"""

import os
import sys
import csv
import argparse
import logging
from datetime import datetime

from config import (
    PLANTILLA_PATH,
    TRABAJOS_DIR,
    REVISADOS_DIR,
    LOG_NOTAS_PATH,
    EXTENSIONES_VALIDAS,
)
from auditor import auditar_libro


# ============================================================================
# CONFIGURACIÓN DE LOGGING
# ============================================================================
def configurar_logging():
    """Configura el sistema de logging con salida a consola y archivo."""
    log_format = "%(asctime)s | %(levelname)-7s | %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"

    # Logger raíz
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Handler de consola
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter(log_format, datefmt=date_format))
    logger.addHandler(ch)

    # Handler de archivo
    log_file = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        f"auditoria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(log_format, datefmt=date_format))
    logger.addHandler(fh)

    return logger


# ============================================================================
# VALIDACIONES DE ENTRADA
# ============================================================================
def validar_rutas(ruta_plantilla, ruta_trabajos):
    """
    Valida que la plantilla exista y la carpeta de trabajos exista y
    contenga archivos .xlsx.

    Retorna:
        list: Lista de rutas absolutas a los archivos .xlsx encontrados.

    Lanza:
        FileNotFoundError: Si la plantilla o la carpeta no existen.
        ValueError: Si no se encuentran archivos .xlsx.
    """
    # Validar plantilla
    if not os.path.isfile(ruta_plantilla):
        raise FileNotFoundError(
            f"❌ No se encontró la plantilla: '{ruta_plantilla}'\n"
            f"   Asegúrate de que el archivo PLANTILLA.xlsx exista en la ruta indicada."
        )

    # Validar carpeta de trabajos
    if not os.path.isdir(ruta_trabajos):
        raise FileNotFoundError(
            f"❌ No se encontró la carpeta de trabajos: '{ruta_trabajos}'\n"
            f"   Asegúrate de que la carpeta TRABAJOS_ESTUDIANTES exista."
        )

    # Buscar archivos .xlsx
    archivos = []
    for archivo in sorted(os.listdir(ruta_trabajos)):
        if archivo.lower().endswith(EXTENSIONES_VALIDAS) and not archivo.startswith("~$"):
            archivos.append(os.path.join(ruta_trabajos, archivo))

    if not archivos:
        raise ValueError(
            f"❌ No se encontraron archivos .xlsx en: '{ruta_trabajos}'\n"
            f"   Coloca los trabajos de los estudiantes en esa carpeta."
        )

    return archivos


# ============================================================================
# GENERACIÓN DEL LOG CSV
# ============================================================================
def generar_log_csv(resultados, ruta_csv):
    """
    Genera el archivo LOG_NOTAS.csv con el resumen de la auditoría.

    Columnas:
        - Nombre del Estudiante (sin extensión)
        - Cantidad de Aciertos
        - Cantidad de Errores
        - Porcentaje de Aciertos
        - Hojas Analizadas
        - Detalles de Objetos (resumen)
    """
    with open(ruta_csv, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")

        # Encabezados
        writer.writerow([
            "Nombre del Estudiante",
            "Cantidad de Aciertos",
            "Cantidad de Errores",
            "Total Evaluaciones",
            "Porcentaje de Aciertos (%)",
            "Hojas Analizadas",
            "Errores en Objetos/COM"
        ])

        for res in resultados:
            nombre = os.path.splitext(res["archivo"])[0]
            aciertos = res["total_aciertos"]
            errores_val = res["total_errores"]
            total = aciertos + errores_val if errores_val >= 0 else 0
            porcentaje = (aciertos / total * 100) if total > 0 else 0

            # Contar errores de objetos
            errores_objetos = sum(
                len(h.get("detalles_objetos", []))
                for h in res.get("detalle_hojas", [])
            )
            errores_com = len(res.get("detalles_com", []))

            hojas = ", ".join(
                h["hoja"] for h in res.get("detalle_hojas", [])
            )

            writer.writerow([
                nombre,
                aciertos,
                errores_val,
                total,
                f"{porcentaje:.1f}",
                hojas,
                errores_objetos + errores_com
            ])

    logging.info(f"\n📊 LOG_NOTAS.csv generado en: {ruta_csv}")


# ============================================================================
# FLUJO PRINCIPAL
# ============================================================================
def main():
    """Punto de entrada principal del script de auditoría."""
    logger = configurar_logging()

    # Parsear argumentos opcionales
    parser = argparse.ArgumentParser(
        description="🔍 Auditor de Excel — Comparación de trabajos de estudiantes vs plantilla"
    )
    parser.add_argument(
        "--plantilla",
        default=PLANTILLA_PATH,
        help=f"Ruta al archivo PLANTILLA.xlsx (default: {PLANTILLA_PATH})"
    )
    parser.add_argument(
        "--trabajos",
        default=TRABAJOS_DIR,
        help=f"Ruta a la carpeta TRABAJOS_ESTUDIANTES (default: {TRABAJOS_DIR})"
    )
    parser.add_argument(
        "--salida",
        default=REVISADOS_DIR,
        help=f"Ruta a la carpeta de salida REVISADOS (default: {REVISADOS_DIR})"
    )
    parser.add_argument(
        "--log",
        default=LOG_NOTAS_PATH,
        help=f"Ruta al archivo LOG_NOTAS.csv (default: {LOG_NOTAS_PATH})"
    )
    args = parser.parse_args()

    # Banner
    logger.info("=" * 70)
    logger.info("  🔍 AUDITOR DE EXCEL — Comparación de Fidelidad Total")
    logger.info(f"  📅 Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)

    # Validar rutas
    try:
        archivos_estudiantes = validar_rutas(args.plantilla, args.trabajos)
    except (FileNotFoundError, ValueError) as e:
        logger.error(str(e))
        sys.exit(1)

    logger.info(f"\n📂 Plantilla:  {args.plantilla}")
    logger.info(f"📂 Trabajos:   {args.trabajos}")
    logger.info(f"📂 Salida:     {args.salida}")
    logger.info(f"📄 Archivos encontrados: {len(archivos_estudiantes)}")

    # Crear carpeta de salida
    os.makedirs(args.salida, exist_ok=True)

    # ----------------------------------------------------------------
    # ITERAR POR CADA ARCHIVO DE ESTUDIANTE
    # ----------------------------------------------------------------
    resultados = []
    total_archivos = len(archivos_estudiantes)

    for idx, ruta_estudiante in enumerate(archivos_estudiantes, 1):
        nombre = os.path.basename(ruta_estudiante)
        logger.info(f"\n[{idx}/{total_archivos}] Procesando: {nombre}")

        # Agregar sufijo _REV al nombre del archivo revisado
        nombre_sin_ext, extension = os.path.splitext(nombre)
        nombre_revisado = f"{nombre_sin_ext}_REV{extension}"
        ruta_salida = os.path.join(args.salida, nombre_revisado)

        try:
            resultado = auditar_libro(args.plantilla, ruta_estudiante, ruta_salida)
            resultados.append(resultado)
        except Exception as e:
            logger.error(f"❌ Error inesperado procesando '{nombre}': {e}")
            resultados.append({
                "archivo": nombre,
                "total_aciertos": 0,
                "total_errores": -1,
                "detalle_hojas": [],
                "detalles_com": [f"Error fatal: {e}"]
            })

    # ----------------------------------------------------------------
    # GENERAR LOG CSV
    # ----------------------------------------------------------------
    generar_log_csv(resultados, args.log)

    # ----------------------------------------------------------------
    # RESUMEN FINAL
    # ----------------------------------------------------------------
    logger.info("\n" + "=" * 70)
    logger.info("  📊 RESUMEN FINAL DE AUDITORÍA")
    logger.info("=" * 70)
    logger.info(f"  Archivos procesados: {total_archivos}")

    total_aciertos_global = sum(r["total_aciertos"] for r in resultados)
    total_errores_global = sum(
        r["total_errores"] for r in resultados if r["total_errores"] >= 0
    )

    logger.info(f"  Total aciertos (global):  {total_aciertos_global}")
    logger.info(f"  Total errores (global):   {total_errores_global}")
    logger.info(f"\n  📁 Archivos revisados en: {args.salida}")
    logger.info(f"  📄 Log de notas en:       {args.log}")
    logger.info("=" * 70)

    # Tabla resumen por estudiante
    logger.info(f"\n{'Estudiante':<40} {'Aciertos':>10} {'Errores':>10} {'%':>8}")
    logger.info("-" * 70)
    for res in resultados:
        nombre = os.path.splitext(res["archivo"])[0]
        ac = res["total_aciertos"]
        er = res["total_errores"]
        total = ac + er if er >= 0 else 0
        pct = (ac / total * 100) if total > 0 else 0
        estado = "✅" if er == 0 else "⚠️" if er < 5 else "❌"
        logger.info(f"  {estado} {nombre:<37} {ac:>10} {er:>10} {pct:>7.1f}%")

    logger.info("\n✅ Auditoría completada.")


if __name__ == "__main__":
    main()
