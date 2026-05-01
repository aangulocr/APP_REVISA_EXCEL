# ============================================================================
# comparador_com.py — Inspección profunda vía COM (win32com) para Windows
# ============================================================================
"""
Módulo que utiliza la API COM de Excel (vía win32com.client) para inspeccionar
objetos que openpyxl no puede leer completamente:
  - Tablas Dinámicas (Pivot Tables)
  - Gráficos Dinámicos (Pivot Charts)
  - Gráficos con detalle fino (series, ejes, formato)

Requiere: Microsoft Excel instalado en el equipo.
"""

import os
import logging

logger = logging.getLogger(__name__)

# Intentar importar win32com; si no está disponible, deshabilitar este módulo
try:
    import win32com.client
    import pythoncom
    COM_DISPONIBLE = True
except ImportError:
    COM_DISPONIBLE = False
    logger.warning(
        "⚠️  win32com no está disponible. "
        "La inspección profunda de Tablas Dinámicas y Gráficos COM estará deshabilitada."
    )


def iniciar_excel_com():
    """
    Inicia una instancia de Excel vía COM.
    Retorna el objeto Application o None si COM no está disponible.
    """
    if not COM_DISPONIBLE:
        return None
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        return excel
    except Exception as e:
        logger.error(f"❌ No se pudo iniciar Excel vía COM: {e}")
        return None


def cerrar_excel_com(excel):
    """Cierra la instancia de Excel COM de forma segura."""
    if excel is None:
        return
    try:
        excel.Quit()
    except Exception:
        pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def comparar_tablas_dinamicas_com(ruta_plantilla, ruta_estudiante, excel=None):
    """
    Compara las Tablas Dinámicas (Pivot Tables) entre la plantilla y el
    archivo del estudiante usando COM.

    Parámetros:
        ruta_plantilla: Ruta absoluta al archivo PLANTILLA.xlsx.
        ruta_estudiante: Ruta absoluta al archivo del estudiante.
        excel: Instancia de Excel COM (opcional, se crea si no se provee).

    Retorna:
        tuple: (errores_count: int, detalles: list[str])
    """
    if not COM_DISPONIBLE:
        return 0, ["⚠️ Inspección COM deshabilitada (win32com no disponible)."]

    errores = 0
    detalles = []
    excel_propio = False

    try:
        if excel is None:
            excel = iniciar_excel_com()
            excel_propio = True

        if excel is None:
            return 0, ["⚠️ No se pudo iniciar Excel para inspección COM."]

        # Abrir ambos archivos
        wb_p = excel.Workbooks.Open(os.path.abspath(ruta_plantilla), ReadOnly=True)
        wb_e = excel.Workbooks.Open(os.path.abspath(ruta_estudiante), ReadOnly=True)

        # Iterar por hojas de la plantilla
        for i in range(1, wb_p.Worksheets.Count + 1):
            ws_p = wb_p.Worksheets(i)
            nombre_hoja = ws_p.Name

            # Buscar hoja homónima en el estudiante
            try:
                ws_e = wb_e.Worksheets(nombre_hoja)
            except Exception:
                continue  # La hoja faltante ya se detecta en el flujo principal

            # Comparar Pivot Tables
            pt_count_p = ws_p.PivotTables().Count
            pt_count_e = ws_e.PivotTables().Count

            if pt_count_p != pt_count_e:
                msg = (f"❌ Hoja '{nombre_hoja}': Cantidad de Tablas Dinámicas incorrecta. "
                       f"Esperadas: {pt_count_p} | Encontradas: {pt_count_e}")
                detalles.append(msg)
                errores += abs(pt_count_p - pt_count_e)

            # Comparar cada tabla dinámica por nombre
            for j in range(1, pt_count_p + 1):
                pt_p = ws_p.PivotTables(j)
                nombre_pt = pt_p.Name

                # Buscar en el estudiante
                pt_e = None
                for k in range(1, pt_count_e + 1):
                    if ws_e.PivotTables(k).Name == nombre_pt:
                        pt_e = ws_e.PivotTables(k)
                        break

                if pt_e is None:
                    msg = (f"❌ Hoja '{nombre_hoja}': Tabla Dinámica faltante: '{nombre_pt}'")
                    detalles.append(msg)
                    errores += 1
                    continue

                # Comparar fuente de datos
                source_p = pt_p.SourceData if hasattr(pt_p, 'SourceData') else ""
                source_e = pt_e.SourceData if hasattr(pt_e, 'SourceData') else ""
                try:
                    if str(source_p) != str(source_e):
                        msg = (f"❌ TD '{nombre_pt}' en '{nombre_hoja}': "
                               f"Fuente de datos incorrecta. "
                               f"Esperada: '{source_p}' | Encontrada: '{source_e}'")
                        detalles.append(msg)
                        errores += 1
                except Exception:
                    pass

                # Comparar campos de fila
                try:
                    campos_fila_p = [pt_p.RowFields(f).Name
                                     for f in range(1, pt_p.RowFields.Count + 1)]
                    campos_fila_e = [pt_e.RowFields(f).Name
                                     for f in range(1, pt_e.RowFields.Count + 1)]
                    if campos_fila_p != campos_fila_e:
                        msg = (f"❌ TD '{nombre_pt}' en '{nombre_hoja}': "
                               f"Campos de fila incorrectos. "
                               f"Esperados: {campos_fila_p} | Encontrados: {campos_fila_e}")
                        detalles.append(msg)
                        errores += 1
                except Exception:
                    pass

                # Comparar campos de columna
                try:
                    campos_col_p = [pt_p.ColumnFields(f).Name
                                    for f in range(1, pt_p.ColumnFields.Count + 1)]
                    campos_col_e = [pt_e.ColumnFields(f).Name
                                    for f in range(1, pt_e.ColumnFields.Count + 1)]
                    if campos_col_p != campos_col_e:
                        msg = (f"❌ TD '{nombre_pt}' en '{nombre_hoja}': "
                               f"Campos de columna incorrectos. "
                               f"Esperados: {campos_col_p} | Encontrados: {campos_col_e}")
                        detalles.append(msg)
                        errores += 1
                except Exception:
                    pass

                # Comparar campos de datos (valores)
                try:
                    campos_datos_p = [pt_p.DataFields(f).Name
                                      for f in range(1, pt_p.DataFields.Count + 1)]
                    campos_datos_e = [pt_e.DataFields(f).Name
                                      for f in range(1, pt_e.DataFields.Count + 1)]
                    if campos_datos_p != campos_datos_e:
                        msg = (f"❌ TD '{nombre_pt}' en '{nombre_hoja}': "
                               f"Campos de datos incorrectos. "
                               f"Esperados: {campos_datos_p} | Encontrados: {campos_datos_e}")
                        detalles.append(msg)
                        errores += 1
                except Exception:
                    pass

        # Cerrar archivos sin guardar
        wb_p.Close(SaveChanges=False)
        wb_e.Close(SaveChanges=False)

    except Exception as e:
        detalles.append(f"⚠️ Error en inspección COM de Tablas Dinámicas: {e}")

    finally:
        if excel_propio and excel:
            cerrar_excel_com(excel)

    return errores, detalles


def comparar_graficos_com(ruta_plantilla, ruta_estudiante, excel=None):
    """
    Compara los gráficos entre plantilla y estudiante usando COM.
    Proporciona detalle más profundo que openpyxl.

    Retorna:
        tuple: (errores_count: int, detalles: list[str])
    """
    if not COM_DISPONIBLE:
        return 0, ["⚠️ Inspección COM deshabilitada (win32com no disponible)."]

    errores = 0
    detalles = []
    excel_propio = False

    try:
        if excel is None:
            excel = iniciar_excel_com()
            excel_propio = True

        if excel is None:
            return 0, ["⚠️ No se pudo iniciar Excel para inspección COM."]

        wb_p = excel.Workbooks.Open(os.path.abspath(ruta_plantilla), ReadOnly=True)
        wb_e = excel.Workbooks.Open(os.path.abspath(ruta_estudiante), ReadOnly=True)

        for i in range(1, wb_p.Worksheets.Count + 1):
            ws_p = wb_p.Worksheets(i)
            nombre_hoja = ws_p.Name

            try:
                ws_e = wb_e.Worksheets(nombre_hoja)
            except Exception:
                continue

            # Gráficos incrustados (ChartObjects)
            charts_p = ws_p.ChartObjects()
            charts_e = ws_e.ChartObjects()

            count_p = charts_p.Count
            count_e = charts_e.Count

            if count_p != count_e:
                msg = (f"❌ Hoja '{nombre_hoja}': Cantidad de gráficos incorrecta. "
                       f"Esperados: {count_p} | Encontrados: {count_e}")
                detalles.append(msg)
                errores += abs(count_p - count_e)

            for j in range(1, min(count_p, count_e) + 1):
                chart_p = charts_p(j).Chart
                chart_e = charts_e(j).Chart

                # Tipo de gráfico (constante numérica de Excel)
                tipo_p = chart_p.ChartType
                tipo_e = chart_e.ChartType
                if tipo_p != tipo_e:
                    msg = (f"❌ Gráfico #{j} en '{nombre_hoja}': "
                           f"Tipo incorrecto. "
                           f"Esperado: {tipo_p} | Encontrado: {tipo_e}")
                    detalles.append(msg)
                    errores += 1

                # Título
                tiene_titulo_p = chart_p.HasTitle
                tiene_titulo_e = chart_e.HasTitle
                if tiene_titulo_p and tiene_titulo_e:
                    titulo_p = chart_p.ChartTitle.Text
                    titulo_e = chart_e.ChartTitle.Text
                    if titulo_p != titulo_e:
                        msg = (f"❌ Gráfico #{j} en '{nombre_hoja}': "
                               f"Título incorrecto. "
                               f"Esperado: '{titulo_p}' | Encontrado: '{titulo_e}'")
                        detalles.append(msg)
                        errores += 1
                elif tiene_titulo_p != tiene_titulo_e:
                    msg = (f"❌ Gráfico #{j} en '{nombre_hoja}': "
                           f"{'Falta título' if tiene_titulo_p else 'Título extra'}")
                    detalles.append(msg)
                    errores += 1

                # Series de datos
                series_count_p = chart_p.SeriesCollection().Count
                series_count_e = chart_e.SeriesCollection().Count
                if series_count_p != series_count_e:
                    msg = (f"❌ Gráfico #{j} en '{nombre_hoja}': "
                           f"Series de datos incorrectas. "
                           f"Esperadas: {series_count_p} | Encontradas: {series_count_e}")
                    detalles.append(msg)
                    errores += 1

                # Comparar cada serie
                for s in range(1, min(series_count_p, series_count_e) + 1):
                    try:
                        nombre_serie_p = chart_p.SeriesCollection(s).Name
                        nombre_serie_e = chart_e.SeriesCollection(s).Name
                        if nombre_serie_p != nombre_serie_e:
                            msg = (f"❌ Gráfico #{j}, Serie #{s} en '{nombre_hoja}': "
                                   f"Nombre incorrecto. "
                                   f"Esperado: '{nombre_serie_p}' | "
                                   f"Encontrado: '{nombre_serie_e}'")
                            detalles.append(msg)
                            errores += 1
                    except Exception:
                        pass

        wb_p.Close(SaveChanges=False)
        wb_e.Close(SaveChanges=False)

    except Exception as e:
        detalles.append(f"⚠️ Error en inspección COM de gráficos: {e}")

    finally:
        if excel_propio and excel:
            cerrar_excel_com(excel)

    return errores, detalles


def inspeccion_profunda_com(ruta_plantilla, ruta_estudiante):
    """
    Ejecuta la inspección profunda completa vía COM.
    Agrupa Tablas Dinámicas y Gráficos en una sola sesión de Excel.

    Retorna:
        tuple: (total_errores: int, todos_detalles: list[str])
    """
    if not COM_DISPONIBLE:
        return 0, [
            "⚠️ Módulo win32com no disponible. "
            "Instálalo con: pip install pywin32"
        ]

    total_errores = 0
    todos_detalles = []
    excel = None

    try:
        excel = iniciar_excel_com()
        if excel is None:
            return 0, ["⚠️ No se pudo iniciar Excel para inspección COM."]

        # Tablas Dinámicas
        err_td, det_td = comparar_tablas_dinamicas_com(
            ruta_plantilla, ruta_estudiante, excel
        )
        total_errores += err_td
        todos_detalles.extend(det_td)

        # Gráficos vía COM
        err_gc, det_gc = comparar_graficos_com(
            ruta_plantilla, ruta_estudiante, excel
        )
        total_errores += err_gc
        todos_detalles.extend(det_gc)

    finally:
        cerrar_excel_com(excel)

    return total_errores, todos_detalles
