# Auditor de Excel - Comparador de Fidelidad Total

Este es un sistema automatizado (el "Motor principal de auditoría") diseñado para comparar masivamente archivos de Excel entregados por estudiantes contra un archivo "plantilla" o "rúbrica" del profesor.

El programa evalúa celda por celda, revisando fórmulas, valores, formatos (color, bordes), validaciones de datos, tablas, gráficos y otros objetos de Excel, y entrega un archivo calificado marcando los errores, además de un resumen de notas en formato CSV.

---

## 🛠️ Requisitos Previos

Antes de poder utilizar el programa, necesitas tener instalado lo siguiente en tu computadora:

1. **Python 3:**
   - **Windows:** Descárgalo desde [python.org](https://www.python.org/downloads/). ¡Muy importante! Durante la instalación, **marca la casilla que dice "Add Python to PATH"** (Agregar Python al PATH).
   - **Linux:** Generalmente ya viene instalado. Puedes verificarlo abriendo una terminal y escribiendo `python3 --version`.
2. **Dependencias de Python:**
   - Abre una terminal (o Símbolo de sistema / CMD en Windows) en la carpeta de este proyecto y ejecuta el siguiente comando para instalar las librerías necesarias:
     ```bash
     pip install -r requirements.txt
     ```
     *(En algunos sistemas Linux, puede que necesites usar `pip3 install -r requirements.txt`).*

---

## 📁 Preparación de los Archivos

Para que el programa funcione, el directorio principal del proyecto debe tener exactamente la siguiente estructura de archivos preparados por ti:

1. **`PLANTILLA.xlsx`:** Este es tu archivo maestro o rúbrica. Contiene las respuestas correctas, formatos esperados, etc. *Debe llamarse exactamente así y estar en la carpeta principal.*
2. **Carpeta `TRABAJOS_ESTUDIANTES`:** Crea una carpeta con este nombre exacto en el directorio principal.
3. **Archivos de los alumnos:** Coloca todos los archivos `.xlsx` que entregaron los estudiantes dentro de la carpeta `TRABAJOS_ESTUDIANTES`.

La carpeta del proyecto debería verse algo así:
```text
APP_REVISA_EXCEL/
├── EJECUTAR_AUDITORIA.bat
├── EJECUTAR_AUDITORIA.sh
├── main.py
├── auditor.py
├── ... (otros archivos .py)
├── PLANTILLA.xlsx                 <-- ¡Tu archivo maestro!
└── TRABAJOS_ESTUDIANTES/          <-- ¡Carpeta con los trabajos!
    ├── juan_perez.xlsx
    ├── maria_gomez.xlsx
    └── ...
```

*(Nota: Si quieres probar el programa sin tener archivos reales, puedes ejecutar el script `crear_datos_prueba.py` el cual generará automáticamente una plantilla y trabajos simulados para que veas cómo funciona).*

---

## 🚀 Cómo Ejecutar la Auditoría

Una vez que tengas la plantilla y los trabajos en su lugar, iniciar el proceso es muy sencillo dependiendo de tu sistema operativo:

### En Windows
1. Haz **doble clic** en el archivo **`EJECUTAR_AUDITORIA.bat`**.
2. Se abrirá una ventana negra (consola) mostrando el progreso de la revisión hoja por hoja y alumno por alumno.
3. Al finalizar, la ventana te avisará que la auditoría fue completada exitosamente. Presiona cualquier tecla para cerrar.

### En Linux
1. Abre tu explorador de archivos (por ejemplo, Nemo, Nautilus).
2. Ve a la carpeta del proyecto.
3. Haz **doble clic** en el archivo **`EJECUTAR_AUDITORIA.sh`** y selecciona la opción **"Ejecutar en un terminal"**.
4. Alternativamente, puedes abrir tu terminal en esa carpeta y ejecutar:
   ```bash
   ./EJECUTAR_AUDITORIA.sh
   ```

---

## 📊 Resultados de la Auditoría

Cuando el programa termina de ejecutarse, generará automáticamente lo siguiente:

1. **Carpeta `REVISADOS/`:** Se creará una nueva carpeta con este nombre. Adentro encontrarás una copia de cada trabajo del estudiante pero con el sufijo `_REV` (ej. `juan_perez_REV.xlsx`). Si abres estos archivos, **verás celdas pintadas de rojo o con comentarios** indicando exactamente dónde se equivocó el estudiante respecto a la plantilla.
2. **Archivo `LOG_NOTAS.csv`:** Este archivo es tu registro final de calificaciones. Puedes abrirlo con Excel y verás una tabla resumiendo los resultados de toda la clase:
   - Nombre del estudiante.
   - Cantidad de aciertos y errores.
   - Porcentaje de precisión (nota).
   - Errores específicos en objetos complejos (como gráficos o tablas).
3. **Archivo de Log (ej. `auditoria_20260606_120000.log`):** Un archivo de texto detallado por si necesitas revisar internamente qué comparó el programa paso a paso o por si ocurrió algún error de lectura.
