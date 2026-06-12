# Manual de Usuario: Auditor de Excel (Panel de Control Gráfico)

Bienvenido al **Auditor de Excel - Comparador de Fidelidad Total**. Este es un sistema de grado profesional diseñado para automatizar la revisión masiva de trabajos, exámenes y tareas prácticas de Microsoft Excel entregados por estudiantes, comparándolos celda por celda contra una plantilla de referencia.

El sistema evalúa de forma inteligente valores, fórmulas, formatos (colores de celda, fuentes, negritas, cursivas, bordes), anchos de columnas, altos de filas, celdas combinadas, tablas de Excel, validaciones de datos y objetos complejos (como gráficos).

---

## 🛠️ Requisitos Previos

Antes de utilizar la aplicación, debes asegurar que tu computadora cuenta con:

1. **Python 3 instalado**:
   * **Windows**: Descarga el instalador desde [python.org](https://www.python.org/downloads/). Al instalarlo, es **imprescindible marcar la casilla "Add Python to PATH"**.
   * **Linux / macOS**: Viene integrado de forma nativa. Puedes validarlo ejecutando `python3 --version` en una terminal.
2. **Dependencias del sistema**:
   La primera vez que ejecutes el sistema (o usando el lanzador rápido), se instalarán de forma automática las librerías necesarias (`openpyxl` y `pywin32` para Windows).

---

## 📁 Preparación de Archivos (Estructura de Directorios)

Para que el auditor localice tus archivos, tu carpeta del proyecto debe organizarse de la siguiente manera:

1. **La Carpeta `PLANTILLAS`**:
   * Crea una carpeta llamada `PLANTILLAS` en la raíz.
   * Guarda allí tu archivo maestro de respuestas correctas con el nombre exacto **`PLANTILLA.xlsx`**.
2. **La Carpeta `TRABAJOS_ESTUDIANTES`**:
   * Coloca en esta carpeta todos los archivos de Excel entregados por tus estudiantes (deben ser archivos `.xlsx`).
3. **La Carpeta `LOGS`**:
   * Se crea de forma automática la primera vez que ejecutas el programa. Guardará los resultados finales y los registros históricos.

### Vista de la Estructura de Carpetas:
```text
APP_REVISA_EXCEL/
├── PLANTILLAS/
│   └── PLANTILLA.xlsx                 <-- Tu archivo maestro de respuestas
├── TRABAJOS_ESTUDIANTES/              <-- Archivos entregados por alumnos
│   ├── juan_perez.xlsx
│   ├── maria_gomez.xlsx
│   └── ...
├── REVISADOS/                         <-- Archivos calificados (se genera sola)
├── LOGS/                              <-- Reportes Excel e Historiales (se genera sola)
├── EJECUTAR_INTERFAZ.bat              <-- Lanzador de Interfaz (Windows)
├── EJECUTAR_INTERFAZ.sh               <-- Lanzador de Interfaz (Linux)
├── gui_server.py                      <-- Servidor Backend local
└── gui/                               <-- Recursos Web de la Interfaz
```

---

## 🚀 Uso del Panel de Control Gráfico (Recomendado)

La aplicación cuenta con una interfaz web responsiva y moderna con diseño *glassmorphic* (translúcido) que facilita la configuración y ejecución. 

> [!TIP]
> **Uso inalámbrico en el aula:** El servidor se enlaza a la red local. Al iniciar la aplicación en tu computadora, puedes controlarla desde tu teléfono celular o tablet conectándote a la red Wi-Fi del aula e ingresando en tu móvil la dirección `http://<ip-de-tu-computadora>:5000`.

### 1. Iniciar la Interfaz
* **En Windows**: Haz doble clic sobre el archivo **`EJECUTAR_INTERFAZ.bat`**.
* **En Linux / macOS**: Ejecuta el archivo **`EJECUTAR_INTERFAZ.sh`** desde tu gestor de archivos o mediante terminal:
  ```bash
  ./EJECUTAR_INTERFAZ.sh
  ```
Esto iniciará el servidor local y abrirá tu navegador web predeterminado en `http://localhost:5000`.

### 2. Configurar la Auditoría
Una vez abierta la interfaz en tu navegador o móvil:
* **Sistema Operativo**: Elige si el servidor de auditoría corre en **Windows** o **Linux** mediante el selector superior.
* **Plantilla y Trabajos**: La app prellena las rutas por defecto. Puedes utilizar los botones **Buscar...** para abrir una ventana nativa de tu computadora y seleccionar un archivo de plantilla o una carpeta de trabajos diferente.
* **Sección / Grupo**: Escribe la sección correspondiente (ej: `11-5B`).
  * *Auto-sugerencia*: Al buscar y seleccionar tu carpeta de trabajos, la interfaz extraerá el nombre del directorio y autocompletará este campo automáticamente si lo dejaste vacío.
* **Fecha de Creación**: Escribe la fecha descriptiva o código para el reporte (ej: `11MAY26` o `120626`).
  * *Prellenado Inteligente*: Al cargar, la interfaz preconfigura la fecha actual del sistema en formato corto de 6 dígitos (`ddMMYY`).

### 3. Ejecutar y Monitorear
* Haz clic en el gran botón **Ejecutar Auditoría**. Todos los controles se inhabilitarán para evitar clics dobles accidentales.
* La **Consola de Ejecución** aparecerá en la parte inferior mostrando en tiempo real los resultados de la auditoría. Las líneas tienen colores según su estado (Verde: Éxito/Aciertos, Rojo: Errores, Amarillo: Advertencias, Blanco/Gris: Sistema).

### 4. Cancelar o Apagar
* **Cancelar Auditoría**: Si necesitas abortar una revisión en progreso, presiona el botón **Cancelar** (`■`). El flujo SSE se desconectará en el cliente y detendrá de forma segura el subproceso de Python en el servidor.
* **Apagar Servidor**: Cuando termines tus labores de revisión, haz clic en el botón rojo **✕ Apagar Servidor** en la esquina superior derecha. Tras confirmar, el servidor HTTP se detendrá de forma segura y **la ventana de comandos (CMD) se cerrará automáticamente en tu sistema**, liberando el puerto.

---

## 📊 Interpretación de Resultados

Cuando finalice el proceso, la aplicación generará dos salidas principales:

### 1. Retroalimentación Detallada (Carpeta `REVISADOS/`)
Dentro de la carpeta `REVISADOS` se creará una copia calificada de cada archivo del alumno con el sufijo `_REV` (ej: `juan_perez_REV.xlsx`).
* El auditor **pintará de color rojo** el fondo de cualquier celda donde el estudiante tenga un fallo.
* Insertará un **comentario de Excel** detallando qué se esperaba (ej: valor esperado, fórmula esperada, o fallo de formato) y qué fue lo que encontró en el archivo del estudiante.

### 2. Informe Consolidado de Calificaciones (Carpeta `LOGS/`)
Se generará un archivo consolidado en Excel y su respectivo registro de auditoría con los nombres dinámicos especificados:
* **Excel consolidado**: `LOGS/{seccion}_LOG_NOTAS_{fecha}.xlsx` (ej: `11-5B_LOG_NOTAS_11MAY26.xlsx`).
* **Historial log**: `LOGS/{seccion}_auditoria_{fecha}.log` (ej: `11-5B_auditoria_11MAY26.log`).

El libro Excel consolidado cuenta con el siguiente formato profesional:
* **Filtros Automáticos**: Fila de encabezado lista para filtrar calificaciones o buscar estudiantes.
* **Ajuste de Texto**: Celdas con wrap-text habilitado para leer cómodamente textos largos.
* **Visualización de Porcentajes**: Columnas de porcentaje resaltadas con fondo verde suave y texto verde oscuro en negrita de alta legibilidad, facilitando la identificación de notas finales y rendimientos por hoja.
* **Métricas**: Detalla la cantidad exacta de aciertos y errores por cada una de las pestañas visibles del libro de Excel, la suma total de aciertos/errores de todo el libro, el porcentaje total, y los errores de objetos complejos.

---

## 💻 Uso Avanzado por Consola (CLI)

Si prefieres omitir la interfaz gráfica, puedes invocar el script directamente desde la terminal de tu sistema operativo dentro de la carpeta del proyecto:

```bash
.venv/Scripts/python.exe main.py --plantilla "PLANTILLAS/PLANTILLA.xlsx" --trabajos "TRABAJOS_ESTUDIANTES" --seccion "11-5B" --fecha "11MAY26"
```

### Argumentos de Consola Disponibles:
* `--plantilla`: Ruta al archivo de plantilla de Excel.
* `--trabajos`: Ruta a la carpeta con los archivos de los estudiantes.
* `--salida`: Carpeta donde se guardarán los Excel corregidos (default: `REVISADOS`).
* `--log`: Especifica un nombre/ruta particular de salida de reporte (si se omite, se guarda por defecto en `LOGS/` con el patrón `{seccion}_LOG_NOTAS_{fecha}.xlsx`).
* `--seccion`: Nombre del grupo de estudiantes (default: extraído de la carpeta de trabajos).
* `--fecha`: Fecha del reporte en formato texto (default: fecha corta del sistema `ddMMYY`).
