/* ============================================================================
   app.js — Lógica de Interacción para la Interfaz de Auditoría
   ============================================================================ */

document.addEventListener("DOMContentLoaded", () => {
    // Referencias a Elementos del DOM
    const btnWin = document.getElementById("btn-win");
    const btnLin = document.getElementById("btn-lin");
    const inputPlantilla = document.getElementById("plantilla-path");
    const inputTrabajos = document.getElementById("trabajos-path");
    const btnBrowseFile = document.getElementById("btn-browse-file");
    const btnBrowseFolder = document.getElementById("btn-browse-folder");
    const btnRun = document.getElementById("btn-run");
    const consoleArea = document.getElementById("console-area");
    const consoleLog = document.getElementById("console-log");
    const btnClearConsole = document.getElementById("btn-clear-console");

    let selectedOS = "windows";
    let isRunning = false;

    // 1. Cargar Rutas por Defecto al Iniciar
    fetch("/api/defaults")
        .then(response => response.json())
        .then(data => {
            inputPlantilla.value = data.plantilla || "";
            inputTrabajos.value = data.trabajos || "";
        })
        .catch(err => {
            console.error("Error al cargar rutas por defecto:", err);
            appendLog("Error de conexión al cargar configuraciones por defecto.", "error");
        });

    // 2. Control de Selección de Sistema Operativo
    btnWin.addEventListener("click", () => {
        if (isRunning) return;
        selectedOS = "windows";
        btnWin.classList.add("active");
        btnLin.classList.remove("active");
        appendLog("Sistema operativo seleccionado para ejecución: Windows", "system");
    });

    btnLin.addEventListener("click", () => {
        if (isRunning) return;
        selectedOS = "linux";
        btnLin.classList.add("active");
        btnWin.classList.remove("active");
        appendLog("Sistema operativo seleccionado para ejecución: Linux", "system");
    });

    // 3. Selección de Archivo de Plantilla
    btnBrowseFile.addEventListener("click", () => {
        if (isRunning) return;
        
        btnBrowseFile.disabled = true;
        btnBrowseFile.textContent = "Buscando...";
        
        fetch("/api/select-file")
            .then(res => res.json())
            .then(data => {
                if (data.path) {
                    inputPlantilla.value = data.path;
                    appendLog(`Plantilla seleccionada: ${data.path}`, "system");
                }
            })
            .catch(err => {
                console.error("Error al seleccionar archivo:", err);
                appendLog("Error al abrir diálogo de selección de archivo.", "error");
            })
            .finally(() => {
                btnBrowseFile.disabled = false;
                btnBrowseFile.textContent = "Buscar...";
            });
    });

    // 4. Selección de Carpeta de Trabajos
    btnBrowseFolder.addEventListener("click", () => {
        if (isRunning) return;
        
        btnBrowseFolder.disabled = true;
        btnBrowseFolder.textContent = "Buscando...";
        
        fetch("/api/select-folder")
            .then(res => res.json())
            .then(data => {
                if (data.path) {
                    inputTrabajos.value = data.path;
                    appendLog(`Carpeta de trabajos seleccionada: ${data.path}`, "system");
                }
            })
            .catch(err => {
                console.error("Error al seleccionar carpeta:", err);
                appendLog("Error al abrir diálogo de selección de carpeta.", "error");
            })
            .finally(() => {
                btnBrowseFolder.disabled = false;
                btnBrowseFolder.textContent = "Buscar...";
            });
    });

    // 5. Limpieza de Consola
    btnClearConsole.addEventListener("click", () => {
        consoleLog.innerHTML = "";
        appendLog("Consola limpia.", "system");
    });

    // 6. Ejecución de la Auditoría (Stream de Logs en Tiempo Real)
    btnRun.addEventListener("click", () => {
        if (isRunning) return;

        const plantilla = inputPlantilla.value.trim();
        const trabajos = inputTrabajos.value.trim();

        if (!plantilla || !trabajos) {
            alert("Por favor, selecciona tanto el archivo de plantilla como la carpeta de trabajos.");
            return;
        }

        // Preparar UI para ejecución
        isRunning = true;
        btnRun.disabled = true;
        btnRun.innerHTML = `<span class="run-icon">⏳</span> Ejecutando Auditoría...`;
        consoleArea.style.display = "block";
        consoleLog.innerHTML = ""; // Limpiar consola
        appendLog(`Iniciando auditoría en segundo plano (OS: ${selectedOS})...`, "system");

        // Desactivar controles
        btnBrowseFile.disabled = true;
        btnBrowseFolder.disabled = true;
        btnWin.disabled = true;
        btnLin.disabled = true;

        // Crear la URL con los parámetros correspondientes
        const params = new URLSearchParams({
            plantilla: plantilla,
            trabajos: trabajos,
            os: selectedOS
        });

        // Conectar al endpoint SSE de ejecución
        const eventSource = new EventSource(`/api/run-audit?${params.toString()}`);

        eventSource.onmessage = (event) => {
            try {
                const data = JSON.parse(event.data);
                
                if (data.status === "done") {
                    appendLog(`\n🎉 Auditoría completada con código de salida: ${data.code}`, data.code === 0 ? "success" : "error");
                    finishExecution();
                    eventSource.close();
                } else if (data.status === "error") {
                    appendLog(`\n❌ Error de ejecución: ${data.message}`, "error");
                    finishExecution();
                    eventSource.close();
                } else {
                    // Determinar tipo de log para coloreado
                    let logType = "info";
                    const text = data.text || "";
                    
                    if (text.includes("❌") || text.toLowerCase().includes("[error]")) {
                        logType = "error";
                    } else if (text.includes("✅") || text.includes("[ÉXITO]") || text.includes("completada")) {
                        logType = "success";
                    } else if (text.includes("⚠️") || text.toLowerCase().includes("[warning]")) {
                        logType = "warning";
                    } else if (text.startsWith("==") || text.includes("Procesando:") || text.includes("Auditando:")) {
                        logType = "system";
                    }
                    
                    appendLog(text, logType);
                }
            } catch (err) {
                console.error("Error al parsear mensaje de log:", err);
            }
        };

        eventSource.onerror = (err) => {
            console.error("Error en conexión de eventos SSE:", err);
            appendLog("Conexión con el servidor finalizada o perdida.", "system");
            finishExecution();
            eventSource.close();
        };
    });

    // Función auxiliar para añadir logs a la consola
    function appendLog(text, type = "info") {
        const line = document.createElement("div");
        line.className = `log-line ${type}`;
        line.textContent = text;
        consoleLog.appendChild(line);
        
        // Auto-scroll al fondo de la consola
        consoleLog.scrollTop = consoleLog.scrollHeight;
    }

    // Restablecer estado de la interfaz tras ejecución
    function finishExecution() {
        isRunning = false;
        btnRun.disabled = false;
        btnRun.innerHTML = `<span class="run-icon">▶</span> Ejecutar Auditoría`;
        
        btnBrowseFile.disabled = false;
        btnBrowseFolder.disabled = false;
        btnWin.disabled = false;
        btnLin.disabled = false;
    }
});
