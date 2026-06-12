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
    const btnStop = document.getElementById("btn-stop");
    const btnExit = document.getElementById("btn-exit");
    const consoleArea = document.getElementById("console-area");
    const consoleLog = document.getElementById("console-log");
    const btnClearConsole = document.getElementById("btn-clear-console");

    let selectedOS = "windows";
    let isRunning = false;
    let eventSource = null;

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
        btnRun.innerHTML = `<span class="run-icon">⏳</span> Ejecutando...`;
        btnStop.style.display = "flex";
        consoleArea.style.display = "block";
        consoleLog.innerHTML = ""; // Limpiar consola
        appendLog(`Iniciando auditoría en segundo plano (OS: ${selectedOS})...`, "system");

        // Desactivar controles
        btnBrowseFile.disabled = true;
        btnBrowseFolder.disabled = true;
        btnWin.disabled = true;
        btnLin.disabled = true;
        btnExit.disabled = true;

        // Crear la URL con los parámetros correspondientes
        const params = new URLSearchParams({
            plantilla: plantilla,
            trabajos: trabajos,
            os: selectedOS
        });

        // Conectar al endpoint SSE de ejecución
        eventSource = new EventSource(`/api/run-audit?${params.toString()}`);

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

    // 7. Botón Cancelar/Detener Auditoría
    btnStop.addEventListener("click", () => {
        if (!isRunning || !eventSource) return;
        
        appendLog("\n⌛ Deteniendo auditoría...", "warning");
        btnStop.disabled = true;
        btnStop.textContent = "Deteniendo...";
        
        // Cerrar flujo SSE del lado del cliente
        eventSource.close();
        
        // Solicitar al backend detener el subproceso
        fetch("/api/abort")
            .then(res => res.json())
            .then(data => {
                appendLog("✅ Auditoría detenida correctamente por el usuario.", "system");
            })
            .catch(err => {
                console.error("Error al abortar auditoría:", err);
                appendLog("⚠️ El servidor no respondió a la cancelación, pero la conexión local se cerró.", "warning");
            })
            .finally(() => {
                btnStop.disabled = false;
                btnStop.innerHTML = "<span>■</span> Cancelar";
                finishExecution();
            });
    });

    // 8. Botón Apagar Servidor y Salir
    btnExit.addEventListener("click", () => {
        if (isRunning) return;

        if (confirm("¿Estás seguro de que deseas cerrar la aplicación?\nEsto apagará el servidor y cerrará la ventana de CMD de inmediato.")) {
            // Reemplazar cuerpo con pantalla elegante de apagado
            document.body.innerHTML = `
                <div class="mesh-background"></div>
                <div class="app-container" style="text-align: center; max-width: 500px; padding: 40px 20px; margin: auto; height: 100vh; display: flex; align-items: center; justify-content: center;">
                    <div style="background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.12); border-radius: 24px; padding: 40px; backdrop-filter: blur(16px); width: 100%; box-shadow: 0 20px 40px rgba(0,0,0,0.3);">
                        <div style="font-size: 3.5rem; margin-bottom: 20px; filter: drop-shadow(0 0 10px rgba(239, 68, 68, 0.4));">🔌</div>
                        <h1 style="font-size: 1.8rem; margin-bottom: 15px; font-weight: 700; color: #ffffff;">Aplicación Cerrada</h1>
                        <p style="color: #94a3b8; font-size: 0.95rem; line-height: 1.6; margin-bottom: 25px;">El servidor local ha sido desconectado. La ventana de la consola (CMD) se cerrará automáticamente.</p>
                        <p style="color: #64748b; font-size: 0.8rem; border-top: 1px solid rgba(255,255,255,0.05); padding-top: 20px;">Ya puedes cerrar esta pestaña del navegador de forma segura.</p>
                    </div>
                </div>
            `;
            
            // Notificar al backend para apagar
            fetch("/api/shutdown").catch(err => {
                console.log("Conexión con el servidor cerrada por apagado.");
            });
        }
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
        btnStop.style.display = "none";
        
        btnBrowseFile.disabled = false;
        btnBrowseFolder.disabled = false;
        btnWin.disabled = false;
        btnLin.disabled = false;
        btnExit.disabled = false;
        eventSource = null;
    }
});
