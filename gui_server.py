# ============================================================================
# gui_server.py — Servidor Local para la Interfaz Gráfica de APP_REVISA_EXCEL
# ============================================================================
"""
Servidor local en Python puro (sin dependencias externas) que:
  1. Sirve la interfaz web moderna (HTML, CSS, JS) en el puerto 5000.
  2. Expone APIs para abrir diálogos de selección de archivos y carpetas locales.
  3. Ejecuta la auditoría en segundo plano transmitiendo la salida en tiempo real.
  4. Escucha en 0.0.0.0 permitiendo controlar la auditoría desde dispositivos móviles.
"""

import os
import sys
import json
import urllib.parse
import subprocess
import threading
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler

# Directorio raíz del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PLANTILLA = os.path.abspath(os.path.join(BASE_DIR, "PLANTILLA.xlsx"))
DEFAULT_TRABAJOS = os.path.abspath(os.path.join(BASE_DIR, "TRABAJOS_ESTUDIANTES"))

PORT = 5000

# Variables globales para control del subproceso y la instancia del servidor
proceso_activo = None
server_instance = None


# ============================================================================
# DIÁLOGOS DE SELECCIÓN NATIVER DE FICHEROS (AISLADOS EN SUBPROCESOS)
# ============================================================================
def abrir_dialogo_archivo():
    """
    Abre una ventana de selección de archivos (PLANTILLA.xlsx) usando tkinter.
    Se ejecuta en un subproceso aislado para evitar problemas de hilos con la GUI.
    """
    script = (
        "import tkinter as tk; from tkinter import filedialog; "
        "root=tk.Tk(); root.withdraw(); root.attributes('-topmost', True); "
        "print(filedialog.askopenfilename(filetypes=[('Archivos de Excel', '*.xlsx')]))"
    )
    try:
        resultado = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=120)
        return resultado.stdout.strip()
    except Exception as e:
        print(f"Error al abrir diálogo de selección de archivo: {e}")
        return ""


def abrir_dialogo_carpeta():
    """
    Abre una ventana de selección de directorios (TRABAJOS_ESTUDIANTES) usando tkinter.
    Se ejecuta en un subproceso aislado para evitar problemas de hilos con la GUI.
    """
    script = (
        "import tkinter as tk; from tkinter import filedialog; "
        "root=tk.Tk(); root.withdraw(); root.attributes('-topmost', True); "
        "print(filedialog.askdirectory())"
    )
    try:
        resultado = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=120)
        return resultado.stdout.strip()
    except Exception as e:
        print(f"Error al abrir diálogo de selección de carpeta: {e}")
        return ""


# ============================================================================
# CONTROLADOR DE PETICIONES HTTP
# ============================================================================
class GUIHandler(BaseHTTPRequestHandler):
    
    def log_message(self, format, *args):
        # Desactivar logs estándar en la consola para no ensuciar la terminal del servidor
        pass

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        query = urllib.parse.parse_qs(parsed_url.query)

        # Rutas de Archivos Estáticos
        if path == "/":
            self.servir_archivo(os.path.join(BASE_DIR, "gui", "index.html"), "text/html")
        elif path == "/style.css":
            self.servir_archivo(os.path.join(BASE_DIR, "gui", "style.css"), "text/css")
        elif path == "/app.js":
            self.servir_archivo(os.path.join(BASE_DIR, "gui", "app.js"), "application/javascript")
            
        # Endpoints API
        elif path == "/api/defaults":
            self.enviar_json({
                "plantilla": DEFAULT_PLANTILLA,
                "trabajos": DEFAULT_TRABAJOS
            })
        elif path == "/api/select-file":
            ruta_archivo = abrir_dialogo_archivo()
            self.enviar_json({"path": ruta_archivo})
        elif path == "/api/select-folder":
            ruta_carpeta = abrir_dialogo_carpeta()
            self.enviar_json({"path": ruta_carpeta})
        elif path == "/api/run-audit":
            self.ejecutar_auditoria_sse(query)
        elif path == "/api/abort":
            self.abortar_auditoria()
        elif path == "/api/shutdown":
            self.apagar_servidor()
        else:
            self.send_error(404, "Recurso no encontrado")

    def servir_archivo(self, ruta_archivo, content_type):
        """Lee y sirve un archivo físico local con su respectivo tipo MIME."""
        try:
            with open(ruta_archivo, "rb") as f:
                contenido = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(contenido)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(contenido)
        except Exception as e:
            self.send_error(500, f"Error del Servidor: {e}")

    def enviar_json(self, data):
        """Serializa y responde un objeto JSON."""
        try:
            contenido = json.dumps(data).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(contenido)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(contenido)
        except Exception as e:
            self.send_error(500, f"Error al enviar JSON: {e}")

    def ejecutar_auditoria_sse(self, query):
        """
        Ejecuta la auditoría llamando a main.py con los argumentos seleccionados
        y transmite la salida estándar de consola en tiempo real usando Server-Sent Events (SSE).
        """
        plantilla = query.get("plantilla", [""])[0]
        trabajos = query.get("trabajos", [""])[0]
        os_selected = query.get("os", ["windows"])[0]

        if not plantilla or not trabajos:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"data: " + json.dumps({"status": "error", "message": "Faltan parámetros obligatorios"}).encode("utf-8") + b"\n\n")
            return

        # Configurar ejecutable Python según sistema seleccionado
        if os_selected.lower() == "windows":
            python_bin = os.path.join(BASE_DIR, ".venv", "Scripts", "python.exe")
            if not os.path.exists(python_bin):
                python_bin = "python"
        else:
            python_bin = os.path.join(BASE_DIR, ".venv", "bin", "python")
            if not os.path.exists(python_bin):
                python_bin = "python3"

        # Comando para ejecutar con salida sin búfer (-u)
        cmd = [python_bin, "-u", "main.py", "--plantilla", plantilla, "--trabajos", trabajos]

        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('Connection', 'keep-alive')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()

        # Iniciar subproceso de auditoría redireccionando salida estándar y errores
        global proceso_activo
        try:
            proceso_activo = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                cwd=BASE_DIR,
                bufsize=1
            )
            
            # Leer la salida del subproceso línea por línea y enviarla al cliente
            for linea in proceso_activo.stdout:
                data = json.dumps({"text": linea.rstrip()})
                self.wfile.write(f"data: {data}\n\n".encode('utf-8'))
                self.wfile.flush()

            proceso_activo.wait()
            codigo_salida = proceso_activo.returncode
            data = json.dumps({"status": "done", "code": codigo_salida})
            self.wfile.write(f"data: {data}\n\n".encode('utf-8'))
            self.wfile.flush()

        except (ConnectionError, BrokenPipeError):
            print("⚠️ Cliente desconectado. Finalizando subproceso de auditoría.")
            if proceso_activo:
                try:
                    proceso_activo.terminate()
                except Exception:
                    pass
        except Exception as e:
            err_msg = json.dumps({"status": "error", "message": str(e)})
            try:
                self.wfile.write(f"data: {err_msg}\n\n".encode('utf-8'))
                self.wfile.flush()
            except Exception:
                pass
        finally:
            proceso_activo = None

    def abortar_auditoria(self):
        """Detiene de forma forzada el subproceso de auditoría activo."""
        global proceso_activo
        if proceso_activo:
            try:
                proceso_activo.terminate()
                print("Proceso de auditoría detenido por solicitud del usuario.")
            except Exception as e:
                print(f"Error al detener el proceso: {e}")
            proceso_activo = None
        self.enviar_json({"status": "aborted", "message": "Auditoría detenida con éxito."})

    def apagar_servidor(self):
        """Apaga el servidor HTTP y cierra la ventana de CMD del sistema."""
        global proceso_activo
        if proceso_activo:
            try:
                proceso_activo.terminate()
            except Exception:
                pass
            proceso_activo = None
            
        self.enviar_json({"status": "shutdown", "message": "Servidor apagado. Cerrando ventana..."})
        
        def apagar_y_salir():
            import time
            time.sleep(0.5)
            global server_instance
            if server_instance:
                server_instance.shutdown()
            os._exit(0)

        threading.Thread(target=apagar_y_salir).start()


# ============================================================================
# FUNCIÓN DE INICIO Y BINDING DEL SERVIDOR
# ============================================================================
def iniciar_servidor():
    global server_instance
    # Enlazar en 0.0.0.0 para que sea accesible desde otros dispositivos en la red local
    server_address = ('0.0.0.0', PORT)
    server_instance = HTTPServer(server_address, GUIHandler)
    
    print("======================================================================")
    print("  SERVIDOR DE INTERFAZ GRAFICA INICIADO")
    print(f"  Acceso local:             http://localhost:{PORT}")
    
    # Obtener IP local para mostrar al usuario cómo acceder desde el móvil
    try:
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip_local = s.getsockname()[0]
        s.close()
        print(f"  Acceso desde celular/aula: http://{ip_local}:{PORT}")
    except Exception:
        pass
    print("======================================================================")

    # Abrir automáticamente el navegador predeterminado en localhost
    def abrir_navegador():
        try:
            webbrowser.open(f"http://localhost:{PORT}")
        except Exception as e:
            print(f"No se pudo abrir el navegador automáticamente: {e}")

    threading.Timer(1.0, abrir_navegador).start()

    try:
        server_instance.serve_forever()
    except KeyboardInterrupt:
        print("\nCerrando servidor de interfaz gráfica...")
        server_instance.server_close()


if __name__ == "__main__":
    iniciar_servidor()
