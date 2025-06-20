import threading
from pywinauto import Application
import time
from pywinauto.keyboard import send_keys
from pywinauto import Desktop

class SAPAutomation:
    def __init__(self):
        self._stop_thread = False
        self.titulo=None
    def saltar_alerta_log_sap(self):
        while not self._stop_thread:
            try: 
                # Conectar a la ventana con el título "SAP Logon"
                app = Application(backend="uia").connect(title="SAP Logon")
                # Obtener la ventana
                dlg = app.window(title="SAP Logon")
                # Restaurar la ventana si está minimizada
                dlg.restore()
                # Intentamos encontrar el botón "OK" y presionar Enter (o hacer clic en el botón)
                dlg.child_window(title="OK", control_type="Button").click_input()  # Clic en el botón "OK"
                print("Botón 'OK' presionado")
            except Exception:
                pass

    def iniciar_hilo(self):
        self._stop_thread = False
        threading.Thread(target=self.saltar_alerta_log_sap, daemon=True).start()
    def iniciar_hiloZPermitir(self):
        self._stop_thread = False
        threading.Thread(target=self.saltarPermitir, daemon=False).start()
    def getTitulo(self, titulo:str):
        self.titulo=titulo
    def saltarPermitir(self):
        # Accede a la ventana del informe
        while not self._stop_thread:
            try:
                # Conexión con la ventana de la aplicación
                app = Application(backend='uia').connect(title_re="Creación de estimados Rolling")
                dialog = app.window(title_re="Creación de estimados Rolling")
                 # Esperar a que la ventana sea visible

                # Buscar la ventana del Data Browser
                windows = Desktop(backend="uia").windows()
                for win in windows:
                    titulo = win.window_text()
                    if "Data Browser:" in titulo:
                        self.titulo = titulo
                        break  # Detener búsqueda al encontrar la ventana

                if self.titulo:  # Si se ha encontrado un título válido
                    # Conectar con la ventana completa
                    app = Application(backend='uia').connect(title=self.titulo)
                    dialog = app.window(title=self.titulo)
                    dialog.set_focus()

                    # Realizar la acción en el control
                    pane = dialog.child_window(title="Seguridad SAP GUI", control_type="Pane")
                    button = pane.child_window(title="Permitir", control_type="Button")
                    button.set_focus()
                    send_keys('{ENTER}')
                    print("Botón 'Permitir' presionado exitosamente.")

            except Exception as e:
                a=0
            time.sleep(2)  # Pausar un momento antes de continuar el ciclo
            

    def detener(self):
        self._stop_thread = True

