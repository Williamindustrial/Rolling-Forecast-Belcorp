import threading
from pywinauto import Application
import time
import threading

class SAPAutomation:
    def __init__(self):
        # Inicializa cualquier otro atributo si es necesario
        pass

    def saltarAlertaLogSAP(self):
        Finalizo=False
        while(Finalizo==False):
            try:
                # Conectar a la ventana con el título "SAP Logon"
                app = Application(backend="uia").connect(title="SAP Logon")

                # Obtener la ventana
                dlg = app.window(title="SAP Logon")

                # Verificar que estamos encontrando la ventana correctamente
                print(f"Ventana encontrada: {dlg.window_text()}")

                # Restaurar la ventana si está minimizada
                dlg.restore()

                # Intentamos encontrar el botón "OK" y presionar Enter (o hacer clic en el botón)
                dlg.child_window(title="OK", control_type="Button").click_input()  # Clic en el botón "OK"
                print("Botón 'OK' presionado")
                Finalizo=True

            except Exception as e:
                print(f"No se pudo presionar el botón OK: {e}")
            time.sleep(0.0001)

    def iniciar_hilo(self):
        # Crear un nuevo hilo para ejecutar saltarAlertaLogSAP
        hilo = threading.Thread(target=self.saltarAlertaLogSAP)
        hilo.start()  # Iniciar el hilo

# Crear una instancia de la clase y llamar a la función desde otro hilo
automation = SAPAutomation()
automation.iniciar_hilo()
