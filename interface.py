import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox


class RutaSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("Creación de estimados Rolling")
        self.root.state('zoomed')
        self.root.configure(bg="#f9f9f9")

        self.rutas = {}
        self.valores = {}

        # Diccionario de nombres amigables
        self.nombres_amigables = {
            "RutaArchivoGlobal": "Archivo Global",
            "RutaArchivoCDL": "Archivo CDL",
            "RutaArchivoCrecimientos": "Archivo Crecimientos",
            "RutaArchivoVentaHistorica": "Venta Histórica",
            "RutaMacrosNovoApp": "Macros NovoApp",
            "RutaMacrosRolling": "Macros Rolling Forecast"
        }

        # ---------- Título principal ----------
        titulo = ttk.Label(
            self.root,
            text="Creación de estimados Rolling",
            foreground="#800080",  # Morado
            background="#f9f9f9",
            font=("Segoe UI", 30, "bold")
        )
        titulo.pack(pady=20)

        # ---------- Frame principal ----------
        frame_contenido = ttk.Frame(self.root, bootstyle="light")
        frame_contenido.pack(pady=10, padx=20, fill="both", expand=True)

        # ---------- Subframe archivos ----------
        frame_archivos = ttk.Labelframe(
            frame_contenido,
            text="Archivos",
            bootstyle="primary",
            padding=10
        )
        frame_archivos.pack(fill="x", pady=10)

        campos_archivos = list(self.nombres_amigables.items())
        for clave, etiqueta in campos_archivos:
            self.crear_campo_archivo(frame_archivos, clave, etiqueta)

        # ---------- Subframe variables ----------
        frame_variables = ttk.Labelframe(
            frame_contenido,
            text="Otros inputs",
            bootstyle="info",
            padding=10
        )
        frame_variables.pack(fill="x", pady=10)

        campos_texto = [
            ("InicioRollingCORP", "Inicio Rolling corporativo"),
            ("InicioRollingPR", "Inicio Rolling Puerto Rico"),
            ("AñoFinRolling", "Año Fin creación de estimados"),
            ("TipoEstimado", "Tipo Estimado")
        ]

        for clave, etiqueta in campos_texto:
            self.crear_campo_texto(frame_variables, clave, etiqueta)

        # ---------- Botón principal ----------
        btn_guardar = ttk.Button(
            self.root,
            text="Calcular",
            bootstyle="purple-outline",
            command=self.mostrar_valores,
            width=30,
            padding=(20, 15)  # ancho, alto
        )
        btn_guardar.pack(pady=20)

    def crear_campo_archivo(self, parent, nombre, texto):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=5)

        ttk.Label(frame, text=texto + ":", font=("Segoe UI", 14)).pack(side="left", padx=10)

        btn = ttk.Button(
            frame,
            text="Buscar",
            bootstyle="purple-outline",
            command=lambda: self.seleccionar_archivo(nombre)
        )
        btn.pack(side="left", padx=10)

        lbl = ttk.Label(frame, text="No seleccionado", width=100, anchor="w")
        lbl.pack(side="left", padx=10)

        self.rutas[nombre] = {"ruta": "", "label": lbl}

    def crear_campo_texto(self, parent, nombre, texto):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=5)

        ttk.Label(frame, text=texto + ":", font=("Segoe UI", 14)).pack(side="left", padx=10)

        if nombre == "TipoEstimado":
            combobox = ttk.Combobox(
                frame,
                values=["Planit", "SAP", "Planit sin TO122", "TO122", "Demanda AA"],
                width=37,
                state="readonly",
                bootstyle="purple"
            )
            combobox.current(0)
            combobox.pack(side="left", padx=10)
            self.valores[nombre] = combobox
        else:
            entrada = ttk.Entry(frame, width=40)
            entrada.pack(side="left", padx=10)
            self.valores[nombre] = entrada

    def seleccionar_archivo(self, nombre):
        ruta = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if ruta:
            self.rutas[nombre]["ruta"] = ruta
            self.rutas[nombre]["label"].config(text=ruta)

    def mostrar_valores(self):
        for clave, valor in self.rutas.items():
            nombre_amigable = self.nombres_amigables.get(clave, clave)

            if not valor["ruta"]:
                messagebox.showerror("Error", f"Por favor, seleccione el {nombre_amigable}.")
                return

            if clave in ["RutaMacrosNovoApp", "RutaMacrosRolling"]:
                if not valor["ruta"].endswith(".xlsm"):
                    messagebox.showerror("Error", f"El archivo {nombre_amigable} debe ser un archivo de macros (.xlsm).")
                    return
            else:
                if not (valor["ruta"].endswith(".xlsx") or valor["ruta"].endswith(".xls")):
                    messagebox.showerror("Error", f"El archivo {nombre_amigable} debe ser un archivo Excel (.xlsx o .xls).")
                    return

        print("---- Rutas de archivos ----")
        for clave, valor in self.rutas.items():
            print(f"{clave} = {valor['ruta']}")

        print("\n---- Valores de entrada ----")
        for clave, entrada in self.valores.items():
            print(f"{clave} = {entrada.get()}")


# ---------- Ejecutar interfaz ----------
if __name__ == "__main__":
    root = ttk.Window(themename="minty")
    app = RutaSelector(root)
    root.mainloop()
