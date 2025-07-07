import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from datetime import datetime
import os
import pandas as pd
import time
import os
import shutil
import openpyxl


class RutaSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("Creación de estimados Rolling")
        self.root.state('zoomed')
        self.root.configure(bg="#f9f9f9")

        self.rutas = {}
        self.valores = {}

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

        # Crear cada campo manualmente, sin bucles
        self.crear_campo_archivo(frame_archivos, "carpeta: ", "Carpeta Resultado")
        self.crear_campo_archivo(frame_archivos, "RutaArchivoGlobal", "Archivo Global")
        self.crear_campo_archivo(frame_archivos, "RutaArchivoCDL", "Archivo CDL")
        self.crear_campo_archivo(frame_archivos, "RutaArchivoCrecimientos", "Archivo Crecimientos")
        self.crear_campo_archivo(frame_archivos, "RutaArchivoVentaHistorica", "Venta Histórica")
        self.crear_campo_archivo(frame_archivos, "RutaMacrosNovoApp", "Macros NovoApp")
        self.crear_campo_archivo(frame_archivos, "RutaMacrosRolling", "Macros Rolling Forecast")

        # ---------- Subframe variables ----------
        frame_variables = ttk.Labelframe(
            frame_contenido,
            text="Otros inputs",
            bootstyle="info",
            padding=10
        )
        frame_variables.pack(fill="x", pady=10)

        # Crear cada campo de texto manualmente
        self.crear_campo_texto(frame_variables, "InicioRollingCORP", "Campaña Inicio Rolling corporativo")
        self.crear_campo_texto(frame_variables, "InicioRollingPR", "Campaña Inicio Rolling Puerto Rico")
        self.crear_campo_texto(frame_variables, "AñoFinRolling", "Año Fin creación de estimados")
        self.crear_campo_texto(frame_variables, "TipoEstimado", "Tipo Estimado")

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
        if nombre != "carpeta: ":
            ruta = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
            if ruta:
                self.rutas[nombre]["ruta"] = ruta
                self.rutas[nombre]["label"].config(text=ruta)
        else:
            ruta = filedialog.askdirectory(title="Selecciona una carpeta", parent=self.root)
            if ruta:
                self.rutas[nombre]["ruta"] = ruta
                self.rutas[nombre]["label"].config(text=ruta)

    def mostrar_valores(self):
        # Validar rutas de archivo
        for clave, valor in self.rutas.items():
            ruta = valor["ruta"]

            if not ruta:
                messagebox.showerror("Error", f"Por favor, seleccione el archivo {clave}.")
                return

            setattr(self, clave, ruta)  # Asigna ruta a self.RutaArchivo...

        # Validar entradas de texto
        for clave, entrada in self.valores.items():
            valor = entrada.get()
            if not valor:
                messagebox.showerror("Error", f"Por favor, ingrese el valor de {clave}.")
                return

            if clave in ["InicioRollingCORP", "InicioRollingPR", "AñoFinRolling"]:
                try:
                    valor = int(valor)
                except ValueError:
                    messagebox.showerror("Error", f"El valor de {clave} debe ser un número.")
                    return

            setattr(self, clave, valor)  # Asigna valor a self.InicioRolling...

        # Aquí continuarías con tu lógica para realizar los cálculos y guardar el archivo
        
        self.realizar_calculos()
        messagebox.showinfo("Éxito", "Cálculo finalizado.")
        
    
    
    def datos(self,PR:bool, Carpeta:str, NombreCDL:str, InicioRollingCORP:int, InicioRollingPR:int,AñoFinRolling:int, TipoEstimado:str, claseDatos, DireccionMacrosNovoApp:str, DireccionMacrosRolling:str, categoria:int ):
        from novoApp import novoApp
        from Tendencia import Tendencia
        from Plinea import Plinea
        
        #Novoaap
        LeerNovo=novoApp(Carpeta=Carpeta,PR=PR, NombreCDL=NombreCDL,
                        InicioRollingCORP=InicioRollingCORP, InicioRollingPR=InicioRollingPR, 
                        AñoFinRolling=AñoFinRolling, claseDatos=claseDatos, DireccionMacrosNovoApp=DireccionMacrosNovoApp, categoria=categoria)
        LeerNovo.LimpiarData()
        LeerNovo.ejecutarMacros()
        resultadoSAPNOVO= LeerNovo.getSAPResultado()
        #Tendencia
        CalculoTendencia= Tendencia(carpeta=Carpeta,CampañaInicioPR=InicioRollingPR,
                                    CampañaInicioCORP=InicioRollingCORP,PR=PR,
                                    TipoEstimado=TipoEstimado,añoFinRolling=AñoFinRolling,
                                    claseDatos=claseDatos,DireccionMacrosRolling=DireccionMacrosRolling,
                                    categoria=categoria)
        CalculoTendencia.mostrarGraficaTendencia()
        
        print("Inicio linea")
        #Línea
        archivoCrecimientos=LeerNovo.df_Crecimientos
        CDL=claseDatos.getCDL()
        CDL.rename(columns={"CDP": "Centro"}, inplace=True)
        df_NovoApp=LeerNovo.df_NovoApp
        df_Horizonte= claseDatos.getHorizonte().copy(deep=True)
        LineaCorrida= Plinea(Carpeta=Carpeta,inicioRollingCORP=InicioRollingCORP, 
                            inicioRollingPR=InicioRollingPR, añoFinRolling=AñoFinRolling, 
                            PR=PR, NombreCDL= NombreCDL, DireccionMacrosRolling=DireccionMacrosRolling, tipoEstimado= TipoEstimado, categoria=categoria)
        LineaCorrida.diferencia(df_diferencia=CalculoTendencia.calculoUnidadesLinea())
        LineaCorrida.pandasAnteriores(archivoCrecimientos,CDL,df_Horizonte,df_NovoApp)
        resultadoSAPLINEA= LineaCorrida.getSAPResultado()
        #df_resultado = pd.concat([resultadoSAPNOVO, resultadoSAPLINEA], ignore_index=True)
        return None

    def leerDatos(self,carpeta:str, CI:str, CF:str, GLOBAL:str ):
        from DescargaTablas import lecturaInputs
        # Eliminar archivos de la carpeta y cerrar excel abiertos
        os.system("taskkill /f /im excel.exe")
        time.sleep(5)  # Espera 5 segundos antes de iniciar
        for archivo in os.listdir(carpeta):
            ruta_completa = os.path.join(carpeta, archivo)
            if os.path.isfile(ruta_completa):
                os.remove(ruta_completa)
            if os.path.isdir(ruta_completa):
                shutil.rmtree(ruta_completa)  # Elimina la subcarpeta
        # Descargar archivos SAP
        leerTablas=lecturaInputs(carpeta)
        leerTablas.conectarSAP()
        leerTablas.descargaNOVOAPP(campañaInicio=CI, campañaFin=CF)
        leerTablas.archivoGlobal(GLOBAL)
        return leerTablas
    
    def realizar_calculos(self):
        try:
            # Obtener las rutas desde los campos de archivo
            carpeta_resultado = self.rutas["carpeta: "]["ruta"].replace('/', '\\')
            archivo_global = self.rutas["RutaArchivoGlobal"]["ruta"].replace('/', '\\')
            archivo_cdl = self.rutas["RutaArchivoCDL"]["ruta"].replace('/', '\\')
            archivo_crecimientos = self.rutas["RutaArchivoCrecimientos"]["ruta"].replace('/', '\\')
            archivo_venta_historica = self.rutas["RutaArchivoVentaHistorica"]["ruta"].replace('/', '\\')
            macros_novoapp = self.rutas["RutaMacrosNovoApp"]["ruta"].replace('/', '\\')
            macros_rolling = self.rutas["RutaMacrosRolling"]["ruta"].replace('/', '\\')

            # Obtener las variables de entrada
            tipo_estimado = self.valores["TipoEstimado"].get()
            if(tipo_estimado=="SAP"):
                tipo_estimado = "Supply Original"
            inicio_rolling_corp = self.valores["InicioRollingCORP"].get()
            inicio_rolling_pr = self.valores["InicioRollingPR"].get()
            año_fin_rolling = self.valores["AñoFinRolling"].get()

            # Validar y convertir a números
            try:
                inicio_rolling_corpInt = int(inicio_rolling_corp)
                inicio_rolling_pr = int(inicio_rolling_pr)
                año_fin_rolling = int(año_fin_rolling[:4])
            except ValueError:
                messagebox.showerror("Error", "Las fechas deben ser números válidos.")
                return

            año_siguiente = datetime.now().year + 1
            inicio_rolling_corp = f"{año_siguiente}01"  # Resultado: '202601'
            campaña_fin = str(año_fin_rolling) 

            # Leer datos
            lecturaDatos = self.leerDatos(
                carpeta=carpeta_resultado,
                CI=inicio_rolling_corp,
                CF=campaña_fin,
                GLOBAL=archivo_global
            )

            lecturaDatos.leerOtrosInputs(
                RutaCDL=archivo_cdl,
                RutaArchivoCrecimiento=archivo_crecimientos,
                RutaHistorico=archivo_venta_historica
            )
            
            Categorias=[101, 102, 103, 104, 105, 106]
            for categoria in Categorias:
                # Cerrar Excel
                os.system("taskkill /f /im excel.exe")

                # Cálculo CORP
                resultadoCORP = self.datos(
                    PR=False,
                    InicioRollingPR=inicio_rolling_pr,
                    AñoFinRolling=año_fin_rolling,
                    Carpeta=carpeta_resultado + "\\",
                    NombreCDL=archivo_cdl,
                    InicioRollingCORP=inicio_rolling_corpInt,
                    TipoEstimado=tipo_estimado,
                    claseDatos=lecturaDatos,
                    DireccionMacrosNovoApp=macros_novoapp,
                    DireccionMacrosRolling=macros_rolling,
                    categoria=categoria
                )
                print("Corrida con éxito CORP")

        # Cálculo PR03
            resultadoPR03 = self.datos(
                PR=True,
                InicioRollingPR=inicio_rolling_pr,
                AñoFinRolling=año_fin_rolling,
                Carpeta=carpeta_resultado + "\\",
                NombreCDL=archivo_cdl,
                InicioRollingCORP=inicio_rolling_corpInt,
                TipoEstimado=tipo_estimado,
                claseDatos=lecturaDatos,
                DireccionMacrosNovoApp=macros_novoapp,
                DireccionMacrosRolling=macros_rolling,
                categoria="PR"
            )
            print("Corrida con éxito PR03")

            messagebox.showinfo("Éxito", f"Cálculo estimados finalizado.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante el cálculo: {str(e)}")
            print(f"Error: {str(e)}")
            
# Ejecutar la interfaz
if __name__ == "__main__":
    root = ttk.Window(themename="minty")
    app = RutaSelector(root)
    root.mainloop()
