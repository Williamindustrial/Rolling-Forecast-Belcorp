# -*- coding: utf-8 -*-
"""
Created on Wed Apr 23 15:03:54 2025

@author: williamtorres
"""
import pandas as pd
from automatizacionVentanas import SAPAutomation 
import threading
from pywinauto import Desktop
from datetime import datetime
import win32com.client
import pyperclip
import time
import tkinter as tk
from tkinter import messagebox
from pywinauto import Application


class lecturaInputs:
    
    df_NovoApp= None
    df_ZMM206NovoApp=None
    CDL=None
    df_Horizonte=None
    df_ZMM206Linea=None
    df_ZMM206kLinea=None
    df_Crecimientos=None
    df_CrecimientosPaís=None
    df_Historico=None
    
    def __init__(self, direcciónResultado:str):
        self.session=None
        self.direcciónResultado=direcciónResultado
    # Conectar SAP
    
    def conectarSAP(self):
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            sap_helper = SAPAutomation()
            sap_helper.iniciar_hilo()  # Inicia el hilo de alerta
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)  # Conexión activa
            session = connection.Children(0)  # Primera sesión activa
            sap_helper.detener()  # Detiene el hilo después de conectarse
            self.session=session
            return session
        except Exception as e:
            print(e)
            self.mostrar_mensaje_error(f"No esta logeado en SAP")
            return False
    def mostrar_mensaje_error(self, mensaje):
        """Muestra un mensaje de error en una ventana emergente."""
        messagebox.showerror("Error", mensaje)
    
    def saltarAlertaLogSAP(self):
        self.hilo = SAPAutomation()
                
    def descargaNOVOAPP(self, campañaInicio:str, campañaFin:str):
        sap_helper = SAPAutomation()
        sap_helper.iniciar_hiloZPermitir() 
        session= self.session
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZTPPFOTOPLANIT"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 14
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtI2-LOW").text = campañaInicio
        session.findById("wnd[0]/usr/txtI2-HIGH").text = campañaFin
        session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.FindById("wnd[0]").SendVKey(43)
        session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.direcciónResultado
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "NOVOAPP.XLSX"
        time.sleep(1)  # Esperar 2 segundos antes de enviar la tecla
        # Enviar tecla F4 (código 4) en la ventana secundaria
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.leerNovoAPP(Carpeta=self.direcciónResultado)
        time.sleep(1) 
        sap_helper.detener() 
    def descargaZMM206K(self, nombreArchivo:str):
        sap_helper = SAPAutomation()
        sap_helper.iniciar_hiloZPermitir() 
        session= self.session
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n zmm206"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtP_STATM").text = "k"
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(29, "TEXT")
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 23
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "29"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "TEXT_MARA_EXTWG"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.direcciónResultado
        sap_helper.getTitulo("Reporte de Textos Breves del Material y Datos Basicos")
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombreArchivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(1) 
        sap_helper.detener() 
        
    def descargaZMM206D(self):
        sap_helper = SAPAutomation()
        sap_helper.iniciar_hiloZPermitir() 
        session= self.session
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "Marc"
        session.findById("wnd[0]").sendVKey(0)
        # Presionar el botón para ingresar los valores del filtro
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        time.sleep(1)
        # Presionar el botón de la ventana emergente (por ejemplo, el filtro)
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(1)
        # Establecer el filtro de búsqueda (ejemplo: "*3")
        session.findById("wnd[0]/usr/ctxtI2-LOW").text = "*3"
        # Borrar la selección máxima
        session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        # Aplicar el filtro
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(1)
        # Establecer el foco en la etiqueta específica y presionar VKey
        session.findById("wnd[0]/mbar/menu[3]/menu[1]").select()
        session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radRSEUMOD-TBALV_STAN").select()
        session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").select()
        session.findById("wnd[1]/usr/tabsG_TABSTRIP/tabp0400/ssubTOOLAREA:SAPLWB_CUSTOMIZING:0400/radSEUCUSTOM-FIELDNAME").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/lbl[37,9]").setFocus()
        session.findById("wnd[0]/usr/lbl[37,9]").caretPosition = 0
        session.findById("wnd[0]").sendVKey(20)
        # Confirmar la operación de exportación
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        # Establecer el nombre y la ubicación del archivo de exportación
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.direcciónResultado
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MARC.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 3
        # Confirmar la exportación
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        # Esperar unos segundos para asegurarse de que el archivo se guarda
        time.sleep(3)
        sap_helper.detener() 
        print("Exportación completada exitosamente.")
        
    def pegarData(self, listaCodigos):
        texto_a_pegar=""
        for i in range(len(listaCodigos)):
            texto_a_pegar= texto_a_pegar+"\r\n"+str(listaCodigos[i])
        #texto_a_pegar = "\r\n".join(listaCodigos)  # "\r\n" es un salto de línea en SAP
        # Copiar al portapapeles
        pyperclip.copy(texto_a_pegar)
        
    def zmm206Linea(self, carpeta):
        archivoZMM206 = carpeta+"ZMM206 LINEA.XLSX"
        lecturaInputs.df_ZMM206Linea= pd.read_excel(archivoZMM206, sheet_name='Sheet1')
    
    def Historico(self, direccion):
        lecturaInputs.df_Historico= pd.read_excel(direccion, sheet_name='DEMANDA HISTORICA', skiprows=4)
    
    def zmm206kLinea(self, Carpeta):
        archivoZMM206k = Carpeta+"ZMM206 LINEA-K.XLSX"
        lecturaInputs.df_ZMM206kLinea= pd.read_excel(archivoZMM206k, sheet_name='Sheet1')
        
    def leerNovoAPP(self, Carpeta):
        archivoNovoapp = Carpeta+"\\NOVOAPP.XLSX"
        lecturaInputs.df_NovoApp= pd.read_excel(archivoNovoapp, sheet_name='Sheet1')
        columnas = ['MANDT', 'COMWERKS', 'COMCAM', 'COMPROD', 'TIPOOFERTA', 'COMUEST',
       'VTAPROY', 'FUENTE', 'UDFTIME', 'ERSDA', 'ERNAM', 'UDLTIME', 'LAEDA',
       'AENAM', 'PROGRAMM']
        lecturaInputs.df_NovoApp.columns = columnas
        codigos=lecturaInputs.df_NovoApp['COMPROD'].unique().tolist()
        self.pegarData(codigos)
        self.descargaZMM206K("ZMM206NOVOAPP.XLSX")
    
    def LeerZmm206NovoAPP(self, Carpeta):
        archivoZMM206NOVO = Carpeta+"ZMM206NOVOAPP.XLSX"
        lecturaInputs.df_ZMM206NovoApp= pd.read_excel(archivoZMM206NOVO, sheet_name='Sheet1')
        
    # Lectura de archivos
    
    def archivoCrecimiento(self, Carpeta):
        archivoCrecimientos = Carpeta
        lecturaInputs.df_Crecimientos= pd.read_excel(archivoCrecimientos, sheet_name='Crecimiento')
    
    def archivoCrecimientoUU(self, Carpeta):
        archivoCrecimientos = Carpeta
        lecturaInputs.df_CrecimientosPaís= pd.read_excel(archivoCrecimientos, sheet_name='UUPaís')
        
    def archivoGlobal(self, archivoDireccion):
        archivoHorizonte = archivoDireccion
        lecturaInputs.df_Horizonte= pd.read_excel(archivoHorizonte, sheet_name='Horizonte')
        lecturaInputs.df_Horizonte['Año'] = lecturaInputs.df_Horizonte['Período'].str[:4].astype(int)
        # Crear la columna 'Campaña'
        lecturaInputs.df_Horizonte['Campaña'] = lecturaInputs.df_Horizonte['Período'].str[-2:].astype(int)
        codigos=lecturaInputs.df_Horizonte['SAP'].unique().tolist()
        self.pegarData(codigos)
        self.descargaZMM206D()
        self.descargaZMM206K("ZMM206 LINEA-K.XLSX")
        
    def archivoCDL(self, CDL):
        añoActual = datetime.now().year
        CDL_CONCATENADO= None
        for i in range(2022, añoActual+2):
            CDL_AÑOACTUAL= pd.read_excel(CDL, sheet_name=str(i))
            # Concatenar los dos DataFrames
            if(i>2022):
                CDL_CONCATENADO = pd.concat([CDL_AÑOACTUAL, CDL_CONCATENADO], ignore_index=True)
            else:
                CDL_CONCATENADO=CDL_AÑOACTUAL
    
        CDL = CDL_CONCATENADO.groupby(["CodigoSAP","Pais"], as_index=False)["CampaniaDescontinuacion"].max()
        data = [
                ('CO03', 'COLOMBIA'),
                ('PE03', 'PERU'),
                ('MX03', 'MEXICO'),
                ('CL03', 'CHILE'),
                ('GT23', 'GUATEMALA'),
                ('SV13', 'EL SALVADOR'),
                ('CR03', 'COSTA RICA'),
                ('DO03', 'REPUBLICA DOMINICANA'),
                ('PA33', 'PANAMA'),
                ('BO03', 'BOLIVIA'),
                ('EC03', 'ECUADOR'),
                ('PR03', 'PUERTO RICO'),
                ('US03', 'USA')
            ]
        
        # Crear el DataFrame
        CDP = pd.DataFrame(data, columns=['CDP', 'Pais'])
        lecturaInputs.CDL = pd.merge(CDL, CDP, left_on="Pais", right_on="Pais", how="left")
    
    def leerOtrosInputs(self, RutaCDL:str, RutaArchivoCrecimiento:str, RutaHistorico:str):
        self.archivoCDL(RutaCDL)
        self.archivoCrecimientoUU(RutaArchivoCrecimiento)
        self.archivoCrecimiento(RutaArchivoCrecimiento)
        self.Historico(RutaHistorico)
        
    def getHorizonte(self):
        return lecturaInputs.df_Horizonte.copy(deep=True)
    
    def getCDL(self):
        return lecturaInputs.CDL.copy(deep=True)
    
    def get_df_Crecimientos(self):
        return lecturaInputs.df_Crecimientos.copy(deep=True)
    
    def get_df_CrecimientosPaís(self):
        return lecturaInputs.df_CrecimientosPaís.copy(deep=True)
    
    def get_VentaHistorica(self):
        return lecturaInputs.df_Historico.copy(deep=True)
    