# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 09:15:25 2025

@author: williamtorres
"""

import pandas as pd
import os
from datetime import datetime
import win32com.client as win32
from win32com.client import constants
from DescargaTablas import lecturaInputs
import time
import papermill as pm
from datetime import datetime
class novoApp:

    # Variables locales
    df_NovoApp= None
    df_ZMM206NovoApp=None
    df_Crecimientos=None
    CDL_CONCATENADO=None
    CDL=None
    df_resultadoSAP=None
    Traduciones = {
            "Código": ["101", "102", "103", "104", "105", "106"],
            "Categoría": [
                "Fragancias",
                "Maquillaje",
                "Cuidado Personal",
                "Tratamiento Facial",
                "Tratamiento Corporal",
                "Accesorios Cosméticos",
            ],
        }
    
    
    def __init__(self, Carpeta: str, PR:bool, NombreCDL:str, InicioRollingCORP: int, InicioRollingPR: int, AñoFinRolling:int, claseDatos:lecturaInputs, DireccionMacrosNovoApp:str, categoria:int):
        # Inputs
        self.Carpeta= Carpeta
        self.PR= PR
        self.ArchivoCDL=Carpeta+NombreCDL
        self.descargaTablas= claseDatos
        self.DireccionMacrosNovoApp=DireccionMacrosNovoApp
        self.categoria=categoria
        if(PR):
            self.direccionResultado= Carpeta+'Resultado PR03\\'+str(categoria)+'\\'
            # Últimos dos dígitos
            self.InicioRolling= InicioRollingPR
            ultimos_dos = str(InicioRollingPR)[-2:]  # 13
            
            # Primeros cuatro dígitos
            primeros_cuatro = int(str(InicioRollingPR)[:4])  # 2026
            if(ultimos_dos=='01'):
                self.FinGlobal= str(primeros_cuatro-1)+'13'
            else:
                self.FinGlobal=InicioRollingPR-1
            self.NumeroCampañas=13
        else:
            self.direccionResultado= Carpeta+'Resultado CORP\\'+str(categoria)+'\\'
            # Últimos dos dígitos
            self.InicioRolling= InicioRollingCORP
            ultimos_dos = str(InicioRollingCORP)[-2:]  # 13
            
            # Primeros cuatro dígitos
            primeros_cuatro = int(str(InicioRollingCORP)[:4])  # 2026
            if(ultimos_dos=='01'):
                self.FinGlobal= str(InicioRollingCORP-1)+'18'
            else:
                self.FinGlobal=InicioRollingCORP-1
            self.NumeroCampañas=18
        self.FinRollin= AñoFinRolling
            
        # Verifica si la carpeta ya existe, si no, la crea
        if not os.path.exists(self.direccionResultado):
            os.makedirs(self.direccionResultado)
        # Creación Carpeta Resultados
        
    # Aqui se lee NOVOAPP, agregar código descargar desde SAP
    def LeerNovoApp(self):
        
        novoApp.df_NovoApp= self.descargaTablas.getNovoApp()
    # Aqui se lee zmm206, agregar código descargar desde SAP
    def LeerZmm206(self):
        archivoZMM206NOVO = self.Carpeta+"ZMM206NOVOAPP.XLSX"
        df_ZMM206NovoApp= pd.read_excel(archivoZMM206NOVO, sheet_name='Sheet1')
        # Quitar el guion "-" de la columna 'Codigo'
        df_ZMM206NovoApp["Material"] = df_ZMM206NovoApp["Material"].str.replace("-", "", regex=True)
        df_ZMM206NovoApp.columns= ['Material', 'Número de material', 'TpMt', 'Grupo art.',
       'Grupo de artículos externo', 'Tipo de producto',
       'Descripcion de la Jerarquía']
        novoApp.df_ZMM206NovoApp=df_ZMM206NovoApp
        
    def archivoCDL(self):
        novoApp.CDL = self.descargaTablas.getCDL()
        
    def archivoCrecimiento(self):
        novoApp.df_Crecimientos= self.descargaTablas.get_df_Crecimientos()
        
        
    def LimpiarData(self):
        self.LeerNovoApp()
        self.LeerZmm206()
        self.archivoCDL()
        self.archivoCrecimiento()
        año_actual = datetime.now().year+1     
        año_modificado = f"{año_actual}00" 
        novoApp.df_NovoApp = pd.merge(novoApp.df_NovoApp, novoApp.CDL, left_on=["COMPROD","COMWERKS"], right_on=["CodigoSAP","CDP"], how="left")
        novoApp.df_NovoApp=novoApp.df_NovoApp[["COMWERKS",'COMPROD', 'COMCAM' , "FUENTE", "TIPOOFERTA", 'COMUEST', 'CampaniaDescontinuacion']]
        
        novoApp.df_ZMM206NovoApp= novoApp.df_ZMM206NovoApp[["Material", "Número de material", 'Grupo art.','Tipo de producto']].drop_duplicates()
        novoApp.df_ZMM206NovoApp['Material'] = novoApp.df_ZMM206NovoApp['Material'].fillna(0).astype(int)
        novoApp.df_NovoApp = pd.merge(novoApp.df_NovoApp, novoApp.df_ZMM206NovoApp, left_on="COMPROD", right_on="Material", how="left")
        novoApp.df_NovoApp=novoApp.df_NovoApp[["COMWERKS",'Material', 'COMCAM' , "FUENTE", "TIPOOFERTA", 'COMUEST', 'CampaniaDescontinuacion','Número de material','Grupo art.','Tipo de producto']]
        # Lista de valores permitidos
        valores_permitidos = {106, 101, 103, 102, 104, 105}

        # Reemplazar los valores que no están en la lista por 106
        novoApp.df_NovoApp['Grupo art.']= novoApp.df_NovoApp['Grupo art.'].apply(lambda x: x if x in valores_permitidos else 106)
        novoApp.df_NovoApp = pd.merge(novoApp.df_NovoApp,novoApp.df_Crecimientos, left_on="Grupo art.", right_on="Tipo", how="left")
        novoApp.df_NovoApp['EDL'] = novoApp.df_NovoApp['Número de material'].astype(str).apply(lambda x: 'X' if ' EDL ' in x else '-')
        if(self.PR):
           novoApp.df_NovoApp = novoApp.df_NovoApp[
            (novoApp.df_NovoApp['FUENTE'] != "Planit") &  
            (novoApp.df_NovoApp['Tipo de producto'] != 'MUESTRA') &
            (novoApp.df_NovoApp['EDL'] != 'X')&
            (novoApp.df_NovoApp['COMWERKS'] == 'PR03')&
            (novoApp.df_NovoApp['Grupo art.'] != 106)]   
        else:
            novoApp.df_NovoApp = novoApp.df_NovoApp[
            (novoApp.df_NovoApp['FUENTE'] != "Planit") &  
            (novoApp.df_NovoApp['Tipo de producto'] != 'MUESTRA') &
            (novoApp.df_NovoApp['EDL'] != 'X')&
            (novoApp.df_NovoApp['COMWERKS'] != 'PR03')]
        
        nuevos_nombres = [
            "Ce.", "Material", "Campaña", "Fuente", "Tipo de Oferta", "Cantidad estimada",
            "Descontinuación", "Descripción", "Grupo art.", "Tipo de producto",
            "Tipo", "Grupo art.","Crecimiento X", "Crecimiento X+1", "Crecimiento X+2", "Crecimiento X+3", "EDL"
        ]
        if(self.PR  == False):
            novoApp.df_NovoApp = novoApp.df_NovoApp[novoApp.df_NovoApp['Grupo art.'] == self.categoria]
        # Asignar los nuevos nombres al DataFrame
        novoApp.df_NovoApp.columns = nuevos_nombres
        novoApp.df_NovoApp['Descontinuación'] = novoApp.df_NovoApp['Descontinuación'].fillna(0).astype(int)
        año_modificado = int(año_modificado)
        novoApp.df_NovoApp= novoApp.df_NovoApp[novoApp.df_NovoApp['Campaña']>=año_modificado]
        novoApp.df_NovoApp.to_csv(self.direccionResultado+"\\df_NovoApp.csv", index=False, encoding='utf-8-sig')
        print("✅ Limpieza de datos finalizda.")

    def ejecutarMacros(self):
        # Copiar el DataFrame al portapapeles
        CampañaInicioRolling_ = self.InicioRolling
        CampañasxPeriodo_ =self.NumeroCampañas
        UltimaCampañaRolling_= int(str(self.FinRollin)+str(self.NumeroCampañas))
        Carpeta_ = self.direccionResultado
        
        # Ejecutar el notebook con parámetros
        pm.execute_notebook(
            f'{self.DireccionMacrosNovoApp}',          # 📓 tu notebook de entrada
            f'{Carpeta_}\\novoApp.ipynb',   # 📓 notebook de salida con resultados
            parameters=dict(
                CampañasxPeriodo=CampañasxPeriodo_,
                UltimaCampañaRolling=UltimaCampañaRolling_,
                CampañaInicioRolling=CampañaInicioRolling_,
                Carpeta=Carpeta_,
            )
        ) 
    
    def getSAPResultado(self):
        return novoApp.df_resultadoSAP
#LeerNovo=novoApp(Carpeta="C:\\Users\\williamtorres\\OneDrive - CETCO S.A\\Rolling forecast\\Carda 18.03.2025\\",PR=False, NombreCDL='12_17.03.2025 Reporte CDL_2022_2023_2024_2025_2026.xlsx',InicioRollingCORP=202605, InicioRollingPR=202602, AñoFinRolling=2028)
#LeerNovo.LimpiarData()
#LeerNovo.ejecutarMacros()