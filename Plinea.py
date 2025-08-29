# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 15:57:28 2025

@author: williamtorres
"""

import pandas as pd
import os
import time
import win32com.client as win32
from win32com.client import constants
from DescargaTablas import lecturaInputs
from datetime import datetime
import papermill as pm

class Plinea:
    df_Crecimientos=None
    CDL=None
    df_Horizonte=None
    df_ZMM206=None
    df_NovoApp=None
    df_ZMM206k=None
    df_resultadoSAP=None
    df_diferencia=None
    
    def __init__(self, Carpeta:str, PR:bool, NombreCDL:str, inicioRollingCORP: int, inicioRollingPR:int, añoFinRolling: int, DireccionMacrosRolling:str, tipoEstimado:str, categoria:int ):
        self.ArchivoCDL=Carpeta+NombreCDL
        self.PR=PR
        self.Carpeta=Carpeta
        self.DireccionMacrosRolling=DireccionMacrosRolling
        self.añoFinRolling=añoFinRolling
        self.tipoEstimado=tipoEstimado
        self.categoria=categoria
        if(PR):
            self.inicioRolling= inicioRollingPR
            self.direccionResultado= Carpeta+'Resultado PR03\\'+str(categoria)+'\\'
            self.camapañas=13
        else:
            self.inicioRolling= inicioRollingCORP
            self.direccionResultado= Carpeta+'Resultado CORP\\'+str(categoria)+'\\'
            self.camapañas=18
        
        traduciones  = {
                        "PaisRaw": [
                            "B. Colombia", "C. Peru", "D. Mexico", "E. Ecuador", "F. Chile", "G. Bolivia",
                            "I. Guatemala", "J. El Salvador", "K. Costa rica", "M. Rep. Dominicana",
                            "L. Panama", "N. Puerto Rico", "O. Estados Unidos"
                        ],
                        "CodigoCDP": [
                            "CO03", "PE03", "MX03", "EC03", "CL03", "BO03",
                            "GT23", "SV13", "CR03", "DO03", "PA33", "PR03", "US03"
                        ]
                    }

        # Crear DataFrame
        self.df_traduciones = pd.DataFrame(traduciones)
    
    def pandasAnteriores(self,archivoCrecimientos,CDL,df_Horizonte,df_NovoApp):
        Plinea.df_Crecimientos=archivoCrecimientos
        Plinea.CDL=CDL
        Plinea.df_Horizonte=df_Horizonte
        Plinea.df_NovoApp=df_NovoApp
        self.zmm206()
        self.operaciones()
    
    def zmm206(self):
        ruta_archivo = self.Carpeta+"MARC.txt"
        # Leer el archivo delimitado por '|', omitir las primeras 3 líneas decorativas
        df = pd.read_csv(ruta_archivo, sep='|', skiprows=3, engine='python')
        # Eliminar columnas vacías que se generan por separadores dobles o bordes
        df = df.dropna(axis=1, how='all')
        # Quitar espacios al principio y final de los nombres de columnas
        df.columns = df.columns.str.strip()
        # Quitar espacios a todos los datos (opcional pero útil)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        # Leer el archivo delimitado por '|', omitir las primeras 3 líneas decorativas
        df = pd.read_csv(ruta_archivo, sep='|', skiprows=3, engine='python')
        # Eliminar columnas vacías que se generan por separadores dobles o bordes
        df = df.dropna(axis=1, how='all')
        # Quitar espacios al principio y final de los nombres de columnas
        df.columns = df.columns.str.strip()
        # Quitar espacios a todos los datos (opcional pero útil)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df = df.iloc[1:]  # Elimina la primera fila (índice 0)
        df = df.iloc[:, 1:] # Elimina la primera columna (índice 0)
        df.columns =['', 'MANDT', 'MATNR', 'WERKS', 'PSTAT', 'LVORM', 'BWTTY', 'XCHAR',
       'MMSTA', 'MMSTD', 'MAABC', 'KZKRI', 'EKGRP', 'AUSME', 'DISPR', 'DISMM',
       'DISPO', 'KZDIE', 'PLIFZ', 'WEBAZ', 'PERKZ', 'AUSSS', 'DISLS', 'BESKZ',
       'SOBSL', 'MINBE', 'EISBE', 'BSTMI']
        df = df[['MATNR', 'WERKS', 'MMSTA']]
        # Renombrar las columnas
        df.columns = ['Material', 'Ce.', 'SM']
        df = df[df['Material'] != None]
        df = df[df['Material'].notna()]
        df['Material']= df['Material'].astype(int)
        Plinea.df_ZMM206=df
        self.zmm206k()
        
    def zmm206k(self):
        archivoZMM206k = self.Carpeta+"ZMM206 LINEA-K.XLSX"
        df_ZMM206k= pd.read_excel(archivoZMM206k, sheet_name='Sheet1')
        df_ZMM206k["Material"] = df_ZMM206k["Material"].str.replace("-", "", regex=True)
        df_ZMM206k = df_ZMM206k[['Material', 'Tipo de producto','Grupo art.']]  # Filtrás solo las columnas que querés
        df_ZMM206k=df_ZMM206k.drop_duplicates(subset=["Material","Tipo de producto"])
        df_ZMM206k['Material']= df_ZMM206k['Material'].astype(int)
        df_ZMM206k = df_ZMM206k.rename(columns={'Material': 'PT'})  # Renombrás la columna
        df_ZMM206k = df_ZMM206k.rename(columns={'Tipo de producto': 'JERARQUIA'})  # Renombrás la columna
        Plinea.df_ZMM206k=df_ZMM206k
    
    def operaciones(self):
        Plinea.df_Horizonte = pd.merge(Plinea.df_Horizonte, self.df_traduciones, left_on="CDP", right_on="PaisRaw", how="left")
        Plinea.df_Horizonte=Plinea.df_Horizonte[['Tipo', 'Marca', 'SAP',
       'Descripción SAP', 'Categoría', 'Período', 'UU', 'CodigoCDP','CDP' ]]
        Plinea.df_ZMM206['Material']= Plinea.df_ZMM206['Material'].astype(int)
        Plinea.df_Horizonte = pd.merge(Plinea.df_Horizonte, Plinea.df_ZMM206, left_on=["SAP",'CodigoCDP'], right_on=["Material", 'Ce.'], how="left")
        Plinea.df_Horizonte=Plinea.df_Horizonte[['Tipo', 'Marca', 'SAP',
       'Descripción SAP', 'Categoría', 'Período', 'UU', 'Ce.', 'CDP' ,'SM' ]]
        Plinea.df_Horizonte = pd.merge(Plinea.df_Horizonte, Plinea.df_ZMM206k, left_on=["SAP"], right_on=["PT"], how="left")
        Plinea.CDL=Plinea.CDL[['CodigoSAP', 'CampaniaDescontinuacion', 'Centro']]
        Plinea.df_Horizonte= pd.merge(Plinea.df_Horizonte,Plinea.CDL, left_on=['SAP', 'Ce.'], right_on=['CodigoSAP','Centro'], how= 'left')
        Plinea.df_Horizonte[['Tipo', 'Marca', 'SAP', 'Descripción SAP', 'Categoría', 'Período', 'UU',
                             'Grupo art.', 'JERARQUIA',
                             'CampaniaDescontinuacion', 'CDP','SM']]
        Plinea.df_Horizonte['CampaniaDescontinuacion']= Plinea.df_Horizonte['CampaniaDescontinuacion'].fillna(0)
        Plinea.df_Crecimientos.rename(columns={"Tipo": "CATEGORIA"}, inplace=True)
        Plinea.df_Horizonte= pd.merge(Plinea.df_Horizonte, Plinea.df_Crecimientos, left_on='Grupo art.', right_on= 'CATEGORIA', how='left')
        CodigosNovoApp=Plinea.df_NovoApp['Material'].unique()
        Plinea.df_Horizonte['novoApp'] = Plinea.df_Horizonte['SAP'].isin(CodigosNovoApp).map({True: 'X', False: ''})
        Plinea.df_Horizonte['EDL'] = Plinea.df_Horizonte['Descripción SAP'].astype(str).apply(lambda x: 'X' if ' EDL ' in x else '-')
        
        año_actual = datetime.now().year     
        año_modificado = f"{año_actual}00" 
        
        if(self.PR):
            Plinea.df_Horizonte = Plinea.df_Horizonte[
                (Plinea.df_Horizonte['Tipo'] == self.tipoEstimado) &  
                (Plinea.df_Horizonte['JERARQUIA'] != 'MUESTRA') &
                (Plinea.df_Horizonte['EDL'] != 'X') &
                (Plinea.df_Horizonte['SM'] != 'XX') & 
                (Plinea.df_Horizonte['SM'] != 'LQ') &
                (Plinea.df_Horizonte['Ce.']== 'PR03') & 
                (Plinea.df_Horizonte['novoApp']!= 'X') &
                (Plinea.df_Horizonte['Grupo art.']!= 106)
            ]
            print("Entro PR-------------")
        else:
            Plinea.df_Horizonte = Plinea.df_Horizonte[
                (Plinea.df_Horizonte['Tipo'] == self.tipoEstimado) &  
                (Plinea.df_Horizonte['JERARQUIA'] != 'MUESTRA') &
                (Plinea.df_Horizonte['EDL'] != 'X') &
                (Plinea.df_Horizonte['SM'] != 'XX') & 
                (Plinea.df_Horizonte['SM'] != 'LQ') &
                (Plinea.df_Horizonte['Ce.']!= 'PR03') & 
                (Plinea.df_Horizonte['Grupo art.']== self.categoria) & 
                (Plinea.df_Horizonte['novoApp']!= 'X') ]
        Plinea.df_Horizonte=Plinea.df_Horizonte[['Tipo', 'SAP', 'Categoría', 'UU', 'Grupo art.', 'SM', 'CampaniaDescontinuacion', 'Crecimiento X', 'Crecimiento X+1', 'Crecimiento X+2',
       'Crecimiento X+3', 'Descripción SAP', 'Período', 'Ce.', 'JERARQUIA', 'EDL', 'CDP', 'novoApp']]
        Plinea.df_Horizonte['campaña'] = (
    Plinea.df_Horizonte['Período'].str.replace(r' C', '', regex=True).astype(int))
        año_modificado = int(año_modificado)
        Plinea.df_Horizonte= Plinea.df_Horizonte[Plinea.df_Horizonte['campaña']>=año_modificado]
        Plinea.df_Horizonte.to_csv(self.direccionResultado+"\\df_Linea.csv", index=False, encoding='utf-8-sig')
        self.ejecutarMacros()
        
    def diferencia(self, df_diferencia):
        Plinea.df_diferencia=df_diferencia
    
    def ejecutarMacros(self):
        CampañaInicioRolling = self.inicioRolling
        UltimaCampañaRolling = int(str(self.añoFinRolling)+str(self.camapañas))
        CampañasxPeriodo = self.camapañas
        Carpeta= self.direccionResultado
        df_Meta= Plinea.df_diferencia
        df_Meta.to_csv(self.direccionResultado+"\\df_Meta.csv", index=False, encoding='utf-8-sig')
        pm.execute_notebook(
            f'{self.DireccionMacrosRolling}',          # 📓 tu notebook de entrada
            f'{Carpeta}\Rolling.ipynb',   # 📓 notebook de salida con resultados
            parameters=dict(
                CampañasxPeriodo=CampañasxPeriodo,
                UltimaCampañaRolling=UltimaCampañaRolling,
                CampañaInicioRolling=CampañaInicioRolling,
                Carpeta=Carpeta
            )
        ) 
    def guardarDatosCorrida(self):
        archivoNovoapp = self.direccionResultado+"Rolling-Forecast.xlsm"
        df_resultadoSAPNovoAPP= pd.read_excel(archivoNovoapp, sheet_name='Consolidado')
        df_CalculosNovoApp= pd.read_excel(archivoNovoapp, sheet_name='Final')
        df_resultadoSAPNovoAPP['TO'] = df_resultadoSAPNovoAPP['TO'].astype(str).str.zfill(3)
        df_resultadoSAPNovoAPP.to_excel(self.direccionResultado+"\\ResultadoSAPPL.xlsx", index=False)
        df_CalculosNovoApp.to_csv(self.direccionResultado+"\\CalulosPL.csv", index=False, encoding='utf-8-sig')
        Plinea.df_resultadoSAP= df_resultadoSAPNovoAPP
        
    
    def getSAPResultado(self):
        return Plinea.df_resultadoSAP