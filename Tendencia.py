# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 11:55:09 2025

@author: williamtorres
"""

import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import win32com.client as win32
from DescargaTablas import lecturaInputs

class Tendencia:
    df=None
    df_Global=None
    MatrizTendencia=None
    x=None
    df_Horizonte=None
    def __init__(self, carpeta:str, CampañaInicioPR:int, CampañaInicioCORP:int, PR: bool, TipoEstimado:str, añoFinRolling: str, claseDatos:lecturaInputs, DireccionMacrosRolling:str):
        self.carpeta= carpeta
        self.TipoEstimado=TipoEstimado
        self.añoFinRolling= añoFinRolling
        self.PR= PR
        self.Tipo=TipoEstimado
        self.descargaTablas= claseDatos
        self.DireccionMacrosRolling=DireccionMacrosRolling
        if(PR):
            self.CarpetaResultado= self.carpeta+"Resultado PR03//"
            self.CampañaInicioEstimados= CampañaInicioPR
        else:
            self.CarpetaResultado= self.carpeta+"Resultado CORP//"
            self.CampañaInicioEstimados= CampañaInicioCORP
    
    def calcularVentaHistorico(self):
        # Leer desde la fila 5 en adelante
        df_temp = self.descargaTablas.get_VentaHistorica()
        
        # Omitir columnas A-D (es decir, columnas con índice 0 a 3)
        df = df_temp.iloc[:, 4:]
        # Extraer el año (primeros 4 caracteres)
        if(self.PR):
            df = df[df["C-País Desc."] == "N. Puerto Rico"]
            df = df.groupby("Time Periods")["Venta UU (SKU)"].sum().reset_index()
            df['Año'] = df['Time Periods'].str[:4]
        else:
            df = df[df["C-País Desc."] != "N. Puerto Rico"]
            df = df.groupby("Time Periods")["Venta UU (SKU)"].sum().reset_index()
            df['Año'] = df['Time Periods'].str[:4]
        
        # Extraer la campaña (últimos 2 caracteres)
        df['NCampaña'] = df['Time Periods'].str[-2:]
        df['Campaña'] = (df['Año'] + df['NCampaña']).astype(int)
        df['Año']=df['Año'].astype(int)
        df=df[df['Campaña'] <= self.CampañaInicioEstimados]
        df = df[['Campaña', 'Venta UU (SKU)', 'Año']]
        df = df.rename(columns={'Venta UU (SKU)': 'UU'})
        Tendencia.df= df
    
    def archivoGlobal(self):
        df_Horizonte= self.descargaTablas.getHorizonte()
        df_Horizonte['Año'] = df_Horizonte['Período'].str[:4]
        # Extraer la campaña (últimos 2 caracteres)
        df_Horizonte['NCampaña'] = df_Horizonte['Período'].str[-2:]
        df_Horizonte['Campaña'] = (df_Horizonte['Año'] + df_Horizonte['NCampaña']).astype(int)
        df_Horizonte['Año']=df_Horizonte['Año'].astype(int)
        if(self.PR):
            df_Horizonte = df_Horizonte[df_Horizonte["CDP"] == "N. Puerto Rico"]
        else:
            df_Horizonte = df_Horizonte[df_Horizonte["CDP"] != "N. Puerto Rico"]
        df_filtrado = df_Horizonte[df_Horizonte['Tipo'] == self.Tipo]
        Tendencia.df_Horizonte=df_filtrado
        df_Global = df_filtrado.groupby('Campaña')['UU'].sum().reset_index()
        df_Global['Campaña']=df_Global['Campaña'].astype(str)
        df_Global['Año'] = df_Global['Campaña'].str[:4]
        df_Global['Año']=df_Global['Año'].astype(int)
        df_Global['Campaña']=df_Global['Campaña'].astype(int)
        df_Global=df_Global[df_Global['Campaña'] > self.CampañaInicioEstimados]
        Tendencia.df_Global= df_Global
        
    def calculandoTendencia(self):
        #Objetivo= [231592267,	239976841,	248778111,	257222411]
        df_Historico = pd.concat([Tendencia.df_Global, Tendencia.df], ignore_index=True)
        df_Historico['Campaña']=df_Historico['Campaña'].astype(str)
        df_Historico['Ncampaña'] = df_Historico['Campaña'].str[-2:]
        tabla = df_Historico.pivot_table(index='Ncampaña', columns='Año', values='UU', aggfunc='sum', fill_value=0)
        # Convertir la tabla dinámica en una matriz de NumPy
        matriz = tabla.values
        # Agregar dos columnas de ceros (puedes cambiar los valores si lo deseas)
        x = datetime.now().year
        columnas_adicionales = np.zeros((matriz.shape[0], self.añoFinRolling-x))
        
        # Concatenar las columnas adicionales a la matriz existente
        matriz_Completa = np.hstack([matriz, columnas_adicionales])
        CrecimientoCampaña=0
        for año in range(1,len(matriz_Completa[0])):
            for i in range(len(matriz)):
                if(matriz_Completa[i,año]==0):
                    if(i-1<0):
                        matriz_Completa[i,año]=matriz_Completa[i,año-1]/matriz_Completa[len(matriz_Completa)-1,año-2]*matriz_Completa[len(matriz_Completa)-1,año-1]
                    else:
                        matriz_Completa[i,año]=matriz_Completa[i,año-1]/matriz_Completa[i-1,año-1]*matriz_Completa[i-1,año]
        
        vectorcolumns=[]
        for i in range(x-1, self.añoFinRolling+1):
            vectorcolumns.append(i)
        MatrizTendencia = pd.DataFrame(matriz_Completa, columns=vectorcolumns)
        Sumas=MatrizTendencia.sum()
        Sumas=Sumas[1:]
        print(Sumas)
        Objetivo=self.calcularObjetivo(Sumas)
        print(Objetivo)
        porcentajeCrecimiento=[]
        for i in range(len(Sumas)):
            porcentajeCrecimiento.append(Objetivo[i]/Sumas[x+i])
            print(porcentajeCrecimiento[i])
        
        for i in range(0, 4):
            MatrizTendencia[x+i]=MatrizTendencia[x+i]*porcentajeCrecimiento[i]
        
        Tendencia.MatrizTendencia= MatrizTendencia
        Tendencia.x=x
        
    def calcularObjetivo(self,SumasSinPorcentaje)->list:
        df_CrecimientosPaís= self.descargaTablas.get_df_CrecimientosPaís()
        fila_pr = df_CrecimientosPaís[df_CrecimientosPaís["País"] == "PR"]
        fila_corp = df_CrecimientosPaís[df_CrecimientosPaís["País"] == "Unidades (m)"]
        fila_Crecimiento = df_CrecimientosPaís[df_CrecimientosPaís["País"] == "Crecimiento Unidades"]
        vector_Crecimiento= fila_Crecimiento.values
        vector_Crecimiento= vector_Crecimiento[:,1:]
        vector_pr = fila_pr.values
        vector_pr=vector_pr[:,1:]
        vector_Corp = fila_corp.values
        vector_Corp= vector_Corp[:,1:]
        ventaCORSINPR= []
        for i in range(len(vector_Corp[0])):
            Diferencia= vector_pr[0,i]
            if(self.PR==False):
                Diferencia= vector_Corp[0,i]- vector_pr[0,i]
            ventaCORSINPR.append(Diferencia*1000000)
        SumaSS= SumasSinPorcentaje.tolist()
        if(SumaSS[0]<ventaCORSINPR[0]):
            SumaSS[0]=ventaCORSINPR[0]
        for i in range(1,len(SumaSS)):
            SumaSS[i]=SumaSS[i-1]*(1+vector_Crecimiento[0,i])
        return SumaSS
    def mostrarGraficaTendencia(self):
        self.calcularVentaHistorico()
        self.archivoGlobal()
        self.calculandoTendencia()
        MatrizTendencia=Tendencia.MatrizTendencia
        # Convertir el índice a tipo string
        MatrizTendencia.index = MatrizTendencia.index.astype(str)
        
        # Establecer el tamaño de la figura antes de crear el gráfico
        plt.figure(figsize=(10, 10))  # Ancho = 30, Alto = 6
        
        # Graficar los datos
        MatrizTendencia.plot(kind='line', marker='o')
        
        # Título y etiquetas
        plt.xlabel("Campañas")
        plt.ylabel("Ventas UU")
        if(self.PR):
            plt.xticks(ticks=range(1, 14))
        else:
            plt.xticks(ticks=range(1, 19))
        # Aumentar el espacio entre las etiquetas si es necesario
        plt.tight_layout()
        
        # Mostrar la gráfica
        plt.grid(True)
        plt.savefig(self.CarpetaResultado+"VentaCorportativa.pdf", format='pdf')
        print("Fin tendencia")
        
    def calculoUnidadesLinea(self):
        archivo= self.CarpetaResultado+'novoAppForecast.xlsm'
        df_Novoapp= pd.read_excel(archivo, sheet_name='Total Año')
        df_Novoapp=df_Novoapp.drop(columns=["Campaña"])
        x=Tendencia.x
        self.MatrizTendenciaAux= self.MatrizTendencia.drop(columns=[x-1,x])
        self.MatrizTendenciaAux
        df_diferencia = pd.DataFrame(self.MatrizTendenciaAux.values - df_Novoapp.values, columns=self.MatrizTendenciaAux.columns)
        CampañaInicioEstimadosA= int(str(self.CampañaInicioEstimados)[-2:])-1
        for i in range(CampañaInicioEstimadosA):
            df_diferencia[x+1][i]=0
        return df_diferencia
            
        """# Copiar el DataFrame al portapapeles
        df_diferencia.to_clipboard(index=False, excel=True)
        df_diferencia.to_excel(self.CarpetaResultado+"unidadesRolling.xlsx", index=False)
        # Ruta del archivo Excel
        archivo = self.DireccionMacrosRolling
        print(archivo)
        # Abrir Excel
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False  #  Esto desactiva los mensajes como "¿Desea reemplazar?"
        # Abrir el archivo
        workbook = excel.Workbooks.Open(archivo)
        print("despues abrir archivo")
        # Seleccionar la hoja donde pegar
        hoja = workbook.Sheets("UnidadesMeta")
        hoja.Activate()  # MUY IMPORTANTE antes de usar .Select()
        # Limpiar toda la hoja
        hoja.Cells.Clear()
        # Seleccionar la celda A1 y pegar desde el portapapeles
        hoja.Range("A1").Select()
        excel.ActiveSheet.Paste()
        print("Paso limpieza")
        workbook.Save()
        workbook.Close(SaveChanges=0)
        excel.Quit()"""
        
                
#CalculoTendencia= Tendencia(carpeta="C:\\Users\\williamtorres\\OneDrive - CETCO S.A\\Rolling forecast\\Carda 18.03.2025\\",CampañaInicioPR=202504,CampañaInicioCORP=202504,PR=False,TipoEstimado="SAP",añoFinRolling=2028)
#CalculoTendencia.mostrarGraficaTendencia()