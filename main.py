# -*- coding: utf-8 -*-
"""
Created on Tue Apr 15 16:16:10 2025

@author: williamtorres
"""
from novoApp import novoApp
from Tendencia import Tendencia
from Plinea import Plinea
from DescargaTablas import lecturaInputs
import pandas as pd
import time
import os
import shutil

def datos(PR:bool, Carpeta:str, NombreCDL:str, InicioRollingCORP:int, InicioRollingPR:int,AñoFinRolling:int, TipoEstimado:str, claseDatos:lecturaInputs, DireccionMacrosNovoApp:str, DireccionMacrosRolling:str ):
    
    #Novoaap
    LeerNovo=novoApp(Carpeta=Carpeta,PR=PR, NombreCDL=NombreCDL, categoria=103,
                     InicioRollingCORP=InicioRollingCORP, InicioRollingPR=InicioRollingPR, 
                     AñoFinRolling=AñoFinRolling, claseDatos=claseDatos, DireccionMacrosNovoApp=DireccionMacrosNovoApp)
    LeerNovo.LimpiarData()
    LeerNovo.ejecutarMacros()
    resultadoSAPNOVO= LeerNovo.getSAPResultado()
    #Tendencia
    CalculoTendencia= Tendencia(carpeta=Carpeta,CampañaInicioPR=InicioRollingPR,
                                CampañaInicioCORP=InicioRollingCORP,PR=PR,
                                TipoEstimado=TipoEstimado,añoFinRolling=AñoFinRolling,
                                claseDatos=claseDatos,DireccionMacrosRolling=DireccionMacrosRolling,categoria=103)
    CalculoTendencia.mostrarGraficaTendencia()
    
    print("Inicio linea")
    #Línea
    archivoCrecimientos=LeerNovo.df_Crecimientos
    CDL=claseDatos.getCDL()
    CDL.rename(columns={"CDP": "Centro"}, inplace=True)
    df_NovoApp=LeerNovo.df_NovoApp
    df_Horizonte= claseDatos.getHorizonte()
    LineaCorrida= Plinea(Carpeta=Carpeta,inicioRollingCORP=InicioRollingCORP, tipoEstimado=TipoEstimado,categoria=103,
                         inicioRollingPR=InicioRollingPR, añoFinRolling=AñoFinRolling, 
                         PR=True, NombreCDL= NombreCDL, DireccionMacrosRolling=DireccionMacrosRolling)
    LineaCorrida.diferencia(CalculoTendencia.calculoUnidadesLinea())
    LineaCorrida.pandasAnteriores(archivoCrecimientos,CDL,df_Horizonte,df_NovoApp)
    resultadoSAPLINEA= LineaCorrida.getSAPResultado()
    df_resultado = pd.concat([resultadoSAPNOVO, resultadoSAPLINEA], ignore_index=True)
    return df_resultado

def leerDatos(carpeta:str, CI:str, CF:str, GLOBAL:str ):
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
    

def nam():
    carpeta = 'C:\\Users\\williamtorres\\Desktop\\Nueva carpeta (4)'
    campañaInicio="202601"
    campañaFin="202818"
    RutaArchivoGlobal="C:\\Users\\williamtorres\\Downloads\\02 de julio 2025 (M6).xlsx"
    RutaArchivoCDL= "C:\\Users\\williamtorres\\Downloads\\24_13.06.2025 Reporte CDL_2023_2024_2025_2026_2027.xlsx"
    RutaArchivoCrecimientos= "C:\\Users\\williamtorres\\OneDrive - CETCO S.A\\Rolling forecast\\Inputs\\Crecimiento.xlsx"
    RutaArchivoVentaHistorica= "C:\\Users\\williamtorres\\OneDrive - CETCO S.A\\Rolling forecast\\Inputs\\Total_CORP.xlsx"
    RutaMacrosNovoApp=  "C:\\Users\\williamtorres\\OneDrive - CETCO S.A\\Rolling forecast\\Inputs\\novoAppForecast.xlsm"
    RutaMacrosRolling=  "C:\\Users\\williamtorres\\OneDrive - CETCO S.A\\Rolling forecast\\Inputs\\Rolling-Forecast.xlsm"
    InicioRollingCORP=202610
    InicioRollingPR=202607
    AñoFinRolling=2028
    TipoEstimado="Planit"
    lecturaDatos= leerDatos(carpeta=carpeta, CI= campañaInicio, CF= campañaFin, GLOBAL=RutaArchivoGlobal)
    lecturaDatos.leerOtrosInputs(RutaCDL=RutaArchivoCDL, RutaArchivoCrecimiento=RutaArchivoCrecimientos, RutaHistorico=RutaArchivoVentaHistorica)
    #Inicio algoritmo
    os.system("taskkill /f /im excel.exe")
    #resultadoCORP=datos(PR=False,InicioRollingPR=InicioRollingPR,AñoFinRolling=AñoFinRolling,Carpeta=carpeta+"\\",NombreCDL=RutaArchivoCDL, InicioRollingCORP= InicioRollingCORP,TipoEstimado=TipoEstimado,claseDatos=lecturaDatos,DireccionMacrosNovoApp=RutaMacrosNovoApp,DireccionMacrosRolling=RutaMacrosRolling)
    print("Corrida con exito CORP")
    resultadoPR03=datos(PR=True,InicioRollingPR=InicioRollingPR,AñoFinRolling=AñoFinRolling,Carpeta=carpeta+"\\",NombreCDL=RutaArchivoCDL, InicioRollingCORP= InicioRollingCORP,TipoEstimado=TipoEstimado,claseDatos=lecturaDatos,DireccionMacrosNovoApp=RutaMacrosNovoApp,DireccionMacrosRolling=RutaMacrosRolling)
    print("Corrida con exito PR03")
    df_resultado = pd.concat([resultadoCORP, resultadoPR03], ignore_index=True)
    # Verifica si la carpeta ya existe, si no, la crea
    if not os.path.exists(carpeta+"\\Carga SAP"):
        os.mkdir(carpeta+"\\Carga SAP")
    df_resultado.to_excel(carpeta+"\\Carga SAP\\"+'CargaSAP.xlsx', index=False)
    
nam()