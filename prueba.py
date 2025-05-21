import win32com.client
try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)  # Conexión activa
    session = connection.Children(0)  # Primera sesión activa
except Exception as e:
    print(e)
            
session.findById("wnd[0]/tbar[0]/okcd").text = "SE16"
session.findById("wnd[0]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZTPPFOTOPLANIT"
session.findById("wnd[0]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/usr/txtI2-LOW").text = "202601"
session.findById("wnd[0]/usr/txtI2-HIGH").text = "202818"
session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[8]").press()
session.findById("wnd[0]/tbar[1]/btn[45]").press()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\williamtorres\\Desktop\\Nueva carpeta (3)"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DWaES"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press()