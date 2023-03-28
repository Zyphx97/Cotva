import customtkinter

import windowsclasses
from windowsclasses import SubWindowthreebbt

from common import methods
global getentryvalue1
def eCwd():
    eCwdmk = windowsclasses.SubWindowthreebbt(ttle='e-Commerce',  # Titulo
                                              height=200,
                                              width=300,
                                              btt1=lambda: rSnch(),  # Comando 1
                                              bttx1='Plantillas',  # Texto 1
                                              btt2=lambda: gDmkr(),  # Comando 2
                                              bttx2='Guías',  # Texto 2
                                              btt3='Guias',  # Comando 3
                                              bttx3='Precaptura')  # Texto 3
    eCwdmk.grab_set()
def oPcwd():
    oPcwdk = windowsclasses.SubWindowthreebbt(ttle='Operación',  # Titulo
                                              height=200,
                                              width=300,
                                              btt1='Plantillas',  # Comando 1
                                              bttx1='Reportes',  # Texto 1
                                              btt2='print("Adios")',  # Comando 2
                                              bttx2='e-Documents',  # Texto 2
                                              btt3='Guias',  # Comando 3
                                              bttx3='Renombre')  # Texto 3
    oPcwdk.grab_set()
def gDmkr():
    gDmkrk = windowsclasses.SubWindowthreebbt(ttle='Creación de Guías',  # Titulo
                                              height=200,
                                              width=400,
                                              btt1=lambda: methods.getfilepath(),  # Comando 1
                                              bttx1='Selecciona Manifiesto',  # Texto 1
                                              btt2=lambda: methods.GuideCreationLooper(),
                                              bttx2='Crear Guìas')  # Texto 1
    gDmkrk.grab_set()
def rSnch():
    global filename, xvalue, yvalue, lbl1, getentryvalue1
    lbl1 = ""
    rSnchk = windowsclasses.SubWindowthreebbt(ttle='Plantillas',  # Titulo
                                              height=200,
                                              width=400,
                                              xvalue = 155,
                                              yvalue = 33,
                                              btt1=lambda: methods.getfolderpath(),  # Comando 1
                                              bttx1='Seleccionar Carpeta',  # Texto 1
                                              btt2=lambda: methods.createplants(),  # Comando 2
                                              bttx2='Crear Plantillas',  # Texto 2
                                              btt3=lambda: methods.analiceplants(),  # Comando 3
                                              bttx3='Analizar Plantillas')
    rSnchk.grab_set()
def rFrnc(getentryvalue1):
    global getthefuckout
    rFrncwd = windowsclasses.SubWindowthreebbt(ttle='Reference',
                                               height=200,
                                               width=300,
                                               xvalue=15,
                                               yvalue=20,
                                               entrypl="Yes",
                                               xentryvalue=25,
                                               yentryvalue=110,
                                               xlistvalue=25,
                                               ylistvalue=80,
                                               xdestr=25,
                                               ydestr=160,
                                               listbox='True',
                                               getentryvalue=getentryvalue1,
                                               lbl1='Seleccione prefijo e ingrese numero de referencia')
    rFrncwd.grab_set()
    rFrncwd.wait_window()
    global getthefuckout
    getthefuckout = rFrncwd.getentryvalue1
    print(getthefuckout)
    return getthefuckout

