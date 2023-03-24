import windowsclasses
from common import methods
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
                                              bttx3='Crear Guìas')  # Texto 1
    gDmkrk.grab_set()
def rSnch():
    global filename, xvalue, yvalue, lbl1
    lbl1 = ""
    rSnchk = windowsclasses.SubWindowthreebbt(ttle='Plantillas',  # Titulo
                                              height=200,
                                              width=400,
                                              xvalue = 155,
                                              yvalue = 33,
                                              btt1=lambda: methods.getfolderpath(),  # Comando 1
                                              bttx1='Seleccionar Carpeta',  # Texto 1
                                              btt2=print('createplants'),  # Comando 2
                                              bttx2='Crear Plantillas',  # Texto 2
                                              btt3=lambda: methods.analiceplants(),  # Comando 3
                                              bttx3='Analizar Plantillas',
                                              lbl1='Master : ' + str(lbl1))  # Texto 3
    rSnchk.grab_set()