import tkinter as tk, PIL, PIL.ImageTk, os, customtkinter, pandas, datetime, qrcode, textwrap, barcode, glob
from tkinter import Canvas, ttk
from PIL import Image, ImageTk, ImageFont, ImageDraw
from string import ascii_letters
from barcode import Code128
from barcode.writer import ImageWriter
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askdirectory

global filename
global foldername
foldername = ""
filename = ""

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
# Window Functions
def eCwd():
    eCwdmk = SubWindowthreebbt(ttle = 'e-Commerce',                 # Titulo
                               gmtry = '300x200',                   # Tamaño de la ventana
                               rsze = 'heigth=false,width=false',   # Tamaño modificable
                               btt1 = lambda: rSnch(),                 # Comando 1
                               bttx1 = 'Plantillas',                # Texto 1
                               btt2 = lambda: gDmkr(),              # Comando 2
                               bttx2 = 'Guías',                     # Texto 2
                               btt3 = 'Guias',                      # Comando 3
                               bttx3 = 'Precaptura')                # Texto 3
def oPcwd():
    oPcwdk = SubWindowthreebbt(ttle = 'Operación',                  # Titulo
                               gmtry = '300x200',                   # Tamaño de la ventana
                               rsze = 'heigth=false,width=false',   # Tamaño modificable
                               btt1 = 'Plantillas',                 # Comando 1
                               bttx1 = 'Reportes',                  # Texto 1
                               btt2 = 'print("Adios")',             # Comando 2
                               bttx2 = 'e-Documents',               # Texto 2
                               btt3 = 'Guias',                      # Comando 3
                               bttx3 = 'Renombre')                  # Texto 3
def gDmkr():
    gDmkrk = SubWindowthreebbt(ttle = 'Creación de Guías',          # Titulo
                               gmtry = '300x150',                   # Tamaño de la ventana
                               rsze = 'heigth=false,width=false',   # Tamaño Modificable
                               btt1 = lambda:getfilepath(),         # Comando 1
                               bttx1 = 'Selecciona Manifiesto',     # Texto 1
                               btt2 = lambda:GuideCreationLooper(),
                               bttx3= 'Crear Guìas')                     # Texto 1
def rSnch():
    global filename
    rSnchk = SubWindowthreebbt(ttle = 'Plantillas',                 # Titulo
                               gmtry = '300x200',                   # Tamaño de la ventana
                               rsze = 'heigth=false,width=false',   # Tamaño modificable
                               btt1 = lambda: getfolderpath(),        # Comando 1
                               bttx1 = 'Seleccionar Carpeta',       # Texto 1
                               btt2 = lambda: createplants(),       # Comando 2
                               bttx2 = 'Crear Plantillas',              # Texto 2
                               btt3 = lambda: analiceplants(),      # Comando 3
                               bttx3 = 'Analizar Plantillas')          # Texto 3


# Method functions
def getfolderpath():
    global foldername
    foldername = askdirectory()
def getfilepath():
    global filename
    filename = askopenfilename()
def analiceplants():
    global foldername
    pLantsfolder = foldername + '\PLANTILLAS'
    pRefolder = foldername + '\PRE'
    mAsterfile = foldername
    csv_files = glob.glob(os.path.join(foldername, "*.xlsx"))
    print(csv_files)










def GuideCreationLooper():
    global filename
    # Elección de Archivo Excel acorde al Layout
    df = pandas.read_excel(filename)
    filepath = os.path.dirname(filename)
    newpath = filepath + "\e-Docs"
    taildf = len(df.index)

    if not os.path.exists(newpath):
        os.makedirs(newpath)
    finalfilepath = newpath + "\Guia"
    print(filepath)
    print(df)

    for i in range(0,taildf):

        # Guia House
        hawb = df['TRACKING NUMBER(AWB)'].loc[df.index[i]]
        hawbs = str(hawb)
        # Bag id
        bgid = df['Bag ID'].loc[df.index[i]]
        bgids = str(bgid)
        # Proveedor0
        shpn = df['SHIPPER'].loc[df.index[i]]
        shpns = str(shpn)
        # Direccion de proveedor
        shpa = df['SHIPPER ADDRESS'].loc[df.index[i]]
        shpas = str(shpa)
        # Ciudad del proveedor
        shpc = df['COUNTRY NAME SHIPPER'].loc[df.index[i]]
        shpcs = str(shpc)
        # Consignatario
        cnsn = df['CONSIGNEE'].loc[df.index[i]]
        cnsns = str(cnsn)
        # Direccion de consignatario
        cnsa = df['CONSIGNEE ADDRESS'].loc[df.index[i]]
        cnsas = str(cnsa)
        # Codigo postal Consignatario
        cnsz = df['ZIP CODE CONSIGNEE'].loc[df.index[i]]
        cnszs = str(cnsz)
        # Ciudad del consignatario
        cnsc = df['CITY NAME CONSIGNEE'].loc[df.index[i]]
        cnscs = str(cnsc)
        # Pais del consignatario
        ctnsc = df['COUNTRY NAME CONSIGNEE'].loc[df.index[i]]
        ctnscs = str(ctnsc)
        # Peso Paquete
        pkwh = df['WEIGHT'].loc[df.index[i]]
        pkwhs = str(pkwh)
        # Descripcion del paquete
        pkdp = df['PRODUCT DESCRIPTION'].loc[df.index[i]]
        pkdps = str(pkdp)
        # Telefono del consignatario
        cnst = df['TEL CONSIGNEE'].loc[df.index[i]]
        cnsts = str(cnst)
        # Telefono del consignatario
        pkqt = df['TOTAL PACKAGES'].loc[df.index[i]]
        pkqts = str(pkqt)
        # Fecha
        now = datetime.datetime.now()
        now2 = now.date()
        nows = str(now2)
        # Image Creation
        img = Image.new('RGB', (1200, 1800), color='white')

        # Font Selection

        fnt = ImageFont.truetype('resources\Arial.ttf', 28)
        fnt2 = ImageFont.truetype('resources\Arial.ttf', 44)
        fnt3 = ImageFont.truetype('resources\Arial.ttf', 34)
        fnt4 = ImageFont.truetype('resources\Arial.ttf', 30)
        fnt5 = ImageFont.truetype('resources\Arial.ttf', 60)

        # QR Code Gen

        qri = qrcode.QRCode(
            version=2,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=20,
            border=4,
        )
        qri.add_data(hawbs)
        qri.make
        qrimg = qri.make_image(fill_color="black", back_color="white").convert('RGB')

        # Bar Code Gen

        bri = Code128(hawbs, writer=ImageWriter())
        bri.save('bri', {"module_width":0.1, "module_height":10, "font_size": 6, "text_distance": 2, "quiet_zone": 3})
        brii = Image.open('bri.png')
        img.paste(qrimg,(280,1100))
        img.paste(brii,(800,50))
        d = ImageDraw.Draw(img)

        # Texto a Imagen - Texto a Imagen - Texto a Imagen - Texto a Imagen - Texto a Imagen - Texto a Imagen


        d.text((500, 50), 'e-Commerce', font=fnt2, fill=(0, 0, 0))
        d.text((20, 50), 'Ultima Milla', font=fnt, fill=(0, 0, 0))
        d.text((20, 340), ('Guía'), font=fnt3, fill=(0, 0, 0))
        d.text((20, 380), hawbs, font=fnt, fill=(0, 0, 0))
        d.text((600, 340), 'Bag ID', font=fnt3, fill=(0, 0, 0))
        d.text((600, 380), bgids, font=fnt, fill=(0, 0, 0))

        #Datos del Proveedor

        d.text((20, 520), 'Datos del Proveedor', font=fnt4, fill=(0, 0, 0))
        d.text((20, 555), shpns, font=fnt, fill=(0, 0, 0))
        d.text((20, 590), shpcs, font=fnt, fill=(0, 0, 0))

        # Text Wrapping (Shipper Address)

        avg_char_width = sum(fnt.getsize(char)[0] for char in ascii_letters) / len(ascii_letters)
        max_char_count = int(img.size[0] * .418 / avg_char_width)
        shpas = textwrap.fill(text=shpas, width=max_char_count)
        d.text(xy=(230,680), text=shpas, font=fnt, fill='#000000', anchor='mm')

        # Datos del Destinatario

        d.text((600, 520), 'Datos del Destinatario', font=fnt4, fill=(0, 0, 0))
        d.text((600, 555), cnsns, font=fnt, fill=(0, 0, 0))
        d.text((880, 590), cnszs, font=fnt, fill=(0, 0, 0))
        d.text((600, 590), cnscs, font=fnt, fill=(0, 0, 0))
        d.text((600, 660), 'Teléfono: ', font=fnt4, fill=(0, 0, 0))
        d.text((880, 660), cnsts, font=fnt, fill=(0, 0, 0))
        d.text((600, 620), ctnscs, font=fnt, fill=(0, 0, 0))

        # Text Wrapping (Shipper Address)

        avg_char_width = sum(fnt.getsize(char)[0] for char in ascii_letters) / len(ascii_letters)
        max_char_count = int(img.size[0] * .450 / avg_char_width)
        cnsas = textwrap.fill(text=cnsas, width=max_char_count)
        d.text(xy=(850,750), text='Dirección: ' + cnsas, font=fnt, fill='#000000', anchor='mm')

        # Datos de Mercancia

        d.text((20, 905), 'Peso: ', font=fnt4, fill=(0, 0, 0))
        d.text((100, 905), pkwhs, font=fnt, fill=(0, 0, 0))
        d.text((200, 905), 'Bultos: ', font=fnt4, fill=(0, 0, 0))
        d.text((350, 905), pkqts, font=fnt, fill=(0, 0, 0))

        #Datos Extra

        d.text((800, 905), 'Fecha: ', font=fnt4, fill=(0, 0, 0))
        d.text((900, 905), nows, font=fnt, fill=(0, 0, 0))
        d.text((560, 1050), 'EAD', font=fnt5, fill=(0, 0, 0))


        img.save(finalfilepath + " " + str(i) +'.pdf')
# Classes
class SubWindowthreebbt(customtkinter.CTkToplevel):
    def __init__(self,ttle,gmtry,rsze,btt1 = "",btt2 = "",btt3 = "",bttx1 = "",bttx2 = "",bttx3 = "",lbl1 = ""):
        super().__init__()
        self.title(str(ttle))
        self.geometry(gmtry)
        self.resizable=str(rsze)

    # Use CTkButton instead of tkinter Button
        if btt1 != "":
            button = customtkinter.CTkButton(master=self, text=bttx1,command=btt1, height=40, width=120)
            button.place(x=15, y=25)
    # Use CTkButton instead of tkinter Button
        if btt2 != "":
            button = customtkinter.CTkButton(master=self, text=bttx2,command=btt2,height=40, width=120)
            button.place(x=15, y=75)
    # Use CTkButton instead of tkinter Button
        if btt3 != "":
            button = customtkinter.CTkButton(master=self, text=bttx3,command=btt3, height=40, width=120)
            button.place(x=15, y=125)
class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Cotva")
        self.geometry("800x400")
        self.resizable(width=False, height=False)
        # Logo Setup
        logo = "resources/Logo Cotva Tk 2.png"
        my_image = customtkinter.CTkImage(light_image=Image.open(logo),
                                          dark_image=Image.open(logo),
                                          size=(311, 400))
        label = customtkinter.CTkLabel(master=self,
                                       text="",
                                       width=311,
                                       height=400,
                                       image=my_image)
        label.pack()
        # Icon setup
        icon = "resources/icon.png"
        load = Image.open(icon)
        render = ImageTk.PhotoImage(load)
        self.iconphoto(False, render)
        # Use CTkButton instead of tkinter Button
        self.button = customtkinter.CTkButton(master=self, text="e-Commerce",
                                              command=lambda: eCwd())
        self.button.place(x=80,y=180)
        # First button setup complete
        self.button = customtkinter.CTkButton(master=self, text="Operaciónes",
                                              command=lambda: oPcwd(),
                                              height=40,width=120)
        self.button.place(x=600,y=180)
        # Second button setup complete
# Main Call
if __name__ == "__main__":

    mw = MainWindow()
    mw.mainloop()