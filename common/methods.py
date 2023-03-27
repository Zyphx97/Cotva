import win32com.client
import xlwt
import glob
import math
import os
import pandas
import datetime
import qrcode
import textwrap
from tkinter.filedialog import askopenfilename, askdirectory
from PIL import Image, ImageTk, ImageFont, ImageDraw
from string import ascii_letters
from barcode import Code128
from barcode.writer import ImageWriter
from wres import windowsclasses

global lbl1, foldername, label, filename


def labelupdates():
    global lbl1


def getfolderpath():
    global foldername
    foldername = askdirectory()


def getfilepath():
    global filename
    filename = askopenfilename()
def split_df_by_total(df, max_total=20000):
    subgroups = []
    current_total = 0
    current_group = pandas.DataFrame(columns=df.columns)

    for index, row in df.iterrows():
        if current_total + row['CAP_VALOR'] > max_total:
            if not current_group.empty:
                # Add the top row to the current group
                current_group = pandas.concat([df.head(1), current_group], ignore_index=True)
                subgroups.append(current_group)
            current_total = row['CAP_VALOR']
            current_group = pandas.DataFrame(columns=df.columns)
        else:
            current_total += row['CAP_VALOR']
        current_group = pandas.concat([current_group, pandas.DataFrame(row).T], ignore_index=True)

    if not current_group.empty:
        # Add the top row to the last group
        current_group = pandas.concat([df.head(1), current_group], ignore_index=True)
        subgroups.append(current_group)

    return subgroups
def analiceplants():
    # Get Inside Paths to work in
    global foldername, label, mnrsdfflt, t1inddf, myrsdf, pLantsfolder
    pLantsfolder = foldername + '/PLANTILLAS/'
    print(pLantsfolder)
    pRefolder = foldername + '\PRE'
    mAsterfile = foldername
    # Get dataframe path
    csv_files = glob.glob(os.path.join(foldername, "*.xlsx"))
    str1 = ""
    # Get Mawb number
    for ele in csv_files:
        str1 += ele
    mfst = str1.index('FORMATO')
    mfstnm = str1[mfst + 21:-5]
    # Read DataFrame and get data from it
    mfstdf = pandas.read_excel(str1)
    # Number of Bags
    btdfnbfn = mfstdf['Bag ID'].unique()
    btdfnbfnfl = str(btdfnbfn.shape[0])
    # Number of Guides
    gddfnb = str(len(mfstdf.index))
    # Number of minor guides
    mnrsdfflt = mfstdf[mfstdf['TOTAL DECLARE VALUE'] < 50]
    mndfnb = str(len(mnrsdfflt.index))
    # Number of T1 Guides
    t1inddf = mfstdf[mfstdf['TOTAL DECLARE VALUE'] > 300]
    t1inddfflt = str(len(t1inddf.index))
    # Number of major guides
    myrsdf = mfstdf[(mfstdf['TOTAL DECLARE VALUE'] >= 50.01) & (mfstdf['TOTAL DECLARE VALUE'] <= 299.99)]
    myrsdfflt = str(len(myrsdf.index))
    # Number of minor, major and t1 referencies
    rfrmn = math.ceil(float(pandas.DataFrame(mnrsdfflt).sum(numeric_only=True)['TOTAL DECLARE VALUE']) / 20000)
    rfrt1 = math.ceil(float(pandas.DataFrame(t1inddf).sum(numeric_only=True)['TOTAL DECLARE VALUE']) / 20000)
    rfrmy = math.ceil(float(pandas.DataFrame(myrsdf).sum(numeric_only=True)['TOTAL DECLARE VALUE']) / 20000)
    rfdfnb = rfrmn + rfrt1 + rfrmy
    # DataFrame modification
    mfstdf = mfstdf.drop(['MWB','bag code','Bag ID','CLIENT REF.NO',
                      'Customer REF. NO','SHIPPER ADDRESS',
                      'CITY NAME SHIPPER','CITY CODE SHIPPER',
                      'COUNTRY NAME SHIPPER','COUNTRY CODE SHIPPER',
                      'CONSIGNEE','CONSIGNEE ADDRESS',
                      'ZIP CODE CONSIGNEE','CITY NAME CONSIGNEE',
                      'TEL CONSIGNEE','CITY CODE CONSIGNEE',
                      'COUNTRY NAME CONSIGNEE','COUNTRY CODE CONSIGNEE',
                      'UNIT OF WEIGHT(kg)','CURRENCY(USD)','ID_PAQUETERIA ',
                      'VUELO','FECHA','AEROLINEA','BULTOS'
], axis=1) # Delete cols
    new_cols = ['GATEWAY','REFERENCIA','GUIA','BULTOS','CAP_PAISORI',
               'CAP_PAISVEN','CANTIDAD','CAP_PESO','CAP_VALOR',
               'UNIDAD','CANTARI','UNITARI','CAP_DESCRIP','NOM004',
               'NOM015','NOM019','NOM020','NOM024','NOM050','NOM003',
               'NOM141','FRACCION','PREVIO','CASTLC','COMTLC','ADVAL',
               'PORCIVA','TIPO_MON','TIP_MERC','USO','VALORACION',
               'VINCULACION','UNIDAD_PESO','CAP_DESCRIPOri','Proveedor',
               'Traducciones','Pre','Piezas','amazonBarCode','MARCA',
               'MODELO','SERIE','OBSERVACIONEs','IDENTIFICADOR',
               'COMPLEMENTO1','COMPLEMENTO2','COMPLEMENTO3','IDENTIFICADOR',
               'COMPLEMENTO1','COMPLEMENTO2','COMPLEMENTO3'
] # New cols dictoinary
    new_cols_dict = {col: pandas.Series(dtype=float) for col in new_cols}
    mfstdf = pandas.concat([mfstdf, pandas.DataFrame(new_cols_dict)], axis=1)
    # Copy cols value to new cols after insert dictionary
    mfstdf['GUIA'] = mfstdf['TRACKING NUMBER(AWB)']
    mfstdf['Proveedor'] = mfstdf['SHIPPER']
    mfstdf['CAP_PESO'] = mfstdf['WEIGHT']
    mfstdf['CAP_VALOR'] = mfstdf['TOTAL DECLARE VALUE']
    mfstdf['CAP_DESCRIP'] = mfstdf['PRODUCT DESCRIPTION']
    mfstdf['CAP_DESCRIPOri'] = mfstdf['PRODUCT DESCRIPTION']
    mfstdf['CANTIDAD'] = mfstdf['TOTAL QTY']
    mfstdf['CANTARI'] = mfstdf['TOTAL QTY']
    mfstdf['Piezas'] = mfstdf['TOTAL QTY']
    mfstdf['BULTOS'] = mfstdf['TOTAL PACKAGES']
    mfstdf = mfstdf.drop(['TRACKING NUMBER(AWB)','SHIPPER',
                          'WEIGHT','TOTAL DECLARE VALUE',
                          'PRODUCT DESCRIPTION','TOTAL QTY',
                          'TOTAL PACKAGES'], axis=1)
    # New data of plants
    mfstdf['GATEWAY'] = 'MEX'
    mfstdf['CAP_PAISORI'] = 'CHN'
    mfstdf['CAP_PAISVEN'] = 'CHN'
    mfstdf['UNIDAD'] = '6'
    mfstdf['UNITARI'] = '6'
    mfstdf['NOM004'] = 0
    mfstdf['NOM015'] = 0
    mfstdf['NOM019'] = 0
    mfstdf['NOM020'] = 0
    mfstdf['NOM024'] = 0
    mfstdf['NOM050'] = 0
    mfstdf['NOM003'] = 0
    mfstdf['NOM141'] = 0
    mfstdf['FRACCION'] = '9901000100'
    mfstdf['ADVAL'] = 0
    mfstdf['PORCIVA'] = 16
    mfstdf['TIPO_MON'] = 'USD'
    mfstdf['TIP_MERC'] = 'N'
    mfstdf['USO'] = 'O'
    mfstdf['VALORACION'] = '1'
    mfstdf['VINCULACION'] = '0'
    mfstdf['UNIDAD_PESO'] = '1'
    mfstdf['Traducciones'] = '1'
    mfstdf['Pre'] = '0'

    # New split df filters Minors Guides
    mnrsdfflt = mfstdf[mfstdf['CAP_VALOR'] < 50]
    mndfnb = str(len(mnrsdfflt.index))
    # New split df filters T1 Guides
    t1inddf = mfstdf[mfstdf['CAP_VALOR'] > 300].copy()
    t1inddf['PORCIVA'] = 19
    t1inddfflt = str(len(t1inddf.index))
    # New split df filters Major guides
    myrsdf = mfstdf[(mfstdf['CAP_VALOR'] >= 50.01) & (mfstdf['CAP_VALOR'] <= 299.99)].copy()
    myrsdf['PORCIVA'] = 19
    myrsdfflt = str(len(myrsdf.index))
    # Windows Dialog "Analisis of manifiesto"
    anlcpltwmb = windowsclasses.msgbxwd(ttle='Análisis de Manifiesto',
                                        height=400,
                                        width=300,
                                        mgbstr1='Master : ' + mfstnm,
                                        mgbstr2='Guías : ' + gddfnb,
                                        mgbstr3='Bultos : ' + btdfnbfnfl,
                                        mgbstr4='Menores : ' + mndfnb,
                                        mgbstr5='Mayores : ' + myrsdfflt,
                                        mgbstr6='T1 Individual : ' + t1inddfflt,
                                        mgbstr7='Referencias : ' + str(rfdfnb))
    anlcpltwmb.grab_set()
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
        brii = bri.render()
        code = bri.get_fullcode()
        width = len(code)*26  # Assuming default "module_width" of 1
        img.paste(brii, ((1200 - width - 25), 70))
        # rest of the code
        img.paste(qrimg, (280, 1100))
        d = ImageDraw.Draw(img)
        # Texto a Imagen - Texto a Imagen - Texto a Imagen - Texto a Imagen - Texto a Imagen - Texto a Imagen


        d.text((400, 35), 'e-Commerce', font=fnt2, fill=(0, 0, 0))
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
        char_widths = []
        for char in ascii_letters:
            bbox = fnt.getbbox(char)
            char_widths.append(bbox[2] - bbox[0])

        avg_char_width = sum(char_widths) / len(char_widths)
        #avg_char_width = sum(fnt.getsize(char)[0] for char in ascii_letters) / len(ascii_letters)
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

        char_widths = []
        for char in ascii_letters:
            bbox = fnt.getbbox(char)
            char_widths.append(bbox[2] - bbox[0])

        avg_char_width = sum(char_widths) / len(char_widths)
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
def createplants():
    global filename
    # loop through each dataframe and split by 'TOTAL DECLARE VALUE'
    global mnrsdfflt, t1inddf, myrsdf, pLantsfolder
    # Scientific notation fix
    df_list = [mnrsdfflt, t1inddf, myrsdf]  # replace with your list of dataframes
    for i, df in enumerate(df_list):
        subgroups = split_df_by_total(df)
        print(subgroups)
        for j, subgroup in enumerate(subgroups):
            # Create a new Workbook object
            workbook = xlwt.Workbook()
            # Add a new Worksheet object to the Workbook object
            worksheet = workbook.add_sheet('Hoja1')
            num_format = xlwt.easyxf(num_format_str='0')
            # Write the header row to the worksheet
            header_row = subgroup.columns
            for col_index, col_data in enumerate(header_row):
                worksheet.write(0, col_index, col_data)
            # Iterate through the DataFrame and write each row to the worksheet
            for row_index, row_data in subgroup.iloc[1:].iterrows():
                for col_index, col_data in enumerate(row_data):
                    # Check if the cell value is NaN and replace it with an empty string
                    if pandas.isna(col_data):
                        col_data = ''
                    if col_index != 2:   # Apply the number format to cells in column C
                        worksheet.write(row_index, col_index, col_data)
                        worksheet.col(2).width = 15 * 256
                    else:
                        worksheet.write(row_index, col_index, col_data, num_format)
            # Save the Workbook object to a file
            filenames = f'df_{i + 1}_group_{j + 1}.xls'
            workbook.save(pLantsfolder + filenames)
            # Open Excel application
            excel = win32com.client.Dispatch("Excel.Application")
            # Open workbook
            workbook = excel.Workbooks.Open(pLantsfolder + filenames)
            # Get the custom document properties object
            custom_properties = workbook.CustomDocumentProperties
            # Set the value of custom property "CustomProperty1" to "CustomValue"
            custom_properties.Add("WorkbookGuid", False, 4, "ae5c34e8-16ac-4df2-85b5-0286647fd3c3")
            # Save and close workbook
            workbook.Save()
            workbook.Close()
            # Quit Excel application
            excel.Quit()

