from PIL import Image, ImageTk
import customtkinter
import windowsfunctions
global xvalue, yvalue

class SubWindowthreebbt(customtkinter.CTkToplevel):
    global xvalue, yvalue, lbl1, label, sbwdmstnm

    def __init__(self, ttle="", xvalue='', yvalue='', height="", width="", btt1="", btt2="", btt3="",
                 bttx1="", bttx2="", bttx3="", lbl1=""):
        super().__init__()
        self.title(str(ttle))
        self.resizable(width='false', height='false')

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) / 2
        y = (screen_height - height) / 2
        # Set the position of the new window
        self.geometry("%dx%d+%d+%d" % (width, height, x, y))
        # Use CTkButton instead of tkinter Button
        if btt1 != "":
            button = customtkinter.CTkButton(master=self, text=bttx1, command=btt1, height=40, width=120)
            button.place(x=15, y=25)
        # Use CTkButton instead of tkinter Button
        if btt2 != "":
            button = customtkinter.CTkButton(master=self, text=bttx2, command=btt2, height=40, width=120)
            button.place(x=15, y=75)
        # Use CTkButton instead of tkinter Button
        if btt3 != "":
            button = customtkinter.CTkButton(master=self, text=bttx3, command=btt3, height=40, width=120)
            button.place(x=15, y=125)
        if lbl1 != "":
            label = customtkinter.CTkLabel(master=self, text=lbl1)
            label.place(x=xvalue, y=yvalue)
class msgbxwd(customtkinter.CTkToplevel):
    def __init__(self, ttle,
                 height="",
                 width="",
                 mgbstr1="",
                 mgbstr2="",
                 mgbstr3="",
                 mgbstr4="",
                 mgbstr5="",
                 mgbstr6="",
                 mgbstr7=""):
        super().__init__()
        self.title(str(ttle))
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) / 2
        y = (screen_height - height) / 2
        self.resizable(width='false', height='false')
        self.geometry("%dx%d+%d+%d" % (width, height, x, y))
        if mgbstr1 != "":
            label1 = customtkinter.CTkLabel(master=self, text=mgbstr1, width=150, height=50)
            label1.place(x=75, y=5)
        if mgbstr2 != "":
            label2 = customtkinter.CTkLabel(master=self, text=mgbstr2, width=150, height=50)
            label2.place(x=75, y=60)
        if mgbstr3 != "":
            label3 = customtkinter.CTkLabel(master=self, text=mgbstr3, width=150, height=50)
            label3.place(x=75, y=115)
        if mgbstr4 != "":
            label4 = customtkinter.CTkLabel(master=self, text=mgbstr4, width=150, height=50)
            label4.place(x=75, y=170)
        if mgbstr5 != "":
            label5 = customtkinter.CTkLabel(master=self, text=mgbstr5, width=150, height=50)
            label5.place(x=75, y=225)
        if mgbstr6 != "":
            label6 = customtkinter.CTkLabel(master=self, text=mgbstr6, width=150, height=50)
            label6.place(x=75, y=285)
        if mgbstr7 != "":
            label7 = customtkinter.CTkLabel(master=self, text=mgbstr7, width=150, height=50)
            label7.place(x=75, y=340)

        self.focus_force
class MainWindow(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Cotva")
        self.resizable(width=False, height=False)
        # Calculate the center of the screen
        width = 800
        height = 400
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) / 2
        y = (screen_height - height) / 2
        # Set the position of the new window
        self.geometry("%dx%d+%d+%d" % (width, height, x, y))
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
                                              command=lambda: windowsfunctions.eCwd())
        self.button.place(x=80, y=180)
        # First button setup complete
        self.button = customtkinter.CTkButton(master=self, text="Operaciónes",
                                              command=lambda: windowsfunctions.oPcwd(),
                                              height=40, width=120)
        self.button.place(x=600, y=180)
        # Second button setup complete
        self.attributes('-topmost', False
                        )