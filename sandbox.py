import customtkinter


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("600x500")
        self.title("CTk example")
        string = ["option 1", "option 2"]
        optionmenu_var = customtkinter.StringVar(value="option 2")  # set initial value
        combobox = customtkinter.CTkOptionMenu(master=self,
                                               values=string,
                                               command=optionmenu_callback,
                                               variable=optionmenu_var)
        combobox.pack(padx=20, pady=10)


def optionmenu_callback(choice):
    print("optionmenu dropdown clicked:", choice)
app = App()
app.mainloop()