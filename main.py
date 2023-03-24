import sys

import customtkinter

sys.path.append('./wres')
from wres import windowsclasses

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

# Main Call
if __name__ == "__main__":
    mw = windowsclasses.MainWindow()
    mw.mainloop()
