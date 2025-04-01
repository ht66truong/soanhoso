from ttkbootstrap.constants import *
from tkinterdnd2 import TkinterDnD
from modules.gui import DataEntryApp


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = DataEntryApp(root)
    root.mainloop()
