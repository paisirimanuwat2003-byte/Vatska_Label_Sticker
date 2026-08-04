
import tkinter as tk
from tkinter.ttk import *
from welcome_page import welcome_page
from agilent import agilent

class excel_data_extractor(tk.Tk):
    def __init__(self):
        super().__init__()
        # ==========================================
        # UI Setup (Tkinter)
        # ==========================================

        # Create the root window using standard tkinter (tk)
        self.title("Excel Data Extractor For Labeling")

        # Setting a minsize is a great idea! It prevents the user from squishing 
        # the window so small that the buttons disappear.
        self.minsize(350, 500) 

        # container is a Frame that will hold all the other frames (pages)
        container = Frame(self)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Load the PNG image file
        img = tk.PhotoImage(file="vatska_icon.png")

        self.iconphoto(True, img)  # Ensure you have an icon.png in the same directory

        # Create a Style object to handle our button colors
        style = Style()

        self.frames = {}

        # Add the welcome_page frame to the frames dictionary
        for F in (welcome_page, agilent):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("welcome_page")

    def show_frame(self, page_name):
        """Show a frame for the given page name"""
        frame = self.frames[page_name]
        frame.tkraise()


if __name__ == "__main__":
    app = excel_data_extractor()
    app.mainloop()