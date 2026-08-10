
import tkinter as tk
from tkinter.ttk import *
from welcome_page import welcome_page
from agilent import agilent
from welco import welco
from taat import taat
from electrolux import electrolux
from utils import resource_path

class excel_data_extractor(tk.Tk):
    def __init__(self):
        super().__init__()
        # ==========================================
        # UI Setup (Tkinter)
        # ==========================================

        # Create the root window using standard tkinter (tk)
        self.title("ระบบพิมพ์สติกเกอร์")

        # Setting a minsize is a great idea! It prevents the user from squishing 
        # the window so small that the buttons disappear.
        self.minsize(350, 500) 

        # container is a Frame that will hold all the other frames (pages)
        container = Frame(self)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Load the PNG image file
        img = tk.PhotoImage(file=resource_path('icons\\cleanroom.png'))

        self.iconphoto(True, img)  # Ensure you have an icon.png in the same directory

        # Create a Style object to handle our button colors
        style = Style()

        # 2. Use the "alt" theme (native Windows/Mac themes often block background overrides)
        style.theme_use("xpnative")

        # set all the frames to white background
        style.configure('TFrame', background='white')

        # Configure All TTK Buttons globally
        style.configure(
            "TButton",
            background="white",
            foreground="black",        # Text color
            bordercolor="white",       # Makes the outer border white
            borderwidth=1,             # Flattens the button completely
            focusthickness=0,          # Removes the inner dotted focus ring
        )

        # Handle Button States (Hover & Click)
        style.map(
            "TButton",
            # Black borders when clicked or hovered
            bordercolor=[("pressed", "black"), ("active", "black")],
            background=[("active", "white"), ("pressed", "#f0f0f0")]
        )

        self.frames = {}

        # Add the welcome_page frame to the frames dictionary
        for F in (welcome_page, agilent, welco, taat, electrolux):
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