import tkinter as tk
from tkinter.ttk import *


class welcome_page(Frame):
    def __init__(self,parent, controller):
        """Class for the first interface."""
        super().__init__(parent)
        self.controller = controller

        # The 4 job Buttons 
        btn_copy1 = Button(self, text="Agilent", command=lambda: self.controller.show_frame("agilent"), state="normal")
        btn_copy1.pack(side=tk.LEFT , padx=5)

        btn_copy2 = Button(self, text="Welco", command=lambda: self.controller.show_frame("welco"), state="normal")
        btn_copy2.pack(side=tk.LEFT, padx=5)

        btn_copy3 = Button(self, text="Electrolux", command=lambda: self.controller.show_frame("electrolux"), state="normal")
        btn_copy3.pack(side=tk.LEFT, padx=5)

        btn_copy4 = Button(self, text="TAAT", command=lambda: self.controller.show_frame("taat"), state="normal")
        btn_copy4.pack(side=tk.LEFT, padx=5)


