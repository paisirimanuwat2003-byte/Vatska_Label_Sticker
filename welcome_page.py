import tkinter as tk
from tkinter.ttk import *
from PIL import Image, ImageTk
from utils import resource_path


class welcome_page(Frame):
    def __init__(self,parent, controller):
        """Class for the first interface."""
        super().__init__(parent)
        self.controller = controller

        # Create and display a basic label
        title_label = tk.Label(self, text="ระบบพิมพ์สติ๊กเกอร์", font=('Arial', 20, 'bold'), bg='white', fg='black')
        title_label.pack(pady=10)

        # -----------------------------------------
        # 1. LOAD THE IMAGE
        # -----------------------------------------
        #agilent
        raw_img1 = Image.open(resource_path('icons\\agilent.png'))
        # 2. Resize it (e.g., 50x50 pixels) using high-quality resampling
        raw_img1 = raw_img1.resize((120, 50), Image.Resampling.LANCZOS)
        self.img1 = ImageTk.PhotoImage(raw_img1)

        #welco
        raw_img2 = Image.open(resource_path('icons\\welco.png'))
        raw_img2 = raw_img2.resize((120, 50), Image.Resampling.LANCZOS)
        self.img2 = ImageTk.PhotoImage(raw_img2)

        #electrolux
        raw_img3 = Image.open(resource_path('icons\\electrolux.jpg'))
        raw_img3 = raw_img3.resize((120, 50), Image.Resampling.LANCZOS)
        self.img3 = ImageTk.PhotoImage(raw_img3)

        #taat
        raw_img4 = Image.open(resource_path('icons\\stgobain.png'))
        raw_img4 = raw_img4.resize((120, 50), Image.Resampling.LANCZOS)
        self.img4 = ImageTk.PhotoImage(raw_img4)

        # The 4 job Buttons 
        btn_copy1 = Button(
            self,
            image=self.img1, 
            command=lambda: self.controller.show_frame("agilent"), 
            state="normal",
            )
        btn_copy1.pack(padx=5, pady=10)

        agilent_label = tk.Label(self, text="อาจิแลนท์", font=('Arial', 18), bg='white', fg='black')
        agilent_label.pack(padx=5)

        btn_copy2 = Button(
            self,
            image=self.img2, 
            command=lambda: self.controller.show_frame("welco"), 
            state="normal",
            )
        btn_copy2.pack(padx=5, pady=10)

        welco_label = tk.Label(self, text="เวลโก้", font=('Arial', 18), bg='white', fg='black')
        welco_label.pack(padx=5)

        btn_copy3 = Button(
            self,
            image=self.img3, 
            command=lambda: self.controller.show_frame("electrolux"), 
            state="normal",
            )
        btn_copy3.pack(padx=5, pady=10)

        electrolux_label = tk.Label(self, text="อิเลคโทรลัคซ์", font=('Arial', 18), bg='white', fg='black')
        electrolux_label.pack(padx=5)

        btn_copy4 = Button(
            self,
            image=self.img4, 
            command=lambda: self.controller.show_frame("taat"), 
            state="normal",
            )
        btn_copy4.pack(padx=5, pady=10)

        taat_label = tk.Label(self, text="เซนต์โกแบง TAAT", font=('Arial', 18), bg='white', fg='black')
        taat_label.pack(padx=5)

        #Bottom left decor
        raw_img5 = Image.open(resource_path('icons\\cleanroom.png'))
        raw_img5 = raw_img5.resize((50, 50), Image.Resampling.LANCZOS)
        self.img5 = ImageTk.PhotoImage(raw_img5)
        # 2. Create a Label to hold the image
        decor_label = tk.Label(self, image=self.img5, bg='white')

        # Force it exactly to the bottom-left corner
        decor_label.place(x=0, y=600, anchor=tk.SW)   

        # version label
        version_label = tk.Label(self, text="Version 1.0", font=('Arial', 8), bg='white', fg='black')    
        version_label.pack(side="bottom",anchor="se")

