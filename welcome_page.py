import tkinter as tk
from tkinter.ttk import *
from PIL import Image, ImageTk


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
        raw_img1 = Image.open(r'D:\vatska\software\icons\agilent.png')
        # 2. Resize it (e.g., 50x50 pixels) using high-quality resampling
        raw_img1 = raw_img1.resize((120, 50), Image.Resampling.LANCZOS)
        self.img1 = ImageTk.PhotoImage(raw_img1)

        #welco
        raw_img2 = Image.open(r'D:\vatska\software\icons\welco.png')
        raw_img2 = raw_img2.resize((120, 50), Image.Resampling.LANCZOS)
        self.img2 = ImageTk.PhotoImage(raw_img2)

        #electrolux
        raw_img3 = Image.open(r'D:\vatska\software\icons\electrolux.jpg')
        raw_img3 = raw_img3.resize((120, 50), Image.Resampling.LANCZOS)
        self.img3 = ImageTk.PhotoImage(raw_img3)

        #taat
        raw_img4 = Image.open(r'D:\vatska\software\icons\stgobain.png')
        raw_img4 = raw_img4.resize((120, 50), Image.Resampling.LANCZOS)
        self.img4 = ImageTk.PhotoImage(raw_img4)

        # The 4 job Buttons 
        btn_copy1 = Button(self, text="Agilent", command=lambda: self.controller.show_frame("agilent"), state="normal")
        btn_copy1.pack(side=tk.LEFT , padx=5)

        btn_copy2 = Button(self, text="Welco", command=lambda: self.controller.show_frame("welco"), state="normal")
        btn_copy2.pack(side=tk.LEFT, padx=5)

        btn_copy3 = Button(self, text="Electrolux", command=lambda: self.controller.show_frame("electrolux"), state="normal")
        btn_copy3.pack(side=tk.LEFT, padx=5)

        btn_copy4 = Button(self, text="TAAT", command=lambda: self.controller.show_frame("taat"), state="normal")
        btn_copy4.pack(side=tk.LEFT, padx=5)


