import tkinter as tk
from tkinter.ttk import *
from tkinter import filedialog, messagebox
import pandas as pd
import pyperclip
from datetime import date
import os
import barcode
from barcode.writer import ImageWriter


class agilent(Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        # Global variables to store the data for our 3 buttons
        copy_data_1 = ""
        copy_data_2 = ""
        copy_data_3 = ""
        copy_data_4 = ""
        copy_data_5 = ""
        copy_data_6 = ""
        copy_data_7 = ""
        copy_data_8 = ""
        copy_amount = ""

        def select_file():
            filepath = filedialog.askopenfilename(
                title="Select an Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if filepath:
                # Extract just the file name from the full path
                file_name = os.path.basename(filepath)
                if len(file_name) <=20:
                    job_number = file_name[14]  # Get the file name without extension
                elif len(file_name) <=21:
                    job_number = file_name[14:16]  # Get the file name without extension

                # (Optional) Show the user which file they selected before processing
                text_preview.delete(1.0, tk.END)
                text_preview.insert(tk.END, f"Loading: {file_name}\n...")
                
                # You can even pass the file_name to your process function if you want to include 
                # it in your formatted copied text!
                process_excel(filepath, job_number)  # Pass the file name as job_number


        def generate_barcode(job_number):
            # Generate the barcode
            CODE128 = barcode.get_barcode_class('code128')
            barcode_instance = CODE128(job_number, writer=ImageWriter())
            barcode_filename = f"{job_number}_barcode"
            barcode_path = os.path.join(os.getcwd(), f"{barcode_filename}.png")
            barcode_instance.save(barcode_path)
            return barcode_path

        def process_excel(filepath, job_number):

            try:
                # Read the Excel file
                df = pd.read_excel(filepath)
                    
                # Update UI to show success
                text_preview.delete(1.0, tk.END)
                text_preview.insert(tk.END, f"Job นี้ต้องปริ้นทั้งหมด {copy_amount} ชิ้น\n\n")
                text_preview.insert(tk.END, "เลือกไฟล์ Excel สำเร็จ!\n\nใช้ปุ่มด้านล่างเพื่อคัดลอกข้อมูลที่ต้องการไปยังคลิปบอร์ด:\n\n")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file:\n{e}")

        # --- Copy Functions for each button ---
        def copy_1():
            if copy_data_1:
                pyperclip.copy(copy_data_1)
                messagebox.showinfo("Copied", "Job .1 คัดลอกไปยังคลิปบอร์ดแล้ว!")

        # Create the text box 
        text_preview = tk.Text(self, wrap=tk.WORD, width=45, height=10, font=("Consolas", 10))
        text_preview.pack(expand=True)

        # Select File Button
        btn_select = Button(self, text="เลือกไฟล์ Excel", command=select_file)
        btn_select.pack(pady=15)

        # The 8 Copy Buttons (Disabled by default until a file is loaded)
        btn_copy1 = Button(self, text="สร้างไฟล์สติกเกอร์", command=copy_1, state="disabled")
        btn_copy1.pack(pady=5, padx=50)

        # return home button
        btn_home = Button(self, text="กลับหน้าแรก", command=lambda: self.controller.show_frame("welcome_page"))
        btn_home.pack(pady=20, padx=50)