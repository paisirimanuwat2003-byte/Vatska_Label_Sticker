import tkinter as tk
from tkinter.ttk import *
from tkinter import filedialog, messagebox
import pandas as pd
import pyperclip
from datetime import date
import os
import barcode
from barcode.writer import ImageWriter
from weasyprint import HTML
import barcode
from barcode.writer import ImageWriter
import base64


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

        # 1. Create a helper function to convert images to Base64 strings
        def get_image_b64(filepath):
            try:
                with open(filepath, "rb") as image_file:
                    encoded = base64.b64encode(image_file.read()).decode('utf-8')
                    return f"data:image/png;base64,{encoded}"
            except FileNotFoundError:
                print(f"Could not find {filepath}")
                return ""


        def generate_barcode(job_number):
            # Generate a standard EAN13 barcode
            # ImageWriter ensures it saves as a standard .png file
            EAN = barcode.get_barcode_class('ean13')

            # Define your custom sizing options
            # module_width: The width of a single barcode line (default is 0.2)
            # module_height: The height of the barcode lines (default is 15.0)
            # font_size: The size of the text under the barcode (default is 10)
            # quiet_zone: The white space margin on the left/right (default is 6.5)
            my_options = {
                'module_width': 0.9,  # Makes the barcode wider/thicker
                'module_height': 11.0, # Makes the barcode taller
                'font_size': 12,       # Makes the text bigger
                'quiet_zone': 2.0,      # Reduces the white border around the image
            }

            my_barcode = EAN('123456789012', writer=ImageWriter())

            # This saves a file named "my_barcode.png" in your folder
            my_barcode.save(r'D:\vatska\software\barcode_cache\my_barcode', options=my_options)  # Ensure this path exists and is writable

        def generate_small_pdf(barcode_path, job_number):
            # Create a simple HTML template for the PDF
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    /* This @page rule defines the physical paper size and margins */
                    @page {{
                        width: 85mm;
                        height: 50mm;
                        margin: 20mm;
                        background-color: #ffffff;
                    }}
                    body {{
                        font-family: Arial, sans-serif;
                        color: #333;
                    }}
                    .header {{
                        background-color: #2c3e50;
                        color: white;
                        padding: 15px;
                        text-align: center;
                    }}
                    .layout-table {{
                        width: 100%;
                        margin-top: 20px;
                    }}
                    .layout-table td {{
                        vertical-align: top;
                        padding: 10px;
                    }}
                    .barcode-section {{
                        text-align: center;
                        margin-top: 30px;
                        padding-top: 20px;
                        border-top: 2px dashed #ccc;
                    }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h1>Product Label</h1>
                </div>

                <table class="layout-table">
                    <tr>
                        <!-- Left Column: Text Information -->
                        <td style="width: 60%;">
                            <h2>Item: Quantum Engine</h2>
                            <p><strong>SKU:</strong> QX-9920</p>
                            <p>This layout is completely custom. You can add as much text, lists, or data as you need right here.</p>
                        </td>
                        
                        <!-- Right Column: Logo/Image -->
                        <td style="width: 40%; text-align: center;">
                            <img src="{logo_path}" style="max-width: 150px; border: 1px solid #ddd;">
                        </td>
                    </tr>
                </table>

                <!-- Bottom Section: The Barcode -->
                <div class="barcode-section">
                    <h3>Tracking Barcode</h3>
                    <img src="{barcode_path}" style="height: 80px;">
                </div>
            </body>
            </html>
            """
            pdf_filename = f"{job_number}_barcode.pdf"
            HTML(string=html_content).write_pdf("final_output.pdf")

        def generate_large_pdf(barcode_path, job_number):
            # Create a simple HTML template for the PDF
            html_content = f"""
            <html>
            <body>
                <h1>Job Number: {job_number}</h1>
                <img src="{barcode_path}" alt="Barcode">
            </body>
            </html>
            """
            pdf_filename = f"{job_number}_barcode.pdf"
            HTML(string=html_content).write_pdf("final_output_large.pdf")

        def generate_box_pdf(barcode_path, job_number):
            # Create a simple HTML template for the PDF
            html_content = f"""
            <html>
            <body>
                <h1>Job Number: {job_number}</h1>
                <img src="{barcode_path}" alt="Barcode">
            </body>
            </html>
            """
            pdf_filename = f"{job_number}_barcode.pdf"
            HTML(string=html_content).write_pdf("final_output_box.pdf")

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
        def btn_generate_click():
            # Here you can define what happens when the "สร้างไฟล์สติกเกอร์" button is clicked
            messagebox.showinfo("Info", "สร้างไฟล์สติกเกอร์ button clicked!")

        # Create the text box 
        text_preview = tk.Text(self, wrap=tk.WORD, width=45, height=10, font=("Consolas", 10))
        text_preview.pack(expand=True)

        # Select File Button
        btn_select = Button(self, text="เลือกไฟล์ Excel", command=select_file)
        btn_select.pack(pady=15)

        # The 8 Copy Buttons (Disabled by default until a file is loaded)
        btn_generate_small = Button(self, text="สร้างไฟล์สติกเกอร์ซองย่อย", command=btn_generate_click, state="disabled")
        btn_generate_small.pack(pady=5)
        btn_generate_large = Button(self, text="สร้างไฟล์สติกเกอร์ซองหลัก", command=btn_generate_click, state="disabled")
        btn_generate_large.pack(pady=5)
        btn_generate_box = Button(self, text="สร้างไฟล์สติกเกอร์กล่อง", command=btn_generate_click, state="disabled")
        btn_generate_box.pack(pady=5)

        # return home button
        btn_home = Button(self, text="กลับหน้าแรก", command=lambda: self.controller.show_frame("welcome_page"))
        btn_home.pack(pady=20, padx=50)