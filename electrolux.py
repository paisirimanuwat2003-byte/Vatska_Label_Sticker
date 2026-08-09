import tkinter as tk
from tkinter.ttk import *
from tkinter import filedialog, messagebox
import pandas as pd
import pyperclip
from datetime import date
import os
import barcode
from weasyprint import HTML
from barcode.writer import ImageWriter
import base64
from PIL import Image, ImageTk

class electrolux(Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
       
        # Load the image
        logo_img = Image.open(r'D:\vatska\software\icons\electrolux.jpg')
        logo_img = logo_img.resize((120, 50), Image.Resampling.LANCZOS)
        self.logo_img = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(self, image=self.logo_img, bg='white')
        logo_label.pack(pady=10)    

        # Global variables to store file path 
        file_path = ""

        def select_file():
            filepath = filedialog.askopenfilename(
                title="Select an Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if filepath:
                # Extract just the file name from the full path
                file_name = os.path.basename(filepath)

                # (Optional) Show the user which file they selected before processing
                text_preview.delete(1.0, tk.END)
                text_preview.insert(tk.END, f"Loading: {file_name}\n...")

                nonlocal file_path
                file_path = filepath

                btn_generate_label.config(state = 'normal')
                # Read the workbook
                workbook = pd.read_excel(filepath, sheet_name=None, header=None)  # Read all sheets into a dictionary of DataFrames and disable header inference

                text_preview.insert(tk.END, "เลือกไฟล์ Excel สำเร็จ!\n\nใช้ปุ่มด้านล่างเพื่อสร้างไฟล์ PDF\n\n")


        def process_excel(filepath):

            try:
                # Read the workbook
                workbook = pd.read_excel(filepath, sheet_name=None, header=None)  # Read all sheets into a dictionary of DataFrames and disable header inference
                    
                # Update UI to show success
                text_preview.delete(1.0, tk.END)

                # 1. Button 1 Data: Job .1 sheet
                if  f'ใบงาน Electrolux' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Electrolux']  # Access the specific sheet by name

                    product_code = df.iloc[31, 2]  # Rows 32, column C
                    product_desc = df.iloc[28, 2]  # Row 29, Column B
                    lot_no = df.iloc[6,5] # Row 7, Column F
                    qty = df.iloc[15, 4]  # Row 16, Column D
                    date_mfg = df.iloc[6, 2] # Row 7, Column C
                    po_no = df.iloc[8, 13] # Row 9, Column N

                    copy_amount = df.iloc[7, 13] # Row 8, Column N

                    text_preview.insert(tk.END, f"Job นี้ต้องปริ้นทั้งหมด {copy_amount} ชิ้น\n\n")

                    return product_code, product_desc, lot_no, qty, date_mfg, po_no

                return 0, 0, 0, 0, 0, 0

            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file:\n{e}")
                return 0, 0, 0, 0, 0, 0


        # Create a helper function to convert images to Base64 strings
        def get_image_b64(filepath):
            try:
                with open(filepath, "rb") as image_file:
                    encoded = base64.b64encode(image_file.read()).decode('utf-8')
                    return f"data:image/png;base64,{encoded}"
            except FileNotFoundError:
                print(f"Could not find {filepath}")
                return ""


        def generate_barcode(part_number):
            # 1. Get the Code 128 class instead of ean13
            Code128 = barcode.get_barcode_class('code128')

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
            # 2. Pass your exact string with the dash!
            my_barcode = Code128(f'{part_number}', writer=ImageWriter())

            # This saves a file named "my_barcode.png" in your folder
            my_barcode.save(r'D:\vatska\software\barcode_cache\electrolux_barcode', options=my_options)  # Ensure this path exists and is writable

        def generate_pdf():
            product_code, product_desc, lot_no, qty, date_mfg, po_no = process_excel(file_path)
            generate_barcode(product_code)  # Generate the barcode for the part number
            # Create a simple HTML template for the PDF
            html_content = f"""
            <!DOCTYPE html>
            <html>
                <head>
                    <meta charset="utf-8">
                    <title>Label Layout</title>
                    <style>
                    @page {{
                    size: 85mm 50mm; /* Set the exact physical PDF size here */
                    margin: 0; /* Remove default page margins */
                    }}
                
                    body {{
                    font-family: Arial, sans-serif;
                    margin: 0;
                    padding: 2mm; /* Use padding inside the body instead of page margins */
                    }}
                        /* The main outer border matching the image */
                        .label-container {{
                            padding-left: 5mm; /* Add some padding to the left for better alignment */
                            width: 85mm; /* Adjust width as needed */
                            height: 50mm; /* Adjust height as needed */
                            box-sizing: border-box;
                            position: relative;
                        }}

                        /* The top header text */
                        .header-text {{
                            font-size: 10pt;
                            letter-spacing: 4px; /* Matches the spaced-out 'a s d f v d' look */
                            font-weight: normal;
                        }}

                        /* Table used for perfect vertical alignment of the labels and values */
                        .data-table {{
                            padding-left: 5mm; /* Add some padding to the left for better alignment */
                            width: 85mm; /* Full width of the label */
                            border-collapse: collapse;
                            font-size: 9pt;
                            margin-bottom: 10px;
                        }}

                        .data-table td {{
                            vertical-align: top;
                        }}

                        /* Set a fixed width for the left column so the right column aligns perfectly */
                        .label-col {{
                            width: 25mm; /* Adjust width as needed */
                        }}

                        /* The barcode placeholder box with the "X" drawn using CSS gradients */
                        .barcode-placeholder {{
                            padding-left: 2mm; /* Add some padding to the left for better alignment */
                            item-align: center;
                            width: 60mm; /* Adjust width as needed */
                            height: 11mm; /* Adjust height as needed */
                            position: absolute;
                            bottom: 2mm; /* Position it at the bottom of the label */
                        }}
                        .black_line {{
                            width: 75mm;
                            height: 0.1mm;
                            background-color: black;
                        }}
                    </style>
                </head>
                <body>

                    <div class="label-container">
                        <!-- Header Text -->
                        <div class="header-text">
                        <img src = "{get_image_b64(r'./icons/electrolux.png')}" style = "width :15mm"  >
                        </div>
                        <div class="black_line"></div>
                        <!-- Data Grid -->
                        <table class="data-table">
                            <tr>
                                <td class="label-col">Product code:</td>
                                <td>{product_code}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Product Desc:</td>
                                <td>{product_desc}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Lot No.:</td>
                                <td>{lot_no}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Quantity:</td>
                                <td>{qty}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Date of Mfg:</td>
                                <td>{date_mfg}</td>
                            </tr>
                            <tr>
                                <td class="label-col">PO No.:</td>
                                <td>{po_no}</td>
                            </tr>
                        </table>
                        <div class="black_line"></div>
                        <!-- Placeholder Box -->
                        <div class="barcode-placeholder">
                            <img src="{get_image_b64(r'D:\vatska\software\barcode_cache\electrolux_barcode.png')}" style="width: 100%; height: 100%; object-fit: contain;">
                        </div>
                    </div>
                </body>
            </html>
            """
            HTML(string=html_content).write_pdf(f"Electrolux.pdf")

        # --- Copy Functions for each button ---
        def btn_generate_pdf():
            # Here you can define what happens when the "สร้างไฟล์สติกเกอร์" button is clicked
            messagebox.showinfo("Info", "สร้างไฟล์สติกเกอร์ button clicked!")
            generate_pdf()  # Call the function to generate the box PDF

        # Create the text box 
        text_preview = tk.Text(self, wrap=tk.WORD, width=45, height=10, font=("Consolas", 10))
        text_preview.pack(expand=True)

        # Select File Button
        btn_select = Button(self, text="เลือก Job order", command=select_file)
        btn_select.pack(pady=15)

        # The 8 Copy Buttons (Disabled by default until a file is loaded)
        btn_generate_label = Button(self, text="ซองหลัก", command=btn_generate_pdf, state="disabled")
        btn_generate_label.pack(pady=5)

        # return home button
        btn_home = Button(self, text="กลับหน้าแรก", command=lambda: self.controller.show_frame("welcome_page"))
        btn_home.pack(pady=20, padx=50)