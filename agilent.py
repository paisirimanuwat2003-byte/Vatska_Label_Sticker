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
import base64
from PIL import Image, ImageTk
from utils import resource_path, get_save_path


class agilent(Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Load the image
        logo_img = Image.open(resource_path('icons\\agilent.png'))
        logo_img = logo_img.resize((120, 50), Image.Resampling.LANCZOS)
        self.logo_img = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(self, image=self.logo_img, bg='white')
        logo_label.pack(pady=10)    

        # Global variables to store the amount of label to be printed
        copy_amount = ""
        big_bag_count = ""

        # Global variables to store file path and job number
        file_path = ""
        job_number = ""

        def select_file():
            filepath = filedialog.askopenfilename(
                title="Select an Excel File",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if filepath:
                # Extract just the file name from the full path
                file_name = os.path.basename(filepath)
                if len(file_name) <=20:
                    jobNumber = file_name[14]  # Get the file name without extension for single-digit job numbers
                elif len(file_name) <=21:
                    jobNumber = file_name[14:16]  # Get the file name without extension for double-digit job numbers

                # (Optional) Show the user which file they selected before processing
                text_preview.delete(1.0, tk.END)
                text_preview.insert(tk.END, f"Loading: {file_name}\n...")

                btn_generate_large_label.config(state = 'normal')
                btn_generate_box_label.config(state = 'normal')
                # Read the workbook
                workbook = pd.read_excel(filepath, sheet_name=None, header=None)  # Read all sheets into a dictionary of DataFrames and disable header inference

                big_count = len(pd.ExcelFile(filepath).sheet_names)-2

                nonlocal file_path, job_number, big_bag_count
                file_path = filepath
                job_number = jobNumber
                big_bag_count = big_count

                for i in range(1,6):
                    if f'ใบงาน Agilent_JOB {job_number}.{i}' in workbook:
                        if i == 1:
                            btn_generate_small_1.config(state = 'normal')
                        if i == 2:
                            btn_generate_small_2.config(state = 'normal')
                        if i == 3:
                            btn_generate_small_3.config(state = 'normal')
                        if i == 4:
                            btn_generate_small_4.config(state = 'normal')
                        if i == 5:
                            btn_generate_small_5.config(state = 'normal')
                    else:
                        if i == 1:
                            btn_generate_small_1.config(state = 'disabled')
                        if i == 2:
                            btn_generate_small_2.config(state = 'disabled')
                        if i == 3:
                            btn_generate_small_3.config(state = 'disabled')
                        if i == 4:
                            btn_generate_small_4.config(state = 'disabled')
                        if i == 5:
                            btn_generate_small_5.config(state = 'disabled')

                text_preview.insert(tk.END, "เลือกไฟล์ Excel สำเร็จ!\n\nใช้ปุ่มด้านล่างเพื่อสร้างไฟล์ PDF:\n\n")

        def process_excel(filepath, job_number, x_number):

            try:
                # Read the workbook
                workbook = pd.read_excel(filepath, sheet_name=None, header=None)  # Read all sheets into a dictionary of DataFrames and disable header inference
                    
                # Update UI to show success
                text_preview.delete(1.0, tk.END)

                # 1. Button 1 Data: Job .1 sheet
                if  f'ใบงาน Agilent_JOB {job_number}.{x_number}' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Agilent_JOB {job_number}.{x_number}']  # Access the specific sheet by name

                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    mpn = df.iloc[7, 5]   # Row 8, Column F
                    model_pn = df.iloc[16, 2] # Row 17, Column C
                    description = df.iloc[29, 1]   # Row 30, Column B
                    date_cell = pd.to_datetime(df.iloc[6, 1], dayfirst=True, errors='coerce')
                    if pd.isna(date_cell):
                        dom = 'Date Not Found'
                    else:
                        dom = date_cell.strftime('%d-%b-%Y')
                    lot_no = df.iloc[6, 5]      # Row 7, Column F                    
                    lot_qty = df.iloc[7, 12]      # Row 8, Column M (placeholder)
                    sg_po = df.iloc[8, 12]      # Row 9, Column M
                    small_qty = df.iloc[16, 3]      # Row 17, Column D

                    print(df.iloc[7,12])

                    print(df.iloc[8,12])



                    copy_amount = df.iloc[7, 12] if pd.notna(df.iloc[7, 12]) else df.iloc[7, 13] # Row 8, Column M (fallback N)

                    text_preview.insert(tk.END, f"Job นี้ต้องปริ้นทั้งหมด {copy_amount} ชิ้น\n\n")

                    return part_number, mpn, model_pn, description, dom, lot_no, lot_qty, sg_po, small_qty

                elif x_number == 'large' or x_number == 'box':
                    # Assuming we are inside the process_excel() function from the previous code
                    if f'ใบงาน Agilent_JOB {job_number}.1' in workbook:
                        df = workbook[f'ใบงาน Agilent_JOB {job_number}.1']  # Access the specific sheet by name
                    else :
                        df = workbook[f'ใบงาน Agilent_JOB {job_number}']  # Access the specific sheet by name
                    
                    df1 = workbook['P2-QC กล่อง']

                    #specific rows and columns
                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    mpn = df.iloc[7, 5]   # Row 8, Column F
                    model_pn = df.iloc[16, 2] # Row 17, Column C
                    description = df.iloc[29, 1]   # Row 30, Column B
                    date_cell = pd.to_datetime(df.iloc[6, 1], dayfirst=True, errors='coerce')
                    if pd.isna(date_cell):
                        dom = 'Date Not Found'
                    else:
                        dom = date_cell.strftime('%d-%b-%Y')
                    lot_no = df.iloc[6, 5]      # Row 7, Column F
                    lot_qty = df1.iloc[9, 1]     # Row 10, Column B
                    sg_po = df.iloc[8, 12] if pd.notna(df.iloc[8, 12]) else df.iloc[8, 13]      # Row 9, Column M (fallback N)
                    small_qty = df.iloc[16, 3]      # Row 17, Column D
                    

                    copy_amount = df.iloc[7, 12] if pd.notna(df.iloc[7, 12]) else df.iloc[7, 13] # Row 8, Column M (fallback N)

                    text_preview.insert(tk.END, f"Job นี้ต้องปริ้นกล่องทั้งหมด {lot_qty} ชิ้น\n")
                    text_preview.insert(tk.END, f"Job นี้ต้องปริ้นถุงใหญ่ทั้งหมด {big_bag_count} ชิ้น\n\n")

                    return part_number, mpn, model_pn, description, dom, lot_no, lot_qty, sg_po, small_qty

                return 0, 0, 0, 0, 0, 0, 0, 0, 0

            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file:\n{e}")
                return 0, 0, 0, 0, 0, 0, 0, 0, 0


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
                'module_width': 1.2,  # Makes the barcode wider/thicker
                'module_height': 20.0, # Makes the barcode taller
                'font_size': 16,       # Makes the text bigger
                'quiet_zone': 2.0,      # Reduces the white border around the image
            }
            # 2. Pass your exact string with the dash!
            my_barcode = Code128(f'{part_number}', writer=ImageWriter())

            # This saves a file named "my_barcode.png" in your folder
            my_barcode.save(get_save_path('agilent_barcode','agilent_barcode'), options=my_options)  # Ensure this path exists and is writable

        def generate_small_pdf(x_number): # x_number is the number of the sheet, for example sheet x.1 or x.2 to x.5

            part_number, mpn, model_pn, description, dom, lot_no, lot_qty, sg_po, small_qty = process_excel(file_path, job_number, x_number)
            generate_barcode(part_number)  # Generate the barcode for the part number
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
                        font-family: 'Bahnschrift', sans-serif;
                        font-variation-settings: 'wght' 25;
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
                            letter-spacing: 0.5mm; /* Matches the spaced-out 'a s d f v d' look */
                            margin-bottom: 10px;
                            font-weight: normal;
                            margin-top: 10px;
                        }}

                        /* Table used for perfect vertical alignment of the labels and values */
                        .data-table {{
                            padding-left: 5mm; /* Add some padding to the left for better alignment */
                            width: 85mm; /* Full width of the label */
                            border-collapse: collapse;
                            font-size: 9pt;
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
                            bottom: 5mm; /* Position it at the bottom of the label */
                        }}
                    </style>
                </head>
                <body>

                    <div class="label-container">
                        <!-- Header Text -->
                        <div class="header-text">A G I L E N T</div>

                        <!-- Data Grid -->
                        <table class="data-table">
                            <tr>
                                <td class="label-col">Part Number:</td>
                                <td>{part_number}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Model:</td>
                                <td>{model_pn}</td>
                            </tr>
                            <tr>
                                <td class="label-col">DOM:</td>
                                <td>{dom}</td>
                            </tr>
                            <tr>
                                <td class="label-col">LOT NO:</td>
                                <td>{lot_no}</td>
                            </tr>
                            <tr>
                                <td class="label-col">QTY:</td>
                                <td>{small_qty} EA</td>
                            </tr>
                        </table>

                        <!-- Placeholder Box -->
                        <div class="barcode-placeholder">
                            <img src="{get_image_b64(get_save_path('agilent_barcode.png','agilent_barcode'))}" style="width: 100%; height: 100%; object-fit: contain;">
                        </div>
                        <div style="position: absolute; bottom: 5mm; right: 5mm; font-size: 4.5pt;">
                            {job_number}.{x_number}
                        </div>
                    </div>
                </body>
            </html>
            """
            HTML(string=html_content).write_pdf(get_save_path(f"ถุงย่อย {job_number}.{x_number}.pdf", 'agilent\\ถุงย่อย'))

        def generate_large_pdf():
            part_number, mpn, model_pn, description, dom, lot_no, lot_qty, sg_po, small_qty = process_excel(file_path, job_number, 'large')
            generate_barcode(mpn)  # Generate the barcode for the part number
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
                        font-family: 'Bahnschrift', sans-serif;
                        font-variation-settings: 'wght' 25;
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
                            letter-spacing: 0.5mm; /* Matches the spaced-out 'a s d f v d' look */
                            margin-bottom: 10px;
                            font-weight: normal;
                            margin-top: 10px;
                        }}

                        /* Table used for perfect vertical alignment of the labels and values */
                        .data-table {{
                            padding-left: 5mm; /* Add some padding to the left for better alignment */
                            width: 85mm; /* Full width of the label */
                            border-collapse: collapse;
                            font-size: 9pt;
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
                            bottom: 5mm; /* Position it at the bottom of the label */
                        }}
                    </style>
                </head>
                <body>

                    <div class="label-container">
                        <!-- Header Text -->
                        <div class="header-text">A G I L E N T</div>

                        <!-- Data Grid -->
                        <table class="data-table">
                            <tr>
                                <td class="label-col">Part Number:</td>
                                <td>{mpn}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Description:</td>
                                <td>{description}</td>
                            </tr>
                            <tr>
                                <td class="label-col">DOM:</td>
                                <td>{dom}</td>
                            </tr>
                            <tr>
                                <td class="label-col">LOT NO:</td>
                                <td>{lot_no}</td>
                            </tr>
                            <tr>
                                <td class="label-col">QTY:</td>
                                <td>{small_qty} EA</td>
                            </tr>
                        </table>

                        <!-- Placeholder Box -->
                        <div class="barcode-placeholder">
                            <img src="{get_image_b64(get_save_path('agilent_barcode.png','agilent_barcode'))}" style="width: 100%; height: 100%; object-fit: contain;">
                        </div>
                        <div style="position: absolute; bottom: 5mm; right: 5mm; font-size: 4.5pt;">
                            {job_number}
                        </div>
                    </div>
                </body>
            </html>
            """
            HTML(string=html_content).write_pdf(get_save_path(f"ถุงหลัก {job_number}.pdf", 'agilent\\ถุงหลัก'))

        def generate_box_pdf():
            part_number, mpn, model_pn, description, dom, lot_no, lot_qty, sg_po, small_qty = process_excel(file_path, job_number, 'box')
            generate_barcode(mpn)  # Generate the barcode for the part number
            # Create a simple HTML template for the PDF
            html_content = f"""
            <!DOCTYPE html>
            <html>
                <head>
                    <meta charset="utf-8">
                    <title>Label Layout</title>
                    <style>
                    @page {{
                    size: 150mm 100mm; /* Set the exact physical PDF size here */
                    margin: 0; /* Remove default page margins */
                    }}
                
                    body {{
                        font-family: 'Bahnschrift', sans-serif;
                        font-variation-settings: 'wght' 25;
                        margin: 0;
                        padding: 2mm; /* Use padding inside the body instead of page margins */
                    }}
                        /* The main outer border matching the image */
                        .label-container {{
                            padding-left: 5mm; /* Add some padding to the left for better alignment */
                            width: 150mm; /* Adjust width as needed */
                            height: 100mm; /* Adjust height as needed */
                            box-sizing: border-box;
                            position: relative;
                        }}

                        /* The top header text */
                        .header-text {{
                            font-size: 26pt;
                            letter-spacing: 0.5mm; /* Matches the spaced-out 'a s d f v d' look */
                            margin-bottom: 10px;
                            margin-top: 10px;
                            font-weight: normal;
                        }}

                        /* Table used for perfect vertical alignment of the labels and values */
                        .data-table {{
                            padding-left: 5mm; /* Add some padding to the left for better alignment */
                            width: 150mm; /* Full width of the label */
                            border-collapse: collapse;
                            font-size: 16pt;
                        }}

                        .data-table td {{
                            vertical-align: top;
                        }}

                        /* Set a fixed width for the left column so the right column aligns perfectly */
                        .label-col {{
                            width: 50mm; /* Adjust width as needed */
                        }}

                        /* The barcode placeholder box with the "X" drawn using CSS gradients */
                        .barcode-placeholder {{
                            padding-left: 2mm; /* Add some padding to the left for better alignment */
                            item-align: center;
                            width: 100mm; /* Adjust width as needed */
                            height: 20mm; /* Adjust height as needed */
                            position: absolute;
                            bottom: 7mm; /* Position it at the bottom of the label */
                        }}
                    </style>
                </head>
                <body>

                    <div class="label-container">
                        <!-- Header Text -->
                        <div class="header-text">A G I L E N T</div>

                        <!-- Data Grid -->
                        <table class="data-table">
                            <tr>
                                <td class="label-col">Part Number:</td>
                                <td>{mpn}</td>
                            </tr>
                            <tr>
                                <td class="label-col">Description:</td>
                                <td>{description}</td>
                            </tr>
                            <tr>
                                <td class="label-col">DOM:</td>
                                <td>{dom}</td>
                            </tr>
                            <tr>
                                <td class="label-col">LOT NO:</td>
                                <td>{lot_no}</td>
                            </tr>
                            <tr>
                                <td class="label-col">LOT QTY:</td>
                                <td>{lot_qty} PK</td>
                            </tr>
                            <tr>
                                <td class="label-col">SG PO:</td>
                                <td>{sg_po}</td>
                            </tr>
                            <tr>
                                <td class="label-col">QTY/BOX:</td>
                                <td>{lot_qty} PK</td>
                            </tr>
                        </table>

                        <!-- Placeholder Box -->
                        <div class="barcode-placeholder">
                            <img src="{get_image_b64(get_save_path('agilent_barcode.png','agilent_barcode'))}" style="width: 100%; height: 100%; object-fit: contain;">
                        </div>
                        <div style="position: absolute; bottom: 5mm; right: 5mm; font-size: 8pt;">
                            {job_number}
                        </div>
                    </div>
                </body>
            </html>
            """
            HTML(string=html_content).write_pdf(get_save_path(f"กล่อง {job_number}.pdf", 'agilent\\กล่อง'))

        # --- Copy Functions for each button ---
        def btn_generate_small(job_number, x_number):
            # Here you can define what happens when the "สร้างไฟล์สติกเกอร์" button is clicked
            messagebox.showinfo("Info", f"สร้างไฟล์สติกเกอร์ button clicked! ({job_number}.{x_number})")
            generate_small_pdf(x_number)  # Call the function to generate the small PDF for the specific sheet

        def btn_generate_large():
            # Here you can define what happens when the "สร้างไฟล์สติกเกอร์ซองหลัก" button is clicked
            messagebox.showinfo("Info", "สร้างไฟล์สติกเกอร์ซองหลัก button clicked!")
            generate_large_pdf()  # Call the function to generate the large PDF

        def btn_generate_box():
            # Here you can define what happens when the "สร้างไฟล์สติกเกอร์กล่อง" button is clicked
            messagebox.showinfo("Info", "สร้างไฟล์สติกเกอร์กล่อง button clicked!")
            generate_box_pdf()  # Call the function to generate the box PDF

        # Create the text box 
        text_preview = tk.Text(self, wrap=tk.WORD, width=45, height=10, font=("Consolas", 10))
        text_preview.pack(expand=True)

        # Select File Button
        btn_select = Button(self, text="เลือก Job order", command=select_file)
        btn_select.pack(pady=15)

        # The 8 Copy Buttons (Disabled by default until a file is loaded)
        btn_generate_small_1 = Button(self, text="ซองย่อย x.1", command=lambda: btn_generate_small(job_number, 1), state="disabled")
        btn_generate_small_1.pack(pady=5)
        btn_generate_small_2 = Button(self, text="ซองย่อย x.2", command=lambda: btn_generate_small(job_number, 2), state="disabled")
        btn_generate_small_2.pack(pady=5)
        btn_generate_small_3 = Button(self, text="ซองย่อย x.3", command=lambda: btn_generate_small(job_number, 3), state="disabled")
        btn_generate_small_3.pack(pady=5)
        btn_generate_small_4 = Button(self, text="ซองย่อย x.4", command=lambda: btn_generate_small(job_number, 4), state="disabled")
        btn_generate_small_4.pack(pady=5)
        btn_generate_small_5 = Button(self, text="ซองย่อย x.5", command=lambda: btn_generate_small(job_number, 5), state="disabled")
        btn_generate_small_5.pack(pady=5)
        btn_generate_large_label = Button(self, text="ซองหลัก", command=btn_generate_large, state="disabled")
        btn_generate_large_label.pack(pady=5)
        btn_generate_box_label = Button(self, text="กล่อง", command=btn_generate_box, state="disabled")
        btn_generate_box_label.pack(pady=5)

        # return home button
        btn_home = Button(self, text="กลับหน้าแรก", command=lambda: self.controller.show_frame("welcome_page"))
        btn_home.pack(pady=20, padx=50)