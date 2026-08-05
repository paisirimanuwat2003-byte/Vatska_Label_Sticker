import tkinter as tk
from tkinter.ttk import *
from tkinter import filedialog, messagebox

import pandas as pd
import pyperclip
from datetime import date
import os


class welco(Frame):
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

        # Preview text widget (ensure it's defined before use in select_file/process_excel)
        text_preview = tk.Text(self, height=10, width=60)

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

        def process_excel(filepath, job_number):
            global copy_data_1, copy_data_2, copy_data_3, copy_data_4, copy_data_5, copy_data_6, copy_data_7, copy_data_8, copy_amount

            try:
                # Read the workbook
                workbook = pd.read_excel(filepath, sheet_name=None, header=None)  # Read all sheets into a dictionary of DataFrames and disable header inference
                # --- EXTRACION LOGIC ---
                # Note: pandas uses 0-based indexing. 
                # Row 1 in Excel = Index 0. Column A = Index 0, Column B = Index 1, etc.
                
                # 1. Button 1 Data: Job .1 sheet
                if  f'ใบงาน Agilent_JOB {job_number}.1' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Agilent_JOB {job_number}.1']

                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[8, 4]      # Row 9, Column D
                    qty = df.iloc[16, 3]      # Row 17, Column D

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""
                    A G I L E N T

                    Part Number:     {part_number}
                    Description:       {Description}
                    DOM:                 {Dom}
                    Lot NO:              {lot_no}
                    QTY:                  {qty}
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_1
                    copy_data_1 = final_copied_text

                
                # 2. Button 2 Data: Job .2 sheet
                if f'ใบงาน Agilent_JOB {job_number}.2' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Agilent_JOB {job_number}.2']

                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[8, 4]      # Row 9, Column D
                    qty = df.iloc[16, 3]      # Row 17, Column D

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""
                    A G I L E N T

                    Part Number:     {part_number}
                    Description:       {Description}
                    DOM:                 {Dom}
                    Lot NO:              {lot_no}
                    QTY:                  {qty}
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_2
                    copy_data_2 = final_copied_text

                # 3. Button 3 Data: Job .3 sheet
                if f'ใบงาน Agilent_JOB {job_number}.3' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Agilent_JOB {job_number}.3']

                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[8, 4]      # Row 9, Column D
                    qty = df.iloc[16, 3]      # Row 17, Column D

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""
                    A G I L E N T

                    Part Number:     {part_number}
                    Description:       {Description}
                    DOM:                 {Dom}
                    Lot NO:              {lot_no}
                    QTY:                  {qty}
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_3
                    copy_data_3 = final_copied_text

                # 4. Button 4 Data: Job .4 sheet
                if f'ใบงาน Agilent_JOB {job_number}.4' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Agilent_JOB {job_number}.4']

                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[8, 4]      # Row 9, Column D
                    qty = df.iloc[16, 3]      # Row 17, Column D

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""
                    A G I L E N T

                    Part Number:     {part_number}
                    Description:       {Description}
                    DOM:                 {Dom}
                    Lot NO:              {lot_no}
                    QTY:                  {qty}
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_4
                    copy_data_4 = final_copied_text
                #5. Button 5 Data: Job .5 sheet
                if f'ใบงาน Agilent_JOB {job_number}.5' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    df = workbook[f'ใบงาน Agilent_JOB {job_number}.5']

                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[8, 5]  # Row 9, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[6, 5]      # Row 7, Column F
                    qty = df.iloc[16, 3]      # Row 17, Column D

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""
                    A G I L E N T

                    Part Number:     {part_number}
                    Description:       {Description}
                    DOM:                 {Dom}
                    Lot NO:              {lot_no}
                    QTY:                  {qty}
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_5
                    copy_data_5 = final_copied_text

                # 6. Button 6 Data: large bag
                if f'ใบงาน Agilent_JOB {job_number}.1' in workbook or f'ใบงาน Agilent_JOB {job_number}' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    if f'ใบงาน Agilent_JOB {job_number}.1' in workbook:
                        df = workbook[f'ใบงาน Agilent_JOB {job_number}.1']
                        btn_copy1.config(state="normal")
                        btn_copy2.config(state="normal")
                        btn_copy3.config(state="normal")
                        btn_copy4.config(state="normal")
                        btn_copy5.config(state="normal") #enable for Job with .1, .2, .3, .4, .5 sheets

                    if f'ใบงาน Agilent_JOB {job_number}' in workbook:
                        df = workbook[f'ใบงาน Agilent_JOB {job_number}']
                        btn_copy1.config(state="disabled")
                        btn_copy2.config(state="disabled")
                        btn_copy3.config(state="disabled")
                        btn_copy4.config(state="disabled")
                        btn_copy5.config(state="disabled") #disable for Job without .1, .2, .3, .4, .5 sheets


                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[7, 5]  # Row 8, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[6, 5]      # Row 7, Column F
                    qty = "1 EA"

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""
                    A G I L E N T

                    Part Number:     {part_number}
                    Description:       {Description}
                    DOM:                 {Dom}
                    Lot NO:              {lot_no}
                    QTY:                  {qty}
                                
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_6
                    copy_data_6 = final_copied_text
                    copy_amount = df.iloc[7, 13]  # Row 8, Column N ,number of copies to print

                # 7. Button 7 Data: box
                if f'ใบงาน Agilent_JOB {job_number}.1' in workbook or f'ใบงาน Agilent_JOB {job_number}' in workbook:
                    # Assuming we are inside the process_excel() function from the previous code
                    if f'ใบงาน Agilent_JOB {job_number}.1' in workbook:
                        df = workbook[f'ใบงาน Agilent_JOB {job_number}.1']

                    if f'ใบงาน Agilent_JOB {job_number}' in workbook:
                        df = workbook[f'ใบงาน Agilent_JOB {job_number}']

                    # Create an empty string to hold our final rearranged text
                    final_copied_text = ""

                    # 1.1 EXTRACT: Pull the specific cells and save them as variables
                    part_number = df.iloc[7, 5]  # Row 8, Column F
                    Description = df.iloc[29, 1]   # Row 30, Column B
                    Dom = df.iloc[6, 1]   # Row 7, Column B
                    lot_no = df.iloc[6, 5]      # Row 7, Column F
                    lot_qty = df.iloc[7, 12]      # Row 8, Column M
                    sg_po = df.iloc[8, 12]      # Row 9, Column M
                    qty = " "

                    # 1.2 REARRANGE & FORMAT: Use an f-string to build your perfect layout
                    # The \n creates a new line. The \t creates a tab indent.
                    template = f"""         
                    A G I L E N T

                    Part Number:    {part_number}
                    Description:      {Description}
                    DOM:                {Dom}        
                    Lot NO:             {lot_no}       
                    LOT QTY:         {lot_qty}           
                    SG PO:             {sg_po}       
                    QTY:             {qty}        
                    """
                    # Add this formatted block to our final text
                    final_copied_text += template

                    # 1.3 READY FOR BUTTON: Save it to the global variable linked to your copy button
                    global copy_data_7
                    copy_data_7 = final_copied_text

                # 8. Button 8 Data: Barcode from Summary sheet
                if f'ใบงาน Agilent_JOB {job_number}.1' in workbook:
                    df_summary = workbook[f'ใบงาน Agilent_JOB {job_number}.1']
                    single_number = df_summary.iloc[6 , 5]  # Row 7, Column F
                    copy_data_8 = str(single_number) # Just the lot no as barcode as a string
                    
                # Update UI to show success
                text_preview.delete(1.0, tk.END)
                text_preview.insert(tk.END, f"Job นี้ต้องปริ้นทั้งหมด {copy_amount} ชิ้น\n\n")
                text_preview.insert(tk.END, "เลือกไฟล์ Excel สำเร็จ!\n\nใช้ปุ่มด้านล่างเพื่อคัดลอกข้อมูลที่ต้องการไปยังคลิปบอร์ด:\n\n")
                
                # Enable the buttons
                btn_copy6.config(state="normal")
                btn_copy7.config(state="normal")
                btn_copy8.config(state="normal")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process file:\n{e}")

        # --- Copy Functions for each button ---
        def copy_1():
            if copy_data_1:
                pyperclip.copy(copy_data_1)
                messagebox.showinfo("Copied", "Job .1 คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_2():
            if copy_data_2:
                pyperclip.copy(copy_data_2)
                messagebox.showinfo("Copied", "Job .2 คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_3():
            if copy_data_3:
                pyperclip.copy(copy_data_3)
                messagebox.showinfo("Copied", "Job .3 คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_4():
            if copy_data_4:
                pyperclip.copy(copy_data_4)
                messagebox.showinfo("Copied", "Job .4 คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_5():
            if copy_data_5:
                pyperclip.copy(copy_data_5)
                messagebox.showinfo("Copied", "Job .5 คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_6():
            if copy_data_6:
                pyperclip.copy(copy_data_6)
                messagebox.showinfo("Copied", "ถุงใหญ่ คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_7():
            if copy_data_7:
                pyperclip.copy(copy_data_7)
                messagebox.showinfo("Copied", "กล่อง คัดลอกไปยังคลิปบอร์ดแล้ว!")

        def copy_8():
            if copy_data_8:
                pyperclip.copy(copy_data_8)
                messagebox.showinfo("Copied", f"Barcode '{copy_data_8}' คัดลอกไปยังคลิปบอร์ดแล้ว!")

        # The 8 Copy Buttons (Disabled by default until a file is loaded)
        btn_copy1 = tk.Button(self, text="ถุงเล็ก (x.1 หรือ x)", command=copy_1, state="disabled")
        btn_copy1.pack(pady=5, padx=50)

        btn_copy2 = tk.Button(self, text="ถุงเล็ก (x.2)", command=copy_2, state="disabled")
        btn_copy2.pack(pady=5, padx=50)

        btn_home = Button(self, text="กลับหน้าแรก", command=lambda: self.controller.show_frame("welcome_page"))
        btn_home.pack(pady=20, padx=50)

