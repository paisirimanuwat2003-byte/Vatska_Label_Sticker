import barcode
from barcode.writer import ImageWriter
import os
import base64

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

# 1. Create a helper function to convert images to Base64 strings
def get_image_b64(filepath):
    try:
        with open(filepath, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode('utf-8')
            return f"data:image/png;base64,{encoded}"
    except FileNotFoundError:
        print(f"Could not find {filepath}")
        return ""

# 2. Convert your images
# We will use an f-string to inject the paths to our images directly into the HTML
logo_path = get_image_b64(r"D:\vatska\software\icons\vatska_icon.png")  # Ensure this path is correct and the file exists
barcode_path = get_image_b64(r"D:\vatska\software\barcode_cache\my_barcode.png")


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
                    margin-bottom: 10px;
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
            </style>
        </head>
        <body>

            <div class="label-container">
                <!-- Header Text -->
                <div class="header-text">a s d f v d</div>

                <!-- Data Grid -->
                <table class="data-table">
                    <tr>
                        <td class="label-col">Part Number:</td>
                        <td>ascasc</td>
                    </tr>
                    <tr>
                        <td class="label-col">MPN:</td>
                        <td>ascascasc</td>
                    </tr>
                    <tr>
                        <td class="label-col">DOM:</td>
                        <td>ascascasc</td>
                    </tr>
                    <tr>
                        <td class="label-col">LOT NO:</td>
                        <td>ascascasc</td>
                    </tr>
                    <tr>
                        <td class="label-col">QTY:</td>
                        <td>ascascasc</td>
                    </tr>
                </table>

                <!-- Placeholder Box -->
                <div class="barcode-placeholder">
                    <img src="{barcode_path}" style="width: 100%; height: 100%; object-fit: contain;">
                </div>
            </div>

        </body>
        </html>
            """

from weasyprint import HTML

# Create a temporary HTML file or pass the string directly
HTML(string=html_content).write_pdf("Final_Product_Label.pdf")

print("PDF successfully generated!")