import docx
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import pandas as pd
import re

def clean_filename(filename):
    # Remove invalid filename characters and clean up spaces
    # First replace tab with space
    filename = str(filename).replace('\t', ' ')
    # Remove invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    # Replace multiple spaces with single space
    filename = re.sub(r'\s+', ' ', filename)
    # Remove spaces at the beginning and end
    filename = filename.strip()
    # Replace remaining spaces with underscores
    filename = filename.replace(' ', '_')
    return filename

def fill_form(input_file, excel_file, output_folder, output_filename):
    # Read Word document
    doc = docx.Document(input_file)
    
    # Read Excel file
    df = pd.read_excel(excel_file)
    # Get data from second row (index 1), using column indexes (0-based)
    location_data = df.iloc[1, 1]     # Column B (index 1)
    intersection_data = df.iloc[1, 2]  # Column C (index 2)

    def set_cell_text(cell, text, bold=False):
        cell.text = text
        if cell.paragraphs:
            run = cell.paragraphs[0].runs[0]
            run.font.name = "Times New Roman"
            run.font.bold = bold

    def add_paragraph_to_cell(cell, text, bold=False):
        paragraph = cell.add_paragraph()
        run = paragraph.add_run(text)
        run.font.name = "Times New Roman"
        run.font.bold = bold

    # Fill out text fields
    field_data = {
        "Design Consultant:": ("NCE", True),
        "Design Engineer:": ("Jim Bui", True),
        "Phone Number:": ("(123) - 456- 7890", False),
        "Email:": ("", False),
        "Project Name:": ("Redondo Beach Curb Ramp Rehab", False),
        "Project #:": ("123.45.678", False),
        "Intersection:": (str(intersection_data), False),
        "Return position:": ("", False)
    }

    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()

                if "Design Consultant:" in cell_text:
                    set_cell_text(cell, "Design Consultant: " + field_data["Design Consultant:"][0], field_data["Design Consultant:"][1])
                elif "Design Engineer:" in cell_text:
                    set_cell_text(cell, "Design Engineer: " + field_data["Design Engineer:"][0], field_data["Design Engineer:"][1])
                elif "Phone Number:" in cell_text:
                    set_cell_text(cell, "Phone Number: " + field_data["Phone Number:"][0])
                    add_paragraph_to_cell(cell, "Email: " + field_data["Email:"][0])
                elif "Curb Ramp Location:" in cell_text:
                    set_cell_text(cell, "Curb Ramp Location:")
                    add_paragraph_to_cell(cell, str(location_data))
                    add_paragraph_to_cell(cell, "Intersection: " + str(intersection_data))
                elif "Project Name:" in cell_text:
                    set_cell_text(cell, "Project Name: " + field_data["Project Name:"][0])
                elif "Project #:" in cell_text:
                    set_cell_text(cell, "Project #: " + field_data["Project #:"][0])

    # Ensure all text in the document uses Times New Roman
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Times New Roman"

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Create clean output filename based on location
    clean_location = clean_filename(str(location_data))
    output_filename = f'Filled_Form_Location_{clean_location}.docx'
    output_path = os.path.join(output_folder, output_filename)
    
    # Save the modified document
    doc.save(output_path)
    print(f"Document has been filled out and saved as '{output_path}'")

# Define paths
input_file = 'Curb Ramp Hardship Form - City of Redondo Beach.docx'
excel_file = r'C:\Users\JLin\Downloads\New folder (3)\Exemption Memo Code - VBA\Redondo Beach - FY 22 Rehab - Curb Ramp Exemptions v02.xlsx'
output_folder = 'Exemption Memos'

# Run the function
fill_form(input_file, excel_file, output_folder, None)
