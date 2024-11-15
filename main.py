import docx
import pandas as pd
import os
from docx.shared import Pt

def replace_first_instance(doc_path, excel_path, output_folder="output"):
    """Replace values while preserving labels"""
    print("\n=== Starting Document Processing ===")
    
    # Read first row from Excel (skipping header)
    print("Reading Excel values...")
    df = pd.read_excel(excel_path, skiprows=1)
    location = str(df.iloc[0, 1])  # First row, column B
    intersection = str(df.iloc[0, 2])  # First row, column C
    
    print(f"Values to insert:")
    print(f"Location: '{location}'")
    print(f"Intersection: '{intersection}'")
    
    # Create new document
    doc = docx.Document(doc_path)
    
    print("\nReplacing values and setting fonts...")
    # Replace first instances
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if 'CR location' in cell.text:
                    print(f"Found 'CR location' in cell: '{cell.text}'")
                    for para in cell.paragraphs:
                        if 'CR location' in para.text:
                            para.clear()
                            run = para.add_run(location)
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(12)
                            print(f"Replaced with: '{location}' in Times New Roman 12pt")
                
                if 'I location' in cell.text:
                    print(f"Found 'I location' in cell: '{cell.text}'")
                    for para in cell.paragraphs:
                        if 'I location' in para.text:
                            # Preserve "Intersection:" label
                            para.clear()
                            label_run = para.add_run("Intersection: ")
                            label_run.font.name = "Times New Roman"
                            label_run.font.size = Pt(12)
                            
                            value_run = para.add_run(intersection)
                            value_run.font.name = "Times New Roman"
                            value_run.font.size = Pt(12)
                            print(f"Replaced with: 'Intersection: {intersection}' in Times New Roman 12pt")
    
    # Save modified document
    os.makedirs(output_folder, exist_ok=True)
    output_path = os.path.join(output_folder, 'Location_001.docx')
    doc.save(output_path)
    print(f"\nSaved modified document to: {output_path}")

if __name__ == "__main__":
    doc_path = "Curb Ramp Hardship Form - City of Redondo Beach.docx"
    excel_path = r"C:\Users\JLin\Downloads\Exemption Memo Code - Python\Redondo Beach - FY 22 Rehab - Curb Ramp Exemptions v02.xlsx"
    
    replace_first_instance(doc_path, excel_path)
