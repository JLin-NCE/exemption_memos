import docx
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import pandas as pd
import re
import win32com.client
import logging
from typing import Dict, Any

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class FormFiller:
    def __init__(self, input_file: str, excel_file: str, output_folder: str):
        self.input_file = os.path.abspath(input_file)
        self.excel_file = os.path.abspath(excel_file)
        self.output_folder = os.path.abspath(output_folder)

    @staticmethod
    def format_street_name(text: str) -> str:
        """Format street names from all caps to proper case"""
        # Dictionary of common street abbreviations and their proper format
        street_abbr = {
            'RD': 'Rd',
            'ST': 'St',
            'AVE': 'Ave',
            'BLVD': 'Blvd',
            'LN': 'Ln',
            'DR': 'Dr',
            'CT': 'Ct',
            'PL': 'Pl',
            'TER': 'Ter',
            'PKY': 'Pky',
            'CIR': 'Cir',
            'HWY': 'Hwy',
            'WAY': 'Way'
        }

        words = text.strip().split()
        formatted_words = []

        for word in words:
            upper_word = word.upper()
            if upper_word in street_abbr:
                formatted_words.append(street_abbr[upper_word])
            else:
                formatted_words.append(word.capitalize())

        return ' '.join(formatted_words)

    @staticmethod
    def extract_and_format_location(location_string: str) -> tuple[str, str]:
        """Extract location number and format remaining text"""
        try:
            # Find the decimal point
            decimal_index = str(location_string).find('.')
            
            if decimal_index != -1:
                # Get everything after decimal
                after_decimal = str(location_string)[decimal_index + 1:]
                
                # Remove parentheses and their contents
                no_parens = re.sub(r'\s*\([^)]*\)', '', after_decimal)
                
                # Extract the location number
                number_match = re.match(r'(\d+)', no_parens)
                if number_match:
                    number = number_match.group(1)
                    remaining_text = no_parens[len(number):].strip()
                    formatted_text = FormFiller.format_street_name(remaining_text) if remaining_text else ""
                    return number, formatted_text
                else:
                    formatted_text = FormFiller.format_street_name(no_parens)
                    return "", formatted_text
            else:
                formatted_text = FormFiller.format_street_name(str(location_string))
                return "", formatted_text
                
        except Exception as e:
            logger.error(f"Error processing location string: {str(e)}")
            return str(location_string).strip(), ""

    def clean_filename(self, filename: str) -> str:
        """Clean filename for Windows compatibility"""
        filename = str(filename).replace('\t', ' ')
        filename = re.sub(r'[<>:"/\\|?*]', '', filename)
        filename = re.sub(r'\s+', ' ', filename)
        filename = filename.strip()
        filename = filename.replace(' ', '_')
        return filename

    def set_cell_text(self, cell: Any, text: str, bold: bool = False):
        """Set text in a table cell with formatting"""
        cell.text = text
        if cell.paragraphs:
            run = cell.paragraphs[0].runs[0]
            run.font.name = "Times New Roman"
            run.font.bold = bold

    def add_paragraph_to_cell(self, cell: Any, text: str, bold: bool = False):
        """Add a new paragraph to a cell with formatting"""
        paragraph = cell.add_paragraph()
        run = paragraph.add_run(text)
        run.font.name = "Times New Roman"
        run.font.bold = bold

    def process_checkboxes(self, doc_path: str):
        """Process checkboxes in the document"""
        word_app = None
        doc = None
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            doc = word_app.Documents.Open(doc_path)
            
            # Log all checkboxes and their states
            for i in range(1, doc.FormFields.Count + 1):
                field = doc.FormFields.Item(i)
                if field.Type == 71:  # Checkbox type
                    current_state = "Checked" if field.CheckBox.Value == 1 else "Unchecked"
                    logger.info(f"Checkbox {i}: {field.Name} - {current_state}")
            
            # Example: Check the first checkbox
            if doc.FormFields.Count >= 1:
                first_checkbox = doc.FormFields.Item(1)
                if first_checkbox.Type == 71:
                    first_checkbox.CheckBox.Value = 1
                    logger.info("Set first checkbox (Fire Hydrant) to checked")
            
            doc.Save()
            logger.info("Checkboxes processed and document saved")
            
        except Exception as e:
            logger.error(f"Error processing checkboxes: {str(e)}")
        finally:
            try:
                if doc:
                    doc.Close(False)
                if word_app:
                    word_app.Quit()
            except Exception as e:
                logger.error(f"Error closing Word application: {str(e)}")

    def process_document(self):
        """Main document processing method"""
        try:
            # Read Excel data
            logger.info(f"Reading Excel file: {self.excel_file}")
            df = pd.read_excel(self.excel_file)
            
            # Get and process location data
            raw_location = str(df.iloc[1, 1])  # Column B
            location_number, location_text = self.extract_and_format_location(raw_location)
            
            # Combine number and formatted text
            location_data = f"{location_number} {location_text}".strip()
            
            # Get and format intersection data
            raw_intersection = str(df.iloc[1, 2])  # Column C
            intersection_data = self.format_street_name(raw_intersection)
            
            logger.info(f"Raw location: {raw_location}")
            logger.info(f"Processed location: {location_data}")
            logger.info(f"Raw intersection: {raw_intersection}")
            logger.info(f"Processed intersection: {intersection_data}")
            
            # Read document
            logger.info("Processing Word document...")
            doc = docx.Document(self.input_file)
            
            # Field data
            field_data = {
                "Design Consultant:": ("NCE", True),
                "Design Engineer:": ("Jim Bui", True),
                "Phone Number:": ("(123) - 456- 7890", False),
                "Email:": ("", False),
                "Project Name:": ("Redondo Beach Curb Ramp Rehab", False),
                "Project #:": ("123.45.678", False)
            }

            # Process tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text.strip()
                        
                        if "Design Consultant:" in cell_text:
                            self.set_cell_text(cell, "Design Consultant: " + field_data["Design Consultant:"][0], 
                                             field_data["Design Consultant:"][1])
                        elif "Design Engineer:" in cell_text:
                            self.set_cell_text(cell, "Design Engineer: " + field_data["Design Engineer:"][0], 
                                             field_data["Design Engineer:"][1])
                        elif "Phone Number:" in cell_text:
                            self.set_cell_text(cell, "Phone Number: " + field_data["Phone Number:"][0])
                            self.add_paragraph_to_cell(cell, "Email: " + field_data["Email:"][0])
                        elif "Curb Ramp Location:" in cell_text:
                            self.set_cell_text(cell, "Curb Ramp Location:")
                            self.add_paragraph_to_cell(cell, location_data)
                        elif "Intersection:" in cell_text:
                            self.set_cell_text(cell, "Intersection:")
                            self.add_paragraph_to_cell(cell, intersection_data)
                            self.add_paragraph_to_cell(cell, "Return position:")
                        elif "Project Name:" in cell_text:
                            self.set_cell_text(cell, "Project Name: " + field_data["Project Name:"][0])
                        elif "Project #:" in cell_text:
                            self.set_cell_text(cell, "Project #: " + field_data["Project #:"][0])

            # Set font for all paragraphs
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"

            # Create output folder if needed
            os.makedirs(self.output_folder, exist_ok=True)

            # Save document
            clean_location = self.clean_filename(location_data)
            output_filename = f'Filled_Form_Location_{clean_location}.docx'
            output_path = os.path.join(self.output_folder, output_filename)
            doc.save(output_path)
            logger.info(f"Saved initial document to: {output_path}")
            
            # Process checkboxes
            self.process_checkboxes(output_path)
            
            logger.info("Document processing completed successfully")
            
        except Exception as e:
            logger.error(f"Error processing document: {str(e)}", exc_info=True)
            raise

if __name__ == "__main__":
    # Define paths
    input_file = 'Curb Ramp Hardship Form - City of Redondo Beach.docx'
    excel_file = r'C:\Users\JLin\Downloads\New folder (3)\Exemption Memo Code - VBA\Redondo Beach - FY 22 Rehab - Curb Ramp Exemptions v02.xlsx'
    output_folder = 'Exemption Memos'
    
    try:
        form_filler = FormFiller(input_file, excel_file, output_folder)
        form_filler.process_document()
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}", exc_info=True)
