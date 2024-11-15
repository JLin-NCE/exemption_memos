# Curb Ramp Form Automation

This Python script automates the filling of Curb Ramp Hardship Exception Forms for the City of Redondo Beach Public Works Department.

## Setup

### Prerequisites
```bash
pip install python-docx pandas pywin32
```

### File Structure
```
project/
│
├── main.py                 # Main script
├── README.md              # This documentation
├── templates/             # Template files
│   └── Curb Ramp Hardship Form - City of Redondo Beach.docx
│
└── Exemption Memos/       # Output directory for filled forms
```

## Checkbox Mappings

### Design Hardships Section
```python
CHECKBOX_MAPPING = {
    # Left Column
    1: "fire_hydrant",          # Fire Hydrant
    2: "utility_pull_box",      # Utility Pull Box or Vault
    3: "utility_power_pole",    # Utility/Power Pole
    4: "catch_basin",           # Catch Basin
    5: "traffic_signal_pole",   # Traffic Signal Pole
    6: "street_grade",          # Street Grade
    7: "other_reasons",         # Other Reasons (describe)
    
    # Right Column
    8: "sidewalk_cross_slope",  # Sidewalk Cross Slope
    9: "no_legal_crossing",     # No Legal Crossing
    10: "building_entrance",    # Building Entrance or Driveway
    11: "private_fence",        # Private fence / wall
    12: "narrow_sidewalks",     # Narrow Sidewalks
    13: "right_of_way_conflict" # Right-of-Way or Easement Conflict
}

# Design Exceptions Section
EXCEPTIONS_MAPPING = {
    # Left Column
    14: "narrow_ramp",          # Less than 48" wide ramp
    15: "landing_clearance",    # Less than 48" landing clearance
    16: "steep_ramp",          # Greater than 8.3% (1:12) ramp slope
    17: "transition_slope",     # Greater than 2.1% (1:48) transition area slope
    18: "wing_flare_slope",    # Greater than 10.0% (1:10) wing/flare slope
    19: "other_exceptions",     # Other Exceptions (describe)
    
    # Right Column
    20: "blended_ramp",        # Blended/merged curb ramps/driveway
    21: "narrow_width",        # Clear Width less than 48"
    22: "narrow_width_short"   # Clear Width less than 48", but greater than 32" for a length of less than 24"
}
```

## Street Abbreviation Mappings
```python
STREET_ABBREVIATIONS = {
    'RD': 'Rd',    # Road
    'ST': 'St',    # Street
    'AVE': 'Ave',  # Avenue
    'BLVD': 'Blvd', # Boulevard
    'LN': 'Ln',    # Lane
    'DR': 'Dr',    # Drive
    'CT': 'Ct',    # Court
    'PL': 'Pl',    # Place
    'TER': 'Ter',  # Terrace
    'PKY': 'Pky',  # Parkway
    'CIR': 'Cir',  # Circle
    'HWY': 'Hwy',  # Highway
    'WAY': 'Way'   # Way
}
```

## Excel File Format

The script expects an Excel file with the following structure:
- Column B (index 1): Location data (format: "XX.YYY Street Name")
- Column C (index 2): Intersection data

### Example Excel Data:
```
| Column A | Column B                              | Column C      |
|----------|---------------------------------------|---------------|
| Header   | Header                                | Header        |
| Data     | 02.BATAAN RD (AVIATION BLVD-VAIL AVE) | Blossom Ln    |
```

## Text Formatting Rules

### Location Processing
1. Extracts number after decimal point
2. Removes parentheses and their contents
3. Converts street names to proper case
4. Formats street abbreviations according to mapping

Example:
```python
Input: "02.BATAAN RD (AVIATION BLVD-VAIL AVE)"
Output: "Bataan Rd"
```

### Form Fields
```python
FIELD_DATA = {
    "Design Consultant:": ("NCE", True),
    "Design Engineer:": ("Jim Bui", True),
    "Phone Number:": ("(123) - 456- 7890", False),
    "Email:": ("", False),
    "Project Name:": ("Redondo Beach Curb Ramp Rehab", False),
    "Project #:": ("123.45.678", False)
}
```

## Usage

```python
from form_filler import FormFiller

input_file = 'Curb Ramp Hardship Form - City of Redondo Beach.docx'
excel_file = 'path/to/your/excel/file.xlsx'
output_folder = 'Exemption Memos'

form_filler = FormFiller(input_file, excel_file, output_folder)
form_filler.process_document()
```

## Error Handling

The script includes handling for:
- File permission issues
- File access conflicts
- Excel data formatting
- Word document processing
- Checkbox manipulation

## Output

Generated files will be saved as:
```
Filled_Form_Location_{location}.docx
```
Example: `Filled_Form_Location_Bataan_Rd.docx`

## Logging

The script logs:
- Input/output operations
- Data processing steps
- Checkbox states
- Errors and exceptions

Logs include:
- Raw and processed location data
- Checkbox operations
- File operations
- Error messages

## Troubleshooting

Common issues:
1. Permission Denied
   - Close any open Word documents
   - Run as administrator
   - Check folder permissions

2. Excel Data Format
   - Ensure correct column structure
   - Check for proper data formatting

3. Word Document
   - Ensure template is not read-only
   - Close any open instances
   - Check form field structure
