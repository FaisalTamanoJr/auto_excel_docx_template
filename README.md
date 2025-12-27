# auto_LOA_LOR

## What this program does

This script is used to extract values from an Excel Sheet (.xlsx) then apply it to a Word Document (.docx) and produce multiple outputs based on the template and Excel values.

## Installation

1. Install [Python](https://wiki.python.org/moin/BeginnersGuide/Download) and add it into your PATH
2. [Activate the virtual environment](https://www.w3schools.com/python/python_virtualenv.asp)
3. Enter `pip install -r requirements.txt` to install the dependencies

## Usage
```
usage: autoLOALOR.py [-h] [-o OUTPUT_DIRECTORY]
                     doc_template excel_sheet suffix

Applies values in an Excel file to a Word Template and Export them to PDF

positional arguments:
  doc_template          .docx template for Excel values to substitute to
  excel_sheet           .xlsx file containing columns to represent document
                        template variable and rows to represent the different
                        values associated with the template variable
  suffix                the suffix added to the filename of the outputs

options:
  -h, --help            show this help message and exit
  -o OUTPUT_DIRECTORY, --output_directory OUTPUT_DIRECTORY
                        Optional directory for output. Default is the current
                        directory.
```
### Example inputs:

Example sample docx file and .xlsx file are `sample_document_template.docx` and `sample_sheet.xlsx` in the project folder.

Sample outputs can be found in the `sample_outputs` directory

The command used to produced the outputs:

```
python .\autoLOALOR.py .\sample_document_template.docx .\sample_sheets.xlsx _sample -o .\sample_outputs\
```