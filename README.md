# SaveAsPDF VBA Macro

## Overview

The `SaveAsPDF` macro is a Visual Basic for Applications (VBA) script designed to automate the process of saving the active Excel worksheet as a PDF file. The generated PDF file is named based on the current date and a serial number, ensuring that each file has a unique name.

## Features

- Automatically generates a PDF file name using the current date in MMDDYYYY format.
- Starts numbering from a predefined serial number (37) and increments it if a file with the same name already exists.
- Exports the active worksheet to a specified directory as a PDF file.
- Includes error handling to notify the user in case of any issues during the saving process.

## Usage Instructions

1. **Open Excel**: Open the Excel workbook that contains the worksheet you want to save as a PDF.
2. **Access the VBA Editor**:
   - Press `ALT + F11` to open the VBA editor.
3. **Insert the Macro**:
   - In the VBA editor, insert a new module by right-clicking on any of the items in the Project Explorer, selecting `Insert`, and then `Module`.
   - Copy and paste the `SaveAsPDF` subroutine code into the new module.
4. **Run the Macro**:
   - Close the VBA editor and return to Excel.
   - Press `ALT + F8`, select `SaveAsPDF`, and click `Run` to execute the macro.
5. **Check the Output**: The PDF will be saved in the specified directory: `C:\Users\user\OneDrive\Documents\OA Tutor\Docs\Finance\Invoice\All\`.

## File Naming Convention

The generated PDF files will follow this naming format: MMDDYYYY - Invoice XXX.pdf

Where `MMDDYYYY` is the current date and `XXX` is a three-digit serial number starting from 037. If a file with the same name already exists, the serial number will increment automatically.

## Error Handling

If an error occurs during the PDF saving process, a message box will display the error description to help diagnose the issue.

## Requirements

- Microsoft Excel with VBA support.
- Basic knowledge of how to run macros in Excel.

## License

This script is provided as-is. Feel free to modify and use it according to your needs.

## Contact

For any questions or feedback, please contact:
- Email: [tech@olaarowolo.com](mailto:tech@olaarowolo.com)
- Phone: 07487397751