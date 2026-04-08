# Excel Image Reader

This is more of a pet project of main, it is not very clean.

Excel Image Reader is a Python desktop tool that scans images with Tesseract OCR and exports the results to an Excel workbook. It uses Google Tesseract for text extraction, and openpyxl for interfacing with excel using python.

## General

This project is meant to catalog pictures of equipment serial numbers into a Excel sheet. For each item, OCR is used to attempt to read a serial value from the image and creates an Excel sheet with the assumed serial value plus links and confidence values.

## Requirements

- Python
- Tesseract OCR installed
- PyQt6
- Pillow
- pytesseract
- openpyxl

## Program Process

1. Pick the folder containing your images.
2. Pick an output folder.
3. Enter a folder name for the generated results.
4. Choose the image order:
   - ModelSerial
   - SerialModel
5. Adjust the OCR confidence threshold if needed.
6. Run the program.

The program output is an Excel workbook named Dataset.xlsx.

## Output

When the run finishes, the program creates a directory with the following:
- A Data folder containing copied images
- A Dataset.xlsx file with:
  - Serial Value
  - Model Image
  - Serial Image

## Notes

- The project currently assumes a Tesseract install path at "C:\Program Files\Tesseract-OCR\tesseract.exe".
  - The program will try to install tesseract if it cannot find it. 
- The OCR model needs additional tuning for better accuracy.
- There is a manual; however, it is fairly barebones.

