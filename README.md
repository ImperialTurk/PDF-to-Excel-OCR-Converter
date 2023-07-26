# PDF-to-Excel-OCR-Converter
A Python script that extracts data from business cards stored in a PDF file and saves it to an Excel spreadsheet using OCR. 
I utilized Google Colab for development; therefore, file paths and the 'get_uploaded_file_path' function may not work properly in other environments. Please create your own function to handle file paths and replace them accordingly.
# Installation
To run this project, you need to have the following dependencies installed:

* Python 3 (3.6 or higher)
* pytesseract
* tesseract-ocr
* poppler-utils
* PyPDF2
* Pillow
* opencv-python
* openpyxl

You can install the Python dependencies using the following command:

```pip install pytesseract tesseract-ocr poppler-utils PyPDF2 Pillow opencv-python openpyxl```

# Usage
1- Clone the repository to your local machine:

```git clone https://github.com/ImperialTurk/PDF-to-Excel-OCR-Converter.git```

```cd PDF-to-Excel-OCR-Converter ```

2- Place your PDF file containing business cards in the project directory.

3- Modify the pdf_file_name variable in the Python script (pdf2xl.py) to match the name of your PDF file.

4- Run the Python script:

``` python pdf2xl.py ```

The script will process the PDF, extract data from the business cards using OCR, and save it to an Excel file (output.xlsx) in the project directory.



