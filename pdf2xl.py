import os
from pdf2image import convert_from_path
from PIL import Image, ImageOps
import pytesseract
import cv2
import numpy as np
from openpyxl import Workbook

# Function to preprocess the image
def preprocess_image(image):
    # Convert to grayscale
    gray_img = ImageOps.grayscale(image)

    # Convert to binary using thresholding
    threshold = 200
    binary_img = gray_img.point(lambda x: 0 if x < threshold else 255, '1')

    # Convert to RGB mode
    img_rgb = binary_img.convert('RGB')

    # Convert PIL image to NumPy array
    img_array = np.array(img_rgb)

    # Denoising with OpenCV
    denoised_img_array = cv2.fastNlMeansDenoisingColored(img_array, None, 10, 10, 7, 21)

    # Convert NumPy array back to PIL image (RGB mode)
    denoised_img = Image.fromarray(denoised_img_array)

    return denoised_img

# Function to extract data from PDF business cards
def extract_data_from_pdf(pdf_path):
    data = []
    images = convert_from_path(pdf_path, dpi=300)
    for i, image in enumerate(images, start=1):
        # Preprocess the image
        preprocessed_image = preprocess_image(image)

        # Extract text from the preprocessed image
        text = pytesseract.image_to_string(preprocessed_image, lang='eng')

        # Append the extracted text to the data list
        data.append(text.strip())

    return data

# Function to save data to an Excel file
def save_to_excel(data, excel_file):
    wb = Workbook()
    ws = wb.active

    for i, text in enumerate(data, start=1):
        ws.cell(row=i, column=1, value=text)

    wb.save(excel_file)

# Function to get the absolute path of the uploaded file
def get_uploaded_file_path(file_name):
    return os.path.join("/content", file_name)

def main():
    # Provide the file name here (without the full path)
    pdf_file_name = 'hindistan-kartvizit.pdf'
    pdf_file_path = get_uploaded_file_path(pdf_file_name)

    if not os.path.isfile(pdf_file_path):
        print(f"Error: File not found at {pdf_file_path}")
        return

    excel_file_path = 'output.xlsx'

    extracted_data = extract_data_from_pdf(pdf_file_path)
    save_to_excel(extracted_data, excel_file_path)

if __name__ == "__main__":
    main()
