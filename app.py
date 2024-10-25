import os
import streamlit as st
from PIL import Image, UnidentifiedImageError
import pytesseract
from openpyxl import Workbook
from tqdm import tqdm
from io import BytesIO

# Streamlit app title and description
st.title("Screenshot Text Extraction")
st.write("Upload a folder containing images, and extract text from specific areas in each image into an Excel file.")

# Define the crop areas for text extraction
crop_areas = [
    (37, 26, 183, 75),        # First set of params (row 1)
    (793, 833, 970, 915),     # Second set of params (row 2)
    (787, 1215, 935, 1280),   # Third set of params (row 3)
    (461, 1225, 597, 1277)    # Fourth set of params (row 4)
]

# Define the function to extract text from a cropped image
def extract_text_from_image(image, crop_area):
    cropped_image = image.crop(crop_area)
    text = pytesseract.image_to_string(cropped_image, lang='eng')
    return text.strip()

# Preprocess image (convert to grayscale)
def preprocess_image(image):
    grayscale_image = image.convert("L")
    return grayscale_image

# Function to process a screenshot and extract data
def process_screenshot(image, filename):
    try:
        # Convert image to grayscale and extract text for each crop area
        preprocessed_image = preprocess_image(image)
        extracted_data = [extract_text_from_image(preprocessed_image, area) for area in crop_areas]
        return [filename, "Date Extracted"] + extracted_data
    except UnidentifiedImageError:
        st.warning(f"Error: Could not identify {filename}, skipping this file.")
        return None

# File upload
uploaded_files = st.file_uploader("Upload images", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if uploaded_files:
    st.info("Processing uploaded images...")

    # Initialize a new workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    ws.append(["Screenshot Name", "Date", "Param1", "Param2", "Param3", "Param4"])

    # Set up a progress bar
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)

    # Process each image file in the uploaded files
    for i, uploaded_file in enumerate(uploaded_files):
        image = Image.open(uploaded_file)
        filename = uploaded_file.name
        result = process_screenshot(image, filename)
        if result:
            ws.append(result)

        # Update the progress bar
        progress_bar.progress((i + 1) / total_files)

    # Save the workbook to a BytesIO stream to download
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)
    st.success("Data extraction complete.")
    st.download_button(
        label="Download Excel File",
        data=excel_output,
        file_name="extracted_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
