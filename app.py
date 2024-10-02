import os
import re
from PIL import Image, ImageDraw
import pytesseract
from openpyxl import Workbook
import streamlit as st
from io import BytesIO

# Set tesseract_cmd to the correct location of the Tesseract executable (you can adjust this path based on your setup)
pytesseract.pytesseract.tesseract_cmd = r'/usr/bin/tesseract'  # For Linux, or the appropriate path for your OS

# Function to extract text from an image for a given crop area
def extract_text_from_image(image, crop_area):
    cropped_image = image.crop(crop_area)
    text = pytesseract.image_to_string(cropped_image, lang='eng')
    return text.strip()

# Function to extract the date from the filename using regex
def extract_date_from_filename(filename):
    match = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    if match:
        return match.group(1)
    return None

# Function to visualize the extraction areas on the image
def visualize_extraction_areas(image, crop_areas):
    draw = ImageDraw.Draw(image)
    for area in crop_areas:
        draw.rectangle(area, outline="red", width=3)  # Red rectangles for visualization
    return image

# Streamlit UI
st.title("Screenshot Text Extraction to Excel")

# Upload files
uploaded_files = st.file_uploader("Upload Screenshot Files", accept_multiple_files=True, type=["png", "jpg"])

# Check if files are uploaded
if uploaded_files:
    # Show the first image as a sample
    sample_image = uploaded_files[0]
    image = Image.open(sample_image)

    # Display sample image
    st.image(image, caption="Sample Screenshot", use_column_width=True)

    # Define the crop areas
    crop_areas = [
        (45, 29, 201, 72),        # Data 1
        (813, 877, 1011, 1000),   # Data 2
        (580, 1414, 752, 1472),   # Data 3
        (842, 1419, 996, 1474),   # Data 4
        (73, 1650, 218, 1731)     # Data 5
    ]

    # Visualize the extraction areas on the sample image
    visualized_image = visualize_extraction_areas(image.copy(), crop_areas)
    st.image(visualized_image, caption="Visualization of Extraction Areas", use_column_width=True)

    # Process button
    if st.button("Process Screenshots and Download Excel"):
        # Create an in-memory workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Data"

        # Add column headers
        ws.append(["Screenshot Name", "Date", "Data 1", "Data 2", "Data 3", "Data 4", "Data 5"])

        # Process each uploaded file
        for uploaded_file in uploaded_files:
            image = Image.open(uploaded_file)
            filename = uploaded_file.name

            # Extract date from filename
            date = extract_date_from_filename(filename)

            # Extract text for each crop area
            extracted_data = [extract_text_from_image(image, area) for area in crop_areas]

            # Append the data row (with the screenshot name and date)
            ws.append([filename, date] + extracted_data)

        # Save the Excel file to a BytesIO object
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Provide download link for the Excel file
        st.download_button(
            label="Download Excel",
            data=output,
            file_name="extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
