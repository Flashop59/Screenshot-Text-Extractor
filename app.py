import os
import re
from PIL import Image, ImageDraw
import easyocr
import streamlit as st
from io import BytesIO
from openpyxl import Workbook

# Initialize the OCR reader
reader = easyocr.Reader(['en'])

# Function to extract text from an image for a given crop area
def extract_text_from_image(image):
    result = reader.readtext(image)
    extracted_text = " ".join([res[1] for res in result])
    return extracted_text

# Function to extract the date from the filename using regex
def extract_date_from_filename(filename):
    match = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    if match:
        return match.group(1)
    return None

# Streamlit app
st.title("Screenshot Text Extraction to Excel")

# File uploader
uploaded_files = st.file_uploader("Upload Screenshot Files", accept_multiple_files=True, type=["png", "jpg"])

# Display first uploaded image as a sample
if uploaded_files:
    sample_image = uploaded_files[0]
    image = Image.open(sample_image)
    st.image(image, caption="Sample Screenshot", use_column_width=True)

    # Define your crop areas (adjust based on the image)
    crop_areas = [
        (45, 29, 201, 72),        # Data 1
        (813, 877, 1011, 1000),   # Data 2
        (580, 1414, 752, 1472),   # Data 3
        (842, 1419, 996, 1474),   # Data 4
        (73, 1650, 218, 1731)     # Data 5
    ]

    # Process button
    if st.button("Process Screenshots and Download Excel"):
        # Create in-memory workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Data"

        # Add headers
        ws.append(["Screenshot Name", "Date", "Data 1", "Data 2", "Data 3", "Data 4", "Data 5"])

        # Process each uploaded file
        for uploaded_file in uploaded_files:
            image = Image.open(uploaded_file)
            filename = uploaded_file.name

            # Extract the date from filename
            date = extract_date_from_filename(filename)

            # Extract text for each crop area
            extracted_data = [extract_text_from_image(image) for _ in crop_areas]

            # Append the data
            ws.append([filename, date] + extracted_data)

        # Save to BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Provide download link
        st.download_button(
            label="Download Excel",
            data=output,
            file_name="extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
