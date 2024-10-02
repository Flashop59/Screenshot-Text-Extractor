import os
import json
from PIL import Image
from google.cloud import vision
from openpyxl import Workbook
import streamlit as st
from io import BytesIO

# Load the Google Cloud credentials from the secret and save it as a file
if 'GOOGLE_CLOUD_CREDENTIALS' in os.environ:
    credentials_json = os.environ['GOOGLE_CLOUD_CREDENTIALS']
    with open('google_credentials.json', 'w') as f:
        f.write(credentials_json)
    
    # Set the GOOGLE_APPLICATION_CREDENTIALS environment variable to point to the saved JSON
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = 'google_credentials.json'

# Initialize the Google Cloud Vision API client
client = vision.ImageAnnotatorClient()

# Function to extract text from an image for a given crop area using Google Vision API
def extract_text_from_image(image, crop_area):
    # Crop the image
    cropped_image = image.crop(crop_area)
    
    # Convert the image to bytes
    img_byte_arr = BytesIO()
    cropped_image.save(img_byte_arr, format='PNG')
    img_byte_arr = img_byte_arr.getvalue()

    # Prepare image content for Vision API
    vision_image = vision.Image(content=img_byte_arr)

    # Perform text detection
    response = client.text_detection(image=vision_image)
    texts = response.text_annotations

    if texts:
        return texts[0].description.strip()
    return ""

# Function to extract the date from the filename using regex
def extract_date_from_filename(filename):
    import re
    match = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    if match:
        return match.group(1)
    return None

# Streamlit UI
st.title("Screenshot Text Extraction to Excel (Google Vision)")

# Upload files
uploaded_files = st.file_uploader("Upload Screenshot Files", accept_multiple_files=True, type=["png", "jpg"])

# Check if files are uploaded
if uploaded_files:
    # Show the first image as a sample
    sample_image = uploaded_files[0]
    image = Image.open(sample_image)

    # Display sample image
    st.image(image, caption="Sample Screenshot", use_column_width=True)

    # Define the crop areas (adjust the pixel coordinates as needed)
    crop_areas = [
        (45, 29, 201, 72),        # Data 1
        (813, 877, 1011, 1000),   # Data 2
        (580, 1414, 752, 1472),   # Data 3
        (842, 1419, 996, 1474),   # Data 4
        (73, 1650, 218, 1731)     # Data 5
    ]

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
