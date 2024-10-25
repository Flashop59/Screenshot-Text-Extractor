import streamlit as st
from PIL import Image
import pytesseract
from openpyxl import Workbook
from io import BytesIO

# Function to extract text from a specified region in an image
def extract_text_from_image(image, left, top, width, height):
    # Crop the region from the image
    cropped_image = image.crop((left, top, left + width, top + height))
    # Extract text from the cropped region using pytesseract
    extracted_text = pytesseract.image_to_string(cropped_image, lang='eng').strip()
    return extracted_text

# Streamlit app title and description
st.title("Screenshot Text Extraction with Manual Coordinates")
st.write("Upload multiple images, manually specify the coordinates for text extraction, and save the extracted text in an Excel file.")

# File uploader to upload multiple images
uploaded_files = st.file_uploader("Upload image files", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if uploaded_files:
    st.write("You can provide the coordinates for the region you want to extract text from. These coordinates will be applied to all uploaded images.")
    
    # Input fields to enter coordinates for cropping
    left = st.number_input("Left", min_value=0, step=1, help="X coordinate of the top-left corner")
    top = st.number_input("Top", min_value=0, step=1, help="Y coordinate of the top-left corner")
    width = st.number_input("Width", min_value=1, step=1, help="Width of the rectangle")
    height = st.number_input("Height", min_value=1, step=1, help="Height of the rectangle")
    
    # Display the first image as a sample
    image_file = uploaded_files[0]
    image = Image.open(image_file)
    
    # Display the uploaded image with the specified rectangle as a preview
    st.image(image, caption="Sample Image with Specified Region", use_column_width=True)
    
    # Button to start text extraction
    if st.button("Extract Text from All Images and Save to Excel"):
        # Create a new Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Data"
        ws.append(["Screenshot Name", "Coordinates (Left, Top, Width, Height)", "Extracted Text"])
        
        # Process each uploaded image
        for image_file in uploaded_files:
            image = Image.open(image_file)
            # Extract text from the specified region
            extracted_text = extract_text_from_image(image, left, top, width, height)
            # Append the results to the Excel sheet
            ws.append([image_file.name, f"{left},{top},{width},{height}", extracted_text])
        
        # Save the workbook to a BytesIO stream for download
        excel_output = BytesIO()
        wb.save(excel_output)
        excel_output.seek(0)
        
        st.success("Text extraction complete. Download the Excel file below.")
        st.download_button(
            label="Download Excel File",
            data=excel_output,
            file_name="extracted_text.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
