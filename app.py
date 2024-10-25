import streamlit as st
from PIL import Image
import numpy as np
from streamlit_drawable_canvas import st_canvas
import pytesseract
from io import BytesIO
from openpyxl import Workbook

# Streamlit app title and description
st.title("Image Annotation and Text Extraction")
st.write("Upload an image, draw rectangles directly on it to select regions, and extract text from those regions.")

# File uploader to upload a single image
uploaded_file = st.file_uploader("Upload an image file", type=["png", "jpg", "jpeg"])

if uploaded_file:
    # Open the uploaded image
    image = Image.open(uploaded_file)
    img_width, img_height = image.size

    # Display the image
    st.image(image, caption="Uploaded Image", use_column_width=True)

    # Create a drawable canvas on top of the image
    canvas_result = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",  # Fill color with transparency
        stroke_width=3,
        stroke_color="#ff0000",
        background_image=image,  # Use image as canvas background
        update_streamlit=True,
        height=img_height,  # Set canvas height to image height
        width=img_width,    # Set canvas width to image width
        drawing_mode="rect",  # Restrict to drawing rectangles only
        key="canvas",
    )

    # When rectangles are drawn, show the extract button
    if canvas_result.json_data is not None:
        rect_data = canvas_result.json_data["objects"]

        if len(rect_data) > 0:
            st.write(f"You have drawn {len(rect_data)} rectangles. Click the button below to extract text.")

            # Button to start text extraction
            if st.button("Extract Text from Selected Regions"):
                wb = Workbook()
                ws = wb.active
                ws.title = "Extracted Data"
                ws.append(["Rectangle Coordinates", "Extracted Text"])

                # Process the image for text extraction
                for rect in rect_data:
                    # Get the coordinates from the drawn rectangle
                    left = int(rect["left"])
                    top = int(rect["top"])
                    width = int(rect["width"])
                    height = int(rect["height"])

                    # Crop the region from the image
                    cropped_image = image.crop((left, top, left + width, top + height))

                    # Extract text from the cropped region using pytesseract
                    extracted_text = pytesseract.image_to_string(cropped_image, lang='eng').strip()

                    # Append the results to the Excel sheet
                    ws.append([f"{left},{top},{width},{height}", extracted_text])

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
