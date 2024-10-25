import streamlit as st
from PIL import Image
import pytesseract
from streamlit_drawable_canvas import st_canvas
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
st.title("Interactive Screenshot Text Extraction")
st.write("Upload multiple images, draw rectangles on one image to select regions, and extract text from the selected regions across all images.")

# File uploader to upload multiple images
uploaded_files = st.file_uploader("Upload image files", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if uploaded_files:
    # Select the first image as a sample for drawing
    image_file = uploaded_files[0]
    image = Image.open(image_file)

    # Display the first image
    st.write("Use this image to draw the regions for text extraction")
    st.image(image, caption="Uploaded Image", use_column_width=True)

    # Get the image dimensions
    img_width, img_height = image.size

    # Create a drawable canvas on top of the image
    st.write("Draw rectangles on the image to select regions for text extraction")
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

    # Check if rectangles are drawn on the canvas
    if canvas_result.json_data is not None:
        rect_data = canvas_result.json_data["objects"]

        if len(rect_data) > 0:
            st.write(f"You have drawn {len(rect_data)} rectangles. The same regions will be used to extract text from all images.")

            # Button to start text extraction
            if st.button("Extract Text from All Images and Save to Excel"):
                wb = Workbook()
                ws = wb.active
                ws.title = "Extracted Data"
                ws.append(["Screenshot Name", "Rectangle Coordinates (Left, Top, Width, Height)", "Extracted Text"])

                # Process each uploaded image
                for image_file in uploaded_files:
                    image = Image.open(image_file)

                    for rect in rect_data:
                        left = int(rect["left"])
                        top = int(rect["top"])
                        width = int(rect["width"])
                        height = int(rect["height"])

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
