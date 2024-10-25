import streamlit as st
from PIL import Image
import pytesseract
from streamlit_drawable_canvas import st_canvas
from io import BytesIO
from openpyxl import Workbook

# Streamlit app title and description
st.title("Interactive Screenshot Text Extraction")
st.write("Upload multiple images, draw rectangles on one image to select regions, and extract text from the selected regions across all images.")

# File uploader to upload multiple images
uploaded_files = st.file_uploader("Upload image files", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if uploaded_files:
    # Display the first image for drawing rectangles
    image_file = uploaded_files[0]
    image = Image.open(image_file)

    # Display the image for selecting areas
    st.image(image, caption="Uploaded Image", use_column_width=True)

    # Get the image dimensions
    img_width, img_height = image.size

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
        st.write("Drawn rectangles data:", canvas_result.json_data)

        # Extract the drawn rectangles from the canvas
        rect_data = canvas_result.json_data["objects"]

        if len(rect_data) > 0:
            st.write(f"You have drawn {len(rect_data)} rectangles. The same regions will be used to extract text from all images.")

            # Button to start text extraction
            if st.button("Extract Text from All Images and Save to Excel"):
                wb = Workbook()
                ws = wb.active
                ws.title = "Extracted Data"
                ws.append(["Screenshot Name", "Rectangle Coordinates", "Extracted Text"])

                # Process each uploaded image
                for image_file in uploaded_files:
                    image = Image.open(image_file)

                    for rect in rect_data:
                        # Get the coordinates from the first drawn rectangle
                        left = int(rect["left"])
                        top = int(rect["top"])
                        width = int(rect["width"])
                        height = int(rect["height"])

                        # Crop the region from the image
                        cropped_image = image.crop((left, top, left + width, top + height))

                        # Extract text from the cropped region using pytesseract
                        extracted_text = pytesseract.image_to_string(cropped_image, lang='eng').strip()

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
