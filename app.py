import streamlit as st
from PIL import Image
import pytesseract
from streamlit_drawable_canvas import st_canvas
from io import BytesIO
from openpyxl import Workbook

# Streamlit app title and description
st.title("Interactive Screenshot Text Extraction")
st.write("Upload an image, draw rectangles on it to select regions, and extract text from the selected regions into an Excel file.")

# File uploader to upload multiple images
uploaded_files = st.file_uploader("Upload image files", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if uploaded_files:
    # Select the first image as a sample
    image_file = uploaded_files[0]
    image = Image.open(image_file)

    # Display the image
    st.image(image, caption="Uploaded Image", use_column_width=True)

    # Get the image size
    img_width, img_height = image.size

    # Create a drawable canvas with an image as background
    canvas_result = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",  # Fill color with transparency
        stroke_width=3,
        stroke_color="#ff0000",
        background_image=image,
        update_streamlit=True,
        height=img_height,  # Match canvas height to image height
        width=img_width,    # Match canvas width to image width
        drawing_mode="rect",  # You can only draw rectangles
        key="canvas",
    )

    # When rectangles are drawn, show the extract button
    if canvas_result.json_data is not None:
        st.write("Drawn rectangles data:", canvas_result.json_data)

        # Extract the drawn rectangles from the canvas
        rect_data = canvas_result.json_data["objects"]
        if len(rect_data) > 0:
            st.write("You have drawn rectangles. Now, click the button to extract text.")

            # Button to start text extraction
            if st.button("Extract Text and Save to Excel"):
                wb = Workbook()
                ws = wb.active
                ws.title = "Extracted Data"
                ws.append(["Screenshot Name", "Rectangle Coordinates", "Extracted Text"])

                # Iterate over drawn rectangles and extract text
                for rect in rect_data:
                    # Get the coordinates from the rectangle drawn on the canvas
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
