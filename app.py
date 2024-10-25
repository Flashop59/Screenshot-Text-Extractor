import streamlit as st
import cv2
import numpy as np
from PIL import Image, ImageOps
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

# Function to allow drawing rectangles on images using OpenCV
def draw_rectangles(img):
    rectangles = []
    drawing = False
    ix, iy = -1, -1

    # Mouse callback function
    def draw_rectangle(event, x, y, flags, param):
        nonlocal ix, iy, drawing, rectangles
        if event == cv2.EVENT_LBUTTONDOWN:
            drawing = True
            ix, iy = x, y
        elif event == cv2.EVENT_MOUSEMOVE:
            if drawing:
                img_copy = img.copy()
                cv2.rectangle(img_copy, (ix, iy), (x, y), (0, 255, 0), 2)
                cv2.imshow("Image", img_copy)
        elif event == cv2.EVENT_LBUTTONUP:
            drawing = False
            rectangles.append((ix, iy, x, y))
            cv2.rectangle(img, (ix, iy), (x, y), (0, 255, 0), 2)
            cv2.imshow("Image", img)

    # Show the image in a window
    cv2.imshow("Image", img)
    cv2.setMouseCallback("Image", draw_rectangle)

    # Wait for the user to press 'q' to finish
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    return rectangles

# Streamlit app title and description
st.title("Screenshot Text Extraction with Interactive Rectangle Drawing")
st.write("Upload multiple images, draw rectangles on one image to select regions, and extract text from the selected regions across all images.")

# File uploader to upload multiple images
uploaded_files = st.file_uploader("Upload image files", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if uploaded_files:
    # Display the first image as a sample for drawing
    image_file = uploaded_files[0]
    pil_image = Image.open(image_file)
    
    # Convert the image to OpenCV format (from PIL to OpenCV format)
    open_cv_image = np.array(pil_image.convert('RGB'))  # Convert to RGB
    open_cv_image = open_cv_image[:, :, ::-1].copy()  # Convert RGB to BGR format for OpenCV

    # Display instruction
    st.write("Now, draw rectangles on the image to select regions for text extraction. Press 'q' to quit drawing.")
    
    # Allow users to draw rectangles on the image
    rectangles = draw_rectangles(open_cv_image)

    if rectangles:
        st.write(f"You have drawn {len(rectangles)} rectangles. The same regions will be used to extract text from all images.")

        # Button to start text extraction
        if st.button("Extract Text from All Images and Save to Excel"):
            wb = Workbook()
            ws = wb.active
            ws.title = "Extracted Data"
            ws.append(["Screenshot Name", "Rectangle Coordinates (Left, Top, Width, Height)", "Extracted Text"])

            # Process each uploaded image
            for image_file in uploaded_files:
                pil_image = Image.open(image_file)

                for rect in rectangles:
                    left, top, right, bottom = rect
                    width = right - left
                    height = bottom - top

                    # Extract text from the specified region
                    extracted_text = extract_text_from_image(pil_image, left, top, width, height)

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
