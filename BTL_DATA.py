import streamlit as st
import pandas as pd
import os
from PIL import Image
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# Configuration
EXCEL_FILE = "shops_data.xlsx"
TEMP_IMAGE = "temp_img.jpg"
MAX_IMAGE_HEIGHT = 300  # pixels (constrains height only)
IMAGE_COLUMN = 'C'  # Column where images will be placed
DATA_COLUMNS = ['Shop_ID', 'Region', 'last_updated']  # Other data columns

def load_or_create_excel():
    if os.path.exists(EXCEL_FILE):
        return load_workbook(EXCEL_FILE)
    else:
        wb = load_workbook()
        ws = wb.active
        ws.append(DATA_COLUMNS)
        # Set initial column widths
        ws.column_dimensions['A'].width = 15  # Shop_ID
        ws.column_dimensions['B'].width = 20  # Region
        ws.column_dimensions[IMAGE_COLUMN].width = 30  # Image column (initial)
        wb.save(EXCEL_FILE)
        return wb

def save_image_to_excel(shop_id, region, uploaded_file):
    try:
        wb = load_or_create_excel()
        ws = wb.active
        
        # Find existing row or create new
        row_idx = None
        for idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if row[0] == shop_id and row[1] == region:
                row_idx = idx
                break
        
        if not row_idx:
            row_idx = ws.max_row + 1
            ws.cell(row=row_idx, column=1, value=shop_id)
            ws.cell(row=row_idx, column=2, value=region)
        
        # Process image (convert RGBA to RGB if needed)
        img = Image.open(uploaded_file)
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        
        # Calculate new dimensions maintaining aspect ratio
        # Only constrain height, let width follow naturally
        if img.height > MAX_IMAGE_HEIGHT:
            ratio = MAX_IMAGE_HEIGHT / img.height
            new_height = MAX_IMAGE_HEIGHT
            new_width = int(img.width * ratio)
        else:
            new_width = img.width
            new_height = img.height
        
        img = img.resize((new_width, new_height))
        
        # Save temp image
        img.save(TEMP_IMAGE, "JPEG", quality=90)
        
        # Add to Excel with perfect width matching
        excel_img = ExcelImage(TEMP_IMAGE)
        
        # Convert image width to Excel column width
        # Excel column width: 1 unit ≈ 0.9cm of 72-point font characters
        # Approximation: 7 pixels ≈ 1 Excel width unit
        col_width = max(10, new_width / 7)  # Minimum width 10
        
        # Set column width to exactly match image width
        ws.column_dimensions[IMAGE_COLUMN].width = col_width
        
        # Set row height (approximate conversion)
        # Excel row height: 1 point = 1/72 inch ≈ 1.33 pixels
        row_height = max(15, new_height / 1.33)  # Minimum height 15
        ws.row_dimensions[row_idx].height = row_height
        
        # Add image to cell (will automatically fill the cell width)
        cell_ref = f"{IMAGE_COLUMN}{row_idx}"
        ws.add_image(excel_img, cell_ref)
        
        # Center align other cells vertically
        for col in range(1, len(DATA_COLUMNS) + 1):
            if get_column_letter(col) != IMAGE_COLUMN:
                ws.cell(row=row_idx, column=col).alignment = Alignment(vertical='center')
        
        # Add timestamp
        ws.cell(row=row_idx, column=len(DATA_COLUMNS), 
               value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")).alignment = Alignment(vertical='center')
        
        wb.save(EXCEL_FILE)
        return True
        
    except Exception as e:
        st.error(f"Error saving image: {str(e)}")
        return False
    finally:
        if os.path.exists(TEMP_IMAGE):
            os.remove(TEMP_IMAGE)

def display_excel_data():
    try:
        wb = load_or_create_excel()
        ws = wb.active
        
        data = []
        for row in ws.iter_rows(values_only=True):
            row_data = list(row[:2]) + [row[-1]] if len(row) > 2 else list(row) + [None]
            data.append(row_data)
        
        df = pd.DataFrame(data, columns=DATA_COLUMNS)
        return df, wb
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return pd.DataFrame(), None

def main():
    st.title("Shop Photo Upload System")
    
    st.sidebar.header("Navigation")
    app_mode = st.sidebar.radio("Go to", ["Upload Photo", "View Data"])
    
    if app_mode == "Upload Photo":
        df, _ = display_excel_data()
        
        regions = sorted(df["Region"].unique()) if not df.empty else []
        shop_ids = sorted(df["Shop_ID"].unique()) if not df.empty else []
        
        if not regions or not shop_ids:
            st.warning("No shop data found. Please add shop data first.")
            return
        
        selected_region = st.selectbox("Select Region", regions)
        filtered_shops = df[df["Region"] == selected_region]["Shop_ID"].unique()
        selected_shop = st.selectbox("Select Shop ID", filtered_shops)
        
        uploaded_file = st.file_uploader("Upload Shop Photo", type=["jpg", "jpeg", "png"])
        
        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption=f"Preview for Shop {selected_shop}", use_column_width=True)
            
            # Show image dimensions
            st.info(f"Image dimensions: {image.width}px × {image.height}px")
            
            if st.button("Save Photo"):
                if save_image_to_excel(selected_shop, selected_region, uploaded_file):
                    st.success(f"Photo saved successfully for Shop {selected_shop}!")
                    st.info("Cell width exactly matches image width")
                else:
                    st.error("Failed to save photo")
    
    elif app_mode == "View Data":
        st.subheader("Shop Data")
        
        df, _ = display_excel_data()
        st.dataframe(df)
        
        st.markdown("### Download Updated Excel File")
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                bytes_data = f.read()
            st.download_button(
                label="Download Excel File",
                data=bytes_data,
                file_name=EXCEL_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()