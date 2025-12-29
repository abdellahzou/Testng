import requests
import pandas as pd
import os
import time
import streamlit as st
from typing import List, Tuple
import io
import zipfile
import base64
from PIL import Image  # Import PIL for thumbnail creation

def create_thumbnail(image_data: bytes, size: Tuple[int, int] = (128, 128)) -> bytes:
    """Creates a thumbnail from image data."""
    try:
        img = Image.open(io.BytesIO(image_data))
        img.thumbnail(size)
        buf = io.BytesIO()
        img.save(buf, format="PNG")  # Save as PNG to preserve transparency
        return buf.getvalue()
    except Exception as e:
        st.error(f"Error creating thumbnail: {e}")
        return b""  # Return empty bytes on error


def download_puma_images_individual(df: pd.DataFrame, product_type: str):
    """Downloads images, shows thumbnails, and provides download links."""
    df['full_product_id'] = df['full_product_id'].astype(str)

    if 'full_product_id' not in df.columns:
        st.error("Error: DataFrame must contain 'full_product_id' column.")
        return

    if product_type == "shoes":
        views = ["sv01", "sv03", "bv", "sv02"]
        base_url = "https://images.puma.com/image/upload/f_auto,q_auto,b_rgb:f7f7f7,w_2000,h_2000/global/{product_id}/{color_code}/{view}/PNA/fmt/png/"
    elif product_type == "textile":
        views = ["fnd", "bv", "mod01", "mod02", "mod03"]
        base_url = "https://images.puma.com/image/upload/f_auto,q_auto,b_rgb:f7f7f7,w_2000,h_2000/global/{product_id}/{color_code}/{view}/EEA/fmt/png/"
    else:
        st.error("Invalid product_type. Must be 'shoes' or 'textile'.")
        return

    for index, row in df.iterrows():
        full_product_id = row['full_product_id']
        if pd.isna(full_product_id):
            st.write(f"Skipping row {index + 1}: Missing 'full_product_id'")
            continue

        if len(full_product_id) < 2:
            st.write(f"Skipping row {index + 1}: Invalid 'full_product_id' length ({full_product_id})")
            continue

        product_id = full_product_id[:-2]
        color_code = full_product_id[-2:]

        for view_index, view in enumerate(views):
            if product_type == "textile" and view == "bv":
                final_url = base_url.format(product_id=product_id, color_code=color_code, view="bv/fnd")
            else:
                final_url = base_url.format(product_id=product_id, color_code=color_code, view=view)

            filename = f"{full_product_id}-{view_index}.png"
            st.write(filename) #show file name

            try:
                with st.spinner(f'Preparing: {filename}...'):
                    response = requests.get(final_url, stream=True)
                    response.raise_for_status()

                    # Create thumbnail
                    image_data = response.content
                    thumbnail_data = create_thumbnail(image_data)

                    # Display thumbnail using st.image, use_container_width=True
                    st.image(thumbnail_data, caption=f"Thumbnail for {filename}", use_container_width=False)


                    # Encode image data as base64 for download link
                    image_data_base64 = base64.b64encode(image_data).decode("utf-8")

                    # Create download link
                    download_link = f'<a href="data:image/png;base64,{image_data_base64}" download="{filename}">Download {filename}</a>'
                    st.markdown(download_link, unsafe_allow_html=True)

                time.sleep(0.5)


            except requests.exceptions.RequestException as e:
                st.error(f"Error preparing: {filename} - {e}")


def download_puma_images_zip(df: pd.DataFrame, product_type: str) -> io.BytesIO:
    """Downloads images and returns them as a zipped BytesIO object."""
    # ... (The same zip-based download function from previous responses) ...
    df['full_product_id'] = df['full_product_id'].astype(str)

    if 'full_product_id' not in df.columns:
        st.error("Error: DataFrame must contain 'full_product_id' column.")
        return None

    if product_type == "shoes":
        views = ["sv01", "sv03", "bv", "sv02"]
        base_url = "https://images.puma.com/image/upload/f_auto,q_auto,b_rgb:f7f7f7,w_2000,h_2000/global/{product_id}/{color_code}/{view}/PNA/fmt/png/"
    elif product_type == "textile":
        views = ["fnd", "bv", "mod01", "mod02", "mod03"]
        base_url = "https://images.puma.com/image/upload/f_auto,q_auto,b_rgb:f7f7f7,w_2000,h_2000/global/{product_id}/{color_code}/{view}/EEA/fmt/png/"
    else:
        st.error("Invalid product_type. Must be 'shoes' or 'textile'.")
        return None

    successful_downloads = 0
    failed_downloads = 0
    zip_buffer = io.BytesIO()  # Create the zip buffer *outside* the loop

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            full_product_id = row['full_product_id']
            if pd.isna(full_product_id):
                st.write(f"Skipping row {index + 1}: Missing 'full_product_id'")
                failed_downloads += 1
                continue

            if len(full_product_id) < 2:
                st.write(f"Skipping row {index + 1}: Invalid 'full_product_id' length ({full_product_id})")
                failed_downloads += 1
                continue

            product_id = full_product_id[:-2]
            color_code = full_product_id[-2:]

            for view_index, view in enumerate(views):
                if product_type == "textile" and view == "bv":
                    final_url = base_url.format(product_id=product_id, color_code=color_code, view="bv/fnd")
                else:
                    final_url = base_url.format(product_id=product_id, color_code=color_code, view=view)

                filename = f"{full_product_id}-{view_index}.png"

                try:
                    with st.spinner(f'Downloading: {filename}...'):
                        response = requests.get(final_url, stream=True)
                        response.raise_for_status()

                        # Write directly to the zip file *in memory*
                        zip_file.writestr(filename, response.content)
                        st.success(f"Downloaded: {filename}")
                        successful_downloads += 1
                        time.sleep(0.5)

                except requests.exceptions.RequestException as e:
                    st.error(f"Error downloading: {filename} - {e}")
                    failed_downloads += 1

    zip_buffer.seek(0)  # Rewind the buffer to the beginning

    if successful_downloads > 0:
        return zip_buffer
    else:
        return None

def create_sample_excel(product_type: str) -> io.BytesIO:
    """Creates sample Excel files (same as before)."""
    if product_type == "shoes":
        data = {
            'full_product_id': ['38227819', '09095102', '09095101', '12345699']
        }
    elif product_type == "textile":
        data = {
            'full_product_id': ['58684150', '62115801']
        }
    else:
        raise ValueError("Invalid product_type. Must be 'shoes' or 'textile'.")

    df = pd.DataFrame(data)
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    return excel_buffer

def main():
    st.title("Puma Image Downloader")

    # --- Product Type Selection ---
    product_type = st.radio("Select Product Type:", ("shoes", "textile"))

    # --- File Upload ---
    uploaded_file = st.file_uploader(f"Upload an Excel file for {product_type}", type=["xlsx"])

    # --- Sample File Download ---
    if st.checkbox(f"Use Sample Excel File for {product_type}"):
        sample_excel = create_sample_excel(product_type)
        st.download_button(
            label=f"Download Sample Excel File ({product_type})",
            data=sample_excel,
            file_name=f"sample_products_{product_type}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- Download Method Selection ---
    download_method = st.radio("Select Download Method:", ("Individual Images", "Zip File"))

    # --- Download Button and Logic ---
    if st.button(f"Download {product_type.capitalize()} Images"):
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file)
                if download_method == "Zip File":
                    zip_buffer = download_puma_images_zip(df, product_type)
                    if zip_buffer:
                        st.download_button(
                            label=f"Download All {product_type.capitalize()} Images as ZIP",
                            data=zip_buffer,
                            file_name=f"downloaded_{product_type}_images.zip",
                            mime="application/zip"
                        )
                    else:
                        st.warning("No images were downloaded.")
                elif download_method == "Individual Images":
                    download_puma_images_individual(df, product_type)

            except Exception as e:
                st.error(f"Error processing file: {e}")
        else:
            st.warning(f"Please upload an Excel file for {product_type}.")

if __name__ == "__main__":
    main()