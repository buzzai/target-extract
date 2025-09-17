import streamlit as st
from docx import Document
import csv
import re
from io import BytesIO

def process_docx(uploaded_file):
    """
    Processes the uploaded docx file to extract target data and returns it as a list of lists.
    """
    doc = Document(uploaded_file)
    plain_text = "\n".join([p.text for p in doc.paragraphs])
    
    target_pattern = re.compile(r"Target(\d+)", re.IGNORECASE)
    center_pattern = re.compile(
        r"Center\s*:\s*([\d\.\-]+)\s*m;\s*([\d\.\-]+)\s*m;\s*([\d\.\-]+)\s*m",
        re.IGNORECASE
    )

    data = []
    current_target = None

    for line in plain_text.splitlines():
        line = line.strip()
        target_match = target_pattern.search(line)
        if target_match:
            current_target = f"Target{target_match.group(1)}"
        
        center_match = center_pattern.search(line)
        if center_match and current_target:
            x, y, z = center_match.groups()
            center_str = f"{x} m; {y} m; {z} m"
            data.append([current_target, center_str, x, y, z])
            current_target = None  # Reset current_target for the next match

    return data

st.title("Word to CSV Converter")

st.markdown("Upload a `.docx` file to extract target coordinates and convert them to a CSV file.")

uploaded_file = st.file_uploader("Choose a `.docx` file", type=["docx", "doc"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.doc'):
        st.warning("`.doc` files are not directly supported by this tool. Please convert your file to `.docx` format and try again.")
    else:
        st.success("File uploaded successfully!")
        st.info("Processing your file...")

        try:
            extracted_data = process_docx(uploaded_file)
            
            if extracted_data:
                st.subheader("Extracted Data")
                
                # Create a DataFrame for better display
                import pandas as pd
                df = pd.DataFrame(extracted_data, columns=["target", "center", "x", "y", "z"])
                st.dataframe(df)

                # Create CSV in-memory for download
                csv_buffer = BytesIO()
                df.to_csv(csv_buffer, index=False)
                csv_buffer.seek(0)
                
                st.download_button(
                    label="Download CSV",
                    data=csv_buffer,
                    file_name="extracted_targets.csv",
                    mime="text/csv",
                )
                st.success(f"Successfully processed {len(extracted_data)} rows and ready for download.")
            else:
                st.warning("No target data found in the document.")
        except Exception as e:
            st.error(f"An error occurred: {e}")