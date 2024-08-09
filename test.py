import streamlit as st
from io import BytesIO
import os
import tempfile
from docx2pdf import convert as docx_convert
import zipfile
import platform

def word_to_pdf(docx_file):
    pdf_buffer = BytesIO()

    if platform.system() != "Windows":
        st.error("This functionality is only available on Windows.")
        return pdf_buffer

    # Initialize COM
    try:
        import comtypes.client
        comtypes.CoInitialize()
    except ImportError:
        st.error("comtypes library is required on Windows.")
        return pdf_buffer

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx_file:
            temp_docx_file.write(docx_file.read())
            temp_docx_path = temp_docx_file.name

        temp_pdf_path = tempfile.mktemp(suffix=".pdf")

        docx_convert(temp_docx_path, temp_pdf_path)

        with open(temp_pdf_path, "rb") as f:
            pdf_buffer.write(f.read())

        os.remove(temp_docx_path)
        os.remove(temp_pdf_path)
    except Exception as e:
        st.error(f"Error converting file: {e}")
    finally:
        # Uninitialize COM
        comtypes.CoUninitialize()

    pdf_buffer.seek(0)
    return pdf_buffer

st.title("Word to PDF Conversion")

uploaded_files = st.file_uploader("Upload Word files (.docx)", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Confirm Conversion"):
        if len(uploaded_files) == 1:
            uploaded_file = uploaded_files[0]
            file_name, file_extension = os.path.splitext(uploaded_file.name)
            
            if file_extension == ".docx":
                pdf_buffer = word_to_pdf(uploaded_file)
                st.download_button(
                    label=f"Download PDF",
                    data=pdf_buffer,
                    file_name=f"{file_name}.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("Only Word files (.docx) are supported.")
        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for uploaded_file in uploaded_files:
                    file_name, file_extension = os.path.splitext(uploaded_file.name)
                    
                    if file_extension == ".docx":
                        pdf_buffer = word_to_pdf(uploaded_file)
                        zip_file.writestr(f"{file_name}.pdf", pdf_buffer.getvalue())
                    else:
                        st.error("Only Word files (.docx) are supported.")
            
            zip_buffer.seek(0)
            
            st.download_button(
                label="Download ZIP File",
                data=zip_buffer,
                file_name="converted_files.zip",
                mime="application/zip"
            )
