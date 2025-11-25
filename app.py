import streamlit as st
import pandas as pd
from io import BytesIO
from ibm import extract_ibm_data_from_pdf, create_styled_excel, correct_descriptions, extract_last_page_text
# ✅ Must be first Streamlit command
st.set_page_config(page_title="IBM Quotation Extractor", layout="wide")
# ---------------------------
# Static content
# ---------------------------
compliance_text = """<Paste compliance text here>"""
logo_path = "image.png"
# ---------------------------
# Streamlit UI
# ---------------------------
st.title("IBM Quotation PDF to Styled Excel Converter")
uploaded_file = st.file_uploader("Upload IBM Quotation PDF", type=["pdf"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_stream = BytesIO(file_bytes)
    ibm_terms_text = extract_last_page_text(BytesIO(file_bytes))
    st.success("✅ PDF uploaded successfully.")
    
    data, header_info = extract_ibm_data_from_pdf(BytesIO(file_bytes))
    corrected_data = correct_descriptions(data)  # Now handled in ibm.py
    if corrected_data and len(corrected_data) > 0:
        st.subheader("Corrected BoQ Data")
        df = pd.DataFrame(corrected_data, columns=[
            "SKU", "Product Description", "Quantity", "Start Date", "End Date",
            "Unit Price in AED", "Total Price in AED"
        ])
        st.dataframe(df, use_container_width=True)
        output = BytesIO()
        
       # create_styled_excel(corrected_data, header_info, logo_path, output, compliance_text, ibm_terms_text)
        create_styled_excel(corrected_data, header_info, logo_path, output, compliance_text, ibm_terms_text)
        st.download_button(
            label="📥 Download Styled Excel File",
            data=output.getvalue(),
            file_name="Styled_IBM_Quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("✅ Excel file includes logo, header details, styled table, summary row, terms, compliance text, and IBM Terms.")
    else:
        st.warning("⚠ No valid line items found in the PDF. Please check the format or try another file.")
else:
    st.info("📤 Please upload a PDF file to begin.")
