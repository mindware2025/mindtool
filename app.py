import streamlit as st
import streamlit as st
import pandas as pd
from io import BytesIO
from ibm import extract_ibm_data_from_pdf, create_styled_excel, create_styled_excel_template2, correct_descriptions, extract_last_page_text
from ibm_template2 import extract_ibm_template2_from_pdf, get_extraction_debug
from sales.ibm_v2 import compare_mep_and_cost
from template_detector import detect_ibm_template
import logging

# Configure logging
logging.basicConfig(
    filename="output_log.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ✅ Must be first Streamlit command
st.set_page_config(page_title="IBM Quotation Extractor", layout="wide")

# ---------------------------
# Tool selection UI
# ---------------------------
tool_choice = st.radio(
    "Select Tool:",
    [
       
        "IBM Excel to Excel+ pdf to excel",
        "IBM PDF to Excel (Template 2 Only) Disabled for now",  # Disabled for now
        "IBM Excel to Excel (Template 1 Only) Disabled for now"  # Disabled for now
    ]
    
)


if tool_choice == "IBM Excel to Excel+ pdf to excel":
    st.header("🆕 IBM Excel to Excel + PDF to Excel (Combo)")
    st.info("Upload an IBM quotation PDF and (optionally) an Excel file. The tool will auto-detect the template and use the best logic for each.")

    # Country selection
    country = st.selectbox("Choose a country:", ["UAE", "Qatar"])

    logo_path = "image.png"
    compliance_text = ""  # Add compliance text if needed

    st.subheader("📤 Upload IBM Quotation Files")

    uploaded_pdf = st.file_uploader(
        "Upload IBM Quotation PDF (.pdf)",
        type=["pdf"],
        help="Supports .pdf files. The tool will extract header information from the PDF."
    )

    uploaded_excel = st.file_uploader(
        "Upload IBM Quotation Excel (.xlsx, .xlsm, .xls)",
        type=["xlsx", "xlsm", "xls"],
        help="Supports .xlsx, .xlsm, and .xls files. The tool will extract line items from the second sheet."
    )

    if uploaded_pdf:
        from sales.ibm_v2_combo import process_ibm_combo
        import io
        pdf_bytes = io.BytesIO(uploaded_pdf.getbuffer())
        excel_bytes = io.BytesIO(uploaded_excel.getbuffer()) if uploaded_excel else None
        result = process_ibm_combo(pdf_bytes, excel_bytes, country=country)

        if result['error']:
            st.error(f"❌ {result['error']}")
        else:
            st.success(f"✅ Detected Template: {result['template']}")
            if result['mep_cost_msg']:
                st.info(result['mep_cost_msg'])
            if result['bid_number_error']:
                st.error(result['bid_number_error'])
            if result['data']:
                if result.get('columns'):
                    st.dataframe(pd.DataFrame(result['data'], columns=result['columns']))
                else:
                    st.dataframe(pd.DataFrame(result['data']))
            if result.get('excel_bytes'):
                st.download_button(
                    label="📥 Download Styled Excel File",
                    data=result['excel_bytes'],
                    file_name="Styled_Quotation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

