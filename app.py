import streamlit as st
import streamlit as st
import pandas as pd
from io import BytesIO
from ibm import extract_ibm_data_from_pdf, create_styled_excel, create_styled_excel_template2, correct_descriptions, extract_last_page_text
from ibm_template2 import extract_ibm_template2_from_pdf, get_extraction_debug
from template_detector import detect_ibm_template

# ✅ Must be first Streamlit command
st.set_page_config(page_title="IBM Quotation Extractor", layout="wide")

# ---------------------------
# Tool selection UI
# ---------------------------
tool_choice = st.radio(
    "Select Tool:",
    ["IBM PDF to Excel (Existing)", "IBM Excel to Excel (New)"]
)

if tool_choice == "IBM PDF to Excel (Existing)":
    # ---------------------------
    # Static content
    # ---------------------------
    compliance_text = """<Paste compliance text here>"""
    logo_path = "image.png"

    st.title("🎯 IBM Quotation PDF to Excel Converter")
    st.markdown("Upload your IBM quotation PDF - the system will automatically detect the template type")

    # Sidebar info for supported templates
    with st.sidebar:
        st.header("📋 Supported Templates")
        st.info("""
        **Auto-Detection Available:**
        
        📦 **Template 1: Parts Information**
        - Coverage dates
        - Entitled/Bid pricing
        - Parts table structure
        
        ☁️ **Template 2: Software as a Service**
        - Subscription parts
        - Service agreements
        - Commit values
        """)

    # Create two columns for layout
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📁 Upload Master Price List (Optional)")
        master_csv = st.file_uploader(
            "Upload IBM Price List CSV", 
            type=["csv"], 
            key="ibm_master_csv",
            help="Upload the master CSV file to enhance quotation processing"
        )

    with col2:
        # Show upload status
        if master_csv:
            st.success("✅ Master CSV uploaded")
        else:
            st.info("📄 No master CSV uploaded")

    # Process master CSV if uploaded
    master_data = None
    if master_csv:
        try:
            master_data = pd.read_csv(master_csv)
            st.success(f"✅ Master data loaded: **{len(master_data)}** SKUs")
            
        except Exception as e:
            st.error(f"❌ Error reading master CSV: {e}")

    st.markdown("---")

    # PDF Upload Section
    st.subheader("📤 Upload IBM Quotation PDF")
    uploaded_file = st.file_uploader(
        "Upload IBM Quotation PDF (Auto-detects template)", 
        type=["pdf"],
        help="Supports both Parts Information and Software as a Service templates"
    )

    if uploaded_file:
        st.success("✅ PDF uploaded successfully!")
        
        # Create columns for template detection display
        col1, col2 = st.columns([3, 1])

elif tool_choice == "IBM Excel to Excel (New)":
    st.header("🆕 IBM Excel to Excel (v2)")
    st.info("Upload both an IBM quotation PDF and an Excel file. The tool will extract header information from the PDF and line items from the Excel file to generate a styled quotation Excel file.")

    logo_path = "image.png"
    compliance_text = ""  # Add compliance text if needed

    st.subheader("📤 Upload IBM Quotation Files")

    # File uploaders for both PDF and Excel
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

    if uploaded_pdf or uploaded_excel:
        import io
        import pandas as pd
        from sales.ibm_v2 import create_styled_excel_v2
        from ibm import extract_ibm_data_from_pdf
        output = io.BytesIO()

        # Initialize header_info and data
        header_info = {}
        data = []

        # Extract header info from PDF if uploaded
        if uploaded_pdf:
            _, extracted_header_info = extract_ibm_data_from_pdf(uploaded_pdf)
            header_info.update(extracted_header_info)

        # Extract table data from Excel if uploaded
        if uploaded_excel:
            file_type = uploaded_excel.type
            file_name = uploaded_excel.name.lower()
            if file_type == "application/vnd.ms-excel" or file_name.endswith(".xls"):
                try:
                    xls = pd.ExcelFile(uploaded_excel, engine="xlrd")
                    if len(xls.sheet_names) < 2:
                        st.error("❌ The uploaded Excel file does not have a second sheet.")
                    else:
                        df = xls.parse(xls.sheet_names[1])
                        data = df.values.tolist()  # Convert DataFrame to list of lists
                except Exception as e:
                    st.error(f"❌ Failed to read Excel file: {e}")
            else:
                try:
                    xls = pd.ExcelFile(uploaded_excel)
                    if len(xls.sheet_names) < 2:
                        st.error("❌ The uploaded Excel file does not have a second sheet.")
                    else:
                        df = xls.parse(xls.sheet_names[1])
                        data = df.values.tolist()  # Convert DataFrame to list of lists
                except Exception as e:
                    st.error(f"❌ Failed to read Excel file: {e}")

        # Call the function to create the styled Excel file
        create_styled_excel_v2(
            data=data,
            header_info=header_info,
            logo_path=logo_path,
            output=output,
            compliance_text=compliance_text,
            ibm_terms_text=""
        )

        # Provide download link for the generated Excel file
        st.download_button(
            label="📥 Download Styled Excel File",
            data=output.getvalue(),
            file_name="Styled_Quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

