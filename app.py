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
        from extract_ibm_terms import extract_ibm_terms_text
        output = io.BytesIO()

        # Initialize header_info and data
        header_info = {}
        data = []

        # Extract data and header info from PDF if uploaded
        if uploaded_pdf:
            try:
                data_from_pdf, extracted_header_info = extract_ibm_data_from_pdf(uploaded_pdf)
                header_info.update(extracted_header_info)
                logging.info("Header information extracted from PDF: %s", header_info)
                # Extract IBM Terms using robust multi-page logic
                uploaded_pdf.seek(0)
                ibm_terms_text = extract_ibm_terms_text(uploaded_pdf)
            except Exception as e:
                logging.error("Failed to extract header information or IBM Terms from PDF: %s", e)
                ibm_terms_text = ""

        # Updated to use the centralized `parse_uploaded_excel` function from `ibm_v2`
        if uploaded_excel:
            try:
                from sales.ibm_v2 import parse_uploaded_excel
                # Use BytesIO for in-memory parsing (no disk write)
                excel_bytes = BytesIO(uploaded_excel.getbuffer())
                data = parse_uploaded_excel(excel_bytes)
                logging.info("Data extracted from Excel using ibm_v2: %s", data)
            except Exception as e:
                logging.error("Failed to extract data from Excel: %s", e)
                st.error(f"❌ Failed to extract data from Excel: {e}")

        # Show preview of parsed data before generating Excel

        bid_number_match = True
        bid_number_error = None
        b13_val = None
        c13_val = None
        pdf_bid_number = None
        if uploaded_excel and uploaded_pdf:
            from sales.ibm_v2 import check_bid_number_match
            pdf_bid_number = header_info.get('Bid Number', '')
            excel_bytes_for_check = BytesIO(uploaded_excel.getbuffer())
            # Extract B13 and C13 for debug
            import pandas as pd
            try:
                xls_dbg = pd.ExcelFile(excel_bytes_for_check)
                df_dbg = xls_dbg.parse(xls_dbg.sheet_names[0], header=None)
                b13_val = str(df_dbg.iloc[12, 1]).strip() if df_dbg.shape[0] > 12 and df_dbg.shape[1] > 1 else ""
                c13_val = str(df_dbg.iloc[12, 2]).strip() if df_dbg.shape[0] > 12 and df_dbg.shape[1] > 2 else ""
            except Exception as e:
                b13_val = f"Error: {e}"
                c13_val = f"Error: {e}"
            # Reset BytesIO for actual check
            excel_bytes_for_check.seek(0)
            bid_number_match, bid_number_error = check_bid_number_match(excel_bytes_for_check, pdf_bid_number)

        # Show debug info for bid number matching
        if uploaded_excel and uploaded_pdf:
            st.info(f"PDF Bid Number: {pdf_bid_number}")
            st.info(f"Excel C13: {c13_val}")
        
        if uploaded_excel and uploaded_pdf and data:
            mep_cost_msg = compare_mep_and_cost(header_info, data)
            st.info(mep_cost_msg)

        if uploaded_excel:
            if not data:
                st.error("❌ No line items found in the uploaded Excel file. Please check the file format and ensure the second sheet contains valid data.")
            else:
                st.success(f"✅ Parsed {len(data)} line items from Excel.")
                st.dataframe(pd.DataFrame(data, columns=["SKU", "Description", "Quantity", "Start Date", "End Date", "Cost"]))

        # Only allow output if bid number matches
        if data and (uploaded_pdf and uploaded_excel):
            if not bid_number_match:
                st.error(bid_number_error)
            else:
                try:
                    create_styled_excel_v2(
                        data=data,
                        header_info=header_info,
                        logo_path=logo_path,
                        output=output,
                        compliance_text=compliance_text,
                        ibm_terms_text=ibm_terms_text if uploaded_pdf else ""
                    )
                    logging.info("Styled Excel file created successfully.")
                except Exception as e:
                    logging.error("Failed to create styled Excel file: %s", e)
                    st.error(f"❌ Failed to create styled Excel file: {e}")

                # Provide download link for the generated Excel file
                st.download_button(
                    label="📥 Download Styled Excel File",
                    data=output.getvalue(),
                    file_name="Styled_Quotation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

