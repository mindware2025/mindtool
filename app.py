import streamlit as st
import streamlit as st
import pandas as pd
from io import BytesIO
from ibm import extract_ibm_data_from_pdf, create_styled_excel, create_styled_excel_template2, correct_descriptions, extract_last_page_text
from ibm_template2 import extract_ibm_template2_from_pdf, get_extraction_debug
from sales.ibm_v2 import compare_mep_and_cost
from sales.mibb import correct_mibb_descriptions, create_mibb_excel, extract_mibb_header_from_pdf, extract_mibb_table_from_pdf
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
       
        "IBM Quotation",
        "IBM PDF to Excel (Template 2 Only) Disabled for now",  # Disabled for now
        "IBM Excel to Excel (Template 1 Only) Disabled for now",  # Disabled for now
        "MIBB Quotations"
    ]
    
)

def load_master_map(master_file):
    df = pd.read_excel(master_file) if master_file.name.endswith(".xlsx") else pd.read_csv(master_file)
    df = df.iloc[:, :2]   # FIRST TWO COLUMNS ONLY
    df.columns = ["part", "desc"]
    df["part"] = df["part"].astype(str).str.upper().str.replace(" ", "").str.replace("-", "")
    df["desc"] = df["desc"].fillna("").astype(str)
    return dict(zip(df["part"], df["desc"]))


if tool_choice == "IBM Quotation":

    st.header("🆕 IBM Excel to Excel + PDF to Excel (Combo)")
    st.info("Upload an IBM quotation PDF and (optionally) an Excel file. The tool will auto-detect the template and use the best logic for each.")

    # Country selection
    country = st.selectbox("Choose a country:", ["UAE", "Qatar", "KSA"])

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
            if result.get('date_validation_msg'):
                st.info(f"📅 Date Validation:\n{result['date_validation_msg']}")
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
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    
                    
                )
elif tool_choice == "MIBB Quotations":
    st.header("📋 MIBB Quotations")
    st.info("Upload a MIBB quotation PDF. The tool will extract header information and table data automatically.")

    logo_path = "image.png"

    st.subheader("📤 Upload MIBB Quotation PDF")
    uploaded_pdf = st.file_uploader(
        "Upload MIBB Quotation PDF (.pdf)",
        type=["pdf"],
        help="Upload a MIBB quotation PDF. The tool will extract header information and table data automatically."
    )

    st.subheader("📄 Upload Pricelist / Master File (Descriptions)")
    master_file = st.file_uploader(
        "Upload (.csv or .xlsx) – only first 2 columns used",
        type=["csv", "xlsx"]
    )

    if uploaded_pdf:
        # Extract header
        pdf_bytes = BytesIO(uploaded_pdf.getbuffer())
        header_info = extract_mibb_header_from_pdf(pdf_bytes)

        # Extract table data
        pdf_bytes.seek(0)
        table_data = extract_mibb_table_from_pdf(pdf_bytes)

        if master_file:
            master_map = load_master_map(master_file)
        else:
            master_map = None
            st.warning("please upload pricelist")

        # Correct descriptions
        table_data = correct_mibb_descriptions(table_data, master_map)

        # Missing SKUs warning (only once)
        missing = []
        if master_map:
            for r in table_data:
                part = str(r[0]).strip().upper()
                if part not in master_map:
                    missing.append(part)

        missing = list(dict.fromkeys(missing))

        if missing:
            st.warning(
                "⚠️ Some part numbers were not found in the master file. "
                "Descriptions were kept blank in Excel. Please double-check:\n\n"
                + ", ".join(missing)
            )

        # Create Excel
        if table_data:
            output = BytesIO()
            create_mibb_excel(
                data=table_data,
                header_info=header_info,
                logo_path=logo_path,
                output=output
            )

            st.success("✅ Excel file generated successfully!")

            st.download_button(
                label="📥 Download MIBB Quotation Excel",
                data=output.getvalue(),
                file_name="MIBB_Quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("👆 Please upload a MIBB quotation PDF to get started.")