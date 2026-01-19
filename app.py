import streamlit as st
import streamlit as st
import pandas as pd
from io import BytesIO
from ibm import extract_ibm_data_from_pdf, create_styled_excel, create_styled_excel_template2, correct_descriptions, extract_last_page_text
from ibm_template2 import extract_ibm_template2_from_pdf, get_extraction_debug
from sales.ibm_v2 import compare_mep_and_cost
from sales.mibb import create_mibb_excel, extract_mibb_header_from_pdf, extract_mibb_table_from_pdf
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
        "IBM Excel to Excel (Template 1 Only) Disabled for now",  # Disabled for now
        "MIBB Quotations"
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
        pdf_bytes = BytesIO(uploaded_pdf.getbuffer())
        excel_bytes = BytesIO(uploaded_excel.getbuffer()) if uploaded_excel else None
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

elif tool_choice == "MIBB Quotations":
    st.header("📋 MIBB Quotations")
    st.info("Upload a MIBB quotation PDF. The tool will extract header information from page 1 and table data from page 2 (Parts Information table).")
    
    logo_path = "image.png"
    
    st.subheader("📤 Upload MIBB Quotation PDF")
    
    uploaded_pdf = st.file_uploader(
        "Upload MIBB Quotation PDF (.pdf)",
        type=["pdf"],
        help="Upload a MIBB quotation PDF. The tool will extract header information and table data automatically."
    )
    
    if uploaded_pdf:
        # Save uploaded PDF for debugging
        from datetime import datetime
        from pathlib import Path
        
        mibb_logs_dir = Path("mibb_logs")
        mibb_logs_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        pdf_filename = uploaded_pdf.name if hasattr(uploaded_pdf, 'name') else f"uploaded_{timestamp}.pdf"
        saved_pdf_path = mibb_logs_dir / f"uploaded_pdf_{timestamp}_{pdf_filename}"
        
        # Save PDF
        with open(saved_pdf_path, "wb") as f:
            f.write(uploaded_pdf.getbuffer())
        
        logging.info(f"MIBB PDF saved: {saved_pdf_path}")
        st.success(f"📄 PDF uploaded and saved for debugging: `{saved_pdf_path.name}`")
        
        # Extract header from PDF
        pdf_bytes = BytesIO(uploaded_pdf.getbuffer())
        header_info = extract_mibb_header_from_pdf(pdf_bytes)
        
        # Extract table data from page 2
        pdf_bytes.seek(0)  # Reset stream
        table_data = extract_mibb_table_from_pdf(pdf_bytes)
        
        # Show log file location
        from pathlib import Path
        log_files = sorted(Path("mibb_logs").glob("mibb_extraction_*.log"), reverse=True)
        if log_files:
            latest_log = log_files[0]
            st.info(f"📋 Debug log saved: `{latest_log.name}` (check `mibb_logs/` folder)")
        
        # Display extracted header info
        st.subheader("📄 Extracted Header Information")
        st.json(header_info)
        
        # Display extracted table data
        if table_data:
            st.subheader("📊 Extracted Table Data")
            df = pd.DataFrame(
                table_data,
                columns=["Part Number", "Description", "Start Date", "End Date", "QTY", "Price USD"]
            )
            st.dataframe(df)
            
            # Allow editing if needed
            st.subheader("✏️ Edit Table Data (Optional)")
            edited_df = st.data_editor(df, num_rows="dynamic")
            
            # Convert edited dataframe back to list
            table_data = edited_df.values.tolist()
        else:
            st.warning("⚠️ No table data found on page 2. Please check if the PDF contains a 'Parts Information' table.")
        
        if table_data:
            # Create Excel
            pdf_bytes.seek(0)  # Reset stream again
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