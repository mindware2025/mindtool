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
    st.info("Upload an IBM quotation Excel file. The tool will extract line items from the second sheet and generate a styled quotation Excel file, matching the original format.")

    logo_path = "image.png"
    compliance_text = ""  # Add compliance text if needed

    st.subheader("📤 Upload IBM Quotation Excel File")

    uploaded_excel = st.file_uploader(
        "Upload IBM Quotation Excel (.xlsx, .xlsm, .xls)",
        type=["xlsx", "xlsm", "xls"],
        help="Supports .xlsx, .xlsm, and .xls files. The tool will extract line items from the second sheet using the specified column mapping."
    )


    if uploaded_excel:
        import pandas as pd
        from sales.ibm_v2 import create_styled_excel_v2
        import io
        file_type = uploaded_excel.type
        file_name = uploaded_excel.name.lower()
        df = None
        # Always use pandas/xlrd for .xls
        if file_type == "application/vnd.ms-excel" or file_name.endswith(".xls"):
            try:
                xls = pd.ExcelFile(uploaded_excel, engine="xlrd")
                if len(xls.sheet_names) < 2:
                    st.error("❌ The uploaded Excel file does not have a second sheet.")
                    st.stop()
                df = xls.parse(xls.sheet_names[1], header=None)
            except Exception as e:
                st.error(f"❌ Error reading .xls file: {e}")
                st.stop()
        else:
            # Try openpyxl for .xlsx/.xlsm, fallback to pandas if it fails
            try:
                from openpyxl import load_workbook
                wb = load_workbook(uploaded_excel, data_only=True)
                sheetnames = wb.sheetnames
                if len(sheetnames) < 2:
                    st.error("❌ The uploaded Excel file does not have a second sheet.")
                    st.stop()
                ws = wb[sheetnames[1]]
                data_rows = list(ws.values)
                df = pd.DataFrame(data_rows)
            except Exception as e:
                # Fallback: try pandas
                try:
                    xls = pd.ExcelFile(uploaded_excel)
                    if len(xls.sheet_names) < 2:
                        st.error("❌ The uploaded Excel file does not have a second sheet.")
                        st.stop()
                    df = xls.parse(xls.sheet_names[1], header=None)
                except Exception as e2:
                    st.error(f"❌ Error reading Excel file: {e}\nFallback also failed: {e2}")
                    st.stop()

     
        def safe_cell(df, row, col):
            try:
                return df.iloc[row, col] if row < len(df.index) and col < len(df.columns) else ""
            except Exception:
                return ""

        header_info = {
            "Customer Name": safe_cell(df, 3, 2),
            "Bid Number": safe_cell(df, 4, 2),
            "PA Agreement Number": safe_cell(df, 5, 2),
            "PA Site Number": safe_cell(df, 6, 2),
            "Select Territory": safe_cell(df, 7, 2),
            "Government Entity (GOE)": safe_cell(df, 8, 2),
            "Reseller Name": safe_cell(df, 9, 2),
            "City": safe_cell(df, 10, 2),
            "Country": safe_cell(df, 11, 2),
            "Maximum End User Price (MEP)": safe_cell(df, 12, 2),
            "Bid Expiration Date": safe_cell(df, 13, 2),
        }

        data = []
        # Find the first data row (skip headers, look for first non-empty SKU)
        data = []
        for i in range(9, len(df)):
            row = df.iloc[i]
            sku = row[0] if len(row) > 0 else ""
            desc = row[1] if len(row) > 1 else ""
            qty = row[6] if len(row) > 6 else ""
            start_date = row[7] if len(row) > 7 else ""
            end_date = row[8] if len(row) > 8 else ""
            cost = row[18] if len(row) > 18 else ""
            # Stop if we hit a summary/total row or empty SKU
            if str(sku).strip() == "" or str(sku).strip().lower().startswith("total"):
                break
            data.append([sku, desc, qty, start_date, end_date, cost])

        if data:
          
            if st.button("🎯 Generate Styled Quotation Excel", type="primary", use_container_width=True, key="generate_excel_v2_btn"):
                with st.spinner("📊 Creating styled Excel quotation..."):
                    try:
                        output = BytesIO()
                        logo_path = "image.png"
                        compliance_text = ""
                        ibm_terms_text = ""
                        create_styled_excel_v2(
                            data,
                            header_info,
                            logo_path,
                            output,
                            compliance_text,
                            ibm_terms_text
                        )
                        st.download_button(
                            label="📥 Download Excel Quotation",
                            data=output.getvalue(),
                            file_name="IBM_Quotation_v2.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.success("✅ Excel file generated successfully!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Error generating Excel: {str(e)}")
                        st.exception(e)


    # (Removed duplicate/old upload logic for Excel files. Only robust, type-checked logic remains above.)

           