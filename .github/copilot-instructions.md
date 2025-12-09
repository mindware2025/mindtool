# Copilot Instructions for IBM Quotation Extractor

## Project Overview
This is a **Streamlit web application** that extracts line item data from IBM quotation PDFs and generates styled Excel quotations with proper formatting, terms, and compliance text. The app handles complex PDF parsing with European number formatting, automatic SKU description correction, and sophisticated quantity inference.

## Architecture & Core Components

### 1. **Data Flow Pipeline**
```
PDF Upload → PDF Text Extraction → Line Item Parsing → SKU Correction → Excel Generation → Download
```

- **`app.py`**: Streamlit frontend - handles file upload, displays data tables, triggers download
- **`ibm.py`**: Core PDF extraction engine with advanced parsing algorithms  
- **`terms_template.py`**: Business terms and compliance text generation
- **`xlsx_helpers.py`**: Excel formatting utilities for rich text styling

### 2. **Critical PDF Parsing Logic (`ibm.py`)**
The PDF parser uses **sliding window chunking** to handle wrapped table rows:
- Tries window sizes 12→1 to capture complete line items
- Uses strict regex patterns: `money_with_sep_re` requires separators (prevents false positives)
- **Quantity inference**: Matches `Entitled_Ext ≈ Qty × Entitled_Unit` with 2-cent tolerance
- **European number parsing**: Handles formats like `114.030,00` → `114030.00`

```python
# Key pattern matching in extract_ibm_data_from_pdf()
date_re = re.compile(r'\b\d{2}[‐‑–-][A-Za-z]{3}[‐‑–-]\d{4}\b')  # Multiple hyphen types
money_with_sep_re = re.compile(r'\d[\d.,]*[.,]\d+')  # Requires separators
```

### 3. **SKU Description Correction**
- Loads `Quotation IBM PriceList csv.csv` (60K+ SKUs) into memory on startup
- Replaces extracted descriptions with canonical ones from CSV
- Handles missing CSV gracefully with logging

## Key Development Patterns

### 1. **Error Handling & Logging**
- Comprehensive logging to `pdf_extraction_debug.log` 
- Graceful degradation when CSV missing or parsing fails
- Each parsing step wrapped in try/catch with fallbacks

### 2. **Currency Conventions**
- **USD to AED conversion**: Fixed rate `3.6725` (constant in `ibm.py`)
- All final prices in AED for UAE market compliance
- European decimal formatting throughout (`parse_euro_number()`)

### 3. **Excel Generation Strategy**
- **Two sheets**: Main quotation + IBM Terms (raw text)
- Logo insertion at fixed position with error handling
- Landscape orientation, fit-to-width printing
- Rich text formatting for compliance sections

## External Dependencies & Files

### Required Files
- **`image.png`**: Company logo (inserted into Excel header)
- **`Quotation IBM PriceList csv.csv`**: SKU master data (60K+ records)
- Both files must exist in project root

### Key Dependencies
```python
streamlit          # Web framework
pandas            # Data manipulation  
PyMuPDF (fitz)    # PDF text extraction
openpyxl          # Excel generation with styling
xlsxwriter        # Rich text Excel formatting
```

## Development Workflows

### Running the Application
```bash
streamlit run app.py
```

### Testing PDF Extraction
```python
# Test with sample PDF in Python REPL
from ibm import extract_ibm_data_from_pdf
with open("sample.pdf", "rb") as f:
    data, header = extract_ibm_data_from_pdf(f)
```

### Debugging Extraction Issues
1. Check `pdf_extraction_debug.log` for raw PDF text
2. Verify date patterns match: `DD-MMM-YYYY` format
3. Ensure money amounts contain separators (not plain integers)
4. Test quantity inference with different window sizes

## Project-Specific Conventions

### 1. **Data Structure Standards**
- Line items: `[sku, desc, qty, start_date, end_date, unit_price_aed, total_price_aed]`
- Header info: Dict with customer details, bid numbers, territory
- All dates in IBM format: `DD-MMM-YYYY`

### 2. **Parsing Robustness**
- **Multi-window approach**: Larger chunks first to capture wrapped content
- **Fallback mechanisms**: If main parsing fails, try division-only quantity inference  
- **Regex specificity**: Strict patterns prevent false matches in headers/footers

### 3. **Excel Styling Patterns**
- Blue theme (`1F497D`) for headers and terms
- Automatic column width adjustment for descriptions
- Print-optimized layout (landscape, scaled to fit)
- Compliance text formatting with specific cell positioning

## Common Gotchas

- **PDF parsing sensitivity**: IBM PDFs have inconsistent formatting across versions
- **European number formats**: Always use `parse_euro_number()` for currency parsing
- **SKU validation**: Must contain both letters and digits, 5-20 characters
- **Excel cell references**: Terms template uses specific cell addresses (B29, C30, etc.)
- **Streamlit config**: `st.set_page_config()` must be first command in `app.py`