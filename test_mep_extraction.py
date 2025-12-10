#!/usr/bin/env python3
"""
Quick test to verify MEP extraction from ibm_template2.py
"""
from ibm_template2 import extract_ibm_template2_from_pdf, get_extraction_debug
import sys

# Assuming you have a test PDF file
test_pdf_path = "sample_template2.pdf"

try:
    with open(test_pdf_path, "rb") as f:
        data, header = extract_ibm_template2_from_pdf(f)
    
    print("\n" + "="*80)
    print("EXTRACTION RESULT")
    print("="*80)
    print(f"MEP Value: {header.get('Maximum End User Price (MEP)', 'NOT FOUND')}")
    print(f"Customer: {header.get('Customer Name', 'N/A')}")
    print(f"Bid Number: {header.get('Bid Number', 'N/A')}")
    print(f"Items extracted: {len(data)}")
    
    print("\n" + "="*80)
    print("DEBUG LOG (MEP related)")
    print("="*80)
    debug_info = get_extraction_debug()
    for msg in debug_info:
        if "MEP" in msg or "Maximum End User Price" in msg:
            print(msg)
    
except FileNotFoundError:
    print(f"Error: {test_pdf_path} not found")
    print("Please provide a test PDF file")
    sys.exit(1)
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
