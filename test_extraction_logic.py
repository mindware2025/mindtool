#!/usr/bin/env python3
"""Test the multi-row extraction logic"""

import re
import logging
from datetime import datetime

# Setup logging
logging.basicConfig(
    filename='test_extraction_logic.log',
    level=logging.DEBUG,
    format='%(message)s'
)
logger = logging.getLogger()

def add_debug(msg):
    """Add debug message"""
    print(msg)
    logger.info(msg)

# Test patterns
subscription_part_re = re.compile(r'\b[A-Z][A-Z0-9]{4,8}\b')
table_row_pattern = re.compile(r'^00[1-9]$|^0[1-9][0-9]$')
date_pattern = re.compile(r'\b\d{2}-[A-Za-z]{3}-\d{4}\b')

def parse_euro_number(value):
    """Parse European format number"""
    if not value:
        return 0
    value = str(value).strip()
    # Replace European separators
    value = value.replace(',', '#COMMA#').replace('.', ',').replace('#COMMA#', '.')
    try:
        return float(value)
    except:
        return 0

# Simulated PDF lines (from the multi-row case)
test_lines = [
    "IBM Maximo Application Suite per AppPoint Subscription License",
    "Subscription Part#: D28B4LL",
    "Billing: Annual",
    "Current Transaction Customer Unit Price: 1.272,00",
    "Channel Discount: 3%",
    "Subscription Length: 72 Months",
    "Price Change within Subscription: Increase 5% every 12 months",
    "Renewal Type: Expires at end of Subscription",
    "Renewal: No",
    "",
    "Line",
    "Item",
    "Quantity",
    "Months",
    "Bid Unit Price",
    "Bid Total Commit Value",
    "001",
    "1",
    "1-12",
    "1.272,00",
    "15.264,00",
    "002",
    "1",
    "13-24",
    "1.272,00",
    "15.264,00",
    "003",
    "1",
    "25-36",
    "1.272,00",
    "15.264,00",
    "004",
    "1",
    "37-48",
    "1.272,00",
    "15.264,00",
    "005",
    "1",
    "49-60",
    "1.272,00",
    "15.264,00",
    "006",
    "1",
    "61-72",
    "1.272,00",
    "15.264,00",
]

add_debug("="*80)
add_debug("TESTING MULTI-ROW EXTRACTION")
add_debug("="*80)

# Pre-scan
table_row_count = 0
for line in test_lines:
    if table_row_pattern.match(line.strip()):
        table_row_count += 1
        add_debug(f"Found table row: {line.strip()}")

add_debug(f"\nTable rows found: {table_row_count}")
is_multi_row_case = table_row_count >= 2
add_debug(f"Is multi-row case: {is_multi_row_case}")

# Extract rows
extracted = []
for i, line in enumerate(test_lines):
    line_stripped = line.strip()
    
    if table_row_pattern.match(line_stripped):
        add_debug(f"\n[ROW {line_stripped}] Processing...")
        
        # Extract quantity
        qty = 1
        if i + 1 < len(test_lines):
            qty_line = test_lines[i + 1].strip()
            add_debug(f"  Qty line: '{qty_line}'")
            try:
                qty = int(qty_line)
                add_debug(f"  ✓ Quantity: {qty}")
            except:
                add_debug(f"  ✗ Could not parse quantity")
        
        # Extract duration
        duration = "1-12"
        if i + 2 < len(test_lines):
            duration_line = test_lines[i + 2].strip()
            add_debug(f"  Duration line: '{duration_line}'")
            duration_match = re.search(r'(\d+)-(\d+)', duration_line)
            if duration_match:
                duration = f"{duration_match.group(1)}-{duration_match.group(2)}"
                add_debug(f"  ✓ Duration: {duration}")
        
        # Extract SKU (search backwards)
        sku = None
        for j in range(i - 1, -1, -1):
            search_text = test_lines[j]
            sku_match = subscription_part_re.search(search_text)
            if sku_match:
                potential_sku = sku_match.group()
                if any(c.isalpha() for c in potential_sku) and any(c.isdigit() for c in potential_sku):
                    if 5 <= len(potential_sku) <= 20:
                        sku = potential_sku
                        add_debug(f"  ✓ SKU found: {sku}")
                        break
        
        if not sku:
            add_debug(f"  ✗ No SKU found")
            continue
        
        # Extract price
        price_usd = 0
        if i + 3 < len(test_lines):
            price_line = test_lines[i + 3].strip()
            add_debug(f"  Price line: '{price_line}'")
            try:
                price_usd = parse_euro_number(price_line)
                add_debug(f"  ✓ Price: {price_usd}")
            except:
                add_debug(f"  ✗ Could not parse price")
        
        # Add to results
        extracted.append({
            'Row': line_stripped,
            'SKU': sku,
            'Qty': qty,
            'Duration': duration,
            'Price USD': price_usd
        })

add_debug("\n" + "="*80)
add_debug("EXTRACTION RESULTS")
add_debug("="*80)
for item in extracted:
    add_debug(f"Row {item['Row']}: {item['SKU']} x {item['Qty']} ({item['Duration']}) @ {item['Price USD']}")

add_debug(f"\nTotal rows extracted: {len(extracted)}")
