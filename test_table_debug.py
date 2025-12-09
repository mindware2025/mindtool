#!/usr/bin/env python3

import re

def debug_table_extraction():
    # Test with a sample PDF file
    print("=== TABLE EXTRACTION DEBUG ===")
    
    # Simulated raw text lines from your PDF
    sample_lines = [
        "Line",
        "Item", 
        "Quantity",
        "Customer",
        "Entitled Rate",
        "Months",
        "Customer Entitled",
        "Total Commit Value",
        "Discount %*",
        "Bid Rate*",
        "Bid",
        "Total Commit Value*",
        "Bid", 
        "Unit Price",
        "Partner",
        "Bid Rate",
        "Partner Bid",
        "Total Commit Value",
        "001",
        "1", 
        "0,00",
        "1-12",
        "0,00",
        "0,000",
        "0,00",
        "0,00",
        "0,00",
        "0,00",
        "0,00 USD",
        "",
        "002",
        "672",
        "215.712,00",
        "1-12", 
        "215.712,00",
        "50,000",
        "107.856,00",
        "107.856,00",
        "160,50",
        "99.227,52",
        "99.227,52 USD"
    ]
    
    print("\n=== RAW LINES ===")
    for i, line in enumerate(sample_lines):
        print(f"Line {i:3d}: {line}")
    
    print("\n=== TESTING BID TOTAL COMMIT VALUE SELECTION ===")
    
    # Test improved logic for selecting correct price
    from collections import Counter
    
    for i, line in enumerate(sample_lines):
        if re.match(r'^\s*00[1-9]', line):  # Line starts with 001, 002, etc.
            print(f"\n>>> TABLE ROW {line}:")
            
            # Collect all prices from next 10 lines
            all_prices = []
            for j in range(i+1, min(i+12, len(sample_lines))):
                check_line = sample_lines[j]
                found_prices = re.findall(r'\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}(?:\s*USD)?\b', check_line)
                if found_prices:
                    clean_prices = [p.replace(' USD', '').strip() for p in found_prices]
                    all_prices.extend(clean_prices)
            
            print(f"    All prices: {all_prices}")
            
            # Apply the same logic as the updated code
            total_price_candidates = [p for p in all_prices if not p.startswith('0,')]
            print(f"    Non-zero candidates: {total_price_candidates}")
            
            if total_price_candidates:
                print(f"    Non-zero candidates: {total_price_candidates}")
                print(f"    Positions: {[f'pos{i}:{p}' for i, p in enumerate(total_price_candidates)]}")
                
                # Use POSITIONAL logic - position 4 (index 3)
                if len(total_price_candidates) >= 4:
                    # Use 4th position (index 3) as it's typically "Bid Total Commit Value"
                    selected_price = total_price_candidates[3]
                    print(f"    Selected position 4 (index 3): {selected_price}")
                elif len(total_price_candidates) >= 2:
                    # Use 2nd position for shorter sequences
                    selected_price = total_price_candidates[1]
                    print(f"    Selected position 2 (index 1): {selected_price}")
                else:
                    # Only one price available
                    selected_price = total_price_candidates[0]
                    print(f"    Selected only available: {selected_price}")
            else:
                selected_price = all_prices[0] if all_prices else "0,00"
                print(f"    Selected (all zeros): {selected_price}")
            
            print(f"    FINAL RESULT: {selected_price}")

if __name__ == "__main__":
    debug_table_extraction()