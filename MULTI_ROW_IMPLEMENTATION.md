# Multi-Row Template 2 Extraction Implementation

## Changes Made to `ibm_template2.py`

### 1. Added Pre-Scan Logic (Lines ~275-282)
Before starting the main extraction loop, the code now:
- Scans all PDF lines for table row markers (001-006 format)
- Counts how many table rows exist
- Sets `is_multi_row_case = table_row_count >= 2`

**Purpose**: Detect multi-row subscription table cases early to skip Strategy 1

### 2. Conditional Wrapping of Strategy 1 (Line ~291)
Changed Strategy 1 trigger from:
```python
if is_subscription_part or is_overage_part:
```
To:
```python
if (is_subscription_part or is_overage_part) and not is_multi_row_case:
```

**Purpose**: Skip Strategy 1 entirely when multiple table rows are detected, allowing Strategy 2 to handle all extraction

### 3. Added Strategy 2: Table Row Extraction (Lines ~764-858)
New comprehensive strategy that:
- Detects table row markers (001, 002, 003, etc.)
- Extracts quantity with support for:
  - European period format: `1.550` → 1550
  - Comma format: `1,550` → 1550
  - Mixed format with smart detection
- Extracts duration (e.g., "1-12", "13-24")
- Finds SKU by searching entire document backwards
- Finds description from nearby IBM service lines
- Extracts unit price and total price
  - Parses European format: `1.272,00` → 1.272
  - Converts USD to AED using 3.6725 rate
- Calculates all pricing fields (Cost, Partner Price)
- Adds complete line item to `extracted_data`

### 4. Fixed Function References
- Changed `parse_euro_number()` calls to `parse_number()` (the correct function name)

## Key Improvements

### Quantity Extraction
- **European Period Format**: `1.550` is now correctly parsed as 1550 (thousands separator)
- **Fallback Parsing**: Handles integers, comma-formatted, and mixed decimal formats
- **Validation**: Smart detection of 3-digit groups to distinguish from decimals

### Description Handling
- Strategy 2 searches up to 30 lines backwards for IBM service description
- Fallback: Uses "IBM Subscription Service - {SKU}" if not found

### Pricing
- Uses `parse_number()` which handles European format (period=thousands, comma=decimal)
- Automatic USD→AED conversion for all prices
- Calculates Cost (Unit Price × Quantity)
- Calculates Partner Price with 8% discount

## How It Works

### For Single-Row Cases:
- `is_multi_row_case = False`
- Strategy 1 executes (subscription part extraction)
- Strategy 2 is skipped

### For Multi-Row Cases (001-006):
- `is_multi_row_case = True`
- Strategy 1 is skipped
- Strategy 2 processes each table row:
  1. Detects row marker (001, 002, etc.)
  2. Extracts qty, duration, prices
  3. Finds SKU from earlier in document
  4. Finds description from IBM service line
  5. Converts currencies and calculates totals
  6. Adds to results

## Testing
The implementation correctly handles the Template 2 multi-row case where:
- Single SKU (e.g., D28B4LL) appears in 6 billing period rows
- Each row has same SKU but different duration (months 1-12, 13-24, etc.)
- Each row has same price but appears multiple times in table

## Validation
✓ Pre-scan detects 6 table rows correctly
✓ Strategy 2 processes all 6 rows
✓ Each row has correct duration extracted
✓ SKU lookup works across entire document
✓ Price parsing handles European format
✓ USD→AED conversion applied
✓ No duplicate filtering (all rows extracted)
